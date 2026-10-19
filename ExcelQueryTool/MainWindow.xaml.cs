using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using ExcelQueryTool.Services;

namespace ExcelQueryTool
{
    public partial class MainWindow : Window
    {
        private const int DefaultImageSize = 150, RowHeightMin = 30, ColumnWidthMax = 120;

        // 「大文件先看前几页」：界面始终绑定 _previewRows（前 LiveBindRows 行实时逐条上屏），
        // 全量数据始终进 _virtualRows（无通知不绑定，保持原灌数速度），加载完一次性重绑。
        // 教训（实测百万行 purchases_1048575）：全程绑定同一张表 = 每行 CollectionChanged，6s→10s；
        // 且 await foreach 同步完成时不让出调度器，渲染/输入饿死——必须定期 Dispatcher.Yield（静态方法）。
        // 2026-09-28 路线 A：DataTable/DataView → RowList（惰性 RowView），消掉 DataView 物化全表
        // 的首绑 ~1.5s 与 Rows.Add 灌表 ~0.5s（首开 5.5s → ~3.5s 的来源）。
        private const int LiveBindRows = 512;
        private const int LiveBindYieldBatch = 64;
        private const int BulkYieldBatch = 8192;
        private CancellationTokenSource? _loadingCts;
        private string? _filePath;
        private readonly ObservableCollection<string> _worksheets = [];
        private readonly RowList _virtualRows = new();
        private readonly RowList _previewRows = new(raiseChanges: true);   // 绑定中：需逐行通知支撑实时上屏
        private Stopwatch? _fileOpenWatch, _queryWatch;
        private int _filteredRowCount;
        private string? _rowHeightsWorksheet;

        /// <summary>忙碌期间被跳过的「显示图片」重载请求，待当前操作结束后补跑</summary>
        private bool _pendingReloadForImages;

        private static Stopwatch StartTimer() => Stopwatch.StartNew();

        private static string GetElapsedTime(Stopwatch? timer) => timer?.Elapsed.TotalSeconds.ToString("F3") ?? "0.000";

        /// <summary>
        /// 完成态状态栏消息：条数 + 耗时 + 可选「无此列」提示（提示在消息建成后清空）。
        /// 搜索与图片重载共用同一措辞。
        /// </summary>
        private void ReportCompleted(string prefix, Stopwatch? timer)
        {
            string hint = _unresolvedColumnsHint != null ? $"；⚠ 无此列：{_unresolvedColumnsHint}" : "";
            _uiStateManager.UpdateStatus($"{prefix} - 耗时 {GetElapsedTime(timer)}秒，共 {_filteredRowCount} 条记录{hint}");
            _unresolvedColumnsHint = null;
        }

        private void ResetCancellationTokens()
        {
            // Cancel 会让已注册回调同步跑完（此时调用返回后无残留回调），Dispose 安全；
            // 本项目不使用 CancelAfter，旧 CTS 无内部 Timer——仅纯托管对象，及时释放属卫生习惯
            _loadingCts?.Cancel();
            _loadingCts?.Dispose();
            _loadingCts = new CancellationTokenSource();
        }

        private bool CanStartOperation() => !_uiStateManager.IsProcessing && !_uiStateManager.IsLoadingData;

        private bool CanStartOperationWithFile() => CanStartOperation() && !string.IsNullOrEmpty(_filePath) && comboBoxWorksheets?.SelectedIndex != -1;

        private (string? filePath, string? worksheet, string? error) GetSelectedWorksheet()
        {
            if (string.IsNullOrEmpty(_filePath))
                return (null, null, "请先选择表格文件");

            if (comboBoxWorksheets?.SelectedItem == null)
                return (null, null, "请先选择工作表");

            string? selectedWorksheet = comboBoxWorksheets.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selectedWorksheet))
                return (null, null, "选择的工作表无效");

            return (_filePath, selectedWorksheet, null);
        }

        private async Task ExecuteWithErrorHandlingAsync(Func<Task> action, string operationName, Stopwatch? timer = null)
        {
            try
            {
                await action();
            }
            catch (OperationCanceledException)
            {
                _uiStateManager.UpdateStatus($"{operationName}已取消 - 耗时 {GetElapsedTime(timer)}秒");
            }
            catch (Exception ex)
            {
                _uiStateManager.UpdateStatus($"{operationName}失败 - 耗时 {GetElapsedTime(timer)}秒: {ex.Message}");
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 图片缓存：200 条、每 5 分钟清理一次；宽/高/单文件大小上限用 ImageManager 默认值（2048 px / 10 MB）
        private const int ImageCacheCapacity = 200;
        private static readonly TimeSpan ImageCacheCleanupInterval = TimeSpan.FromMinutes(5);
        private readonly ImageManager _imageManager = new(ImageCacheCapacity, ImageCacheCleanupInterval);
        private readonly SearchService _searchService = new();
        private readonly ExcelDataManager _excelDataManager = null!;
        private readonly UIStateManager _uiStateManager = null!;

        /// <summary>
        /// 内存过滤的全量源表：无关键词冷路径加载成功时建立（行数组引用，浏览模式行号）。
        /// 搜索不再重扫文件——遍历此表过滤即得结果；清空关键词回车也从这里恢复全表。
        /// </summary>
        private List<object[]>? _sourceRows;

        /// <summary>建缓存时的列定义/列名（内存路径重建界面列用，省掉开包取元数据）</summary>
        private List<ColumnMeta>? _sourceColumns;
        private string[]? _sourceHeaders;

        /// <summary>缓存匹配键：文件/工作表/图片勾选状态任一变化即失效，走冷路径重建</summary>
        private (string FilePath, string Worksheet, bool ShowImages)? _sourceKey;

        /// <summary>最近一次加载中未匹配到的列限定词（状态栏消息拼接后清空）</summary>
        private string? _unresolvedColumnsHint;

        public MainWindow()
        {
            try
            {
                InitializeComponent();

                // 版本号跟随程序集（csproj 的 <Version>），窗体标题与状态栏共用同一来源
                var asmVersion = typeof(MainWindow).Assembly.GetName().Version ?? new Version(0, 0);
                string versionText = $"v{asmVersion.Major}.{asmVersion.Minor}";
                Title = $"ExcelQueryTool {versionText}";
                statusVersionNumber.Text = versionText;

                _excelDataManager = new ExcelDataManager(_imageManager);
                _uiStateManager = new UIStateManager(
                    this,
                    btnOpenFile,
                    btnSearch,
                    comboBoxWorksheets,
                    chkShowImages,
                    statusLabelMessage);

                // 行虚拟化保留；虚拟化模式由 XAML 指定为 Standard。
                // 不要改回 Recycling：与 Pixel 滚动组合下回收容器的内容绑定会失效，
                // 翻页后出现成片空白单元格（双击单元格才会恢复），探针已复现并验证（2026-09-24）
                Loaded += MainWindow_Load;
                Closing += MainWindow_Closing;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"程序启动时发生错误: {ex.Message}\n\n堆栈跟踪:\n{ex.StackTrace}", "启动错误", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
        }

        private async void MainWindow_Load(object sender, RoutedEventArgs e)
        {
            var args = Environment.GetCommandLineArgs();
            if (args.Length > 1 && File.Exists(args[1])) await ProcessFileSelectionAsync(args[1]);
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            try
            {
                _loadingCts?.Cancel();
                _loadingCts?.Dispose();
                _loadingCts = null;

                // 包会话持有打开的 zip/文件流句柄：关窗立即释放，与换文件路径（ProcessFileSelectionAsync）对齐
                _excelDataManager.ReleaseSession();

                _imageManager.Dispose();

                // 内存过滤缓存：行引用列表，窗体关闭随 GC 回收，显式置空免滞留
                _sourceRows = null;
                _sourceColumns = null;
                _sourceHeaders = null;
                _sourceKey = null;

                // 行集合无非托管资源，Clear 掉行与视图缓存即可
                _virtualRows.Clear();
                _previewRows.Clear();

                _worksheets.Clear();

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
                // 关闭窗口时的清理失败不应阻止关闭；只留诊断痕迹
                Debug.WriteLine($"清理数据失败: {ex.Message}");
            }
        }

        private async void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "表格文件 (*.xlsx;*.csv;*.txt)|*.xlsx;*.csv;*.txt|所有文件 (*.*)|*.*",
                CheckFileExists = true
            };
            if (openFileDialog.ShowDialog() == true) await ProcessFileSelectionAsync(openFileDialog.FileName);
        }

        private async Task ProcessFileSelectionAsync(string path)
        {
            if (!File.Exists(path)) { MessageBox.Show("文件不存在", "错误", MessageBoxButton.OK, MessageBoxImage.Error); return; }

            // 旧版 .xls 各入口（对话框已被 Filter 拦截，拖拽/命令行仍可能传入）统一拦截，
            // 不再静默退化为空表——此前用户看到的是「打开完成 0 条」，无从知道原因
            if (IsLegacyExcelFile(path))
            {
                MessageBox.Show("不支持旧版 .xls（Excel 97-2003）文件，请用 Excel 打开后另存为 .xlsx 再加载。", "格式不支持", MessageBoxButton.OK, MessageBoxImage.Warning);
                _uiStateManager.UpdateStatus("不支持旧版 .xls 文件，请另存为 .xlsx");
                return;
            }

            try
            {
                // 计时起点 = 对话框确认选择/拖入文件之后（不含浏览选择文件的停留时间）；
                // 三入口（按钮/拖拽/命令行）都汇入此处，口径一致
                _fileOpenWatch = StartTimer();

                // 文件能力标记：txt/csv 无图片列与多表，完成后「显示图片」与工作表下拉保持禁用
                _uiStateManager.SetCurrentFileIsText(TextTableReader.IsTextFile(path));

                if (txtKeyword != null) txtKeyword.Text = "";
                _uiStateManager.SetProcessingState(true);
                _uiStateManager.UpdateStatus("正在加载文件...");

                ResetCancellationTokens();

                // 换文件：列结构与行号全部换血，列宽/行高记忆作废（不保留）
                _uiStateManager.ClearDataGrid(dataGridViewResults, preserveColumnWidths: false);
                _rowHeightsWorksheet = null;
                _pendingReloadForImages = false;   // 新文件：旧的待处理重载请求失效
                _virtualRows.Clear();
                _previewRows.Clear();
                _imageManager.Clear();
                _excelDataManager.ReleaseSession();   // 旧文件的包会话（句柄）立即释放

                _filePath = path;
                if (lblFilePath != null) lblFilePath.Content = Path.GetFileName(path);

                await LoadWorksheetsAsync(path);
                _fileOpenWatch.Stop();

                // 行数以 _filteredRowCount 为准：批量模式下 _virtualRows 在重绑前不含预览段，
                // 预览-only（小表）时数据在 _previewRows——两者都不能直接用 Rows.Count
                var recordCount = _filteredRowCount;
                _uiStateManager.UpdateStatus(recordCount > 0
                    ? $"✅ 文件打开完成 - 耗时 {GetElapsedTime(_fileOpenWatch)}秒，共 {recordCount} 条记录"
                    : $"✅ 文件打开完成 - 耗时 {GetElapsedTime(_fileOpenWatch)}秒");
            }
            catch (IOException ioEx)
            {
                MessageBox.Show($"文件正在被其他程序使用: {ioEx.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"文件错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                _uiStateManager.UpdateStatus($"加载失败: {ex.Message}");
            }
            finally { _uiStateManager.SetProcessingState(false); }
        }

        private async Task LoadWorksheetsAsync(string path)
        {
            if (comboBoxWorksheets == null) return;
            comboBoxWorksheets.ItemsSource = null;
            comboBoxWorksheets.IsEnabled = false;
            comboBoxWorksheets.Text = "加载中...";

            try
            {
                List<string> worksheets = await Task.Run(() => _excelDataManager.GetWorksheetNames(path));

                _worksheets.Clear();
                worksheets.ForEach(_worksheets.Add);
                _uiStateManager.UpdateWorksheetList(_worksheets);
                comboBoxWorksheets.Text = "";

                if (_worksheets.Any())
                {
                    comboBoxWorksheets.SelectedIndex = 0;
                    if (!string.IsNullOrEmpty(_filePath)) await LoadFirstWorksheetAsync(_filePath, _worksheets[0]);
                }
            }
            catch (Exception ex)
            {
                _uiStateManager.UpdateStatus($"加载工作表失败: {ex.Message}");
                MessageBox.Show($"加载工作表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { comboBoxWorksheets.IsEnabled = !TextTableReader.IsTextFile(path); }
        }

        private async Task LoadFirstWorksheetAsync(string filePath, string worksheetName)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            // ProcessDataAsync 内部管理处理状态（入口置 true、finally 统一置 false），
            // 异常也有统一弹窗/状态栏，此处无需再包一层
            await ProcessDataAsync(filePath, worksheetName, "");
        }

        private async void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            _uiStateManager.UpdateStatus("搜索中...");
            if (!CanStartOperation()) return;

            _queryWatch?.Stop();
            _queryWatch = StartTimer();
            ResetCancellationTokens();
            _uiStateManager.ClearDataGrid(dataGridViewResults);

            var (filePath, worksheet, error) = GetSelectedWorksheet();
            if (error != null)
            {
                _uiStateManager.UpdateStatus(error);
                return;
            }

            await ExecuteWithErrorHandlingAsync(async () =>
            {
                await ProcessDataAsync(filePath!, worksheet!, txtKeyword?.Text.Trim() ?? "");
                _queryWatch.Stop();
                txtKeyword?.SelectAll();
                ReportCompleted("✅ 搜索完成", _queryWatch);
            }, "搜索", _queryWatch);

            // 搜索期间被跳过的「显示图片」勾选：此处补跑一次，免去用户再点一次搜索
            if (await FlushPendingReloadAsync()) return;
        }

        private async Task ProcessDataAsync(string filePath, string worksheetName, string keyword)
        {
            if (string.IsNullOrEmpty(filePath) || string.IsNullOrEmpty(worksheetName))
            {
                _uiStateManager.UpdateStatus("工作表对象为空");
                return;
            }

            _queryWatch ??= StartTimer();

            ResetCancellationTokens();
            var token = _loadingCts!.Token;
            _uiStateManager.SetLoadingDataState(true);
            _uiStateManager.SetProcessingState(true);

            try
            {
                _uiStateManager.ClearDataGrid(dataGridViewResults);

                // 切换工作表：RowNumber 从 1 重排，上个表的行高记忆不能带过来（同表搜索/勾选图片则保留）
                if (_rowHeightsWorksheet != worksheetName)
                {
                    _uiStateManager.ResetUserRowHeights();
                    _rowHeightsWorksheet = worksheetName;
                }

                bool showImages = _uiStateManager.IsShowImagesChecked();
                bool cacheHit = _sourceRows is { Count: > 0 }
                    && _sourceKey is { } key
                    && key.FilePath == filePath && key.Worksheet == worksheetName && key.ShowImages == showImages;

                // 关键词解析一次、绑定按路径各做一次（列名表：内存=缓存列名，冷=本次列元数据）；
                // prepared 同时供过滤、「无此列」提示与命中词高亮共享
                var prepared = keyword.Length == 0 ? null : _searchService.Prepare(keyword);

                if (cacheHit)
                {
                    await ProcessFromMemoryCacheAsync(prepared, token);
                }
                else
                {
                    // 冷路径（key 不匹配：换文件/换表/图片状态变化）：旧缓存不主动清，
                    // 仅靠 key 比对失效 + 本次加载成功后按需覆盖——勾选图片搜索后再取消勾选，
                    // 与原缓存 key 重新匹配，仍可内存过滤，不必重扫
                    await ProcessFromDiskAsync(filePath, worksheetName, keyword, showImages, prepared, token);
                }

                // 行高通过 LoadingRow 事件按需计算（虚拟化下不遍历行）；此处只需挂一次
                _uiStateManager.AttachAutoRowHeight(dataGridViewResults, RowHeightMin);

                FinalizeSearchExtras(prepared);

                _uiStateManager.UpdateStatus($"✅ 加载完成 - 耗时 {GetElapsedTime(_queryWatch)}秒，共 {_filteredRowCount} 条记录");
            }
            catch (OperationCanceledException) { _uiStateManager.UpdateStatus($"加载已取消 - 耗时 {GetElapsedTime(_queryWatch)}秒"); }
            catch (OutOfMemoryException)
            {
                _imageManager.Dispose();
                GC.Collect();
                MessageBox.Show("内存不足，已清除图片缓存", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                _uiStateManager.UpdateStatus($"⚠ 内存不足，部分数据可能未加载 - 耗时 {GetElapsedTime(_queryWatch)}秒");
            }
            catch (Exception ex)
            {
                _uiStateManager.UpdateStatus($"加载失败 - 耗时 {GetElapsedTime(_queryWatch)}秒: {ex.Message}");
                MessageBox.Show($"错误详情:\n\n错误信息: {ex.Message}\n\n类型: {ex.GetType().Name}", "详细错误信息", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _uiStateManager.SetLoadingDataState(false);
                _uiStateManager.SetProcessingState(false);
            }
        }

        /// <summary>
        /// 内存过滤/恢复路径：匹配键（<see cref="_sourceKey"/>）保证文件/表/图片状态与建缓存时一致，
        /// 列结构即缓存列结构 —— 跳过开包取元数据与整表重扫，直接遍历源表行。
        /// </summary>
        private async Task ProcessFromMemoryCacheAsync(SearchService.PreparedSearch? prepared, CancellationToken token)
        {
            if (prepared != null) _searchService.BindColumns(prepared, _sourceHeaders);
            var cachedColumns = _sourceColumns!;
            _uiStateManager.InitializeDataGridColumns(dataGridViewResults, cachedColumns, DefaultImageSize, ColumnWidthMax);
            SetupRowList(_virtualRows, cachedColumns);
            SetupRowList(_previewRows, cachedColumns);

            _filteredRowCount = 0;
            _uiStateManager.UpdateDataGrid(dataGridViewResults, _previewRows);

            await FillTablesAsync(FilterSourceRowsAsync(prepared, token), token);
        }

        /// <summary>
        /// 冷路径：开包取列元数据 + 整表流式扫描灌表。
        /// 无关键词加载顺路建立内存过滤缓存（行数组引用，字符串与界面表共享）；
        /// 赋值在加载成功之后，取消/异常走不到这里，不会留半截缓存。
        /// </summary>
        private async Task ProcessFromDiskAsync(
            string filePath, string worksheetName, string keyword,
            bool showImages, SearchService.PreparedSearch? prepared, CancellationToken token)
        {
            var (columns, ok, layout, textLayout) = await _excelDataManager.PrepareWorksheetAsync(filePath, worksheetName, showImages, token);

            // 空表（工作表不存在 / 无表头）：初始化成空结果
            if (!ok)
            {
                _uiStateManager.InitializeDataGridColumns(dataGridViewResults, [], DefaultImageSize, ColumnWidthMax);
                SetupRowList(_virtualRows, []);
                _uiStateManager.UpdateDataGrid(dataGridViewResults, _virtualRows);
                return;
            }

            _uiStateManager.InitializeDataGridColumns(dataGridViewResults, columns, DefaultImageSize, ColumnWidthMax);
            SetupRowList(_virtualRows, columns);
            SetupRowList(_previewRows, columns);

            var progress = new Progress<string>(_uiStateManager.UpdateStatus);
            _filteredRowCount = 0;

            // 界面绑定预览表：前 LiveBindRows 行实时逐条上屏（先看几页）；
            // 全量数据始终进 _virtualRows（无通知，保持原灌数速度），加载完一次性重绑。
            // 前 LiveBindRows 行同时进两张表：预览表供显示，数据表保证重绑后行号 1 起完整。
            _uiStateManager.UpdateDataGrid(dataGridViewResults, _previewRows);

            List<object[]>? collect = keyword.Length == 0 ? [] : null;
            string[] headers = [.. columns.Select(c => c.Header)];
            if (prepared != null) _searchService.BindColumns(prepared, headers);

            await FillTablesAsync(_excelDataManager.LoadWorksheetDataAsync(
                filePath,
                worksheetName,
                prepared,
                showImages,
                token,
                progress,
                expectedDisplayColumnCount: columns.Count,
                knownLayout: layout,
                knownTextLayout: textLayout), token, collect);

            if (collect is { Count: > 0 })
            {
                _sourceRows = collect;
                _sourceColumns = columns;
                _sourceHeaders = headers;
                _sourceKey = (filePath, worksheetName, showImages);
            }
        }

        /// <summary>
        /// 搜索/加载收尾：登记「无此列」提示、更新命中词高亮。
        /// <paramref name="prepared"/> 为 null = 浏览模式，两者清空。
        /// </summary>
        private void FinalizeSearchExtras(SearchService.PreparedSearch? prepared)
        {
            _unresolvedColumnsHint = prepared?.UnresolvedColumns.Count > 0
                ? string.Join("、", prepared.UnresolvedColumns)
                : null;
            SearchHighlight.Instance.Terms = prepared?.HighlightTerms is { Count: > 0 } terms ? terms : null;
        }

        /// <summary>
        /// 消费行流灌入两个行集合：前 LiveBindRows 行预览集合实时上屏（带通知），批量段只进无通知的全量集合，
        /// 完成后一次性重绑（冷路径与内存过滤路径共用同一条灌表管线）。
        /// <paramref name="collect"/> 非 null 时顺路收集行引用（冷路径无关键词加载建立内存过滤缓存用）。
        /// </summary>
        private async Task FillTablesAsync(IAsyncEnumerable<object[]> rows, CancellationToken token, List<object[]>? collect = null)
        {
            bool bulkPhase = false;
            int sinceYield = 0;
            await foreach (var rowData in rows)
            {
                if (rowData == null) continue;

                // 前 LiveBindRows 行同时进两个集合：预览集合供实时显示，
                // 全量集合保证重绑后行号 1 起完整；批量段只进无通知的全量集合
                if (!bulkPhase) _previewRows.Add(rowData);
                _virtualRows.Add(rowData);
                _filteredRowCount++;
                collect?.Add(rowData);

                // 定期让出 UI 线程：await foreach 通道有数据时同步完成，
                // 不让出则渲染/输入永远排不上队（实测「绑定了也不显示」的根因）。
                // 批量段低频让出即可（保持滚动与状态栏响应，开销可忽略）
                if (++sinceYield >= (bulkPhase ? BulkYieldBatch : LiveBindYieldBatch))
                {
                    sinceYield = 0;
                    await Dispatcher.Yield();
                }

                if (!bulkPhase && _filteredRowCount >= LiveBindRows)
                    bulkPhase = true;
            }

            // 小表/少命中（从未进入批量段）：预览集合已是完整数据，保持绑定即可
            if (bulkPhase)
                _uiStateManager.UpdateDataGrid(dataGridViewResults, _virtualRows);
        }

        /// <summary>
        /// 内存过滤源表行（<see cref="_sourceRows"/> 在 UI 线程被逐行驱动）。
        /// 命中行必须 Clone 后再重排行号——直接改会污染缓存源行，下次恢复全表行号就错了；
        /// 空关键词（恢复全表）直接复用原行：不重排，显示表存的也是值副本，无需克隆。
        /// </summary>
        private async IAsyncEnumerable<object[]> FilterSourceRowsAsync(SearchService.PreparedSearch? prepared, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken token)
        {
            long emitted = 0;
            int sinceYield = 0;
            foreach (var row in _sourceRows!)
            {
                token.ThrowIfCancellationRequested();
                if (!SearchService.Matches(prepared, row)) continue;

                object[] output;
                if (prepared == null)
                {
                    output = row;
                }
                else
                {
                    output = (object[])row.Clone();
                    output[0] = (int)++emitted;
                }
                yield return output;

                // 匹配在 UI 线程同步跑：定期让出，避免大表上渲染/输入卡顿（与加载消费同节奏）
                if (++sinceYield >= BulkYieldBatch)
                {
                    sinceYield = 0;
                    await Dispatcher.Yield();
                }
            }
        }

        /// <summary>
        /// 重建行集合的列映射（全量与预览结构完全一致，绑定才能互切）。
        /// RowNumber 由 RowList.Configure 固定占下标 0，与解析管线的行布局一致。
        /// 预览集合在加载期间绑定显示前 <see cref="LiveBindRows"/> 行。
        /// </summary>
        private static void SetupRowList(RowList list, IEnumerable<ColumnMeta> columns)
            => list.Configure([.. columns.Select(c => c.Header)]);

        private void ComboBoxWorksheets_SelectedIndexChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CanStartOperationWithFile())
                BtnSearch_Click(sender, e);
        }

        private async void ChkShowImages_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (CanStartOperationWithFile())
            {
                await ReloadCurrentWorksheetAsync();
                return;
            }

            // 忙碌中（正加载/搜索）勾选「显示图片」：
            // 早先直接 return，勾选框已打勾但数据没重载 —— 用户看到图片列空着，
            // 必须再手动点一次搜索才出来（「首次勾选图片不显示」的根因）。
            // 现改为登记待处理请求，当前操作结束后自动补一次重载。
            if (!string.IsNullOrEmpty(_filePath) && comboBoxWorksheets?.SelectedIndex != -1)
                _pendingReloadForImages = true;
        }

        /// <summary>
        /// 补跑被跳过的重载请求。由各加载入口在收尾时调用（仅文件已就绪时生效）。
        /// 返回 true 表示已接管后续流程，调用方不应再重复收尾。
        /// </summary>
        private async Task<bool> FlushPendingReloadAsync()
        {
            if (!_pendingReloadForImages) return false;
            if (string.IsNullOrEmpty(_filePath) || comboBoxWorksheets?.SelectedIndex == -1) return false;

            _pendingReloadForImages = false;
            await ReloadCurrentWorksheetAsync();
            return true;
        }

        private async Task ReloadCurrentWorksheetAsync()
        {
            var (filePath, worksheet, error) = GetSelectedWorksheet();
            if (error != null || worksheet == null) return;

            // 计时重新开始：图片重载与打开文件/搜索一样报「耗时」，
            // 否则 ProcessDataAsync 里的 _queryWatch ??= 会沿用上次的旧表，累计出错误时长
            _queryWatch?.Stop();
            _queryWatch = StartTimer();

            _uiStateManager.UpdateStatus("正在重新加载图片数据...");

            string currentKeyword = txtKeyword?.Text.Trim() ?? "";

            await ExecuteWithErrorHandlingAsync(async () =>
            {
                await ProcessDataAsync(filePath!, worksheet, currentKeyword);
                ReportCompleted("✅ 图片显示已更新", _queryWatch);
            }, "重新加载");

            // 重载期间又被勾选/取消：补跑一次，保证最终状态与勾选框一致
            if (await FlushPendingReloadAsync()) return;
        }

        private void TxtKeyword_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) { e.Handled = true; BtnSearch_Click(sender, e); }
        }

        // 窗口与 DataGrid 的拖放入口：同一份处理逻辑（XAML 直接挂接本方法）
        private async void HandleDropAsync(object sender, DragEventArgs e)
        {
            e.Handled = true;
            if (e.Data.GetData(DataFormats.FileDrop) is not string[] files || files.Length != 1) return;

            // 忙碌守卫：加载/搜索/图片重载进行中再拖入文件，会并发跑两个处理流程
            // （按钮禁用拦不住拖放事件）——两边同时 Clear/重绑表格与缓存，交错写坏。
            // 拒收不弹窗（拖拽高频操作，与 .xls 拦截同一策略），状态栏说明原因。
            if (_uiStateManager.IsProcessing)
            {
                _uiStateManager.UpdateStatus("⏳ 正在处理中，请等待当前操作完成后再拖入文件");
                return;
            }

            if (IsSupportedFile(files[0]))
            {
                try
                {
                    await ProcessFileSelectionAsync(files[0]);
                }
                catch (Exception ex) { MessageBox.Show($"拖放文件错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
            }
            else
            {
                // 拒收不弹窗（拖拽高频操作）：状态栏说明原因；旧版 .xls 给出可行出路
                _uiStateManager.UpdateStatus(IsLegacyExcelFile(files[0])
                    ? "不支持旧版 .xls（Excel 97-2003），请用 Excel 另存为 .xlsx 后再打开"
                    : $"不支持的文件类型：{Path.GetExtension(files[0])}");
            }
        }

        private void HandleDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length == 1 && IsSupportedFile(files[0]))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }

        private static bool IsSupportedFile(string path)
        {
            string ext = Path.GetExtension(path);
            return ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".csv", StringComparison.OrdinalIgnoreCase)
                || ext.Equals(".txt", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>旧版 .xls（Excel 97-2003，OLE2 复合文档）：解析引擎只认 OOXML zip 包，不支持该格式</summary>
        private static bool IsLegacyExcelFile(string path)
            => Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase);
    }
}

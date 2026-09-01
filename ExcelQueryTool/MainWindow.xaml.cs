using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using OfficeOpenExcel = OfficeOpenXml;
using ExcelQueryTool.Services;

namespace ExcelQueryTool
{
    public partial class MainWindow : Window
    {
        private const int DefaultImageSize = 150, RowHeightMin = 30, ColumnWidthMax = 120;
        private CancellationTokenSource? _loadingCts;
        private string? _filePath;
        private readonly ObservableCollection<string> _worksheets = [];
        private readonly DataTable _virtualDataTable = new();
        private Stopwatch? _fileOpenWatch, _queryWatch;
        private OfficeOpenExcel.ExcelPackage? _cachedPackage;
        private string? _cachedFilePath;
        private readonly Lock _packageLock = new();
        private int _filteredRowCount;

        private static Stopwatch StartTimer() => Stopwatch.StartNew();

        private static string GetElapsedTime(Stopwatch? timer) => timer?.Elapsed.TotalSeconds.ToString("F3") ?? "0.000";

        private void ResetCancellationTokens()
        {
            _loadingCts?.Cancel();
            _loadingCts = new CancellationTokenSource();
        }

        private bool CanStartOperation() => !_uiStateManager.IsProcessing && !_uiStateManager.IsLoadingData;

        private bool CanStartOperationWithFile() => CanStartOperation() && !string.IsNullOrEmpty(_filePath) && comboBoxWorksheets?.SelectedIndex != -1;

        private (OfficeOpenExcel.ExcelPackage? package, OfficeOpenExcel.ExcelWorksheet? worksheet, string? error) GetSelectedWorksheet()
        {
            if (string.IsNullOrEmpty(_filePath))
                return (null, null, "请先选择Excel文件");

            if (comboBoxWorksheets?.SelectedItem == null)
                return (null, null, "请先选择工作表");

            string? selectedWorksheet = comboBoxWorksheets.SelectedItem?.ToString();
            if (selectedWorksheet == null)
                return (null, null, "选择的工作表无效");

            var package = GetOrCreatePackage();
            if (package == null)
                return (null, null, "无法打开Excel文件");

            var worksheet = package.Workbook.Worksheets[selectedWorksheet];
            if (worksheet == null)
                return (null, null, "选择的工作表不存在");

            return (package, worksheet, null);
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

        private readonly ImageManager _imageManager = new(200, TimeSpan.FromMinutes(5), 2048, 2048, 10 * 1024 * 1024);
        private readonly SearchService _searchService = new();
        private readonly ExcelDataManager _excelDataManager = null!;
        private readonly UIStateManager _uiStateManager = null!;

        static MainWindow() => OfficeOpenExcel.ExcelPackage.License.SetNonCommercialPersonal("My Name");

        public MainWindow()
        {
            try
            {
                InitializeComponent();
                
                _excelDataManager = new ExcelDataManager(_imageManager, _searchService);
                _uiStateManager = new UIStateManager(
                    this,
                    btnOpenFile,
                    btnSearch,
                    comboBoxWorksheets,
                    chkShowImages,
                    statusLabelMessage);
                
                if (dataGridViewResults != null)
                {
                    VirtualizingStackPanel.SetIsVirtualizing(dataGridViewResults, true);
                    VirtualizingStackPanel.SetVirtualizationMode(dataGridViewResults, VirtualizationMode.Recycling);
                }
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
                
                _imageManager?.Dispose();
                
                lock (_packageLock)
                {
                    _cachedPackage?.Dispose();
                    _cachedPackage = null;
                    _cachedFilePath = null;
                }
                
                _virtualDataTable?.Dispose();
                _virtualDataTable?.Clear();
                
                _worksheets?.Clear();
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch
            {
            }
        }

        private async void BtnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
                CheckFileExists = true
            };
            if (openFileDialog.ShowDialog() == true) await ProcessFileSelectionAsync(openFileDialog.FileName);
        }

        private async Task ProcessFileSelectionAsync(string path)
        {
            if (!File.Exists(path)) { MessageBox.Show("文件不存在", "错误", MessageBoxButton.OK, MessageBoxImage.Error); return; }

            try
            {
                _fileOpenWatch = StartTimer();
                if (txtKeyword != null) txtKeyword.Text = "";
                _uiStateManager.SetProcessingState(true);
                _uiStateManager.UpdateStatus("正在加载文件...");

                ResetCancellationTokens();

                _uiStateManager.ClearDataGrid(dataGridViewResults);
                _virtualDataTable.Clear();
                _virtualDataTable.Columns.Clear();
                _imageManager.Clear();

                _filePath = path;
                if (lblFilePath != null) lblFilePath.Content = Path.GetFileName(path);

                lock (_packageLock)
                {
                    _cachedPackage?.Dispose();
                    _cachedPackage = null;
                    _cachedFilePath = null;
                }

                using var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                await LoadWorksheetsAsync(fileStream);
                _fileOpenWatch.Stop();

                var recordCount = _virtualDataTable.Rows.Count;
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

        private async Task LoadWorksheetsAsync(FileStream stream)
        {
            if (comboBoxWorksheets == null) return;
            comboBoxWorksheets.ItemsSource = null;
            comboBoxWorksheets.IsEnabled = false;
            comboBoxWorksheets.Text = "加载中...";

            try
            {
                List<string> worksheets;
                using (var package = new OfficeOpenExcel.ExcelPackage(stream))
                {
                    worksheets = ExcelDataManager.GetWorksheetNames(package);
                }

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
            finally { comboBoxWorksheets.IsEnabled = true; }
        }

        private async Task LoadFirstWorksheetAsync(string filePath, string worksheetName)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            try
            {
                _uiStateManager.SetProcessingState(true);
                var package = GetOrCreatePackage();
                if (package == null) { _uiStateManager.UpdateStatus("无法打开Excel文件"); return; }
                var worksheet = package.Workbook.Worksheets[worksheetName];
                if (worksheet != null) await ProcessDataAsync(worksheet, "");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                _uiStateManager.UpdateStatus("加载失败");
            }
            finally { _uiStateManager.SetProcessingState(false); }
        }

        private async void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            _uiStateManager.UpdateStatus("搜索中...");
            if (!CanStartOperation()) return;

            _queryWatch?.Stop();
            _queryWatch = StartTimer();
            _uiStateManager.SetProcessingState(true);
            ResetCancellationTokens();
            _uiStateManager.ClearDataGrid(dataGridViewResults);

            var (_, worksheet, error) = GetSelectedWorksheet();
            if (error != null)
            {
                _uiStateManager.UpdateStatus(error);
                _uiStateManager.SetProcessingState(false);
                return;
            }

            await ExecuteWithErrorHandlingAsync(async () =>
            {
                await ProcessDataAsync(worksheet!, txtKeyword?.Text.Trim() ?? "");
                _queryWatch.Stop();
                txtKeyword?.SelectAll();
                _uiStateManager.UpdateStatus($"✅ 搜索完成 - 耗时 {GetElapsedTime(_queryWatch)}秒，共 {_filteredRowCount} 条记录");
            }, "搜索", _queryWatch);

            _uiStateManager.SetProcessingState(false);
        }

        private OfficeOpenExcel.ExcelPackage? GetOrCreatePackage()
        {
            if (string.IsNullOrEmpty(_filePath)) return null;

            lock (_packageLock)
            {
                if (_cachedPackage != null && _cachedFilePath == _filePath)
                {
                    return _cachedPackage;
                }

                try
                {
                    _cachedPackage?.Dispose();
                    _cachedPackage = new OfficeOpenExcel.ExcelPackage(new FileInfo(_filePath));
                    _cachedFilePath = _filePath;
                    return _cachedPackage;
                }
                catch
                {
                    _cachedPackage = null;
                    _cachedFilePath = null;
                    return null;
                }
            }
        }

        private async Task ProcessDataAsync(OfficeOpenExcel.ExcelWorksheet worksheet, string keyword)
        {
            if (worksheet == null || worksheet.Dimension == null)
            {
                _uiStateManager.UpdateStatus(worksheet == null ? "工作表对象为空" : "工作表为空");
                return;
            }

            _queryWatch ??= StartTimer();

            _loadingCts?.Cancel();
            _loadingCts?.Dispose();
            _loadingCts = new CancellationTokenSource();
            var token = _loadingCts.Token;
            _uiStateManager.SetLoadingDataState(true);

            try
            {
                _uiStateManager.SetProcessingState(true);
                _uiStateManager.ClearDataGrid(dataGridViewResults);

                _imageManager.Clear();

                await Task.Run(() => _imageManager.BuildPictureIndex(worksheet, token), token);
                var columns = ExcelDataManager.GetColumnMetadata(worksheet, _imageManager.PictureMap) ?? [];

                int totalRows = worksheet.Dimension.Rows;
                if (totalRows <= 1)
                {
                    _uiStateManager.InitializeDataGridColumns(dataGridViewResults, columns, DefaultImageSize, ColumnWidthMax);
                    SetupDataTable(columns);
                    _uiStateManager.UpdateDataGrid(dataGridViewResults, _virtualDataTable);
                    _uiStateManager.UpdateStatus($"✅ 加载完成 - 耗时 {GetElapsedTime(_queryWatch)}秒，共 0 条记录");
                    return;
                }

                _uiStateManager.InitializeDataGridColumns(dataGridViewResults, columns, DefaultImageSize, ColumnWidthMax);
                SetupDataTable(columns);

                var progress = new Progress<string>(_uiStateManager.UpdateStatus);
                _filteredRowCount = 0;

                _virtualDataTable.BeginLoadData();
                try
                {
                    await foreach (var rowData in _excelDataManager.LoadWorksheetDataAsync(
                        worksheet,
                        keyword,
                        _uiStateManager.IsShowImagesChecked(),
                        token,
                        progress))
                    {
                        if (rowData != null)
                        {
                            _virtualDataTable.Rows.Add(rowData);
                            _filteredRowCount++;
                        }
                    }
                }
                finally
                {
                    _virtualDataTable.EndLoadData();
                }

                _uiStateManager.UpdateDataGrid(dataGridViewResults, _virtualDataTable);
                _uiStateManager.ApplyAutoRowHeight(dataGridViewResults, RowHeightMin);

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

        private void SetupDataTable(IEnumerable<ColumnMeta> columns)
        {
            _virtualDataTable.BeginLoadData();
            try
            {
                _virtualDataTable.Columns.Clear();
                _virtualDataTable.Rows.Clear();
                _virtualDataTable.Columns.Add("RowNumber", typeof(int));
                foreach (var col in columns)
                {
                    _virtualDataTable.Columns.Add(col.Header, col.IsImage ? typeof(object) : typeof(string));
                }
            }
            finally
            {
                _virtualDataTable.EndLoadData();
            }
        }

        private void ComboBoxWorksheets_SelectedIndexChanged(object sender, SelectionChangedEventArgs e)
        {
            if (CanStartOperationWithFile())
                BtnSearch_Click(sender, e);
        }

        private void ChkShowImages_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (CanStartOperationWithFile())
                ReloadCurrentWorksheet();
        }

        private async void ReloadCurrentWorksheet()
        {
            var (package, worksheet, error) = GetSelectedWorksheet();
            if (error != null || worksheet == null) return;

            _uiStateManager.SetProcessingState(true);
            _uiStateManager.UpdateStatus("正在重新加载图片数据...");

            string currentKeyword = txtKeyword?.Text.Trim() ?? "";

            await ExecuteWithErrorHandlingAsync(async () =>
            {
                await ProcessDataAsync(worksheet, currentKeyword);
                int displayCount = !string.IsNullOrWhiteSpace(currentKeyword) ? _filteredRowCount : _virtualDataTable.Rows.Count;
                _uiStateManager.UpdateStatus($"✅ 图片显示已更新 - 共 {displayCount} 条记录");
            }, "重新加载");

            _uiStateManager.SetProcessingState(false);
        }

        private void TxtKeyword_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) { e.Handled = true; BtnSearch_Click(sender, e); }
        }

        private async void DataGrid_Drop(object sender, DragEventArgs e) => await HandleDrop(sender, e);
        private void DataGrid_DragEnter(object sender, DragEventArgs e) => HandleDragEnter(e);
        private async void Window_Drop(object sender, DragEventArgs e) => await HandleDrop(sender, e);
        private void Window_DragEnter(object sender, DragEventArgs e) => HandleDragEnter(e);
        
        private async Task HandleDrop(object sender, DragEventArgs e)
        {
            e.Handled = true;
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length == 1 && IsExcelFile(files[0]))
            {
                try
                {
                    await ProcessFileSelectionAsync(files[0]);
                }
                catch (Exception ex) { MessageBox.Show($"拖放文件错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
            }
        }

        private static void HandleDragEnter(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length == 1 && IsExcelFile(files[0]))
                e.Effects = DragDropEffects.Copy;
            else
                e.Effects = DragDropEffects.None;
        }

        private static bool IsExcelFile(string path) => string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase) || string.Equals(Path.GetExtension(path), ".xls", StringComparison.OrdinalIgnoreCase);
    }
}

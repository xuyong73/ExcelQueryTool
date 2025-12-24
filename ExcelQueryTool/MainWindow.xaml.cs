using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using OfficeOpenExcel = OfficeOpenXml;
using OfficeOpenExcelDrawing = OfficeOpenXml.Drawing;

namespace ExcelQueryTool
{
    public partial class MainWindow : Window
    {
        private const int DefaultImageSize = 150, RowHeightMax = 100, RowHeightMin = 30, ColumnWidthMax = 120, BaseBatchSize = 10000;
        private CancellationTokenSource? _loadingCts, _cts;
        private bool _isProcessing, _isLoadingData;
        private string? _filePath;
        private readonly ObservableCollection<string> _worksheets = new();
        private ImageCacheManager? _imageCache;
        private Dictionary<(string, int, int), OfficeOpenExcelDrawing.ExcelPicture?>? _pictureMap;
        private readonly DataTable _virtualDataTable = new();
        private Stopwatch? _fileOpenWatch, _queryWatch;

        static MainWindow() => OfficeOpenExcel.ExcelPackage.License.SetNonCommercialPersonal("My Name");

        public MainWindow()
        {
            try
            {
                InitializeComponent();
                _imageCache = new ImageCacheManager(200, TimeSpan.FromMinutes(5));
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
            _cts?.Cancel();
            _loadingCts?.Cancel();
            _imageCache?.Dispose();
            ClearDataGrid();
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
                _fileOpenWatch = Stopwatch.StartNew();
                if (txtKeyword != null) txtKeyword.Text = "";
                SetProcessingState(true);
                UpdateStatus("正在加载文件...");

                _loadingCts?.Cancel();
                _cts?.Cancel();
                _loadingCts = new CancellationTokenSource();
                _cts = new CancellationTokenSource();

                ClearDataGrid();
                _virtualDataTable.Clear();
                _virtualDataTable.Columns.Clear();
                _pictureMap?.Clear();
                _pictureMap ??= new Dictionary<(string, int, int), OfficeOpenExcelDrawing.ExcelPicture?>();

                _filePath = path;
                if (lblFilePath != null) lblFilePath.Content = Path.GetFileName(path);

                using var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                await LoadWorksheetsAsync(fileStream);
                _fileOpenWatch.Stop();

                var recordCount = _virtualDataTable.Rows.Count;
                UpdateStatus(recordCount > 0 
                    ? $"✅ 文件打开完成 - 耗时 {_fileOpenWatch.Elapsed.TotalSeconds:F3}秒，共 {recordCount} 条记录"
                    : $"✅ 文件打开完成 - 耗时 {_fileOpenWatch.Elapsed.TotalSeconds:F3}秒");
            }
            catch (IOException ioEx)
            {
                MessageBox.Show($"文件正在被其他程序使用: {ioEx.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"文件错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                UpdateStatus($"加载失败: {ex.Message}");
            }
            finally { SetProcessingState(false); }
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
                    worksheets = package.Workbook.Worksheets.Where(ws => !string.IsNullOrWhiteSpace(ws.Name)).Select(ws => ws.Name).ToList();
                }

                _worksheets.Clear();
                worksheets.ForEach(ws => _worksheets.Add(ws));
                comboBoxWorksheets.ItemsSource = _worksheets;
                comboBoxWorksheets.IsEnabled = true;
                comboBoxWorksheets.Text = "";

                if (_worksheets.Any())
                {
                    comboBoxWorksheets.SelectedIndex = 0;
                    if (!string.IsNullOrEmpty(_filePath)) await LoadFirstWorksheetAsync(_filePath, _worksheets[0]);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"加载工作表失败: {ex.Message}");
                MessageBox.Show($"加载工作表失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { comboBoxWorksheets.IsEnabled = true; }
        }

        private async Task LoadFirstWorksheetAsync(string filePath, string worksheetName)
        {
            if (string.IsNullOrEmpty(filePath)) return;
            try
            {
                SetProcessingState(true);
                using var package = new OfficeOpenExcel.ExcelPackage(new FileInfo(filePath));
                var worksheet = package.Workbook.Worksheets[worksheetName];
                if (worksheet != null) await ProcessDataAsync(worksheet, "");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                UpdateStatus("加载失败");
            }
            finally { SetProcessingState(false); }
        }

        private async void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            UpdateStatus("搜索中...");
            if (_isProcessing || _isLoadingData) return;

            _queryWatch = Stopwatch.StartNew();
            SetProcessingState(true);
            _loadingCts?.Cancel();
            _cts?.Cancel();
            _cts = new CancellationTokenSource();
            ClearDataGrid();

            if (string.IsNullOrEmpty(_filePath)) { UpdateStatus("请先选择Excel文件"); SetProcessingState(false); return; }
            if (comboBoxWorksheets?.SelectedItem == null) { UpdateStatus("请先选择工作表"); SetProcessingState(false); return; }

            try
            {
                string? selectedWorksheet = comboBoxWorksheets.SelectedItem?.ToString();
                if (selectedWorksheet == null) { UpdateStatus("选择的工作表无效"); SetProcessingState(false); return; }
                using var package = new OfficeOpenExcel.ExcelPackage(new FileInfo(_filePath));
                var worksheet = package.Workbook.Worksheets[selectedWorksheet];
                if (worksheet == null) { UpdateStatus("选择的工作表不存在"); SetProcessingState(false); return; }

                await ProcessDataAsync(worksheet, txtKeyword?.Text.Trim() ?? "");
                _queryWatch.Stop();
                txtKeyword?.SelectAll();
            }
            catch (OperationCanceledException) { UpdateStatus("操作已取消"); }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                UpdateStatus($"搜索失败: {ex.Message}");
            }
            finally { SetProcessingState(false); }
        }

        private void ClearDataGrid()
        {
            if (dataGridViewResults == null) return;
            if (!Dispatcher.CheckAccess()) { Dispatcher.Invoke(ClearDataGrid); return; }

            try
            {
                dataGridViewResults.SelectedItem = null;
                dataGridViewResults.UnselectAll();
                if (dataGridViewResults.Columns.Count > 0) dataGridViewResults.Columns.Clear();

                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        if (dataGridViewResults.ItemsSource is System.Data.DataView dataView)
                        {
                            dataGridViewResults.ItemsSource = null;
                            dataView.Dispose();
                        }
                        else dataGridViewResults.ItemsSource = null;
                    }
                    catch (Exception ex) 
                    { 
                        // 忽略清理数据视图时的错误
                        Debug.WriteLine($"清理数据视图时发生错误: {ex.Message}");
                    }
                }), System.Windows.Threading.DispatcherPriority.Send);

                dataGridViewResults.UpdateLayout();
                Task.Delay(100).Wait();

                if (_virtualDataTable != null)
                {
                    // 不要在这里释放ImageTextPair中的图片，因为图片已经在ImageCacheManager中管理
                    // 只需要清除数据，不要释放图片资源
                    _virtualDataTable.Clear();
                    _virtualDataTable.Columns.Clear();
                }
                _pictureMap?.Clear();
            }
            catch (Exception ex)
            {
                // 忽略清理数据网格时的错误，但尝试基本清理
                try { dataGridViewResults.ItemsSource = null; dataGridViewResults.Columns.Clear(); } 
                catch (Exception innerEx) 
                { 
                    Debug.WriteLine($"清理数据网格时发生错误: {innerEx.Message}");
                }
                Debug.WriteLine($"清理数据网格时发生错误: {ex.Message}");
            }
        }

        private async Task ProcessDataAsync(OfficeOpenExcel.ExcelWorksheet worksheet, string keyword)
        {
            if (worksheet == null || worksheet.Dimension == null)
            {
                UpdateStatus(worksheet == null ? "工作表对象为空" : "工作表为空");
                return;
            }

            _loadingCts?.Cancel();
            _loadingCts?.Dispose();
            _loadingCts = new CancellationTokenSource();
            var token = _loadingCts.Token;
            _isLoadingData = true;

            try
            {
                SetProcessingState(true);
                ClearDataGrid();

                _pictureMap ??= new Dictionary<(string, int, int), OfficeOpenExcelDrawing.ExcelPicture?>();
                _pictureMap.Clear();

                _pictureMap = await Task.Run(() => BuildPictureIndex(worksheet, token), token);
                var columns = GetColumnMetadata(worksheet, _pictureMap) ?? new List<ColumnMeta>();

                int totalRows = worksheet.Dimension.Rows;
                if (totalRows <= 1)
                {
                    InitializeDataGridColumns(worksheet, columns);
                    SetupDataTable(columns);
                    UpdateDataGridDisplay();
                    UpdateStatus("✅ 加载完成 - 共 0 条记录");
                    return;
                }

                InitializeDataGridColumns(worksheet, columns);
                SetupDataTable(columns);

                int loadedRows = 0;
                var stopwatch = Stopwatch.StartNew();
                int lastUpdateTime = 0;

                while (loadedRows < totalRows - 1 && !token.IsCancellationRequested)
                {
                    int currentBatchSize = CalculateDynamicBatchSize(totalRows, loadedRows);
                    var batchData = await LoadBatchAsync(worksheet, loadedRows + 2, currentBatchSize, keyword, token);

                    if (batchData.Count > 0)
                    {
                        _virtualDataTable.BeginLoadData();
                        foreach (var rowData in batchData)
                        {
                            if (rowData != null) _virtualDataTable.Rows.Add(rowData);
                        }
                        _virtualDataTable.EndLoadData();
                    }

                    loadedRows += currentBatchSize;

                    if (stopwatch.ElapsedMilliseconds - lastUpdateTime > 200)
                    {
                        UpdateStatus($"加载中: {loadedRows}/{totalRows - 1} 行 ({stopwatch.Elapsed.TotalSeconds:F1}s)");
                        lastUpdateTime = (int)stopwatch.ElapsedMilliseconds;

                        if (GC.GetTotalMemory(false) > GC.GetGCMemoryInfo().TotalAvailableMemoryBytes * 0.7)
                        {
                            GC.Collect(GC.MaxGeneration, GCCollectionMode.Optimized);
                            await Task.Delay(50, token);
                        }
                    }
                }

                if (!token.IsCancellationRequested) ApplyFilter(keyword);
            }
            catch (OperationCanceledException) { UpdateStatus("加载已取消"); }
            catch (OutOfMemoryException)
            {
                _imageCache?.Dispose();
                GC.Collect();
                MessageBox.Show("内存不足，已清除图片缓存", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                UpdateStatus("⚠ 内存不足，部分数据可能未加载");
            }
            catch (Exception ex)
            {
                UpdateStatus($"加载失败: {ex.Message}");
                MessageBox.Show($"错误详情:\n\n错误信息: {ex.Message}\n\n类型: {ex.GetType().Name}", "详细错误信息", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _isLoadingData = false;
                SetProcessingState(false);
            }
        }

        private void SetupDataTable(IEnumerable<ColumnMeta> columns)
        {
            _virtualDataTable.Columns.Clear();
            _virtualDataTable.Rows.Clear();
            _virtualDataTable.Columns.Add("RowNumber", typeof(int));
            foreach (var col in columns)
            {
                _virtualDataTable.Columns.Add(col.Header, col.IsImage && chkShowImages?.IsChecked == true ? typeof(object) : typeof(string));
            }
        }

        private int CalculateDynamicBatchSize(int totalRows, int loadedRows) => Math.Min(BaseBatchSize, totalRows - 1 - loadedRows);

        private async Task<List<object[]>> LoadBatchAsync(OfficeOpenExcel.ExcelWorksheet worksheet, int startRow, int batchSize, string keyword, CancellationToken token)
        {
            try
            {
                var batchData = new List<object[]>(batchSize);
                int endRow = Math.Min(startRow + batchSize - 1, worksheet.Dimension.End.Row);
                bool showImages = chkShowImages?.IsChecked == true;

                var imageColumns = new HashSet<int>();
                if (showImages && _pictureMap != null)
                {
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        if (_pictureMap.Keys.Any(k => k.Item1 == (worksheet.Name ?? "") && k.Item3 == col))
                        {
                            imageColumns.Add(col);
                        }
                    }
                }

                for (int row = startRow; row <= endRow; row++)
                {
                    token.ThrowIfCancellationRequested();
                    if (worksheet.Row(row).Hidden) continue;

                    var rowData = new object[worksheet.Dimension.Columns + 1];
                    rowData[0] = row - 1;
                    bool hasData = false;

                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        try
                        {
                            object? cellValue = null;
                            var worksheetName = worksheet.Name ?? "";

                            if (showImages && imageColumns.Contains(col) &&
                                _pictureMap != null && _pictureMap.ContainsKey((worksheetName, row, col)))
                            {
                                var img = LoadCellImage(worksheet, row, col);
                                var cellText = GetCellTextSafely(worksheet, row, col);

                                if (img != null && IsImageValid(img))
                                {
                                    cellValue = new ImageTextPair { Image = img, Text = cellText ?? string.Empty };
                                    hasData = true;
                                }
                                else cellValue = cellText ?? string.Empty;
                            }
                            else
                            {
                                cellValue = GetCellTextSafely(worksheet, row, col);
                                hasData = hasData || !string.IsNullOrEmpty(cellValue?.ToString() ?? "");
                            }

                            rowData[col] = cellValue ?? string.Empty;
                        }
                        catch (OutOfMemoryException)
                        {
                            rowData[col] = "[内存不足]";
                            _imageCache?.Dispose();
                            GC.Collect();
                        }
                        catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException)
                        {
                            rowData[col] = "[图片错误]";
                        }
                        catch { rowData[col] = "[错误]"; }
                    }

                    if (hasData || !string.IsNullOrWhiteSpace(keyword)) batchData.Add(rowData);

                    if (row % 100 == 0 && GC.GetTotalMemory(false) > 200 * 1024 * 1024)
                    {
                        GC.Collect(GC.MaxGeneration, GCCollectionMode.Optimized);
                    }
                }

                return batchData;
            }
            catch (OperationCanceledException) { return new List<object[]>(0); }
            catch (Exception ex)
            {
                UpdateStatus($"加载错误: {ex.Message}");
                return new List<object[]>(0);
            }
        }

        private string GetCellTextSafely(OfficeOpenExcel.ExcelWorksheet ws, int row, int col)
        {
            try
            {
                var cell = ws.Cells[row, col];
                if (cell.Value is double || cell.Value is int || cell.Value is long) return cell.Value?.ToString() ?? string.Empty;
                var text = cell.Text ?? string.Empty;
                return text == "#VALUE!" ? string.Empty : text;
            }
            catch { return string.Empty; }
        }

        private System.Drawing.Image? LoadCellImage(OfficeOpenExcel.ExcelWorksheet ws, int row, int col)
        {
            try
            {
                if (_pictureMap == null) return null;
                
                var key = (ws.Name ?? "", row, col);
                string cacheKey = $"{ws.Name}_{row}_{col}";

                // 先尝试从缓存获取
                if (_imageCache != null && _imageCache.TryGet(cacheKey, out var cachedImage))
                {
                    return cachedImage;
                }

                // 从pictureMap获取
                if (_pictureMap.TryGetValue(key, out var excelPicture) && excelPicture != null && excelPicture.Image != null)
                {
                    var image = GetImageFromExcelImage(excelPicture.Image);
                    if (image != null && _imageCache != null)
                    {
                        _imageCache.Add(cacheKey, image);
                    }
                    return image;
                }

                // 从单元格获取
                var cellPicture = ws.Cells[row, col].Picture;
                if (cellPicture.Exists)
                {
                    try
                    {
                        var getMethod = cellPicture.GetType().GetMethod("Get");
                        if (getMethod != null)
                        {
                            var excelCellPicture = getMethod.Invoke(cellPicture, null);
                            if (excelCellPicture != null)
                            {
                                var getImageMethod = excelCellPicture.GetType().GetMethod("GetImage");
                                if (getImageMethod != null)
                                {
                                    var excelImage = getImageMethod.Invoke(excelCellPicture, null);
                                    if (excelImage != null)
                                    {
                                        var image = GetImageFromExcelImage(excelImage);
                                        if (image != null && _imageCache != null)
                                        {
                                            _imageCache.Add(cacheKey, image);
                                        }
                                        return image;
                                    }
                                }

                                var getImageBytesMethod = excelCellPicture.GetType().GetMethod("GetImageBytes");
                                if (getImageBytesMethod != null)
                                {
                                    var bytes = getImageBytesMethod.Invoke(excelCellPicture, null) as byte[];
                                    if (bytes != null && bytes.Length > 0)
                                    {
                                        using (var ms = new MemoryStream(bytes))
                                        {
                                            var image = System.Drawing.Image.FromStream(ms);
                                            if (image != null && _imageCache != null)
                                            {
                                                _imageCache.Add(cacheKey, image);
                                            }
                                            return image;
                                        }
                                    }
                                }
                            }
                        }

                        var pictureType = cellPicture.GetType();
                        var imageProperty = pictureType.GetProperty("Image");
                        if (imageProperty != null)
                        {
                            var excelImage = imageProperty.GetValue(cellPicture);
                            if (excelImage != null)
                            {
                                var image = GetImageFromExcelImage(excelImage);
                                if (image != null && _imageCache != null)
                                {
                                    _imageCache.Add(cacheKey, image);
                                }
                                return image;
                            }
                        }

                        var bytesProperty = pictureType.GetProperty("ImageBytes") ?? pictureType.GetProperty("Bytes");
                        if (bytesProperty != null)
                        {
                            var bytes = bytesProperty.GetValue(cellPicture) as byte[];
                            if (bytes != null && bytes.Length > 0)
                            {
                                using (var ms = new MemoryStream(bytes))
                                {
                                    var image = System.Drawing.Image.FromStream(ms);
                                    if (image != null && _imageCache != null)
                                    {
                                        _imageCache.Add(cacheKey, image);
                                    }
                                    return image;
                                }
                            }
                        }
                    }
                    catch (Exception ex) 
                    { 
                        // 忽略加载图片时的反射错误
                        Debug.WriteLine($"通过反射加载图片时发生错误: {ex.Message}");
                    }
                }
                return null;
            }
            catch (Exception ex) 
            { 
                // 忽略加载图片时的错误
                Debug.WriteLine($"加载单元格图片时发生错误: {ex.Message}");
                return null; 
            }
        }

        private System.Drawing.Image? GetImageFromExcelImage(object excelImage)
        {
            try
            {
                var excelImageType = excelImage.GetType();
                var imageBytesProperty = excelImageType.GetProperty("ImageBytes");
                if (imageBytesProperty != null)
                {
                    var bytes = imageBytesProperty.GetValue(excelImage) as byte[];
                    if (bytes != null && bytes.Length > 0)
                    {
                        using (var ms = new MemoryStream(bytes)) return System.Drawing.Image.FromStream(ms);
                    }
                }

                var getImageBytesMethod = excelImageType.GetMethod("get_ImageBytes");
                if (getImageBytesMethod != null)
                {
                    var bytes = getImageBytesMethod.Invoke(excelImage, null) as byte[];
                    if (bytes != null && bytes.Length > 0)
                    {
                        using (var ms = new MemoryStream(bytes)) return System.Drawing.Image.FromStream(ms);
                    }
                }
                return null;
            }
            catch (Exception ex) 
            { 
                // 忽略从Excel图片对象获取图片时的错误
                Debug.WriteLine($"从Excel图片对象获取图片时发生错误: {ex.Message}");
                return null; 
            }
        }

        private static bool IsImageValid(System.Drawing.Image img)
        {
            try { return img != null && img.Width > 0 && img.Height > 0; }
            catch { return false; }
        }

        private void InitializeDataGridColumns(OfficeOpenExcel.ExcelWorksheet worksheet, IEnumerable<ColumnMeta> columns)
        {
            Dispatcher.Invoke(() =>
            {
                if (dataGridViewResults == null) return;
                dataGridViewResults.Columns.Clear();
                bool showImages = chkShowImages?.IsChecked == true;

                foreach (var col in columns)
                {
                    if (col.IsImage && showImages)
                    {
                        var imageCol = new DataGridTemplateColumn
                        {
                            Header = col.Header + " 🖼️",
                            Width = DefaultImageSize,
                            MinWidth = 50
                        };

                        var gridFactory = new FrameworkElementFactory(typeof(Grid));
                        var imageFactory = new FrameworkElementFactory(typeof(System.Windows.Controls.Image));
                        imageFactory.SetValue(System.Windows.Controls.Image.StretchProperty, Stretch.Uniform);
                        imageFactory.SetValue(System.Windows.Controls.Image.StretchDirectionProperty, StretchDirection.Both);
                        imageFactory.SetValue(System.Windows.Controls.Image.HorizontalAlignmentProperty, HorizontalAlignment.Center);
                        imageFactory.SetValue(System.Windows.Controls.Image.VerticalAlignmentProperty, VerticalAlignment.Center);

                        var widthBinding = new Binding("ActualWidth")
                        {
                            RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(DataGridCell), 1),
                            Converter = new CellSizeConverter(),
                            ConverterParameter = DefaultImageSize
                        };
                        imageFactory.SetBinding(System.Windows.Controls.Image.WidthProperty, widthBinding);

                        var heightBinding = new Binding("ActualHeight")
                        {
                            RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(DataGridCell), 1),
                            Converter = new CellSizeConverter(),
                            ConverterParameter = DefaultImageSize
                        };
                        imageFactory.SetBinding(System.Windows.Controls.Image.HeightProperty, heightBinding);

                        var imageBinding = new Binding(col.Header)
                        {
                            Converter = new ImageConverter(),
                            ConverterParameter = col.Header
                        };
                        imageFactory.SetBinding(System.Windows.Controls.Image.SourceProperty, imageBinding);

                        var textFactory = new FrameworkElementFactory(typeof(TextBlock));
                        textFactory.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Left);
                        textFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
                        textFactory.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
                        textFactory.SetValue(TextBlock.TextAlignmentProperty, TextAlignment.Left);
                        textFactory.SetValue(TextBlock.FontSizeProperty, 10.0);
                        textFactory.SetValue(TextBlock.ForegroundProperty, System.Windows.Media.Brushes.Gray);

                        var textBinding = new Binding(col.Header)
                        {
                            Converter = new ImageTextConverter(),
                            ConverterParameter = col.Header
                        };
                        textFactory.SetBinding(TextBlock.TextProperty, textBinding);

                        var imageVisibilityBinding = new Binding(col.Header)
                        {
                            Converter = new ImageVisibilityConverter()
                        };
                        imageFactory.SetBinding(VisibilityProperty, imageVisibilityBinding);

                        // 文本总是显示，但当有图片时，图片会覆盖文本
                        // 所以这里不需要额外的可见性控制

                        gridFactory.AppendChild(imageFactory);
                        gridFactory.AppendChild(textFactory);

                        var template = new DataTemplate();
                        template.VisualTree = gridFactory;
                        imageCol.CellTemplate = template;
                        dataGridViewResults.Columns.Add(imageCol);
                    }
                    else
                    {
                        var textCol = new DataGridTextColumn
                        {
                            Header = col.Header,
                            Binding = new Binding(col.Header),
                            Width = ColumnWidthMax
                        };

                        var textStyle = new Style(typeof(TextBlock));
                        textStyle.Setters.Add(new Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap));
                        textStyle.Setters.Add(new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center));
                        textStyle.Setters.Add(new Setter(TextBlock.TextAlignmentProperty, TextAlignment.Left));
                        textStyle.Setters.Add(new Setter(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Stretch));
                        textCol.ElementStyle = textStyle;
                        dataGridViewResults.Columns.Add(textCol);
                    }
                }

                if (dataGridViewResults.ItemsSource != null) dataGridViewResults.ItemsSource = null;
            });
        }

        private void UpdateDataGridDisplay()
        {
            if (dataGridViewResults == null) return;
            if (!Dispatcher.CheckAccess()) { Dispatcher.Invoke(UpdateDataGridDisplay); return; }
            dataGridViewResults.ItemsSource = _virtualDataTable.DefaultView;
            ApplyAutoRowHeight();
        }

        private Dictionary<(string, int, int), OfficeOpenExcelDrawing.ExcelPicture?> BuildPictureIndex(OfficeOpenExcel.ExcelWorksheet ws, CancellationToken ct)
        {
            var dict = new Dictionary<(string, int, int), OfficeOpenExcelDrawing.ExcelPicture?>();
            foreach (var pic in ws.Drawings.OfType<OfficeOpenExcelDrawing.ExcelPicture>())
            {
                ct.ThrowIfCancellationRequested();
                var key = (ws.Name ?? "", pic.From.Row + 1, pic.From.Column + 1);
                dict.TryAdd(key, pic);
            }

            var dimension = ws.Dimension;
            if (dimension != null)
            {
                for (int row = dimension.Start.Row; row <= dimension.End.Row; row++)
                {
                    for (int col = dimension.Start.Column; col <= dimension.End.Column; col++)
                    {
                        ct.ThrowIfCancellationRequested();
                        if (ws.Cells[row, col].Picture.Exists)
                        {
                            var key = (ws.Name ?? "", row, col);
                            if (!dict.ContainsKey(key)) dict[key] = null;
                        }
                    }
                }
            }
            return dict;
        }

        private List<ColumnMeta> GetColumnMetadata(OfficeOpenExcel.ExcelWorksheet ws, Dictionary<(string, int, int), OfficeOpenExcelDrawing.ExcelPicture?> pictureMap)
        {
            var columns = new List<ColumnMeta>();
            var headerCounts = new Dictionary<string, int>();

            if (ws == null || ws.Dimension == null) return columns;

            for (int col = 1; col <= ws.Dimension.Columns; col++)
            {
                string originalHeader = ws.Cells[1, col]?.Text ?? "";
                if (string.IsNullOrWhiteSpace(originalHeader)) originalHeader = $"<列{col}>";

                if (!headerCounts.ContainsKey(originalHeader)) headerCounts[originalHeader] = 1;
                else headerCounts[originalHeader]++;

                string finalHeader = headerCounts[originalHeader] > 1 ? $"{originalHeader}_{headerCounts[originalHeader]}" : originalHeader;

                bool hasImage = false;
                for (int row = 2; row <= ws.Dimension.Rows; row++)
                {
                    var key = (ws.Name ?? "", row, col);
                    if (pictureMap.ContainsKey(key)) { hasImage = true; break; }
                }

                columns.Add(new ColumnMeta(finalHeader, hasImage));
            }
            return columns;
        }

        private void SetProcessingState(bool isProcessing)
        {
            Dispatcher.Invoke(() =>
            {
                _isProcessing = isProcessing;
                if (btnOpenFile != null) btnOpenFile.IsEnabled = !isProcessing;
                if (btnSearch != null) btnSearch.IsEnabled = !isProcessing;
                if (comboBoxWorksheets != null) comboBoxWorksheets.IsEnabled = !isProcessing;
                if (chkShowImages != null) chkShowImages.IsEnabled = !isProcessing;
                this.Cursor = isProcessing ? Cursors.Wait : Cursors.Arrow;
            });
        }

        private void UpdateStatus(string message)
        {
            Dispatcher.Invoke(() =>
            {
                if (toolStripStatusLabel != null) toolStripStatusLabel.Text = message;
            });
        }

        private void ApplyFilter(string keyword)
        {
            if (_virtualDataTable == null || _virtualDataTable.Rows.Count == 0)
            {
                UpdateDataGridDisplay();
                return;
            }

            try
            {
                var filteredRows = new List<object[]>();
                bool hasKeyword = !string.IsNullOrWhiteSpace(keyword);

                if (!hasKeyword)
                {
                    foreach (DataRow row in _virtualDataTable.Rows)
                    {
                        var rowData = new object[_virtualDataTable.Columns.Count];
                        for (int i = 0; i < _virtualDataTable.Columns.Count; i++)
                        {
                            if (_virtualDataTable.Columns[i].ColumnName == "RowNumber")
                                rowData[i] = filteredRows.Count + 1;
                            else
                                rowData[i] = row[i];
                        }
                        filteredRows.Add(rowData);
                    }
                }
                else
                {
                    var searchConditions = ParseSearchConditions(keyword);
                    foreach (DataRow row in _virtualDataTable.Rows)
                    {
                        bool match = EvaluateConditions(row, searchConditions);
                        if (match)
                        {
                            var rowData = new object[_virtualDataTable.Columns.Count];
                            for (int i = 0; i < _virtualDataTable.Columns.Count; i++)
                            {
                                if (_virtualDataTable.Columns[i].ColumnName == "RowNumber")
                                    rowData[i] = filteredRows.Count + 1;
                                else
                                    rowData[i] = row[i];
                            }
                            filteredRows.Add(rowData);
                        }
                    }
                }

                var filteredTable = _virtualDataTable.Clone();
                foreach (var rowData in filteredRows)
                {
                    filteredTable.Rows.Add(rowData);
                }

                if (!Dispatcher.CheckAccess())
                {
                    Dispatcher.Invoke(() => dataGridViewResults.ItemsSource = filteredTable.DefaultView);
                }
                else dataGridViewResults.ItemsSource = filteredTable.DefaultView;

                if (_queryWatch != null)
                    UpdateStatus($"✅ 搜索完成 - 共 {filteredRows.Count} 条记录，查询耗时 {_queryWatch.Elapsed.TotalSeconds:F3}秒");
                else
                    UpdateStatus($"✅ 搜索完成 - 共 {filteredRows.Count} 条记录");
                return;
            }
            catch (Exception ex)
            {
                UpdateStatus($"筛选错误: {ex.Message}");
            }

            UpdateDataGridDisplay();
        }

        private class SearchCondition
        {
            public List<string> AndTerms { get; set; } = new();
            public List<string> OrTerms { get; set; } = new();
            public List<string> NotTerms { get; set; } = new();
        }

        private SearchCondition ParseSearchConditions(string keyword)
        {
            var condition = new SearchCondition();
            // 首先规范化标点符号：将中文标点转换为英文标点
            string normalizedKeyword = NormalizePunctuation(keyword);
            var parts = normalizedKeyword.Split(new[] { ' ', '+' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var part in parts)
            {
                if (part.StartsWith("!"))
                {
                    string term = part.Substring(1).Trim();
                    if (!string.IsNullOrEmpty(term)) condition.NotTerms.Add(term);
                }
                else if (part.Contains(",") || part.Contains(";"))
                {
                    var orTerms = part.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var orTerm in orTerms)
                    {
                        string trimmed = orTerm.Trim();
                        if (!string.IsNullOrEmpty(trimmed)) condition.OrTerms.Add(trimmed);
                    }
                }
                else condition.AndTerms.Add(part.Trim());
            }

            return condition;
        }

        private string NormalizePunctuation(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            
            // 将中文标点转换为英文标点
            string normalized = input
                .Replace('，', ',')   // 中文逗号 -> 英文逗号
                .Replace('；', ';')   // 中文分号 -> 英文分号  
                .Replace('！', '!')   // 中文感叹号 -> 英文感叹号
                .Replace('　', ' ');  // 中文全角空格 -> 英文空格
            
            // 移除逗号和分号周围的空格，使"甲 , 乙"等价于"甲,乙"
            // 处理逗号
            normalized = System.Text.RegularExpressions.Regex.Replace(normalized, @"\s*,\s*", ",");
            // 处理分号
            normalized = System.Text.RegularExpressions.Regex.Replace(normalized, @"\s*;\s*", ";");
            
            // 移除感叹号前后的空格，使" ! 甲"等价于"!甲"
            // 匹配：零个或多个空格，感叹号，零个或多个空格
            // 替换为：感叹号（无空格）
            normalized = System.Text.RegularExpressions.Regex.Replace(normalized, @"\s*!\s*", "!");
            
            return normalized;
        }

        private bool EvaluateConditions(DataRow row, SearchCondition condition)
        {
            if (condition.NotTerms.Count > 0)
            {
                foreach (var notTerm in condition.NotTerms)
                {
                    if (RowContainsTerm(row, notTerm)) return false;
                }
            }

            if (condition.AndTerms.Count > 0)
            {
                foreach (var andTerm in condition.AndTerms)
                {
                    if (!RowContainsTerm(row, andTerm)) return false;
                }
            }

            if (condition.OrTerms.Count > 0)
            {
                bool orMatch = false;
                foreach (var orTerm in condition.OrTerms)
                {
                    if (RowContainsTerm(row, orTerm)) { orMatch = true; break; }
                }
                if (!orMatch && (condition.AndTerms.Count > 0 || condition.NotTerms.Count > 0)) return false;
                if (!orMatch && condition.AndTerms.Count == 0 && condition.NotTerms.Count == 0) return false;
            }

            return true;
        }

        private bool RowContainsTerm(DataRow row, string term)
        {
            foreach (DataColumn col in _virtualDataTable.Columns)
            {
                if (col.ColumnName == "RowNumber") continue;
                object? value = row[col];
                if (value != null && value.ToString()?.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            return false;
        }

        private void ComboBoxWorksheets_SelectedIndexChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_isProcessing && !_isLoadingData && !string.IsNullOrEmpty(_filePath) && comboBoxWorksheets?.SelectedIndex != -1)
                BtnSearch_Click(sender, e);
        }

        private void ChkShowImages_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if (!_isProcessing && !_isLoadingData && !string.IsNullOrEmpty(_filePath) && comboBoxWorksheets?.SelectedIndex != -1)
            {
                // 重新加载当前工作表数据
                ReloadCurrentWorksheet();
            }
        }

        private async void ReloadCurrentWorksheet()
        {
            if (string.IsNullOrEmpty(_filePath) || comboBoxWorksheets?.SelectedItem == null) return;

            try
            {
                SetProcessingState(true);
                UpdateStatus("正在重新加载图片数据...");

                string? selectedWorksheet = comboBoxWorksheets.SelectedItem?.ToString();
                if (selectedWorksheet == null) return;

                // 保存当前的搜索关键词
                string currentKeyword = txtKeyword?.Text.Trim() ?? "";

                // 重新打开文件并处理
                using var package = new OfficeOpenExcel.ExcelPackage(new FileInfo(_filePath));
                var worksheet = package.Workbook.Worksheets[selectedWorksheet];
                if (worksheet == null) return;

                // 清除图片缓存，强制重新加载图片
                _imageCache?.Dispose();
                // 重新创建缓存管理器
                _imageCache = new ImageCacheManager(200, TimeSpan.FromMinutes(5));
                
                // 重新处理数据
                await ProcessDataAsync(worksheet, currentKeyword);

                UpdateStatus($"✅ 图片显示已更新 - 共 {_virtualDataTable.Rows.Count} 条记录");
            }
            catch (Exception ex)
            {
                UpdateStatus($"重新加载失败: {ex.Message}");
                MessageBox.Show($"重新加载失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetProcessingState(false);
            }
        }

        private void TxtKeyword_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                e.Handled = true;
                BtnSearch_Click(sender, e);
            }
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
                    if (dataGridViewResults != null)
                    {
                        dataGridViewResults.ItemsSource = null;
                        dataGridViewResults.Columns.Clear();
                    }
                    await ProcessFileSelectionAsync(files[0]);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"拖放文件错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        
        private void HandleDragEnter(DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && e.Data.GetData(DataFormats.FileDrop) is string[] files)
            {
                if (files.Length == 1 && IsExcelFile(files[0]))
                {
                    e.Effects = DragDropEffects.Copy;
                    return;
                }
            }
            e.Effects = DragDropEffects.None;
        }

        private bool IsExcelFile(string path)
        {
            string ext = Path.GetExtension(path).ToLower();
            return ext == ".xlsx" || ext == ".xls";
        }

        private void DataGrid_CleanUpVirtualizedItem(object sender, CleanUpVirtualizedItemEventArgs e)
        {
            // 不要在这里释放ImageTextPair中的图片，因为图片已经在ImageCacheManager中管理
            // 虚拟化清理时，WPF会自动处理UI元素的清理，但数据对象应该保持不变
            // 这样可以避免翻页时图片消失的问题
        }

        private void ApplyAutoRowHeight()
        {
            if (dataGridViewResults == null || dataGridViewResults.Items.Count == 0) return;
            
            Dispatcher.Invoke(() =>
            {
                try
                {
                    // 简单的自动行高逻辑：根据内容自动调整
                    foreach (var item in dataGridViewResults.Items)
                    {
                        var row = dataGridViewResults.ItemContainerGenerator.ContainerFromItem(item) as DataGridRow;
                        if (row != null)
                        {
                            // 只设置最小行高，不限制最大行高，允许用户手动调整
                            row.MinHeight = RowHeightMin;
                            
                            // 如果行中有多行文本，适当增加行高
                            var cell = dataGridViewResults.Columns[0].GetCellContent(item) as TextBlock;
                            if (cell != null && cell.Text != null)
                            {
                                var lineCount = cell.Text.Split('\n').Length;
                                if (lineCount > 1)
                                {
                                    // 根据行数自动调整行高，但不限制最大高度
                                    row.Height = RowHeightMin + (lineCount - 1) * 15;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // 忽略调整行高时的错误
                }
            });
        }

        private void DataGridViewResults_Loaded(object sender, RoutedEventArgs e) => ApplyAutoRowHeight();

        private void DataGridViewResults_LayoutUpdated(object sender, EventArgs e)
        {
            if (dataGridViewResults.Items.Count > 0 && dataGridViewResults.RowHeight == RowHeightMin) ApplyAutoRowHeight();
        }

        public class ImageTextPair : IDisposable
        {
            public System.Drawing.Image? Image { get; set; }
            public string? Text { get; set; }
            public void Dispose() => Image?.Dispose();
        }

        public class ColumnMeta
        {
            public string Header { get; }
            public bool IsImage { get; }
            public ColumnMeta(string header, bool isImage) { Header = header; IsImage = isImage; }
        }

        public class ImageCacheManager : IDisposable
        {
            private readonly int _capacity;
            private readonly ConcurrentDictionary<string, (System.Drawing.Image? Value, DateTime LastAccess)> _cache;
            private readonly System.Threading.Timer _cleanupTimer;
            private readonly object _lock = new object();

            public ImageCacheManager(int capacity, TimeSpan cleanupInterval)
            {
                _capacity = capacity;
                _cache = new ConcurrentDictionary<string, (System.Drawing.Image?, DateTime)>();
                _cleanupTimer = new System.Threading.Timer(Cleanup, null, cleanupInterval, cleanupInterval);
            }

            public bool TryGet(string key, out System.Drawing.Image? value)
            {
                if (_cache.TryGetValue(key, out var entry))
                {
                    lock (_lock) { _cache[key] = (entry.Value, DateTime.UtcNow); }
                    value = entry.Value;
                    return true;
                }
                value = null;
                return false;
            }

            public void Add(string key, System.Drawing.Image? value)
            {
                lock (_lock) { _cache[key] = (value, DateTime.UtcNow); }
                if (_capacity > 0 && _cache.Count > _capacity * 1.2) Cleanup(null);
            }

            private void Cleanup(object? state)
            {
                lock (_lock)
                {
                    if (_cache.Count <= _capacity) return;
                    var toRemove = _cache.OrderBy(x => x.Value.LastAccess).Take(_cache.Count - _capacity).ToList();
                    foreach (var item in toRemove)
                    {
                        if (_cache.TryRemove(item.Key, out var entry)) entry.Value?.Dispose();
                    }
                }
            }

            public void Dispose()
            {
                _cleanupTimer?.Dispose();
                lock (_lock)
                {
                    foreach (var item in _cache.Values) item.Value?.Dispose();
                    _cache.Clear();
                }
                GC.SuppressFinalize(this);
            }
        }

        public class ImageConverter : IValueConverter
        {
            public object Convert(object? value, Type targetType, object? parameter, System.Globalization.CultureInfo culture)
            {
                if (value == null) return null!;
                try
                {
                    if (value is ImageTextPair pair)
                    {
                        if (pair.Image != null)
                        {
                            try
                            {
                                using var ms = new MemoryStream();
                                try { pair.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png); }
                                catch
                                {
                                    pair.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                                }
                                ms.Position = 0;
                                var bitmap = new BitmapImage();
                                bitmap.BeginInit();
                                bitmap.StreamSource = ms;
                                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                                bitmap.EndInit();
                                bitmap.Freeze();
                                return bitmap;
                            }
                            catch
                            {
                                return null!;
                            }
                        }
                        else { return null!; }
                    }
                    else { return null!; }
                }
                catch { return null!; }
            }

            public object ConvertBack(object? value, Type targetType, object? parameter, System.Globalization.CultureInfo culture) => throw new NotImplementedException();
        }

        public class ImageTextConverter : IValueConverter
        {
            public object Convert(object? value, Type targetType, object? parameter, System.Globalization.CultureInfo culture)
            {
                if (value == null) return null!;
                try
                {
                    if (value is ImageTextPair pair) return pair.Text ?? "🖼️ 图片";
                    return value.ToString()!;
                }
                catch { return null!; }
            }

            public object ConvertBack(object? value, Type targetType, object? parameter, System.Globalization.CultureInfo culture) => throw new NotImplementedException();
        }

        public class ImageVisibilityConverter : IValueConverter
        {
            public object Convert(object? value, Type targetType, object? parameter, System.Globalization.CultureInfo culture)
            {
                if (value == null) return Visibility.Collapsed;
                try { return value is ImageTextPair pair && pair.Image != null ? Visibility.Visible : Visibility.Collapsed; }
                catch { return Visibility.Collapsed; }
            }

            public object ConvertBack(object? value, Type targetType, object? parameter, System.Globalization.CultureInfo culture) => throw new NotImplementedException();
        }

        public class CellSizeConverter : IValueConverter
        {
            public object Convert(object? value, Type targetType, object? parameter, System.Globalization.CultureInfo culture)
            {
                // 解析parameter作为默认大小
                double defaultSize = 150.0;
                if (parameter is double paramSize && paramSize > 0)
                {
                    defaultSize = paramSize;
                }
                
                if (value is double cellSize && cellSize > 0)
                {
                    double margin = 10.0;
                    double minSize = 30.0;
                    
                    return Math.Max(minSize, cellSize - margin);
                }
                
                // 如果没有有效的cellSize，返回默认大小
                return defaultSize;
            }

            public object ConvertBack(object? value, Type targetType, object? parameter, System.Globalization.CultureInfo culture) => throw new NotImplementedException();
        }
    }
}

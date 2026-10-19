using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace ExcelQueryTool.Services
{
    public class UIStateManager
    {
        private readonly Button? _btnOpenFile;
        private readonly Button? _btnSearch;
        private readonly ComboBox? _comboBoxWorksheets;
        private readonly CheckBox? _chkShowImages;
        private readonly TextBlock? _statusLabel;
        private readonly Window _window;
        private bool _isProcessing;
        private bool _isLoadingData;
        // 当前文件是 txt/csv（无图片列、无多表）：完成态启用控件时据此保持灰掉——
        // 否则 SetProcessingState(false) 的无条件启用会把文本文件无意义的功能控件点亮
        private bool _currentFileIsText;
        private static readonly Thickness _defaultMargin = new(2);
        private static readonly Style? _sharedTextStyle;
        private static readonly Style? _numericTextStyle;
        private static readonly Style? _headerStyle;

        /// <summary>
        /// 全表统一字号（磅）。此前普通列继承 DataGrid 默认（12）而图片列写死 10，
        /// 同一张表里两种字号并存 —— 这是「字号不统一」的根因。
        /// 现统一到 12，不读 Excel 样式（Excel 里常有局部加粗/合并残留造成的杂态字号）。
        /// </summary>
        private const double UnifiedFontSize = 12.0;

        static UIStateManager()
        {
            _sharedTextStyle = BuildTextStyle(HorizontalAlignment.Left);

            // 数字列：右对齐（Excel「常规」格式的惯例）。等宽数字串右对齐后，
            // 数量级差异一眼可见（左对齐时 9 与 100000 的个位不对齐，难比较）。
            _numericTextStyle = BuildTextStyle(HorizontalAlignment.Right);

            _headerStyle = new Style(typeof(System.Windows.Controls.Primitives.DataGridColumnHeader))
            {
                Setters =
                {
                    new Setter(Control.FontSizeProperty, UnifiedFontSize),
                    new Setter(Control.HorizontalContentAlignmentProperty, HorizontalAlignment.Center),
                    new Setter(Control.VerticalContentAlignmentProperty, VerticalAlignment.Center),
                }
            };
        }

        /// <summary>构造单元格文本样式。全列统一字号，只有水平对齐随列类型变化。</summary>
        private static Style BuildTextStyle(HorizontalAlignment align) => new(typeof(TextBlock))
        {
            Setters =
            {
                new Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap),
                new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center),
                new Setter(TextBlock.TextAlignmentProperty,
                    align == HorizontalAlignment.Right ? TextAlignment.Right : TextAlignment.Left),
                new Setter(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Stretch),
                new Setter(TextBlock.FontSizeProperty, UnifiedFontSize),
                new Setter(TextBlock.ForegroundProperty, new SolidColorBrush(Color.FromRgb(0x1F, 0x1F, 0x1F)))
            }
        };

        public bool IsProcessing => _isProcessing;
        public bool IsLoadingData => _isLoadingData;

        public UIStateManager(
            Window window,
            Button? btnOpenFile = null,
            Button? btnSearch = null,
            ComboBox? comboBoxWorksheets = null,
            CheckBox? chkShowImages = null,
            TextBlock? statusLabel = null)
        {
            _window = window;
            _btnOpenFile = btnOpenFile;
            _btnSearch = btnSearch;
            _comboBoxWorksheets = comboBoxWorksheets;
            _chkShowImages = chkShowImages;
            _statusLabel = statusLabel;
        }

        /// <summary>标记当前文件能力：txt/csv 无图片列与多表，「显示图片」勾选与工作表下拉在非处理态也保持禁用</summary>
        public void SetCurrentFileIsText(bool isTextFile) => _currentFileIsText = isTextFile;

        public void SetProcessingState(bool isProcessing)
        {
            _isProcessing = isProcessing;
            _window.Dispatcher.Invoke(() =>
            {
                if (_btnOpenFile != null) _btnOpenFile.IsEnabled = !isProcessing;
                if (_btnSearch != null) _btnSearch.IsEnabled = !isProcessing;
                if (_comboBoxWorksheets != null) _comboBoxWorksheets.IsEnabled = !isProcessing && !_currentFileIsText;
                if (_chkShowImages != null) _chkShowImages.IsEnabled = !isProcessing && !_currentFileIsText;
                _window.Cursor = isProcessing ? Cursors.Wait : Cursors.Arrow;
            });
        }

        public void SetLoadingDataState(bool isLoadingData)
        {
            _isLoadingData = isLoadingData;
        }

        public void UpdateStatus(string message)
        {
            _window.Dispatcher.Invoke(() =>
            {
                if (_statusLabel != null) _statusLabel.Text = message;
            });
        }

        public void UpdateWorksheetList(System.Collections.ObjectModel.ObservableCollection<string> worksheets)
        {
            _window.Dispatcher.Invoke(() =>
            {
                if (_comboBoxWorksheets != null)
                {
                    _comboBoxWorksheets.ItemsSource = null;
                    _comboBoxWorksheets.ItemsSource = worksheets;
                    if (worksheets.Count > 0) _comboBoxWorksheets.SelectedIndex = 0;
                }
            });
        }

        public bool IsShowImagesChecked()
        {
            return _chkShowImages?.IsChecked == true;
        }

        /// <summary>
        /// 用户调整过的列宽（键 = 列 Header）。
        /// 快照点必须在 <see cref="ClearDataGrid"/> 里 —— 调用方清完列才调
        /// <see cref="InitializeDataGridColumns"/>，若到那时再快照，列早已被清空、
        /// 永远拿不到任何宽度（曾经的时序 bug：保留逻辑写了但从不生效）。
        /// </summary>
        private readonly Dictionary<string, double> _savedColumnWidths = new(StringComparer.Ordinal);

        public void ClearDataGrid(DataGrid? dataGrid, bool preserveColumnWidths = true)
        {
            _window.Dispatcher.Invoke(() =>
            {
                if (dataGrid == null) return;

                if (preserveColumnWidths)
                {
                    // 记住用户拖出的固定宽度：重载/搜索后按 Header 恢复，不再重置回初始值
                    foreach (var existing in dataGrid.Columns)
                    {
                        if (existing.Header is string h && existing.Width.UnitType == DataGridLengthUnitType.Pixel)
                            _savedColumnWidths[h] = existing.Width.Value;
                    }
                }
                else
                {
                    // 换文件：新旧文件列结构无关，按 Header 撞名继承宽度是错的，记忆全部作废
                    _savedColumnWidths.Clear();
                    _userRowHeights.Clear();
                }

                dataGrid.ItemsSource = null;
                dataGrid.Columns.Clear();
            });
        }

        public void UpdateDataGrid(DataGrid? dataGrid, RowList rows)
        {
            _window.Dispatcher.Invoke(() =>
            {
                if (dataGrid != null)
                {
                    dataGrid.ItemsSource = rows;
                }
            });
        }

        public void InitializeDataGridColumns(
            DataGrid? dataGrid,
            System.Collections.Generic.IEnumerable<ColumnMeta> columns,
            int defaultImageSize = 150,
            int columnWidthMax = 120)
        {
            _window.Dispatcher.Invoke(() =>
            {
                if (dataGrid == null) return;

                // 表头统一样式（字号与数据行一致、水平居中）
                if (dataGrid.ColumnHeaderStyle == null && _headerStyle != null)
                    dataGrid.ColumnHeaderStyle = _headerStyle;

                dataGrid.Columns.Clear();

                // 用户拖过的固定宽度优先于程序自动计算值（快照来自 ClearDataGrid，见字段注释）
                foreach (var col in columns)
                {
                    if (col.IsImage)
                    {
                        var imageCol = new DataGridTemplateColumn
                        {
                            Header = col.Header,
                            // 用户之前拖过该列：恢复其宽度，否则回到默认尺寸
                            Width = _savedColumnWidths.TryGetValue(col.Header, out var iw) ? iw : defaultImageSize,
                            MinWidth = _savedColumnWidths.TryGetValue(col.Header, out var imw) ? Math.Max(50, imw) : 50
                        };

                        var stackPanelFactory = new FrameworkElementFactory(typeof(StackPanel));
                        stackPanelFactory.SetValue(StackPanel.OrientationProperty, Orientation.Vertical);

                        // 图片区：ItemsControl 逐张渲染 —— 同一格锚定多张图时从上到下堆叠
                        // （E8 类场景；单图时只有一项，视觉与旧版一致）
                        var imagesListFactory = new FrameworkElementFactory(typeof(ItemsControl));
                        var imageStackFactory = new FrameworkElementFactory(typeof(StackPanel));
                        imageStackFactory.SetValue(StackPanel.OrientationProperty, Orientation.Vertical);
                        imagesListFactory.SetValue(ItemsControl.ItemsPanelProperty, new ItemsPanelTemplate(imageStackFactory));

                        var imageFactory = new FrameworkElementFactory(typeof(System.Windows.Controls.Image));
                        imageFactory.SetValue(System.Windows.Controls.Image.StretchProperty, Stretch.Uniform);
                        imageFactory.SetValue(System.Windows.Controls.Image.HorizontalAlignmentProperty, HorizontalAlignment.Center);
                        imageFactory.SetValue(System.Windows.Controls.Image.VerticalAlignmentProperty, VerticalAlignment.Center);
                        imageFactory.SetValue(System.Windows.Controls.Image.MarginProperty, _defaultMargin);

                        var widthBinding = new Binding("ActualWidth")
                        {
                            RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(DataGridCell), 1),
                            Converter = new CellSizeConverter(),
                            ConverterParameter = defaultImageSize
                        };
                        imageFactory.SetBinding(System.Windows.Controls.Image.WidthProperty, widthBinding);

                        // 不绑定高度：宽度跟随列宽 + Stretch.Uniform 按比例自适应，
                        // 行高（NaN 自动）再由图片实际高度撑开 —— 拉大列宽/行高图片随之缩放。
                        // 之前高度绑定单元格 ActualHeight 会在「行高由内容撑开」时形成
                        // 行↔图互相放大的正反馈循环，是行高失控的根源，故彻底移除。

                        // 当前项（GDI Image，来自 ImageTextPair.Images）→ WPF BitmapImage
                        var sourceBinding = new Binding { Converter = new ImageSourceConverter() };
                        imageFactory.SetBinding(System.Windows.Controls.Image.SourceProperty, sourceBinding);
                        imagesListFactory.SetValue(ItemsControl.ItemTemplateProperty, new DataTemplate { VisualTree = imageFactory });

                        var imagesBinding = CreateRowBinding(col.Header);
                        imagesBinding.Converter = new CellImagesConverter();
                        imagesListFactory.SetBinding(ItemsControl.ItemsSourceProperty, imagesBinding);

                        var textFactory = new FrameworkElementFactory(typeof(TextBlock));
                        textFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
                        textFactory.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
                        textFactory.SetValue(TextBlock.TextAlignmentProperty, TextAlignment.Left);
                        textFactory.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Stretch);
                        // 与普通列同一字号：此前写死 10.0，与普通列的 12 并存 → 同表两种字号
                        textFactory.SetValue(TextBlock.FontSizeProperty, UnifiedFontSize);
                        textFactory.SetValue(TextBlock.ForegroundProperty, new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0x1F, 0x1F, 0x1F)));
                        textFactory.SetValue(TextBlock.MarginProperty, _defaultMargin);

                        var textBinding = CreateRowBinding(col.Header);
                        textBinding.Converter = new ImageTextConverter();
                        textBinding.ConverterParameter = col.Header;
                        textFactory.SetBinding(TextBlock.TextProperty, textBinding);

                        // 命中词高亮：WatchText 与 Text 同源（同一 Binding 实例复用），词表全局
                        textFactory.SetBinding(SearchHighlight.WatchTextProperty, textBinding);
                        textFactory.SetBinding(SearchHighlight.WatchTermsProperty,
                            new Binding(nameof(SearchHighlight.Terms)) { Source = SearchHighlight.Instance });

                        // 无图整块折叠（挂在 ItemsControl 上；此前挂在单个 Image 上）
                        var imageVisibilityBinding = CreateRowBinding(col.Header);
                        imageVisibilityBinding.Converter = new ImageVisibilityConverter();
                        imagesListFactory.SetBinding(UIElement.VisibilityProperty, imageVisibilityBinding);

                        stackPanelFactory.AppendChild(imagesListFactory);
                        stackPanelFactory.AppendChild(textFactory);

                        var template = new DataTemplate();
                        template.VisualTree = stackPanelFactory;
                        imageCol.CellTemplate = template;
                        dataGrid.Columns.Add(imageCol);
                    }
                    else
                    {
                        // 按 Excel 原始列宽展示：字符宽度 → 像素（约 7px/字符 + 5px 补白）。
                        // ExcelWidth=0（未声明）保持原默认。上限 3 倍默认宽度，防止单列过宽独占界面。
                        double width = columnWidthMax;
                        if (col.ExcelWidth > 0)
                        {
                            double px = Math.Ceiling(col.ExcelWidth * 7 + 5);
                            width = Math.Max(30, Math.Min(px, columnWidthMax * 3));
                        }

                        // 命中词高亮挂在 ElementStyle 的附加属性上：基样式（对齐/字号）BasedOn 保留，
                        // WatchText 绑列值、WatchTerms 绑全局词表，任一变化重建 Inlines；
                        // 词表为空或本格无命中时不动容器（Text 与 Inlines 同一存储，
                        // 清空后内容无法恢复——2026-09-24 过滤空白的教训），显示行为不变
                        var cellStyle = new Style(typeof(TextBlock), col.IsNumeric ? _numericTextStyle : _sharedTextStyle);
                        cellStyle.Setters.Add(new Setter(SearchHighlight.WatchTextProperty, CreateRowBinding(col.Header)));
                        cellStyle.Setters.Add(new Setter(SearchHighlight.WatchTermsProperty,
                            new Binding(nameof(SearchHighlight.Terms)) { Source = SearchHighlight.Instance }));

                        var textCol = new DataGridTextColumn
                        {
                            Header = col.Header,
                            Binding = CreateRowBinding(col.Header),
                            // 用户之前拖过该列：恢复其宽度，否则用程序算的初始宽度
                            Width = _savedColumnWidths.TryGetValue(col.Header, out var tw) ? tw : width,
                            // 兜底：后续自动调整也不低于用户拖出的宽度
                            MinWidth = _savedColumnWidths.TryGetValue(col.Header, out var tmw) ? Math.Max(30, tmw) : 30,
                            // 列级对齐：数字列右对齐、文本列左对齐（统一字号在样式里）
                            ElementStyle = cellStyle
                        };
                        dataGrid.Columns.Add(textCol);
                    }
                }
            });
        }

        /// <summary>
        /// 为行项（<see cref="RowView"/>）构造绑定路径：统一走字符串索引器 <c>[header]</c>。
        ///
        /// RowView 只有索引器与 RowNumber 属性，裸路径解析不到列 —— 与旧 DataView 时代
        /// 「正常表头走属性名、特殊表头走索引器」的双轨不同，现在全部走索引器。
        /// 索引器内容不受路径特殊字符影响，表头里的 <c>]</c> <c>^</c> <c>[</c>
        /// 按 WPF 路径规则用 <c>^</c> 转义（202509-客诉.xlsx 第 5 列图片不显示的教训延续）。
        /// </summary>
        private static Binding CreateRowBinding(string header)
        {
            var sb = new System.Text.StringBuilder(header.Length + 2);
            sb.Append('[');
            foreach (char c in header)
            {
                if (c is ']' or '^' or '[')
                    sb.Append('^');
                sb.Append(c);
            }
            sb.Append(']');
            return new Binding(sb.ToString());
        }

        /// <summary>
        /// 切换「按内容自适应行高」。
        ///
        /// 之所以不是「一次性遍历全部行设置高度」：DataGrid 开启了行虚拟化，
        /// 视口外的行没有容器（<c>ContainerFromItem</c> 返回 null），遍历百万行只能改到可见的十几行，
        /// 白费百万次字典查找，且滚动到新行时高度不会重算。
        /// 
        /// 正确做法是挂 <see cref="DataGrid.LoadingRow"/> —— 容器生成时按需计算，
        /// 滚动到时自然生效，代价与行数无关。
        /// </summary>
        public void AttachAutoRowHeight(DataGrid? dataGrid, int rowHeightMin = 30)
        {
            if (dataGrid == null) return;

            _autoRowHeightMin = rowHeightMin;
            dataGrid.LoadingRow -= OnLoadingRow;
            dataGrid.LoadingRow += OnLoadingRow;
            dataGrid.UnloadingRow -= OnUnloadingRow;
            dataGrid.UnloadingRow += OnUnloadingRow;
            // 行高拖拽记忆的捕获点在 OnLoadingRow 里挂（Thumb 只在行容器加载时存在）
        }

        /// <summary>清空用户行高记忆（切换工作表时调用：RowNumber 从 1 重排，跨表继承是错的）</summary>
        public void ResetUserRowHeights() => _userRowHeights.Clear();

        /// <summary>在行容器可视化树里找行头拖拽 Thumb（顶层 DataGridRowHeader 的直接子级）</summary>
        private static System.Windows.Controls.Primitives.Thumb? FindRowHeaderThumb(System.Windows.DependencyObject row)
        {
            System.Windows.Controls.Primitives.Thumb? Find(System.Windows.DependencyObject d)
            {
                int n = System.Windows.Media.VisualTreeHelper.GetChildrenCount(d);
                for (int i = 0; i < n; i++)
                {
                    var child = System.Windows.Media.VisualTreeHelper.GetChild(d, i);
                    if (child is System.Windows.Controls.Primitives.Thumb t) return t;
                    var deep = Find(child);
                    if (deep != null) return deep;
                }
                return null;
            }
            return Find(row);
        }

        private void RowHeaderDragDelta(object sender, System.Windows.Controls.Primitives.DragDeltaEventArgs e)
        {
            if (sender is not System.Windows.Controls.Primitives.Thumb thumb) return;
            if (IsProcessing) return;   // 数据加载中不记录，避免与程序重算互相干扰

            // 必须取行的实际高度：Thumb.Height 是显式设定值、默认 NaN，
            // 曾用它计算导致 NaN > min 恒 false，行高记忆从未生效过
            if (thumb.TemplatedParent is DataGridRow row
                && row.Item is RowView rv && rv.RowNumber > 0)
            {
                double newHeight = row.ActualHeight + e.VerticalChange;
                if (newHeight > _autoRowHeightMin)
                    _userRowHeights[rv.RowNumber] = newHeight;
            }
        }

        /// <summary>
        /// 行容器滚出视口时清掉 MinHeight/Height。
        /// 行虚拟化（Recycling）下容器会被复用给其他行：显式残留的高度会让
        /// 新内容按旧行高布局（错位、内容被顶出可视区）。清零后由 LoadingRow 重算。
        /// </summary>
        private void OnUnloadingRow(object? sender, DataGridRowEventArgs e)
        {
            e.Row.Height = double.NaN;
            e.Row.MinHeight = 0;
        }

        private int _autoRowHeightMin = 30;

        /// <summary>
        /// 用户拖拽行头调整过的行高（键 = RowNumber 数据序号）。
        /// 行虚拟化 Recycling 下容器复用会清掉显式 Height，记忆值用于在
        /// LoadingRow 重算时不低于用户设定（取 max(内容需求, 用户值)）。
        /// </summary>
        private readonly Dictionary<int, double> _userRowHeights = new();

        private void OnLoadingRow(object? sender, DataGridRowEventArgs e)
        {
            if (sender is not DataGrid grid) return;

            try
            {
                var row = e.Row;
                // 挂行头拖拽事件：记录用户手动调大的行高（Thumb 只在行容器加载时存在）
                if (FindRowHeaderThumb(row) is { } thumb)
                {
                    thumb.DragDelta -= RowHeaderDragDelta;
                    thumb.DragDelta += RowHeaderDragDelta;
                }

                // 高度需求一律用 MinHeight 表达、Height 恒为 NaN（自动）：
                // Recycling 行虚拟化下容器跨行复用，显式 Height 会残留到别的内容上
                // （错位、垂直居中的内容被顶出可视区）；MinHeight 随 LoadingRow 重算即清。
                row.Height = double.NaN;
                row.MinHeight = _autoRowHeightMin;

                // 从行数据本身统计换行数：此前只读第 0 列（RowNumber）的显示文本，
                // 永远是单行 —— 数据列里的多行文字（换行、空行）全被压进固定行高而显示不完整。
                if (row.Item is not RowView rv) return;

                int maxLines = 1;
                foreach (var value in rv.Values)
                {
                    // 图片列单元格存的是 ImageTextPair，文字在其 Text 属性里
                    string? text = value switch
                    {
                        ImageTextPair pair => pair.Text,
                        string s => s,
                        _ => null,
                    };
                    if (string.IsNullOrEmpty(text)) continue;

                    int lines = 1;
                    foreach (char c in text)
                    {
                        if (c == '\n') lines++;
                    }
                    if (lines > maxLines) maxLines = lines;
                }

                // 高度需求用 MinHeight 表达；Height 保持 NaN（自动），实际行高由内容撑开
                if (maxLines > 1)
                    row.MinHeight = _autoRowHeightMin + (maxLines - 1) * 15;

                // 用户拖过该行：行高不低于用户设定（MinHeight = max(内容需求, 用户值)）。
                // Height 恒为 NaN —— 曾经在这里固定 Height=userH，勾选「显示图片」后
                // 图片需要的行高更高，固定值反而把图片行压扁；MinHeight 只托底不封顶，
                // 文本行保持用户拖出的高度，图片行由内容自然撑开。
                int rowNo = rv.RowNumber;
                if (rowNo > 0 && _userRowHeights.TryGetValue(rowNo, out double userH)
                    && userH > row.MinHeight)
                {
                    row.MinHeight = userH;
                }
            }
            catch
            {
                // 行高属于显示细节，计算失败不应影响数据展示
            }
        }
    }
}

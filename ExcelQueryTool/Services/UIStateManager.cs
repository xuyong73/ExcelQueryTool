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
        private static readonly Thickness _defaultMargin = new(2);
        private static readonly Style? _sharedTextStyle;

        static UIStateManager()
        {
            _sharedTextStyle = new Style(typeof(TextBlock))
            {
                Setters =
                {
                    new Setter(TextBlock.TextWrappingProperty, TextWrapping.Wrap),
                    new Setter(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center),
                    new Setter(TextBlock.TextAlignmentProperty, TextAlignment.Left),
                    new Setter(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Stretch)
                }
            };
        }

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

        public void SetProcessingState(bool isProcessing)
        {
            _isProcessing = isProcessing;
            _window.Dispatcher.Invoke(() =>
            {
                if (_btnOpenFile != null) _btnOpenFile.IsEnabled = !isProcessing;
                if (_btnSearch != null) _btnSearch.IsEnabled = !isProcessing;
                if (_comboBoxWorksheets != null) _comboBoxWorksheets.IsEnabled = !isProcessing;
                if (_chkShowImages != null) _chkShowImages.IsEnabled = !isProcessing;
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

        public void ClearDataGrid(DataGrid? dataGrid)
        {
            _window.Dispatcher.Invoke(() =>
            {
                if (dataGrid != null)
                {
                    dataGrid.ItemsSource = null;
                    dataGrid.Columns.Clear();
                }
            });
        }

        public void UpdateDataGrid(DataGrid? dataGrid, System.Data.DataTable dataTable)
        {
            _window.Dispatcher.Invoke(() =>
            {
                if (dataGrid != null)
                {
                    dataGrid.ItemsSource = dataTable.DefaultView;
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
                dataGrid.Columns.Clear();

                foreach (var col in columns)
                {
                    if (col.IsImage)
                    {
                        var imageCol = new DataGridTemplateColumn
                        {
                            Header = col.Header,
                            Width = defaultImageSize,
                            MinWidth = 50
                        };

                        var stackPanelFactory = new FrameworkElementFactory(typeof(StackPanel));
                        stackPanelFactory.SetValue(StackPanel.OrientationProperty, Orientation.Vertical);

                        var imageFactory = new FrameworkElementFactory(typeof(System.Windows.Controls.Image));
                        imageFactory.SetValue(System.Windows.Controls.Image.StretchProperty, Stretch.Uniform);
                        imageFactory.SetValue(System.Windows.Controls.Image.StretchDirectionProperty, StretchDirection.Both);
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

                        var heightBinding = new Binding("ActualHeight")
                        {
                            RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(DataGridCell), 1),
                            Converter = new CellSizeConverter(),
                            ConverterParameter = defaultImageSize
                        };
                        imageFactory.SetBinding(System.Windows.Controls.Image.HeightProperty, heightBinding);

                        var imageBinding = new Binding(col.Header)
                        {
                            Converter = new ImageConverter(),
                            ConverterParameter = col.Header
                        };
                        imageFactory.SetBinding(System.Windows.Controls.Image.SourceProperty, imageBinding);

                        var textFactory = new FrameworkElementFactory(typeof(TextBlock));
                        textFactory.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center);
                        textFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
                        textFactory.SetValue(TextBlock.TextWrappingProperty, TextWrapping.Wrap);
                        textFactory.SetValue(TextBlock.TextAlignmentProperty, TextAlignment.Center);
                        textFactory.SetValue(TextBlock.FontSizeProperty, 10.0);
                        textFactory.SetValue(TextBlock.ForegroundProperty, System.Windows.Media.Brushes.Gray);
                        textFactory.SetValue(TextBlock.MarginProperty, _defaultMargin);

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
                        imageFactory.SetBinding(UIElement.VisibilityProperty, imageVisibilityBinding);

                        stackPanelFactory.AppendChild(imageFactory);
                        stackPanelFactory.AppendChild(textFactory);

                        var template = new DataTemplate();
                        template.VisualTree = stackPanelFactory;
                        imageCol.CellTemplate = template;
                        dataGrid.Columns.Add(imageCol);
                    }
                    else
                    {
                        var textCol = new DataGridTextColumn
                        {
                            Header = col.Header,
                            Binding = new Binding(col.Header),
                            Width = columnWidthMax,
                            ElementStyle = _sharedTextStyle
                        };
                        dataGrid.Columns.Add(textCol);
                    }
                }
            });
        }

        public void ApplyAutoRowHeight(DataGrid? dataGrid, int rowHeightMin = 30)
        {
            if (dataGrid == null || dataGrid.Items.Count == 0) return;

            _window.Dispatcher.Invoke(() =>
            {
                try
                {
                    foreach (var item in dataGrid.Items)
                    {
                        var row = dataGrid.ItemContainerGenerator.ContainerFromItem(item) as DataGridRow;
                        if (row != null)
                        {
                            row.MinHeight = rowHeightMin;
                            var cell = dataGrid.Columns[0].GetCellContent(item) as TextBlock;
                            if (cell != null && cell.Text != null)
                            {
                                var lineCount = cell.Text.Split('\n').Length;
                                if (lineCount > 1) row.Height = rowHeightMin + (lineCount - 1) * 15;
                            }
                        }
                    }
                }
                catch { }
            });
        }
    }
}

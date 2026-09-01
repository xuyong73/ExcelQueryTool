using OfficeOpenExcel = OfficeOpenXml;

namespace ExcelQueryTool.Services
{
    public class ExcelDataManager(ImageManager _imageManager, SearchService _searchService)
    {
        public async IAsyncEnumerable<object[]> LoadWorksheetDataAsync(
            OfficeOpenExcel.ExcelWorksheet worksheet,
            string keyword,
            bool showImages,
            [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken token,
            IProgress<string>? progress = null)
        {
            progress?.Report("正在分析工作表结构...");
            token.ThrowIfCancellationRequested();

            var dimension = worksheet.Dimension;
            if (dimension == null) yield break;

            int totalCols = dimension.Columns;
            var columns = GetColumnMetadata(worksheet, _imageManager.PictureMap);
            HashSet<int> imageColumns = [.. columns.Where(c => c.IsImage).Select(c => columns.IndexOf(c) + 1)];

            progress?.Report("正在构建数据表结构...");
            token.ThrowIfCancellationRequested();

            bool hasKeyword = !string.IsNullOrWhiteSpace(keyword);

            int startRow = 2;
            int endRow = dimension.Rows;
            int totalRows = endRow - startRow + 1;
            int batchSize = CalculateBatchSize(totalRows);

            int totalRecords = 0;

            for (int start = startRow; start <= endRow; start += batchSize)
            {
                int end = Math.Min(start + batchSize - 1, endRow);
                progress?.Report($"正在加载数据: {start}-{end}/{endRow}");

                List<object[]> data;
                try
                {
                    data = await LoadBatchDataAsync(
                        worksheet, start, end, totalCols,
                        keyword, showImages, imageColumns, hasKeyword, token);
                }
                catch (OperationCanceledException)
                {
                    progress?.Report("数据加载已取消");
                    yield break;
                }
                catch (Exception ex)
                {
                    progress?.Report($"数据加载错误: {ex.Message}");
                    yield break;
                }

                foreach (var rowData in data)
                {
                    yield return rowData;
                    totalRecords++;
                }

                if (token.IsCancellationRequested) yield break;
            }

            progress?.Report($"数据加载完成，共 {totalRecords} 条记录");
        }

        private static int CalculateBatchSize(int totalRows)
        {
            return totalRows switch
            {
                <= 10000 => totalRows,
                <= 50000 => 10000,
                <= 100000 => 20000,
                _ => 50000
            };
        }

        private async Task<List<object[]>> LoadBatchDataAsync(
            OfficeOpenExcel.ExcelWorksheet worksheet,
            int start,
            int end,
            int totalCols,
            string keyword,
            bool showImages,
            HashSet<int> imageColumns,
            bool hasKeyword,
            CancellationToken token)
        {
            return await Task.Run(() =>
            {
                try
                {
                    List<object[]> data = [];
                    var range = worksheet.Cells[start, 1, end, totalCols];
                    var values = range.Value as object[,];
                    int rowBase = values?.GetLowerBound(0) ?? 0;
                    int colBase = values?.GetLowerBound(1) ?? 0;

                    HashSet<int> dateColumns = [];
                    for (int col = 1; col <= totalCols; col++)
                    {
                        if (IsDateFormatted(worksheet.Cells[start, col]))
                            dateColumns.Add(col);
                    }

                    for (int row = start; row <= end; row++)
                    {
                        token.ThrowIfCancellationRequested();
                        if (worksheet.Row(row).Hidden) continue;

                        var rowData = new object[totalCols + 1];
                        rowData[0] = row - 1;
                        bool hasData = false;
                        int arrRow = rowBase + (row - start);

                        for (int col = 1; col <= totalCols; col++)
                        {
                            try
                            {
                                object? cellValue = values?.GetValue(arrRow, colBase + (col - 1));
                                object? result;

                                if (showImages && imageColumns.Contains(col))
                                {
                                    var img = _imageManager.LoadCellImage(worksheet, row, col);
                                    var text = CellValueToString(cellValue, dateColumns.Contains(col));

                                    if (img != null && ImageManager.IsImageValid(img))
                                    {
                                        result = new ImageTextPair { Image = img, Text = text };
                                        hasData = true;
                                    }
                                    else
                                    {
                                        result = text;
                                    }
                                }
                                else
                                {
                                    result = CellValueToString(cellValue, dateColumns.Contains(col));
                                    hasData = hasData || result is string s && s.Length > 0;
                                }

                                rowData[col] = result;
                            }
                            catch (OutOfMemoryException)
                            {
                                rowData[col] = "[内存不足]";
                                _imageManager.Cache.Dispose();
                                GC.Collect();
                            }
                            catch (Exception ex) when (ex is ArgumentException or InvalidOperationException)
                            {
                                rowData[col] = "[图片错误]";
                            }
                            catch { rowData[col] = "[错误]"; }
                        }

                        bool shouldAddRow = hasKeyword
                            ? RowMatchesKeyword(rowData, keyword)
                            : hasData;

                        if (shouldAddRow) data.Add(rowData);
                    }

                    return data;
                }
                catch (OperationCanceledException) { return []; }
                catch (Exception) { return []; }
            }, token);
        }

        private static bool IsDateFormatted(OfficeOpenExcel.ExcelRange cell)
        {
            try
            {
                var format = cell.Style.Numberformat.Format ?? "";
                if (string.IsNullOrEmpty(format)) return false;

                return format.Contains('y') || format.Contains('Y') ||
                       format.Contains('m') || format.Contains('M') ||
                       format.Contains('d') || format.Contains('D');
            }
            catch
            {
                return false;
            }
        }

        private static string CellValueToString(object? value, bool isDateColumn = false)
        {
            return value switch
            {
                null or DBNull => string.Empty,
                string s => s,
                DateTime dt => dt.ToString("yyyy-MM-dd"),
                double d when isDateColumn && d >= 1 && d <= 2958465 => DateTime.FromOADate(d).ToString("yyyy-MM-dd"),
                _ => value?.ToString() is { Length: > 0 } str ? str : string.Empty
            };
        }

        private bool RowMatchesKeyword(object[] rowData, string keyword)
        {
            if (string.IsNullOrWhiteSpace(keyword)) return true;

            var searchConditions = _searchService.ParseSearchConditions(keyword);
            return searchConditions.Evaluate(rowData, RowContainsTerm);
        }

        private bool RowContainsTerm(object[] rowData, string term)
        {
            if (string.IsNullOrWhiteSpace(term)) return true;

            for (int i = 1; i < rowData.Length; i++)
            {
                var value = rowData[i];
                if (value == null) continue;

                string textValue = value switch
                {
                    ImageTextPair p => p.Text ?? string.Empty,
                    _ => value.ToString() ?? string.Empty
                };

                if (textValue.Contains(term, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        public static List<ColumnMeta> GetColumnMetadata(OfficeOpenExcel.ExcelWorksheet ws, Dictionary<(string, int, int), OfficeOpenExcel.Drawing.ExcelPicture?> pictureMap)
        {
            List<ColumnMeta> columns = [];
            var counts = new Dictionary<string, int>();

            if (ws?.Dimension == null) return columns;

            var worksheetName = ws.Name ?? "";
            int totalCols = ws.Dimension.Columns;
            int totalRows = ws.Dimension.Rows;
            var pictureKeys = pictureMap.Keys.Where(k => k.Item1 == worksheetName).Select(k => k.Item3).ToHashSet();

            var headerRange = ws.Cells[1, 1, 1, totalCols];
            var headerValues = headerRange.Value;

            for (int col = 1; col <= totalCols; col++)
            {
                string header = headerValues is object[,] hv
                    ? hv.GetValue(hv.GetLowerBound(0), hv.GetLowerBound(1) + (col - 1))?.ToString() ?? ""
                    : ws.Cells[1, col]?.Text ?? "";

                bool isHeaderEmpty = string.IsNullOrWhiteSpace(header);

                if (isHeaderEmpty && totalRows > 1)
                {
                    var columnRange = ws.Cells[2, col, totalRows, col];
                    if (columnRange.Value is object[,] columnValues)
                    {
                        bool hasData = false;
                        int rowBase = columnValues.GetLowerBound(0);
                        int colBase = columnValues.GetLowerBound(1);
                        for (int r = rowBase; r <= columnValues.GetUpperBound(0); r++)
                        {
                            var val = columnValues.GetValue(r, colBase);
                            if (val != null && val is not DBNull && (val is not string s || s.Length > 0))
                            {
                                hasData = true;
                                break;
                            }
                        }
                        if (!hasData) continue;
                    }
                }

                if (isHeaderEmpty) header = $"<列{col}>";

                if (!counts.TryGetValue(header, out int count))
                    count = 1;
                else
                    count++;
                counts[header] = count;

                string name = count > 1 ? $"{header}_{count}" : header;
                columns.Add(new ColumnMeta(name, pictureKeys.Contains(col)));
            }

            return columns;
        }

        public static List<string> GetWorksheetNames(OfficeOpenExcel.ExcelPackage package)
        {
            return [.. package.Workbook.Worksheets.Select(ws => ws.Name)];
        }
    }
}

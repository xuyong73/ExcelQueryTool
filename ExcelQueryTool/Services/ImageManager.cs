using System.Collections.Concurrent;
using System.Diagnostics;
using System.Drawing;
using System.IO;

namespace ExcelQueryTool.Services
{
    /// <summary>
    /// 图片单元格的值：一个由 <see cref="ImageManager"/> 的 LRU 缓存<b>持有</b>的位图引用，外加该单元格的文本。
    ///
    /// 所有权约定：<b>本类型不是位图的所有者，因此不实现 IDisposable</b>。
    /// 位图由 <see cref="ImageCacheManager"/> 统一持有；
    /// 单元格只保存引用，不得释放它。
    ///
    /// 这里刻意不提供可写属性：早期实现带 setter + IDisposable，看起来「管理到位」，
    /// 但该 Dispose 从未被任何地方调用，而缓存淘汰仍会销毁位图 —— 结果是同一个
    /// Image 实例被缓存、setter、Dispose 三方各自释放，正在显示的图被销毁后
    /// 由转换器兜成空白格（现象是「滚动后图片随机变空白」，极难复现）。
    /// 现在只有一个所有者，问题根除。
    ///
    /// 2026-09-23 起缓存淘汰/替换/清空<b>只移除引用、不再销毁位图</b>：
    /// 本类型持有的引用在行集合（RowList）的整个生命期内都可能被重新绑定渲染
    /// （虚拟化滚动 = 反复重绑），而行集合存活整个会话 —— 销毁任何
    /// 仍被引用的位图都会复现「空白格」。位图在无引用后由 GC 回收
    /// （System.Drawing.Image 终结器释放原生资源）。
    /// </summary>
    public class ImageTextPair
    {
        /// <summary>
        /// 由缓存持有的位图列表（同一单元格锚定多张图时按包内锚点顺序排列），
        /// 本类型只读引用，不要释放。
        /// </summary>
        public required IReadOnlyList<Image> Images { get; init; }

        public string? Text { get; init; }
    }

    public class ColumnMeta(string header, bool isImage, double excelWidth = 0, bool isNumeric = false)
    {
        public string Header { get; } = header;
        public bool IsImage { get; } = isImage;

        /// <summary>Excel 原始列宽（字符单位；0 = 未声明，用程序默认宽度）</summary>
        public double ExcelWidth { get; } = excelWidth;

        /// <summary>
        /// 该列是否为数字列（数据区非空值中数字占比 ≥ 0.8）。
        /// 供 UI 做列级对齐：数字列右对齐、文本列左对齐（Excel「常规」格式的通用惯例）。
        /// </summary>
        public bool IsNumeric { get; } = isNumeric;
    }

    /// <summary>
    /// 已解码位图的 LRU 缓存。
    ///
    /// 淘汰语义（2026-09-23 起）：<b>淘汰/替换/清空只移除缓存条目，绝不 Dispose 位图</b>。
    /// 位图会以 <see cref="ImageTextPair"/> 的形式被行集合（RowList）长期引用（虚拟化滚动会
    /// 反复重新绑定），任何主动销毁都会让还在显示的单元格变空白 —— 这正是
    /// E4 图片不显示的根因（52 张图解码后约 190MB 超 100MB 缓存上限，最早解码的
    /// 30 余张被淘汰销毁，其所在单元格全部空白）。
    /// 容量/体积上限只决定「缓存是否还记着这条目」；无引用的位图由 GC 回收。
    /// </summary>
    public class ImageCacheManager : IDisposable
    {
        private readonly int _capacity;
        private readonly long _maxCacheSizeBytes;
        private readonly ConcurrentDictionary<string, (Image? Value, DateTime LastAccess, long Size)> _cache = [];
        private readonly System.Threading.Timer _cleanupTimer;
        private readonly Lock _lock = new();
        private long _currentCacheSize;
        private bool _disposed;

        public ImageCacheManager(int capacity, TimeSpan cleanupInterval, long maxCacheSizeBytes = 100 * 1024 * 1024)
        {
            _capacity = capacity;
            _maxCacheSizeBytes = maxCacheSizeBytes;
            _cleanupTimer = new System.Threading.Timer(Cleanup, null, (int)cleanupInterval.TotalMilliseconds, (int)cleanupInterval.TotalMilliseconds);
        }

        public bool TryGet(string key, out Image? value)
        {
            if (_disposed) { value = null; return false; }

            if (_cache.TryGetValue(key, out var entry))
            {
                lock (_lock) { _cache[key] = (entry.Value, DateTime.UtcNow, entry.Size); }
                value = entry.Value;
                return true;
            }
            value = null;
            return false;
        }

        public void Add(string key, Image? value)
        {
            if (_disposed || value == null) return;

            long imageSize = EstimateImageSize(value);

            lock (_lock)
            {
                if (_cache.TryGetValue(key, out var existingEntry))
                {
                    _currentCacheSize -= existingEntry.Size;
                    // 不 Dispose：旧位图可能仍被 ImageTextPair 引用着（行集合生命期 = 整个会话），
                    // 销毁会让正在显示的单元格渲染为空白。无引用后由 GC 回收。
                }

                _cache[key] = (value, DateTime.UtcNow, imageSize);
                _currentCacheSize += imageSize;

                if (_capacity > 0 && _cache.Count > _capacity * 1.2 || _currentCacheSize > _maxCacheSizeBytes)
                {
                    Cleanup(null);
                }
            }
        }

        private static long EstimateImageSize(Image image)
        {
            try
            {
                return (long)image.Width * image.Height * 4;
            }
            catch (Exception ex)
            {
                // 尺寸估算失败（图对象已失效）：按 1 MB 估值参与 LRU 预算
                Debug.WriteLine($"估算图片大小失败: {ex.Message}");
                return 1024 * 1024;
            }
        }

        private void Cleanup(object? state)
        {
            if (_disposed) return;

            lock (_lock)
            {
                if (_cache.Count <= _capacity && _currentCacheSize <= _maxCacheSizeBytes) return;

                var sortedItems = _cache.OrderBy(x => x.Value.LastAccess).ToList();
                int itemsToRemove = Math.Max(0, _cache.Count - _capacity);
                long sizeToRemove = Math.Max(0L, _currentCacheSize - _maxCacheSizeBytes);
                long removedSize = 0;

                foreach (var item in sortedItems)
                {
                    if (itemsToRemove <= 0 && removedSize >= sizeToRemove) break;

                    if (_cache.TryRemove(item.Key, out var entry))
                    {
                        // 只淘汰缓存条目，不销毁位图 —— 位图可能仍被 ImageTextPair
                        // 引用（正在显示或等待重新绑定），销毁即白屏。
                        _currentCacheSize -= entry.Size;
                        removedSize += entry.Size;
                        itemsToRemove--;
                    }
                }
            }
        }

        public void Clear()
        {
            lock (_lock)
            {
                // 只清条目不销毁位图：若仍有 ImageTextPair 引用着（如切表前旧行集合
                // 尚未释放完毕），Dispose 会造成瞬时白屏；无引用的位图由 GC 回收。
                _cache.Clear();
                _currentCacheSize = 0;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _cleanupTimer?.Dispose();
                Clear();
                _disposed = true;
            }
            GC.SuppressFinalize(this);
        }
    }

    public class ImageManager(int cacheCapacity, TimeSpan cleanupInterval, int maxImageWidth = 2048, int maxImageHeight = 2048, long maxImageFileSize = 10 * 1024 * 1024) : IDisposable
    {
        private readonly ImageCacheManager _imageCache = new(cacheCapacity, cleanupInterval);
        private readonly int _maxImageWidth = maxImageWidth;
        private readonly int _maxImageHeight = maxImageHeight;
        private readonly long _maxImageFileSize = maxImageFileSize;
        private bool _disposed;

        /// <summary>(工作表名, 行 1 基, 列 1 基) → 图片二进制列表（同格多图），由 XlsxPictureIndex 填充</summary>
        private readonly Dictionary<(string, int, int), List<byte[]>> _pictureMap = [];

        /// <summary>
        /// 已建立索引的文件路径。<see cref="BuildPictureIndex"/> 以它为判据跳过重复构建 ——
        /// 这是本类最重要的性能约定：切工作表、勾选「显示图片」都会走到建索引这一步，
        /// 而对 109 MB 的包重扫一次就要 387 ms（实测），必须避免。
        /// </summary>
        private string? _indexedFilePath;

        /// <summary>图片索引是否已就绪（已针对该文件构建过）</summary>
        public bool HasPictureIndexFor(string filePath)
            => _indexedFilePath != null
               && string.Equals(_indexedFilePath, filePath, StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// 只清<b>位图缓存</b>，保留图片索引。
        /// 
        /// 用于「不需要重新解析包，但希望释放已解码位图」的场景（例如关闭图片显示、
        /// 加载新工作表前回收内存）。位图缓存可随时按需重建，成本远低于重扫包。
        /// </summary>
        public void ClearImageCache() => _imageCache.Clear();

        /// <summary>清空图片索引三件套（锚点表、按表列缓存、已索引文件记忆），不动位图缓存。</summary>
        private void ResetIndex()
        {
            _pictureMap.Clear();
            _pictureColumnsBySheet.Clear();
            _indexedFilePath = null;
        }

        /// <summary>
        /// 清空图片索引与位图缓存（完整重置）。
        /// 需要重扫包时用 <see cref="BuildPictureIndex"/>，不要用本方法再重建。
        /// </summary>
        public void Clear()
        {
            ResetIndex();
            _imageCache.Clear();
        }

        /// <summary>
        /// 从 xlsx 流建立图片索引。替代原先遍历 EPPlus Drawing 的方式。
        /// 流必须可 Seek（xlsx 为 ZipArchive 读取所需）。
        /// 
        /// 索引按 <paramref name="filePath"/> 记忆：同一文件重复调用直接返回（不清空、不重扫）。
        /// 传 null 表示不参与记忆，强制重建。
        /// </summary>
        public void BuildPictureIndex(Stream stream, CancellationToken ct, string? filePath = null)
        {
            // 已经为该文件建过索引：跳过整包扫描（大包上这是最贵的一步）
            if (filePath != null && HasPictureIndexFor(filePath)) return;

            Clear();
            ct.ThrowIfCancellationRequested();

            var index = XlsxPictureIndex.Build(stream);
            ct.ThrowIfCancellationRequested();

            foreach (var (key, bytesList) in index.Pictures)
            {
                _pictureMap[key] = bytesList;
            }

            _indexedFilePath = filePath;
        }

        /// <summary>工作表名 → 有图片的列号集合（1 基）。索引变更时整体失效重建。</summary>
        private readonly Dictionary<string, HashSet<int>> _pictureColumnsBySheet = new(StringComparer.Ordinal);

        /// <summary>
        /// 取某工作表中有图片的列号集合（1 基）。
        /// 
        /// 按工作表缓存在 <see cref="_pictureColumnsBySheet"/>：本方法每次载入工作表都会调用，
        /// 而原始实现要遍历整个 <see cref="_pictureMap"/>（大包上可达数百项）。索引建好后
        /// 结果不变，故只在索引重建时失效。
        /// </summary>
        public HashSet<int> GetPictureColumns(string worksheetName)
        {
            if (string.IsNullOrEmpty(worksheetName)) return [];

            if (_pictureColumnsBySheet.TryGetValue(worksheetName, out var cached)) return cached;

            var cols = new HashSet<int>();
            foreach (var key in _pictureMap.Keys)
            {
                if (string.Equals(key.Item1, worksheetName, StringComparison.Ordinal)) cols.Add(key.Item3);
            }

            _pictureColumnsBySheet[worksheetName] = cols;
            return cols;
        }

        /// <summary>
        /// 取单元格的全部图片（同格多图按包内锚点顺序）。返回值由缓存<b>持有</b>，
        /// 调用方只读引用、不得释放。缓存按容量/体积淘汰时会销毁位图；
        /// 此时仍显示该图的单元格会渲染为空白（见 ImageTextPair 注释）。
        /// </summary>
        public IReadOnlyList<Image> LoadCellImages(string worksheetName, int row, int col)
        {
            if (_disposed) return [];

            try
            {
                var key = (worksheetName ?? "", row, col);
                if (!_pictureMap.TryGetValue(key, out var byteList) || byteList.Count == 0)
                    return [];

                var images = new List<Image>(byteList.Count);
                for (int i = 0; i < byteList.Count; i++)
                {
                    var bytes = byteList[i];
                    if (bytes.Length == 0 || bytes.Length > _maxImageFileSize) continue;

                    // 每张图独立缓存键：同格第 N 张不与第 0 张互相覆盖
                    string cacheKey = $"{worksheetName}_{row}_{col}_{i}";
                    if (_imageCache.TryGet(cacheKey, out var cachedImage))
                    {
                        images.Add(cachedImage!);
                        continue;
                    }

                    using var ms = new MemoryStream(bytes);
                    using var raw = Image.FromStream(ms);
                    // GDI+ 要求流在图片生命周期内保持打开；这里复制为独立位图，
                    // 使 ms 的释放不影响后续渲染（否则图片可能显示为空白）
                    var image = new Bitmap(raw);
                    var resized = ResizeImageIfNeeded(image);
                    if (resized != image) image.Dispose();
                    _imageCache.Add(cacheKey, resized);
                    images.Add(resized);
                }

                return images;
            }
            catch (Exception ex)
            {
                // 取图失败（字节损坏/格式不支持）：该格按无图处理，不阻塞数据展示
                Debug.WriteLine($"读取单元格图片失败: {ex.Message}");
                return [];
            }
        }

        private Image ResizeImageIfNeeded(Image originalImage)
        {
            try
            {
                if (originalImage.Width <= _maxImageWidth && originalImage.Height <= _maxImageHeight)
                {
                    return originalImage;
                }

                int newWidth = Math.Min(originalImage.Width, _maxImageWidth);
                int newHeight = Math.Min(originalImage.Height, _maxImageHeight);

                if (originalImage.Width > _maxImageWidth)
                {
                    double ratio = (double)_maxImageWidth / originalImage.Width;
                    newHeight = (int)(originalImage.Height * ratio);
                }

                if (newHeight > _maxImageHeight)
                {
                    double ratio = (double)_maxImageHeight / newHeight;
                    newWidth = (int)(newWidth * ratio);
                    newHeight = _maxImageHeight;
                }

                var resizedImage = new Bitmap(newWidth, newHeight);
                using (var graphics = Graphics.FromImage(resizedImage))
                {
                    graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    graphics.DrawImage(originalImage, 0, 0, newWidth, newHeight);
                }

                return resizedImage;
            }
            catch (Exception ex)
            {
                // 缩放失败：用原图顶替（可能超尺寸），不丢图
                Debug.WriteLine($"缩放图片失败: {ex.Message}");
                return originalImage;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _imageCache.Dispose();
                ResetIndex();
                _disposed = true;
            }
            GC.SuppressFinalize(this);
        }
    }
}

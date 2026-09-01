using System.Collections.Concurrent;
using System.Drawing;
using System.IO;
using System.Threading;
using OfficeOpenExcel = OfficeOpenXml;
using OfficeOpenExcelDrawing = OfficeOpenXml.Drawing;

namespace ExcelQueryTool.Services
{
    public class ImageTextPair : IDisposable
    {
        private Image? _image;
        private bool _disposed;

        public Image? Image
        {
            get => _image;
            set
            {
                if (_image != null && _image != value)
                {
                    _image.Dispose();
                }
                _image = value;
            }
        }

        public string? Text { get; set; }

        public void Dispose()
        {
            if (!_disposed)
            {
                _image?.Dispose();
                _image = null;
                _disposed = true;
                GC.SuppressFinalize(this);
            }
        }
    }

    public class ColumnMeta(string header, bool isImage)
    {
        public string Header { get; } = header;
        public bool IsImage { get; } = isImage;
    }

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
                    existingEntry.Value?.Dispose();
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
            catch
            {
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
                        entry.Value?.Dispose();
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
                foreach (var item in _cache.Values)
                {
                    item.Value?.Dispose();
                }
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
        private static readonly ConcurrentDictionary<Type, System.Reflection.MethodInfo?> s_getMethodCache = new();
        private static readonly ConcurrentDictionary<Type, System.Reflection.MethodInfo?> s_getImageMethodCache = new();
        private static readonly ConcurrentDictionary<Type, System.Reflection.MethodInfo?> s_getImageBytesMethodCache = new();
        private static readonly ConcurrentDictionary<Type, System.Reflection.PropertyInfo?> s_imagePropertyCache = new();
        private static readonly ConcurrentDictionary<Type, System.Reflection.PropertyInfo?> s_imageBytesPropertyCache = new();
        private static readonly ConcurrentDictionary<Type, System.Reflection.PropertyInfo?> s_excelImageBytesPropertyCache = new();

        private readonly ImageCacheManager _imageCache = new(cacheCapacity, cleanupInterval);
        private readonly Dictionary<(string, int, int), OfficeOpenExcelDrawing.ExcelPicture?> _pictureMap = [];
        private readonly int _maxImageWidth = maxImageWidth;
        private readonly int _maxImageHeight = maxImageHeight;
        private readonly long _maxImageFileSize = maxImageFileSize;
        private bool _disposed;

        public ImageCacheManager Cache => _imageCache;

        public Dictionary<(string, int, int), OfficeOpenExcelDrawing.ExcelPicture?> PictureMap => _pictureMap;

        public void Clear()
        {
            _pictureMap.Clear();
        }

        public void BuildPictureIndex(OfficeOpenExcel.ExcelWorksheet ws, CancellationToken ct)
        {
            _pictureMap.Clear();
            var worksheetName = ws.Name ?? "";

            foreach (var pic in ws.Drawings.OfType<OfficeOpenExcelDrawing.ExcelPicture>())
            {
                ct.ThrowIfCancellationRequested();
                var key = (worksheetName, pic.From.Row + 1, pic.From.Column + 1);
                _pictureMap.TryAdd(key, pic);
            }
        }

        public Image? LoadCellImage(OfficeOpenExcel.ExcelWorksheet ws, int row, int col)
        {
            if (_disposed) return null;

            try
            {
                string cacheKey = $"{ws.Name}_{row}_{col}";
                if (_imageCache.TryGet(cacheKey, out var cachedImage))
                    return cachedImage;

                var key = (ws.Name ?? "", row, col);
                if (_pictureMap.TryGetValue(key, out var excelPicture) && excelPicture?.Image != null)
                {
                    var image = GetImageFromExcelImage(excelPicture.Image);
                    if (image != null)
                    {
                        var resized = ResizeImageIfNeeded(image);
                        if (resized != image) image.Dispose();
                        _imageCache.Add(cacheKey, resized);
                        return resized;
                    }
                }

                var cellPicture = ws.Cells[row, col].Picture;
                if (!cellPicture.Exists) return null;

                var picType = cellPicture.GetType();

                var getMethod = s_getMethodCache.GetOrAdd(picType, t => t.GetMethod("Get"));
                if (getMethod != null)
                {
                    var excelCellPic = getMethod.Invoke(cellPicture, null);
                    if (excelCellPic != null)
                    {
                        var cellPicType = excelCellPic.GetType();

                        var getImageMethod = s_getImageMethodCache.GetOrAdd(cellPicType, t => t.GetMethod("GetImage"));
                        if (getImageMethod != null)
                        {
                            var excelImg = getImageMethod.Invoke(excelCellPic, null);
                            if (excelImg != null)
                            {
                                var img = GetImageFromExcelImage(excelImg);
                                if (img != null)
                                {
                                    var resized = ResizeImageIfNeeded(img);
                                    if (resized != img) img.Dispose();
                                    _imageCache.Add(cacheKey, resized);
                                    return resized;
                                }
                            }
                        }

                        var getBytesMethod = s_getImageBytesMethodCache.GetOrAdd(cellPicType, t => t.GetMethod("GetImageBytes"));
                        if (getBytesMethod?.Invoke(excelCellPic, null) is byte[] bytes2 && bytes2.Length > 0 && bytes2.Length <= _maxImageFileSize)
                        {
                            using var ms = new MemoryStream(bytes2);
                            var img = Image.FromStream(ms);
                            var resized = ResizeImageIfNeeded(img);
                            if (resized != img) img.Dispose();
                            _imageCache.Add(cacheKey, resized);
                            return resized;
                        }
                    }
                }

                var imageProp = s_imagePropertyCache.GetOrAdd(picType, t => t.GetProperty("Image"));
                if (imageProp?.GetValue(cellPicture) is { } excelImage2 && excelImage2 != null)
                {
                    var img = GetImageFromExcelImage(excelImage2);
                    if (img != null)
                    {
                        var resized = ResizeImageIfNeeded(img);
                        if (resized != img) img.Dispose();
                        _imageCache.Add(cacheKey, resized);
                        return resized;
                    }
                }

                var bytesProp = s_imageBytesPropertyCache.GetOrAdd(picType, t => t.GetProperty("ImageBytes") ?? t.GetProperty("Bytes"));
                if (bytesProp?.GetValue(cellPicture) is byte[] bytes && bytes.Length > 0 && bytes.Length <= _maxImageFileSize)
                {
                    using var ms = new MemoryStream(bytes);
                    var img = Image.FromStream(ms);
                    var resized = ResizeImageIfNeeded(img);
                    if (resized != img) img.Dispose();
                    _imageCache.Add(cacheKey, resized);
                    return resized;
                }

                return null;
            }
            catch { return null; }
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
            catch
            {
                return originalImage;
            }
        }

        private static Bitmap? GetImageFromExcelImage(object excelImage)
        {
            try
            {
                var t = excelImage.GetType();

                var bytesProp = s_excelImageBytesPropertyCache.GetOrAdd(t, type => type.GetProperty("ImageBytes"));
                if (bytesProp?.GetValue(excelImage) is byte[] bytes && bytes.Length > 0)
                {
                    using var ms = new MemoryStream(bytes);
                    return new Bitmap(Image.FromStream(ms));
                }

                var getBytesMethod = s_getImageBytesMethodCache.GetOrAdd(t, type => type.GetMethod("get_ImageBytes"));
                if (getBytesMethod?.Invoke(excelImage, null) is byte[] bytes2 && bytes2.Length > 0)
                {
                    using var ms = new MemoryStream(bytes2);
                    return new Bitmap(Image.FromStream(ms));
                }

                return null;
            }
            catch { return null; }
        }

        public static bool IsImageValid(Image img)
        {
            try { return img != null && img.Width > 0 && img.Height > 0; }
            catch { return false; }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _imageCache.Dispose();
                _pictureMap.Clear();
                _disposed = true;
            }
            GC.SuppressFinalize(this);
        }
    }
}

using System;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace ExcelQueryTool.Services
{
    public class ImageConverter : IValueConverter
    {
        public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is ImageTextPair pair && pair.Image != null)
            {
                try
                {
                    byte[] imageBytes;
                    using (var ms = new MemoryStream())
                    {
                        try { pair.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Png); }
                        catch { pair.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp); }
                        imageBytes = ms.ToArray();
                    }

                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    var streamSource = new MemoryStream(imageBytes);
                    bitmap.StreamSource = streamSource;
                    bitmap.EndInit();
                    streamSource.Dispose();
                    bitmap.Freeze();
                    return bitmap;
                }
                catch
                {
                    return null!;
                }
            }
            return null!;
        }

        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) => throw new NotImplementedException();
    }

    public class ImageTextConverter : IValueConverter
    {
        public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is ImageTextPair pair) return pair.Text ?? "🖼️ 图片";
            return value?.ToString() ?? string.Empty;
        }

        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) => throw new NotImplementedException();
    }

    public class ImageVisibilityConverter : IValueConverter
    {
        public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is ImageTextPair pair && pair.Image != null) return Visibility.Visible;
            return Visibility.Collapsed;
        }

        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) => throw new NotImplementedException();
    }

    public class CellSizeConverter : IValueConverter
    {
        public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            double defaultSize = parameter is double paramSize && paramSize > 0 ? paramSize : 150.0;
            if (value is double cellSize && cellSize > 0) return Math.Max(50.0, cellSize - 20.0);
            return defaultSize;
        }

        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) => throw new NotImplementedException();
    }
}

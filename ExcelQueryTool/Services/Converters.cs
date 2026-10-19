using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace ExcelQueryTool.Services
{
    /// <summary>
    /// 单向值转换器基类：本文件全部转换器只做「源数据 → 界面」单向转换，
    /// ConvertBack 统一 fail-fast——单向绑定不会回调它，真被调到即绑定方向配错。
    /// </summary>
    public abstract class OneWayValueConverter : IValueConverter
    {
        public abstract object Convert(object? value, Type targetType, object? parameter, CultureInfo culture);

        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
            => throw new NotImplementedException("单向转换器不支持回写：请确认绑定方向为 OneWay");
    }

    /// <summary>
    /// 单元格图片集 → ItemsSource：把 ImageTextPair 交出其 Images 列表（同格多图按锚点顺序），
    /// 由模板列里的 ItemsControl 逐张渲染（每张再经 <see cref="ImageSourceConverter"/> 转换）。
    /// </summary>
    public class CellImagesConverter : OneWayValueConverter
    {
        public override object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
            => value is ImageTextPair pair ? pair.Images : null!;
    }

    /// <summary>
    /// 单张 GDI 位图 → WPF <see cref="BitmapImage"/>（经 PNG/BMP 中转，Freeze 供跨线程渲染）。
    /// 输入是 ItemsControl 当前项（GDI Image），不再是 ImageTextPair。
    /// </summary>
    public class ImageSourceConverter : OneWayValueConverter
    {
        public override object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is Image img)
            {
                try
                {
                    byte[] imageBytes;
                    using (var ms = new MemoryStream())
                    {
                        // 个别 GDI+ 格式（如带索引色的图）不支持 PNG 编码：退回 BMP 兜底
                        try { img.Save(ms, System.Drawing.Imaging.ImageFormat.Png); }
                        catch { img.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp); }
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
                catch (Exception ex)
                {
                    // 单张图转换失败：该格显示为空，不阻塞其余单元格
                    Debug.WriteLine($"图片转 ImageSource 失败: {ex.Message}");
                    return null!;
                }
            }
            return null!;
        }
    }

    public class ImageTextConverter : OneWayValueConverter
    {
        public override object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is ImageTextPair pair) return pair.Text ?? "🖼️ 图片";
            return value?.ToString() ?? string.Empty;
        }
    }

    public class ImageVisibilityConverter : OneWayValueConverter
    {
        public override object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            if (value is ImageTextPair pair && pair.Images.Count > 0) return Visibility.Visible;
            return Visibility.Collapsed;
        }
    }

    public class CellSizeConverter : OneWayValueConverter
    {
        public override object Convert(object? value, Type targetType, object? parameter, CultureInfo culture)
        {
            double defaultSize = parameter is double paramSize && paramSize > 0 ? paramSize : 150.0;
            if (value is double cellSize && cellSize > 0) return Math.Max(50.0, cellSize - 20.0);
            return defaultSize;
        }
    }
}

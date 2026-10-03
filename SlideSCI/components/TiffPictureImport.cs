using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace SlideSCI
{
    /// <summary>
    /// 将高位深 TIFF 的首个画面转换为 PowerPoint 可嵌入的 8 位 PNG。
    /// 原文件只读；临时副本仅在插入期间保留，来源和标定仍使用原 TIFF。
    /// </summary>
    internal sealed class TiffPictureImport : IDisposable
    {
        // 显微 TIFF 的分辨率可能是物理标定，不能用于计算幻灯片中的显示尺寸。
        private const double DisplayDpi = 96.0;
        private readonly bool temporary;
        internal string ImportPath { get; }
        internal float WidthPoints { get; }
        internal float HeightPoints { get; }

        private TiffPictureImport(string path, bool temporary = false, int pixelWidth = 0, int pixelHeight = 0)
        {
            ImportPath = path;
            this.temporary = temporary;
            WidthPoints = pixelWidth > 0 ? (float)(pixelWidth * 72.0 / DisplayDpi) : -1;
            HeightPoints = pixelHeight > 0 ? (float)(pixelHeight * 72.0 / DisplayDpi) : -1;
        }

        internal static TiffPictureImport Prepare(string sourcePath)
        {
            string extension = Path.GetExtension(sourcePath);
            if (!string.Equals(extension, ".tif", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(extension, ".tiff", StringComparison.OrdinalIgnoreCase))
                return new TiffPictureImport(sourcePath);

            using (var stream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                // 按需解码，只访问第一页，避免加载整个显微图像堆栈。
                var decoder = new TiffBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat,
                    BitmapCacheOption.OnDemand);
                if (decoder.Frames.Count == 0) throw new InvalidDataException("TIFF 中没有图像。");
                BitmapFrame frame = decoder.Frames[0];
                PixelFormat format = frame.Format;
                bool gray16 = format == PixelFormats.Gray16;
                bool color16 = format == PixelFormats.Rgb48 || format == PixelFormats.Rgba64 ||
                    format == PixelFormats.Prgba64;
                if (!gray16 && !color16)
                    return new TiffPictureImport(sourcePath, pixelWidth: frame.PixelWidth, pixelHeight: frame.PixelHeight);

                BitmapSource converted = gray16 ? ConvertGray16(frame) : ConvertColor16(frame);
                string temporaryPath = Path.Combine(Path.GetTempPath(),
                    "SlideSCI-tiff-8bit-" + Guid.NewGuid().ToString("N") + ".png");
                var result = new TiffPictureImport(temporaryPath, true, frame.PixelWidth, frame.PixelHeight);
                try
                {
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(converted));
                    using (var output = new FileStream(temporaryPath, FileMode.CreateNew, FileAccess.Write))
                        encoder.Save(output);
                    return result;
                }
                catch
                {
                    result.Dispose();
                    throw;
                }
            }
        }

        private static BitmapSource ConvertGray16(BitmapSource source)
        {
            int width = source.PixelWidth;
            int height = source.PixelHeight;
            int count = checked(width * height);
            var pixels = new ushort[count];
            // WIC 输出本机字节序，不依赖 TIFF 文件的大端/小端存储方式。
            source.CopyPixels(pixels, checked(width * sizeof(ushort)), 0);
            GetAutoDisplayLevels(pixels, 1, out double minimum, out double maximum);
            var output = new byte[count];
            double scale = 256.0 / (maximum - minimum + 1.0);
            for (int i = 0; i < count; i++) output[i] = ToDisplayByte(pixels[i], minimum, scale);
            return BitmapSource.Create(width, height, DisplayDpi, DisplayDpi,
                PixelFormats.Gray8, null, output, width);
        }

        private static BitmapSource ConvertColor16(BitmapSource source)
        {
            // 先在 16 位精度下解除预乘 alpha，不能先降为 8 位再调对比度。
            if (source.Format == PixelFormats.Prgba64)
                source = new FormatConvertedBitmap(source, PixelFormats.Rgba64, null, 0);
            int channels = source.Format == PixelFormats.Rgb48 ? 3 : 4;
            int width = source.PixelWidth;
            int height = source.PixelHeight;
            int count = checked(width * height);
            var pixels = new ushort[checked(count * channels)];
            source.CopyPixels(pixels, checked(width * channels * sizeof(ushort)), 0);
            // RGB 共用一套窗口，不分别拉伸各通道；alpha 不参与统计。
            GetAutoDisplayLevels(pixels, channels, out double minimum, out double maximum);
            double scale = 256.0 / (maximum - minimum + 1.0);
            var output = new byte[checked(count * 4)];
            for (int i = 0; i < count; i++)
            {
                int input = i * channels;
                int target = i * 4;
                output[target] = ToDisplayByte(pixels[input + 2], minimum, scale);
                output[target + 1] = ToDisplayByte(pixels[input + 1], minimum, scale);
                output[target + 2] = ToDisplayByte(pixels[input], minimum, scale);
                output[target + 3] = channels == 4
                    ? (byte)((pixels[input + 3] * 255 + ushort.MaxValue / 2) / ushort.MaxValue)
                    : byte.MaxValue;
            }
            return BitmapSource.Create(width, height, DisplayDpi, DisplayDpi,
                PixelFormats.Bgra32, null, output, checked(width * 4));
        }

        /// <summary>
        /// 对齐 CAPilot capilot/utils/display_levels.py 的 auto_display_levels：
        /// ImageJ 首次 Auto，在实际数据范围内统计 256 桶直方图，而非固定百分位裁剪。
        /// </summary>
        private static void GetAutoDisplayLevels(ushort[] pixels, int channels,
            out double minimum, out double maximum)
        {
            int colorChannels = Math.Min(channels, 3);
            int dataMinimum = ushort.MaxValue;
            int dataMaximum = ushort.MinValue;
            int sampleCount = 0;
            for (int offset = 0; offset < pixels.Length; offset += channels)
            {
                // 完全透明的像素不属于可见画面，避免隐藏颜色干扰显示范围。
                if (channels == 4 && pixels[offset + 3] == 0) continue;
                for (int channel = 0; channel < colorChannels; channel++)
                {
                    int value = pixels[offset + channel];
                    dataMinimum = Math.Min(dataMinimum, value);
                    dataMaximum = Math.Max(dataMaximum, value);
                    sampleCount++;
                }
            }
            minimum = sampleCount == 0 ? 0 : dataMinimum;
            maximum = sampleCount == 0 ? 1 : dataMaximum;
            if (maximum <= minimum) maximum = minimum + 1;
            if (sampleCount == 0 || dataMaximum <= dataMinimum) return;

            const int histogramBins = 256;
            const int firstAutoThreshold = 5000;
            int range = dataMaximum - dataMinimum;
            var histogram = new int[histogramBins];
            for (int offset = 0; offset < pixels.Length; offset += channels)
            {
                if (channels == 4 && pixels[offset + 3] == 0) continue;
                for (int channel = 0; channel < colorChannels; channel++)
                {
                    int bin = (pixels[offset + channel] - dataMinimum) * histogramBins / range;
                    // 与 numpy.histogram 一致，最大值计入最后一桶。
                    histogram[Math.Min(histogramBins - 1, bin)]++;
                }
            }
            int threshold = sampleCount / firstAutoThreshold;
            int spikeLimit = sampleCount / 10;
            int first = -1;
            int last = -1;
            for (int bin = 0; bin < histogramBins; bin++)
            {
                int count = histogram[bin];
                if (count <= threshold || count > spikeLimit) continue;
                if (first < 0) first = bin;
                last = bin;
            }
            // 没有有效窗口或只有一桶时，回退到完整数据范围。
            if (first < 0 || last <= first) return;
            double binSize = range / (double)histogramBins;
            minimum = dataMinimum + first * binSize;
            maximum = dataMinimum + last * binSize;
        }

        private static byte ToDisplayByte(ushort value, double minimum, double scale)
        {
            // 对齐 CAPilot _colorize_channel / ImageJ ShortProcessor 的 8 位映射：
            // round((value - min) * 256 / (max - min + 1))，超出窗口的值截断。
            double mapped = Math.Round((value - minimum) * scale, MidpointRounding.ToEven);
            return (byte)Math.Max(0, Math.Min(byte.MaxValue, mapped));
        }

        public void Dispose()
        {
            if (!temporary) return;
            try { File.Delete(ImportPath); }
            catch (Exception ex) { Debug.WriteLine($"清理 TIFF 转换临时文件失败：{ex.Message}"); }
        }
    }
}

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace SlideSCI
{
    /// <summary>显示图像的线性 LUT 与 RGB 饱和相加；不修改亮度范围，也不执行图像配准。</summary>
    internal static class ChannelImageProcessor
    {
        // ImageJ Merge Channels 的 C1–C7 顺序。灰色 LUT 的亮端为白色。
        internal static readonly string[] ColorNames = { "红", "绿", "蓝", "灰", "青", "品红", "黄" };
        internal static readonly Color[] DefaultColors =
        {
            Color.Red, Color.Lime, Color.Blue, Color.White, Color.Cyan, Color.Magenta, Color.Yellow
        };

        internal static Bitmap Resize(Image image, Size size)
        {
            var result = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppArgb);
            result.SetResolution(288F, 288F);
            try
            {
                using (Graphics graphics = Graphics.FromImage(result))
                using (var attributes = new ImageAttributes())
                {
                    graphics.CompositingMode = CompositingMode.SourceCopy;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    attributes.SetWrapMode(WrapMode.TileFlipXY);
                    graphics.DrawImage(image, new Rectangle(Point.Empty, size), 0, 0,
                        image.Width, image.Height, GraphicsUnit.Pixel, attributes);
                }
                return result;
            }
            catch { result.Dispose(); throw; }
        }

        internal static Bitmap ApplyColor(Bitmap source, Color color, int sourcePeak = 255)
        {
            if (sourcePeak < 1 || sourcePeak > 255) sourcePeak = 255;
            Bitmap result = Resize(source, source.Size);
            try
            {
                WithPixels(result, ImageLockMode.ReadWrite, (data, row) =>
                {
                    for (int y = 0; y < result.Height; y++)
                    {
                        IntPtr address = IntPtr.Add(data.Scan0, y * data.Stride);
                        Marshal.Copy(address, row, 0, row.Length);
                        for (int x = 0; x < row.Length; x += 4)
                        {
                            // 单色通道取最大 RGB 分量，灰度值不变；已上色通道可直接换色。
                            int intensity = Math.Max(row[x], Math.Max(row[x + 1], row[x + 2]));
                            // 自定义深色 LUT 再次编辑时，恢复上次 LUT 的亮度比例。
                            intensity = Math.Min(255, (intensity * 255 + sourcePeak / 2) / sourcePeak);
                            // 与 ImageJ LUT.createLutFromColor 一样线性映射，并向下取整。
                            row[x] = (byte)(intensity * color.B / 255);
                            row[x + 1] = (byte)(intensity * color.G / 255);
                            row[x + 2] = (byte)(intensity * color.R / 255);
                        }
                        Marshal.Copy(row, 0, address, row.Length);
                    }
                });
                return result;
            }
            catch { result.Dispose(); throw; }
        }

        internal static Bitmap CreateMergeCanvas(Size size)
        {
            var result = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppArgb);
            try
            {
                result.SetResolution(288F, 288F);
                using (Graphics graphics = Graphics.FromImage(result)) graphics.Clear(Color.Black);
                return result;
            }
            catch { result.Dispose(); throw; }
        }

        /// <summary>逐张累加，避免同时保留七张全分辨率图。透明像素按 alpha 贡献到黑色背景。</summary>
        internal static void AddChannel(Bitmap target, Bitmap channel)
        {
            if (target.Size != channel.Size) throw new ArgumentException("通道图像尺寸必须一致。");
            BitmapData input = null;
            try
            {
                input = channel.LockBits(new Rectangle(Point.Empty, channel.Size), ImageLockMode.ReadOnly,
                    PixelFormat.Format32bppArgb);
                var sourceRow = new byte[channel.Width * 4];
                WithPixels(target, ImageLockMode.ReadWrite, (output, targetRow) =>
                {
                    for (int y = 0; y < target.Height; y++)
                    {
                        IntPtr address = IntPtr.Add(output.Scan0, y * output.Stride);
                        Marshal.Copy(address, targetRow, 0, targetRow.Length);
                        Marshal.Copy(IntPtr.Add(input.Scan0, y * input.Stride), sourceRow, 0, sourceRow.Length);
                        for (int x = 0; x < targetRow.Length; x += 4)
                            for (int component = 0; component < 3; component++)
                                targetRow[x + component] = (byte)Math.Min(255, targetRow[x + component] +
                                    (sourceRow[x + component] * sourceRow[x + 3] + 127) / 255);
                        Marshal.Copy(targetRow, 0, address, targetRow.Length);
                    }
                });
            }
            finally { if (input != null) channel.UnlockBits(input); }
        }

        private static void WithPixels(Bitmap bitmap, ImageLockMode mode, Action<BitmapData, byte[]> action)
        {
            BitmapData data = bitmap.LockBits(new Rectangle(Point.Empty, bitmap.Size), mode,
                PixelFormat.Format32bppArgb);
            try { action(data, new byte[bitmap.Width * 4]); }
            finally { bitmap.UnlockBits(data); }
        }
    }
}

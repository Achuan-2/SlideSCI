using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;

namespace SlideSCI
{
    /// <summary>仅在设置窗口绘制比例尺，不写入图片或幻灯片。图片由调用者负责释放。</summary>
    internal sealed class ScaleBarPreview : Control
    {
        private readonly Image image;
        private readonly SizeF pictureSize;
        private ScaleBarSettings settings;
        private ScaleBarLayout layout;
        private string message = "请填写 FOV 信息以预览比例尺。";

        internal ScaleBarPreview(Image image, SizeF pictureSize)
        {
            this.image = image;
            this.pictureSize = pictureSize;
            DoubleBuffered = true;
            ResizeRedraw = true;
            BackColor = Color.FromArgb(238, 241, 245);
            MinimumSize = new Size(360, 320);
        }

        /// <returns>空字符串表示参数可预览，否则返回可显示的输入提示。</returns>
        internal string UpdateSettings(ImageFieldOfView fov, SizeF visibleFraction, ScaleBarSettings options)
        {
            settings = options;
            layout = null;
            try
            {
                if (options.LengthMicrometers != 0 && !fov.IsValidFor(options.Vertical)) throw new InvalidOperationException(options.Vertical
                    ? "请填写大于 0 的 FOV 高度。" : "请填写大于 0 的 FOV 宽度。");
                layout = MeasureLayout(pictureSize,
                    new SizeF((float)(fov.WidthMicrometers * visibleFraction.Width),
                        (float)(fov.HeightMicrometers * visibleFraction.Height)), options);
                message = "";
            }
            catch (InvalidOperationException ex) { message = ex.Message; }
            Invalidate();
            return message;
        }

        private static Font CreateLabelFont(ScaleBarSettings options) =>
            new Font("Arial", options.FontSize, options.Bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel);

        internal static ScaleBarLayout MeasureLayout(SizeF pictureSize, SizeF visibleFov, ScaleBarSettings options)
        {
            if (options.LengthMicrometers == 0) return ScaleBarLayout.Calculate(pictureSize, visibleFov, options, SizeF.Empty);
            using (var bitmap = new Bitmap(1, 1))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            using (var font = CreateLabelFont(options))
            using (var format = StringFormat.GenericTypographic)
            {
                // 一个逻辑像素代表一个 pt，避免预览画布缩放影响文字测量。
                graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                SizeF textSize = options.ShowText ? graphics.MeasureString(options.Label, font, int.MaxValue, format) : SizeF.Empty;
                return ScaleBarLayout.Calculate(pictureSize, visibleFov, options, textSize);
            }
        }

        /// <summary>调用方负责将画布转换为图片局部坐标（pt）。</summary>
        internal static void DrawOverlay(Graphics graphics, ScaleBarLayout layout, ScaleBarSettings options)
        {
            if (options.LengthMicrometers == 0) return;
            graphics.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            using (var brush = new SolidBrush(Color.FromArgb(options.BarColorArgb)))
                graphics.FillRectangle(brush, layout.Bar);
            if (options.ShowText)
            {
                using (var font = CreateLabelFont(options))
                using (var brush = new SolidBrush(Color.FromArgb(options.FontColorArgb)))
                using (var format = StringFormat.GenericTypographic)
                    graphics.DrawString(options.Label, font, brush, layout.Text.Location, format);
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics graphics = e.Graphics;
            const int padding = 18;
            float scale = Math.Min((ClientSize.Width - 2 * padding) / pictureSize.Width,
                (ClientSize.Height - 2 * padding - 46) / pictureSize.Height);
            if (scale <= 0) return;
            var bounds = new RectangleF((ClientSize.Width - pictureSize.Width * scale) / 2,
                padding + (ClientSize.Height - 2 * padding - 46 - pictureSize.Height * scale) / 2,
                pictureSize.Width * scale, pictureSize.Height * scale);
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.DrawImage(image, bounds);
            if (layout != null)
            {
                GraphicsState state = graphics.Save();
                try
                {
                    graphics.SetClip(bounds);
                    graphics.TranslateTransform(bounds.Left, bounds.Top);
                    graphics.ScaleTransform(scale, scale);
                    DrawOverlay(graphics, layout, settings);
                }
                finally { graphics.Restore(state); }
            }
            using (var pen = new Pen(Color.FromArgb(185, 190, 200))) graphics.DrawRectangle(pen, bounds.X, bounds.Y, bounds.Width, bounds.Height);
            string hint = string.IsNullOrEmpty(message) ? "随参数实时更新；图片按未旋转方向预览。" : message;
            TextRenderer.DrawText(graphics, hint, Font,
                new Rectangle(padding, ClientSize.Height - 44, ClientSize.Width - 2 * padding, 40),
                string.IsNullOrEmpty(message) ? Color.DimGray : Color.Firebrick,
                TextFormatFlags.WordBreak | TextFormatFlags.HorizontalCenter);
        }
    }
}

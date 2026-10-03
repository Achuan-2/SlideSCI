using System;
using System.Drawing;

namespace SlideSCI
{
    /// <summary>预览与 PowerPoint 生成共用的比例尺布局，所有坐标均为图片局部坐标（pt）。</summary>
    internal sealed class ScaleBarLayout
    {
        internal RectangleF Bar { get; private set; }
        internal RectangleF Text { get; private set; }

        internal static ScaleBarLayout Calculate(SizeF imageSize, SizeF visibleFov,
            ScaleBarSettings settings, SizeF textSize)
        {
            if (!ImageFieldOfView.Positive(imageSize.Width) || !ImageFieldOfView.Positive(imageSize.Height) ||
                settings == null || !settings.IsValid)
                throw new InvalidOperationException("请填写有效的 FOV 和比例尺参数。");
            if (settings.LengthMicrometers == 0)
                return new ScaleBarLayout { Bar = RectangleF.Empty, Text = RectangleF.Empty };
            if (!ImageFieldOfView.Positive(settings.Vertical ? visibleFov.Height : visibleFov.Width))
                throw new InvalidOperationException("请填写有效的 FOV 和比例尺参数。");
            float length = (float)(settings.LengthMicrometers * (settings.Vertical
                ? imageSize.Height / visibleFov.Height : imageSize.Width / visibleFov.Width));
            float barWidth = settings.Vertical ? settings.ThicknessPoints : length;
            float barHeight = settings.Vertical ? length : settings.ThicknessPoints;
            float margin = Math.Min(8, Math.Min(imageSize.Width, imageSize.Height) * 0.04f);
            const float gap = 3;
            bool text = settings.ShowText;
            float blockWidth = settings.Vertical ? barWidth + (text ? gap + textSize.Width : 0)
                : Math.Max(barWidth, text ? textSize.Width : 0);
            float blockHeight = settings.Vertical ? Math.Max(barHeight, text ? textSize.Height : 0)
                : barHeight + (text ? gap + textSize.Height : 0);
            if (!ImageFieldOfView.Positive(length) || blockWidth + 2 * margin > imageSize.Width ||
                blockHeight + 2 * margin > imageSize.Height)
                throw new InvalidOperationException("比例尺或文字超出图片范围，请减小长度、厚度或字号。");
            bool right = settings.Corner == ScaleBarCorner.TopRight || settings.Corner == ScaleBarCorner.BottomRight;
            bool bottom = settings.Corner == ScaleBarCorner.BottomLeft || settings.Corner == ScaleBarCorner.BottomRight;
            // 比例尺本体锚定所选角落，文字宽高不参与它的边缘间距计算。
            float x = right ? imageSize.Width - margin - barWidth : margin;
            float y = bottom ? imageSize.Height - margin - barHeight : margin;
            var barBounds = new RectangleF(x, y, barWidth, barHeight);
            RectangleF textBounds = RectangleF.Empty;
            if (text)
            {
                // 文字放在图片内侧，与比例尺在所选角落的外侧边缘对齐。
                float textX = settings.Vertical
                    ? (right ? x - gap - textSize.Width : x + barWidth + gap)
                    : (right ? x + barWidth - textSize.Width : x);
                float textY = settings.Vertical
                    ? (bottom ? y + barHeight - textSize.Height : y)
                    : (bottom ? y - gap - textSize.Height : y + barHeight + gap);
                textBounds = new RectangleF(textX, textY, textSize.Width, textSize.Height);
            }
            return new ScaleBarLayout
            {
                Bar = barBounds,
                Text = textBounds
            };
        }
    }
}

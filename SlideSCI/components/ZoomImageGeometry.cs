using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    public enum ZoomImagePlacement
    {
        Right,
        Left,
        Top,
        Bottom
    }

    public enum ZoomGuideLineExtent
    {
        InsideSourceImage,
        AcrossImages
    }

    /// <summary>弹窗预览和幻灯片生成共用的布局计算，坐标单位由调用者决定。</summary>
    internal static class ZoomImageGeometry
    {
        public const float GapPoints = 0.5f * 72f / 2.54f;

        public static RectangleF GetZoomBounds(RectangleF source, SizeF regionSize,
            ZoomImagePlacement placement, float gap)
        {
            if (regionSize.Width <= 0 || regionSize.Height <= 0)
            {
                throw new ArgumentOutOfRangeException(nameof(regionSize), "放大区域必须有有效的宽度和高度。");
            }

            bool horizontal = placement == ZoomImagePlacement.Left || placement == ZoomImagePlacement.Right;
            float scale = horizontal ? source.Height / regionSize.Height : source.Width / regionSize.Width;
            float width = regionSize.Width * scale;
            float height = regionSize.Height * scale;
            switch (placement)
            {
                case ZoomImagePlacement.Left:
                    return new RectangleF(source.Left - gap - width, source.Top, width, height);
                case ZoomImagePlacement.Top:
                    return new RectangleF(source.Left, source.Top - gap - height, width, height);
                case ZoomImagePlacement.Bottom:
                    return new RectangleF(source.Left, source.Bottom + gap, width, height);
                default:
                    return new RectangleF(source.Right + gap, source.Top, width, height);
            }
        }

        // 角点按左上、右上、右下、左下排列，同时支持旋转的矩形或图片。
        public static PointF[] GetCorners(RectangleF bounds, float rotationDegrees = 0)
        {
            var corners = new[]
            {
                new PointF(bounds.Left, bounds.Top), new PointF(bounds.Right, bounds.Top),
                new PointF(bounds.Right, bounds.Bottom), new PointF(bounds.Left, bounds.Bottom)
            };
            double angle = rotationDegrees * Math.PI / 180;
            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            float centerX = bounds.Left + bounds.Width / 2f;
            float centerY = bounds.Top + bounds.Height / 2f;
            for (int i = 0; i < corners.Length; i++)
            {
                double x = corners[i].X - centerX;
                double y = corners[i].Y - centerY;
                corners[i] = new PointF(centerX + (float)(x * cos - y * sin),
                    centerY + (float)(x * sin + y * cos));
            }
            return corners;
        }

        // 两条辅助线的端点：起点一、终点一、起点二、终点二。
        public static PointF[] GetGuideEndpoints(PointF[] region, PointF[] zoom, ZoomImagePlacement placement)
        {
            switch (placement)
            {
                case ZoomImagePlacement.Left:
                    return new[] { region[0], zoom[1], region[3], zoom[2] };
                case ZoomImagePlacement.Top:
                    return new[] { region[0], zoom[3], region[1], zoom[2] };
                case ZoomImagePlacement.Bottom:
                    return new[] { region[3], zoom[0], region[2], zoom[1] };
                default:
                    return new[] { region[1], zoom[0], region[2], zoom[3] };
            }
        }

        public static ZoomImagePlacement GetSavedPlacement()
        {
            int saved = Properties.Settings.Default.ZoomImagePlacement;
            return Enum.IsDefined(typeof(ZoomImagePlacement), saved)
                ? (ZoomImagePlacement)saved : ZoomImagePlacement.Right;
        }

        public static ZoomGuideLineExtent GetSavedGuideLineExtent()
        {
            int saved = Properties.Settings.Default.ZoomGuideLineExtent;
            return Enum.IsDefined(typeof(ZoomGuideLineExtent), saved)
                ? (ZoomGuideLineExtent)saved : ZoomGuideLineExtent.AcrossImages;
        }

        public static ZoomImagePlacement GetRelativePlacement(RectangleF source, RectangleF zoom,
            ZoomImagePlacement fallback)
        {
            float dx = zoom.Left + zoom.Width / 2f - source.Left - source.Width / 2f;
            float dy = zoom.Top + zoom.Height / 2f - source.Top - source.Height / 2f;
            if (Math.Abs(dx) < 0.001f && Math.Abs(dy) < 0.001f) return fallback;
            return Math.Abs(dx) / Math.Max(0.01f, source.Width) >= Math.Abs(dy) / Math.Max(0.01f, source.Height)
                ? (dx >= 0 ? ZoomImagePlacement.Right : ZoomImagePlacement.Left)
                : (dy >= 0 ? ZoomImagePlacement.Bottom : ZoomImagePlacement.Top);
        }

        /// <summary>将矩形中心转入原图的局部坐标，保存与原图宽高的比例关系。</summary>
        public static RectangleF GetRelativeRegion(RectangleF source, float sourceRotation, RectangleF marker)
        {
            var sourceCenter = new PointF(source.Left + source.Width / 2f, source.Top + source.Height / 2f);
            var markerCenter = new PointF(marker.Left + marker.Width / 2f, marker.Top + marker.Height / 2f);
            PointF localCenter = RotatePoint(markerCenter, sourceCenter, -sourceRotation);
            return new RectangleF((localCenter.X - source.Left - marker.Width / 2f) / source.Width,
                (localCenter.Y - source.Top - marker.Height / 2f) / source.Height,
                marker.Width / source.Width, marker.Height / source.Height);
        }

        public static RectangleF GetRegionBounds(RectangleF source, float sourceRotation, RectangleF region)
        {
            float width = source.Width * region.Width, height = source.Height * region.Height;
            var sourceCenter = new PointF(source.Left + source.Width / 2f, source.Top + source.Height / 2f);
            var localCenter = new PointF(source.Left + (region.Left + region.Width / 2f) * source.Width,
                source.Top + (region.Top + region.Height / 2f) * source.Height);
            PointF center = RotatePoint(localCenter, sourceCenter, sourceRotation);
            return new RectangleF(center.X - width / 2f, center.Y - height / 2f, width, height);
        }

        /// <summary>将辅助线截到原图内部；先转入图片局部坐标，也支持旋转的原图。</summary>
        public static bool TryClipLineToRectangle(PointF start, PointF end, RectangleF bounds,
            float rotationDegrees, out PointF clippedStart, out PointF clippedEnd)
        {
            var center = new PointF(bounds.Left + bounds.Width / 2f, bounds.Top + bounds.Height / 2f);
            PointF localStart = RotatePoint(start, center, -rotationDegrees);
            PointF localEnd = RotatePoint(end, center, -rotationDegrees);
            float dx = localEnd.X - localStart.X;
            float dy = localEnd.Y - localStart.Y;
            float first = 0, last = 1;
            clippedStart = clippedEnd = PointF.Empty;
            if (bounds.Width <= 0 || bounds.Height <= 0 ||
                !ClipEdge(-dx, localStart.X - bounds.Left, ref first, ref last) ||
                !ClipEdge(dx, bounds.Right - localStart.X, ref first, ref last) ||
                !ClipEdge(-dy, localStart.Y - bounds.Top, ref first, ref last) ||
                !ClipEdge(dy, bounds.Bottom - localStart.Y, ref first, ref last) ||
                (last - first) * Math.Max(Math.Abs(dx), Math.Abs(dy)) < 0.0001f)
            {
                return false;
            }
            clippedStart = RotatePoint(new PointF(localStart.X + first * dx, localStart.Y + first * dy),
                center, rotationDegrees);
            clippedEnd = RotatePoint(new PointF(localStart.X + last * dx, localStart.Y + last * dy),
                center, rotationDegrees);
            return true;
        }

        private static bool ClipEdge(float direction, float distance, ref float first, ref float last)
        {
            if (Math.Abs(direction) < 0.000001f)
            {
                return distance >= 0;
            }
            float ratio = distance / direction;
            if (direction < 0)
            {
                if (ratio > last) return false;
                first = Math.Max(first, ratio);
            }
            else
            {
                if (ratio < first) return false;
                last = Math.Min(last, ratio);
            }
            return true;
        }

        public static PointF RotatePoint(PointF point, PointF center, float rotationDegrees)
        {
            double angle = rotationDegrees * Math.PI / 180;
            double cos = Math.Cos(angle), sin = Math.Sin(angle);
            float x = point.X - center.X, y = point.Y - center.Y;
            return new PointF(center.X + (float)(x * cos - y * sin),
                center.Y + (float)(x * sin + y * cos));
        }

        public static Office.MsoLineDashStyle GetOfficeDashStyle(DashStyle style)
        {
            switch (style)
            {
                case DashStyle.Dash: return Office.MsoLineDashStyle.msoLineDash;
                case DashStyle.Dot: return Office.MsoLineDashStyle.msoLineRoundDot;
                case DashStyle.DashDot: return Office.MsoLineDashStyle.msoLineDashDot;
                case DashStyle.DashDotDot: return Office.MsoLineDashStyle.msoLineDashDotDot;
                default: return Office.MsoLineDashStyle.msoLineSolid;
            }
        }

        public static DashStyle GetDrawingDashStyle(Office.MsoLineDashStyle style)
        {
            switch (style)
            {
                case Office.MsoLineDashStyle.msoLineDash:
                case Office.MsoLineDashStyle.msoLineLongDash:
                case Office.MsoLineDashStyle.msoLineSysDash:
                    return DashStyle.Dash;
                case Office.MsoLineDashStyle.msoLineRoundDot:
                case Office.MsoLineDashStyle.msoLineSquareDot:
                case Office.MsoLineDashStyle.msoLineSysDot:
                    return DashStyle.Dot;
                case Office.MsoLineDashStyle.msoLineDashDot:
                case Office.MsoLineDashStyle.msoLineLongDashDot:
                case Office.MsoLineDashStyle.msoLineSysDashDot:
                    return DashStyle.DashDot;
                case Office.MsoLineDashStyle.msoLineDashDotDot:
                case Office.MsoLineDashStyle.msoLineLongDashDotDot:
                    return DashStyle.DashDotDot;
                default:
                    return DashStyle.Solid;
            }
        }
    }
}

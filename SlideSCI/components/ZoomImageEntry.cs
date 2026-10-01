using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    /// <summary>弹窗中的一个放大区域；原始布局用于保留用户在幻灯片上调整过的位置。</summary>
    public sealed class ZoomImageEntry
    {
        public string RecordKey { get; }
        public string DisplayName { get; }
        public ZoomImageSelection Options { get; set; }
        public ZoomImageSelection InitialOptions { get; }
        public RectangleF? OriginalZoomRegion { get; }
        public float OriginalZoomRotation { get; }
        public int DisplayOrder => int.TryParse(DisplayName.Replace("放大图 ", ""), out int number) && number > 0
            ? number : int.MaxValue;

        public ZoomImageEntry(string recordKey, string name, ZoomImageSelection options,
            RectangleF? originalZoomRegion = null, float originalZoomRotation = 0)
        {
            RecordKey = recordKey;
            DisplayName = name;
            Options = InitialOptions = options;
            OriginalZoomRegion = originalZoomRegion;
            OriginalZoomRotation = originalZoomRotation;
        }

        public bool HasRegion => Options.Region.Width > 0 && Options.Region.Height > 0;
        public bool PreservesLayout => OriginalZoomRegion.HasValue && Options.Placement == InitialOptions.Placement;
        public override string ToString() => DisplayName;

        public static ZoomImageSelection WithRegion(ZoomImageSelection style, RectangleF region,
            float rotation = 0, bool edited = true) => new ZoomImageSelection(region,
                style.OutlineColor, style.OutlineWidthPoints, style.OutlineDashStyle,
                style.UseRectangleColorForZoomImage, style.AddGuideLines, style.Placement,
                rotation, edited, style.NativeOutlineDashStyle, style.GuideLineExtent);

        public static bool SameOptions(ZoomImageSelection first, ZoomImageSelection second) =>
            first.Region == second.Region && first.RegionRotationDegrees == second.RegionRotationDegrees &&
            first.OutlineColor.ToArgb() == second.OutlineColor.ToArgb() &&
            first.OutlineWidthPoints == second.OutlineWidthPoints &&
            (first.NativeOutlineDashStyle ?? ZoomImageGeometry.GetOfficeDashStyle(first.OutlineDashStyle)) ==
            (second.NativeOutlineDashStyle ?? ZoomImageGeometry.GetOfficeDashStyle(second.OutlineDashStyle)) &&
            first.UseRectangleColorForZoomImage == second.UseRectangleColorForZoomImage &&
            first.AddGuideLines == second.AddGuideLines && first.Placement == second.Placement &&
            first.GuideLineExtent == second.GuideLineExtent;
    }

    internal sealed class ZoomImageExistingObjects
    {
        public string Key { get; }
        public PowerPoint.Shape Marker { get; }
        public PowerPoint.Shape Zoom { get; }
        public int MarkerId { get; }
        public int? ZoomId { get; }
        public IList<PowerPoint.Shape> Guides { get; }
        public IList<int> GuideIds { get; }
        public ZoomImageEntry Entry { get; }
        public bool IsManualRectangle { get; }
        public bool NeedsImageRefresh { get; }

        public ZoomImageExistingObjects(string key, PowerPoint.Shape marker, PowerPoint.Shape zoom,
            IList<PowerPoint.Shape> guides, ZoomImageEntry entry, bool isManualRectangle = false, bool needsImageRefresh = false)
        {
            Key = key;
            Marker = marker;
            Zoom = zoom;
            MarkerId = marker.Id;
            ZoomId = zoom?.Id;
            Guides = guides;
            var guideIds = new List<int>();
            foreach (PowerPoint.Shape guide in guides) guideIds.Add(guide.Id);
            GuideIds = guideIds;
            Entry = entry;
            IsManualRectangle = isManualRectangle;
            NeedsImageRefresh = needsImageRefresh;
        }
    }

    internal static class ZoomImageSettingsCodec
    {
        public static string Serialize(ZoomImageSelection options, string name) => string.Join(";", new[]
        {
            "ZOOM", "4", options.OutlineColor.ToArgb().ToString(CultureInfo.InvariantCulture),
            options.OutlineWidthPoints.ToString("R", CultureInfo.InvariantCulture),
            ((int)options.OutlineDashStyle).ToString(CultureInfo.InvariantCulture),
            options.UseRectangleColorForZoomImage ? "1" : "0", options.AddGuideLines ? "1" : "0",
            ((int)options.Placement).ToString(CultureInfo.InvariantCulture),
            ((int)options.GuideLineExtent).ToString(CultureInfo.InvariantCulture),
            (options.NativeOutlineDashStyle.HasValue ? (int)options.NativeOutlineDashStyle.Value : -1)
                .ToString(CultureInfo.InvariantCulture),
            options.Region.X.ToString("R", CultureInfo.InvariantCulture),
            options.Region.Y.ToString("R", CultureInfo.InvariantCulture),
            options.Region.Width.ToString("R", CultureInfo.InvariantCulture),
            options.Region.Height.ToString("R", CultureInfo.InvariantCulture),
            options.RegionRotationDegrees.ToString("R", CultureInfo.InvariantCulture),
            Uri.EscapeDataString(name ?? "")
        });

        public static ZoomImageSelection Parse(string[] parts)
        {
            if (!((parts.Length == 10 && parts[1] == "3") || (parts.Length == 16 && parts[1] == "4")) ||
                !int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out int color) ||
                !float.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out float width) ||
                float.IsNaN(width) || float.IsInfinity(width) || width < 0.01f || width > 1000 ||
                !int.TryParse(parts[4], out int dash) || dash < 0 || dash > (int)DashStyle.DashDotDot ||
                (parts[5] != "0" && parts[5] != "1") || (parts[6] != "0" && parts[6] != "1") ||
                !int.TryParse(parts[7], out int placement) || !Enum.IsDefined(typeof(ZoomImagePlacement), placement) ||
                !int.TryParse(parts[8], out int extent) || !Enum.IsDefined(typeof(ZoomGuideLineExtent), extent) ||
                !int.TryParse(parts[9], out int native)) return null;
            Office.MsoLineDashStyle? nativeStyle = Enum.IsDefined(typeof(Office.MsoLineDashStyle), native) &&
                native != (int)Office.MsoLineDashStyle.msoLineDashStyleMixed ? (Office.MsoLineDashStyle?)native : null;
            RectangleF region = RectangleF.Empty;
            float rotation = 0;
            if (parts[1] == "4")
            {
                var values = new float[5];
                for (int i = 0; i < values.Length; i++)
                    if (!float.TryParse(parts[i + 10], NumberStyles.Float, CultureInfo.InvariantCulture, out values[i]) ||
                        float.IsNaN(values[i]) || float.IsInfinity(values[i])) return null;
                if (values[2] <= 0 || values[3] <= 0) return null;
                region = new RectangleF(values[0], values[1], values[2], values[3]);
                rotation = values[4];
            }
            return new ZoomImageSelection(region, Color.FromArgb(color), width, (DashStyle)dash,
                parts[5] == "1", parts[6] == "1", (ZoomImagePlacement)placement, rotation, false,
                nativeStyle, (ZoomGuideLineExtent)extent);
        }

        public static string ReadName(string[] parts)
        {
            if (parts.Length != 16 || parts[1] != "4") return null;
            try { return Uri.UnescapeDataString(parts[15]); }
            catch (UriFormatException) { return null; }
        }

        public static bool SameCrop(ZoomImageSelection first, ZoomImageSelection second)
        {
            if (first == null || first.Region.Width <= 0 || first.Region.Height <= 0) return false;
            RectangleF a = first.Region, b = second.Region;
            float angle = Math.Abs((first.RegionRotationDegrees - second.RegionRotationDegrees) % 360);
            return Math.Abs(a.X - b.X) < 0.00001f && Math.Abs(a.Y - b.Y) < 0.00001f &&
                Math.Abs(a.Width - b.Width) < 0.00001f && Math.Abs(a.Height - b.Height) < 0.00001f &&
                Math.Min(angle, 360 - angle) < 0.001f;
        }
    }

    /// <summary>预览与生成共同使用的多图布局，新增区域沿选定方向向外排布，避免互相遮挡。</summary>
    internal static class ZoomImageLayout
    {
        public static Dictionary<ZoomImageEntry, RectangleF> Calculate(RectangleF source,
            IList<ZoomImageEntry> entries, float gap, Func<ZoomImageEntry, SizeF> getCropSize = null)
        {
            var result = new Dictionary<ZoomImageEntry, RectangleF>();
            // 已有布局先占位，避免新增图覆盖后面才加载的已有图。
            foreach (ZoomImageEntry entry in entries)
            {
                if (!entry.HasRegion || !entry.PreservesLayout) continue;
                RectangleF old = entry.OriginalZoomRegion.Value;
                var bounds = new RectangleF(source.Left + old.Left * source.Width,
                    source.Top + old.Top * source.Height, old.Width * source.Width, old.Height * source.Height);
                if (entry.Options.Region != entry.InitialOptions.Region ||
                    entry.Options.RegionRotationDegrees != entry.InitialOptions.RegionRotationDegrees)
                {
                    SizeF crop = GetCropSize(source, entry, getCropSize);
                    if (crop.Width <= 0 || crop.Height <= 0) continue;
                    if (entry.Options.Placement == ZoomImagePlacement.Left || entry.Options.Placement == ZoomImagePlacement.Right)
                        bounds.Width = bounds.Height * crop.Width / crop.Height;
                    else bounds.Height = bounds.Width * crop.Height / crop.Width;
                }
                result.Add(entry, bounds);
            }
            foreach (ZoomImageEntry entry in entries)
            {
                if (!entry.HasRegion || result.ContainsKey(entry)) continue;
                SizeF crop = GetCropSize(source, entry, getCropSize);
                if (crop.Width <= 0 || crop.Height <= 0) continue;
                RectangleF bounds = ZoomImageGeometry.GetZoomBounds(source, crop, entry.Options.Placement, gap);
                for (int attempt = 0; attempt <= result.Count; attempt++)
                {
                    bool moved = false;
                    foreach (RectangleF occupied in result.Values)
                    {
                        if (!bounds.IntersectsWith(occupied)) continue;
                        switch (entry.Options.Placement)
                        {
                            case ZoomImagePlacement.Left: bounds.X = occupied.Left - gap - bounds.Width; break;
                            case ZoomImagePlacement.Top: bounds.Y = occupied.Top - gap - bounds.Height; break;
                            case ZoomImagePlacement.Bottom: bounds.Y = occupied.Bottom + gap; break;
                            default: bounds.X = occupied.Right + gap; break;
                        }
                        moved = true;
                    }
                    if (!moved) break;
                }
                result.Add(entry, bounds);
            }
            return result;
        }

        private static SizeF GetCropSize(RectangleF source, ZoomImageEntry entry, Func<ZoomImageEntry, SizeF> getCropSize)
        {
            if (getCropSize != null) return getCropSize(entry);
            RectangleF region = entry.Options.Region;
            var crop = new RectangleF(source.Left + region.Left * source.Width, source.Top + region.Top * source.Height,
                region.Width * source.Width, region.Height * source.Height);
            if (entry.Options.RegionRotationDegrees == 0) crop = RectangleF.Intersect(source, crop);
            return crop.Size;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    /// <summary>
    /// 通过形状标签保存放大图的关联。PowerPoint 没有通用的形状位置变更事件，
    /// 因此在 UI 线程定期检查当前页，先让放大矩形跟随原图，再更新辅助线。
    /// </summary>
    internal sealed class ZoomGuideLineTracker : IDisposable
    {
        private const string TagPrefix = "SLIDESCI_ZOOM_LINK_";
        private readonly PowerPoint.Application application;
        private readonly Timer timer;
        private readonly List<TrackedGroup> groups = new List<TrackedGroup>();
        private readonly HashSet<string> cleanedOrphanGroups = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private int cachedWindowHandle;
        private int cachedSlideId = -1;
        private int cachedShapeCount = -1;
        private DateTime nextDiscovery = DateTime.MinValue;
        private bool updating;
        private bool disposed;
        private int suspensionCount;
        private string lastError;

        public ZoomGuideLineTracker(PowerPoint.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            application.WindowSelectionChange += OnSelectionChanged;
            timer = new Timer { Interval = 250 };
            timer.Tick += (sender, e) => UpdateActiveSlide();
            timer.Start();
        }

        // 标签随演示文稿保存；用关联键而不是 Shape.Id，复制整页后仍能重新找到对象。
        public static string Attach(PowerPoint.Shape picture, PowerPoint.Shape marker, PowerPoint.Shape zoom,
            IList<PowerPoint.Shape> guides, ZoomGuideLineExtent extent, ZoomImagePlacement placement,
            ZoomImageSelection options = null, string name = null)
        {
            string key = TagPrefix + Guid.NewGuid().ToString("N").ToUpperInvariant();
            var taggedShapes = new List<PowerPoint.Shape>();
            try
            {
                AddTag(picture, key, "SOURCE", taggedShapes);
                AddTag(marker, key, MarkerAnchor.Capture(new GeometryState(picture, marker, zoom)).ToMetadata(),
                    taggedShapes);
                AddTag(zoom, key, options == null ? "ZOOM" : ZoomImageSettingsCodec.Serialize(options, name), taggedShapes);
                for (int i = 0; i < guides.Count; i++)
                {
                    string metadata = string.Format(CultureInfo.InvariantCulture, "GUIDE;{0};{1};{2}",
                        i, (int)extent, (int)placement);
                    AddTag(guides[i], key, metadata, taggedShapes);
                }
            }
            catch
            {
                // 生成失败时不在用户原有的图片或矩形上留下半成品关联。
                foreach (PowerPoint.Shape shape in taggedShapes)
                {
                    try { shape.Tags.Delete(key); }
                    catch (Exception ex) { Debug.WriteLine($"清理放大图关联失败: {ex.Message}"); }
                }
                throw;
            }
            return key;
        }

        public void RefreshNow() => UpdateActiveSlide(true);

        public IDisposable Suspend()
        {
            suspensionCount++;
            return new Suspension(() => { suspensionCount--; InvalidateCache(); });
        }

        private sealed class Suspension : IDisposable
        {
            private Action release;
            public Suspension(Action release) { this.release = release; }
            public void Dispose() { Action action = release; release = null; action?.Invoke(); }
        }

        public static IList<ZoomImageExistingObjects> ReadExisting(PowerPoint.Slide slide, PowerPoint.Shape picture)
        {
            var discovered = new Dictionary<string, GroupBuilder>(StringComparer.OrdinalIgnoreCase);
            foreach (PowerPoint.Shape shape in slide.Shapes) DiscoverShapeLinks(shape, false, discovered);
            var result = new List<ZoomImageExistingObjects>();
            var reservedNames = new HashSet<string>(discovered.Values.Where(group => group.Picture != null &&
                group.Picture.Id == picture.Id && !string.IsNullOrEmpty(group.DisplayName)).Select(group => group.DisplayName));
            var assignedNames = new HashSet<string>();
            var source = new RectangleF(picture.Left, picture.Top, picture.Width, picture.Height);
            foreach (var item in discovered.OrderBy(pair => pair.Value.Zoom == null ? int.MaxValue : pair.Value.Zoom.Id))
            {
                GroupBuilder group = item.Value;
                if (group.Ambiguous || group.HasGroupedShapes || group.Picture == null ||
                    group.Picture.Id != picture.Id || group.Marker == null || group.Zoom == null) continue;
                PowerPoint.Shape marker = group.Marker, zoom = group.Zoom;
                RectangleF markerBounds = new RectangleF(marker.Left, marker.Top, marker.Width, marker.Height);
                RectangleF region = ZoomImageGeometry.GetRelativeRegion(source, picture.Rotation, markerBounds);
                Office.MsoLineDashStyle dash = marker.Line.DashStyle;
                if (dash == Office.MsoLineDashStyle.msoLineDashStyleMixed) dash = Office.MsoLineDashStyle.msoLineSolid;
                ZoomImageSelection saved = group.ZoomSettings;
                bool useColor = saved?.UseRectangleColorForZoomImage ??
                    (zoom.Line.Visible == Office.MsoTriState.msoTrue && zoom.Line.ForeColor.RGB == marker.Line.ForeColor.RGB);
                ZoomImagePlacement placement = saved?.Placement ?? (group.Guides.Count > 0
                    ? group.Guides[0].InitialPlacement : ZoomImageGeometry.GetRelativePlacement(source,
                        new RectangleF(zoom.Left, zoom.Top, zoom.Width, zoom.Height), ZoomImagePlacement.Right));
                var options = new ZoomImageSelection(region, ColorTranslator.FromOle(marker.Line.ForeColor.RGB),
                    Math.Max(0.01f, marker.Line.Weight), ZoomImageGeometry.GetDrawingDashStyle(dash),
                    useColor, saved?.AddGuideLines ?? group.Guides.Count > 0, placement,
                    marker.Rotation - picture.Rotation, false, dash,
                    saved?.GuideLineExtent ?? (group.Guides.Count > 0 ? group.Guides[0].Extent : ZoomGuideLineExtent.AcrossImages));
                var zoomRegion = new RectangleF((zoom.Left - picture.Left) / picture.Width,
                    (zoom.Top - picture.Top) / picture.Height, zoom.Width / picture.Width, zoom.Height / picture.Height);
                string name = group.DisplayName;
                if (string.IsNullOrEmpty(name) || assignedNames.Contains(name))
                {
                    int number = 1;
                    while (reservedNames.Contains($"放大图 {number}") || assignedNames.Contains($"放大图 {number}")) number++;
                    name = $"放大图 {number}";
                }
                assignedNames.Add(name);
                var entry = new ZoomImageEntry(item.Key, name, options, zoomRegion, zoom.Rotation);
                result.Add(new ZoomImageExistingObjects(item.Key, marker, zoom,
                    group.Guides.Select(guide => guide.Shape).ToArray(), entry, false,
                    !ZoomImageSettingsCodec.SameCrop(saved, options)));
            }
            result.Sort((first, second) => first.Entry.DisplayOrder.CompareTo(second.Entry.DisplayOrder));
            return result;
        }

        public static void RemoveLink(string key, PowerPoint.Shape picture, PowerPoint.Shape marker,
            PowerPoint.Shape zoom, IEnumerable<PowerPoint.Shape> guides)
        {
            var shapes = new List<PowerPoint.Shape> { picture, marker, zoom };
            if (guides != null) shapes.AddRange(guides);
            foreach (PowerPoint.Shape shape in shapes)
                if (shape != null) shape.Tags.Delete(key);
        }

        private static void AddTag(PowerPoint.Shape shape, string key, string value,
            List<PowerPoint.Shape> taggedShapes)
        {
            taggedShapes.Add(shape);
            shape.Tags.Add(key, value);
        }

        public static void RemoveCopiedLinks(PowerPoint.Shape shape)
        {
            PowerPoint.Tags tags = shape.Tags;
            for (int i = tags.Count; i >= 1; i--)
            {
                string key = tags.Name(i);
                if (key.StartsWith(TagPrefix, StringComparison.OrdinalIgnoreCase)) tags.Delete(key);
            }
        }

        public void InvalidateCache() => nextDiscovery = DateTime.MinValue;

        private void OnSelectionChanged(PowerPoint.Selection selection)
        {
            InvalidateCache();
            UpdateActiveSlide();
        }

        private void UpdateActiveSlide(bool force = false)
        {
            // 拖动期间不争用 PowerPoint 的编辑操作；放开鼠标后在下一次检查中更新。
            if (disposed || updating || suspensionCount > 0 || (!force && Control.MouseButtons != MouseButtons.None)) return;
            updating = true;
            try
            {
                PowerPoint.DocumentWindow window = application.ActiveWindow;
                if (window == null || (window.View.Type != PowerPoint.PpViewType.ppViewNormal &&
                    window.View.Type != PowerPoint.PpViewType.ppViewSlide))
                {
                    ClearCache();
                    return;
                }
                var slide = window.View.Slide as PowerPoint.Slide;
                if (slide == null) return;
                if (window.HWND != cachedWindowHandle || slide.SlideID != cachedSlideId)
                {
                    ClearCache();
                    cachedWindowHandle = window.HWND;
                    cachedSlideId = slide.SlideID;
                }
                int count = slide.Shapes.Count;
                if (count != cachedShapeCount || DateTime.UtcNow >= nextDiscovery)
                {
                    DiscoverGroups(slide);
                    cachedShapeCount = slide.Shapes.Count;
                    nextDiscovery = DateTime.UtcNow.AddSeconds(2);
                }
                foreach (TrackedGroup group in groups) UpdateGroup(group);
                lastError = null;
            }
            catch (Exception ex)
            {
                // 对象删除、文稿关闭或 Office 暂时繁忙时，下次重新发现关联，不打断编辑。
                InvalidateCache();
                if (lastError != ex.Message)
                {
                    Debug.WriteLine($"更新放大图辅助线失败: {ex.Message}");
                    lastError = ex.Message;
                }
            }
            finally { updating = false; }
        }

        private void DiscoverGroups(PowerPoint.Slide slide)
        {
            var discovered = new Dictionary<string, GroupBuilder>(StringComparer.OrdinalIgnoreCase);
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                DiscoverShapeLinks(shape, false, discovered);
            }

            var previous = new Dictionary<string, TrackedGroup>(StringComparer.OrdinalIgnoreCase);
            foreach (TrackedGroup group in groups) previous[group.Key] = group;
            // 旧版可能让多个放大图共用一个矩形；仍有放大图使用时保留它。
            // 提前保存 ID，避免多个已删除的放大图共享同一矩形时访问失效的 COM 对象。
            var markerIdsByGroup = discovered.Values.Where(builder => builder.Marker != null)
                .ToDictionary(builder => builder, builder => builder.Marker.Id);
            var liveMarkerIds = new HashSet<int>(markerIdsByGroup.Where(pair => pair.Key.Zoom != null)
                .Select(pair => pair.Value));
            var deletedMarkerIds = new HashSet<int>();
            var updated = new List<TrackedGroup>();
            foreach (var entry in discovered)
            {
                GroupBuilder builder = entry.Value;
                // 同一页粘贴整个关联组可能保留重复标签；有歧义时不把线连到另一组图片。
                if (builder.Ambiguous) continue;
                string orphanKey = string.Format(CultureInfo.InvariantCulture, "{0}:{1}:{2}",
                    cachedWindowHandle, cachedSlideId, entry.Key);
                if (builder.Zoom == null)
                {
                    // 放大图已删除，清理对应矩形与辅助线，保留原图。
                    // 每次删除只清理一次，避免用户撤销恢复对象后又被立即删除。
                    if ((builder.Marker != null || builder.Guides.Count > 0) &&
                        !cleanedOrphanGroups.Contains(orphanKey))
                    {
                        foreach (Guide guide in builder.Guides) guide.Shape.Delete();
                        if (builder.Marker != null)
                        {
                            int markerId = markerIdsByGroup[builder];
                            if (!liveMarkerIds.Contains(markerId) && deletedMarkerIds.Add(markerId))
                                builder.Marker.Delete();
                        }
                        cleanedOrphanGroups.Add(orphanKey);
                    }
                    continue;
                }
                else cleanedOrphanGroups.Remove(orphanKey);
                // 组合内的局部坐标不能直接套用幻灯片坐标；仍识别标签，避免把分组误判为删除。
                if (builder.Picture == null || builder.Marker == null || builder.HasGroupedShapes) continue;
                var group = new TrackedGroup(entry.Key, builder);
                if (previous.TryGetValue(entry.Key, out TrackedGroup old) && group.HasSameShapes(old))
                {
                    group.LastState = old.LastState;
                    group.PendingMarkerUpdate = old.PendingMarkerUpdate;
                }
                updated.Add(group);
            }
            groups.Clear();
            groups.AddRange(updated);
        }

        private static void DiscoverShapeLinks(PowerPoint.Shape shape, bool inGroup,
            Dictionary<string, GroupBuilder> discovered)
        {
            PowerPoint.Tags tags = shape.Tags;
            for (int i = 1; i <= tags.Count; i++)
            {
                string key = tags.Name(i);
                if (!key.StartsWith(TagPrefix, StringComparison.OrdinalIgnoreCase)) continue;
                if (!discovered.TryGetValue(key, out GroupBuilder builder))
                {
                    builder = new GroupBuilder();
                    discovered.Add(key, builder);
                }
                builder.Add(shape, tags.Value(i));
                builder.HasGroupedShapes |= inGroup;
            }
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape child in shape.GroupItems) DiscoverShapeLinks(child, true, discovered);
            }
        }

        private static void UpdateGroup(TrackedGroup group)
        {
            var state = new GeometryState(group.Picture, group.Marker, group.Zoom);
            state = SynchronizeMarker(group, state);
            if (group.LastState != null && state.EqualsGeometry(group.LastState)) return;
            if (group.Guides.Count == 0)
            {
                group.LastState = state;
                return;
            }
            ZoomImagePlacement placement = ZoomImageGeometry.GetRelativePlacement(state.PictureBounds,
                state.ZoomBounds, group.Guides[0].InitialPlacement);
            PointF[] endpoints = ZoomImageGeometry.GetGuideEndpoints(
                ZoomImageGeometry.GetCorners(state.MarkerBounds, state.MarkerRotation),
                ZoomImageGeometry.GetCorners(state.ZoomBounds, state.ZoomRotation), placement);
            foreach (Guide guide in group.Guides)
            {
                PointF start = endpoints[guide.Index * 2], end = endpoints[guide.Index * 2 + 1];
                bool visible = guide.Extent != ZoomGuideLineExtent.InsideSourceImage ||
                    ZoomImageGeometry.TryClipLineToRectangle(start, end, state.PictureBounds,
                        state.PictureRotation, out start, out end);
                visible = visible && (Math.Abs(end.X - start.X) > 0.0001f ||
                    Math.Abs(end.Y - start.Y) > 0.0001f);
                if (visible) SetLineEndpoints(guide.Shape, start, end);
                Office.MsoTriState visibility = visible ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (guide.Shape.Visible != visibility) guide.Shape.Visible = visibility;
            }
            // 全部完成才记下状态；COM 操作中断时，下次可重试未完成的部分。
            group.LastState = state;
        }

        private static GeometryState SynchronizeMarker(TrackedGroup group, GeometryState state)
        {
            MarkerAnchor anchor = group.Anchor;
            if (anchor == null)
            {
                // 兼容旧版只有 MARKER 角色的关联，从当前坐标建立跟随基准。
                SaveMarkerAnchor(group, state);
                return state;
            }

            RectangleF oldMarker = anchor.GetMarkerBounds(anchor.PictureBounds, anchor.PictureRotation);
            float oldRotation = anchor.PictureRotation + anchor.RotationOffset;
            bool sourceChanged = !SameBounds(state.PictureBounds, anchor.PictureBounds) ||
                !SameRotation(state.PictureRotation, anchor.PictureRotation);
            bool markerUnchanged = SameBounds(state.MarkerBounds, oldMarker) &&
                SameRotation(state.MarkerRotation, oldRotation);

            if (group.PendingMarkerUpdate || (sourceChanged && markerUnchanged))
            {
                // 先记下待完成状态，Office 繁忙导致操作中断时，下次仍按同一基准重试。
                group.PendingMarkerUpdate = true;
                SetMarkerBounds(group.Marker, anchor.GetMarkerBounds(state.PictureBounds, state.PictureRotation),
                    state.PictureRotation + anchor.RotationOffset);
                state = new GeometryState(group.Picture, group.Marker, group.Zoom);
                SaveMarkerAnchor(group, state);
            }
            else if (sourceChanged || !markerUnchanged)
            {
                // 用户同时移动了原图与矩形，或单独编辑了矩形：直接接受当前位置，
                // 不再叠加原图的位移，并用新的相对位置作为以后跟随的基准。
                SaveMarkerAnchor(group, state);
            }
            return state;
        }

        private static void SaveMarkerAnchor(TrackedGroup group, GeometryState state)
        {
            MarkerAnchor anchor = MarkerAnchor.Capture(state);
            group.Marker.Tags.Add(group.Key, anchor.ToMetadata());
            group.Anchor = anchor;
            group.PendingMarkerUpdate = false;
        }

        private static void SetMarkerBounds(PowerPoint.Shape marker, RectangleF bounds, float rotation)
        {
            Office.MsoTriState aspect = marker.LockAspectRatio;
            marker.LockAspectRatio = Office.MsoTriState.msoFalse;
            try
            {
                marker.Width = bounds.Width;
                marker.Height = bounds.Height;
                marker.Rotation = (rotation % 360 + 360) % 360;
                marker.Left = bounds.Left;
                marker.Top = bounds.Top;
            }
            finally { marker.LockAspectRatio = aspect; }
        }

        private static bool SameBounds(RectangleF first, RectangleF second) =>
            Math.Abs(first.X - second.X) < 0.001f && Math.Abs(first.Y - second.Y) < 0.001f &&
            Math.Abs(first.Width - second.Width) < 0.001f && Math.Abs(first.Height - second.Height) < 0.001f;

        private static bool SameRotation(float first, float second)
        {
            float difference = Math.Abs((first - second) % 360);
            return Math.Min(difference, 360 - difference) < 0.001f;
        }

        private static void SetLineEndpoints(PowerPoint.Shape line, PointF start, PointF end)
        {
            float dx = end.X - start.X, dy = end.Y - start.Y;
            float width = (float)Math.Sqrt(dx * dx + dy * dy);
            if (width < 0.0001f) return;
            float rotation = (float)((Math.Atan2(dy, dx) * 180 / Math.PI + 360) % 360);
            float left = (start.X + end.X - width) / 2f;
            float top = (start.Y + end.Y) / 2f;
            // 不替换线对象，保留用户修改的颜色、线宽、线型、层级及形状标签。
            if (Math.Abs(line.Left - left) < 0.001f && Math.Abs(line.Top - top) < 0.001f &&
                Math.Abs(line.Width - width) < 0.001f && Math.Abs(line.Height) < 0.001f &&
                Math.Abs(line.Rotation - rotation) < 0.001f) return;
            line.LockAspectRatio = Office.MsoTriState.msoFalse;
            line.Rotation = 0;
            line.Width = width;
            line.Height = 0;
            line.Rotation = rotation;
            line.Left = left;
            line.Top = top;
        }

        private void ClearCache()
        {
            groups.Clear();
            cachedSlideId = -1;
            cachedWindowHandle = 0;
            cachedShapeCount = -1;
            InvalidateCache();
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            timer.Stop();
            timer.Dispose();
            application.WindowSelectionChange -= OnSelectionChanged;
            ClearCache();
            cleanedOrphanGroups.Clear();
        }

        private sealed class Guide
        {
            public PowerPoint.Shape Shape { get; }
            public int ShapeId { get; }
            public int Index { get; }
            public ZoomGuideLineExtent Extent { get; }
            public ZoomImagePlacement InitialPlacement { get; }

            public Guide(PowerPoint.Shape shape, int index, ZoomGuideLineExtent extent, ZoomImagePlacement placement)
            {
                Shape = shape;
                ShapeId = shape.Id;
                Index = index;
                Extent = extent;
                InitialPlacement = placement;
            }
        }

        private sealed class GroupBuilder
        {
            public PowerPoint.Shape Picture;
            public PowerPoint.Shape Marker;
            public PowerPoint.Shape Zoom;
            public bool Ambiguous;
            public bool HasGroupedShapes;
            public MarkerAnchor Anchor;
            public ZoomImageSelection ZoomSettings;
            public string DisplayName;
            public readonly List<Guide> Guides = new List<Guide>();

            public void Add(PowerPoint.Shape shape, string metadata)
            {
                switch (metadata)
                {
                    case "SOURCE": SetRole(ref Picture, shape); return;
                    case "MARKER": SetRole(ref Marker, shape); return;
                    case "ZOOM": SetRole(ref Zoom, shape); return;
                }
                string[] parts = metadata.Split(';');
                if (parts.Length > 1 && parts[0] == "ZOOM")
                {
                    SetRole(ref Zoom, shape);
                    ZoomSettings = ZoomImageSettingsCodec.Parse(parts);
                    DisplayName = ZoomImageSettingsCodec.ReadName(parts);
                    return;
                }
                if (parts.Length > 1 && parts[0] == "MARKER")
                {
                    SetRole(ref Marker, shape);
                    Anchor = MarkerAnchor.Parse(parts);
                    return;
                }
                if (parts.Length == 4 && parts[0] == "GUIDE" &&
                    int.TryParse(parts[1], out int index) && index >= 0 && index < 2 &&
                    int.TryParse(parts[2], out int extent) && Enum.IsDefined(typeof(ZoomGuideLineExtent), extent) &&
                    int.TryParse(parts[3], out int placement) && Enum.IsDefined(typeof(ZoomImagePlacement), placement))
                {
                    Guides.Add(new Guide(shape, index, (ZoomGuideLineExtent)extent, (ZoomImagePlacement)placement));
                }
            }

            private void SetRole(ref PowerPoint.Shape role, PowerPoint.Shape shape)
            {
                if (role != null) Ambiguous = true;
                else role = shape;
            }
        }

        private sealed class TrackedGroup
        {
            public string Key { get; }
            public PowerPoint.Shape Picture { get; }
            public PowerPoint.Shape Marker { get; }
            public PowerPoint.Shape Zoom { get; }
            private readonly int pictureId;
            private readonly int markerId;
            private readonly int zoomId;
            public List<Guide> Guides { get; }
            public GeometryState LastState { get; set; }
            public MarkerAnchor Anchor { get; set; }
            public bool PendingMarkerUpdate { get; set; }

            public TrackedGroup(string key, GroupBuilder builder)
            {
                Key = key;
                Picture = builder.Picture;
                Marker = builder.Marker;
                Zoom = builder.Zoom;
                pictureId = Picture.Id;
                markerId = Marker.Id;
                zoomId = Zoom?.Id ?? 0;
                Guides = builder.Guides;
                Anchor = builder.Anchor;
            }

            public bool HasSameShapes(TrackedGroup other)
            {
                if (pictureId != other.pictureId || markerId != other.markerId || zoomId != other.zoomId ||
                    Guides.Count != other.Guides.Count) return false;
                for (int i = 0; i < Guides.Count; i++)
                {
                    if (Guides[i].ShapeId != other.Guides[i].ShapeId) return false;
                }
                return true;
            }
        }

        /// <summary>相对区域和上次原图坐标一并保存，以便文件重开后仍可判断是谁发生变化。</summary>
        private sealed class MarkerAnchor
        {
            public RectangleF Region { get; }
            public float RotationOffset { get; }
            public RectangleF PictureBounds { get; }
            public float PictureRotation { get; }

            private MarkerAnchor(RectangleF region, float rotationOffset, RectangleF pictureBounds, float pictureRotation)
            {
                Region = region;
                RotationOffset = rotationOffset;
                PictureBounds = pictureBounds;
                PictureRotation = pictureRotation;
            }

            public static MarkerAnchor Capture(GeometryState state) => new MarkerAnchor(
                ZoomImageGeometry.GetRelativeRegion(state.PictureBounds, state.PictureRotation, state.MarkerBounds),
                state.MarkerRotation - state.PictureRotation, state.PictureBounds, state.PictureRotation);

            public RectangleF GetMarkerBounds(RectangleF pictureBounds, float pictureRotation) =>
                ZoomImageGeometry.GetRegionBounds(pictureBounds, pictureRotation, Region);

            public string ToMetadata()
            {
                float[] values = { Region.X, Region.Y, Region.Width, Region.Height, RotationOffset,
                    PictureBounds.X, PictureBounds.Y, PictureBounds.Width, PictureBounds.Height, PictureRotation };
                var parts = new string[values.Length + 2];
                parts[0] = "MARKER";
                parts[1] = "2";
                for (int i = 0; i < values.Length; i++)
                    parts[i + 2] = values[i].ToString("R", CultureInfo.InvariantCulture);
                return string.Join(";", parts);
            }

            public static MarkerAnchor Parse(string[] parts)
            {
                if (parts.Length != 12 || parts[1] != "2") return null;
                var values = new float[10];
                for (int i = 0; i < values.Length; i++)
                {
                    if (!float.TryParse(parts[i + 2], NumberStyles.Float, CultureInfo.InvariantCulture, out values[i]) ||
                        float.IsNaN(values[i]) || float.IsInfinity(values[i])) return null;
                }
                if (values[2] <= 0 || values[3] <= 0 || values[7] <= 0 || values[8] <= 0) return null;
                return new MarkerAnchor(new RectangleF(values[0], values[1], values[2], values[3]), values[4],
                    new RectangleF(values[5], values[6], values[7], values[8]), values[9]);
            }
        }

        private sealed class GeometryState
        {
            public RectangleF PictureBounds { get; }
            public RectangleF MarkerBounds { get; }
            public RectangleF ZoomBounds { get; }
            public float PictureRotation { get; }
            public float MarkerRotation { get; }
            public float ZoomRotation { get; }

            public GeometryState(PowerPoint.Shape picture, PowerPoint.Shape marker, PowerPoint.Shape zoom)
            {
                PictureBounds = GetBounds(picture);
                MarkerBounds = GetBounds(marker);
                ZoomBounds = zoom == null ? RectangleF.Empty : GetBounds(zoom);
                PictureRotation = picture.Rotation;
                MarkerRotation = marker.Rotation;
                ZoomRotation = zoom?.Rotation ?? 0;
            }

            private static RectangleF GetBounds(PowerPoint.Shape shape) =>
                new RectangleF(shape.Left, shape.Top, shape.Width, shape.Height);

            public bool EqualsGeometry(GeometryState other) => PictureBounds == other.PictureBounds &&
                MarkerBounds == other.MarkerBounds && ZoomBounds == other.ZoomBounds &&
                PictureRotation == other.PictureRotation && MarkerRotation == other.MarkerRotation &&
                ZoomRotation == other.ZoomRotation;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    /// <summary>图片区域编辑器。视图缩放和平移与保存的归一化矩形坐标相互独立。</summary>
    internal sealed class ZoomRegionCanvas : Control
    {
        private enum DragMode { None, Draw, Move, Resize, Pan }

        private readonly Image preview;
        private readonly float pictureWidthPoints;
        private float zoomFactor = 1;
        private PointF panOffset;
        private DragMode dragMode;
        private int resizeHandle;
        private PointF dragStartImage;
        private Point dragStartScreen;
        private PointF dragStartPan;
        private RectangleF dragStartRegion;
        private float dragStartRotation;
        private bool dragStartEdited;
        private MouseButtons dragButton;
        private bool panToolActive;
        private bool spaceHeld;

        public event EventHandler SelectionChanged;
        public event EventHandler ViewChanged;
        public event Action<ZoomImageEntry> EntryPicked;
        public event EventHandler DeleteEntryRequested;
        public IList<ZoomImageEntry> Entries { get; set; }
        public ZoomImageEntry ActiveEntry { get; set; }
        public bool CanDrawRegion { get; set; } = true;
        public bool KeepSquare { get; set; } = true;
        public bool IsEditing => dragMode != DragMode.None && dragMode != DragMode.Pan;
        public RectangleF SelectedRegion { get; private set; }
        public float RegionRotationDegrees { get; private set; }
        public bool RegionWasEdited { get; private set; }
        public Office.MsoLineDashStyle? NativeOutlineDashStyle { get; set; }
        public Color OutlineColor { get; set; } = Color.Red;
        public float OutlineWidthPoints { get; set; } = 1.5f;
        public DashStyle OutlineDashStyle { get; set; } = DashStyle.Dash;
        public bool HasRectangle => SelectedRegion.Width > 0 && SelectedRegion.Height > 0;
        public bool HasSelection => HasRectangle &&
            (dragMode == DragMode.None || dragMode == DragMode.Pan);
        public float ZoomPercent => zoomFactor * 100;
        public bool PanToolActive
        {
            get => panToolActive;
            set
            {
                panToolActive = value;
                if (dragMode == DragMode.None) UpdateCursor(PointToClient(MousePosition));
            }
        }

        public ZoomRegionCanvas(Image preview, float pictureWidthPoints)
        {
            this.preview = preview ?? throw new ArgumentNullException(nameof(preview));
            this.pictureWidthPoints = Math.Max(0.01f, pictureWidthPoints);
            DoubleBuffered = true;
            ResizeRedraw = true;
            SetStyle(ControlStyles.Selectable, true);
            TabStop = true;
            BackColor = Color.FromArgb(235, 237, 240);
            Cursor = Cursors.Cross;
        }

        public void SetInitialRegion(RectangleF region, float rotationDegrees, bool edited = false)
        {
            SelectedRegion = region;
            RegionRotationDegrees = rotationDegrees;
            RegionWasEdited = edited;
            NotifySelectionChanged();
        }

        public void ClearSelection()
        {
            dragMode = DragMode.None;
            Capture = false;
            SelectedRegion = RectangleF.Empty;
            RegionRotationDegrees = 0;
            RegionWasEdited = true;
            UpdateCursor(PointToClient(MousePosition));
            NotifySelectionChanged();
        }

        public void FitImage()
        {
            if (dragMode != DragMode.None) return;
            zoomFactor = 1;
            panOffset = PointF.Empty;
            NotifyViewChanged();
        }

        private float FitScale => Math.Min(Math.Max(1, ClientSize.Width - 24) / (float)preview.Width,
            Math.Max(1, ClientSize.Height - 24) / (float)preview.Height);

        private RectangleF ImageBounds
        {
            get
            {
                float width = preview.Width * FitScale * zoomFactor;
                float height = preview.Height * FitScale * zoomFactor;
                return new RectangleF((ClientSize.Width - width) / 2f + panOffset.X,
                    (ClientSize.Height - height) / 2f + panOffset.Y, width, height);
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            RectangleF image = ImageBounds;
            e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            e.Graphics.DrawImage(preview, image);
            if (Entries != null)
            {
                foreach (ZoomImageEntry entry in Entries)
                {
                    if (entry == ActiveEntry || !entry.HasRegion) continue;
                    RectangleF region = entry.Options.Region;
                    var screen = new RectangleF(image.Left + region.Left * image.Width,
                        image.Top + region.Top * image.Height, region.Width * image.Width, region.Height * image.Height);
                    using (var pen = new Pen(entry.Options.OutlineColor,
                        Math.Max(0.75f, entry.Options.OutlineWidthPoints * image.Width / pictureWidthPoints))
                        { DashStyle = entry.Options.OutlineDashStyle })
                    {
                        e.Graphics.DrawPolygon(pen, ZoomImageGeometry.GetCorners(screen, entry.Options.RegionRotationDegrees));
                    }
                    DrawEntryLabel(e.Graphics, entry.DisplayName, screen.Location);
                }
            }
            if (!HasRectangle) return;

            var regionScreen = new RectangleF(image.Left + SelectedRegion.Left * image.Width,
                image.Top + SelectedRegion.Top * image.Height,
                SelectedRegion.Width * image.Width, SelectedRegion.Height * image.Height);
            float width = Math.Max(0.75f, OutlineWidthPoints * image.Width / pictureWidthPoints);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(OutlineColor, width) { DashStyle = OutlineDashStyle })
            {
                e.Graphics.DrawPolygon(pen, ZoomImageGeometry.GetCorners(regionScreen, RegionRotationDegrees));
            }
            if (ActiveEntry != null) DrawEntryLabel(e.Graphics, ActiveEntry.DisplayName, regionScreen.Location);
            if (dragMode != DragMode.Draw)
            {
                using (var handlePen = new Pen(OutlineColor))
                {
                    foreach (PointF point in GetHandlePoints())
                    {
                        var handle = new RectangleF(point.X - 3.5f, point.Y - 3.5f, 7, 7);
                        e.Graphics.FillRectangle(Brushes.White, handle);
                        e.Graphics.DrawRectangle(handlePen, handle.X, handle.Y, handle.Width, handle.Height);
                    }
                }
            }
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            // 鼠标回到画布后滚轮操作画布，而不是右侧设置控件。
            Focus();
            UpdateCursor(PointToClient(MousePosition));
        }

        protected override bool IsInputKey(Keys keyData) =>
            (keyData & Keys.KeyCode) == Keys.Space || base.IsInputKey(keyData);

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space && !e.Control && !e.Alt)
            {
                spaceHeld = true;
                if (dragMode == DragMode.None) UpdateCursor(PointToClient(MousePosition));
                e.Handled = e.SuppressKeyPress = true;
            }
            base.OnKeyDown(e);
        }

        protected override void OnKeyUp(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space)
            {
                spaceHeld = false;
                if (dragMode == DragMode.None) UpdateCursor(PointToClient(MousePosition));
                e.Handled = true;
            }
            base.OnKeyUp(e);
        }

        protected override void OnLostFocus(EventArgs e)
        {
            spaceHeld = false;
            if (dragMode == DragMode.Pan) FinishDrag();
            UpdateCursor(PointToClient(MousePosition));
            base.OnLostFocus(e);
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if (dragMode != DragMode.None || e.Delta == 0) return;
            PointF anchor = ToImagePoint(e.Location);
            zoomFactor = (float)Math.Max(0.1, Math.Min(20,
                zoomFactor * Math.Pow(1.1, e.Delta / 120.0)));
            PointF shiftedAnchor = ToScreenPoint(anchor);
            panOffset = new PointF(panOffset.X + e.X - shiftedAnchor.X,
                panOffset.Y + e.Y - shiftedAnchor.Y);
            ClampPan();
            NotifyViewChanged();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (dragMode != DragMode.None) return;
            Focus();
            bool startPan = e.Button == MouseButtons.Middle ||
                (e.Button == MouseButtons.Left && (PanToolActive || spaceHeld));
            if (!startPan && e.Button == MouseButtons.Left && HitHandle(e.Location) < 0 &&
                (!CanDrawRegion || HasRectangle))
            {
                ZoomImageEntry picked = HitEntry(ToImagePoint(e.Location));
                if (picked != null && picked != ActiveEntry) EntryPicked?.Invoke(picked);
            }
            dragStartScreen = e.Location;
            dragStartPan = panOffset;
            dragStartImage = ToImagePoint(e.Location);
            dragStartRegion = SelectedRegion;
            dragStartRotation = RegionRotationDegrees;
            dragStartEdited = RegionWasEdited;

            if (startPan)
            {
                dragMode = DragMode.Pan;
                Cursor = Cursors.Hand;
            }
            else if (e.Button == MouseButtons.Left)
            {
                if (HasRectangle)
                {
                    resizeHandle = HitHandle(e.Location);
                    if (resizeHandle >= 0) dragMode = DragMode.Resize;
                    else if (ContainsSelection(dragStartImage)) dragMode = DragMode.Move;
                    // 已有矩形时，点击外部不重画；必须先删除矩形。
                    else return;
                }
                else
                {
                    if (!CanDrawRegion) return;
                    if (!ImageBounds.Contains(e.Location)) return;
                    dragMode = DragMode.Draw;
                    RegionRotationDegrees = 0;
                    RegionWasEdited = true;
                }
            }
            else return;

            dragButton = e.Button;
            Capture = true;
            NotifySelectionChanged();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (dragMode != DragMode.None) UpdateDrag(e.Location);
            else UpdateCursor(e.Location);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (dragMode != DragMode.None && e.Button == dragButton)
            {
                UpdateDrag(e.Location);
                FinishDrag();
                UpdateCursor(e.Location);
            }
        }

        protected override void OnMouseCaptureChanged(EventArgs e)
        {
            base.OnMouseCaptureChanged(e);
            if (!Capture && dragMode != DragMode.None) FinishDrag();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape && dragMode != DragMode.None)
            {
                SelectedRegion = dragStartRegion;
                RegionRotationDegrees = dragStartRotation;
                RegionWasEdited = dragStartEdited;
                panOffset = dragStartPan;
                dragMode = DragMode.None;
                Capture = false;
                UpdateCursor(PointToClient(MousePosition));
                NotifySelectionChanged();
                NotifyViewChanged();
                return true;
            }
            if (keyData == Keys.Delete && ActiveEntry != null)
            {
                if (dragMode != DragMode.None) return true;
                if (DeleteEntryRequested != null) DeleteEntryRequested.Invoke(this, EventArgs.Empty);
                else ClearSelection();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (preview == null) return;
            ClampPan();
            NotifyViewChanged();
        }

        private void UpdateDrag(Point location)
        {
            if (dragMode == DragMode.Pan)
            {
                panOffset = new PointF(dragStartPan.X + location.X - dragStartScreen.X,
                    dragStartPan.Y + location.Y - dragStartScreen.Y);
                ClampPan();
                NotifyViewChanged();
                return;
            }
            PointF end = ToImagePoint(location);
            RectangleF region;
            if (dragMode == DragMode.Draw)
            {
                end = new PointF(Math.Max(0, Math.Min(preview.Width, end.X)),
                    Math.Max(0, Math.Min(preview.Height, end.Y)));
                if (KeepSquare)
                {
                    float dx = end.X - dragStartImage.X, dy = end.Y - dragStartImage.Y;
                    float availableWidth = dx < 0 ? dragStartImage.X : preview.Width - dragStartImage.X;
                    float availableHeight = dy < 0 ? dragStartImage.Y : preview.Height - dragStartImage.Y;
                    float side = Math.Min(Math.Max(Math.Abs(dx), Math.Abs(dy)),
                        Math.Min(availableWidth, availableHeight));
                    end = new PointF(dragStartImage.X + (dx < 0 ? -side : side),
                        dragStartImage.Y + (dy < 0 ? -side : side));
                }
                region = RectangleF.FromLTRB(Math.Min(dragStartImage.X, end.X), Math.Min(dragStartImage.Y, end.Y),
                    Math.Max(dragStartImage.X, end.X), Math.Max(dragStartImage.Y, end.Y));
            }
            else
            {
                var start = ToPixelRegion(dragStartRegion);
                var delta = new PointF(end.X - dragStartImage.X, end.Y - dragStartImage.Y);
                region = dragMode == DragMode.Move ? MoveRegion(start, delta) : ResizeRegion(start, delta);
            }
            var normalized = new RectangleF(region.X / preview.Width, region.Y / preview.Height,
                region.Width / preview.Width, region.Height / preview.Height);
            // 单纯点选或拖回原位时，保留手动矩形的精确原始坐标。
            bool changed = Math.Abs(normalized.X - dragStartRegion.X) > 0.0000001f ||
                Math.Abs(normalized.Y - dragStartRegion.Y) > 0.0000001f ||
                Math.Abs(normalized.Width - dragStartRegion.Width) > 0.0000001f ||
                Math.Abs(normalized.Height - dragStartRegion.Height) > 0.0000001f;
            SelectedRegion = changed ? normalized : dragStartRegion;
            RegionWasEdited = dragStartEdited || changed || dragMode == DragMode.Draw;
            NotifySelectionChanged();
        }

        private RectangleF MoveRegion(RectangleF start, PointF delta)
        {
            RectangleF bounds = GetRotatedBounds(start);
            // 继承的矩形可能部分超出图片，允许保留原位，避免点选时突然跳动。
            float minX = Math.Min(0, Math.Min(-bounds.Left, preview.Width - bounds.Right));
            float maxX = Math.Max(0, Math.Max(-bounds.Left, preview.Width - bounds.Right));
            float minY = Math.Min(0, Math.Min(-bounds.Top, preview.Height - bounds.Bottom));
            float maxY = Math.Max(0, Math.Max(-bounds.Top, preview.Height - bounds.Bottom));
            delta.X = Math.Max(minX, Math.Min(maxX, delta.X));
            delta.Y = Math.Max(minY, Math.Min(maxY, delta.Y));
            return new RectangleF(start.X + delta.X, start.Y + delta.Y, start.Width, start.Height);
        }

        private RectangleF ResizeRegion(RectangleF start, PointF delta)
        {
            PointF localDelta = ZoomImageGeometry.RotatePoint(delta, PointF.Empty, -RegionRotationDegrees);
            RectangleF candidate = ResizeLocalRegion(start, localDelta);
            if (!IsInsideImage(start) || IsInsideImage(candidate)) return candidate;

            if (KeepSquare)
            {
                // 按边长约束边界，宽高同时停止增长，避免贴边时破坏正方形。
                float minimum = Math.Min(2, Math.Min(start.Width, start.Height));
                float maximum = candidate.Width;
                for (int i = 0; i < 20; i++)
                {
                    float side = (minimum + maximum) / 2f;
                    if (IsInsideImage(GetSquareResizeRegion(start, side))) minimum = side;
                    else maximum = side;
                }
                return GetSquareResizeRegion(start, minimum);
            }

            // 保持对边或对角固定，缩放到图片边缘即停止；旋转矩形也遵循相同规则。
            float first = 0, last = 1;
            for (int i = 0; i < 20; i++)
            {
                float fraction = (first + last) / 2f;
                candidate = ResizeLocalRegion(start, new PointF(localDelta.X * fraction, localDelta.Y * fraction));
                if (IsInsideImage(candidate)) first = fraction;
                else last = fraction;
            }
            return ResizeLocalRegion(start, new PointF(localDelta.X * first, localDelta.Y * first));
        }

        private RectangleF ResizeLocalRegion(RectangleF start, PointF delta)
        {
            float left = start.Left, right = start.Right, top = start.Top, bottom = start.Bottom;
            float minimumWidth = Math.Min(2, start.Width);
            float minimumHeight = Math.Min(2, start.Height);
            if (resizeHandle == 0 || resizeHandle == 6 || resizeHandle == 7)
                left = Math.Min(right - minimumWidth, left + delta.X);
            if (resizeHandle == 2 || resizeHandle == 3 || resizeHandle == 4)
                right = Math.Max(left + minimumWidth, right + delta.X);
            if (resizeHandle == 0 || resizeHandle == 1 || resizeHandle == 2)
                top = Math.Min(bottom - minimumHeight, top + delta.Y);
            if (resizeHandle == 4 || resizeHandle == 5 || resizeHandle == 6)
                bottom = Math.Max(top + minimumHeight, bottom + delta.Y);
            if (KeepSquare)
            {
                float side = resizeHandle == 1 || resizeHandle == 5 ? bottom - top
                    : resizeHandle == 3 || resizeHandle == 7 ? right - left
                    : Math.Max(right - left, bottom - top);
                return GetSquareResizeRegion(start, side);
            }
            var center = new PointF(start.Left + start.Width / 2f, start.Top + start.Height / 2f);
            PointF newCenter = ZoomImageGeometry.RotatePoint(new PointF((left + right) / 2f, (top + bottom) / 2f),
                center, RegionRotationDegrees);
            return new RectangleF(newCenter.X - (right - left) / 2f, newCenter.Y - (bottom - top) / 2f,
                right - left, bottom - top);
        }

        private RectangleF GetSquareResizeRegion(RectangleF start, float side)
        {
            // 角点拖动固定对角，边中点拖动固定对边中点。
            float left = (start.Left + start.Right - side) / 2f;
            float top = (start.Top + start.Bottom - side) / 2f;
            if (resizeHandle == 0 || resizeHandle == 6 || resizeHandle == 7) left = start.Right - side;
            if (resizeHandle == 2 || resizeHandle == 3 || resizeHandle == 4) left = start.Left;
            if (resizeHandle == 0 || resizeHandle == 1 || resizeHandle == 2) top = start.Bottom - side;
            if (resizeHandle == 4 || resizeHandle == 5 || resizeHandle == 6) top = start.Top;
            var center = new PointF(start.Left + start.Width / 2f, start.Top + start.Height / 2f);
            PointF newCenter = ZoomImageGeometry.RotatePoint(new PointF(left + side / 2f, top + side / 2f),
                center, RegionRotationDegrees);
            return new RectangleF(newCenter.X - side / 2f, newCenter.Y - side / 2f, side, side);
        }

        private RectangleF GetRotatedBounds(RectangleF region)
        {
            PointF[] corners = ZoomImageGeometry.GetCorners(region, RegionRotationDegrees);
            float left = corners[0].X, right = left, top = corners[0].Y, bottom = top;
            foreach (PointF point in corners)
            {
                left = Math.Min(left, point.X); right = Math.Max(right, point.X);
                top = Math.Min(top, point.Y); bottom = Math.Max(bottom, point.Y);
            }
            return RectangleF.FromLTRB(left, top, right, bottom);
        }

        private bool IsInsideImage(RectangleF region)
        {
            RectangleF bounds = GetRotatedBounds(region);
            return bounds.Left >= -0.001f && bounds.Top >= -0.001f &&
                bounds.Right <= preview.Width + 0.001f && bounds.Bottom <= preview.Height + 0.001f;
        }

        private void FinishDrag()
        {
            if (dragMode == DragMode.Draw &&
                (SelectedRegion.Width * ImageBounds.Width < 4 || SelectedRegion.Height * ImageBounds.Height < 4))
            {
                SelectedRegion = RectangleF.Empty;
            }
            dragMode = DragMode.None;
            Capture = false;
            UpdateCursor(PointToClient(MousePosition));
            NotifySelectionChanged();
        }

        private PointF[] GetHandlePoints()
        {
            PointF[] corners = ZoomImageGeometry.GetCorners(ToPixelRegion(SelectedRegion), RegionRotationDegrees);
            var handles = new PointF[8];
            for (int i = 0; i < 4; i++)
            {
                PointF next = corners[(i + 1) % 4];
                handles[i * 2] = ToScreenPoint(corners[i]);
                handles[i * 2 + 1] = ToScreenPoint(new PointF((corners[i].X + next.X) / 2f,
                    (corners[i].Y + next.Y) / 2f));
            }
            return handles;
        }

        private int HitHandle(Point location)
        {
            if (!HasRectangle) return -1;
            PointF[] handles = GetHandlePoints();
            for (int i = 0; i < handles.Length; i++)
            {
                if (Math.Abs(location.X - handles[i].X) <= 7 && Math.Abs(location.Y - handles[i].Y) <= 7)
                    return i;
            }
            return -1;
        }

        private bool ContainsSelection(PointF imagePoint)
        {
            RectangleF region = ToPixelRegion(SelectedRegion);
            var center = new PointF(region.Left + region.Width / 2f, region.Top + region.Height / 2f);
            return region.Contains(ZoomImageGeometry.RotatePoint(imagePoint, center, -RegionRotationDegrees));
        }

        private ZoomImageEntry HitEntry(PointF imagePoint)
        {
            if (Entries == null) return null;
            // 当前矩形优先，重叠区域也可用窗口中的下拉框切换。
            if (HasRectangle && ContainsSelection(imagePoint)) return ActiveEntry;
            for (int i = Entries.Count - 1; i >= 0; i--)
            {
                ZoomImageEntry entry = Entries[i];
                if (!entry.HasRegion) continue;
                RectangleF region = ToPixelRegion(entry.Options.Region);
                var center = new PointF(region.Left + region.Width / 2f, region.Top + region.Height / 2f);
                if (region.Contains(ZoomImageGeometry.RotatePoint(imagePoint, center, -entry.Options.RegionRotationDegrees)))
                    return entry;
            }
            return null;
        }

        private void DrawEntryLabel(Graphics graphics, string text, PointF location)
        {
            SizeF size = graphics.MeasureString(text, Font);
            var label = new RectangleF(location.X, location.Y - size.Height - 2, size.Width + 4, size.Height + 2);
            graphics.FillRectangle(Brushes.White, label);
            graphics.DrawString(text, Font, Brushes.Black, label.Left + 2, label.Top + 1);
        }

        private void UpdateCursor(Point location)
        {
            if (dragMode == DragMode.Pan || PanToolActive || spaceHeld)
            {
                Cursor = Cursors.Hand;
                return;
            }
            int handle = HitHandle(location);
            if (handle >= 0)
            {
                // 将控制点的方向一同旋转，鼠标图标与矩形的实际边角对应。
                int direction = ((int)Math.Round(handle + RegionRotationDegrees / 45f) % 8 + 8) % 8;
                Cursor = direction % 4 == 0 ? Cursors.SizeNWSE : direction % 4 == 1 ? Cursors.SizeNS :
                    direction % 4 == 2 ? Cursors.SizeNESW : Cursors.SizeWE;
            }
            else Cursor = HasRectangle
                ? (ContainsSelection(ToImagePoint(location)) ? Cursors.SizeAll : Cursors.Default)
                : (CanDrawRegion ? Cursors.Cross : Cursors.Default);
        }

        private RectangleF ToPixelRegion(RectangleF region) => new RectangleF(region.X * preview.Width,
            region.Y * preview.Height, region.Width * preview.Width, region.Height * preview.Height);

        private PointF ToImagePoint(Point location)
        {
            RectangleF image = ImageBounds;
            return new PointF((location.X - image.Left) * preview.Width / image.Width,
                (location.Y - image.Top) * preview.Height / image.Height);
        }

        private PointF ToScreenPoint(PointF point)
        {
            RectangleF image = ImageBounds;
            return new PointF(image.Left + point.X * image.Width / preview.Width,
                image.Top + point.Y * image.Height / preview.Height);
        }

        private void ClampPan()
        {
            // 适应窗口时也可拖动，但至少保留一小段图片可见，避免移出画布后找不到。
            float maxX = Math.Max(0, (preview.Width * FitScale * zoomFactor + ClientSize.Width) / 2f - 24);
            float maxY = Math.Max(0, (preview.Height * FitScale * zoomFactor + ClientSize.Height) / 2f - 24);
            panOffset = new PointF(Math.Max(-maxX, Math.Min(maxX, panOffset.X)),
                Math.Max(-maxY, Math.Min(maxY, panOffset.Y)));
        }

        private void NotifySelectionChanged()
        {
            Invalidate();
            SelectionChanged?.Invoke(this, EventArgs.Empty);
        }

        private void NotifyViewChanged()
        {
            Invalidate();
            ViewChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        private void CreateZoomImage()
        {
            try
            {
                if (app.ActiveWindow == null) { MessageBox.Show("请先打开一个演示文稿。", "提示"); return; }
                Selection selection = app.ActiveWindow.Selection;
                if (selection.Type != PpSelectionType.ppSelectionShapes || selection.HasChildShapeRange ||
                    (selection.ShapeRange.Count != 1 && selection.ShapeRange.Count != 2))
                {
                    MessageBox.Show("请选择一张图片，或一张图片和一个普通矩形（不支持组合内的对象）。", "提示");
                    return;
                }
                Shape picture = null, rectangle = null;
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.Type == Office.MsoShapeType.msoPicture || shape.Type == Office.MsoShapeType.msoLinkedPicture)
                        picture = shape;
                    else if (shape.Type == Office.MsoShapeType.msoAutoShape &&
                        shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle) rectangle = shape;
                }
                if (picture == null || (selection.ShapeRange.Count == 2 && rectangle == null))
                {
                    MessageBox.Show("请选择一张图片，或一张图片和一个普通矩形。", "提示");
                    return;
                }
                Globals.ThisAddIn.ZoomGuideLines?.RefreshNow();
                using (Globals.ThisAddIn.ZoomGuideLines?.Suspend())
                {
                    Slide slide = app.ActiveWindow.View.Slide;
                    IList<ZoomImageExistingObjects> records = ZoomGuideLineTracker.ReadExisting(slide, picture);
                    int selectedIndex = 0;
                    if (rectangle != null)
                    {
                        var found = records.FirstOrDefault(record => record.Marker.Id == rectangle.Id);
                        if (found != null) selectedIndex = records.IndexOf(found);
                        else
                        {
                            string key = "MANUAL_" + Guid.NewGuid().ToString("N");
                            int number = Math.Max(records.Count + 1, records.Select(record => record.Entry.DisplayOrder)
                                .Where(order => order < int.MaxValue).DefaultIfEmpty(0).Max() + 1);
                            var entry = new ZoomImageEntry(key, $"放大图 {number}",
                                GetManualZoomRectangleSelection(picture, rectangle));
                            records.Add(new ZoomImageExistingObjects(key, rectangle, null, new Shape[0], entry, true));
                            selectedIndex = records.Count - 1;
                        }
                    }
                    if (!TryEditZoomImages(picture, rectangle, records.Select(record => record.Entry).ToArray(),
                        selectedIndex, out IList<ZoomImageEntry> entries)) return;
                    ApplyZoomImageEdits(slide, picture, records, entries);
                }
            }
            catch (Exception ex) { MessageBox.Show($"制作放大图时出错: {ex.Message}", "操作失败"); }
        }

        private bool TryEditZoomImages(Shape picture, Shape rectangle, IList<ZoomImageEntry> entries,
            int selectedIndex, out IList<ZoomImageEntry> result)
        {
            result = null;
            string previewPath = Path.Combine(Path.GetTempPath(), "SlideSCI-zoom-" + Guid.NewGuid() + ".png");
            Shape previewCopy = null;
            try
            {
                previewCopy = picture.Duplicate()[1];
                ZoomGuideLineTracker.RemoveCopiedLinks(previewCopy);
                previewCopy.Rotation = 0;
                previewCopy.Line.Visible = Office.MsoTriState.msoFalse;
                previewCopy.Shadow.Visible = Office.MsoTriState.msoFalse;
                previewCopy.Glow.Radius = 0;
                previewCopy.SoftEdge.Radius = 0;
                previewCopy.Reflection.Type = Office.MsoReflectionType.msoReflectionTypeNone;
                var page = app.ActivePresentation.PageSetup;
                previewCopy.Export(previewPath, PpShapeFormat.ppShapeFormatPNG,
                    (int)Math.Ceiling(page.SlideWidth * 2), (int)Math.Ceiling(page.SlideHeight * 2), PpExportMode.ppRelativeToSlide);
                previewCopy.Delete();
                previewCopy = null;
                picture.Select(Office.MsoTriState.msoTrue);
                using (var preview = Image.FromFile(previewPath))
                using (var dialog = new ZoomImageForm(preview, picture.Width, entries, selectedIndex))
                {
                    if (dialog.ShowDialog(new PowerPointDialogOwner(app.HWND)) != DialogResult.OK) return false;
                    result = dialog.SelectedEntries;
                    return true;
                }
            }
            finally
            {
                try
                {
                    previewCopy?.Delete();
                    picture.Select(Office.MsoTriState.msoTrue);
                    rectangle?.Select(Office.MsoTriState.msoFalse);
                }
                catch (COMException ex) { Debug.WriteLine($"清理放大图预览时出错: {ex.Message}"); }
                try { if (File.Exists(previewPath)) File.Delete(previewPath); }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                { Debug.WriteLine($"清理放大图预览文件时出错: {ex.Message}"); }
            }
        }

        /// <summary>先生成所有需要更新的对象，全部成功后再提交关联并清理旧图。</summary>
        private void ApplyZoomImageEdits(Slide slide, Shape picture, IList<ZoomImageExistingObjects> originals,
            IList<ZoomImageEntry> entries)
        {
            var byKey = originals.ToDictionary(record => record.Key);
            var temporary = new List<Shape>();
            var drafts = new List<ZoomImageDraft>();
            var restores = new List<Action>();
            bool committed = false;
            try
            {
                app.StartNewUndoEntry();
                foreach (ZoomImageEntry entry in entries)
                {
                    if (!entry.HasRegion) throw new InvalidOperationException("请为每个放大图框选有效区域。");
                    ZoomImageExistingObjects previous = null;
                    if (entry.RecordKey != null) byKey.TryGetValue(entry.RecordKey, out previous);
                    if (previous?.Zoom != null && !previous.NeedsImageRefresh &&
                        ZoomImageEntry.SameOptions(entry.Options, entry.InitialOptions) &&
                        previous.Marker.Fill.Visible == Office.MsoTriState.msoFalse &&
                        previous.Marker.Line.Visible == Office.MsoTriState.msoTrue &&
                        previous.Guides.Count == (entry.Options.AddGuideLines ? 2 : 0)) continue;
                    drafts.Add(CreateZoomImageDraft(slide, picture, entry, previous, temporary));
                }
                var source = new RectangleF(picture.Left, picture.Top, picture.Width, picture.Height);
                var draftMap = drafts.ToDictionary(draft => draft.Entry);
                var positions = ZoomImageLayout.Calculate(source, entries, ZoomImageGeometry.GapPoints, entry =>
                {
                    if (draftMap.TryGetValue(entry, out ZoomImageDraft draft)) return new SizeF(draft.Zoom.Width, draft.Zoom.Height);
                    return new SizeF(entry.Options.Region.Width * picture.Width, entry.Options.Region.Height * picture.Height);
                });
                foreach (ZoomImageDraft draft in drafts)
                {
                    if (!positions.TryGetValue(draft.Entry, out RectangleF bounds))
                        throw new InvalidOperationException($"{draft.Entry.DisplayName}的矩形与原图没有重叠。");
                    SetZoomShapeBounds(draft.Zoom, bounds, draft.Entry.PreservesLayout
                        ? draft.Entry.OriginalZoomRotation : draft.Zoom.Rotation);
                    draft.Zoom.LockAspectRatio = Office.MsoTriState.msoTrue;
                    if (draft.Entry.Options.AddGuideLines)
                        draft.Guides.AddRange(AddZoomGuideLines(slide, picture, draft.Marker, draft.Zoom,
                            ZoomImageGeometry.GetRelativePlacement(source, bounds, draft.Entry.Options.Placement),
                            draft.Entry.Options.GuideLineExtent, temporary));
                }

                var keptMarkers = new HashSet<int>();
                foreach (ZoomImageEntry entry in entries.Where(entry => !draftMap.ContainsKey(entry)))
                    if (entry.RecordKey != null && byKey.TryGetValue(entry.RecordKey, out ZoomImageExistingObjects old))
                        keptMarkers.Add(old.Marker.Id);
                foreach (ZoomImageDraft draft in drafts)
                {
                    draft.TargetMarker = draft.Marker;
                    if (draft.Previous != null && !keptMarkers.Contains(draft.Previous.Marker.Id))
                    {
                        draft.TargetMarker = draft.Previous.Marker;
                        restores.Add(CaptureZoomRectangleRestoreAction(draft.TargetMarker));
                        CopyZoomRectangleBounds(draft.Marker, draft.TargetMarker);
                        ApplyZoomRectangleStyle(draft.TargetMarker, draft.Entry.Options);
                    }
                    keptMarkers.Add(draft.TargetMarker.Id);
                    draft.LinkKey = ZoomGuideLineTracker.Attach(picture, draft.TargetMarker, draft.Zoom, draft.Guides,
                        draft.Entry.Options.GuideLineExtent, draft.Entry.Options.Placement,
                        draft.Entry.Options, draft.Entry.DisplayName);
                }
                // 到此所有新对象和关联均已完成，后续只做清理，失败时保留完成的新图。
                committed = true;
                var errors = new List<string>();
                var unchangedKeys = new HashSet<string>(entries.Where(entry => !draftMap.ContainsKey(entry))
                    .Select(entry => entry.RecordKey).Where(key => key != null));
                var removedIds = new HashSet<int>();
                var outdated = originals.Where(old => !unchangedKeys.Contains(old.Key)).ToArray();
                // 共享矩形的旧关联先统一移除，再删除对象，避免第二组访问已删除的矩形。
                foreach (ZoomImageExistingObjects old in outdated)
                {
                    if (!old.IsManualRectangle)
                    {
                        try { ZoomGuideLineTracker.RemoveLink(old.Key, picture, old.Marker, old.Zoom, old.Guides); }
                        catch (COMException ex) { errors.Add(ex.Message); }
                    }
                }
                foreach (ZoomImageExistingObjects old in outdated)
                {
                    DeleteZoomObject(old.Zoom, removedIds, errors, old.ZoomId);
                    for (int i = 0; i < old.Guides.Count; i++)
                        DeleteZoomObject(old.Guides[i], removedIds, errors, old.GuideIds[i]);
                    if (!keptMarkers.Contains(old.MarkerId)) DeleteZoomObject(old.Marker, removedIds, errors, old.MarkerId);
                }
                foreach (ZoomImageDraft draft in drafts)
                    if (draft.TargetMarker != draft.Marker) DeleteZoomObject(draft.Marker, removedIds, errors);
                temporary.Clear();
                Globals.ThisAddIn.ZoomGuideLines?.InvalidateCache();
                if (drafts.Count > 0) drafts[drafts.Count - 1].Zoom.Select(Office.MsoTriState.msoTrue);
                else picture.Select(Office.MsoTriState.msoTrue);
                if (errors.Count > 0) MessageBox.Show("放大图已更新，但部分旧对象未能清理：\n" +
                    string.Join("\n", errors.Distinct()), "提示");
            }
            catch
            {
                if (!committed)
                {
                    foreach (ZoomImageDraft draft in drafts.Where(draft => draft.LinkKey != null))
                    {
                        try { ZoomGuideLineTracker.RemoveLink(draft.LinkKey, picture, draft.TargetMarker, draft.Zoom, draft.Guides); }
                        catch (COMException ex) { Debug.WriteLine(ex.Message); }
                    }
                    foreach (Action restore in restores.AsEnumerable().Reverse())
                    {
                        try { restore(); }
                        catch (COMException ex) { Debug.WriteLine(ex.Message); }
                    }
                    foreach (Shape shape in temporary)
                    {
                        try { shape.Delete(); }
                        catch (COMException ex) { Debug.WriteLine(ex.Message); }
                    }
                }
                throw;
            }
        }

        private ZoomImageDraft CreateZoomImageDraft(Slide slide, Shape picture, ZoomImageEntry entry,
            ZoomImageExistingObjects previous, List<Shape> temporary)
        {
            Shape marker;
            if (previous != null)
            {
                marker = previous.Marker.Duplicate()[1];
                temporary.Add(marker);
                ZoomGuideLineTracker.RemoveCopiedLinks(marker);
                marker.Left = previous.Marker.Left;
                marker.Top = previous.Marker.Top;
                if (entry.Options.RegionWasEdited || entry.Options.Region != entry.InitialOptions.Region ||
                    entry.Options.RegionRotationDegrees != entry.InitialOptions.RegionRotationDegrees)
                    SetZoomShapeBounds(marker, ZoomImageGeometry.GetRegionBounds(
                        new RectangleF(picture.Left, picture.Top, picture.Width, picture.Height),
                        picture.Rotation, entry.Options.Region), picture.Rotation + entry.Options.RegionRotationDegrees);
            }
            else
            {
                marker = CreateZoomRegionRectangle(slide, picture, entry.Options.Region);
                temporary.Add(marker);
                marker.Rotation = picture.Rotation + entry.Options.RegionRotationDegrees;
            }
            ApplyZoomRectangleStyle(marker, entry.Options);
            Shape pictureCopy = picture.Duplicate()[1];
            temporary.Add(pictureCopy);
            ZoomGuideLineTracker.RemoveCopiedLinks(pictureCopy);
            pictureCopy.Left = picture.Left;
            pictureCopy.Top = picture.Top;
            Shape mask = marker.Duplicate()[1];
            temporary.Add(mask);
            mask.Left = marker.Left;
            mask.Top = marker.Top;
            var before = new HashSet<int>(slide.Shapes.Cast<Shape>().Select(shape => shape.Id));
            int pictureId = pictureCopy.Id, maskId = mask.Id;
            pictureCopy.Select(Office.MsoTriState.msoTrue);
            mask.Select(Office.MsoTriState.msoFalse);
            app.ActiveWindow.Selection.ShapeRange.MergeShapes(Office.MsoMergeCmd.msoMergeIntersect, pictureCopy);
            temporary.Remove(pictureCopy);
            temporary.Remove(mask);
            Shape zoom = slide.Shapes.Cast<Shape>().FirstOrDefault(shape => shape.Id == pictureId || shape.Id == maskId ||
                !before.Contains(shape.Id));
            if (zoom == null) throw new InvalidOperationException($"{entry.DisplayName}没有有效的图片相交区域。");
            temporary.Add(zoom);
            if (previous?.Zoom != null) zoom.Name = previous.Zoom.Name;
            if (entry.Options.UseRectangleColorForZoomImage)
            {
                zoom.Line.Visible = Office.MsoTriState.msoTrue;
                zoom.Line.ForeColor.RGB = marker.Line.ForeColor.RGB;
                zoom.Line.Transparency = 0;
                zoom.Line.Weight = marker.Line.Weight;
            }
            else if (previous?.Zoom != null && !entry.InitialOptions.UseRectangleColorForZoomImage)
            {
                // 更新取图区域时保留用户独立设置的放大图边框。
                zoom.Line.Visible = previous.Zoom.Line.Visible;
                zoom.Line.ForeColor.RGB = previous.Zoom.Line.ForeColor.RGB;
                zoom.Line.Weight = previous.Zoom.Line.Weight;
                zoom.Line.Transparency = previous.Zoom.Line.Transparency;
                zoom.Line.DashStyle = previous.Zoom.Line.DashStyle;
                zoom.Line.Style = previous.Zoom.Line.Style;
            }
            return new ZoomImageDraft(entry, previous, marker, zoom);
        }

        private static void SetZoomShapeBounds(Shape shape, RectangleF bounds, float rotation)
        {
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            shape.Width = bounds.Width;
            shape.Height = bounds.Height;
            shape.Rotation = (rotation % 360 + 360) % 360;
            shape.Left = bounds.Left;
            shape.Top = bounds.Top;
        }

        private static void DeleteZoomObject(Shape shape, HashSet<int> deletedIds, List<string> errors, int? knownId = null)
        {
            if (shape == null) return;
            try { if (deletedIds.Add(knownId ?? shape.Id)) shape.Delete(); }
            catch (COMException ex) { errors.Add(ex.Message); }
        }

        private sealed class ZoomImageDraft
        {
            public ZoomImageEntry Entry { get; }
            public ZoomImageExistingObjects Previous { get; }
            public Shape Marker { get; }
            public Shape Zoom { get; }
            public List<Shape> Guides { get; } = new List<Shape>();
            public Shape TargetMarker { get; set; }
            public string LinkKey { get; set; }
            public ZoomImageDraft(ZoomImageEntry entry, ZoomImageExistingObjects previous, Shape marker, Shape zoom)
            { Entry = entry; Previous = previous; Marker = marker; Zoom = zoom; }
        }
    }
}

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
                    selection.ShapeRange.Count == 0)
                {
                    MessageBox.Show("请选择一张或多张图片，或一张图片和一个普通矩形（不支持组合内的对象）。", "提示");
                    return;
                }
                // 导出预览和裁剪都会改变当前选择，必须先保存原始对象及其顺序。
                Shape[] selectedShapes = selection.ShapeRange.Cast<Shape>().ToArray();
                var pictures = new List<Shape>();
                Shape rectangle = null;
                bool hasUnsupportedShape = false;
                foreach (Shape shape in selectedShapes)
                {
                    if (ScaleBarService.IsImage(shape))
                        pictures.Add(shape);
                    else if (shape.Type == Office.MsoShapeType.msoAutoShape &&
                        shape.AutoShapeType == Office.MsoAutoShapeType.msoShapeRectangle && rectangle == null)
                        rectangle = shape;
                    else hasUnsupportedShape = true;
                }
                if (pictures.Count == 0 || hasUnsupportedShape || (rectangle != null && pictures.Count != 1))
                {
                    MessageBox.Show("请选择一张或多张图片，或一张图片和一个普通矩形。", "提示");
                    return;
                }
                Shape picture = pictures[0];
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
                    if (!TryEditZoomImages(picture, selectedShapes, records.Select(record => record.Entry).ToArray(),
                        selectedIndex, out IList<ZoomImageEntry> entries)) return;
                    ApplySelectedZoomImageEdits(slide, pictures, records, entries);
                }
            }
            catch (Exception ex) { MessageBox.Show($"制作放大图时出错: {ex.Message}", "操作失败"); }
        }

        private bool TryEditZoomImages(Shape picture, IList<Shape> selectedShapes, IList<ZoomImageEntry> entries,
            int selectedIndex, out IList<ZoomImageEntry> result)
        {
            result = null;
            string previewPath = Path.Combine(Path.GetTempPath(), "SlideSCI-zoom-" + Guid.NewGuid() + ".png");
            Shape previewCopy = null;
            try
            {
                previewCopy = ScaleBarService.DuplicateContent(picture);
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
                ImageFieldOfView sourceFov = ScaleBarService.ReadFov(picture);
                SizeF? visibleFov = sourceFov == null ? (SizeF?)null : ScaleBarService.VisibleFov(picture, sourceFov);
                ScaleBarSettings sourceScaleBar = ScaleBarService.ReadSettings(picture);
                using (var preview = Image.FromFile(previewPath))
                using (var dialog = new ZoomImageForm(preview, picture.Width, entries, selectedIndex,
                    sourceScaleBar?.LengthMicrometers, visibleFov, sourceScaleBar))
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
                    for (int i = 0; i < selectedShapes.Count; i++)
                        selectedShapes[i].Select(i == 0 ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
                }
                catch (COMException ex) { Debug.WriteLine($"清理放大图预览时出错: {ex.Message}"); }
                try { if (File.Exists(previewPath)) File.Delete(previewPath); }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                { Debug.WriteLine($"清理放大图预览文件时出错: {ex.Message}"); }
            }
        }

        /// <summary>第一张图编辑全部条目；本次新增区域按相对坐标添加到其他所选图片。</summary>
        private void ApplySelectedZoomImageEdits(Slide slide, IList<Shape> pictures,
            IList<ZoomImageExistingObjects> firstOriginals, IList<ZoomImageEntry> entries)
        {
            ZoomImageEntry[] addedEntries = entries.Where(entry => entry.RecordKey == null).ToArray();
            var errors = new List<string>();
            app.StartNewUndoEntry();
            for (int i = 0; i < pictures.Count; i++)
            {
                if (i > 0 && addedEntries.Length == 0) continue;
                Shape picture = pictures[i];
                try
                {
                    IList<ZoomImageExistingObjects> originals = i == 0 ? firstOriginals
                        : ZoomGuideLineTracker.ReadExisting(slide, picture);
                    IList<ZoomImageEntry> targetEntries = i == 0 ? entries
                        : AppendBatchZoomEntries(originals, addedEntries);
                    ApplyZoomImageEdits(slide, picture, originals, targetEntries);
                }
                catch (Exception ex)
                {
                    // 单张图失败时沿用原有回滚，继续处理剩余图片。
                    errors.Add($"第 {i + 1} 张图片：{ex.Message}");
                }
            }
            if (errors.Count > 0) MessageBox.Show("制作放大图时遇到以下问题：\n" +
                string.Join("\n", errors), "操作失败");
        }

        private static IList<ZoomImageEntry> AppendBatchZoomEntries(IList<ZoomImageExistingObjects> originals,
            IList<ZoomImageEntry> addedEntries)
        {
            // 保留目标图已有条目，并分配独立编号和关联，避免重复或误用第一张图的对象。
            var result = originals.Select(record => record.Entry).ToList();
            var names = new HashSet<string>(result.Select(entry => entry.DisplayName));
            int nextNumber = Math.Max(result.Count + 1, result.Where(entry => entry.DisplayOrder < int.MaxValue)
                .Select(entry => entry.DisplayOrder + 1).DefaultIfEmpty(1).Max());
            foreach (ZoomImageEntry entry in addedEntries)
            {
                string name = entry.DisplayName;
                if (names.Contains(name))
                {
                    while (names.Contains($"放大图 {nextNumber}")) nextNumber++;
                    name = $"放大图 {nextNumber++}";
                }
                names.Add(name);
                // Options 只含相对区域及样式；实际矩形和放大图由各自原图生成。
                result.Add(new ZoomImageEntry(null, name, entry.Options));
            }
            return result;
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
                foreach (ZoomImageEntry entry in entries)
                {
                    if (!entry.HasRegion) throw new InvalidOperationException("请为每个放大图框选有效区域。");
                    ZoomImageExistingObjects previous = null;
                    if (entry.RecordKey != null) byKey.TryGetValue(entry.RecordKey, out previous);
                    if (previous?.Zoom != null && !previous.NeedsImageRefresh &&
                        ScaleBarService.HasCurrentCalibration(picture, previous.Zoom) &&
                        ScaleBarService.IsZoomScaleBarWithinLimit(previous.Zoom) &&
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
                    ScaleBarSettings bar = (draft.Previous?.Zoom == null ? null : ScaleBarService.ReadSettings(draft.Previous.Zoom))
                        ?? ScaleBarService.ReadSettings(picture);
                    bar = bar?.Copy();
                    if (bar != null && draft.Entry.Options.ScaleBarLengthMicrometers.HasValue)
                    {
                        bar = ScaleBarService.FitZoomSettings(draft.Zoom, bar,
                            draft.Entry.Options.ScaleBarLengthMicrometers, draft.Entry.Options.ScaleBarShowText);
                        draft.Entry.Options = ZoomImageEntry.WithScaleBar(draft.Entry.Options, bar);
                        draft.Zoom = ScaleBarService.Add(slide, draft.Zoom, bar, temporary);
                    }
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
            Shape zoom = ZoomGuideLineTracker.CreateCrop(slide, picture, marker);
            temporary.Add(zoom);
            if (previous?.Zoom != null) zoom.Name = previous.Zoom.Name;
            if (entry.Options.UseRectangleColorForZoomImage)
            {
                zoom.Line.Visible = Office.MsoTriState.msoTrue;
                zoom.Line.Transparency = 0;
                zoom.Line.Weight = marker.Line.Weight;
                zoom.Line.ForeColor.RGB = marker.Line.ForeColor.RGB;
            }
            else if (previous?.Zoom != null && !entry.InitialOptions.UseRectangleColorForZoomImage)
            {
                // 更新取图区域时保留用户独立设置的放大图边框。
                ZoomGuideLineTracker.CopyZoomOutline(previous.Zoom, zoom);
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
            try
            {
                if (deletedIds.Add(knownId ?? shape.Id))
                    ScaleBarService.DeleteWithAnnotations((Slide)shape.Parent, shape);
            }
            catch (COMException ex) { errors.Add(ex.Message); }
        }

        private sealed class ZoomImageDraft
        {
            public ZoomImageEntry Entry { get; }
            public ZoomImageExistingObjects Previous { get; }
            public Shape Marker { get; }
            public Shape Zoom { get; set; }
            public List<Shape> Guides { get; } = new List<Shape>();
            public Shape TargetMarker { get; set; }
            public string LinkKey { get; set; }
            public ZoomImageDraft(ZoomImageEntry entry, ZoomImageExistingObjects previous, Shape marker, Shape zoom)
            { Entry = entry; Previous = previous; Marker = marker; Zoom = zoom; }
        }
    }
}

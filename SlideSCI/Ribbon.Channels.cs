using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        private const string PseudoColorTag = "SLIDESCI_PSEUDO_COLOR";
        private const int ChannelImageMaxPixels = 4096;

        private void PseudoColor_Click(object sender, RibbonControlEventArgs e)
        {
            var previews = new List<Bitmap>();
            PowerPoint.Shape[] pictures = null;
            string step = "检查所选图片";
            try
            {
                if (!TryGetChannelPictures(1, int.MaxValue, out PowerPoint.Slide slide,
                    out PowerPoint.ShapeRange originalRange, out pictures)) return;
                using (Globals.ThisAddIn.ZoomGuideLines?.Suspend())
                {
                    var colors = new Color[pictures.Length];
                    var sourcePeaks = new int[pictures.Length];
                    bool hasRememberedColor = int.TryParse(Properties.Settings.Default.PseudoColorArgb,
                        NumberStyles.Integer, CultureInfo.InvariantCulture, out int rememberedArgb);
                    for (int i = 0; i < pictures.Length; i++)
                    {
                        step = $"读取第 {i + 1} 张图片的预览";
                        previews.Add(ExportChannelPicture(pictures[i], GetChannelRasterSize(pictures[i], 640)));
                        step = $"读取第 {i + 1} 张图片的伪彩设置";
                        bool wasColored = int.TryParse(ReadPseudoColorTag(pictures[i]), NumberStyles.Integer,
                            CultureInfo.InvariantCulture, out int argb);
                        Color sourceColor = wasColored ? Color.FromArgb(argb) : Color.White;
                        Color defaultColor = wasColored ? sourceColor
                            : ChannelImageProcessor.DefaultColors[i % ChannelImageProcessor.DefaultColors.Length];
                        colors[i] = hasRememberedColor ? Color.FromArgb(rememberedArgb) : defaultColor;
                        // 默认颜色可来自上次操作，亮度恢复比例始终由当前图片的原 LUT 决定。
                        sourcePeaks[i] = Math.Max(sourceColor.R, Math.Max(sourceColor.G, sourceColor.B));
                    }
                    RestoreChannelSelection(pictures);
                    step = "打开伪彩设置窗口";
                    using (var dialog = new PseudoColorForm(pictures.Select(picture => picture.Name).ToArray(),
                        previews, colors, sourcePeaks))
                    {
                        if (dialog.ShowDialog(new PowerPointDialogOwner(app.HWND)) != DialogResult.OK) return;
                        step = "应用伪彩";
                        ApplyPseudoColors(slide, originalRange, pictures, dialog.SelectedColors, sourcePeaks, dialog.KeepOriginal);
                        pictures = null; // 已由提交操作选中新图，不能再访问删除的原对象。
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show($"设置伪彩时出错（{step}）：{ex.Message}", "伪彩", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                foreach (Bitmap preview in previews) preview.Dispose();
                if (pictures != null) RestoreChannelSelection(pictures);
            }
        }

        private static string ReadPseudoColorTag(PowerPoint.Shape picture)
        {
            // 普通图片尚无伪彩标签；只读取存在的标签，避免向 COM 索取不存在的项目。
            PowerPoint.Tags tags = picture.Tags;
            for (int i = 1; i <= tags.Count; i++)
                if (string.Equals(tags.Name(i), PseudoColorTag, StringComparison.OrdinalIgnoreCase))
                    return tags.Value(i);
            return null;
        }

        private bool TryGetChannelPictures(int minimum, int maximum, out PowerPoint.Slide slide,
            out PowerPoint.ShapeRange originalRange, out PowerPoint.Shape[] pictures)
        {
            slide = null;
            originalRange = null;
            pictures = null;
            if (app.Windows.Count == 0 || app.ActiveWindow == null)
            {
                MessageBox.Show("请先打开一个演示文稿，并选中图片。", "提示");
                return false;
            }
            PowerPoint.Selection selection = app.ActiveWindow.Selection;
            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.HasChildShapeRange)
            {
                MessageBox.Show("请选中幻灯片上的图片；组合中的图片请先取消组合。", "提示");
                return false;
            }
            originalRange = selection.ShapeRange;
            // 导出临时副本会改变选区，先固定所选图片及顺序。
            pictures = GetSelectedShapesInSelectionOrder(selection).ToArray();
            if (pictures.Length < minimum || pictures.Length > maximum || pictures.Any(picture =>
                picture.Type != Office.MsoShapeType.msoPicture && picture.Type != Office.MsoShapeType.msoLinkedPicture))
            {
                MessageBox.Show(maximum == int.MaxValue ? "请选择一张或多张图片。"
                    : $"合并通道需要选择 {minimum}–{maximum} 张图片，请不要同时选择文字、形状或组合。", "提示");
                pictures = null;
                return false;
            }
            slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
            if (slide != null) return true;
            MessageBox.Show("请切换到图片所在的幻灯片后再操作。", "提示");
            return false;
        }

        private void ApplyPseudoColors(PowerPoint.Slide slide, PowerPoint.ShapeRange originalRange,
            PowerPoint.Shape[] pictures, Color[] colors, int[] sourcePeaks, bool keepOriginal)
        {
            var replacements = new List<PowerPoint.Shape>();
            bool committed = false;
            string step = "开始生成伪彩图";
            float offset = keepOriginal
                ? pictures.Max(picture => picture.Left + picture.Width) - pictures.Min(picture => picture.Left) + 18F : 0;
            app.StartNewUndoEntry();
            try
            {
                for (int i = 0; i < pictures.Length; i++)
                {
                    step = $"读取第 {i + 1} 张图片的画面";
                    PowerPoint.Shape source = pictures[i];
                    using (Bitmap image = ExportChannelPicture(source, GetChannelRasterSize(source, ChannelImageMaxPixels)))
                    using (Bitmap colored = ChannelImageProcessor.ApplyColor(image, colors[i], sourcePeaks[i]))
                    {
                        step = $"插入第 {i + 1} 张伪彩图";
                        PowerPoint.Shape replacement = InsertChannelBitmap(slide, colored, source.Left + offset,
                            source.Top, source.Width, source.Height);
                        replacements.Add(replacement);
                        step = $"保留第 {i + 1} 张图片的布局";
                        replacement.Rotation = source.Rotation;
                        replacement.Name = keepOriginal ? source.Name + "_伪彩" : source.Name;
                        replacement.AlternativeText = source.AlternativeText;
                        replacement.Title = source.Title;
                        replacement.Visible = source.Visible;
                        replacement.LockAspectRatio = source.LockAspectRatio;
                        step = $"保留第 {i + 1} 张图片的描边";
                        // 无描边图片的线型、颜色等属性可能未定义，不能直接回写给新图片。
                        if (source.Line.Visible == Office.MsoTriState.msoFalse)
                            replacement.Line.Visible = Office.MsoTriState.msoFalse;
                        else ZoomGuideLineTracker.CopyZoomOutline(source, replacement);
                        step = $"保留第 {i + 1} 张图片的关联设置";
                        // 原位替换时保留关联标签；新增副本不能与原图共用放大图关联。
                        for (int tag = 1; tag <= source.Tags.Count; tag++)
                            replacement.Tags.Add(source.Tags.Name(tag), source.Tags.Value(tag));
                        if (keepOriginal) ZoomGuideLineTracker.RemoveCopiedLinks(replacement);
                        replacement.Tags.Add(PseudoColorTag, colors[i].ToArgb().ToString(CultureInfo.InvariantCulture));
                        if (!keepOriginal)
                        {
                            step = $"保留第 {i + 1} 张图片的层级";
                            int position = source.ZOrderPosition + 1;
                            while (replacement.ZOrderPosition > position)
                                replacement.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                        }
                    }
                }
                // 所有新图成功后一次删除原选区；生成失败时原图保持完整。
                step = "替换所选原图";
                if (!keepOriginal) originalRange.Delete();
                committed = true;
                Globals.ThisAddIn.ZoomGuideLines?.InvalidateCache();
                RestoreChannelSelection(replacements);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"{step}失败：{ex.Message}", ex);
            }
            finally
            {
                if (!committed)
                    foreach (PowerPoint.Shape replacement in replacements) TryDeleteChannelShape(replacement);
            }
        }

        private void MergeChannels_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Shape[] pictures = null;
            bool committed = false;
            try
            {
                if (!TryGetChannelPictures(2, 7, out PowerPoint.Slide slide, out _, out pictures)) return;
                PowerPoint.Shape first = pictures[0];
                double aspect = first.Width / first.Height;
                if (pictures.Any(picture => Math.Abs((picture.Width / picture.Height) / aspect - 1) > 0.005))
                {
                    MessageBox.Show("各通道的可见画面宽高比必须一致，请先调整图片裁剪范围。\n" +
                        "合并按画面坐标叠加，不自动配准；同一比例的图片会统一到第一张图片的输出尺寸。",
                        "合并通道", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                // 移除旋转后按局部画面坐标合并，输出继承第一张图的旋转角度。
                float angle = first.Rotation;
                if (pictures.Any(picture => Math.Abs(Math.IEEERemainder(picture.Rotation - angle, 360)) > 0.01))
                {
                    MessageBox.Show("各通道图片的旋转角度必须一致，请先统一旋转角度。", "合并通道");
                    return;
                }
                using (Globals.ThisAddIn.ZoomGuideLines?.Suspend())
                using (Bitmap merged = ChannelImageProcessor.CreateMergeCanvas(GetChannelRasterSize(first, ChannelImageMaxPixels)))
                {
                    foreach (PowerPoint.Shape picture in pictures)
                        using (Bitmap channel = ExportChannelPicture(picture, merged.Size))
                            ChannelImageProcessor.AddChannel(merged, channel);
                    app.StartNewUndoEntry();
                    PowerPoint.Shape output = InsertChannelBitmap(slide, merged,
                        pictures.Max(picture => picture.Left + picture.Width) + 18F, first.Top, first.Width, first.Height);
                    try
                    {
                        output.Rotation = angle;
                        output.Name = $"合并通道_{pictures.Length}通道";
                        output.AlternativeText = "合并通道（RGB 相加，超过 255 截断；不自动配准）：\n" +
                            string.Join("\n", pictures.Select(picture => picture.Name));
                        output.LockAspectRatio = Office.MsoTriState.msoTrue;
                        committed = true;
                        RestoreChannelSelection(new[] { output });
                    }
                    finally { if (!committed) TryDeleteChannelShape(output); }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"合并通道时出错：{ex.Message}", "合并通道", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally { if (!committed && pictures != null) RestoreChannelSelection(pictures); }
        }

        private static Size GetChannelRasterSize(PowerPoint.Shape picture, int maximum)
        {
            // 288 dpi 的显示画面，长边限制为 4096，预览单独降采样。
            double scale = Math.Min(4, maximum / (double)Math.Max(picture.Width, picture.Height));
            return new Size(Math.Max(1, (int)Math.Round(picture.Width * scale)),
                Math.Max(1, (int)Math.Round(picture.Height * scale)));
        }

        private Bitmap ExportChannelPicture(PowerPoint.Shape picture, Size size)
        {
            string path = Path.Combine(Path.GetTempPath(), "SlideSCI-channel-source-" + Guid.NewGuid().ToString("N") + ".png");
            PowerPoint.Shape copy = null;
            string step = "创建临时图片";
            try
            {
                copy = picture.Duplicate()[1];
                step = "清理临时图片关联";
                ZoomGuideLineTracker.RemoveCopiedLinks(copy);
                step = "移除图片旋转";
                copy.Rotation = 0;
                step = "移除图片描边";
                copy.Line.Visible = Office.MsoTriState.msoFalse;
                step = "移除图片阴影";
                copy.Shadow.Visible = Office.MsoTriState.msoFalse;
                step = "移除图片发光";
                copy.Glow.Radius = 0;
                step = "移除图片柔化边缘";
                copy.SoftEdge.Radius = 0;
                step = "移除图片映像";
                copy.Reflection.Type = Office.MsoReflectionType.msoReflectionTypeNone;
                step = "设置临时图片导出尺寸";
                // 在临时副本上调整尺寸，保持幻灯片导出参数有界。原图很小的时候，
                // 用目标像素 / 原图宽高反推幻灯片尺寸可能产生巨大的参数，触发 COM 越界。
                copy.LockAspectRatio = Office.MsoTriState.msoFalse;
                copy.Width = size.Width * 0.75F;
                copy.Height = size.Height * 0.75F;
                copy.Left = 0;
                copy.Top = 0;
                var page = app.ActivePresentation.PageSetup;
                step = "导出图片画面";
                // 默认 96 dpi：0.75 pt 对应 1 px。宽高参数保持为幻灯片尺寸。
                copy.Export(path, PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                    Math.Max(1, (int)Math.Round(page.SlideWidth)),
                    Math.Max(1, (int)Math.Round(page.SlideHeight)),
                    PowerPoint.PpExportMode.ppRelativeToSlide);
                step = "读取导出的图片";
                using (Image image = Image.FromFile(path)) return ChannelImageProcessor.Resize(image, size);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"{step}失败：{ex.Message}", ex);
            }
            finally
            {
                if (copy != null) TryDeleteChannelShape(copy);
                TryDeleteChannelFile(path);
            }
        }

        private static PowerPoint.Shape InsertChannelBitmap(PowerPoint.Slide slide, Bitmap image,
            float left, float top, float width, float height)
        {
            string path = Path.Combine(Path.GetTempPath(), "SlideSCI-channel-result-" + Guid.NewGuid().ToString("N") + ".png");
            try
            {
                image.Save(path, ImageFormat.Png);
                return slide.Shapes.AddPicture(path, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                    left, top, width, height);
            }
            finally { TryDeleteChannelFile(path); }
        }

        private static void RestoreChannelSelection(IEnumerable<PowerPoint.Shape> pictures)
        {
            try
            {
                bool first = true;
                foreach (PowerPoint.Shape picture in pictures)
                {
                    picture.Select(first ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
                    first = false;
                }
            }
            catch (Exception ex) { Debug.WriteLine($"恢复通道图片选区失败：{ex.Message}"); }
        }

        private static void TryDeleteChannelShape(PowerPoint.Shape shape)
        {
            try { shape.Delete(); }
            catch (Exception ex) { Debug.WriteLine($"清理通道临时图片失败：{ex.Message}"); }
        }

        private static void TryDeleteChannelFile(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            { Debug.WriteLine($"清理通道临时文件失败：{ex.Message}"); }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using Newtonsoft.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    /// <summary>FOV 持久化、原生比例尺形状与组合处理；Ribbon 和放大图联动共用。</summary>
    internal static class ScaleBarService
    {
        private const string SettingsPrefix = "比例尺（SlideSCI）：";
        private const string LegacySettingsPrefix = "标尺（SlideSCI）：";
        private const string GroupTag = "SLIDESCI_SCALE_GROUP";
        private const string ContentTag = "SLIDESCI_SCALE_CONTENT";
        private const string AnnotationTag = "SLIDESCI_SCALE_ANNOTATION";
        private const string VisibleFovTag = "SLIDESCI_FOV_VISIBLE";
        private const string SourceFovTag = "SLIDESCI_FOV_SOURCE";
        private const string ZoomTagPrefix = "SLIDESCI_ZOOM_";

        internal static string Tag(PowerPoint.Shape shape, string key)
        {
            for (int i = 1; i <= shape.Tags.Count; i++)
                if (string.Equals(shape.Tags.Name(i), key, StringComparison.OrdinalIgnoreCase)) return shape.Tags.Value(i);
            return null;
        }
        internal static bool IsScaleGroup(PowerPoint.Shape shape) =>
            shape.Type == Office.MsoShapeType.msoGroup && Tag(shape, GroupTag) == "1";
        internal static PowerPoint.Shape Content(PowerPoint.Shape shape)
        {
            if (!IsScaleGroup(shape)) return shape;
            return shape.GroupItems.Cast<PowerPoint.Shape>().Single(item => Tag(item, ContentTag) == "1");
        }
        internal static bool IsImage(PowerPoint.Shape shape)
        {
            if (IsScaleGroup(shape)) return true;
            return shape.Type == Office.MsoShapeType.msoPicture || shape.Type == Office.MsoShapeType.msoLinkedPicture ||
                (shape.Type != Office.MsoShapeType.msoGroup && shape.Fill.Type == Office.MsoFillType.msoFillPicture);
        }
        internal static ImageFieldOfView ReadFov(PowerPoint.Shape shape, bool trySource = true)
        {
            string text = Content(shape).AlternativeText ?? "";
            ImageFieldOfView saved = ImageFieldOfView.Read(text);
            if (saved != null || !trySource) return saved;
            string path = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .LastOrDefault(line => line.StartsWith(MediaSourceInsertion.PathPrefix, StringComparison.Ordinal));
            return path == null ? null : TiffFieldOfViewReader.TryRead(path.Substring(MediaSourceInsertion.PathPrefix.Length));
        }
        internal static ScaleBarSettings ReadSettings(PowerPoint.Shape shape)
        {
            string line = (Content(shape).AlternativeText ?? "").Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .LastOrDefault(IsSettingsLine);
            if (line == null) return null;
            string prefix = line.StartsWith(SettingsPrefix, StringComparison.Ordinal) ? SettingsPrefix : LegacySettingsPrefix;
            return ScaleBarSettings.Parse(line.Substring(prefix.Length));
        }
        private static bool IsSettingsLine(string line) => line.StartsWith(SettingsPrefix, StringComparison.Ordinal) ||
            line.StartsWith(LegacySettingsPrefix, StringComparison.Ordinal);

        internal static ScaleBarSettings FitZoomSettings(PowerPoint.Shape zoom, ScaleBarSettings settings,
            double? requestedLength = null, bool? showText = null)
        {
            if ((requestedLength ?? settings.LengthMicrometers) == 0)
                return settings.ForZoom(SizeF.Empty, 0, showText);
            ImageFieldOfView fov = ReadFov(zoom, false);
            if (fov == null) throw new InvalidOperationException("放大图没有有效的 FOV 信息。");
            return settings.ForZoom(VisibleFov(zoom, fov), requestedLength, showText);
        }

        internal static bool IsZoomScaleBarWithinLimit(PowerPoint.Shape zoom)
        {
            ScaleBarSettings settings = ReadSettings(zoom);
            return settings == null || settings.LengthMicrometers == 0 ||
                settings.LengthMicrometers <= FitZoomSettings(zoom, settings).LengthMicrometers;
        }
        private static void WriteSettings(PowerPoint.Shape picture, ScaleBarSettings settings)
        {
            string text = string.Join(Environment.NewLine, (picture.AlternativeText ?? "")
                .Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')
                .Where(line => !IsSettingsLine(line))).TrimEnd('\r', '\n');
            picture.AlternativeText = text + (text.Length == 0 ? "" : Environment.NewLine) +
                SettingsPrefix + JsonConvert.SerializeObject(settings);
        }
        internal static void WriteFov(PowerPoint.Shape shape, ImageFieldOfView fov)
        {
            PowerPoint.Shape picture = Content(shape);
            picture.AlternativeText = fov.Write(picture.AlternativeText);
            if (IsScaleGroup(shape)) shape.AlternativeText = picture.AlternativeText;
        }

        /// <summary>显示画面的物理范围；普通图片采用原图 FOV，局部取图采用可见画面的 FOV。</summary>
        internal static SizeF VisibleFov(PowerPoint.Shape shape, ImageFieldOfView fov)
        {
            PowerPoint.Shape picture = Content(shape);
            double width = fov.WidthMicrometers, height = fov.HeightMicrometers;
            if (Tag(picture, VisibleFovTag) != "1" &&
                (picture.Type == Office.MsoShapeType.msoPicture || picture.Type == Office.MsoShapeType.msoLinkedPicture))
            {
                // Crop.PictureWidth/Height 是包含被裁掉部分的整张图片在当前缩放下的尺寸。
                Office.Crop crop = picture.PictureFormat.Crop;
                if (!ImageFieldOfView.Positive(crop.PictureWidth) || !ImageFieldOfView.Positive(crop.PictureHeight))
                    throw new InvalidOperationException("无法读取图片的裁剪尺寸。");
                width *= picture.Width / crop.PictureWidth;
                height *= picture.Height / crop.PictureHeight;
            }
            if ((width != 0 && !ImageFieldOfView.Positive(width)) ||
                (height != 0 && !ImageFieldOfView.Positive(height)) ||
                (width == 0 && height == 0) || width > float.MaxValue || height > float.MaxValue)
                throw new InvalidOperationException("图片 FOV 或裁剪尺寸无效。");
            return new SizeF((float)width, (float)height);
        }

        /// <summary>复制组合里的纯图片到幻灯片，避免将原图比例尺烘焙进放大图。</summary>
        internal static PowerPoint.Shape DuplicateContent(PowerPoint.Shape source)
        {
            PowerPoint.Shape copy = source.Duplicate()[1];
            copy.Left = source.Left;
            copy.Top = source.Top;
            if (!IsScaleGroup(copy)) return copy;
            var links = LinkTags(source);
            PowerPoint.Shape[] children = null;
            try
            {
                float rotation = copy.Rotation;
                copy.Rotation = 0;
                children = copy.Ungroup().Cast<PowerPoint.Shape>().ToArray();
                PowerPoint.Shape picture = children.Single(item => Tag(item, ContentTag) == "1");
                foreach (PowerPoint.Shape child in children) if (child.Id != picture.Id) child.Delete();
                picture.Tags.Delete(ContentTag);
                picture.Rotation = rotation;
                foreach (var link in links) picture.Tags.Add(link.Key, link.Value);
                return picture;
            }
            catch
            {
                if (children == null) TryDelete(copy);
                else foreach (PowerPoint.Shape child in children) TryDelete(child);
                throw;
            }
        }

        internal static IList<PowerPoint.Shape> Annotations(PowerPoint.Slide slide, PowerPoint.Shape image)
        {
            if (IsScaleGroup(image)) return new PowerPoint.Shape[0];
            string owner = image.Id.ToString(CultureInfo.InvariantCulture);
            return slide.Shapes.Cast<PowerPoint.Shape>().Where(shape => Tag(shape, AnnotationTag) == owner).ToArray();
        }
        internal static void DeleteWithAnnotations(PowerPoint.Slide slide, PowerPoint.Shape image)
        {
            foreach (PowerPoint.Shape annotation in Annotations(slide, image)) annotation.Delete();
            image.Delete();
        }

        /// <summary>按对象 ID 从当前 Shapes 集合定位实际下标，避免复制、取消编组后使用失效的层级位置。</summary>
        internal static PowerPoint.ShapeRange RangeByIds(PowerPoint.Slide slide, IEnumerable<PowerPoint.Shape> members)
        {
            var requestedIds = new HashSet<int>(members.Select(member => member.Id));
            if (requestedIds.Count == 0) throw new InvalidOperationException("没有需要组合的形状。");
            PowerPoint.Shapes shapes = slide.Shapes;
            var indices = new List<object>();
            int count = shapes.Count;
            // 不依赖形状名称：原图与用于更新的副本可以重名。
            for (int index = 1; index <= count; index++)
                if (requestedIds.Remove(shapes[index].Id)) indices.Add(index);
            if (requestedIds.Count != 0)
                throw new InvalidOperationException("待处理形状已不在当前幻灯片中，请重新选中图片后再试。");
            return shapes.Range(indices.ToArray());
        }

        /// <summary>在副本上完成比例尺后再替换原对象，任何生成失败都保留原图与原比例尺。</summary>
        internal static PowerPoint.Shape Replace(PowerPoint.Slide slide, PowerPoint.Shape original,
            ImageFieldOfView fov, ScaleBarSettings settings)
        {
            var created = new List<PowerPoint.Shape>();
            PowerPoint.Shape result = null;
            bool committed = false;
            try
            {
                PowerPoint.Shape picture = DuplicateContent(original);
                created.Add(picture);
                picture.Name = original.Name;
                if (fov.IsValid) WriteFov(picture, fov);
                result = Add(slide, picture, settings, created);
                result.Name = original.Name;
                int target = original.ZOrderPosition + 1;
                // 整组位于原对象的正上方，删除原对象后保持层级。
                if (settings.GroupWithImage)
                    while (result.ZOrderPosition > target) result.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                else
                    while (picture.ZOrderPosition > target) picture.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                DeleteWithAnnotations(slide, original);
                committed = true;
                return result;
            }
            finally
            {
                if (!committed) foreach (PowerPoint.Shape shape in created.AsEnumerable().Reverse()) TryDelete(shape);
                Globals.ThisAddIn.ZoomGuideLines?.InvalidateCache();
            }
        }

        /// <summary>对新图片添加比例尺；返回组合或纯图片。created 用于调用方事务回滚。</summary>
        internal static PowerPoint.Shape Add(PowerPoint.Slide slide, PowerPoint.Shape picture,
            ScaleBarSettings settings, IList<PowerPoint.Shape> created)
        {
            if (settings == null || !settings.IsValid) throw new InvalidOperationException("比例尺参数无效。");
            if (settings.LengthMicrometers == 0)
            {
                WriteSettings(picture, settings);
                return picture;
            }
            ImageFieldOfView fov = ReadFov(picture, false);
            if (fov == null) throw new InvalidOperationException("图片没有有效的 FOV 信息。");
            if (!fov.IsValidFor(settings.Vertical)) throw new InvalidOperationException(settings.Vertical
                ? "纵向比例尺需要填写 FOV 高度。" : "横向比例尺需要填写 FOV 宽度。");
            SizeF visible = VisibleFov(picture, fov);
            // 标签先以原生文本框测量，再共同定位，确保完整落在图片内部。
            PowerPoint.Shape text = null;
            float textWidth = 0, textHeight = 0;
            float rotation = picture.Rotation;
            var members = new List<PowerPoint.Shape> { picture };
            string step = "创建比例尺文字";
            try
            {
                if (settings.ShowText)
                {
                    text = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 1, 1);
                    created.Add(text);
                    members.Add(text);
                    text.Name = "比例尺文字 " + picture.Id;
                    text.Fill.Visible = text.Line.Visible = Office.MsoTriState.msoFalse;
                    text.TextFrame.MarginLeft = text.TextFrame.MarginRight = 0;
                    text.TextFrame.MarginTop = text.TextFrame.MarginBottom = 0;
                    text.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
                    text.TextFrame.TextRange.Text = settings.Label;
                    var font = text.TextFrame.TextRange.Font;
                    font.Name = "Arial";
                    font.Size = settings.FontSize;
                    font.Bold = settings.Bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    font.Color.RGB = ColorTranslator.ToOle(Color.FromArgb(settings.FontColorArgb));
                    text.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                    textWidth = text.Width;
                    textHeight = text.Height;
                }
                step = "计算比例尺布局";
                ScaleBarLayout layout = ScaleBarLayout.Calculate(new SizeF(picture.Width, picture.Height), visible,
                    settings, new SizeF(textWidth, textHeight));
                step = "创建比例尺矩形";
                var bar = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, 0, 0, layout.Bar.Width, layout.Bar.Height);
                created.Add(bar);
                members.Add(bar);
                bar.Name = "比例尺 " + picture.Id;
                bar.Line.Visible = Office.MsoTriState.msoFalse;
                bar.Fill.Solid();
                bar.Fill.Transparency = 0;
                bar.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(settings.BarColorArgb));
                SetLocalBounds(bar, picture, layout.Bar.X, layout.Bar.Y, rotation, !settings.GroupWithImage);
                if (text != null)
                {
                    text.TextFrame.TextRange.ParagraphFormat.Alignment =
                        settings.Corner == ScaleBarCorner.TopRight || settings.Corner == ScaleBarCorner.BottomRight
                            ? PowerPoint.PpParagraphAlignment.ppAlignRight : PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    SetLocalBounds(text, picture, layout.Text.X, layout.Text.Y, rotation, !settings.GroupWithImage);
                }
                step = "保存比例尺参数";
                WriteSettings(picture, settings);
                foreach (PowerPoint.Shape member in members.Skip(1))
                    member.Tags.Add(AnnotationTag, picture.Id.ToString(CultureInfo.InvariantCulture));
                if (!settings.GroupWithImage) return picture;
                var links = LinkTags(picture);
                // 先按图片局部坐标编组，再整体恢复旋转，组的宽高与图片画框一致。
                picture.Rotation = 0;
                picture.Tags.Add(ContentTag, "1");
                step = "编组图片与比例尺";
                PowerPoint.Shape group = RangeByIds(slide, members).Group();
                created.Add(group);
                step = "保存组合信息";
                group.Tags.Add(GroupTag, "1");
                group.AlternativeText = picture.AlternativeText;
                foreach (var link in links) { group.Tags.Add(link.Key, link.Value); picture.Tags.Delete(link.Key); }
                group.Rotation = rotation;
                return group;
            }
            catch (Exception ex)
            {
                // 对新对象的清理由调用者负责；先恢复尚未编组的图片旋转。
                try { picture.Rotation = rotation; } catch (Exception restoreError) { Debug.WriteLine(restoreError); }
                throw new InvalidOperationException($"{step}失败：{ex.Message}", ex);
            }
        }

        private static void SetLocalBounds(PowerPoint.Shape item, PowerPoint.Shape picture, float x, float y,
            float rotation, bool rotate)
        {
            float cx = picture.Left + x + item.Width / 2, cy = picture.Top + y + item.Height / 2;
            if (rotate)
            {
                double angle = rotation * Math.PI / 180;
                float px = picture.Left + picture.Width / 2, py = picture.Top + picture.Height / 2;
                float dx = cx - px, dy = cy - py;
                cx = px + (float)(dx * Math.Cos(angle) - dy * Math.Sin(angle));
                cy = py + (float)(dx * Math.Sin(angle) + dy * Math.Cos(angle));
                item.Rotation = rotation;
            }
            item.Left = cx - item.Width / 2;
            item.Top = cy - item.Height / 2;
        }
        private static IDictionary<string, string> LinkTags(PowerPoint.Shape shape)
        {
            var links = new Dictionary<string, string>();
            for (int i = 1; i <= shape.Tags.Count; i++)
                if (shape.Tags.Name(i).StartsWith(ZoomTagPrefix, StringComparison.OrdinalIgnoreCase))
                    links[shape.Tags.Name(i)] = shape.Tags.Value(i);
            return links;
        }
        internal static void TryDelete(PowerPoint.Shape shape)
        { try { shape.Delete(); } catch (Exception ex) { Debug.WriteLine(ex); } }

        /// <summary>在缩放前从实际取图的尺寸反推其 FOV，旋转区域使用两个像素标定轴换算。</summary>
        internal static void SetCropFov(PowerPoint.Shape source, PowerPoint.Shape crop)
        {
            ImageFieldOfView fov = ReadFov(source);
            if (fov == null) return;
            SizeF visible = VisibleFov(source, fov);
            ImageFieldOfView cropFov = CalculateCropFov(new SizeF(source.Width, source.Height), visible,
                new SizeF(crop.Width, crop.Height), crop.Rotation - source.Rotation);
            if (!cropFov.IsValid)
                throw new InvalidOperationException("旋转取图需要 FOV 宽度和高度，请先补充标定。");
            WriteFov(crop, cropFov);
            crop.Tags.Add(VisibleFovTag, "1");
            crop.Tags.Add(SourceFovTag, CalibrationStamp(source));
        }

        /// <summary>取图前的尺寸和 FOV 使用各自一致的单位；返回 μm，预览与生成共用。</summary>
        internal static ImageFieldOfView CalculateCropFov(SizeF sourceSize, SizeF visibleFov, SizeF cropSize, float rotation)
        {
            double x = visibleFov.Width / sourceSize.Width, y = visibleFov.Height / sourceSize.Height;
            double angle = rotation * Math.PI / 180;
            double cos = Math.Cos(angle), sin = Math.Sin(angle);
            return new ImageFieldOfView
            {
                Width = cropSize.Width * ProjectCalibration(x, y, cos, sin),
                Height = cropSize.Height * ProjectCalibration(x, y, sin, cos),
                Unit = "μm", Source = "由原图标定和实际取图范围换算"
            };
        }

        private static double ProjectCalibration(double x, double y, double xWeight, double yWeight)
        {
            // 任意角度旋转可能同时依赖两个标定轴；未知轴不能按零像素尺寸参与换算。
            const double epsilon = 0.000001;
            if ((Math.Abs(xWeight) > epsilon && !ImageFieldOfView.Positive(x)) ||
                (Math.Abs(yWeight) > epsilon && !ImageFieldOfView.Positive(y))) return 0;
            return Math.Sqrt(x * x * xWeight * xWeight + y * y * yWeight * yWeight);
        }

        private static string CalibrationStamp(PowerPoint.Shape source)
        {
            ImageFieldOfView fov = ReadFov(source);
            if (fov == null) return "";
            SizeF visible = VisibleFov(source, fov);
            return visible.Width.ToString("R", CultureInfo.InvariantCulture) + ";" +
                visible.Height.ToString("R", CultureInfo.InvariantCulture);
        }
        internal static bool HasCurrentCalibration(PowerPoint.Shape source, PowerPoint.Shape crop) =>
            string.Equals(Tag(Content(crop), SourceFovTag) ?? "", CalibrationStamp(source), StringComparison.Ordinal);
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    /// <summary>文件选择与文件粘贴共用的插入流程，路径与返回的 Shape 直接对应。</summary>
    internal static class MediaSourceInsertion
    {
        internal const string PathPrefix = "原图路径：";
        internal const string VideoPathPrefix = "视频路径：";
        private static readonly HashSet<string> ImageExtensions = new HashSet<string>(
            new[] { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".svg", ".emf", ".wmf" },
            StringComparer.OrdinalIgnoreCase);
        // 扩展名用于判断是否尝试视频插入；实际编码能否播放由 PowerPoint 决定。
        private static readonly HashSet<string> VideoExtensions = new HashSet<string>(
            new[] { ".mp4", ".m4v", ".mov", ".webm", ".avi", ".wmv", ".asf", ".mpg", ".mpeg", ".mpe" },
            StringComparer.OrdinalIgnoreCase);

        internal static bool IsSupportedMediaFile(string path) =>
            !string.IsNullOrEmpty(path) && File.Exists(path) &&
            (ImageExtensions.Contains(Path.GetExtension(path)) || VideoExtensions.Contains(Path.GetExtension(path)));

        internal static IList<string> Insert(Application application, Slide slide, IEnumerable<string> paths,
            PointF? dropPosition = null)
        {
            var insertedNames = new List<object>();
            var failures = new List<string>();
            var usedNames = new HashSet<string>(slide.Shapes.Cast<Shape>().Select(shape => shape.Name),
                StringComparer.OrdinalIgnoreCase);
            float slideWidth = application.ActivePresentation.PageSetup.SlideWidth;
            float slideHeight = application.ActivePresentation.PageSetup.SlideHeight;
            application.StartNewUndoEntry();
            using (Globals.ThisAddIn.ZoomGuideLines?.Suspend())
            {
                foreach (string path in paths)
                {
                    Shape media = null;
                    try
                    {
                        string sourcePath = Path.GetFullPath(path);
                        bool isVideo = VideoExtensions.Contains(Path.GetExtension(sourcePath));
                        // 视频必须插入为可播放的媒体对象，不能作为图片或文件图标插入。
                        media = isVideo
                            ? slide.Shapes.AddMediaObject2(sourcePath, Office.MsoTriState.msoFalse,
                                Office.MsoTriState.msoTrue, 0, 0, -1, -1)
                            : slide.Shapes.AddPicture(sourcePath, Office.MsoTriState.msoFalse,
                                Office.MsoTriState.msoTrue, 0, 0, -1, -1);
                        if (isVideo && media.MediaType != PpMediaType.ppMediaTypeMovie)
                            throw new InvalidOperationException("PowerPoint 未将此文件识别为视频。");
                        float scale = Math.Min(1f, Math.Min(slideWidth * 0.9f / media.Width,
                            slideHeight * 0.9f / media.Height));
                        media.LockAspectRatio = Office.MsoTriState.msoTrue;
                        if (scale < 1f) media.Width *= scale;
                        // 拖放时以落点为中心；文件选择和粘贴仍放在页面中心。
                        media.Left = dropPosition.HasValue
                            ? Math.Max(0, Math.Min(slideWidth - media.Width, dropPosition.Value.X - media.Width / 2f))
                            : (slideWidth - media.Width) / 2f;
                        media.Top = dropPosition.HasValue
                            ? Math.Max(0, Math.Min(slideHeight - media.Height, dropPosition.Value.Y - media.Height / 2f))
                            : (slideHeight - media.Height) / 2f;

                        string description = media.AlternativeText ?? string.Empty;
                        media.AlternativeText = description +
                            (description.Length == 0 ? string.Empty : Environment.NewLine) +
                            (isVideo ? VideoPathPrefix : PathPrefix) + sourcePath;
                        media.Name = GetAvailableName(sourcePath, usedNames);
                        usedNames.Add(media.Name);
                        insertedNames.Add(media.Name);
                    }
                    catch (Exception ex)
                    {
                        // 只清理本次新建且未完成记录的对象。
                        if (media != null)
                        {
                            try { media.Delete(); }
                            catch (Exception cleanupError) { Debug.WriteLine(cleanupError); }
                        }
                        failures.Add($"{Path.GetFileName(path)}：{ex.Message}");
                    }
                }

                // 选中失败不影响文件已经插入和记录路径的事实。
                if (insertedNames.Count > 0)
                {
                    try { slide.Shapes.Range(insertedNames.ToArray()).Select(); }
                    catch (Exception ex) { Debug.WriteLine($"选中插入对象失败：{ex.Message}"); }
                }
            }
            return failures;
        }

        private static string GetAvailableName(string sourcePath, ISet<string> usedNames)
        {
            string fileName = Path.GetFileName(sourcePath);
            string candidate = fileName;
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            int suffix = 2;
            while (usedNames.Contains(candidate)) candidate = $"{baseName} ({suffix++}){extension}";
            return candidate;
        }
    }
}

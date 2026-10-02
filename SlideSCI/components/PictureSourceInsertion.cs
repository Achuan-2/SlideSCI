using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    /// <summary>文件选择与文件粘贴共用的插入流程，路径与返回的 Shape 直接对应。</summary>
    internal static class PictureSourceInsertion
    {
        internal const string PathPrefix = "原图路径：";
        private static readonly HashSet<string> ImageExtensions = new HashSet<string>(
            new[] { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".svg", ".emf", ".wmf" },
            StringComparer.OrdinalIgnoreCase);

        internal static bool IsSupportedImageFile(string path) =>
            !string.IsNullOrEmpty(path) && ImageExtensions.Contains(Path.GetExtension(path)) && File.Exists(path);

        internal static IList<string> Insert(Application application, Slide slide, IEnumerable<string> paths)
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
                    Shape picture = null;
                    try
                    {
                        string sourcePath = Path.GetFullPath(path);
                        picture = slide.Shapes.AddPicture(sourcePath, Office.MsoTriState.msoFalse,
                            Office.MsoTriState.msoTrue, 0, 0, -1, -1);
                        float scale = Math.Min(1f, Math.Min(slideWidth * 0.9f / picture.Width,
                            slideHeight * 0.9f / picture.Height));
                        picture.LockAspectRatio = Office.MsoTriState.msoTrue;
                        if (scale < 1f) picture.Width *= scale;
                        picture.Left = (slideWidth - picture.Width) / 2f;
                        picture.Top = (slideHeight - picture.Height) / 2f;

                        string description = picture.AlternativeText ?? string.Empty;
                        picture.AlternativeText = description +
                            (description.Length == 0 ? string.Empty : Environment.NewLine) + PathPrefix + sourcePath;
                        picture.Name = GetAvailableName(sourcePath, usedNames);
                        usedNames.Add(picture.Name);
                        insertedNames.Add(picture.Name);
                    }
                    catch (Exception ex)
                    {
                        // 只清理本次新建且未完成记录的图片。
                        if (picture != null)
                        {
                            try { picture.Delete(); }
                            catch (Exception cleanupError) { Debug.WriteLine(cleanupError); }
                        }
                        failures.Add($"{Path.GetFileName(path)}：{ex.Message}");
                    }
                }

                // 选中失败不影响图片已经插入和记录路径的事实。
                if (insertedNames.Count > 0)
                {
                    try { slide.Shapes.Range(insertedNames.ToArray()).Select(); }
                    catch (Exception ex) { Debug.WriteLine($"选中插入图片失败：{ex.Message}"); }
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

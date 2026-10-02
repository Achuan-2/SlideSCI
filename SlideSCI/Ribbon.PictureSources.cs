using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        private void InsertPicturesWithSource_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (app.Windows.Count == 0 || app.ActiveWindow == null)
                {
                    MessageBox.Show("请先打开一个演示文稿，并切换到要插入图片的幻灯片。", "提示");
                    return;
                }

                Slide slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null)
                {
                    MessageBox.Show("请先切换到要插入图片的幻灯片。", "提示");
                    return;
                }

                using (var dialog = new OpenFileDialog
                {
                    Title = "插入图片并记录原图路径",
                    Filter = "图片文件 (*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.svg;*.emf;*.wmf)|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff;*.svg;*.emf;*.wmf|所有文件 (*.*)|*.*",
                    Multiselect = true,
                    CheckFileExists = true,
                    RestoreDirectory = true
                })
                {
                    if (dialog.ShowDialog(new PowerPointDialogOwner(app.HWND)) != DialogResult.OK) return;

                    var failures = MediaSourceInsertion.Insert(app, slide, dialog.FileNames);
                    if (failures.Count > 0)
                        MessageBox.Show($"已插入 {dialog.FileNames.Length - failures.Count} 张图片。以下文件插入失败：\n\n" +
                            string.Join(Environment.NewLine, failures), "插入图片", MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入图片时出错：{ex.Message}", "插入图片", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void OpenPictureSourceLocation_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (app.Windows.Count == 0 || app.ActiveWindow == null)
                {
                    MessageBox.Show("请先打开一个演示文稿，并选中一张记录了路径的图片。", "提示");
                    return;
                }

                Selection selection = app.ActiveWindow.Selection;
                if (selection.Type != PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("请选中一张记录了路径的图片。", "提示");
                    return;
                }
                ShapeRange shapes = selection.HasChildShapeRange ? selection.ChildShapeRange : selection.ShapeRange;
                if (shapes.Count != 1)
                {
                    MessageBox.Show("请只选中一张图片；组合中的图片请先单独选中。", "提示");
                    return;
                }

                string sourceLine = (shapes[1].AlternativeText ?? string.Empty)
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .LastOrDefault(line => line.StartsWith(MediaSourceInsertion.PathPrefix, StringComparison.Ordinal) ||
                        line.StartsWith(MediaSourceInsertion.VideoPathPrefix, StringComparison.Ordinal));
                if (sourceLine == null)
                {
                    MessageBox.Show("这张图片没有记录原图路径。请通过“插入图片”插入，或从资源管理器复制图片文件后在幻灯片中普通粘贴。", "提示");
                    return;
                }

                string prefix = sourceLine.StartsWith(MediaSourceInsertion.VideoPathPrefix, StringComparison.Ordinal)
                    ? MediaSourceInsertion.VideoPathPrefix : MediaSourceInsertion.PathPrefix;
                string sourcePath = sourceLine.Substring(prefix.Length);
                // 替代文字可由用户编辑；只接受完整文件路径，不把内容作为命令执行。
                if (!Path.IsPathRooted(sourcePath) || sourcePath.IndexOf('"') >= 0)
                {
                    MessageBox.Show($"记录的原图路径无效：\n{sourcePath}", "提示");
                    return;
                }
                sourcePath = Path.GetFullPath(sourcePath);
                if (!File.Exists(sourcePath))
                {
                    MessageBox.Show($"找不到原图，文件可能已移动、删除，或所在磁盘未连接。\n\n记录的路径：\n{sourcePath}",
                        "原图不存在", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = "/select,\"" + sourcePath + "\"",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开原图位置时出错：{ex.Message}", "打开原图位置", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}

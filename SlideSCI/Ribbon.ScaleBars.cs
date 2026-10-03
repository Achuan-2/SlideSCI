using System;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        private void ScaleBar_Click(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Shape[] pictures = null;
            string step = "检查所选图片";
            try
            {
                if (app.Windows.Count == 0 || app.ActiveWindow == null ||
                    app.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                { MessageBox.Show("请先选中图片或本插件生成的图片比例尺组合。", "提示"); return; }
                var selection = app.ActiveWindow.Selection;
                if (selection.HasChildShapeRange)
                { MessageBox.Show("请选中整个图片比例尺组合后再操作。", "提示"); return; }
                pictures = selection.ShapeRange.Cast<PowerPoint.Shape>().ToArray();
                if (pictures.Length == 0 || pictures.Any(shape => !ScaleBarService.IsImage(shape)))
                { MessageBox.Show("请选择图片或本插件生成的图片比例尺组合。", "提示"); return; }
                var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null) { MessageBox.Show("请切换到图片所在的幻灯片。", "提示"); return; }
                using (Globals.ThisAddIn.ZoomGuideLines?.Suspend())
                {
                    // 每张图分别显示自身 FOV 和已有比例尺，避免把首图标定误用于其他图。
                    for (int i = 0; i < pictures.Length; i++)
                    {
                        PowerPoint.Shape picture = pictures[i];
                        step = "读取 FOV 与比例尺参数";
                        ImageFieldOfView fov = ScaleBarService.ReadFov(picture);
                        ScaleBarSettings existing = ScaleBarService.ReadSettings(picture);
                        ScaleBarSettings defaults = ScaleBarSettings.Parse(Properties.Settings.Default.ScaleBarOptions) ?? new ScaleBarSettings();
                        SizeF pictureSize = new SizeF(picture.Width, picture.Height);
                        // 用单位 FOV 获取裁剪比例，编辑 FOV 时预览仍按同一可见范围换算。
                        SizeF visibleFraction = ScaleBarService.VisibleFov(picture,
                            new ImageFieldOfView { Width = 1, Height = 1 });
                        step = "准备图片预览";
                        using (Bitmap preview = ExportChannelPicture(picture, GetChannelRasterSize(picture, 960), true))
                        using (var dialog = new ScaleBarForm(picture.Name, fov, existing ?? defaults,
                            preview, pictureSize, visibleFraction))
                        {
                            RestoreChannelSelection(pictures);
                            step = "编辑比例尺设置";
                            if (dialog.ShowDialog(new PowerPointDialogOwner(app.HWND)) != DialogResult.OK) break;
                            app.StartNewUndoEntry();
                            ScaleBarSettings options = dialog.BarSettings;
                            step = "生成比例尺";
                            pictures[i] = ScaleBarService.Replace(slide, picture, dialog.FieldOfView, options);
                            Properties.Settings.Default.ScaleBarOptions = JsonConvert.SerializeObject(options);
                            try { Properties.Settings.Default.Save(); }
                            catch (Exception ex) { Debug.WriteLine($"保存比例尺默认设置失败：{ex.Message}"); }
                            pictures[i].Select();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show($"设置比例尺时出错（{step}）：{ex.Message}", "操作失败",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (pictures != null) RestoreChannelSelection(pictures);
            }
        }
    }
}

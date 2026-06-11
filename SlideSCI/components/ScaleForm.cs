using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    public class ScaleForm : Form
    {
        private PowerPoint.Application app;
        private List<int> selectedShapeIdsByOrder;

        // UI Controls
        private Label lblInfo;
        private Panel cardScale;
        private Label lblScale;
        private NumericUpDown numScale;
        private Label lblPercentSign;
        private TrackBar trackScale;
        private Button btnClose;

        // Data cache
        private class ShapeScaleInfo
        {
            public PowerPoint.Shape ActualShape { get; set; }
            public float OriginalLeft { get; set; }
            public float OriginalTop { get; set; }
            public float OriginalWidth { get; set; }
            public float OriginalHeight { get; set; }
            public List<TextRunInfo> TextRuns { get; set; } = new List<TextRunInfo>();
            public List<ShapeScaleInfo> GroupChildren { get; set; } = new List<ShapeScaleInfo>();
        }

        private class TextRunInfo
        {
            public PowerPoint.TextRange TextRange { get; set; }
            public float OriginalFontSize { get; set; }
        }

        private List<ShapeScaleInfo> scaleInfos = new List<ShapeScaleInfo>();
        private bool isInitializing = false;
        private bool isApplying = false;
        private bool isUpdatingControls = false;

        public ScaleForm(PowerPoint.Application app, List<int> selectedShapeIdsByOrder)
        {
            this.app = app;
            this.selectedShapeIdsByOrder = selectedShapeIdsByOrder;
            
            InitializeComponent();
            RefreshSelection();
        }

        private void InitializeComponent()
        {
            // Form properties
            this.Text = "图文同缩比例设置";
            this.ClientSize = new Size(304, 200); // Compact layout
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.TopMost = true;
            this.BackColor = Color.FromArgb(240, 242, 245); // Fluent light gray
            this.Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;

            // 1. Info Label
            lblInfo = new Label
            {
                Location = new Point(15, 10),
                Size = new Size(274, 38),
                ForeColor = Color.FromArgb(64, 64, 64),
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "正在读取当前选中的形状..."
            };
            this.Controls.Add(lblInfo);

            // 2. Scale Card
            cardScale = new Panel
            {
                Location = new Point(15, 53),
                Size = new Size(274, 85),
                BackColor = Color.White
            };
            cardScale.Paint += DrawCardBorder;
            this.Controls.Add(cardScale);

            lblScale = new Label
            {
                Text = "缩放比例",
                Location = new Point(12, 12),
                AutoSize = true, // Auto-size to prevent clipping
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(205, 74, 38)
            };
            cardScale.Controls.Add(lblScale);

            numScale = new NumericUpDown
            {
                Minimum = 10,
                Maximum = 500,
                Value = 100,
                Location = new Point(160, 10),
                Size = new Size(80, 23)
            };
            numScale.ValueChanged += numScale_ValueChanged;
            cardScale.Controls.Add(numScale);

            lblPercentSign = new Label
            {
                Text = "%",
                Location = new Point(242, 12),
                AutoSize = true, // Auto-size to prevent clipping
                Font = new Font("Microsoft YaHei", 9.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(64, 64, 64)
            };
            cardScale.Controls.Add(lblPercentSign);

            trackScale = new TrackBar
            {
                Minimum = 10,
                Maximum = 300,
                Value = 100,
                TickStyle = TickStyle.None,
                Location = new Point(8, 42),
                Size = new Size(258, 30)
            };
            trackScale.Scroll += trackScale_Scroll;
            cardScale.Controls.Add(trackScale);

            // 3. Close Button
            btnClose = new Button
            {
                Text = "关闭",
                Width = 80,
                Height = 32,
                Location = new Point(112, 153),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(224, 224, 224),
                ForeColor = Color.FromArgb(51, 51, 51),
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.FlatAppearance.MouseOverBackColor = Color.FromArgb(200, 200, 200);
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);
        }

        private void DrawCardBorder(object sender, PaintEventArgs e)
        {
            Panel panel = sender as Panel;
            if (panel != null)
            {
                using (Pen pen = new Pen(Color.FromArgb(224, 224, 224), 1))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, panel.Width - 1, panel.Height - 1);
                }
            }
        }

        public void OnSelectionChanged()
        {
            RefreshSelection();
        }

        public void RefreshSelection()
        {
            if (isApplying) return;

            scaleInfos.Clear();
            try
            {
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0)
                {
                    lblInfo.Text = "当前未选中任何形状。";
                    SetControlsEnabled(false);
                    return;
                }

                SetControlsEnabled(true);
                lblInfo.Text = $"已选中 {sel.ShapeRange.Count} 个形状";

                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    RecordShapeScaleInfo(shape, scaleInfos);
                }

                // Reset controls back to 100% representing original size
                isInitializing = true;
                numScale.Value = 100;
                trackScale.Value = 100;
                isInitializing = false;
            }
            catch (Exception ex)
            {
                lblInfo.Text = $"加载选区失败：{ex.Message}";
                SetControlsEnabled(false);
            }
        }

        private void RecordShapeScaleInfo(PowerPoint.Shape shape, List<ShapeScaleInfo> list)
        {
            var info = new ShapeScaleInfo
            {
                ActualShape = shape,
                OriginalLeft = shape.Left,
                OriginalTop = shape.Top,
                OriginalWidth = shape.Width,
                OriginalHeight = shape.Height
            };

            RecordTextRuns(shape, info.TextRuns);

            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                try
                {
                    foreach (PowerPoint.Shape child in shape.GroupItems)
                    {
                        RecordGroupChildShape(child, info.GroupChildren);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error traversing group items: {ex.Message}");
                }
            }

            list.Add(info);
        }

        private void RecordGroupChildShape(PowerPoint.Shape child, List<ShapeScaleInfo> list)
        {
            var info = new ShapeScaleInfo
            {
                ActualShape = child,
                OriginalLeft = child.Left,
                OriginalTop = child.Top,
                OriginalWidth = child.Width,
                OriginalHeight = child.Height
            };

            RecordTextRuns(child, info.TextRuns);

            if (child.Type == Office.MsoShapeType.msoGroup)
            {
                try
                {
                    foreach (PowerPoint.Shape nestedChild in child.GroupItems)
                    {
                        RecordGroupChildShape(nestedChild, info.GroupChildren);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error traversing nested group: {ex.Message}");
                }
            }

            list.Add(info);
        }

        private void RecordTextRuns(PowerPoint.Shape shape, List<TextRunInfo> runList)
        {
            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    var textRange = shape.TextFrame.TextRange;
                    var runs = textRange.Runs();
                    if (runs != null && runs.Count > 0)
                    {
                        for (int i = 1; i <= runs.Count; i++)
                        {
                            var run = textRange.Runs(i, 1);
                            runList.Add(new TextRunInfo
                            {
                                TextRange = run,
                                OriginalFontSize = run.Font.Size
                            });
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error recording text runs: {ex.Message}");
            }
        }

        private void SetControlsEnabled(bool enabled)
        {
            lblScale.Enabled = enabled;
            numScale.Enabled = enabled;
            trackScale.Enabled = enabled;
        }

        private void ApplyScaling(float percent)
        {
            if (isInitializing || scaleInfos == null || scaleInfos.Count == 0) return;
            if (isApplying) return;

            isApplying = true;
            try
            {
                float S = percent / 100f;

                // Find bounding box of the top-level selection
                float minLeft = scaleInfos.Min(info => info.OriginalLeft);
                float minTop = scaleInfos.Min(info => info.OriginalTop);

                foreach (var info in scaleInfos)
                {
                    // 1. Scale size of the top-level shape
                    info.ActualShape.Width = info.OriginalWidth * S;
                    info.ActualShape.Height = info.OriginalHeight * S;

                    // 2. Scale position relative to top-left of selection
                    float dX = info.OriginalLeft - minLeft;
                    float dY = info.OriginalTop - minTop;
                    info.ActualShape.Left = minLeft + dX * S;
                    info.ActualShape.Top = minTop + dY * S;

                    // 3. Scale text runs in top-level shape
                    foreach (var run in info.TextRuns)
                    {
                        run.TextRange.Font.Size = run.OriginalFontSize * S;
                    }

                    // 4. Scale text runs in nested child shapes recursively
                    ScaleGroupChildrenText(info, S);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying scale: {ex.Message}");
            }
            finally
            {
                isApplying = false;
            }
        }

        private void ScaleGroupChildrenText(ShapeScaleInfo info, float S)
        {
            foreach (var child in info.GroupChildren)
            {
                foreach (var run in child.TextRuns)
                {
                    run.TextRange.Font.Size = run.OriginalFontSize * S;
                }
                ScaleGroupChildrenText(child, S);
            }
        }

        private void numScale_ValueChanged(object sender, EventArgs e)
        {
            if (isUpdatingControls) return;
            isUpdatingControls = true;

            int val = (int)numScale.Value;
            if (val >= trackScale.Minimum && val <= trackScale.Maximum)
            {
                trackScale.Value = val;
            }

            isUpdatingControls = false;
            ApplyScaling(val);
        }

        private void trackScale_Scroll(object sender, EventArgs e)
        {
            if (isUpdatingControls) return;
            isUpdatingControls = true;

            numScale.Value = trackScale.Value;

            isUpdatingControls = false;
            ApplyScaling(trackScale.Value);
        }
    }
}

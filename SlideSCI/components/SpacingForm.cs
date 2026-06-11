using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    public class SpacingForm : Form
    {
        private PowerPoint.Application app;
        private List<int> selectedShapeIdsByOrder;

        private const float PointsPerCm = 28.3464593f;

        // UI Controls
        private Label lblInfo;
        
        private Panel cardDistribute;
        private Button btnDistributeH;
        private Button btnDistributeV;

        private Panel cardHorizontal;
        private Label lblHorizontal;
        private NumericUpDown numHorizontal;
        private TrackBar trackHorizontal;
        
        private Panel cardVertical;
        private Label lblVertical;
        private NumericUpDown numVertical;
        private TrackBar trackVertical;
        
        private Button btnClose;

        // Data cache
        private class ShapeSnapshot
        {
            public PowerPoint.Shape ActualShape { get; set; }
            public int Id { get; set; }
            public float OriginalLeft { get; set; }
            public float OriginalTop { get; set; }
            public float Width { get; set; }
            public float Height { get; set; }
            public bool IsBaseline { get; set; }
        }

        private List<ShapeSnapshot> snapshotShapes = new List<ShapeSnapshot>();
        private bool isInitializing = false;
        private bool isApplying = false;
        private bool isUpdatingControls = false;

        public SpacingForm(PowerPoint.Application app, List<int> selectedShapeIdsByOrder)
        {
            this.app = app;
            this.selectedShapeIdsByOrder = selectedShapeIdsByOrder;
            
            InitializeComponent();
            RefreshSelection();
        }

        private void InitializeComponent()
        {
            // Form properties
            this.Text = "对齐间距设置";
            this.ClientSize = new Size(304, 395); // Client size increased to fit stacked buttons
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.TopMost = true;
            this.BackColor = Color.FromArgb(240, 242, 245); // Fluent light gray
            this.Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;

            // 1. Info Label (top-most)
            lblInfo = new Label
            {
                Location = new Point(15, 15),
                AutoSize = true, // Auto-size to prevent clipping
                ForeColor = Color.FromArgb(64, 64, 64),
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold),
                Text = "正在读取当前选中的形状..."
            };
            this.Controls.Add(lblInfo);

            // 2. Uniform Distribute Card (taller to fit vertical stacking)
            cardDistribute = new Panel
            {
                Location = new Point(15, 53),
                Size = new Size(274, 90),
                BackColor = Color.White
            };
            cardDistribute.Paint += DrawCardBorder;
            this.Controls.Add(cardDistribute);

            btnDistributeH = new Button
            {
                Text = "水平均匀分布",
                Width = 254, // Full width of card minus padding
                Height = 32,
                Location = new Point(10, 10),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(243, 243, 243),
                ForeColor = Color.FromArgb(51, 51, 51),
                Cursor = Cursors.Hand,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular)
            };
            btnDistributeH.FlatAppearance.BorderSize = 1;
            btnDistributeH.FlatAppearance.BorderColor = Color.FromArgb(220, 220, 220);
            btnDistributeH.FlatAppearance.MouseOverBackColor = Color.FromArgb(225, 225, 225);
            btnDistributeH.Click += btnDistributeH_Click;
            cardDistribute.Controls.Add(btnDistributeH);

            btnDistributeV = new Button
            {
                Text = "垂直均匀分布",
                Width = 254, // Full width of card minus padding
                Height = 32,
                Location = new Point(10, 48), // Stacked below H button
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(243, 243, 243),
                ForeColor = Color.FromArgb(51, 51, 51),
                Cursor = Cursors.Hand,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular)
            };
            btnDistributeV.FlatAppearance.BorderSize = 1;
            btnDistributeV.FlatAppearance.BorderColor = Color.FromArgb(220, 220, 220);
            btnDistributeV.FlatAppearance.MouseOverBackColor = Color.FromArgb(225, 225, 225);
            btnDistributeV.Click += btnDistributeV_Click;
            cardDistribute.Controls.Add(btnDistributeV);

            // 3. Horizontal Spacing Card (pushed down)
            cardHorizontal = new Panel
            {
                Location = new Point(15, 153),
                Size = new Size(274, 85),
                BackColor = Color.White
            };
            cardHorizontal.Paint += DrawCardBorder;
            this.Controls.Add(cardHorizontal);

            lblHorizontal = new Label
            {
                Text = "水平间距 (厘米)",
                Location = new Point(12, 12),
                Size = new Size(150, 20),
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(205, 74, 38) // PowerPoint Accent color
            };
            cardHorizontal.Controls.Add(lblHorizontal);

            numHorizontal = new NumericUpDown
            {
                Minimum = -100,
                Maximum = 100,
                DecimalPlaces = 2,
                Value = 0.5m,
                Increment = 0.1m,
                Location = new Point(182, 10),
                Size = new Size(80, 23)
            };
            numHorizontal.ValueChanged += numHorizontal_ValueChanged;
            cardHorizontal.Controls.Add(numHorizontal);

            trackHorizontal = new TrackBar
            {
                Minimum = -50,
                Maximum = 150,
                Value = 5,
                TickStyle = TickStyle.None,
                Location = new Point(8, 42),
                Size = new Size(258, 30)
            };
            trackHorizontal.Scroll += trackHorizontal_Scroll;
            cardHorizontal.Controls.Add(trackHorizontal);

            // 4. Vertical Spacing Card (pushed down)
            cardVertical = new Panel
            {
                Location = new Point(15, 248),
                Size = new Size(274, 85),
                BackColor = Color.White
            };
            cardVertical.Paint += DrawCardBorder;
            this.Controls.Add(cardVertical);

            lblVertical = new Label
            {
                Text = "垂直间距 (厘米)",
                Location = new Point(12, 12),
                Size = new Size(150, 20),
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold),
                ForeColor = Color.FromArgb(205, 74, 38)
            };
            cardVertical.Controls.Add(lblVertical);

            numVertical = new NumericUpDown
            {
                Minimum = -100,
                Maximum = 100,
                DecimalPlaces = 2,
                Value = 0.5m,
                Increment = 0.1m,
                Location = new Point(182, 10),
                Size = new Size(80, 23)
            };
            numVertical.ValueChanged += numVertical_ValueChanged;
            cardVertical.Controls.Add(numVertical);

            trackVertical = new TrackBar
            {
                Minimum = -50,
                Maximum = 150,
                Value = 5,
                TickStyle = TickStyle.None,
                Location = new Point(8, 42),
                Size = new Size(258, 30)
            };
            trackVertical.Scroll += trackVertical_Scroll;
            cardVertical.Controls.Add(trackVertical);

            // 5. Close Button (centered at bottom, pushed down)
            btnClose = new Button
            {
                Text = "关闭",
                Width = 80,
                Height = 32,
                Location = new Point(112, 348), // Centered horizontally
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

            snapshotShapes.Clear();
            try
            {
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count < 2)
                {
                    lblInfo.Text = "请在 PPT 中选中至少两个形状。";
                    SetControlsEnabled(false);
                    return;
                }

                SetControlsEnabled(true);

                // Find baseline shape (first selected shape based on tracking list)
                PowerPoint.Shape baselineShape = null;
                if (selectedShapeIdsByOrder != null && selectedShapeIdsByOrder.Count > 0)
                {
                    int firstId = selectedShapeIdsByOrder[0];
                    foreach (PowerPoint.Shape shape in sel.ShapeRange)
                    {
                        if (shape.Id == firstId)
                        {
                            baselineShape = shape;
                            break;
                        }
                    }
                }

                if (baselineShape == null)
                {
                    baselineShape = sel.ShapeRange[1]; // Fallback to first in COM collection
                }

                lblInfo.Text = $"已选中 {sel.ShapeRange.Count} 个形状";

                foreach (PowerPoint.Shape shape in sel.ShapeRange)
                {
                    snapshotShapes.Add(new ShapeSnapshot
                    {
                        ActualShape = shape,
                        Id = shape.Id,
                        OriginalLeft = shape.Left,
                        OriginalTop = shape.Top,
                        Width = shape.Width,
                        Height = shape.Height,
                        IsBaseline = (shape.Id == baselineShape.Id)
                    });
                }

                // Enable/disable distribute buttons based on count (require >= 3)
                bool canDistribute = snapshotShapes.Count >= 3;
                btnDistributeH.Enabled = canDistribute;
                btnDistributeV.Enabled = canDistribute;

                // Calculate current average horizontal/vertical spacing as default values
                // For Horizontal spacing (sorted by Left)
                var sortedH = snapshotShapes.OrderBy(s => s.OriginalLeft).ToList();
                float totalHGap = 0f;
                for (int i = 0; i < sortedH.Count - 1; i++)
                {
                    totalHGap += (sortedH[i + 1].OriginalLeft - sortedH[i].OriginalLeft - sortedH[i].Width);
                }
                float avgHGapPoints = sortedH.Count > 1 ? totalHGap / (sortedH.Count - 1) : 0f;
                float avgHGapCm = avgHGapPoints / PointsPerCm;

                // For Vertical spacing (sorted by Top)
                var sortedV = snapshotShapes.OrderBy(s => s.OriginalTop).ToList();
                float totalVGap = 0f;
                for (int i = 0; i < sortedV.Count - 1; i++)
                {
                    totalVGap += (sortedV[i + 1].OriginalTop - sortedV[i].OriginalTop - sortedV[i].Height);
                }
                float avgVGapPoints = sortedV.Count > 1 ? totalVGap / (sortedV.Count - 1) : 0f;
                float avgVGapCm = avgVGapPoints / PointsPerCm;

                // Update UI values in centimeters without triggering ApplySpacing
                isInitializing = true;

                float hVal = Math.Max(-100f, Math.Min(100f, avgHGapCm));
                float vVal = Math.Max(-100f, Math.Min(100f, avgVGapCm));

                numHorizontal.Value = (decimal)hVal;
                numVertical.Value = (decimal)vVal;

                trackHorizontal.Value = (int)Math.Max(-50f, Math.Min(150f, hVal * 10f));
                trackVertical.Value = (int)Math.Max(-50f, Math.Min(150f, vVal * 10f));

                isInitializing = false;
            }
            catch (Exception ex)
            {
                lblInfo.Text = $"加载选区失败：{ex.Message}";
                SetControlsEnabled(false);
            }
        }

        private void SetControlsEnabled(bool enabled)
        {
            btnDistributeH.Enabled = enabled && snapshotShapes.Count >= 3;
            btnDistributeV.Enabled = enabled && snapshotShapes.Count >= 3;
            
            lblHorizontal.Enabled = enabled;
            numHorizontal.Enabled = enabled;
            trackHorizontal.Enabled = enabled;

            lblVertical.Enabled = enabled;
            numVertical.Enabled = enabled;
            trackVertical.Enabled = enabled;
        }

        private void ApplyHorizontalSpacing()
        {
            if (isInitializing || snapshotShapes == null || snapshotShapes.Count < 2) return;
            if (isApplying) return;

            isApplying = true;
            try
            {
                float hSpaceCm = (float)numHorizontal.Value;
                float hSpacePoints = hSpaceCm * PointsPerCm;

                var sortedH = snapshotShapes.OrderBy(s => s.OriginalLeft).ToList();
                int baseIndex = sortedH.FindIndex(s => s.IsBaseline);
                if (baseIndex >= 0)
                {
                    // Ensure baseline shape doesn't move horizontally
                    sortedH[baseIndex].ActualShape.Left = sortedH[baseIndex].OriginalLeft;

                    // Propagate right
                    for (int i = baseIndex + 1; i < sortedH.Count; i++)
                    {
                        sortedH[i].ActualShape.Left = sortedH[i - 1].ActualShape.Left + sortedH[i - 1].Width + hSpacePoints;
                    }

                    // Propagate left
                    for (int i = baseIndex - 1; i >= 0; i--)
                    {
                        sortedH[i].ActualShape.Left = sortedH[i + 1].ActualShape.Left - sortedH[i].Width - hSpacePoints;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying horizontal spacing: {ex.Message}");
            }
            finally
            {
                isApplying = false;
            }
        }

        private void ApplyVerticalSpacing()
        {
            if (isInitializing || snapshotShapes == null || snapshotShapes.Count < 2) return;
            if (isApplying) return;

            isApplying = true;
            try
            {
                float vSpaceCm = (float)numVertical.Value;
                float vSpacePoints = vSpaceCm * PointsPerCm;

                var sortedV = snapshotShapes.OrderBy(s => s.OriginalTop).ToList();
                int baseIndex = sortedV.FindIndex(s => s.IsBaseline);
                if (baseIndex >= 0)
                {
                    // Ensure baseline shape doesn't move vertically
                    sortedV[baseIndex].ActualShape.Top = sortedV[baseIndex].OriginalTop;

                    // Propagate down
                    for (int i = baseIndex + 1; i < sortedV.Count; i++)
                    {
                        sortedV[i].ActualShape.Top = sortedV[i - 1].ActualShape.Top + sortedV[i - 1].Height + vSpacePoints;
                    }

                    // Propagate up
                    for (int i = baseIndex - 1; i >= 0; i--)
                    {
                        sortedV[i].ActualShape.Top = sortedV[i + 1].ActualShape.Top - sortedV[i].Height - vSpacePoints;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error applying vertical spacing: {ex.Message}");
            }
            finally
            {
                isApplying = false;
            }
        }

        private void numHorizontal_ValueChanged(object sender, EventArgs e)
        {
            if (isUpdatingControls) return;
            isUpdatingControls = true;

            int val = (int)Math.Round(numHorizontal.Value * 10m);
            if (val >= trackHorizontal.Minimum && val <= trackHorizontal.Maximum)
            {
                trackHorizontal.Value = val;
            }

            isUpdatingControls = false;
            ApplyHorizontalSpacing();
        }

        private void trackHorizontal_Scroll(object sender, EventArgs e)
        {
            if (isUpdatingControls) return;
            isUpdatingControls = true;

            numHorizontal.Value = (decimal)trackHorizontal.Value / 10m;

            isUpdatingControls = false;
            ApplyHorizontalSpacing();
        }

        private void numVertical_ValueChanged(object sender, EventArgs e)
        {
            if (isUpdatingControls) return;
            isUpdatingControls = true;

            int val = (int)Math.Round(numVertical.Value * 10m);
            if (val >= trackVertical.Minimum && val <= trackVertical.Maximum)
            {
                trackVertical.Value = val;
            }

            isUpdatingControls = false;
            ApplyVerticalSpacing();
        }

        private void trackVertical_Scroll(object sender, EventArgs e)
        {
            if (isUpdatingControls) return;
            isUpdatingControls = true;

            numVertical.Value = (decimal)trackVertical.Value / 10m;

            isUpdatingControls = false;
            ApplyVerticalSpacing();
        }

        private void btnDistributeH_Click(object sender, EventArgs e)
        {
            if (snapshotShapes == null || snapshotShapes.Count < 3)
            {
                MessageBox.Show("请选中至少三个形状进行水平均匀分布。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // 1. Sort by Left
                var sortedH = snapshotShapes.OrderBy(s => s.OriginalLeft).ToList();
                int n = sortedH.Count;

                // 2. Calculate total width and sum of shape widths
                float totalWidth = sortedH[n - 1].OriginalLeft + sortedH[n - 1].Width - sortedH[0].OriginalLeft;
                float sumWidth = sortedH.Sum(s => s.Width);
                float totalGap = totalWidth - sumWidth;
                float gapPoints = totalGap / (n - 1);

                // 3. Apply positions
                for (int i = 1; i < n - 1; i++)
                {
                    sortedH[i].ActualShape.Left = sortedH[i - 1].ActualShape.Left + sortedH[i - 1].Width + gapPoints;
                }
                
                // The last shape stays at its original Left
                sortedH[n - 1].ActualShape.Left = sortedH[n - 1].OriginalLeft;

                // 4. Update the Horizontal NumericUpDown and TrackBar to reflect the new gap
                float gapCm = gapPoints / PointsPerCm;
                isInitializing = true;
                numHorizontal.Value = (decimal)Math.Max(-100f, Math.Min(100f, gapCm));
                trackHorizontal.Value = (int)Math.Max(-50f, Math.Min(150f, gapCm * 10f));
                isInitializing = false;

                // 5. Update our snapshot's OriginalLeft to match the new coordinates!
                foreach (var s in snapshotShapes)
                {
                    s.OriginalLeft = s.ActualShape.Left;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"水平均匀分布失败: {ex.Message}");
            }
        }

        private void btnDistributeV_Click(object sender, EventArgs e)
        {
            if (snapshotShapes == null || snapshotShapes.Count < 3)
            {
                MessageBox.Show("请选中至少三个形状进行垂直均匀分布。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // 1. Sort by Top
                var sortedV = snapshotShapes.OrderBy(s => s.OriginalTop).ToList();
                int n = sortedV.Count;

                // 2. Calculate total height and sum of shape heights
                float totalHeight = sortedV[n - 1].OriginalTop + sortedV[n - 1].Height - sortedV[0].OriginalTop;
                float sumHeight = sortedV.Sum(s => s.Height);
                float totalGap = totalHeight - sumHeight;
                float gapPoints = totalGap / (n - 1);

                // 3. Apply positions
                for (int i = 1; i < n - 1; i++)
                {
                    sortedV[i].ActualShape.Top = sortedV[i - 1].ActualShape.Top + sortedV[i - 1].Height + gapPoints;
                }

                // The last shape stays at its original Top
                sortedV[n - 1].ActualShape.Top = sortedV[n - 1].OriginalTop;

                // 4. Update the Vertical NumericUpDown and TrackBar to reflect the new gap
                float gapCm = gapPoints / PointsPerCm;
                isInitializing = true;
                numVertical.Value = (decimal)Math.Max(-100f, Math.Min(100f, gapCm));
                trackVertical.Value = (int)Math.Max(-50f, Math.Min(150f, gapCm * 10f));
                isInitializing = false;

                // 5. Update our snapshot's OriginalTop to match the new coordinates!
                foreach (var s in snapshotShapes)
                {
                    s.OriginalTop = s.ActualShape.Top;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"垂直均匀分布失败: {ex.Message}");
            }
        }
    }
}

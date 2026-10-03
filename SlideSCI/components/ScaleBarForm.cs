using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlideSCI
{
    /// <summary>查看和编辑 FOV 与比例尺，并在原图预览中即时显示效果。</summary>
    internal sealed class ScaleBarForm : Form
    {
        private readonly TableLayoutPanel fields;
        private readonly NumericUpDown fovWidth, fovHeight;
        private readonly ComboBox unit, direction, position;
        private readonly NumericUpDown barWidth, barHeight, thickness, fontSize;
        private readonly CheckBox showText, bold, group;
        private readonly Button barColor, fontColor;
        private readonly ImageFieldOfView initialFov;
        private readonly ScaleBarPreview preview;
        private readonly SizeF visibleFraction;
        private readonly Button confirm;
        private string barUnit;
        private bool updatingBarUnit;

        internal ImageFieldOfView FieldOfView => new ImageFieldOfView
        {
            Width = (double)fovWidth.Value, Height = (double)fovHeight.Value, Unit = unit.Text,
            Source = initialFov != null && (double)fovWidth.Value == initialFov.Width &&
                (double)fovHeight.Value == initialFov.Height && unit.Text == initialFov.Unit ? initialFov.Source : "手动设置"
        };
        internal ScaleBarSettings BarSettings => new ScaleBarSettings
        {
            Vertical = direction.SelectedIndex == 1,
            // 保存时统一换算为 μm，兼容已有图片和放大图的物理长度计算。
            WidthMicrometers = (double)barWidth.Value * ImageFieldOfView.UnitFactor(barUnit),
            HeightMicrometers = (double)barHeight.Value * ImageFieldOfView.UnitFactor(barUnit),
            LengthUnit = barUnit,
            ThicknessPoints = (float)thickness.Value, ShowText = showText.Checked,
            BarColorArgb = barColor.BackColor.ToArgb(), FontColorArgb = fontColor.BackColor.ToArgb(),
            FontSize = (float)fontSize.Value, Bold = bold.Checked,
            Corner = (ScaleBarCorner)position.SelectedIndex, GroupWithImage = group.Checked
        };

        internal ScaleBarForm(string imageName, ImageFieldOfView fov, ScaleBarSettings settings,
            Image previewImage, SizeF pictureSize, SizeF visibleFraction)
        {
            initialFov = fov;
            this.visibleFraction = visibleFraction;
            settings = settings ?? new ScaleBarSettings();
            Text = "比例尺 — " + imageName;
            Font = new Font("微软雅黑", 9);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = MinimizeBox = ShowIcon = ShowInTaskbar = false;
            var content = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, Dock = DockStyle.Fill, Padding = new Padding(10) };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 420));
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 460));
            Controls.Add(content);
            fields = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, Padding = new Padding(8), Dock = DockStyle.Fill };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 235));
            content.Controls.Add(fields, 0, 0);
            preview = new ScaleBarPreview(previewImage, pictureSize) { Dock = DockStyle.Fill };
            var previewGroup = new GroupBox { Text = "实时预览", Dock = DockStyle.Fill, Padding = new Padding(8), Margin = new Padding(10, 0, 0, 0) };
            previewGroup.Controls.Add(preview);
            content.Controls.Add(previewGroup, 1, 0);
            Add("", new Label { Text = "FOV 为未裁剪原图尺寸。\n横向填写宽度，纵向填写高度。", AutoSize = true });
            fovWidth = Numeric("FOV 宽度", fov?.Width ?? 0, 0, 1000000000);
            fovHeight = Numeric("FOV 高度", fov?.Height ?? 0, 0, 1000000000);
            unit = Choice("FOV 单位", new[] { "μm", "nm", "mm", "cm", "m", "in" }, 0);
            if (fov != null)
            {
                unit.SelectedItem = ImageFieldOfView.NormalizeUnit(fov.Unit);
            }
            Add("读取来源", new Label { Text = fov?.Source ?? "未找到标定，请手动填写 FOV。", AutoSize = true, MaximumSize = new Size(235, 0) });
            direction = Choice("比例尺方向", new[] { "横向", "纵向" }, settings.Vertical ? 1 : 0);
            barUnit = unit.Text;
            double unitFactor = ImageFieldOfView.UnitFactor(barUnit);
            barWidth = Numeric($"Width（{barUnit}）", settings.WidthMicrometers / unitFactor,
                0, 1000000000 / unitFactor);
            barHeight = Numeric($"Height（{barUnit}）", settings.HeightMicrometers / unitFactor,
                0, 1000000000 / unitFactor);
            Add("", new Label { Text = "Width / Height 为 0 时隐藏比例尺及文字。", AutoSize = true });
            thickness = Numeric("Thickness（pt）", settings.ThicknessPoints, 0.01, 1000);
            barColor = ColorButton("比例尺颜色", settings.BarColorArgb);
            position = Choice("比例尺位置", new[] { "左上", "左下", "右上", "右下" }, (int)settings.Corner);
            showText = Check($"显示比例尺文字（例如 50{barUnit}）", settings.ShowText);
            fontColor = ColorButton("字体颜色", settings.FontColorArgb);
            fontSize = Numeric("字体大小（pt）", settings.FontSize, 1, 400);
            bold = Check("文字加粗", settings.Bold);
            group = Check("比例尺与图片编组", settings.GroupWithImage);
            direction.SelectedIndexChanged += (sender, e) => UpdateEnabled();
            showText.CheckedChanged += (sender, e) => UpdateEnabled();
            UpdateEnabled();
            var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft };
            var cancel = new Button { Text = "取消", AutoSize = true, DialogResult = DialogResult.Cancel };
            confirm = new Button { Text = "确定", AutoSize = true };
            confirm.Click += (sender, e) =>
            {
                if (!BarSettings.IsValid || (BarSettings.LengthMicrometers != 0 && !FieldOfView.IsValidFor(BarSettings.Vertical)))
                { MessageBox.Show(this, "请填写大于 0 的 FOV 尺寸及比例尺参数：横向需要宽度，纵向需要高度。", Text); return; }
                DialogResult = DialogResult.OK;
            };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(confirm);
            Add("", buttons);
            AcceptButton = confirm;
            CancelButton = cancel;
            unit.SelectedIndexChanged += (sender, e) => UpdateBarUnit();
            foreach (Control control in fields.Controls)
            {
                if (control is NumericUpDown number) number.ValueChanged += (sender, e) => UpdatePreview();
                else if (control is ComboBox choice) choice.SelectedIndexChanged += (sender, e) => UpdatePreview();
                else if (control is CheckBox check) check.CheckedChanged += (sender, e) => UpdatePreview();
            }
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (!updatingBarUnit)
                confirm.Enabled = string.IsNullOrEmpty(preview.UpdateSettings(FieldOfView, visibleFraction, BarSettings));
        }

        private void UpdateBarUnit()
        {
            string nextUnit = unit.Text;
            double previousFactor = ImageFieldOfView.UnitFactor(barUnit);
            double nextFactor = ImageFieldOfView.UnitFactor(nextUnit);
            updatingBarUnit = true;
            try
            {
                foreach (NumericUpDown number in new[] { barWidth, barHeight })
                {
                    decimal converted = number.Value * (decimal)previousFactor / (decimal)nextFactor;
                    // 先放宽范围，再恢复换算后的边界，防止控件自动截断原值。
                    number.Minimum = 0;
                    number.Maximum = 1000000000000m;
                    number.Value = converted;
                    number.Maximum = 1000000000m / (decimal)nextFactor;
                    number.Minimum = 0;
                }
                barUnit = nextUnit;
                ((Label)fields.GetControlFromPosition(0, fields.GetRow(barWidth))).Text = $"Width（{barUnit}）";
                ((Label)fields.GetControlFromPosition(0, fields.GetRow(barHeight))).Text = $"Height（{barUnit}）";
                showText.Text = $"显示比例尺文字（例如 50{barUnit}）";
            }
            finally { updatingBarUnit = false; }
            UpdatePreview();
        }

        private void UpdateEnabled()
        {
            barWidth.Enabled = direction.SelectedIndex == 0;
            barHeight.Enabled = direction.SelectedIndex == 1;
            fovWidth.Enabled = barWidth.Enabled;
            fovHeight.Enabled = barHeight.Enabled;
            fontColor.Enabled = fontSize.Enabled = bold.Enabled = showText.Checked;
        }
        private void Add(string label, Control control, bool stretch = true)
        {
            int row = fields.RowCount++;
            fields.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            fields.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row);
            control.Anchor = stretch ? AnchorStyles.Left | AnchorStyles.Right : AnchorStyles.Left;
            control.Margin = new Padding(0, 5, 0, 5);
            fields.Controls.Add(control, 1, row);
        }
        private NumericUpDown Numeric(string label, double value, double minimum, double maximum)
        {
            var control = new LiveNumericUpDown { Minimum = (decimal)minimum, Maximum = (decimal)maximum, DecimalPlaces = 2 };
            control.Value = (decimal)Math.Max(minimum, Math.Min(maximum, value));
            Add(label, control);
            return control;
        }
        private ComboBox Choice(string label, string[] choices, int index)
        {
            var control = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
            control.Items.AddRange(choices);
            control.SelectedIndex = index;
            Add(label, control);
            return control;
        }
        private CheckBox Check(string label, bool value)
        {
            var control = new CheckBox { Text = label, Checked = value, AutoSize = true };
            Add("", control);
            return control;
        }
        private Button ColorButton(string label, int argb)
        {
            var button = new Button
            {
                Size = new Size(44, 24), BackColor = Color.FromArgb(argb), UseVisualStyleBackColor = false,
                FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
                AccessibleName = label, AccessibleDescription = "点击选择颜色"
            };
            button.FlatAppearance.BorderColor = Color.Gray;
            button.Click += (sender, e) =>
            {
                using (var dialog = new ColorDialog { Color = button.BackColor, FullOpen = true })
                    if (dialog.ShowDialog(this) == DialogResult.OK) { button.BackColor = dialog.Color; UpdatePreview(); }
            };
            Add(label, button, false);
            return button;
        }
    }
}

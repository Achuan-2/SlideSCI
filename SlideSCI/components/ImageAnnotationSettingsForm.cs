using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace SlideSCI
{
    /// <summary>图片标题和标签设置共用的弹窗布局与输入检查。</summary>
    internal abstract class ImageAnnotationSettingsForm : Form
    {
        private readonly TableLayoutPanel fields;
        protected readonly ComboBox fontName;
        protected readonly NumericUpDown fontSize;

        protected ImageAnnotationSettingsForm(string title, IEnumerable<string> fonts,
            string currentFontName, string currentFontSize)
        {
            Text = title;
            Font = new Font("微软雅黑", 9F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;

            fields = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                Dock = DockStyle.Fill,
                Padding = new Padding(16),
                Width = 420
            };
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            fields.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 270));
            Controls.Add(fields);

            fontName = AddComboBox("字体", currentFontName, fonts);
            fontSize = AddNumericBox("字号", currentFontSize, 14m, 0.01m);
        }

        protected ComboBox AddComboBox(string label, string value, IEnumerable<string> choices,
            bool allowCustomValue = true)
        {
            var control = new ComboBox
            {
                DropDownStyle = allowCustomValue ? ComboBoxStyle.DropDown : ComboBoxStyle.DropDownList,
                AutoCompleteMode = allowCustomValue ? AutoCompleteMode.SuggestAppend : AutoCompleteMode.None,
                AutoCompleteSource = AutoCompleteSource.ListItems,
                MaxDropDownItems = 16,
                IntegralHeight = false
            };
            foreach (string choice in choices) control.Items.Add(choice);
            if (allowCustomValue) control.Text = value;
            else
            {
                control.SelectedItem = value;
                if (control.SelectedIndex < 0) control.SelectedIndex = 0;
            }
            AddField(label, control);
            return control;
        }

        protected NumericUpDown AddNumericBox(string label, string value, decimal defaultValue = 0m,
            decimal minimum = int.MinValue)
        {
            var control = new NumericUpDown
            {
                Minimum = minimum,
                Maximum = int.MaxValue,
                DecimalPlaces = 2,
                Increment = 1m
            };
            decimal currentValue;
            if (!decimal.TryParse(value, out currentValue)) currentValue = defaultValue;
            control.Value = Math.Max(control.Minimum, Math.Min(control.Maximum, currentValue));
            AddField(label, control);
            return control;
        }

        protected TextBox AddTextBox(string label, string value)
        {
            var control = new TextBox { Text = value };
            AddField(label, control);
            return control;
        }

        protected CheckBox AddCheckBox(string text, bool value)
        {
            var control = new CheckBox { Text = text, Checked = value, AutoSize = true };
            AddField("", control);
            return control;
        }

        private void AddField(string label, Control control)
        {
            int row = fields.RowCount++;
            fields.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            fields.Controls.Add(new Label
            {
                Text = label,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 8, 8, 8)
            }, 0, row);
            control.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            control.Margin = new Padding(0, 5, 0, 5);
            control.TabIndex = row;
            fields.Controls.Add(control, 1, row);
        }

        protected void AddDialogButtons()
        {
            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 12, 0, 0)
            };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };
            var confirm = new Button { Text = "确定", AutoSize = true };
            confirm.Click += (sender, e) =>
            {
                if (ValidateSettings()) DialogResult = DialogResult.OK;
            };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(confirm);
            int row = fields.RowCount++;
            fields.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            fields.Controls.Add(buttons, 0, row);
            fields.SetColumnSpan(buttons, 2);
            AcceptButton = confirm;
            CancelButton = cancel;
        }

        protected virtual bool ValidateSettings()
        {
            if (string.IsNullOrWhiteSpace(fontName.Text))
                return ShowInputError(fontName, "请输入字体名称。");
            return ValidateNumber(fontSize, "字号", true);
        }

        protected bool ValidateNumber(Control control, string label, bool positive = false)
        {
            float value;
            if (!float.TryParse(control.Text, out value) || float.IsNaN(value) ||
                float.IsInfinity(value) || (positive && value <= 0))
                return ShowInputError(control, positive ? $"{label}必须是大于 0 的数字。" : $"请输入有效的{label}。");
            return true;
        }

        protected bool ShowInputError(Control control, string message)
        {
            MessageBox.Show(this, message, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            control.Focus();
            return false;
        }
    }

    internal sealed class ImageTitleSettingsForm : ImageAnnotationSettingsForm
    {
        private readonly NumericUpDown distance;
        private readonly TextBox titleText;
        private readonly CheckBox autoGroup;
        private readonly CheckBox center;

        internal ImageTitleSettingsForm(IEnumerable<string> fonts)
            : base("图片标题设置", fonts, Properties.Settings.Default.TitleFontName,
                Properties.Settings.Default.TitleFontSize)
        {
            var settings = Properties.Settings.Default;
            distance = AddNumericBox("图距 (pt)", settings.TitleDistanceFromBottom);
            titleText = AddTextBox("标题占位", settings.TitleText);
            autoGroup = AddCheckBox("自动编组图片与标题", settings.AutoGroup);
            center = AddCheckBox("标题居中", settings.imgAddTitleCenter);
            AddDialogButtons();
        }

        protected override bool ValidateSettings()
        {
            return base.ValidateSettings() && ValidateNumber(distance, "图片与标题距离");
        }

        internal void SaveSettings()
        {
            var settings = Properties.Settings.Default;
            settings.TitleFontName = fontName.Text.Trim();
            settings.TitleFontSize = fontSize.Value.ToString("0.##");
            settings.TitleDistanceFromBottom = distance.Value.ToString("0.##");
            settings.TitleText = titleText.Text;
            settings.AutoGroup = autoGroup.Checked;
            settings.imgAddTitleCenter = center.Checked;
            settings.Save();
        }
    }

    internal sealed class ImageLabelSettingsForm : ImageAnnotationSettingsForm
    {
        private readonly ComboBox template;
        private readonly NumericUpDown offsetX;
        private readonly NumericUpDown offsetY;
        private readonly CheckBox bold;

        internal ImageLabelSettingsForm(IEnumerable<string> fonts)
            : base("图片标签设置", fonts, Properties.Settings.Default.LabelFontName,
                Properties.Settings.Default.LabelFontSize)
        {
            var settings = Properties.Settings.Default;
            template = AddComboBox("标签模板", settings.LabelTemplate, new[]
            {
                "A", "a", "A)", "a)", "(A)", "(a)", "1", "1)", "Ⅰ", "Ⅰ)", "①)", "①", "一)", "一"
            }, false);
            offsetX = AddNumericBox("X 偏移 (pt)", settings.LabelOffsetX, -20m);
            offsetY = AddNumericBox("Y 偏移 (pt)", settings.LabelOffsetY, -7m);
            bold = AddCheckBox("加粗", settings.LabelBold);
            AddDialogButtons();
        }

        protected override bool ValidateSettings()
        {
            return base.ValidateSettings() && ValidateNumber(offsetX, "X 偏移") &&
                ValidateNumber(offsetY, "Y 偏移");
        }

        internal void SaveSettings()
        {
            var settings = Properties.Settings.Default;
            settings.LabelFontName = fontName.Text.Trim();
            settings.LabelFontSize = fontSize.Value.ToString("0.##");
            settings.LabelTemplate = template.Text;
            settings.LabelOffsetX = offsetX.Value.ToString("0.##");
            settings.LabelOffsetY = offsetY.Value.ToString("0.##");
            settings.LabelBold = bold.Checked;
            settings.Save();
        }
    }
}

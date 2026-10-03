using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace SlideSCI
{
    internal sealed class PseudoColorForm : Form
    {
        private readonly IList<Bitmap> originals;
        private readonly Color[] colors;
        private readonly int[] sourcePeaks;
        private readonly ComboBox currentPicture;
        private readonly ComboBox colorChoice;
        private readonly Button customColor;
        private readonly PictureBox preview;
        private readonly CheckBox keepOriginal;
        private bool loading;

        internal Color[] SelectedColors => (Color[])colors.Clone();
        internal bool KeepOriginal => keepOriginal.Checked;

        // 预览图由调用方持有，弹窗仅负责释放生成的伪彩预览。
        internal PseudoColorForm(IList<string> names, IList<Bitmap> images, Color[] initialColors, int[] imagePeaks)
        {
            originals = images;
            colors = (Color[])initialColors.Clone();
            sourcePeaks = (int[])imagePeaks.Clone();
            Text = "伪彩设置";
            Font = new Font("微软雅黑", 9F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(700, 570);
            MinimumSize = new Size(620, 570);
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowIcon = false;
            ShowInTaskbar = false;

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6,
                Padding = new Padding(14) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 44));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            Controls.Add(layout);
            layout.Controls.Add(new Label { Text = $"已选择 {images.Count} 张图片，可逐张设色或统一应用。",
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft }, 0, 0);
            currentPicture = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill };
            for (int i = 0; i < names.Count; i++) currentPicture.Items.Add($"{i + 1}. {names[i]}");
            layout.Controls.Add(currentPicture, 0, 1);

            var colorRow = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            colorRow.Controls.Add(new Label { Text = "颜色", AutoSize = true, Margin = new Padding(0, 7, 10, 0) });
            colorChoice = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120 };
            colorChoice.Items.AddRange(ChannelImageProcessor.ColorNames);
            colorChoice.Items.Add("自定义");
            colorRow.Controls.Add(colorChoice);
            customColor = new Button { Text = "自定义颜色…", AutoSize = true };
            colorRow.Controls.Add(customColor);
            var applyAll = new Button { Text = "应用颜色到全部图片", AutoSize = true };
            colorRow.Controls.Add(applyAll);
            layout.Controls.Add(colorRow, 0, 2);
            preview = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.FromArgb(32, 32, 32) };
            layout.Controls.Add(preview, 0, 3);
            keepOriginal = new CheckBox { Text = "保留原图，在右侧生成伪彩图", Dock = DockStyle.Fill,
                Checked = Properties.Settings.Default.PseudoColorKeepOriginal };
            layout.Controls.Add(keepOriginal, 0, 4);
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft };
            var cancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, AutoSize = true };
            var confirm = new Button { Text = "确定", DialogResult = DialogResult.OK, AutoSize = true };
            actions.Controls.Add(cancel);
            actions.Controls.Add(confirm);
            layout.Controls.Add(actions, 0, 5);
            AcceptButton = confirm;
            CancelButton = cancel;

            currentPicture.SelectedIndexChanged += (sender, args) => LoadColor();
            colorChoice.SelectedIndexChanged += (sender, args) =>
            {
                if (loading || currentPicture.SelectedIndex < 0 || colorChoice.SelectedIndex < 0) return;
                if (colorChoice.SelectedIndex < ChannelImageProcessor.DefaultColors.Length)
                {
                    colors[currentPicture.SelectedIndex] = ChannelImageProcessor.DefaultColors[colorChoice.SelectedIndex];
                    UpdatePreview();
                }
                else ChooseCustomColor();
            };
            customColor.Click += (sender, args) => ChooseCustomColor();
            applyAll.Click += (sender, args) =>
            {
                Color color = colors[currentPicture.SelectedIndex];
                for (int i = 0; i < colors.Length; i++) colors[i] = color;
                LoadColor();
            };
            currentPicture.SelectedIndex = 0;
        }

        private void ChooseCustomColor()
        {
            using (var dialog = new ColorDialog { Color = colors[currentPicture.SelectedIndex], FullOpen = true })
                if (dialog.ShowDialog(this) == DialogResult.OK) colors[currentPicture.SelectedIndex] = dialog.Color;
            LoadColor();
        }

        private void LoadColor()
        {
            loading = true;
            try
            {
                int index = Array.FindIndex(ChannelImageProcessor.DefaultColors,
                    color => color.ToArgb() == colors[currentPicture.SelectedIndex].ToArgb());
                colorChoice.SelectedIndex = index < 0 ? ChannelImageProcessor.DefaultColors.Length : index;
            }
            finally { loading = false; }
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            Color color = colors[currentPicture.SelectedIndex];
            Image next = ChannelImageProcessor.ApplyColor(originals[currentPicture.SelectedIndex], color,
                sourcePeaks[currentPicture.SelectedIndex]);
            Image previous = preview.Image;
            preview.Image = next;
            previous?.Dispose();
            customColor.BackColor = color;
            customColor.ForeColor = color.GetBrightness() > 0.5F ? Color.Black : Color.White;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try
            {
                // 与放大图弹窗一致：记忆关闭时的选项，取消只撤销本次图片修改。
                var settings = Properties.Settings.Default;
                settings.PseudoColorArgb = colors[currentPicture.SelectedIndex].ToArgb().ToString(CultureInfo.InvariantCulture);
                settings.PseudoColorKeepOriginal = keepOriginal.Checked;
                settings.Save();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"保存伪彩设置时出错：{ex.Message}");
                MessageBox.Show("伪彩设置未能保存，下次打开将使用上次保存的设置。", "提示");
            }
            base.OnFormClosed(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && preview != null)
            {
                Image previous = preview.Image;
                preview.Image = null;
                previous?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}

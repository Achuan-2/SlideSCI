using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    /// <summary>图片框选区域及保留在原图上的矩形描边。</summary>
    public sealed class ZoomImageSelection
    {
        public RectangleF Region { get; }
        public Color OutlineColor { get; }
        public float OutlineWidthPoints { get; }
        public DashStyle OutlineDashStyle { get; }
        public bool UseRectangleColorForZoomImage { get; }
        public bool AddGuideLines { get; }
        public ZoomImagePlacement Placement { get; }
        public ZoomGuideLineExtent GuideLineExtent { get; }
        public float RegionRotationDegrees { get; }
        public bool RegionWasEdited { get; }
        public Office.MsoLineDashStyle? NativeOutlineDashStyle { get; }

        public ZoomImageSelection(RectangleF region, Color color, float widthPoints, DashStyle dashStyle,
            bool useRectangleColorForZoomImage, bool addGuideLines, ZoomImagePlacement placement,
            float regionRotationDegrees = 0, bool regionWasEdited = false,
            Office.MsoLineDashStyle? nativeOutlineDashStyle = null,
            ZoomGuideLineExtent guideLineExtent = ZoomGuideLineExtent.AcrossImages)
        {
            Region = region;
            OutlineColor = color;
            OutlineWidthPoints = widthPoints;
            OutlineDashStyle = dashStyle;
            UseRectangleColorForZoomImage = useRectangleColorForZoomImage;
            AddGuideLines = addGuideLines;
            Placement = placement;
            GuideLineExtent = guideLineExtent;
            RegionRotationDegrees = regionRotationDegrees;
            RegionWasEdited = regionWasEdited;
            NativeOutlineDashStyle = nativeOutlineDashStyle;
        }
    }

    /// <summary>在图片预览中框选区域，返回相对于图片的归一化矩形。</summary>
    public sealed class ZoomImageForm : Form
    {
        private readonly ZoomRegionCanvas canvas;
        private readonly Button confirmButton;
        private readonly CheckBox useRectangleColorCheckBox;
        private readonly CheckBox addGuideLinesCheckBox;
        private readonly ComboBox placementComboBox;
        private readonly ComboBox guideLineExtentComboBox;
        private readonly LayoutPreviewCanvas layoutPreview;
        private readonly List<ZoomImageEntry> entries = new List<ZoomImageEntry>();
        private readonly ComboBox entrySelector;
        private readonly Button addEntryButton;
        private readonly Control settingsGroups;
        private readonly Font settingsTitleFont;
        private readonly bool hadExistingEntries;
        private ZoomImageEntry activeEntry;
        private NumericUpDown lineWidth;
        private Button colorButton;
        private ComboBox lineStyle;
        private bool loadingEntry = true;
        private int nextEntryNumber;

        public IList<ZoomImageEntry> SelectedEntries
        {
            get { SaveCurrentEntry(); return entries.ToArray(); }
        }

        public ZoomImageSelection SelectedOptions => new ZoomImageSelection(
            canvas.SelectedRegion, canvas.OutlineColor, canvas.OutlineWidthPoints, canvas.OutlineDashStyle,
            useRectangleColorCheckBox.Checked, addGuideLinesCheckBox.Checked,
            placementComboBox.SelectedIndex >= 0 ? (ZoomImagePlacement)placementComboBox.SelectedIndex : ZoomImagePlacement.Right,
            canvas.RegionRotationDegrees, canvas.RegionWasEdited, canvas.NativeOutlineDashStyle,
            guideLineExtentComboBox.SelectedIndex >= 0
                ? (ZoomGuideLineExtent)guideLineExtentComboBox.SelectedIndex : ZoomGuideLineExtent.AcrossImages);

        // 图片由调用者持有，关闭弹窗时不释放图片。
        public ZoomImageForm(Image preview, float pictureWidthPoints, ZoomImageSelection initialSelection = null)
            : this(preview, pictureWidthPoints, initialSelection == null ? null : new[]
                { new ZoomImageEntry(null, "放大图 1", initialSelection) }, 0)
        {
        }

        public ZoomImageForm(Image preview, float pictureWidthPoints, IList<ZoomImageEntry> initialEntries, int selectedIndex)
        {
            if (initialEntries != null) entries.AddRange(initialEntries);
            hadExistingEntries = entries.Any(entry => entry.RecordKey != null);
            nextEntryNumber = Math.Max(entries.Count + 1, entries.Where(entry => entry.DisplayOrder < int.MaxValue)
                .Select(entry => entry.DisplayOrder + 1).DefaultIfEmpty(1).Max());
            Text = "制作放大图";
            ClientSize = new Size(1200, 900);
            MinimumSize = new Size(600, 540);
            StartPosition = FormStartPosition.CenterParent;
            ShowIcon = false;
            ShowInTaskbar = false;
            MinimizeBox = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            settingsTitleFont = new Font(Font, FontStyle.Bold);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                ColumnCount = 1,
                RowCount = 4
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 114));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 46));
            canvas = new ZoomRegionCanvas(preview, pictureWidthPoints) { Dock = DockStyle.Fill };
            canvas.Entries = entries;
            LoadOutlineSettings();
            if (entries.Count > 0)
            {
                ZoomImageSelection initialSelection = entries[Math.Max(0, Math.Min(entries.Count - 1, selectedIndex))].Options;
                canvas.OutlineColor = initialSelection.OutlineColor;
                canvas.OutlineWidthPoints = initialSelection.OutlineWidthPoints;
                canvas.OutlineDashStyle = initialSelection.OutlineDashStyle;
                canvas.NativeOutlineDashStyle = initialSelection.NativeOutlineDashStyle;
                canvas.SetInitialRegion(initialSelection.Region, initialSelection.RegionRotationDegrees);
            }
            layoutPreview = new LayoutPreviewCanvas(preview, pictureWidthPoints, () => entries, () => activeEntry)
            {
                Dock = DockStyle.Fill
            };
            useRectangleColorCheckBox = new CheckBox
            {
                Text = "使用矩形描边颜色",
                Checked = Properties.Settings.Default.ZoomUseRectangleColor,
                AutoSize = true,
                Anchor = AnchorStyles.Left
            };
            addGuideLinesCheckBox = new CheckBox
            {
                Text = "添加辅助线", Checked = Properties.Settings.Default.ZoomAddGuideLines,
                AutoSize = true, Anchor = AnchorStyles.Left
            };
            placementComboBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 80 };
            placementComboBox.Items.AddRange(new object[] { "右侧", "左侧", "上侧", "下侧" });
            placementComboBox.SelectedIndex = (int)ZoomImageGeometry.GetSavedPlacement();
            useRectangleColorCheckBox.CheckedChanged += (sender, e) => RefreshPreviews();
            placementComboBox.SelectedIndexChanged += (sender, e) => RefreshPreviews();
            var zoomSettings = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            zoomSettings.Controls.Add(new Label
            {
                Text = "放置位置", AutoSize = true, Margin = new Padding(0, 7, 4, 0)
            });
            zoomSettings.Controls.Add(placementComboBox);
            useRectangleColorCheckBox.Margin = new Padding(14, 5, 3, 0);
            zoomSettings.Controls.Add(useRectangleColorCheckBox);

            guideLineExtentComboBox = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList, Width = 120,
                Enabled = addGuideLinesCheckBox.Checked
            };
            guideLineExtentComboBox.Items.AddRange(new object[] { "仅在原图内", "可超出图片" });
            guideLineExtentComboBox.SelectedIndex = (int)ZoomImageGeometry.GetSavedGuideLineExtent();
            guideLineExtentComboBox.SelectedIndexChanged += (sender, e) => RefreshPreviews();
            addGuideLinesCheckBox.CheckedChanged += (sender, e) =>
            {
                guideLineExtentComboBox.Enabled = addGuideLinesCheckBox.Checked;
                RefreshPreviews();
            };
            var guideSettings = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            addGuideLinesCheckBox.Margin = new Padding(0, 5, 14, 0);
            guideSettings.Controls.Add(addGuideLinesCheckBox);
            guideSettings.Controls.Add(new Label
            {
                Text = "绘制范围", AutoSize = true, Margin = new Padding(0, 7, 4, 0)
            });
            guideSettings.Controls.Add(guideLineExtentComboBox);

            var groupedSettings = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 3,
                Margin = Padding.Empty, Padding = new Padding(0, 0, 0, 6)
            };
            groupedSettings.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            groupedSettings.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 3; i++) groupedSettings.RowStyles.Add(new RowStyle(SizeType.Percent, 100f / 3));
            AddSettingsRow(groupedSettings, 0, "放大矩形设置", CreateOutlineSettings());
            AddSettingsRow(groupedSettings, 1, "放大图设置", zoomSettings);
            AddSettingsRow(groupedSettings, 2, "辅助线设置", guideSettings);
            settingsGroups = groupedSettings;
            layout.Controls.Add(settingsGroups, 0, 1);

            entrySelector = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 170 };
            entrySelector.SelectedIndexChanged += (sender, e) =>
            {
                if (!loadingEntry) SelectEntry(entrySelector.SelectedItem as ZoomImageEntry);
            };
            addEntryButton = new Button { Text = "添加新放大图", AutoSize = true };
            addEntryButton.Click += (sender, e) => AddEntry();
            var entryControls = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            entryControls.Controls.Add(new Label { Text = "当前放大图", AutoSize = true, Margin = new Padding(0, 7, 6, 0) });
            entryControls.Controls.Add(entrySelector);
            entryControls.Controls.Add(addEntryButton);
            layout.Controls.Add(entryControls, 0, 0);

            var zoomLabel = new Label
            {
                Text = "缩放：100%", AutoSize = true, Margin = new Padding(0, 7, 8, 0)
            };
            canvas.ViewChanged += (sender, e) => zoomLabel.Text = $"缩放：{canvas.ZoomPercent:0}%";
            var fitButton = new Button { Text = "适应窗口", AutoSize = true };
            fitButton.Click += (sender, e) => canvas.FitImage();
            var panButton = new CheckBox
            {
                Text = "抓手", Appearance = Appearance.Button, AutoSize = true,
                TextAlign = ContentAlignment.MiddleCenter
            };
            panButton.CheckedChanged += (sender, e) =>
            {
                canvas.PanToolActive = panButton.Checked;
                canvas.Focus();
            };
            var panHint = new Label
            {
                Text = "按住空格+左键可以拖动画布", AutoSize = true,
                ForeColor = Color.DimGray, Margin = new Padding(3, 2, 0, 4)
            };
            var viewControls = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 3, RowCount = 2, Margin = Padding.Empty,
                Padding = new Padding(6, 2, 0, 0)
            };
            viewControls.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            viewControls.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            viewControls.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            viewControls.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            viewControls.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            viewControls.Controls.Add(zoomLabel, 0, 0);
            viewControls.Controls.Add(fitButton, 1, 0);
            viewControls.Controls.Add(panButton, 2, 0);
            viewControls.Controls.Add(panHint, 2, 1);

            var previews = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2 };
            previews.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60));
            previews.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40));
            previews.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            previews.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            previews.Controls.Add(canvas, 0, 0);
            previews.Controls.Add(viewControls, 0, 1);
            previews.Controls.Add(layoutPreview, 1, 0);
            previews.SetRowSpan(layoutPreview, 2);
            layout.Controls.Add(previews, 0, 2);

            confirmButton = new Button { Text = "确定", Enabled = false, AutoSize = true };
            confirmButton.Click += (sender, e) =>
            {
                SaveCurrentEntry();
                if (CanConfirm())
                {
                    DialogResult = DialogResult.OK;
                }
            };
            canvas.SelectionChanged += (sender, e) =>
            {
                if (loadingEntry) return;
                SaveCurrentEntry();
                confirmButton.Enabled = CanConfirm();
                addEntryButton.Enabled = !canvas.IsEditing;
                layoutPreview.Invalidate();
            };
            canvas.EntryPicked += entry => SelectEntry(entry);
            canvas.DeleteEntryRequested += (sender, e) => DeleteCurrentEntry();
            var cancelButton = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                AutoSize = true
            };
            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 10, 0, 0)
            };
            buttons.Controls.Add(cancelButton);
            buttons.Controls.Add(confirmButton);
            layout.Controls.Add(buttons, 0, 3);
            Controls.Add(layout);
            AcceptButton = confirmButton;
            CancelButton = cancelButton;
            loadingEntry = false;
            if (entries.Count == 0) AddEntry();
            else SelectEntry(entries[Math.Max(0, Math.Min(entries.Count - 1, selectedIndex))]);
        }

        private bool CanConfirm() => !canvas.IsEditing && entries.All(entry => entry.HasRegion) &&
            (entries.Count > 0 || hadExistingEntries);

        private void SaveCurrentEntry()
        {
            if (!loadingEntry && activeEntry != null) activeEntry.Options = SelectedOptions;
        }

        private void AddEntry()
        {
            SaveCurrentEntry();
            ZoomImageEntry pending = entries.FirstOrDefault(entry => !entry.HasRegion);
            if (pending != null)
            {
                pending.Options = ZoomImageEntry.WithRegion(SelectedOptions, RectangleF.Empty);
                SelectEntry(pending);
                canvas.Focus();
                return;
            }
            var newEntry = new ZoomImageEntry(null, $"放大图 {nextEntryNumber++}",
                ZoomImageEntry.WithRegion(SelectedOptions, RectangleF.Empty));
            entries.Add(newEntry);
            SelectEntry(newEntry);
            canvas.Focus();
        }

        private void DeleteCurrentEntry()
        {
            if (activeEntry == null) return;
            int index = entries.IndexOf(activeEntry);
            entries.Remove(activeEntry);
            activeEntry = null;
            SelectEntry(entries.Count == 0 ? null : entries[Math.Min(index, entries.Count - 1)]);
        }

        private void SelectEntry(ZoomImageEntry entry)
        {
            if (canvas.IsEditing) return;
            SaveCurrentEntry();
            loadingEntry = true;
            try
            {
                activeEntry = entry;
                canvas.ActiveEntry = entry;
                canvas.CanDrawRegion = entry != null && !entry.HasRegion;
                entrySelector.Items.Clear();
                foreach (ZoomImageEntry item in entries) entrySelector.Items.Add(item);
                entrySelector.SelectedItem = entry;
                settingsGroups.Enabled = entry != null;
                if (entry == null) canvas.SetInitialRegion(RectangleF.Empty, 0);
                else
                {
                    ZoomImageSelection options = entry.Options;
                    canvas.OutlineColor = options.OutlineColor;
                    canvas.OutlineWidthPoints = options.OutlineWidthPoints;
                    canvas.OutlineDashStyle = options.OutlineDashStyle;
                    canvas.NativeOutlineDashStyle = options.NativeOutlineDashStyle;
                    canvas.SetInitialRegion(options.Region, options.RegionRotationDegrees, options.RegionWasEdited);
                    lineWidth.Maximum = Math.Max(20M, (decimal)options.OutlineWidthPoints);
                    lineWidth.Value = (decimal)options.OutlineWidthPoints;
                    colorButton.BackColor = options.OutlineColor;
                    colorButton.ForeColor = options.OutlineColor.GetBrightness() > 0.5f ? Color.Black : Color.White;
                    ResetLineStyleChoices();
                    useRectangleColorCheckBox.Checked = options.UseRectangleColorForZoomImage;
                    addGuideLinesCheckBox.Checked = options.AddGuideLines;
                    placementComboBox.SelectedIndex = (int)options.Placement;
                    guideLineExtentComboBox.SelectedIndex = (int)options.GuideLineExtent;
                    guideLineExtentComboBox.Enabled = options.AddGuideLines;
                }
            }
            finally { loadingEntry = false; }
            confirmButton.Enabled = CanConfirm();
            canvas.Invalidate();
            layoutPreview.Invalidate();
        }

        private void LoadOutlineSettings()
        {
            var settings = Properties.Settings.Default;
            double width = settings.ZoomOutlineWidthPoints;
            canvas.OutlineWidthPoints = double.IsNaN(width) || double.IsInfinity(width)
                ? 1.5f : (float)Math.Max(0.01, Math.Min(1000, width));
            canvas.OutlineColor = Color.FromArgb(255, Color.FromArgb(settings.ZoomOutlineColorArgb));
            int dash = settings.ZoomOutlineDashStyle;
            canvas.OutlineDashStyle = dash >= (int)DashStyle.Solid && dash <= (int)DashStyle.DashDotDot
                ? (DashStyle)dash : DashStyle.Dash;
            int nativeDash = settings.ZoomOutlineNativeDashStyle;
            if (Enum.IsDefined(typeof(Office.MsoLineDashStyle), nativeDash) &&
                nativeDash != (int)Office.MsoLineDashStyle.msoLineDashStyleMixed)
            {
                canvas.NativeOutlineDashStyle = (Office.MsoLineDashStyle)nativeDash;
                canvas.OutlineDashStyle = ZoomImageGeometry.GetDrawingDashStyle(canvas.NativeOutlineDashStyle.Value);
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            try
            {
                // 只记忆操作设置，框选范围属于当前图片，下次重新框选。
                var settings = Properties.Settings.Default;
                ZoomImageSelection options = SelectedOptions;
                settings.ZoomOutlineWidthPoints = options.OutlineWidthPoints;
                settings.ZoomOutlineColorArgb = options.OutlineColor.ToArgb();
                settings.ZoomOutlineDashStyle = (int)options.OutlineDashStyle;
                settings.ZoomOutlineNativeDashStyle = options.NativeOutlineDashStyle.HasValue
                    ? (int)options.NativeOutlineDashStyle.Value : -1;
                settings.ZoomUseRectangleColor = options.UseRectangleColorForZoomImage;
                settings.ZoomAddGuideLines = options.AddGuideLines;
                settings.ZoomImagePlacement = (int)options.Placement;
                settings.ZoomGuideLineExtent = (int)options.GuideLineExtent;
                settings.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"保存放大图设置时出错: {ex.Message}");
                MessageBox.Show("放大图设置未能保存，下次打开将使用上次保存的设置。", "提示");
            }
            base.OnFormClosed(e);
        }

        private void RefreshPreviews()
        {
            if (loadingEntry) return;
            SaveCurrentEntry();
            canvas.Invalidate();
            layoutPreview?.Invalidate();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) settingsTitleFont?.Dispose();
            base.Dispose(disposing);
        }

        private void AddSettingsRow(TableLayoutPanel table, int row, string title, Control content)
        {
            table.Controls.Add(new Label
            {
                Text = title + "：", Font = settingsTitleFont, AutoSize = true,
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, 8, 0)
            }, 0, row);
            content.Margin = Padding.Empty;
            table.Controls.Add(content, 1, row);
        }

        private Control CreateOutlineSettings()
        {
            var settings = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            settings.Controls.Add(new Label
            {
                Text = "线宽（pt）", AutoSize = true, Margin = new Padding(0, 8, 4, 0)
            });
            lineWidth = new NumericUpDown
            {
                Minimum = 0.01M, Maximum = Math.Max(20M, (decimal)canvas.OutlineWidthPoints), Increment = 0.25M,
                DecimalPlaces = 2, Value = (decimal)canvas.OutlineWidthPoints, Width = 70
            };
            lineWidth.ValueChanged += (sender, e) =>
            {
                if (loadingEntry) return;
                canvas.OutlineWidthPoints = (float)lineWidth.Value;
                RefreshPreviews();
            };
            settings.Controls.Add(lineWidth);

            colorButton = new Button
            {
                Text = "描边颜色", Width = 90, Height = 26,
                BackColor = canvas.OutlineColor,
                ForeColor = canvas.OutlineColor.GetBrightness() > 0.5f ? Color.Black : Color.White,
                UseVisualStyleBackColor = false,
                Margin = new Padding(14, 2, 10, 0)
            };
            colorButton.Click += (sender, e) =>
            {
                using (var dialog = new ColorDialog { Color = canvas.OutlineColor, FullOpen = true })
                {
                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        canvas.OutlineColor = dialog.Color;
                        colorButton.BackColor = dialog.Color;
                        colorButton.ForeColor = dialog.Color.GetBrightness() > 0.5f ? Color.Black : Color.White;
                        RefreshPreviews();
                    }
                }
            };
            settings.Controls.Add(colorButton);
            settings.Controls.Add(new Label
            {
                Text = "线型", AutoSize = true, Margin = new Padding(0, 8, 4, 0)
            });
            lineStyle = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 110 };
            lineStyle.SelectedIndexChanged += (sender, e) =>
            {
                if (loadingEntry || lineStyle.SelectedItem == null) return;
                var option = (LineStyleOption)lineStyle.SelectedItem;
                canvas.OutlineDashStyle = option.Style;
                canvas.NativeOutlineDashStyle = option.NativeStyle;
                RefreshPreviews();
            };
            ResetLineStyleChoices();
            settings.Controls.Add(lineStyle);
            var squareCheckBox = new CheckBox
            {
                Text = "正方形", Checked = canvas.KeepSquare, AutoSize = true,
                Margin = new Padding(14, 5, 3, 0)
            };
            squareCheckBox.CheckedChanged += (sender, e) =>
            {
                canvas.KeepSquare = squareCheckBox.Checked;
                canvas.Focus();
            };
            settings.Controls.Add(squareCheckBox);
            return settings;
        }

        private void ResetLineStyleChoices()
        {
            lineStyle.Items.Clear();
            lineStyle.Items.AddRange(new object[]
            {
                new LineStyleOption("虚线", DashStyle.Dash),
                new LineStyleOption("实线", DashStyle.Solid),
                new LineStyleOption("点线", DashStyle.Dot),
                new LineStyleOption("点划线", DashStyle.DashDot),
                new LineStyleOption("双点划线", DashStyle.DashDotDot)
            });
            int initialStyleIndex = -1;
            if (canvas.NativeOutlineDashStyle.HasValue && canvas.NativeOutlineDashStyle.Value !=
                ZoomImageGeometry.GetOfficeDashStyle(canvas.OutlineDashStyle))
            {
                // 保留 Office 的长虚线、方点线等原生线型，避免继承时被近似替换。
                initialStyleIndex = lineStyle.Items.Add(new LineStyleOption("原矩形线型",
                    canvas.OutlineDashStyle, canvas.NativeOutlineDashStyle));
            }
            for (int i = 0; initialStyleIndex < 0 && i < lineStyle.Items.Count; i++)
            {
                if (((LineStyleOption)lineStyle.Items[i]).Style == canvas.OutlineDashStyle)
                {
                    initialStyleIndex = i;
                    break;
                }
            }
            lineStyle.SelectedIndex = initialStyleIndex;
        }

        private sealed class LineStyleOption
        {
            private readonly string label;
            public DashStyle Style { get; }
            public Office.MsoLineDashStyle? NativeStyle { get; }

            public LineStyleOption(string label, DashStyle style, Office.MsoLineDashStyle? nativeStyle = null)
            {
                this.label = label;
                Style = style;
                NativeStyle = nativeStyle;
            }

            public override string ToString() => label;
        }

        /// <summary>只显示布局结果，不修改框选坐标；与幻灯片生成共用位置和连线规则。</summary>
        private sealed class LayoutPreviewCanvas : Control
        {
            private readonly Image preview;
            private readonly float pictureWidthPoints;
            private readonly Func<IList<ZoomImageEntry>> getEntries;
            private readonly Func<ZoomImageEntry> getActiveEntry;

            public LayoutPreviewCanvas(Image preview, float pictureWidthPoints,
                Func<IList<ZoomImageEntry>> getEntries, Func<ZoomImageEntry> getActiveEntry)
            {
                this.preview = preview;
                this.pictureWidthPoints = Math.Max(0.01f, pictureWidthPoints);
                this.getEntries = getEntries;
                this.getActiveEntry = getActiveEntry;
                DoubleBuffered = true;
                ResizeRedraw = true;
                BackColor = Color.FromArgb(245, 246, 248);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                e.Graphics.DrawString("布局预览", Font, Brushes.DimGray, 10, 8);
                IList<ZoomImageEntry> items = getEntries();
                if (!items.Any(entry => entry.HasRegion))
                {
                    e.Graphics.DrawString("框选后显示放大图和辅助线", Font, Brushes.DimGray, 10, 34);
                    return;
                }

                var source = new RectangleF(0, 0, preview.Width, preview.Height);
                float pixelsPerPoint = preview.Width / pictureWidthPoints;
                var positions = ZoomImageLayout.Calculate(source, items, ZoomImageGeometry.GapPoints * pixelsPerPoint);
                RectangleF union = source;
                foreach (var pair in positions)
                {
                    PointF[] corners = ZoomImageGeometry.GetCorners(pair.Value,
                        pair.Key.PreservesLayout ? pair.Key.OriginalZoomRotation : 0);
                    RectangleF visibleBounds = RectangleF.FromLTRB(corners.Min(point => point.X), corners.Min(point => point.Y),
                        corners.Max(point => point.X), corners.Max(point => point.Y));
                    union = RectangleF.Union(union, visibleBounds);
                }
                float scale = Math.Min(Math.Max(1, ClientSize.Width - 24) / union.Width,
                    Math.Max(1, ClientSize.Height - 50) / union.Height);
                var origin = new PointF((ClientSize.Width - union.Width * scale) / 2f - union.Left * scale,
                    34 + (ClientSize.Height - 50 - union.Height * scale) / 2f - union.Top * scale);
                RectangleF sourceScreen = ToScreen(source, scale, origin);
                e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.DrawImage(preview, sourceScreen);
                foreach (ZoomImageEntry entry in items)
                {
                    if (!positions.TryGetValue(entry, out RectangleF zoom)) continue;
                    ZoomImageSelection options = entry.Options;
                    var region = new RectangleF(options.Region.X * preview.Width, options.Region.Y * preview.Height,
                        options.Region.Width * preview.Width, options.Region.Height * preview.Height);
                    RectangleF crop = options.RegionRotationDegrees == 0 ? RectangleF.Intersect(region, source) : region;
                    if (crop.Width <= 0 || crop.Height <= 0) continue;
                    RectangleF zoomScreen = ToScreen(zoom, scale, origin);
                    RectangleF regionScreen = ToScreen(region, scale, origin);
                    float zoomRotation = entry.PreservesLayout ? entry.OriginalZoomRotation : 0;
                    GraphicsState state = e.Graphics.Save();
                    try
                    {
                        float centerX = zoomScreen.Left + zoomScreen.Width / 2f;
                        float centerY = zoomScreen.Top + zoomScreen.Height / 2f;
                        e.Graphics.TranslateTransform(centerX, centerY);
                        e.Graphics.RotateTransform(zoomRotation);
                        e.Graphics.TranslateTransform(-centerX, -centerY);
                        DrawZoomContent(e.Graphics, source, zoomScreen, crop, options.RegionRotationDegrees);
                    }
                    finally { e.Graphics.Restore(state); }
                    DrawEntryOutlines(e.Graphics, entry, sourceScreen, zoomScreen, regionScreen,
                        pixelsPerPoint * scale, zoomRotation,
                        ZoomImageGeometry.GetRelativePlacement(source, zoom, options.Placement));
                    e.Graphics.DrawString(entry.DisplayName + (entry == getActiveEntry() ? "（当前）" : ""),
                        Font, Brushes.DimGray, zoomScreen.Left, zoomScreen.Bottom + 2);
                }
            }

            private static void DrawEntryOutlines(Graphics graphics, ZoomImageEntry entry, RectangleF sourceScreen,
                RectangleF zoomScreen, RectangleF regionScreen, float pixelsPerPoint, float zoomRotation,
                ZoomImagePlacement placement)
            {
                ZoomImageSelection options = entry.Options;
                float width = Math.Max(0.75f, options.OutlineWidthPoints * pixelsPerPoint);
                PointF[] regionCorners = ZoomImageGeometry.GetCorners(regionScreen, options.RegionRotationDegrees);
                using (var regionPen = new Pen(options.OutlineColor, width) { DashStyle = options.OutlineDashStyle })
                {
                    if (options.AddGuideLines)
                    {
                        PointF[] endpoints = ZoomImageGeometry.GetGuideEndpoints(
                            regionCorners, ZoomImageGeometry.GetCorners(zoomScreen, zoomRotation),
                            placement);
                        for (int i = 0; i < endpoints.Length; i += 2)
                        {
                            PointF start = endpoints[i], end = endpoints[i + 1];
                            if (options.GuideLineExtent == ZoomGuideLineExtent.InsideSourceImage &&
                                !ZoomImageGeometry.TryClipLineToRectangle(start, end, sourceScreen, 0,
                                    out start, out end)) continue;
                            graphics.DrawLine(regionPen, start, end);
                        }
                    }
                    graphics.DrawPolygon(regionPen, regionCorners);
                }
                if (options.UseRectangleColorForZoomImage)
                {
                    using (var zoomPen = new Pen(options.OutlineColor, width))
                    {
                        graphics.DrawPolygon(zoomPen, ZoomImageGeometry.GetCorners(zoomScreen, zoomRotation));
                    }
                }
            }

            private void DrawZoomContent(Graphics graphics, RectangleF source, RectangleF destination,
                RectangleF region, float rotationDegrees)
            {
                if (rotationDegrees == 0)
                {
                    graphics.DrawImage(preview, destination, region, GraphicsUnit.Pixel);
                    return;
                }

                GraphicsState state = graphics.Save();
                try
                {
                    graphics.SetClip(destination);
                    graphics.TranslateTransform(destination.Left + destination.Width / 2f,
                        destination.Top + destination.Height / 2f);
                    graphics.ScaleTransform(destination.Width / region.Width, destination.Height / region.Height);
                    graphics.RotateTransform(-rotationDegrees);
                    graphics.TranslateTransform(-region.Left - region.Width / 2f, -region.Top - region.Height / 2f);
                    graphics.DrawImage(preview, source);
                }
                finally
                {
                    graphics.Restore(state);
                }
            }

            private static RectangleF ToScreen(RectangleF bounds, float scale, PointF origin)
            {
                return new RectangleF(origin.X + bounds.Left * scale, origin.Y + bounds.Top * scale,
                    bounds.Width * scale, bounds.Height * scale);
            }

        }

    }
}

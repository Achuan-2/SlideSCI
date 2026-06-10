using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    public class ShapeLibraryControl : UserControl
    {
        private Panel topPanel;
        private TextBox txtSearch;
        private Button btnSave;
        private Button btnRefresh;
        private Button btnOpenFolder;
        private Button btnImport;
        private Button btnExport;
        private Button btnBack;
        private Label lblPath;
        private Button btnNewFolder;
        private FlowLayoutPanel flowLayout;
        private ContextMenuStrip contextMenu;
        private string libraryDir;
        private string currentDir;
        private Button btnToggleViewMode;
        private static bool isCardMode = false;

        public bool IsCardMode => isCardMode;

        public PowerPoint.DocumentWindow AssociatedWindow { get; set; }

        public ShapeLibraryControl()
        {
            InitializeLibraryDir();
            InitializeComponent();
            LoadLibraryItems();
        }

        private void InitializeLibraryDir()
        {
            libraryDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SlideSCI", "ShapeLibrary");
            currentDir = libraryDir;
            try
            {
                if (!Directory.Exists(libraryDir))
                {
                    Directory.CreateDirectory(libraryDir);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"创建素材库目录失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponent()
        {
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(248, 249, 250);
            this.Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point);

            // 1. FlowLayoutPanel for Grid of Cards (Added FIRST to DockStyle.Fill correctly below topPanel)
            flowLayout = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(10, 10, 10, 10),
                BackColor = Color.FromArgb(243, 244, 246)
            };
            // Enable double buffering for smooth scrolling
            typeof(FlowLayoutPanel).InvokeMember("DoubleBuffered",
                System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                null, flowLayout, new object[] { true });

            flowLayout.MouseUp += (s, e) =>
            {
                if (e.Button == MouseButtons.Right)
                {
                    ShowContextMenu(flowLayout, e.Location);
                }
            };

            this.Controls.Add(flowLayout);

            // 2. Top Control Panel (Added SECOND)
            topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 120, // Increased from 114 to accommodate the taller path indicator with back button
                Padding = new Padding(12, 10, 12, 10),
                BackColor = Color.White
            };
            this.Controls.Add(topPanel);

            // Initialize ToolTip early so it can be used by btnBack and others
            ToolTip toolTip = new ToolTip();

            // Path indicator panel (Bottom-most in topPanel, added FIRST)
            Panel pathRow = new Panel
            {
                Dock = DockStyle.Top,
                Height = 24, // Increased from 20 to fit the back button nicely
                Padding = new Padding(0)
            };
            topPanel.Controls.Add(pathRow);

            // Back Button (Icon only, docks Left in pathRow)
            btnBack = new Button
            {
                Text = "",
                Image = CreateIcon("back", 16, Color.FromArgb(51, 51, 51)),
                Width = 24,
                Height = 24,
                Dock = DockStyle.Left,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White, // Match pathRow background (white)
                ForeColor = Color.FromArgb(51, 51, 51),
                Cursor = Cursors.Hand,
                Enabled = false
            };
            btnBack.FlatAppearance.BorderSize = 0;
            btnBack.FlatAppearance.MouseOverBackColor = Color.FromArgb(240, 240, 240);
            btnBack.FlatAppearance.MouseDownBackColor = Color.FromArgb(220, 220, 220);
            btnBack.Click += (s, e) => GoBackToParent();
            toolTip.SetToolTip(btnBack, "返回上级文件夹");

            // Drag-and-drop to Back button moves items up one level
            btnBack.AllowDrop = true;
            btnBack.DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(typeof(LibraryCard)) || e.Data.GetDataPresent(typeof(FolderCard)))
                {
                    e.Effect = DragDropEffects.Move;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            };
            btnBack.DragDrop += (s, e) =>
            {
                if (currentDir == libraryDir) return;
                string parentDir = Path.GetDirectoryName(currentDir);
                if (e.Data.GetDataPresent(typeof(LibraryCard)))
                {
                    LibraryCard card = (LibraryCard)e.Data.GetData(typeof(LibraryCard));
                    if (card != null) MoveTargetTo(card, parentDir);
                }
                else if (e.Data.GetDataPresent(typeof(FolderCard)))
                {
                    FolderCard fCard = (FolderCard)e.Data.GetData(typeof(FolderCard));
                    if (fCard != null) MoveTargetTo(fCard, parentDir);
                }
            };

            // Initialize lblPath
            lblPath = new Label
            {
                Text = "素材库",
                Dock = DockStyle.Fill,
                AutoSize = false, // Disable AutoSize so Dock = Fill works correctly and aligns text relative to the remaining space
                Font = new Font("Microsoft YaHei", 8.5F, FontStyle.Bold),
                ForeColor = Color.FromArgb(120, 120, 120),
                TextAlign = ContentAlignment.MiddleLeft
            };

            // Spacer between back button and breadcrumb path
            Panel pathBackSpacer = new Panel { Dock = DockStyle.Left, Width = 6 };

            // Add controls in reverse docking order (largest index docks first, index 0 docks last)
            pathRow.Controls.Add(lblPath);         // index 0 - Dock = Fill (docks last/innermost)
            pathRow.Controls.Add(pathBackSpacer);  // index 1 - Dock = Left
            pathRow.Controls.Add(btnBack);         // index 2 - Dock = Left (docks first/outermost)

            // Spacer between Action Panel and Path Row
            Panel actionPathSpacer = new Panel { Dock = DockStyle.Top, Height = 6 };
            topPanel.Controls.Add(actionPathSpacer);

            // Action Buttons Panel (Middle-bottom, added SECOND)
            FlowLayoutPanel actionPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 32,
                Margin = new Padding(0),
                Padding = new Padding(0),
                FlowDirection = FlowDirection.LeftToRight
            };
            topPanel.Controls.Add(actionPanel);

            // Save Selected Shapes Button (Primary CTA)
            btnSave = new Button
            {
                Text = "保存素材",
                Image = CreateIcon("save", 16, Color.White),
                ImageAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText,
                Height = 32,
                AutoSize = true,
                Margin = new Padding(0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 212), // Theme Blue
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold)
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 90, 158);
            btnSave.Click += BtnSave_Click;
            actionPanel.Controls.Add(btnSave);

            // Spacer
            Panel actionSpacer1 = new Panel { Width = 6, Height = 32, Margin = new Padding(0) };
            actionPanel.Controls.Add(actionSpacer1);

            // New Folder Button
            btnNewFolder = new Button
            {
                Text = " 新建文件夹",
                Image = CreateIcon("newfolder", 16, Color.FromArgb(51, 51, 51)),
                ImageAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText,
                Height = 32,
                AutoSize = true,
                Margin = new Padding(0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(51, 51, 51),
                Cursor = Cursors.Hand,
                Font = new Font("Microsoft YaHei", 9F)
            };
            btnNewFolder.FlatAppearance.BorderSize = 0;
            btnNewFolder.FlatAppearance.MouseOverBackColor = Color.FromArgb(220, 220, 220);
            btnNewFolder.Click += BtnNewFolder_Click;
            actionPanel.Controls.Add(btnNewFolder);

            // Spacer
            Panel actionSpacerNew = new Panel { Width = 6, Height = 32, Margin = new Padding(0) };
            actionPanel.Controls.Add(actionSpacerNew);

            // Import Button
            btnImport = new Button
            {
                Text = " 导入",
                Image = CreateIcon("import", 16, Color.FromArgb(51, 51, 51)),
                ImageAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText,
                Height = 32,
                AutoSize = true,
                Margin = new Padding(0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(51, 51, 51),
                Cursor = Cursors.Hand,
                Font = new Font("Microsoft YaHei", 9F)
            };
            btnImport.FlatAppearance.BorderSize = 0;
            btnImport.FlatAppearance.MouseOverBackColor = Color.FromArgb(220, 220, 220);
            btnImport.Click += BtnImport_Click;
            actionPanel.Controls.Add(btnImport);

            // Spacer
            Panel actionSpacer2 = new Panel { Width = 6, Height = 32, Margin = new Padding(0) };
            actionPanel.Controls.Add(actionSpacer2);

            // Export Button
            btnExport = new Button
            {
                Text = " 导出",
                Image = CreateIcon("export", 16, Color.FromArgb(51, 51, 51)),
                ImageAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText,
                Height = 32,
                AutoSize = true,
                Margin = new Padding(0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(51, 51, 51),
                Cursor = Cursors.Hand,
                Font = new Font("Microsoft YaHei", 9F)
            };
            btnExport.FlatAppearance.BorderSize = 0;
            btnExport.FlatAppearance.MouseOverBackColor = Color.FromArgb(220, 220, 220);
            btnExport.Click += BtnExport_Click;
            actionPanel.Controls.Add(btnExport);

            // Spacer between Search Row and Action Panel (Middle-top, added THIRD)
            Panel topSpacer = new Panel
            {
                Dock = DockStyle.Top,
                Height = 6
            };
            topPanel.Controls.Add(topSpacer);

            // Search & Icons panel (Top-most, added FOURTH)
            Panel searchRow = new Panel
            {
                Dock = DockStyle.Top,
                Height = 28
            };
            topPanel.Controls.Add(searchRow);

            // Open Folder Button (Icon only)
            btnOpenFolder = new Button
            {
                Text = "",
                Image = CreateIcon("folder", 16, Color.FromArgb(51, 51, 51)),
                Width = 28,
                Height = 28,
                Dock = DockStyle.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(51, 51, 51),
                Cursor = Cursors.Hand
            };
            btnOpenFolder.FlatAppearance.BorderSize = 0;
            btnOpenFolder.FlatAppearance.MouseOverBackColor = Color.FromArgb(220, 220, 220);
            btnOpenFolder.Click += (s, e) => OpenLibraryFolder();
            toolTip.SetToolTip(btnOpenFolder, "打开素材文件夹");
            searchRow.Controls.Add(btnOpenFolder);

            // Refresh Button (Icon only)
            btnRefresh = new Button
            {
                Text = "",
                Image = CreateIcon("refresh", 16, Color.FromArgb(51, 51, 51)),
                Width = 28,
                Height = 28,
                Dock = DockStyle.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.FromArgb(51, 51, 51),
                Cursor = Cursors.Hand
            };
            btnRefresh.FlatAppearance.BorderSize = 0;
            btnRefresh.FlatAppearance.MouseOverBackColor = Color.FromArgb(220, 220, 220);
            btnRefresh.Click += (s, e) => LoadLibraryItems();
            toolTip.SetToolTip(btnRefresh, "刷新素材列表");
            searchRow.Controls.Add(btnRefresh);

            // Search input wrapper (to simulate flat border)
            Panel searchWrapper = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 28,
                BackColor = Color.White,
                Padding = new Padding(6, 4, 6, 4)
            };
            searchRow.Controls.Add(searchWrapper);

            // Search TextBox
            txtSearch = new TextBox
            {
                Text = "",
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei", 9F),
                ForeColor = Color.Gray
            };
            // Set Placeholder text
            SetSearchPlaceholder();
            txtSearch.Enter += TxtSearch_Enter;
            txtSearch.Leave += TxtSearch_Leave;
            txtSearch.TextChanged += TxtSearch_TextChanged;
            searchWrapper.Controls.Add(txtSearch);

            // Customize search border drawing
            searchWrapper.Paint += (s, e) =>
            {
                Color borderColor = txtSearch.Focused ? Color.FromArgb(0, 120, 212) : Color.FromArgb(220, 220, 220);
                using (Pen pen = new Pen(borderColor, 1))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, searchWrapper.Width - 1, searchWrapper.Height - 1);
                }
            };
            txtSearch.GotFocus += (s, e) => searchWrapper.Invalidate();
            txtSearch.LostFocus += (s, e) => searchWrapper.Invalidate();

            // 3. Context Menu Setup
            contextMenu = new ContextMenuStrip();
            contextMenu.Opening += ContextMenu_Opening;

            // 4. View Mode Toggle Button Setup
            btnToggleViewMode = new Button
            {
                Size = new Size(36, 36),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                Cursor = Cursors.Hand,
            };
            btnToggleViewMode.FlatAppearance.BorderSize = 0;

            // Make it circular
            try
            {
                System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
                path.AddEllipse(0, 0, btnToggleViewMode.Width, btnToggleViewMode.Height);
                btnToggleViewMode.Region = new Region(path);
            }
            catch { }

            // Custom drawing with hover state
            bool isBtnHovered = false;
            btnToggleViewMode.MouseEnter += (s, e) => { isBtnHovered = true; btnToggleViewMode.Invalidate(); };
            btnToggleViewMode.MouseLeave += (s, e) => { isBtnHovered = false; btnToggleViewMode.Invalidate(); };

            btnToggleViewMode.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                Color circleColor;
                Color iconColor;
                if (isCardMode)
                {
                    circleColor = isBtnHovered ? Color.FromArgb(0, 90, 158) : Color.FromArgb(0, 120, 212);
                    iconColor = Color.White;
                }
                else
                {
                    circleColor = isBtnHovered ? Color.FromArgb(240, 240, 240) : Color.White;
                    iconColor = Color.FromArgb(51, 51, 51);
                }

                using (Brush b = new SolidBrush(circleColor))
                {
                    e.Graphics.FillEllipse(b, 0, 0, btnToggleViewMode.Width, btnToggleViewMode.Height);
                }

                // Draw border for white button
                if (!isCardMode)
                {
                    using (Pen p = new Pen(Color.FromArgb(220, 220, 220), 1))
                    {
                        e.Graphics.DrawEllipse(p, 0, 0, btnToggleViewMode.Width - 1, btnToggleViewMode.Height - 1);
                    }
                }

                // Draw the icon centered (grid when card mode is active, card when grid mode is active)
                string iconType = isCardMode ? "grid" : "card";
                using (Bitmap icon = CreateIcon(iconType, 16, iconColor))
                {
                    int x = (btnToggleViewMode.Width - icon.Width) / 2;
                    int y = (btnToggleViewMode.Height - icon.Height) / 2;
                    e.Graphics.DrawImage(icon, x, y);
                }
            };

            btnToggleViewMode.Click += (s, e) =>
            {
                isCardMode = !isCardMode;
                btnToggleViewMode.Invalidate();
                toolTip.SetToolTip(btnToggleViewMode, isCardMode ? "切换至网格视图" : "切换至卡片视图");
                LoadLibraryItems();
            };

            toolTip.SetToolTip(btnToggleViewMode, isCardMode ? "切换至网格视图" : "切换至卡片视图");
            this.Controls.Add(btnToggleViewMode);
            btnToggleViewMode.BringToFront();

            // Hook layout resize events
            this.Resize += (s, e) => UpdateToggleButtonPosition();
            flowLayout.Resize += (s, e) => {
                if (isCardMode)
                {
                    UpdateCardsLayout();
                }
            };
        }

        private void SetSearchPlaceholder()
        {
            if (string.IsNullOrEmpty(txtSearch.Text))
            {
                txtSearch.Text = "🔍 搜索素材...";
                txtSearch.ForeColor = Color.Gray;
            }
        }

        private void TxtSearch_Enter(object sender, EventArgs e)
        {
            if (txtSearch.Text == "🔍 搜索素材...")
            {
                txtSearch.Text = "";
                txtSearch.ForeColor = Color.FromArgb(33, 33, 33);
            }
        }

        private void TxtSearch_Leave(object sender, EventArgs e)
        {
            SetSearchPlaceholder();
        }

        private void TxtSearch_TextChanged(object sender, EventArgs e)
        {
            string query = txtSearch.Text.Trim();
            if (string.IsNullOrEmpty(query) || txtSearch.Text == "🔍 搜索素材...")
            {
                LoadLibraryItems();
            }
            else
            {
                SearchAndFlattenMaterials(query);
            }
        }

        public void LoadLibraryItems()
        {
            string query = (txtSearch == null || txtSearch.Text == "🔍 搜索素材...") ? "" : txtSearch.Text.Trim();
            if (!string.IsNullOrEmpty(query))
            {
                SearchAndFlattenMaterials(query);
                return;
            }

            flowLayout.SuspendLayout();
            // Clear existing controls and dispose bitmaps to avoid memory leaks
            foreach (Control ctrl in flowLayout.Controls)
            {
                if (ctrl is LibraryCard card)
                {
                    card.DisposeCard();
                }
                else if (ctrl is FolderCard fCard)
                {
                    fCard.Dispose();
                }
            }
            flowLayout.Controls.Clear();

            if (!Directory.Exists(currentDir))
            {
                currentDir = libraryDir;
            }

            try
            {
                // Update Back button state
                btnBack.Enabled = (currentDir != libraryDir);
                btnBack.Image = CreateIcon("back", 16, btnBack.Enabled ? Color.FromArgb(51, 51, 51) : Color.LightGray);

                // Load subfolders first
                string[] subDirs = Directory.GetDirectories(currentDir);
                Array.Sort(subDirs);
                foreach (string dirPath in subDirs)
                {
                    FolderCard fCard = new FolderCard(dirPath, this);
                    flowLayout.Controls.Add(fCard);
                }

                // Load library items (shapes)
                string[] pngFiles = Directory.GetFiles(currentDir, "*.png");
                Array.Sort(pngFiles); // Sort alphabetically

                foreach (string pngPath in pngFiles)
                {
                    string assetName = Path.GetFileNameWithoutExtension(pngPath);
                    string pptxPath = Path.Combine(currentDir, assetName + ".pptx");

                    if (File.Exists(pptxPath))
                    {
                        LibraryCard card = new LibraryCard(assetName, pngPath, pptxPath, this);
                        flowLayout.Controls.Add(card);
                    }
                }

                // Update path label
                string relPath = currentDir.Substring(libraryDir.Length).Replace(Path.DirectorySeparatorChar, '>');
                if (string.IsNullOrEmpty(relPath))
                {
                    lblPath.Text = "素材库";
                }
                else
                {
                    lblPath.Text = "素材库 " + relPath;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载素材库列表时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            flowLayout.ResumeLayout();
            UpdateToggleButtonPosition();
        }

        private void SearchAndFlattenMaterials(string query)
        {
            flowLayout.SuspendLayout();
            
            // Clear existing controls and dispose bitmaps to avoid memory leaks
            foreach (Control ctrl in flowLayout.Controls)
            {
                if (ctrl is LibraryCard card)
                {
                    card.DisposeCard();
                }
                else if (ctrl is FolderCard fCard)
                {
                    fCard.Dispose();
                }
            }
            flowLayout.Controls.Clear();

            if (!Directory.Exists(currentDir))
            {
                currentDir = libraryDir;
            }

            try
            {
                // Disable Back button during search to avoid confusion
                btnBack.Enabled = false;
                btnBack.Image = CreateIcon("back", 16, Color.LightGray);

                // Recursively get all .pptx files in currentDir
                string[] pptxFiles = Directory.GetFiles(currentDir, "*.pptx", SearchOption.AllDirectories);
                
                // Sort alphabetically by file name (not full path) for consistent display
                Array.Sort(pptxFiles, (a, b) => string.Compare(Path.GetFileName(a), Path.GetFileName(b), StringComparison.OrdinalIgnoreCase));

                foreach (string pptxPath in pptxFiles)
                {
                    string assetName = Path.GetFileNameWithoutExtension(pptxPath);
                    if (assetName.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string pngPath = Path.Combine(Path.GetDirectoryName(pptxPath), assetName + ".png");
                        if (File.Exists(pngPath))
                        {
                            LibraryCard card = new LibraryCard(assetName, pngPath, pptxPath, this);
                            flowLayout.Controls.Add(card);
                        }
                    }
                }

                // Show breadcrumbs path with searching info
                string relPath = currentDir.Substring(libraryDir.Length).Replace(Path.DirectorySeparatorChar, '>');
                string prefix = string.IsNullOrEmpty(relPath) ? "素材库" : "素材库 " + relPath;
                lblPath.Text = $"{prefix} (搜索: \"{query}\")";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"搜索素材时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            flowLayout.ResumeLayout();
            UpdateToggleButtonPosition();
        }

        private void OpenLibraryFolder()
        {
            try
            {
                if (Directory.Exists(libraryDir))
                {
                    System.Diagnostics.Process.Start("explorer.exe", libraryDir);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开素材库文件夹失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            if (app.Presentations.Count == 0 || app.ActiveWindow == null || app.ActiveWindow.View == null)
            {
                MessageBox.Show("当前未打开任何幻灯片，请先打开演示文稿并选中形状。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            PowerPoint.Selection selection = null;
            try
            {
                selection = app.ActiveWindow.Selection;
            }
            catch
            {
                MessageBox.Show("获取选择失败，请确保您选中了幻灯片中的形状。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count == 0)
            {
                MessageBox.Show("请先在幻灯片中选择要保存的形状（可框选多个形状组合）。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                if (selection != null) Marshal.ReleaseComObject(selection);
                return;
            }

            // Prompt name using InputDialog, showing folder selection (passing libraryDir as root, and currentDir as default)
            using (InputDialog dlg = new InputDialog("保存素材", "请输入素材名称：", "自定义素材", libraryDir, currentDir, isFolder: false, showFolderSelect: true))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    string assetName = dlg.InputText;
                    SaveSelectedShapes(selection.ShapeRange, assetName, dlg.SelectedFolder);
                }
            }

            Marshal.ReleaseComObject(selection);
        }

        private void BtnImport_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Zip文件 (*.zip)|*.zip";
                openFileDialog.Title = "导入素材库";
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        using (ZipArchive archive = ZipFile.OpenRead(openFileDialog.FileName))
                        {
                            bool? overwriteAll = null; // null = ask, true = overwrite, false = skip
                            int importCount = 0;

                            // First, validate zip structure to make sure it contains expected files
                            bool hasValidEntries = false;
                            foreach (ZipArchiveEntry entry in archive.Entries)
                            {
                                string ext = Path.GetExtension(entry.FullName).ToLower();
                                if (ext == ".pptx" || ext == ".png")
                                {
                                    hasValidEntries = true;
                                    break;
                                }
                            }

                            if (!hasValidEntries)
                            {
                                MessageBox.Show("所选 ZIP 文件不包含任何有效的素材文件 (.pptx / .png)。", "导入失败", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            foreach (ZipArchiveEntry entry in archive.Entries)
                            {
                                string ext = Path.GetExtension(entry.FullName).ToLower();
                                if (ext != ".pptx" && ext != ".png") continue;

                                // Resolve target path preserving subdirectory structure
                                string entryPath = entry.FullName.Replace('/', Path.DirectorySeparatorChar);
                                string targetPath = Path.GetFullPath(Path.Combine(libraryDir, entryPath));

                                // Prevent zip slip: check if targetPath starts with libraryDir
                                if (!targetPath.StartsWith(libraryDir, StringComparison.OrdinalIgnoreCase))
                                {
                                    continue; // Skip unsafe paths
                                }

                                string targetDir = Path.GetDirectoryName(targetPath);
                                if (!Directory.Exists(targetDir))
                                {
                                    Directory.CreateDirectory(targetDir);
                                }

                                if (File.Exists(targetPath))
                                {
                                    if (overwriteAll == null)
                                    {
                                        var result = MessageBox.Show(
                                            "导入的素材中包含与本地重名的素材，是否覆盖现有本地素材？\n\n- 点击【是】: 覆盖所有同名素材\n- 点击【否】: 跳过重名素材，只导入新素材\n- 点击【取消】: 中止导入",
                                            "导入冲突",
                                            MessageBoxButtons.YesNoCancel,
                                            MessageBoxIcon.Question
                                        );
                                        if (result == DialogResult.Yes)
                                        {
                                            overwriteAll = true;
                                        }
                                        else if (result == DialogResult.No)
                                        {
                                            overwriteAll = false;
                                        }
                                        else
                                        {
                                            return; // Abort
                                        }
                                    }

                                    if (overwriteAll == false)
                                    {
                                        continue; // Skip
                                    }
                                }

                                entry.ExtractToFile(targetPath, overwrite: true);
                                if (ext == ".pptx")
                                {
                                    importCount++;
                                }
                            }

                            MessageBox.Show($"素材导入成功！共导入 {importCount} 个素材。", "导入成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            LoadLibraryItems();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"导入素材时发生错误: {ex.Message}", "导入错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            int exportCount = Directory.GetFiles(libraryDir, "*.pptx", SearchOption.AllDirectories).Length;
            if (exportCount == 0)
            {
                MessageBox.Show("当前素材库没有任何素材可以导出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Zip文件 (*.zip)|*.zip";
                saveFileDialog.Title = "导出素材库";
                saveFileDialog.FileName = "SlideSCI_ShapeLibrary.zip";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (File.Exists(saveFileDialog.FileName))
                        {
                            File.Delete(saveFileDialog.FileName);
                        }

                        // Dispose all cards to release any potential GDI+ locks on images before zipping
                        foreach (Control ctrl in flowLayout.Controls)
                        {
                            if (ctrl is LibraryCard card) card.DisposeCard();
                        }
                        flowLayout.Controls.Clear();

                        ZipFile.CreateFromDirectory(libraryDir, saveFileDialog.FileName);

                        MessageBox.Show($"素材库成功导出至:\n{saveFileDialog.FileName}\n\n共导出 {exportCount} 个素材。", "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadLibraryItems();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"导出素材时发生错误: {ex.Message}", "导出错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadLibraryItems();
                    }
                }
            }
        }

        private void SaveSelectedShapes(PowerPoint.ShapeRange shapeRange, string assetName, string targetFolder)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Presentation tempPres = null;
            string pptxPath = Path.Combine(targetFolder, assetName + ".pptx");
            string pngPath = Path.Combine(targetFolder, assetName + ".png");

            try
            {
                // 1. Export Preview PNG directly from selection (transparent and high quality, original size)
                shapeRange.Export(pngPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG, 0, 0, PowerPoint.PpExportMode.ppScaleToFit);

                // 2. Create a hidden pptx to hold shape data
                tempPres = app.Presentations.Add(Office.MsoTriState.msoFalse);
                tempPres.PageSetup.SlideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                tempPres.PageSetup.SlideHeight = app.ActivePresentation.PageSetup.SlideHeight;

                PowerPoint.Slide tempSlide = tempPres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);

                // Copy selection shapes to clipboard and paste to temporary slide
                int retries = 3;
                bool copySuccess = false;
                while (retries > 0)
                {
                    try
                    {
                        shapeRange.Copy();
                        copySuccess = true;
                        break;
                    }
                    catch
                    {
                        retries--;
                        if (retries == 0) throw;
                        System.Threading.Thread.Sleep(50);
                    }
                }

                if (copySuccess)
                {
                    tempSlide.Shapes.Paste();
                    tempPres.SaveAs(pptxPath);
                }

                MessageBox.Show($"素材「{assetName}」保存成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadLibraryItems();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存素材时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Clean up corrupted files if exist
                try
                {
                    if (File.Exists(pptxPath)) File.Delete(pptxPath);
                    if (File.Exists(pngPath)) File.Delete(pngPath);
                }
                catch { }
            }
            finally
            {
                if (tempPres != null)
                {
                    try { tempPres.Close(); } catch { }
                    Marshal.ReleaseComObject(tempPres);
                }
            }
        }

        private void UpdateAsset(LibraryCard card)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            if (app.Presentations.Count == 0 || app.ActiveWindow == null || app.ActiveWindow.View == null)
            {
                MessageBox.Show("当前未打开任何幻灯片，请先打开演示文稿并选中形状。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            PowerPoint.Selection selection = null;
            try
            {
                selection = app.ActiveWindow.Selection;
            }
            catch
            {
                MessageBox.Show("获取选择失败，请确保您选中了幻灯片中的形状。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count == 0)
            {
                MessageBox.Show("请先在幻灯片中选择要更新的形状（可框选多个形状组合）。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                if (selection != null) Marshal.ReleaseComObject(selection);
                return;
            }

            var result = MessageBox.Show($"确定要用当前选中的形状更新素材「{card.AssetName}」吗？", "确认更新", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                card.DisposeCard(); // Release image locks before overwriting
                SaveSelectedShapes(selection.ShapeRange, card.AssetName, Path.GetDirectoryName(card.PptxPath));
            }

            Marshal.ReleaseComObject(selection);
        }

        public void InsertAsset(string pptxPath)
        {
            if (!File.Exists(pptxPath))
            {
                MessageBox.Show("素材源文件丢失，请尝试刷新列表。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            PowerPoint.Application app = Globals.ThisAddIn.Application;
            if (app.Presentations.Count == 0 || app.ActiveWindow == null || app.ActiveWindow.View == null)
            {
                MessageBox.Show("当前未打开任何活动幻灯片，无法插入素材。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            PowerPoint.Presentation activePres = app.ActivePresentation;
            PowerPoint.Slide activeSlide = null;
            try
            {
                activeSlide = app.ActiveWindow.View.Slide;
            }
            catch
            {
                MessageBox.Show("请先选择或切换到一张幻灯片。", "操作提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            PowerPoint.Presentation tempPres = null;

            // 1. Backup user's clipboard to prevent pollution
            System.Windows.Forms.IDataObject clipboardBackup = null;
            try
            {
                clipboardBackup = System.Windows.Forms.Clipboard.GetDataObject();
            }
            catch { }

            try
            {
                // Ensure active window is focused
                try { app.ActiveWindow.Activate(); } catch { }

                // 2. Open material pptx hidden
                tempPres = app.Presentations.Open(
                    pptxPath,
                    ReadOnly: Office.MsoTriState.msoTrue,
                    WithWindow: Office.MsoTriState.msoFalse
                );

                if (tempPres.Slides.Count > 0)
                {
                    PowerPoint.Slide tempSlide = tempPres.Slides[1];
                    if (tempSlide.Shapes.Count > 0)
                    {
                        PowerPoint.ShapeRange shapesToCopy = tempSlide.Shapes.Range();

                        // Clear clipboard before copy to avoid pasting stale content if copy fails
                        try { System.Windows.Forms.Clipboard.Clear(); } catch { }
                        System.Windows.Forms.Application.DoEvents();

                        // Copy shapes from the temporary hidden slide
                        int copyRetries = 10;
                        bool copySuccess = false;
                        while (copyRetries > 0)
                        {
                            try
                            {
                                shapesToCopy.Copy();
                                System.Windows.Forms.Application.DoEvents();

                                // Verify clipboard contains active object (not empty)
                                if (System.Windows.Forms.Clipboard.GetDataObject() != null)
                                {
                                    copySuccess = true;
                                    break;
                                }
                            }
                            catch { }
                            copyRetries--;
                            System.Windows.Forms.Application.DoEvents();
                            System.Threading.Thread.Sleep(50);
                        }

                        if (copySuccess)
                        {
                            // Paste onto the active slide
                            PowerPoint.ShapeRange pastedRange = null;
                            int pasteRetries = 15;
                            while (pasteRetries > 0)
                            {
                                try
                                {
                                    pastedRange = activeSlide.Shapes.Paste();
                                    break;
                                }
                                catch
                                {
                                    pasteRetries--;
                                    if (pasteRetries == 0) throw;
                                    System.Windows.Forms.Application.DoEvents();
                                    System.Threading.Thread.Sleep(50);
                                }
                            }

                            if (pastedRange != null)
                            {
                                try
                                {
                                    float minLeft = float.MaxValue;
                                    float minTop = float.MaxValue;
                                    float maxRight = float.MinValue;
                                    float maxBottom = float.MinValue;

                                    for (int i = 1; i <= pastedRange.Count; i++)
                                    {
                                        PowerPoint.Shape shape = pastedRange[i];
                                        float left = shape.Left;
                                        float top = shape.Top;
                                        float right = left + shape.Width;
                                        float bottom = top + shape.Height;

                                        if (left < minLeft) minLeft = left;
                                        if (top < minTop) minTop = top;
                                        if (right > maxRight) maxRight = right;
                                        if (bottom > maxBottom) maxBottom = bottom;
                                    }

                                    if (minLeft != float.MaxValue)
                                    {
                                        float rangeWidth = maxRight - minLeft;
                                        float rangeHeight = maxBottom - minTop;

                                        float slideWidth = activePres.PageSetup.SlideWidth;
                                        float slideHeight = activePres.PageSetup.SlideHeight;

                                        float targetLeft = (slideWidth - rangeWidth) / 2f;
                                        float targetTop = (slideHeight - rangeHeight) / 2f;
                                        float offsetX = targetLeft - minLeft;
                                        float offsetY = targetTop - minTop;

                                        for (int i = 1; i <= pastedRange.Count; i++)
                                        {
                                            PowerPoint.Shape shape = pastedRange[i];
                                            shape.Left += offsetX;
                                            shape.Top += offsetY;
                                        }
                                    }
                                }
                                catch { }

                                try
                                {
                                    pastedRange.Select(); // Highlight/select the pasted shapes
                                }
                                catch { }
                            }
                        }
                        else
                        {
                            throw new Exception("无法复制素材至剪贴板，请重试。");
                        }
                    }
                    else
                    {
                        MessageBox.Show("该素材中不包含任何形状。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("该素材中不包含任何幻灯片。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"插入素材时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 3. Clean up the temporary hidden presentation
                if (tempPres != null)
                {
                    try
                    {
                        tempPres.Close();
                    }
                    catch { }
                    Marshal.ReleaseComObject(tempPres);
                }

                // 4. Restore original slide selection to make sure focus is correct
                try
                {
                    activeSlide.Select();
                }
                catch { }

                // 5. Restore user's clipboard so it is not polluted
                if (clipboardBackup != null)
                {
                    try
                    {
                        System.Windows.Forms.Clipboard.SetDataObject(clipboardBackup, true);
                    }
                    catch { }
                }
            }
        }

        // Context Menu Handlers
        private object GetSelectedCardOrFolder()
        {
            Control ctrl = contextMenu.SourceControl;
            while (ctrl != null)
            {
                if (ctrl is LibraryCard card) return card;
                if (ctrl is FolderCard fCard) return fCard;
                ctrl = ctrl.Parent;
            }
            return null;
        }

        private void ContextMenu_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            contextMenu.Items.Clear();
            object selected = GetSelectedCardOrFolder();

            if (selected is LibraryCard card)
            {
                ToolStripMenuItem menuInsert = new ToolStripMenuItem("插入素材 (Enter)");
                menuInsert.Click += (s, ev) => InsertAsset(card.PptxPath);

                ToolStripMenuItem menuUpdate = new ToolStripMenuItem("更新素材");
                menuUpdate.Click += (s, ev) => UpdateAsset(card);

                ToolStripMenuItem menuRename = new ToolStripMenuItem("重命名");
                menuRename.Click += (s, ev) => RenameCard(card);

                ToolStripMenuItem menuMoveTo = new ToolStripMenuItem("移动到文件夹");
                PopulateMoveToMenu(menuMoveTo, card);

                ToolStripMenuItem menuDelete = new ToolStripMenuItem("删除素材 (Delete)");
                menuDelete.Click += (s, ev) => DeleteCardPrompt(card);

                ToolStripMenuItem menuOpenDir = new ToolStripMenuItem("打开文件位置");
                menuOpenDir.Click += (s, ev) => OpenFileLocation(card.PptxPath);

                contextMenu.Items.Add(menuInsert);
                contextMenu.Items.Add(menuUpdate);
                contextMenu.Items.Add(menuRename);
                contextMenu.Items.Add(menuMoveTo);
                contextMenu.Items.Add(menuDelete);
                contextMenu.Items.Add(new ToolStripSeparator());
                contextMenu.Items.Add(menuOpenDir);
            }
            else if (selected is FolderCard fCard)
            {
                ToolStripMenuItem menuOpen = new ToolStripMenuItem("打开文件夹");
                menuOpen.Click += (s, ev) => EnterFolder(fCard.FolderPath);

                ToolStripMenuItem menuRename = new ToolStripMenuItem("重命名");
                menuRename.Click += (s, ev) => RenameFolder(fCard);

                ToolStripMenuItem menuMoveTo = new ToolStripMenuItem("移动到文件夹");
                PopulateMoveToMenu(menuMoveTo, fCard);

                ToolStripMenuItem menuDelete = new ToolStripMenuItem("删除文件夹");
                menuDelete.Click += (s, ev) => DeleteFolderPrompt(fCard);

                ToolStripMenuItem menuOpenDir = new ToolStripMenuItem("打开文件夹位置");
                menuOpenDir.Click += (s, ev) => OpenFileLocation(fCard.FolderPath);

                contextMenu.Items.Add(menuOpen);
                contextMenu.Items.Add(menuRename);
                contextMenu.Items.Add(menuMoveTo);
                contextMenu.Items.Add(menuDelete);
                contextMenu.Items.Add(new ToolStripSeparator());
                contextMenu.Items.Add(menuOpenDir);
            }
            else
            {
                ToolStripMenuItem menuNewFolder = new ToolStripMenuItem("新建文件夹");
                menuNewFolder.Click += BtnNewFolder_Click;

                ToolStripMenuItem menuRefresh = new ToolStripMenuItem("刷新");
                menuRefresh.Click += (s, ev) => LoadLibraryItems();

                ToolStripMenuItem menuImport = new ToolStripMenuItem("导入素材库");
                menuImport.Click += BtnImport_Click;

                ToolStripMenuItem menuExport = new ToolStripMenuItem("导出素材库");
                menuExport.Click += BtnExport_Click;

                contextMenu.Items.Add(menuNewFolder);
                contextMenu.Items.Add(menuRefresh);
                contextMenu.Items.Add(new ToolStripSeparator());
                contextMenu.Items.Add(menuImport);
                contextMenu.Items.Add(menuExport);
            }
        }

        private void PopulateMoveToMenu(ToolStripMenuItem parentMenu, object target)
        {
            string targetParent = (target is LibraryCard card) 
                ? Path.GetDirectoryName(card.PptxPath) 
                : Path.GetDirectoryName(((FolderCard)target).FolderPath);

            // Add Root Folder option if target is not already in root
            if (targetParent != libraryDir)
            {
                ToolStripMenuItem rootItem = new ToolStripMenuItem("根目录");
                rootItem.Click += (s, e) => MoveTargetTo(target, libraryDir);
                parentMenu.DropDownItems.Add(rootItem);
            }

            // Find all directories under libraryDir recursively
            if (Directory.Exists(libraryDir))
            {
                try
                {
                    string[] dirs = Directory.GetDirectories(libraryDir, "*", SearchOption.AllDirectories);
                    Array.Sort(dirs);

                    foreach (string dir in dirs)
                    {
                        // Check if it's the current parent directory
                        if (dir == targetParent) continue;

                        // If we are moving a folder, make sure it cannot be moved to itself or a child of itself
                        if (target is FolderCard fCard)
                        {
                            if (dir == fCard.FolderPath || dir.StartsWith(fCard.FolderPath + Path.DirectorySeparatorChar))
                            {
                                continue;
                            }
                        }

                        // Get relative path for display
                        string relPath = dir.Substring(libraryDir.Length).TrimStart(Path.DirectorySeparatorChar);
                        ToolStripMenuItem dirItem = new ToolStripMenuItem(relPath);
                        string destDir = dir;
                        dirItem.Click += (s, e) => MoveTargetTo(target, destDir);
                        parentMenu.DropDownItems.Add(dirItem);
                    }
                }
                catch (Exception)
                {
                    // Handle exception
                }
            }

            if (parentMenu.DropDownItems.Count == 0)
            {
                parentMenu.Enabled = false;
                ToolStripMenuItem noFolderItem = new ToolStripMenuItem("(无可用文件夹)");
                noFolderItem.Enabled = false;
                parentMenu.DropDownItems.Add(noFolderItem);
            }
        }

        private void MoveTargetTo(object target, string destDir)
        {
            try
            {
                if (target is LibraryCard card)
                {
                    string destPptx = Path.Combine(destDir, Path.GetFileName(card.PptxPath));
                    string destPng = Path.Combine(destDir, Path.GetFileName(card.PngPath));

                    card.DisposeCard(); // Release locks

                    if (File.Exists(card.PptxPath)) File.Move(card.PptxPath, destPptx);
                    if (File.Exists(card.PngPath)) File.Move(card.PngPath, destPng);
                }
                else if (target is FolderCard fCard)
                {
                    string destFolder = Path.Combine(destDir, Path.GetFileName(fCard.FolderPath));
                    Directory.Move(fCard.FolderPath, destFolder);
                }

                LoadLibraryItems();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"移动失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LoadLibraryItems();
            }
        }

        private void RenameCard(LibraryCard card)
        {
            string parentDir = Path.GetDirectoryName(card.PptxPath);
            using (InputDialog dlg = new InputDialog("重命名素材", "请输入新素材名称：", card.AssetName, parentDir))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    string newName = dlg.InputText;
                    string newPptxPath = Path.Combine(parentDir, newName + ".pptx");
                    string newPngPath = Path.Combine(parentDir, newName + ".png");

                    try
                    {
                        card.DisposeCard();
                        File.Move(card.PptxPath, newPptxPath);
                        File.Move(card.PngPath, newPngPath);
                        LoadLibraryItems();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"重命名失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadLibraryItems();
                    }
                }
            }
        }

        private void RenameFolder(FolderCard fCard)
        {
            string parentDir = Path.GetDirectoryName(fCard.FolderPath);
            string folderName = Path.GetFileName(fCard.FolderPath);
            using (InputDialog dlg = new InputDialog("重命名文件夹", "请输入新文件夹名称：", folderName, parentDir, isFolder: true))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    string newName = dlg.InputText;
                    string newFolderPath = Path.Combine(parentDir, newName);

                    try
                    {
                        Directory.Move(fCard.FolderPath, newFolderPath);
                        LoadLibraryItems();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"重命名文件夹失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LoadLibraryItems();
                    }
                }
            }
        }

        private void DeleteCardPrompt(LibraryCard card)
        {
            var result = MessageBox.Show($"确定要永久删除素材「{card.AssetName}」吗？", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                DeleteCardAsset(card);
            }
        }

        private void DeleteFolderPrompt(FolderCard fCard)
        {
            string folderName = Path.GetFileName(fCard.FolderPath);
            var result = MessageBox.Show($"确定要永久删除文件夹「{folderName}」及其包含的所有子目录和素材吗？此操作不可恢复！", "确认删除文件夹", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                try
                {
                    if (Directory.Exists(fCard.FolderPath))
                    {
                        Directory.Delete(fCard.FolderPath, true);
                    }
                    LoadLibraryItems();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"删除文件夹失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadLibraryItems();
                }
            }
        }

        private void OpenFileLocation(string path)
        {
            try
            {
                string argument = "/select, \"" + path + "\"";
                System.Diagnostics.Process.Start("explorer.exe", argument);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"定位失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void DeleteCardAsset(LibraryCard card)
        {
            try
            {
                card.DisposeCard();

                if (File.Exists(card.PptxPath)) File.Delete(card.PptxPath);
                if (File.Exists(card.PngPath)) File.Delete(card.PngPath);

                LoadLibraryItems();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LoadLibraryItems();
            }
        }

        public void ShowContextMenu(Control source, Point loc)
        {
            contextMenu.Show(source, loc);
        }

        private void UpdateToggleButtonPosition()
        {
            if (btnToggleViewMode != null)
            {
                int x = this.ClientSize.Width - btnToggleViewMode.Width - 24;
                int y = this.ClientSize.Height - btnToggleViewMode.Height - 24;

                if (x < 10) x = 10;
                if (y < (topPanel != null ? topPanel.Bottom : 120) + 10)
                    y = (topPanel != null ? topPanel.Bottom : 120) + 10;

                btnToggleViewMode.Location = new Point(x, y);
                btnToggleViewMode.BringToFront();
            }
        }

        public int GetFlowLayoutWidth()
        {
            if (flowLayout == null) return 200;
            int w = flowLayout.ClientSize.Width - flowLayout.Padding.Left - flowLayout.Padding.Right;
            return w > 130 ? w : 200;
        }

        public void UpdateCardsLayout()
        {
            if (!isCardMode || flowLayout == null) return;

            flowLayout.SuspendLayout();
            int targetWidth = GetFlowLayoutWidth() - 24;
            if (targetWidth < 130) targetWidth = 130;

            foreach (Control ctrl in flowLayout.Controls)
            {
                if (ctrl is LibraryCard card)
                {
                    card.ApplyCardModeLayout(targetWidth);
                }
                else if (ctrl is FolderCard fCard)
                {
                    fCard.ApplyCardModeLayout(targetWidth);
                }
            }
            flowLayout.ResumeLayout();
        }

        public void GoBackToParent()
        {
            if (currentDir != libraryDir)
            {
                string parentDir = Path.GetDirectoryName(currentDir);
                if (Directory.Exists(parentDir))
                {
                    currentDir = parentDir;
                    LoadLibraryItems();
                }
            }
        }

        public void EnterFolder(string path)
        {
            if (Directory.Exists(path))
            {
                currentDir = path;
                LoadLibraryItems();
            }
        }

        private void BtnNewFolder_Click(object sender, EventArgs e)
        {
            using (InputDialog dlg = new InputDialog("新建文件夹", "请输入文件夹名称：", "新建文件夹", currentDir, isFolder: true))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    string folderName = dlg.InputText;
                    string folderPath = Path.Combine(currentDir, folderName);
                    try
                    {
                        if (!Directory.Exists(folderPath))
                        {
                            Directory.CreateDirectory(folderPath);
                        }
                        LoadLibraryItems();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"创建文件夹失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        public static Bitmap CreateIcon(string type, int size, Color color)
        {
            Bitmap bmp = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                using (Pen pen = new Pen(color, 1.8f))
                using (Brush brush = new SolidBrush(color))
                {
                    if (type == "save")
                    {
                        g.DrawRectangle(pen, 2, 2, size - 4, size - 4);
                        g.FillRectangle(brush, 5, 2, size - 10, 4);
                        g.DrawRectangle(pen, 4, size - 7, size - 8, 5);
                    }
                    else if (type == "import")
                    {
                        g.DrawLine(pen, 2, size - 5, 2, size - 2);
                        g.DrawLine(pen, 2, size - 2, size - 3, size - 2);
                        g.DrawLine(pen, size - 3, size - 2, size - 3, size - 5);
                        int mid = size / 2;
                        g.DrawLine(pen, mid, 2, mid, size - 6);
                        g.DrawLine(pen, mid - 3, size - 9, mid, size - 6);
                        g.DrawLine(pen, mid + 3, size - 9, mid, size - 6);
                    }
                    else if (type == "export")
                    {
                        g.DrawLine(pen, 2, size - 5, 2, size - 2);
                        g.DrawLine(pen, 2, size - 2, size - 3, size - 2);
                        g.DrawLine(pen, size - 3, size - 2, size - 3, size - 5);
                        int mid = size / 2;
                        g.DrawLine(pen, mid, size - 6, mid, 2);
                        g.DrawLine(pen, mid - 3, 5, mid, 2);
                        g.DrawLine(pen, mid + 3, 5, mid, 2);
                    }
                    else if (type == "refresh")
                    {
                        g.DrawArc(pen, 3, 3, size - 6, size - 6, 45, 270);
                        PointF[] arrowPoints = new PointF[] {
                            new PointF(size - 3, 3),
                            new PointF(size - 6, 7),
                            new PointF(size - 9, 3)
                        };
                        g.FillPolygon(brush, arrowPoints);
                    }
                    else if (type == "folder")
                    {
                        g.DrawLine(pen, 2, 4, 6, 4);
                        g.DrawLine(pen, 6, 4, 8, 6);
                        g.DrawLine(pen, 8, 6, size - 3, 6);
                        g.DrawLine(pen, size - 3, 6, size - 3, size - 3);
                        g.DrawLine(pen, size - 3, size - 3, 2, size - 3);
                        g.DrawLine(pen, 2, size - 3, 2, 4);
                    }
                    else if (type == "newfolder")
                    {
                        g.DrawLine(pen, 2, 4, 6, 4);
                        g.DrawLine(pen, 6, 4, 8, 6);
                        g.DrawLine(pen, 8, 6, size - 3, 6);
                        g.DrawLine(pen, size - 3, 6, size - 3, size - 3);
                        g.DrawLine(pen, size - 3, size - 3, 2, size - 3);
                        g.DrawLine(pen, 2, size - 3, 2, 4);
                        // Draw a plus (+) symbol at the bottom right inside the folder
                        g.DrawLine(pen, size - 8, size - 6, size - 4, size - 6);
                        g.DrawLine(pen, size - 6, size - 8, size - 6, size - 4);
                    }
                    else if (type == "back")
                    {
                        int mid = size / 2;
                        g.DrawLine(pen, 3, mid, size - 4, mid);
                        g.DrawLine(pen, 3, mid, 8, mid - 4);
                        g.DrawLine(pen, 3, mid, 8, mid + 4);
                    }
                    else if (type == "trash")
                    {
                        g.DrawLine(pen, 2, 3, size - 3, 3);
                        g.DrawLine(pen, 5, 3, 5, 1);
                        g.DrawLine(pen, 5, 1, size - 6, 1);
                        g.DrawLine(pen, size - 6, 1, size - 6, 3);
                        g.DrawLine(pen, 3, 5, 4, size - 2);
                        g.DrawLine(pen, 4, size - 2, size - 5, size - 2);
                        g.DrawLine(pen, size - 5, size - 2, size - 4, 5);
                        g.DrawLine(pen, 3, 5, size - 4, 5);
                        g.DrawLine(pen, 6, 6, 6, size - 4);
                        g.DrawLine(pen, size - 7, 6, size - 7, size - 4);
                    }
                    else if (type == "grid")
                    {
                        int s = size / 2 - 2;
                        g.DrawRectangle(pen, 2, 2, s, s);
                        g.DrawRectangle(pen, size / 2 + 1, 2, s, s);
                        g.DrawRectangle(pen, 2, size / 2 + 1, s, s);
                        g.DrawRectangle(pen, size / 2 + 1, size / 2 + 1, s, s);
                    }
                    else if (type == "card")
                    {
                        int h = size / 2 - 2;
                        g.DrawRectangle(pen, 2, 2, size - 4, h);
                        g.DrawRectangle(pen, 2, size / 2 + 1, size - 4, h);
                    }
                }
            }
            return bmp;
        }
    }

    public class LibraryCard : Panel
    {
        private PictureBox picPreview;
        private Label lblName;
        private Button btnTrash;
        private bool isHovered = false;
        private ShapeLibraryControl parentContainer;

        public string AssetName { get; private set; }
        public string PngPath { get; private set; }
        public string PptxPath { get; private set; }

        public LibraryCard(string assetName, string pngPath, string pptxPath, ShapeLibraryControl parent)
        {
            this.AssetName = assetName;
            this.PngPath = pngPath;
            this.PptxPath = pptxPath;
            this.parentContainer = parent;

            InitializeCard();
        }

        private void InitializeCard()
        {
            this.Size = new Size(130, 135);
            this.Margin = new Padding(6, 6, 6, 6);
            this.BackColor = Color.White;
            this.Cursor = Cursors.Hand;

            // Load bitmap using a stream to avoid locking the PNG file on disk
            Image previewImg = null;
            try
            {
                if (File.Exists(PngPath))
                {
                    using (FileStream fs = new FileStream(PngPath, FileMode.Open, FileAccess.Read))
                    {
                        using (Image tempImg = Image.FromStream(fs))
                        {
                            previewImg = new Bitmap(tempImg);
                        }
                    }
                }
            }
            catch
            {
                // Fallback if image fails to load
            }

            // Picture Box (displays shape preview)
            picPreview = new PictureBox
            {
                Image = previewImg,
                SizeMode = PictureBoxSizeMode.Zoom,
                Size = new Size(118, 90),
                Location = new Point(6, 6),
                BackColor = Color.FromArgb(245, 245, 245), // Soft grey background to make white shapes visible
            };
            this.Controls.Add(picPreview);

            // Asset Name Label
            lblName = new Label
            {
                Text = AssetName,
                Size = new Size(118, 25),
                Location = new Point(6, 102),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei", 8.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            this.Controls.Add(lblName);

            // Trash Button (Float on hover)
            btnTrash = new Button
            {
                Text = "",
                Image = ShapeLibraryControl.CreateIcon("trash", 12, Color.White),
                Size = new Size(20, 20),
                Location = new Point(104, 8),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(230, 220, 53, 69), // Translucent red
                Cursor = Cursors.Hand,
                Visible = false
            };
            btnTrash.FlatAppearance.BorderSize = 0;
            btnTrash.Click += BtnTrash_Click;
            this.Controls.Add(btnTrash);
            btnTrash.BringToFront();

            // Setup Mouse Event Hooks for Hover effects
            BindHoverEvents(this);
            BindHoverEvents(picPreview);
            BindHoverEvents(lblName);

            // MouseClick Events (Insert Shapes only on Left Click)
            picPreview.MouseClick += Card_MouseClick;
            lblName.MouseClick += Card_MouseClick;
            this.MouseClick += Card_MouseClick;

            // MouseDoubleClick Events
            picPreview.MouseDoubleClick += Card_MouseDoubleClick;
            lblName.MouseDoubleClick += Card_MouseDoubleClick;
            this.MouseDoubleClick += Card_MouseDoubleClick;

            // Context Menu Registration
            picPreview.MouseUp += Card_MouseUp;
            lblName.MouseUp += Card_MouseUp;
            this.MouseUp += Card_MouseUp;

            // Setup Drag & Drop
            picPreview.MouseDown += Card_MouseDown;
            lblName.MouseDown += Card_MouseDown;
            this.MouseDown += Card_MouseDown;

            picPreview.MouseMove += Card_MouseMove;
            lblName.MouseMove += Card_MouseMove;
            this.MouseMove += Card_MouseMove;

            // Keyboard navigation (delete key)
            this.KeyUp += LibraryCard_KeyUp;

            if (parentContainer.IsCardMode)
            {
                int targetWidth = parentContainer.GetFlowLayoutWidth() - 24;
                ApplyCardModeLayout(targetWidth);
            }
        }

        private void LibraryCard_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                BtnTrash_Click(this, EventArgs.Empty);
            }
        }

        private void Card_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                parentContainer.InsertAsset(PptxPath);
            }
        }

        private void Card_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                parentContainer.InsertAsset(PptxPath);
            }
        }

        private void Card_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                parentContainer.ShowContextMenu((Control)sender, e.Location);
            }
        }

        private Point dragStartPoint;

        private void Card_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                dragStartPoint = e.Location;
            }
        }

        private void Card_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (Math.Abs(e.X - dragStartPoint.X) > SystemInformation.DragSize.Width ||
                    Math.Abs(e.Y - dragStartPoint.Y) > SystemInformation.DragSize.Height)
                {
                    this.DoDragDrop(this, DragDropEffects.Move);
                }
            }
        }

        private void BindHoverEvents(Control ctrl)
        {
            ctrl.MouseEnter += (s, e) => SetHoverState(true);
            ctrl.MouseLeave += (s, e) =>
            {
                Point clientMouse = this.PointToClient(Cursor.Position);
                if (!this.ClientRectangle.Contains(clientMouse))
                {
                    SetHoverState(false);
                }
            };
        }

        private void SetHoverState(bool hover)
        {
            if (isHovered == hover) return;
            isHovered = hover;

            if (isHovered)
            {
                this.BackColor = Color.FromArgb(243, 249, 255); // Light theme tint blue
                lblName.ForeColor = Color.FromArgb(0, 120, 212); // Blue highlights
                btnTrash.Visible = true;
            }
            else
            {
                this.BackColor = Color.White;
                lblName.ForeColor = Color.FromArgb(51, 51, 51);
                btnTrash.Visible = false;
            }
            this.Invalidate(); // Trigger repaint for border updates
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            Color borderColor = isHovered ? Color.FromArgb(0, 120, 212) : Color.FromArgb(220, 224, 230);
            int borderThickness = isHovered ? 2 : 1;

            using (Pen pen = new Pen(borderColor, borderThickness))
            {
                int offset = borderThickness == 1 ? 0 : 1;
                e.Graphics.DrawRectangle(pen, offset, offset, this.Width - borderThickness - offset, this.Height - borderThickness - offset);
            }
        }

        private void BtnTrash_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show($"确定要永久删除素材「{AssetName}」吗？", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                parentContainer.DeleteCardAsset(this);
            }
        }

        public void ApplyCardModeLayout(int targetWidth)
        {
            int padding = 6;
            int labelHeight = 25;
            int spacing = 6;

            int picWidth = targetWidth - (padding * 2);
            int picHeight = picWidth * 3 / 4; // 4:3 aspect ratio

            int cardHeight = padding + picHeight + spacing + labelHeight + padding;

            this.Size = new Size(targetWidth, cardHeight);

            picPreview.Location = new Point(padding, padding);
            picPreview.Size = new Size(picWidth, picHeight);

            lblName.Location = new Point(padding, padding + picHeight + spacing);
            lblName.Size = new Size(picWidth, labelHeight);

            btnTrash.Location = new Point(targetWidth - btnTrash.Width - padding - 2, padding + 2);
        }

        public void DisposeCard()
        {
            if (picPreview.Image != null)
            {
                var img = picPreview.Image;
                picPreview.Image = null;
                img.Dispose();
            }
            this.Dispose();
        }
    }

    public class FolderCard : Panel
    {
        private PictureBox picPreview;
        private Label lblName;
        private Button btnTrash;
        private bool isHovered = false;
        private ShapeLibraryControl parentContainer;

        public string FolderPath { get; private set; }

        public FolderCard(string folderPath, ShapeLibraryControl parent)
        {
            this.FolderPath = folderPath;
            this.parentContainer = parent;

            InitializeCard();
        }

        private void InitializeCard()
        {
            this.Size = new Size(130, 135);
            this.Margin = new Padding(6, 6, 6, 6);
            this.BackColor = Color.White;
            this.Cursor = Cursors.Hand;

            // Picture Box (displays folder icon)
            picPreview = new PictureBox
            {
                Image = ShapeLibraryControl.CreateIcon("folder", 48, Color.FromArgb(242, 175, 41)), // Classic folder yellow
                SizeMode = PictureBoxSizeMode.CenterImage,
                Size = new Size(118, 90),
                Location = new Point(6, 6),
                BackColor = Color.FromArgb(245, 245, 245),
            };
            this.Controls.Add(picPreview);

            // Folder Name Label
            string folderName = Path.GetFileName(FolderPath);
            lblName = new Label
            {
                Text = folderName,
                Size = new Size(118, 25),
                Location = new Point(6, 102),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei", 8.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(51, 51, 51)
            };
            this.Controls.Add(lblName);

            // Trash Button (Float on hover)
            btnTrash = new Button
            {
                Text = "",
                Image = ShapeLibraryControl.CreateIcon("trash", 12, Color.White),
                Size = new Size(20, 20),
                Location = new Point(104, 8),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(230, 220, 53, 69), // Translucent red
                Cursor = Cursors.Hand,
                Visible = false
            };
            btnTrash.FlatAppearance.BorderSize = 0;
            btnTrash.Click += BtnTrash_Click;
            this.Controls.Add(btnTrash);
            btnTrash.BringToFront();

            // Material count badge
            int materialCount = 0;
            if (Directory.Exists(FolderPath))
            {
                try
                {
                    materialCount = Directory.GetFiles(FolderPath, "*.pptx").Length;
                }
                catch { }
            }

            Label lblBadge = null;
            if (materialCount > 0)
            {
                string text = materialCount.ToString();
                int badgeWidth = text.Length > 1 ? 24 : 18;
                lblBadge = new Label
                {
                    Text = text,
                    AutoSize = false,
                    Size = new Size(badgeWidth, 18),
                    Location = new Point(10, 10),
                    BackColor = Color.FromArgb(0, 120, 212), // Theme Blue
                    ForeColor = Color.White,
                    Font = new Font("Microsoft YaHei", 7.5F, FontStyle.Bold),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand
                };

                try
                {
                    System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
                    int radius = 9; // Half of height (18)
                    path.AddArc(0, 0, radius * 2, radius * 2, 90, 180);
                    path.AddArc(lblBadge.Width - radius * 2 - 1, 0, radius * 2, radius * 2, 270, 180);
                    path.CloseFigure();
                    lblBadge.Region = new Region(path);
                }
                catch { }

                this.Controls.Add(lblBadge);
                lblBadge.BringToFront();

                // Click / Double Click / Context Menu / Drag Events for Badge
                lblBadge.MouseClick += Card_MouseClick;
                lblBadge.MouseDoubleClick += Card_MouseDoubleClick;
                lblBadge.MouseUp += Card_MouseUp;
                lblBadge.MouseDown += Card_MouseDown;
                lblBadge.MouseMove += Card_MouseMove;
                
                // Allow Drag-and-drop targets on badge
                lblBadge.AllowDrop = true;
                lblBadge.DragEnter += FolderCard_DragEnter;
                lblBadge.DragDrop += FolderCard_DragDrop;
            }

            // Setup Mouse Event Hooks for Hover effects
            BindHoverEvents(this);
            BindHoverEvents(picPreview);
            BindHoverEvents(lblName);
            if (lblBadge != null)
            {
                BindHoverEvents(lblBadge);
            }

            // Click / Double Click Events (Open Folder)
            picPreview.MouseClick += Card_MouseClick;
            lblName.MouseClick += Card_MouseClick;
            this.MouseClick += Card_MouseClick;

            picPreview.MouseDoubleClick += Card_MouseDoubleClick;
            lblName.MouseDoubleClick += Card_MouseDoubleClick;
            this.MouseDoubleClick += Card_MouseDoubleClick;

            // Context Menu Registration
            picPreview.MouseUp += Card_MouseUp;
            lblName.MouseUp += Card_MouseUp;
            this.MouseUp += Card_MouseUp;

            // Setup Drag & Drop
            picPreview.MouseDown += Card_MouseDown;
            lblName.MouseDown += Card_MouseDown;
            this.MouseDown += Card_MouseDown;

            picPreview.MouseMove += Card_MouseMove;
            lblName.MouseMove += Card_MouseMove;
            this.MouseMove += Card_MouseMove;

            // Make FolderCard a Drop Target
            this.AllowDrop = true;
            picPreview.AllowDrop = true;
            lblName.AllowDrop = true;

            this.DragEnter += FolderCard_DragEnter;
            picPreview.DragEnter += FolderCard_DragEnter;
            lblName.DragEnter += FolderCard_DragEnter;

            this.DragDrop += FolderCard_DragDrop;
            picPreview.DragDrop += FolderCard_DragDrop;
            lblName.DragDrop += FolderCard_DragDrop;

            // Keyboard navigation (delete key)
            this.KeyUp += FolderCard_KeyUp;

            if (parentContainer.IsCardMode)
            {
                int targetWidth = parentContainer.GetFlowLayoutWidth() - 24;
                ApplyCardModeLayout(targetWidth);
            }
        }

        private void FolderCard_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                BtnTrash_Click(this, EventArgs.Empty);
            }
        }

        private void Card_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                parentContainer.EnterFolder(FolderPath);
            }
        }

        private void Card_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                parentContainer.EnterFolder(FolderPath);
            }
        }

        private void Card_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                parentContainer.ShowContextMenu((Control)sender, e.Location);
            }
        }

        private Point dragStartPoint;

        private void Card_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                dragStartPoint = e.Location;
            }
        }

        private void Card_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (Math.Abs(e.X - dragStartPoint.X) > SystemInformation.DragSize.Width ||
                    Math.Abs(e.Y - dragStartPoint.Y) > SystemInformation.DragSize.Height)
                {
                    this.DoDragDrop(this, DragDropEffects.Move);
                }
            }
        }

        private void FolderCard_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(LibraryCard)) || e.Data.GetDataPresent(typeof(FolderCard)))
            {
                if (e.Data.GetDataPresent(typeof(FolderCard)))
                {
                    FolderCard dragged = (FolderCard)e.Data.GetData(typeof(FolderCard));
                    if (dragged != null && (dragged.FolderPath == this.FolderPath || this.FolderPath.StartsWith(dragged.FolderPath + Path.DirectorySeparatorChar)))
                    {
                        e.Effect = DragDropEffects.None;
                        return;
                    }
                }
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void FolderCard_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(LibraryCard)))
            {
                LibraryCard card = (LibraryCard)e.Data.GetData(typeof(LibraryCard));
                if (card != null)
                {
                    string destPptx = Path.Combine(this.FolderPath, Path.GetFileName(card.PptxPath));
                    string destPng = Path.Combine(this.FolderPath, Path.GetFileName(card.PngPath));
                    try
                    {
                        card.DisposeCard();
                        if (File.Exists(card.PptxPath)) File.Move(card.PptxPath, destPptx);
                        if (File.Exists(card.PngPath)) File.Move(card.PngPath, destPng);
                        parentContainer.LoadLibraryItems();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"移动素材失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        parentContainer.LoadLibraryItems();
                    }
                }
            }
            else if (e.Data.GetDataPresent(typeof(FolderCard)))
            {
                FolderCard dragged = (FolderCard)e.Data.GetData(typeof(FolderCard));
                if (dragged != null)
                {
                    if (dragged.FolderPath == this.FolderPath || this.FolderPath.StartsWith(dragged.FolderPath + Path.DirectorySeparatorChar))
                    {
                        return;
                    }
                    string destDir = Path.Combine(this.FolderPath, Path.GetFileName(dragged.FolderPath));
                    try
                    {
                        Directory.Move(dragged.FolderPath, destDir);
                        parentContainer.LoadLibraryItems();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"移动文件夹失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        parentContainer.LoadLibraryItems();
                    }
                }
            }
        }

        private void BindHoverEvents(Control ctrl)
        {
            ctrl.MouseEnter += (s, e) => SetHoverState(true);
            ctrl.MouseLeave += (s, e) =>
            {
                Point clientMouse = this.PointToClient(Cursor.Position);
                if (!this.ClientRectangle.Contains(clientMouse))
                {
                    SetHoverState(false);
                }
            };
        }

        private void SetHoverState(bool hover)
        {
            if (isHovered == hover) return;
            isHovered = hover;

            if (isHovered)
            {
                this.BackColor = Color.FromArgb(243, 249, 255);
                lblName.ForeColor = Color.FromArgb(0, 120, 212);
                btnTrash.Visible = true;
            }
            else
            {
                this.BackColor = Color.White;
                lblName.ForeColor = Color.FromArgb(51, 51, 51);
                btnTrash.Visible = false;
            }
            this.Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Color borderColor = isHovered ? Color.FromArgb(0, 120, 212) : Color.FromArgb(220, 224, 230);
            int borderThickness = isHovered ? 2 : 1;
            using (Pen pen = new Pen(borderColor, borderThickness))
            {
                int offset = borderThickness == 1 ? 0 : 1;
                e.Graphics.DrawRectangle(pen, offset, offset, this.Width - borderThickness - offset, this.Height - borderThickness - offset);
            }
        }

        public void ApplyCardModeLayout(int targetWidth)
        {
            this.Size = new Size(targetWidth, 60);

            picPreview.Location = new Point(6, 6);
            picPreview.Size = new Size(48, 48);
            picPreview.SizeMode = PictureBoxSizeMode.Zoom;

            lblName.Location = new Point(60, 17);
            lblName.Size = new Size(targetWidth - 140, 25);
            lblName.TextAlign = ContentAlignment.MiddleLeft;

            Label lblBadge = null;
            foreach (Control c in this.Controls)
            {
                if (c is Label && c != lblName)
                {
                    lblBadge = (Label)c;
                    break;
                }
            }

            if (lblBadge != null)
            {
                lblBadge.Location = new Point(targetWidth - 70, 21);
            }

            btnTrash.Location = new Point(targetWidth - btnTrash.Width - 10, 20);
        }

        private void BtnTrash_Click(object sender, EventArgs e)
        {
            string folderName = Path.GetFileName(FolderPath);
            var result = MessageBox.Show($"确定要永久删除文件夹「{folderName}」及其包含的所有内容吗？", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes)
            {
                try
                {
                    if (Directory.Exists(FolderPath)) Directory.Delete(FolderPath, true);
                    parentContainer.LoadLibraryItems();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"删除文件夹失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    parentContainer.LoadLibraryItems();
                }
            }
        }
    }
}

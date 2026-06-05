using System;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SlideSCI
{
    public class InputDialog : Form
    {
        private TextBox txtInput;
        private Label lblError;
        private Button btnOk;
        private Button btnCancel;
        private string libraryDir;
        private string libraryRoot;
        private string currentDir;
        private bool isFolder;
        private bool showFolderSelect;
        private ComboBox cmbFolder;
        private Label lblFolderPrompt;

        public string SelectedFolder { get; private set; }
        public string InputText { get; private set; }

        private class FolderItem
        {
            public string DisplayName { get; set; }
            public string FullPath { get; set; }
            public override string ToString() => DisplayName;
        }

        // Overload constructor for backwards compatibility
        public InputDialog(string title, string prompt, string defaultText, string libraryDir, bool isFolder = false)
            : this(title, prompt, defaultText, libraryDir, libraryDir, isFolder, false)
        {
        }

        // Main constructor supporting folder selection
        public InputDialog(string title, string prompt, string defaultText, string libraryRoot, string currentDir, bool isFolder = false, bool showFolderSelect = false)
        {
            this.libraryRoot = libraryRoot;
            this.currentDir = currentDir;
            this.libraryDir = currentDir; // Validation will use the selected folder
            this.isFolder = isFolder;
            this.showFolderSelect = showFolderSelect;
            this.SelectedFolder = currentDir; // Default to current directory
            InitializeComponent(title, prompt, defaultText);
        }

        private void InitializeComponent(string title, string prompt, string defaultText)
        {
            // Form setup
            this.Text = title;
            this.Size = new Size(380, showFolderSelect ? 290 : 220);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.BackColor = Color.FromArgb(248, 249, 250); // Soft grey-white background
            this.Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point);

            // Container Panel (for margins/padding)
            Panel container = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(20, 15, 20, 15)
            };
            this.Controls.Add(container);

            // 1. Declare spacer at the bottom of top container
            Panel spacer = new Panel
            {
                Dock = DockStyle.Top,
                Height = 8
            };

            // 2. Declare lblError (for validation messages)
            lblError = new Label
            {
                Dock = DockStyle.Top,
                Height = 22,
                ForeColor = Color.FromArgb(220, 53, 69), // Bootstrap danger red
                Font = new Font("Microsoft YaHei", 8.5F),
                TextAlign = ContentAlignment.MiddleLeft,
                Text = ""
            };

            // Folder Selection ComboBox (if enabled)
            if (showFolderSelect)
            {
                cmbFolder = new ComboBox
                {
                    Dock = DockStyle.Top,
                    Height = 28,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Font = new Font("Microsoft YaHei", 9.5F),
                    BackColor = Color.White,
                    ForeColor = Color.FromArgb(33, 33, 33)
                };

                lblFolderPrompt = new Label
                {
                    Text = "保存位置：",
                    Dock = DockStyle.Top,
                    Height = 24,
                    ForeColor = Color.FromArgb(51, 51, 51),
                    Font = new Font("Microsoft YaHei", 9.5F, FontStyle.Bold),
                    TextAlign = ContentAlignment.BottomLeft
                };

                PopulateFolders();
            }

            // Folder spacer between textbox and folder prompt
            Panel spacerFolder = showFolderSelect ? new Panel { Dock = DockStyle.Top, Height = 10 } : null;

            // 3. Declare txtWrapper (Input box border wrapper to simulate flat border)
            Panel txtWrapper = new Panel
            {
                Dock = DockStyle.Top,
                Height = 32,
                BackColor = Color.White,
                Padding = new Padding(6, 6, 6, 6)
            };

            // TextBox
            txtInput = new TextBox
            {
                Text = defaultText,
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei", 10F),
                ForeColor = Color.FromArgb(33, 33, 33)
            };
            txtWrapper.Controls.Add(txtInput);

            // 4. Declare Prompt Label
            Label lblPrompt = new Label
            {
                Text = prompt,
                Dock = DockStyle.Top,
                Height = 25,
                ForeColor = Color.FromArgb(51, 51, 51),
                Font = new Font("Microsoft YaHei", 9.5F, FontStyle.Bold),
                TextAlign = ContentAlignment.BottomLeft
            };

            // Add to container in reverse order to dock correctly from top to bottom
            container.Controls.Add(spacer);
            container.Controls.Add(lblError);
            if (showFolderSelect)
            {
                container.Controls.Add(cmbFolder);
                container.Controls.Add(lblFolderPrompt);
                container.Controls.Add(spacerFolder);
            }
            container.Controls.Add(txtWrapper);
            container.Controls.Add(lblPrompt);

            // Bottom Buttons Panel
            Panel btnPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 36
            };
            container.Controls.Add(btnPanel);

            FlowLayoutPanel btnFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Margin = new Padding(0),
                Padding = new Padding(0)
            };
            btnPanel.Controls.Add(btnFlow);

            // Cancel Button
            btnCancel = new Button
            {
                Text = "取消",
                DialogResult = DialogResult.Cancel,
                Width = 80,
                Height = 34,
                Margin = new Padding(10, 0, 0, 0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(224, 224, 224),
                ForeColor = Color.FromArgb(51, 51, 51),
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.FlatAppearance.MouseOverBackColor = Color.FromArgb(200, 200, 200);
            btnFlow.Controls.Add(btnCancel);

            // OK Button
            btnOk = new Button
            {
                Text = "保存",
                DialogResult = DialogResult.None, // Handle click manually for validation
                Width = 80,
                Height = 34,
                Margin = new Padding(0),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 212), // Microsoft Active Blue
                ForeColor = Color.White,
                Cursor = Cursors.Hand,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold)
            };
            btnOk.FlatAppearance.BorderSize = 0;
            btnOk.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 90, 158);
            btnOk.Click += BtnOk_Click;
            btnFlow.Controls.Add(btnOk);

            // Customize TextBox Border Drawing
            txtWrapper.Paint += (s, e) =>
            {
                Color borderColor = txtInput.Focused ? Color.FromArgb(0, 120, 212) : Color.FromArgb(200, 200, 200);
                using (Pen pen = new Pen(borderColor, 1))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, txtWrapper.Width - 1, txtWrapper.Height - 1);
                }
            };

            txtInput.GotFocus += (s, e) => txtWrapper.Invalidate();
            txtInput.LostFocus += (s, e) => txtWrapper.Invalidate();

            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;

            // Focus textbox and select all text
            this.Load += (s, e) =>
            {
                txtInput.Focus();
                if (!string.IsNullOrEmpty(txtInput.Text))
                {
                    txtInput.SelectionStart = 0;
                    txtInput.SelectionLength = txtInput.Text.Length;
                }
            };
        }

        private void PopulateFolders()
        {
            if (cmbFolder == null) return;

            cmbFolder.Items.Clear();
            cmbFolder.Items.Add(new FolderItem { DisplayName = "根目录 (素材库)", FullPath = libraryRoot });

            if (Directory.Exists(libraryRoot))
            {
                try
                {
                    string[] subDirs = Directory.GetDirectories(libraryRoot, "*", SearchOption.AllDirectories);
                    Array.Sort(subDirs);
                    foreach (string dir in subDirs)
                    {
                        string relPath = dir.Substring(libraryRoot.Length).TrimStart(Path.DirectorySeparatorChar);
                        cmbFolder.Items.Add(new FolderItem { DisplayName = relPath, FullPath = dir });
                    }
                }
                catch { }
            }

            // Find currentDir in items
            int selectedIndex = 0;
            for (int i = 0; i < cmbFolder.Items.Count; i++)
            {
                if (((FolderItem)cmbFolder.Items[i]).FullPath == currentDir)
                {
                    selectedIndex = i;
                    break;
                }
            }
            cmbFolder.SelectedIndex = selectedIndex;
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            string val = txtInput.Text.Trim();

            if (string.IsNullOrEmpty(val))
            {
                lblError.Text = "提示：名称不能为空。";
                return;
            }

            // Check for invalid file name characters
            if (Regex.IsMatch(val, @"[\\/:*?""<>|]"))
            {
                lblError.Text = "提示：不能包含 \\ / : * ? \" < > | 字符。";
                return;
            }

            // Update SelectedFolder and libraryDir based on selection before validation
            if (showFolderSelect && cmbFolder != null && cmbFolder.SelectedItem is FolderItem selectedItem)
            {
                this.SelectedFolder = selectedItem.FullPath;
                this.libraryDir = selectedItem.FullPath;
            }
            else
            {
                this.SelectedFolder = currentDir;
            }

            // Check if directory or file already exists in library
            if (isFolder)
            {
                string folderPath = Path.Combine(libraryDir, val);
                if (Directory.Exists(folderPath))
                {
                    lblError.Text = "提示：文件夹名称已存在，请重新输入。";
                    return;
                }
            }
            else
            {
                string pptxPath = Path.Combine(libraryDir, val + ".pptx");
                if (File.Exists(pptxPath))
                {
                    var result = MessageBox.Show(
                        $"已存在名为「{val}」的素材，是否覆盖？",
                        "确认覆盖",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );
                    if (result != DialogResult.Yes)
                    {
                        lblError.Text = "提示：名称已存在，请重新输入。";
                        return;
                    }
                }
            }

            this.InputText = val;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}

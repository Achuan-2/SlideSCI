using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public class AISidebarControl : UserControl
    {
        private FlowLayoutPanel chatFlowLayout;
        
        private Panel inputPanel;
        private TextBox txtInput;
        private Button btnSend;
        private Button btnClear;
        private Button btnSettings;
        private ToolTip toolTip;

        private string configPath;
        private AIConfig currentConfig;
        private List<ChatMessage> conversationHistory = new List<ChatMessage>();
        private static readonly HttpClient httpClient = new HttpClient();
        private bool isGenerating = false;
        private CancellationTokenSource cts;

        public PowerPoint.DocumentWindow AssociatedWindow { get; set; }

        public AISidebarControl()
        {
            // Enable TLS 1.2 and TLS 1.3 for secure API calls
            try
            {
                System.Net.ServicePointManager.SecurityProtocol |= 
                    System.Net.SecurityProtocolType.Tls12 | 
                    System.Net.SecurityProtocolType.Tls11 |
                    (System.Net.SecurityProtocolType)12288; // TLS 1.3
            }
            catch { }

            InitializeConfigPath();
            LoadConfig();
            InitializeComponent();
            InitializeSystemMessage();
        }

        private void InitializeConfigPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string slideSciDir = Path.Combine(appData, "SlideSCI");
            if (!Directory.Exists(slideSciDir))
            {
                Directory.CreateDirectory(slideSciDir);
            }
            configPath = Path.Combine(slideSciDir, "ai_config.json");
        }

        private void LoadConfig()
        {
            if (File.Exists(configPath))
            {
                try
                {
                    string json = File.ReadAllText(configPath);
                    currentConfig = JsonConvert.DeserializeObject<AIConfig>(json);

                    // Migration: If no presets exist but we have old single-config fields, create a default preset
                    if (currentConfig != null && (currentConfig.Presets == null || currentConfig.Presets.Count == 0))
                    {
                        var dict = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);
                        string oldUrl = dict.ContainsKey("ApiUrl") ? dict["ApiUrl"]?.ToString() : null;
                        string oldKey = dict.ContainsKey("ApiKey") ? dict["ApiKey"]?.ToString() : null;
                        string oldModel = dict.ContainsKey("Model") ? dict["Model"]?.ToString() : null;
                        string oldPrompt = dict.ContainsKey("SystemPrompt") ? dict["SystemPrompt"]?.ToString() : null;

                        currentConfig.Presets = new List<AIPreset>();
                        if (!string.IsNullOrEmpty(oldUrl) || !string.IsNullOrEmpty(oldKey))
                        {
                            currentConfig.Presets.Add(new AIPreset
                            {
                                Name = "Default",
                                ApiUrl = oldUrl ?? "https://api.openai.com/v1",
                                ApiKey = oldKey ?? "",
                                Model = oldModel ?? "gpt-4o",
                                SystemPrompt = oldPrompt ?? GetDefaultSystemPrompt()
                            });
                            currentConfig.CurrentPresetName = "Default";
                        }
                    }
                }
                catch
                {
                    currentConfig = new AIConfig();
                }
            }
            else
            {
                currentConfig = new AIConfig();
            }

            if (currentConfig == null) currentConfig = new AIConfig();

            // Ensure we have some default presets if list is still empty
            if (currentConfig.Presets == null || currentConfig.Presets.Count == 0)
            {
                currentConfig.Presets = new List<AIPreset>
                {
                    new AIPreset { Name = "Kimi", ApiUrl = "https://api.moonshot.cn/v1", Model = "kimi-latest", SystemPrompt = GetDefaultSystemPrompt() },
                    new AIPreset { Name = "DeepSeek", ApiUrl = "https://api.deepseek.com/v1", Model = "deepseek-chat", SystemPrompt = GetDefaultSystemPrompt() },
                    new AIPreset { Name = "OpenAI", ApiUrl = "https://api.openai.com/v1", Model = "gpt-4o", SystemPrompt = GetDefaultSystemPrompt() }
                };
                currentConfig.CurrentPresetName = "Kimi";
                SaveConfig();
            }

            if (string.IsNullOrWhiteSpace(currentConfig.SystemPrompt) ||
                currentConfig.SystemPrompt.Contains("Office.MsoShapeType.msoShapeRectangle") ||
                currentConfig.SystemPrompt.Contains("msoShapeRoundRectangle"))
            {
                currentConfig.SystemPrompt = GetDefaultSystemPrompt();
                SaveConfig();
            }
        }

        private void SaveConfig()
        {
            try
            {
                string json = JsonConvert.SerializeObject(currentConfig, Formatting.Indented);
                File.WriteAllText(configPath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存设置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponent()
        {
            this.Dock = DockStyle.Fill;
            this.BackColor = Color.FromArgb(245, 246, 247);
            this.Font = new Font("Microsoft YaHei", 9F, FontStyle.Regular, GraphicsUnit.Point);

            toolTip = new ToolTip();

            // 1. Input Panel (Docked at Bottom)
            inputPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 105,
                BackColor = Color.White,
                Padding = new Padding(10, 8, 10, 8)
            };
            inputPanel.Paint += (s, e) =>
            {
                using (Pen pen = new Pen(Color.FromArgb(225, 225, 225), 1))
                {
                    e.Graphics.DrawLine(pen, 0, 0, inputPanel.Width, 0);
                }
            };

            // Custom border panel for textbox
            Panel txtWrapper = new Panel
            {
                Location = new Point(10, 8),
                Width = this.Width - 20,
                Height = 52,
                BackColor = Color.White,
                Padding = new Padding(4)
            };
            txtWrapper.Paint += (s, e) =>
            {
                Color borderColor = txtInput.Focused ? Color.FromArgb(0, 120, 212) : Color.FromArgb(220, 220, 220);
                using (Pen pen = new Pen(borderColor, 1))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, txtWrapper.Width - 1, txtWrapper.Height - 1);
                }
            };

            txtInput = new TextBox
            {
                Multiline = true,
                BorderStyle = BorderStyle.None,
                Dock = DockStyle.Fill,
                Font = new Font("Microsoft YaHei", 9F),
                ScrollBars = ScrollBars.Vertical
            };
            txtInput.GotFocus += (s, e) => txtWrapper.Invalidate();
            txtInput.LostFocus += (s, e) => txtWrapper.Invalidate();
            txtInput.KeyDown += TxtInput_KeyDown;
            txtWrapper.Controls.Add(txtInput);
            inputPanel.Controls.Add(txtWrapper);

            // Icon-only Clear button
            btnClear = new Button
            {
                Image = CreateIcon("trash", 16, Color.FromArgb(100, 100, 100)),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                Location = new Point(10, 68),
                Width = 28,
                Height = 26,
                Cursor = Cursors.Hand
            };
            btnClear.FlatAppearance.BorderSize = 0;
            btnClear.FlatAppearance.MouseOverBackColor = Color.FromArgb(240, 240, 240);
            btnClear.Click += (s, e) => ClearChat();
            toolTip.SetToolTip(btnClear, "清空对话");
            inputPanel.Controls.Add(btnClear);

            // Icon-only Settings button
            btnSettings = new Button
            {
                Image = CreateIcon("settings", 16, Color.FromArgb(100, 100, 100)),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Transparent,
                Location = new Point(44, 68),
                Width = 28,
                Height = 26,
                Cursor = Cursors.Hand
            };
            btnSettings.FlatAppearance.BorderSize = 0;
            btnSettings.FlatAppearance.MouseOverBackColor = Color.FromArgb(240, 240, 240);
            btnSettings.Click += BtnSettings_Click;
            toolTip.SetToolTip(btnSettings, "AI 设置");
            inputPanel.Controls.Add(btnSettings);

            btnSend = new Button
            {
                Text = "发送",
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold),
                Location = new Point(this.Width - 75, 68),
                Width = 65,
                Height = 26,
                Cursor = Cursors.Hand
            };
            btnSend.FlatAppearance.BorderSize = 0;
            btnSend.Click += async (s, e) => await HandleSendAsync();
            inputPanel.Controls.Add(btnSend);

            this.Controls.Add(inputPanel);

            // 2. Chat Flow Layout Panel (Docked Fill)
            chatFlowLayout = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Padding = new Padding(8),
                BackColor = Color.FromArgb(245, 246, 247)
            };
            typeof(FlowLayoutPanel).InvokeMember("DoubleBuffered",
                System.Reflection.BindingFlags.SetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                null, chatFlowLayout, new object[] { true });

            this.Controls.Add(chatFlowLayout);

            // Ensure Z-order prevents bottom layout overlapping
            chatFlowLayout.BringToFront();

            this.Resize += AISidebarControl_Resize;
            chatFlowLayout.Resize += (s, e) => UpdateBubbleWidths();
        }

        private void AISidebarControl_Resize(object sender, EventArgs e)
        {
            if (inputPanel != null)
            {
                foreach (Control ctrl in inputPanel.Controls)
                {
                    if (ctrl is Panel p) // txtWrapper
                    {
                        p.Width = inputPanel.Width - 20;
                    }
                    else if (ctrl == btnSend)
                    {
                        btnSend.Left = inputPanel.Width - 75;
                    }
                }
            }
            UpdateBubbleWidths();
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            using (var settingsForm = new AISettingsForm(currentConfig))
            {
                if (settingsForm.ShowDialog(this) == DialogResult.OK)
                {
                    settingsForm.UpdateConfig(currentConfig);
                    SaveConfig();

                    // Update the active system message in the conversation history
                    if (conversationHistory.Count > 0 && conversationHistory[0].Role == "system")
                    {
                        conversationHistory[0].Content = currentConfig.SystemPrompt;
                    }
                }
            }
        }

        private void InitializeSystemMessage()
        {
            string systemPrompt = currentConfig.SystemPrompt;
            if (string.IsNullOrWhiteSpace(systemPrompt))
            {
                systemPrompt = GetDefaultSystemPrompt();
                currentConfig.SystemPrompt = systemPrompt;
            }

            conversationHistory.Add(new ChatMessage { Role = "system", Content = systemPrompt });
            AppendBubble("你好！我是你的 SlideSCI AI 助手。我可以执行以下命令：\n1. 把选中的文字或文本框颜色改为红色、蓝色等。\n2. 添加各种形状（矩形、圆形等）。\n3. 新增一页空白幻灯片。\n4. 将幻灯片中的指定文字进行查找替换。\n\n请在下方输入你的指令！", isUser: false, isSystem: false);
        }

        public static string GetDefaultSystemPrompt()
        {
            return @"你是一个PowerPoint AI助手，你可以通过编写和执行C#代码来帮用户操作当前幻灯片。
你有调用 `execute_csharp_code` 工具的能力。
当你需要对幻灯片进行操作时（如修改文字颜色、添加形状、新建页面、替换文字等），你应该写C#代码，然后通过调用 `execute_csharp_code` 来执行。

编写C#代码时请遵守以下规则：
1. 代码会在一个包含 `PowerPoint.Application Application` 参数的方法体内执行：`public static void Run(PowerPoint.Application Application)`。
2. 已经默认导入了以下命名空间，请勿再在代码中写 using 声明：
   - System
   - System.Collections.Generic
   - System.Linq
   - System.Drawing
   - System.Windows.Forms
   - Microsoft.Office.Interop.PowerPoint (别名为 PowerPoint)
   - Microsoft.Office.Core (别名为 Office)
3. 安全获取当前幻灯片对象 (Slide) 的方法：
   PowerPoint.Slide slide = null;
   try { slide = Application.ActiveWindow.View.Slide as PowerPoint.Slide; } catch {}
   if (slide == null) {
       try { slide = Application.ActivePresentation.Slides[Application.ActiveWindow.Selection.SlideRange[1].SlideIndex]; } catch {}
   }
   if (slide == null) {
       throw new Exception(""没有找到活动的幻灯片，请确保PPT处于正常视图。"");
   }
4. 新增幻灯片页面请直接对 Application.ActivePresentation.Slides 进行 Add 或 Add2 操作，例如：
   var pres = Application.ActivePresentation;
   pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
5. 操作文字颜色时，请使用 `System.Drawing.ColorTranslator.ToOle(Color.Red)` 将颜色转换为 OLE 颜色。例如：
   shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Red);
6. 插入形状时，第一个参数必须是 `Office.MsoAutoShapeType` 枚举类型（注意：千万不可写错成 `Office.MsoShapeType`，否则会报编译错误）。常用的形状枚举如下：
   - 矩形：`Office.MsoAutoShapeType.msoShapeRectangle`
   - 圆角矩形（注意是 Rounded，不是 Round）：`Office.MsoAutoShapeType.msoShapeRoundedRectangle`
   - 椭圆/圆形：`Office.MsoAutoShapeType.msoShapeOval`
   示例：`slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height)`。位置与尺寸参数皆为 float 类型。
7. 修改文字内容时，如果需要查找并替换，可以遍历幻灯片中的 Shape 并检查 `shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue`。
8. 请尽量编写健壮的代码，用 try-catch 包裹可能出错的地方，遇到不支持的操作时抛出异常。

【核心PPT设计审美与排版规范（PPT Aesthetics & Layout Design Rules）】
为了确保AI添加的PPT样式美观、专业且符合现代高端设计规范，请在编写代码生成/修改PPT内容时，严格遵守以下排版与审美原则：

1. 配色系统与 60-30-10 黄金比例：
   - 禁止在没有用户明确要求的情况下使用刺眼的高饱和度纯原色（例如纯红 Color.Red、纯蓝 Color.Blue、纯绿 Color.Green、纯黄 Color.Yellow）。这会让 PPT 显得极不专业。
   - 使用成熟、高对比度且优雅的经典主题配色（使用 Color.FromArgb 定义颜色，然后使用 ColorTranslator.ToOle 转换）：
     * 极简科技蓝主题：
       - 主背景/主色：浅灰白 Color.FromArgb(248, 249, 250) 或 极深蓝 Color.FromArgb(10, 34, 64)
       - 辅助色/卡片背景：浅天蓝 Color.FromArgb(232, 240, 254) 或 经典蓝 Color.FromArgb(26, 115, 232)
       - 点缀强调色（10%）：温暖橙/金色 Color.FromArgb(245, 166, 35) 或 珊瑚红 Color.FromArgb(234, 67, 53)
     * 现代商务炭黑主题：
       - 主色：炭黑色 Color.FromArgb(43, 43, 43)
       - 辅助色：浅灰色 Color.FromArgb(240, 240, 240)
       - 点缀色：青/湖蓝色 Color.FromArgb(0, 168, 181)
     * 暖调人文/学术主题：
       - 主色：深林绿 Color.FromArgb(27, 67, 50)
       - 辅助色：柔和米色 Color.FromArgb(245, 242, 235)
       - 点缀色：赤陶红/土黄 Color.FromArgb(190, 130, 60)
   - 遵守 60-30-10 比例：背景占 60% 视觉面积，文字/卡片容器等次要元素占 30%，画龙点睛的重点高亮（如关键数字、重点结论词、边框线装饰）占 10%。

2. 视觉层级与字号梯度（Visual Hierarchy & Text Scales）：
   - 建立清晰鲜明的字号差，中文默认使用“微软雅黑” (Microsoft YaHei)，英文默认使用“Segoe UI”：
     * 幻灯片大标题（Slide Title）：36pt - 44pt，粗体 (Bold)，深色高对比度。
     * 章节/卡片小标题（Card Header）：20pt - 24pt，粗体。
     * 正文内容（Body Text）：14pt - 18pt，常规体 (Regular)，中等对比度（如深灰）。
     * 关键数据/大数字（Big Stat/Metric）：48pt - 72pt 粗体，可搭配 12pt 的解释性小字。
   - 为确保文字排版完美，对所有新创建的文本框设置以下属性以防文本溢出或不对齐：
     * shape.TextFrame.WordWrap = Office.MsoTriState.msoTrue; // 开启自动换行
     * 适当清除或缩小默认内边距：
       shape.TextFrame.MarginLeft = 10f;
       shape.TextFrame.MarginRight = 10f;
       shape.TextFrame.MarginTop = 10f;
       shape.TextFrame.MarginBottom = 10f;

3. 负空间与留白（Negative Space & Whitespace）：
   - 不要把幻灯片塞满！页面必须保留至少 30% - 40% 的空白区域，让观众的眼睛有呼吸的空间。
   - 绝不允许使用大段文字，请将文字提炼为短句，以列表或卡片形式展现。

4. 动态网格布局与防重叠（Dynamic Grid Layout & Math Positioning）：
   - 永远通过代码动态获取幻灯片的物理大小，并在此基础上进行百分比定位或网格分栏计算，严禁使用写死的绝对坐标导致在不同比例的PPT中排版错乱：
     float slideWidth = Application.ActivePresentation.PageSetup.SlideWidth;
     float slideHeight = Application.ActivePresentation.PageSetup.SlideHeight;
   - 留出四周安全边距（Margin，通常设为 slideWidth * 0.06f 与 slideHeight * 0.08f）。
   - 如果需要展示并列的多个观点，应自动计算分栏位置：
     * 两栏布局：左栏占宽 45%，右栏占宽 45%，中间留空 10%。
     * 三分栏卡片布局示例：
       float margin = slideWidth * 0.05f;
       float availableWidth = slideWidth - 2 * margin;
       float gap = 20f; // 卡片间距
       float cardWidth = (availableWidth - 2 * gap) / 3f;
       float cardHeight = slideHeight * 0.5f;
       float top = slideHeight * 0.3f;
       for (int i = 0; i < 3; i++) {
           float left = margin + i * (cardWidth + gap);
           // 在 left, top, cardWidth, cardHeight 处绘制卡片容器
       }
   - 通过数学计算绝对保证生成的文本框、容器形状之间不会重叠。

5. 扁平化容器卡片化（Flat Card-Based Design）：
   - 列表或并列观点应使用“浅色卡片”作为容器装载，卡片背景填充非常浅的颜色（如 RGB: 245, 245, 245），并且必须去掉边框轮廓以符合现代扁平化审美：
     cardShape.Line.Visible = Office.MsoTriState.msoFalse;
     cardShape.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(245, 245, 245));
   - 在卡片之上叠放文字形状时，通过 Y 坐标偏移（如 cardTop + 15 放置标题，cardTop + 45 放置正文）使排版整洁大方。
   - 尽量使用无轮廓（`Line.Visible = msoFalse`）的形状，如果使用线条，应保持极细（`Line.Weight = 1.0f`）且使用温和的淡灰色。";
        }

        private void TxtInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true; // Prevent newline
                btnSend.PerformClick();
            }
        }

        private void ClearChat()
        {
            chatFlowLayout.Controls.Clear();
            conversationHistory.Clear();
            InitializeSystemMessage();
        }

        private void AppendBubble(string text, bool isUser, bool isSystem, bool isError = false)
        {
            // Strip emojis from the text to prevent square box characters in RichTextBox
            text = StripEmojis(text);

            ChatRowPanel row = new ChatRowPanel(text, isUser, isSystem, isError, chatFlowLayout.Width - 25);
            chatFlowLayout.Controls.Add(row);
            
            chatFlowLayout.PerformLayout();
            chatFlowLayout.ScrollControlIntoView(row);
            
            try
            {
                chatFlowLayout.VerticalScroll.Value = chatFlowLayout.VerticalScroll.Maximum;
            }
            catch {}
        }

        private static string StripEmojis(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            try
            {
                // Remove surrogate pairs and symbol ranges (emojis, icons, private use areas)
                return Regex.Replace(text, @"\p{Cs}|\p{So}|\p{Cn}", "");
            }
            catch
            {
                return text;
            }
        }

        private void UpdateBubbleWidths()
        {
            if (chatFlowLayout == null) return;
            chatFlowLayout.SuspendLayout();
            int targetWidth = chatFlowLayout.Width - 25;
            if (targetWidth < 100) targetWidth = 100;
            foreach (Control ctrl in chatFlowLayout.Controls)
            {
                if (ctrl is ChatRowPanel row)
                {
                    row.UpdateWidth(targetWidth);
                }
            }
            chatFlowLayout.ResumeLayout(true);
        }

        private async Task HandleSendAsync()
        {
            if (isGenerating)
            {
                cts?.Cancel();
                return;
            }

            string message = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(message)) return;

            if (string.IsNullOrEmpty(currentConfig.ApiKey))
            {
                MessageBox.Show("请先配置 API Key！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                BtnSettings_Click(null, null);
                return;
            }

            txtInput.Text = "";
            AppendBubble(message, isUser: true, isSystem: false);
            conversationHistory.Add(new ChatMessage { Role = "user", Content = message });

            isGenerating = true;
            btnSend.Text = "停止";
            btnSend.BackColor = Color.FromArgb(209, 52, 56);
            txtInput.Enabled = false;
            btnClear.Enabled = false;
            btnSettings.Enabled = false;

            cts = new CancellationTokenSource();

            try
            {
                await RunAgentLoopAsync(cts.Token);
            }
            catch (OperationCanceledException)
            {
                AppendBubble("[已终止]: 用户中止了 AI 响应。", isUser: false, isSystem: true, isError: false);
            }
            catch (Exception ex)
            {
                AppendBubble($"[AI 调用出错]:\n{ex.ToString()}", isUser: false, isSystem: true, isError: true);
            }
            finally
            {
                isGenerating = false;
                btnSend.Text = "发送";
                btnSend.BackColor = Color.FromArgb(0, 120, 212);
                
                txtInput.Enabled = true;
                btnClear.Enabled = true;
                btnSettings.Enabled = true;

                cts?.Dispose();
                cts = null;
            }
        }

        private async Task RunAgentLoopAsync(CancellationToken token)
        {
            int maxIterations = 5;
            int currentIteration = 0;

            var tools = new List<ChatTool>
            {
                new ChatTool
                {
                    Function = new ToolFunction
                    {
                        Name = "execute_csharp_code",
                        Description = "在幻灯片中动态编译并执行C#代码来操作系统。代码将被放置在 `public static void Run(PowerPoint.Application Application)` 方法体中运行。你可以使用任何 PowerPoint API 进行操作。",
                        Parameters = new ToolParameters
                        {
                            Properties = new Dictionary<string, object>
                            {
                                {
                                    "code", new Dictionary<string, string>
                                    {
                                        { "type", "string" },
                                        { "description", "要执行的 C# 语句。必须使用合法的 C# 语法，不用声明 class 或 method，只写方法体内部的执行代码。" }
                                    }
                                }
                            },
                            Required = new List<string> { "code" }
                        }
                    }
                }
            };

            while (currentIteration < maxIterations)
            {
                token.ThrowIfCancellationRequested();
                currentIteration++;

                var request = new ChatRequest
                {
                    Model = currentConfig.Model,
                    Messages = conversationHistory,
                    Tools = tools,
                    ToolChoice = "auto"
                };

                string requestJson = JsonConvert.SerializeObject(request);
                var content = new StringContent(requestJson, Encoding.UTF8, "application/json");

                httpClient.DefaultRequestHeaders.Clear();
                httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {currentConfig.ApiKey}");

                string url = currentConfig.ApiUrl.TrimEnd('/') + "/chat/completions";
                
                HttpResponseMessage response;
                try
                {
                    response = await httpClient.PostAsync(url, content, token);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    throw new Exception($"无法连接到 AI 服务，网络连接异常: {ex.Message}\n请求地址: {url}\n\n详细信息:\n{ex.ToString()}");
                }

                token.ThrowIfCancellationRequested();
                string responseJson = await response.Content.ReadAsStringAsync();
                if (!response.IsSuccessStatusCode)
                {
                    string errorDetails = "";
                    try
                    {
                        var errRes = JsonConvert.DeserializeObject<ChatResponse>(responseJson);
                        errorDetails = errRes?.Error?.Message;
                    }
                    catch {}
                    if (string.IsNullOrEmpty(errorDetails))
                    {
                        errorDetails = responseJson;
                    }
                    throw new Exception($"API 请求失败 (HTTP {response.StatusCode} - {response.ReasonPhrase}):\n{errorDetails}");
                }

                var chatRes = JsonConvert.DeserializeObject<ChatResponse>(responseJson);
                if (chatRes == null || chatRes.Choices == null || chatRes.Choices.Count == 0)
                {
                    throw new Exception("API 返回了空响应，无法处理。");
                }

                var choiceMessage = chatRes.Choices[0].Message;
                conversationHistory.Add(choiceMessage);

                if (!string.IsNullOrEmpty(choiceMessage.Content))
                {
                    AppendBubble(choiceMessage.Content, isUser: false, isSystem: false);
                }

                if (choiceMessage.ToolCalls != null && choiceMessage.ToolCalls.Count > 0)
                {
                    token.ThrowIfCancellationRequested();
                    foreach (var toolCall in choiceMessage.ToolCalls)
                    {
                        if (toolCall.Function.Name == "execute_csharp_code")
                        {
                            string code = "";
                            try
                            {
                                var args = JsonConvert.DeserializeObject<Dictionary<string, string>>(toolCall.Function.Arguments);
                                if (args != null && args.TryGetValue("code", out string argCode))
                                {
                                    code = argCode;
                                }
                            }
                            catch (Exception ex)
                            {
                                string toolErr = $"解析参数失败: {ex.Message}";
                                conversationHistory.Add(new ChatMessage
                                {
                                    Role = "tool",
                                    Name = "execute_csharp_code",
                                    Content = toolErr,
                                    ToolCallId = toolCall.Id
                                });
                                AppendBubble($"[工具调用解析失败]\n{toolErr}", isUser: false, isSystem: true, isError: true);
                                continue;
                            }

                            AppendBubble($"[AI 请求执行 C# 代码]:\n{code}", isUser: false, isSystem: true);

                            string runResult = "";
                            try
                            {
                                runResult = CSharpExecutor.Execute(code, Globals.ThisAddIn.Application);
                            }
                            catch (Exception ex)
                            {
                                runResult = $"Critical Runtime Error: {ex.Message}";
                            }

                            conversationHistory.Add(new ChatMessage
                            {
                                Role = "tool",
                                Name = "execute_csharp_code",
                                Content = runResult,
                                ToolCallId = toolCall.Id
                            });

                            if (runResult == "Success")
                            {
                                AppendBubble("[执行结果]: 成功！", isUser: false, isSystem: true);
                            }
                            else
                            {
                                AppendBubble($"[执行结果]: 失败！\n{runResult}", isUser: false, isSystem: true, isError: true);
                            }
                        }
                    }
                    continue;
                }

                break;
            }
        }

        public static Bitmap CreateIcon(string type, int size, Color color)
        {
            Bitmap bmp = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                using (Pen pen = new Pen(color, 1.8f))
                using (Brush brush = new SolidBrush(color))
                {
                    if (type == "settings")
                    {
                        int r1 = size / 2 - 1;
                        int r2 = size / 2 - 4;
                        int cx = size / 2;
                        int cy = size / 2;
                        g.DrawEllipse(pen, cx - r2, cy - r2, r2 * 2, r2 * 2);
                        
                        for (int i = 0; i < 8; i++)
                        {
                            double angle = i * Math.PI / 4;
                            float x1 = (float)(cx + r2 * Math.Cos(angle));
                            float y1 = (float)(cy + r2 * Math.Sin(angle));
                            float x2 = (float)(cx + r1 * Math.Cos(angle));
                            float y2 = (float)(cy + r1 * Math.Sin(angle));
                            g.DrawLine(pen, x1, y1, x2, y2);
                        }
                    }
                    else if (type == "trash")
                    {
                        g.DrawRectangle(pen, 4, 5, size - 9, size - 8);
                        g.DrawLine(pen, 2, 3, size - 3, 3);
                        g.DrawLine(pen, 6, 1, size - 7, 1);
                        g.DrawLine(pen, 7, 7, 7, size - 5);
                        g.DrawLine(pen, size / 2, 7, size / 2, size - 5);
                        g.DrawLine(pen, size - 8, 7, size - 8, size - 5);
                    }
                }
            }
            return bmp;
        }
    }

    // --- Settings Dialog Form ---
    public class AISettingsForm : Form
    {
        private ComboBox cmbPresets;
        private Button btnAddPreset;
        private Button btnDeletePreset;
        private TextBox txtPresetName;
        private TextBox txtApiUrl;
        private TextBox txtApiKey;
        private CheckBox chkShowKey;
        private TextBox txtModel;
        private TextBox txtSystemPrompt;
        private Button btnSave;
        private Button btnCancel;
        private Button btnRestoreDefault;
        private AIConfig config;
        private List<AIPreset> tempPresets;
        private string activePresetName;
        private AIPreset currentEditingPreset = null;

        public AISettingsForm(AIConfig currentConfig)
        {
            this.config = currentConfig;
            this.tempPresets = new List<AIPreset>();
            foreach (var p in currentConfig.Presets)
                tempPresets.Add(new AIPreset { Name = p.Name, ApiUrl = p.ApiUrl, ApiKey = p.ApiKey, Model = p.Model, SystemPrompt = p.SystemPrompt });
            this.activePresetName = currentConfig.CurrentPresetName;
            InitializeComponent();
            LoadPresetsToCombo();
        }

        private void InitializeComponent()
        {
            this.Text = "AI 助手设置";
            this.Size = new Size(620, 580);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Microsoft YaHei", 9F);
            this.BackColor = Color.White;

            Label lblSelectPreset = new Label { Text = "选择预设:", Location = new Point(15, 20), Width = 80, Height = 22, TextAlign = ContentAlignment.MiddleLeft };
            cmbPresets = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(100, 20), Width = 280, Height = 24 };

            btnAddPreset = new Button { Text = "新增", Location = new Point(390, 18), Width = 90, Height = 28, BackColor = Color.FromArgb(240, 240, 240), FlatStyle = FlatStyle.Flat };
            btnAddPreset.FlatAppearance.BorderColor = Color.LightGray;
            btnDeletePreset = new Button { Text = "删除", Location = new Point(490, 18), Width = 90, Height = 28, BackColor = Color.FromArgb(240, 240, 240), FlatStyle = FlatStyle.Flat };
            btnDeletePreset.FlatAppearance.BorderColor = Color.LightGray;

            Label lblPresetName = new Label { Text = "预设名称:", Location = new Point(15, 60), Width = 80, Height = 22, TextAlign = ContentAlignment.MiddleLeft };
            txtPresetName = new TextBox { Location = new Point(100, 60), Width = 480 };
            txtPresetName.TextChanged += TxtPresetName_TextChanged;

            Label lblUrl = new Label { Text = "API 地址:", Location = new Point(15, 100), Width = 80, Height = 22, TextAlign = ContentAlignment.MiddleLeft };
            txtApiUrl = new TextBox { Location = new Point(100, 100), Width = 480 };

            Label lblKey = new Label { Text = "API Key:", Location = new Point(15, 140), Width = 80, Height = 22, TextAlign = ContentAlignment.MiddleLeft };
            txtApiKey = new TextBox { Location = new Point(100, 140), Width = 400, PasswordChar = '*' };
            chkShowKey = new CheckBox { Text = "显示", Location = new Point(510, 140), Width = 70, Height = 22, TextAlign = ContentAlignment.MiddleLeft };
            chkShowKey.CheckedChanged += (s, e) => { txtApiKey.PasswordChar = chkShowKey.Checked ? '\0' : '*'; };

            Label lblModel = new Label { Text = "模型名称:", Location = new Point(15, 180), Width = 80, Height = 22, TextAlign = ContentAlignment.MiddleLeft };
            txtModel = new TextBox { Location = new Point(100, 180), Width = 480 };

            Label lblPrompt = new Label { Text = "系统提示词:", Location = new Point(15, 220), Width = 85, Height = 22, TextAlign = ContentAlignment.MiddleLeft };
            txtSystemPrompt = new TextBox { Location = new Point(100, 220), Width = 480, Height = 240, Multiline = true, ScrollBars = ScrollBars.Vertical, Font = new Font("Microsoft YaHei", 9F) };

            btnRestoreDefault = new Button { Text = "恢复默认提示词", Location = new Point(100, 485), Width = 150, Height = 28, FlatStyle = FlatStyle.Flat };
            btnRestoreDefault.FlatAppearance.BorderColor = Color.LightGray;
            btnRestoreDefault.Click += (s, e) => { txtSystemPrompt.Text = AISidebarControl.GetDefaultSystemPrompt(); };

            btnSave = new Button { Text = "保存", DialogResult = DialogResult.OK, Location = new Point(380, 485), Width = 90, Height = 28, BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnSave.FlatAppearance.BorderSize = 0;
            btnCancel = new Button { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(490, 485), Width = 90, Height = 28, FlatStyle = FlatStyle.Flat };
            btnCancel.FlatAppearance.BorderColor = Color.LightGray;

            this.Controls.AddRange(new Control[] {
                lblSelectPreset, cmbPresets, btnAddPreset, btnDeletePreset,
                lblPresetName, txtPresetName, lblUrl, txtApiUrl,
                lblKey, txtApiKey, chkShowKey, lblModel, txtModel,
                lblPrompt, txtSystemPrompt, btnRestoreDefault, btnSave, btnCancel
            });
            this.AcceptButton = btnSave;
            this.CancelButton = btnCancel;
            btnAddPreset.Click += BtnAddPreset_Click;
            btnDeletePreset.Click += BtnDeletePreset_Click;
        }

        private void LoadPresetsToCombo()
        {
            cmbPresets.SelectedIndexChanged -= CmbPresets_SelectedIndexChanged;
            cmbPresets.DataSource = null;
            cmbPresets.DataSource = tempPresets;
            cmbPresets.DisplayMember = "Name";
            int idx = tempPresets.FindIndex(p => p.Name == activePresetName);
            if (idx < 0 && tempPresets.Count > 0) idx = 0;
            cmbPresets.SelectedIndex = idx;
            if (idx >= 0) { LoadPresetToFields(tempPresets[idx]); activePresetName = tempPresets[idx].Name; }
            cmbPresets.SelectedIndexChanged += CmbPresets_SelectedIndexChanged;
        }

        private void SaveCurrentPresetEdits()
        {
            if (currentEditingPreset != null)
            {
                currentEditingPreset.Name = txtPresetName.Text.Trim();
                currentEditingPreset.ApiUrl = txtApiUrl.Text.Trim();
                currentEditingPreset.ApiKey = txtApiKey.Text.Trim();
                currentEditingPreset.Model = txtModel.Text.Trim();
                currentEditingPreset.SystemPrompt = txtSystemPrompt.Text;
            }
        }

        private void LoadPresetToFields(AIPreset preset)
        {
            currentEditingPreset = preset;
            if (preset != null)
            {
                txtPresetName.Text = preset.Name;
                txtApiUrl.Text = preset.ApiUrl;
                txtApiKey.Text = preset.ApiKey;
                txtModel.Text = preset.Model;
                txtSystemPrompt.Text = preset.SystemPrompt;
            }
        }

        private void CmbPresets_SelectedIndexChanged(object sender, EventArgs e)
        {
            SaveCurrentPresetEdits();
            if (cmbPresets.SelectedIndex >= 0)
            {
                var selected = cmbPresets.SelectedItem as AIPreset;
                LoadPresetToFields(selected);
                this.activePresetName = selected?.Name ?? "";
            }
        }

        private void TxtPresetName_TextChanged(object sender, EventArgs e)
        {
            if (currentEditingPreset != null && cmbPresets.SelectedItem == currentEditingPreset)
            {
                string newName = txtPresetName.Text.Trim();
                if (!string.IsNullOrEmpty(newName) && currentEditingPreset.Name != newName)
                {
                    currentEditingPreset.Name = newName;
                    int idx = cmbPresets.SelectedIndex;
                    cmbPresets.SelectedIndexChanged -= CmbPresets_SelectedIndexChanged;
                    cmbPresets.DataSource = null;
                    cmbPresets.DataSource = tempPresets;
                    cmbPresets.DisplayMember = "Name";
                    cmbPresets.SelectedIndex = idx;
                    cmbPresets.SelectedIndexChanged += CmbPresets_SelectedIndexChanged;
                }
            }
        }

        private void BtnAddPreset_Click(object sender, EventArgs e)
        {
            SaveCurrentPresetEdits();
            int count = 1;
            string newName = "New Preset";
            while (tempPresets.Exists(p => p.Name == $"{newName} {count}")) count++;
            string finalName = $"{newName} {count}";
            var newPreset = new AIPreset { Name = finalName, ApiUrl = "https://api.openai.com/v1", ApiKey = "", Model = "gpt-4o", SystemPrompt = AISidebarControl.GetDefaultSystemPrompt() };
            tempPresets.Add(newPreset);
            activePresetName = finalName;
            LoadPresetsToCombo();
        }

        private void BtnDeletePreset_Click(object sender, EventArgs e)
        {
            if (tempPresets.Count <= 1) { MessageBox.Show("必须保留至少一个预设！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (cmbPresets.SelectedItem is AIPreset selected)
            {
                if (MessageBox.Show($"确定删除预设 '{selected.Name}' 吗？", "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    tempPresets.Remove(selected);
                    currentEditingPreset = null;
                    activePresetName = tempPresets[0].Name;
                    LoadPresetsToCombo();
                }
            }
        }

        public void UpdateConfig(AIConfig targetConfig)
        {
            SaveCurrentPresetEdits();
            targetConfig.Presets = this.tempPresets;
            targetConfig.CurrentPresetName = this.activePresetName;
        }
    }

    // --- Data Models for API ---
    public class AIPreset
    {
        public string Name { get; set; } = "Default";
        public string ApiUrl { get; set; } = "https://api.openai.com/v1";
        public string ApiKey { get; set; } = "";
        public string Model { get; set; } = "gpt-4o";
        public string SystemPrompt { get; set; } = "";
        public override string ToString() => Name;
    }

    public class AIConfig
    {
        public List<AIPreset> Presets { get; set; } = new List<AIPreset>();
        public string CurrentPresetName { get; set; } = "";

        [JsonIgnore]
        public AIPreset CurrentPreset
        {
            get
            {
                if (Presets == null || Presets.Count == 0)
                {
                    Presets = new List<AIPreset> { new AIPreset { Name = "Default" } };
                    CurrentPresetName = "Default";
                }
                var active = Presets.Find(p => p.Name == CurrentPresetName);
                if (active == null) { active = Presets[0]; CurrentPresetName = active.Name; }
                return active;
            }
        }

        [JsonIgnore] public string ApiUrl { get => CurrentPreset.ApiUrl; set => CurrentPreset.ApiUrl = value; }
        [JsonIgnore] public string ApiKey { get => CurrentPreset.ApiKey; set => CurrentPreset.ApiKey = value; }
        [JsonIgnore] public string Model { get => CurrentPreset.Model; set => CurrentPreset.Model = value; }
        [JsonIgnore] public string SystemPrompt { get => CurrentPreset.SystemPrompt; set => CurrentPreset.SystemPrompt = value; }
    }

    public class ChatRequest
    {
        [JsonProperty("model")]
        public string Model { get; set; }
        [JsonProperty("messages")]
        public List<ChatMessage> Messages { get; set; }
        [JsonProperty("tools", NullValueHandling = NullValueHandling.Ignore)]
        public List<ChatTool> Tools { get; set; }
        [JsonProperty("tool_choice", NullValueHandling = NullValueHandling.Ignore)]
        public string ToolChoice { get; set; }
    }

    public class ChatMessage
    {
        [JsonProperty("role")]
        public string Role { get; set; }
        [JsonProperty("content", NullValueHandling = NullValueHandling.Ignore)]
        public string Content { get; set; }
        [JsonProperty("reasoning_content", NullValueHandling = NullValueHandling.Ignore)]
        public string ReasoningContent { get; set; }
        [JsonProperty("tool_calls", NullValueHandling = NullValueHandling.Ignore)]
        public List<ToolCall> ToolCalls { get; set; }
        [JsonProperty("tool_call_id", NullValueHandling = NullValueHandling.Ignore)]
        public string ToolCallId { get; set; }
        [JsonProperty("name", NullValueHandling = NullValueHandling.Ignore)]
        public string Name { get; set; }
    }

    public class ChatTool
    {
        [JsonProperty("type")]
        public string Type { get; set; } = "function";
        [JsonProperty("function")]
        public ToolFunction Function { get; set; }
    }

    public class ToolFunction
    {
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("description")]
        public string Description { get; set; }
        [JsonProperty("parameters")]
        public ToolParameters Parameters { get; set; }
    }

    public class ToolParameters
    {
        [JsonProperty("type")]
        public string Type { get; set; } = "object";
        [JsonProperty("properties")]
        public Dictionary<string, object> Properties { get; set; }
        [JsonProperty("required")]
        public List<string> Required { get; set; }
    }

    public class ToolCall
    {
        [JsonProperty("id")]
        public string Id { get; set; }
        [JsonProperty("type")]
        public string Type { get; set; }
        [JsonProperty("function")]
        public ToolCallFunction Function { get; set; }
    }

    public class ToolCallFunction
    {
        [JsonProperty("name")]
        public string Name { get; set; }
        [JsonProperty("arguments")]
        public string Arguments { get; set; }
    }

    public class ChatResponse
    {
        [JsonProperty("choices")]
        public List<ChatChoice> Choices { get; set; }
        [JsonProperty("error")]
        public ApiError Error { get; set; }
    }

    public class ChatChoice
    {
        [JsonProperty("message")]
        public ChatMessage Message { get; set; }
    }

    public class ApiError
    {
        [JsonProperty("message")]
        public string Message { get; set; }
    }

    // --- Custom UI Controls ---
    public class ChatRowPanel : Panel
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool HideCaret(IntPtr hWnd);

        private Panel bubblePanel;
        private RichTextBox rtbText;
        private bool isUser;
        private bool isSystem;
        private bool isError;

        // Collapsible properties
        private bool isCollapsible = false;
        private bool isExpanded = false;
        private string headerText = "";
        private string bodyText = "";
        private Panel headerPanel;
        private Button btnToggle;

        public ChatRowPanel(string text, bool isUser, bool isSystem, bool isError, int width)
        {
            this.isUser = isUser;
            this.isSystem = isSystem;
            this.isError = isError;
            this.Width = width;
            this.BackColor = Color.Transparent;

            // Check if it is a collapsible code block or output log
            if (isSystem && (text.StartsWith("[AI 请求执行 C# 代码]:") || text.StartsWith("[执行结果]:")))
            {
                int firstNewLine = text.IndexOf('\n');
                if (firstNewLine >= 0)
                {
                    headerText = text.Substring(0, firstNewLine).TrimEnd(':');
                    bodyText = text.Substring(firstNewLine + 1);
                    isCollapsible = !string.IsNullOrWhiteSpace(bodyText);
                }
                else
                {
                    headerText = text;
                    bodyText = "";
                    isCollapsible = false;
                }
                isExpanded = false; // Collapsed by default
            }

            bubblePanel = new Panel
            {
                BackColor = Color.Transparent,
                Padding = new Padding(10, 8, 10, 8),
                AutoSize = false
            };

            rtbText = new RichTextBox
            {
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.None,
                ReadOnly = true,
                Multiline = true,
                WordWrap = true,
                BackColor = Color.White,
                Font = new Font("Microsoft YaHei", 9F)
            };

            rtbText.GotFocus += (s, e) => { HideCaret(rtbText.Handle); };
            rtbText.MouseDown += (s, e) => { HideCaret(rtbText.Handle); };

            if (isCollapsible)
            {
                headerPanel = new Panel
                {
                    Height = 24,
                    BackColor = Color.Transparent,
                    Location = new Point(10, 6)
                };

                Label lblHeader = new Label
                {
                    Text = headerText,
                    Font = new Font("Microsoft YaHei", 9F, FontStyle.Bold),
                    ForeColor = isError ? Color.FromArgb(180, 40, 40) : Color.FromArgb(80, 80, 80),
                    Location = new Point(0, 3),
                    AutoSize = true
                };
                headerPanel.Controls.Add(lblHeader);

                btnToggle = new Button
                {
                    Text = "▶ 展开",
                    Font = new Font("Microsoft YaHei", 8F),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = Color.Transparent,
                    ForeColor = Color.FromArgb(0, 120, 212),
                    Location = new Point(lblHeader.Right + 6, 1),
                    Width = 48,
                    Height = 22,
                    Cursor = Cursors.Hand
                };
                btnToggle.FlatAppearance.BorderSize = 0;
                btnToggle.FlatAppearance.MouseOverBackColor = Color.FromArgb(230, 230, 230);
                btnToggle.Click += BtnToggle_Click;
                headerPanel.Controls.Add(btnToggle);

                lblHeader.SizeChanged += (s, e) => { btnToggle.Left = lblHeader.Right + 6; };

                bubblePanel.Controls.Add(headerPanel);

                rtbText.Text = bodyText;
                rtbText.Font = new Font("Consolas", 8.5F);
                rtbText.ForeColor = Color.FromArgb(60, 60, 60);
                rtbText.Location = new Point(10, 32);
                rtbText.Visible = false;

                bubblePanel.Controls.Add(rtbText);
            }
            else
            {
                FormatMarkdown(rtbText, text, isUser);
                rtbText.Location = new Point(10, 8);
                bubblePanel.Controls.Add(rtbText);
            }

            bubblePanel.Paint += BubblePanel_Paint;
            this.Controls.Add(bubblePanel);
            UpdatePosition(width);
        }

        private void BtnToggle_Click(object sender, EventArgs e)
        {
            isExpanded = !isExpanded;
            rtbText.Visible = isExpanded;
            btnToggle.Text = isExpanded ? "▼ 收起" : "▶ 展开";
            
            UpdatePosition(this.Width);
            
            if (this.Parent is FlowLayoutPanel flp)
            {
                flp.PerformLayout();
            }
        }

        private void BubblePanel_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Color backColor;
            if (isError) backColor = Color.FromArgb(253, 231, 233);
            else if (isSystem) backColor = Color.FromArgb(240, 240, 240);
            else if (isUser) backColor = Color.FromArgb(0, 120, 212);
            else backColor = Color.White;

            using (GraphicsPath path = GetRoundedPath(new Rectangle(0, 0, bubblePanel.Width - 1, bubblePanel.Height - 1), 8))
            using (SolidBrush brush = new SolidBrush(backColor))
            {
                e.Graphics.FillPath(brush, path);
            }

            if (!isUser && !isError)
            {
                using (GraphicsPath path = GetRoundedPath(new Rectangle(0, 0, bubblePanel.Width - 1, bubblePanel.Height - 1), 8))
                using (Pen pen = new Pen(Color.FromArgb(220, 220, 220), 1))
                {
                    e.Graphics.DrawPath(pen, path);
                }
            }
        }

        private GraphicsPath GetRoundedPath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int d = radius * 2;
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        public void UpdateWidth(int width)
        {
            this.Width = width;
            UpdatePosition(width);
        }

        private void UpdatePosition(int width)
        {
            int maxBubbleWidth = Convert.ToInt32(width * 0.85);
            int innerPaddingWidth = 20;

            int textWidth = maxBubbleWidth - innerPaddingWidth;
            if (textWidth < 50) textWidth = 50;

            // 1. Calculate the final bubble and RichTextBox widths
            Size measuredSize = TextRenderer.MeasureText(rtbText.Text, rtbText.Font, new Size(textWidth - 20, 0), TextFormatFlags.WordBreak);
            
            int finalBubbleWidth;
            if (isCollapsible)
            {
                finalBubbleWidth = maxBubbleWidth;
            }
            else
            {
                finalBubbleWidth = Math.Min(measuredSize.Width + innerPaddingWidth + 10, maxBubbleWidth);
            }

            int finalRtbWidth = finalBubbleWidth - innerPaddingWidth;
            if (finalRtbWidth < 50) finalRtbWidth = 50;

            rtbText.Width = finalRtbWidth;

            // Force handle creation so GetPositionFromCharIndex is completely accurate
            var handle = rtbText.Handle;

            // 2. Measure precise text wrapping height natively
            int calculatedHeight = 0;
            if (!string.IsNullOrEmpty(rtbText.Text))
            {
                Point pos = rtbText.GetPositionFromCharIndex(rtbText.Text.Length - 1);
                
                // Fallback if handle isn't fully ready or pos.Y returns 0 unexpectedly for multi-line text
                if (pos.Y == 0 && rtbText.Text.Contains("\n"))
                {
                    int lineCount = rtbText.Lines.Length;
                    int listBonus = 0;
                    foreach (var line in rtbText.Lines)
                    {
                        if (line.TrimStart().StartsWith("- ") || line.TrimStart().StartsWith("* "))
                        {
                            listBonus += 15;
                        }
                    }
                    int boldAndFontPadding = 24 + (lineCount * 4) + listBonus;
                    calculatedHeight = measuredSize.Height + boldAndFontPadding;
                }
                else
                {
                    calculatedHeight = pos.Y + rtbText.Font.Height + 16;
                }
            }

            // 3. Set control heights and layout bounds
            if (isCollapsible)
            {
                headerPanel.Width = finalRtbWidth;
                if (isExpanded)
                {
                    rtbText.Height = calculatedHeight;
                    bubblePanel.Width = maxBubbleWidth;
                    bubblePanel.Height = rtbText.Bottom + 8;
                }
                else
                {
                    bubblePanel.Width = maxBubbleWidth;
                    bubblePanel.Height = headerPanel.Bottom + 6;
                }
            }
            else
            {
                rtbText.Height = calculatedHeight;
                bubblePanel.Width = finalBubbleWidth;
                bubblePanel.Height = rtbText.Height + 16;
            }

            Color bubbleColor;
            if (isError) bubbleColor = Color.FromArgb(253, 231, 233);
            else if (isSystem) bubbleColor = Color.FromArgb(240, 240, 240);
            else if (isUser) bubbleColor = Color.FromArgb(0, 120, 212);
            else bubbleColor = Color.White;

            rtbText.BackColor = bubbleColor;
            rtbText.ForeColor = isUser ? Color.White : Color.FromArgb(33, 33, 33);

            this.Height = bubblePanel.Height + 10;

            if (isUser)
            {
                bubblePanel.Left = width - bubblePanel.Width - 5;
            }
            else
            {
                bubblePanel.Left = 5;
            }
            bubblePanel.Top = 5;
            
            bubblePanel.Invalidate();
        }

        private static void FormatMarkdown(RichTextBox rtb, string markdown, bool isUser)
        {
            rtb.Text = "";
            rtb.SelectionBullet = false;
            rtb.SelectionIndent = 0;

            string[] lines = markdown.Split('\n');
            for (int l = 0; l < lines.Length; l++)
            {
                string line = lines[l];
                
                if (line.TrimStart().StartsWith("- ") || line.TrimStart().StartsWith("* "))
                {
                    int bulletIndex = line.IndexOf('-');
                    if (bulletIndex < 0) bulletIndex = line.IndexOf('*');
                    line = line.Substring(bulletIndex + 2);
                    rtb.SelectionBullet = true;
                    rtb.SelectionIndent = 12;
                }
                else
                {
                    rtb.SelectionBullet = false;
                    rtb.SelectionIndent = 0;
                }

                int i = 0;
                while (i < line.Length)
                {
                    if (i + 1 < line.Length && line[i] == '*' && line[i + 1] == '*')
                    {
                        int next = line.IndexOf("**", i + 2);
                        if (next >= 0)
                        {
                            string boldText = line.Substring(i + 2, next - (i + 2));
                            rtb.SelectionFont = new Font(rtb.Font, FontStyle.Bold);
                            rtb.AppendText(boldText);
                            i = next + 2;
                        }
                        else
                        {
                            rtb.SelectionFont = new Font(rtb.Font, FontStyle.Regular);
                            rtb.AppendText("**");
                            i += 2;
                        }
                    }
                    else if (line[i] == '`')
                    {
                        int next = line.IndexOf('`', i + 1);
                        if (next >= 0)
                        {
                            string codeText = line.Substring(i + 1, next - (i + 1));
                            
                            rtb.SelectionFont = new Font("Consolas", rtb.Font.Size);
                            rtb.SelectionBackColor = isUser ? Color.FromArgb(0, 60, 120) : Color.FromArgb(228, 229, 230);
                            rtb.SelectionColor = isUser ? Color.FromArgb(200, 230, 255) : Color.FromArgb(190, 50, 50);

                            rtb.AppendText(codeText);
                            
                            rtb.SelectionBackColor = rtb.BackColor;
                            rtb.SelectionColor = rtb.ForeColor;
                            
                            i = next + 1;
                        }
                        else
                        {
                            rtb.SelectionFont = new Font(rtb.Font, FontStyle.Regular);
                            rtb.AppendText("`");
                            i++;
                        }
                    }
                    else
                    {
                        rtb.SelectionFont = new Font(rtb.Font, FontStyle.Regular);
                        rtb.SelectionColor = isUser ? Color.White : Color.FromArgb(33, 33, 33);
                        rtb.AppendText(line[i].ToString());
                        i++;
                    }
                }

                if (l < lines.Length - 1)
                {
                    rtb.AppendText(Environment.NewLine);
                }
            }
        }
    }
}

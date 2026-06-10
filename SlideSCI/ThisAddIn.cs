namespace SlideSCI
{
    public partial class ThisAddIn
    {
        public LatexToSvgConverter LatexSvgConverter { get; private set; }

        public static bool AreWindowsEqual(Microsoft.Office.Interop.PowerPoint.DocumentWindow win1, Microsoft.Office.Interop.PowerPoint.DocumentWindow win2)
        {
            if (win1 == null || win2 == null) return win1 == win2;
            try
            {
                if (win1.Caption != win2.Caption) return false;
                if (win1.Presentation.Name != win2.Presentation.Name) return false;
                if (win1.Presentation.FullName != win2.Presentation.FullName) return false;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public Microsoft.Office.Tools.CustomTaskPane ShapeLibraryTaskPane
        {
            get
            {
                try
                {
                    var app = this.Application;
                    if (app.Windows.Count == 0) return null;
                    var activeWindow = app.ActiveWindow;
                    if (activeWindow == null) return null;

                    foreach (Microsoft.Office.Tools.CustomTaskPane pane in this.CustomTaskPanes)
                    {
                        try
                        {
                            if (pane.Control is ShapeLibraryControl ctrl)
                            {
                                if (ctrl.AssociatedWindow != null && AreWindowsEqual(ctrl.AssociatedWindow, activeWindow))
                                    return pane;
                            }
                        }
                        catch { }
                    }
                }
                catch { }
                return null;
            }
        }

        public Microsoft.Office.Tools.CustomTaskPane AISidebarTaskPane
        {
            get
            {
                try
                {
                    var app = this.Application;
                    if (app.Windows.Count == 0) return null;
                    var activeWindow = app.ActiveWindow;
                    if (activeWindow == null) return null;

                    foreach (Microsoft.Office.Tools.CustomTaskPane pane in this.CustomTaskPanes)
                    {
                        try
                        {
                            if (pane.Control is AISidebarControl ctrl)
                            {
                                if (ctrl.AssociatedWindow != null && AreWindowsEqual(ctrl.AssociatedWindow, activeWindow))
                                    return pane;
                            }
                        }
                        catch { }
                    }
                }
                catch { }
                return null;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            LatexSvgConverter = new LatexToSvgConverter();
        }

        public void ToggleShapeLibraryTaskPane(Microsoft.Office.Interop.PowerPoint.DocumentWindow contextWindow = null)
        {
            try
            {
                var app = this.Application;
                if (app.Windows.Count == 0) return;
                var activeWindow = contextWindow ?? app.ActiveWindow;
                if (activeWindow == null) return;

                Microsoft.Office.Tools.CustomTaskPane activePane = null;
                foreach (Microsoft.Office.Tools.CustomTaskPane pane in this.CustomTaskPanes)
                {
                    try
                    {
                        if (pane.Control is ShapeLibraryControl ctrl)
                        {
                            if (ctrl.AssociatedWindow != null && AreWindowsEqual(ctrl.AssociatedWindow, activeWindow))
                            {
                                activePane = pane;
                                break;
                            }
                        }
                    }
                    catch { }
                }

                if (activePane == null)
                {
                    var control = new ShapeLibraryControl();
                    control.AssociatedWindow = activeWindow;
                    activePane = this.CustomTaskPanes.Add(control, "PPT素材库", activeWindow);
                    activePane.Width = 360;
                }

                activePane.Visible = !activePane.Visible;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"无法打开素材库面板: {ex.Message}", "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void ToggleAISidebarTaskPane(Microsoft.Office.Interop.PowerPoint.DocumentWindow contextWindow = null)
        {
            try
            {
                var app = this.Application;
                if (app.Windows.Count == 0) return;
                var activeWindow = contextWindow ?? app.ActiveWindow;
                if (activeWindow == null) return;

                Microsoft.Office.Tools.CustomTaskPane activePane = null;
                foreach (Microsoft.Office.Tools.CustomTaskPane pane in this.CustomTaskPanes)
                {
                    try
                    {
                        if (pane.Control is AISidebarControl ctrl)
                        {
                            if (ctrl.AssociatedWindow != null && AreWindowsEqual(ctrl.AssociatedWindow, activeWindow))
                            {
                                activePane = pane;
                                break;
                            }
                        }
                    }
                    catch { }
                }

                if (activePane == null)
                {
                    var control = new AISidebarControl();
                    control.AssociatedWindow = activeWindow;
                    activePane = this.CustomTaskPanes.Add(control, "SlideSCI AI 助手", activeWindow);
                    activePane.Width = 320;
                }

                activePane.Visible = !activePane.Visible;
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"无法打开 AI 助手面板: {ex.Message}", "错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}

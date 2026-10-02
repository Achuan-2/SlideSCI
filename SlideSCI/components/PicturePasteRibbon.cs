using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    /// <summary>
    /// 在设计器生成的 Ribbon XML 中添加普通 Paste 命令回调。
    /// VSTO 的动态回调通过 IReflect 完整转发，保留现有按钮与 Ribbon Load 处理。
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class PicturePasteRibbon : Office.IRibbonExtensibility, IReflect
    {
        private const string PasteCallback = nameof(OnPictureFilePaste);
        private readonly Office.IRibbonExtensibility ribbon;
        private readonly IReflect callbacks;
        private readonly Type ownType = typeof(PicturePasteRibbon);
        private bool pasting;

        internal PicturePasteRibbon(Office.IRibbonExtensibility ribbon)
        {
            this.ribbon = ribbon ?? throw new ArgumentNullException(nameof(ribbon));
            callbacks = (IReflect)ribbon;
        }

        public string GetCustomUI(string ribbonID)
        {
            string xml = ribbon.GetCustomUI(ribbonID);
            if (string.IsNullOrEmpty(xml)) return xml;
            XDocument document = XDocument.Parse(xml);
            XElement root = document.Root;
            XNamespace ns = root.Name.Namespace;
            XElement commands = root.Element(ns + "commands");
            if (commands == null)
            {
                commands = new XElement(ns + "commands");
                root.AddFirst(commands);
            }
            XElement paste = commands.Elements(ns + "command")
                .FirstOrDefault(command => (string)command.Attribute("idMso") == "Paste");
            if (paste == null)
            {
                paste = new XElement(ns + "command", new XAttribute("idMso", "Paste"));
                commands.Add(paste);
            }
            paste.SetAttributeValue("onAction", PasteCallback);
            return document.ToString(SaveOptions.DisableFormatting);
        }

        public void OnPictureFilePaste(Office.IRibbonControl control, ref bool cancelDefault)
        {
            // 无法确定是图片文件粘贴时，交还给 Office；不修改剪贴板内容。
            cancelDefault = pasting;
            if (pasting) return;
            try
            {
                if (!Clipboard.ContainsFileDropList()) return;
                string[] paths = Clipboard.GetFileDropList().Cast<string>().ToArray();
                if (paths.Length == 0 || !paths.All(PictureSourceInsertion.IsSupportedImageFile)) return;
                var application = Globals.ThisAddIn.Application;
                if (application.Windows.Count == 0) return;
                DocumentWindow window = application.ActiveWindow;
                if (window == null || window.View.Type != PpViewType.ppViewNormal ||
                    window.Selection.Type == PpSelectionType.ppSelectionText) return;
                Slide slide = window.View.Slide as Slide;
                if (slide == null) return;

                // 从此处起由插件负责本次插入；即使部分失败也不能再默认粘贴一次。
                cancelDefault = true;
                pasting = true;
                var failures = PictureSourceInsertion.Insert(application, slide, paths);
                if (failures.Count > 0)
                    MessageBox.Show("以下图片粘贴失败：\n\n" + string.Join(Environment.NewLine, failures),
                        "粘贴图片", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                if (cancelDefault)
                    MessageBox.Show($"粘贴图片时出错：{ex.Message}", "粘贴图片", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                else
                    System.Diagnostics.Debug.WriteLine($"读取文件剪贴板失败，使用默认粘贴：{ex.Message}");
            }
            finally { pasting = false; }
        }

        // Office 查询动态回调的方法信息，再通过 InvokeMember 调用。
        public MethodInfo GetMethod(string name, BindingFlags bindingAttr) => name == PasteCallback
            ? ownType.GetMethod(name, bindingAttr) : callbacks.GetMethod(name, bindingAttr);
        public MethodInfo GetMethod(string name, BindingFlags bindingAttr, Binder binder, Type[] types,
            ParameterModifier[] modifiers) => name == PasteCallback
            ? ownType.GetMethod(name, bindingAttr, binder, types, modifiers)
            : callbacks.GetMethod(name, bindingAttr, binder, types, modifiers);
        public MethodInfo[] GetMethods(BindingFlags bindingAttr) => callbacks.GetMethods(bindingAttr)
            .Concat(new[] { ownType.GetMethod(PasteCallback) }).ToArray();
        public MemberInfo[] GetMember(string name, BindingFlags bindingAttr) => name == PasteCallback
            ? ownType.GetMember(name, bindingAttr) : callbacks.GetMember(name, bindingAttr);
        public MemberInfo[] GetMembers(BindingFlags bindingAttr) => callbacks.GetMembers(bindingAttr)
            .Concat(ownType.GetMember(PasteCallback)).ToArray();
        public FieldInfo GetField(string name, BindingFlags bindingAttr) => callbacks.GetField(name, bindingAttr);
        public FieldInfo[] GetFields(BindingFlags bindingAttr) => callbacks.GetFields(bindingAttr);
        public PropertyInfo GetProperty(string name, BindingFlags bindingAttr) => callbacks.GetProperty(name, bindingAttr);
        public PropertyInfo GetProperty(string name, BindingFlags bindingAttr, Binder binder, Type returnType,
            Type[] types, ParameterModifier[] modifiers) =>
            callbacks.GetProperty(name, bindingAttr, binder, returnType, types, modifiers);
        public PropertyInfo[] GetProperties(BindingFlags bindingAttr) => callbacks.GetProperties(bindingAttr);
        public Type UnderlyingSystemType => callbacks.UnderlyingSystemType;
        public object InvokeMember(string name, BindingFlags invokeAttr, Binder binder, object target,
            object[] args, ParameterModifier[] modifiers, CultureInfo culture, string[] namedParameters) =>
            name == PasteCallback
                ? ownType.InvokeMember(name, invokeAttr, binder, this, args, modifiers, culture, namedParameters)
                : callbacks.InvokeMember(name, invokeAttr, binder, ribbon, args, modifiers, culture, namedParameters);
    }
}

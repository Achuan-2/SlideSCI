using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows.Forms;
using IDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    /// <summary>
    /// 仅在 Office UI 线程代理已注册的 OLE 拖放目标；拖放数据直接来自 IDataObject。
    /// 兼容性边界：OleDropTargetInterface 是 Windows OLE 的内部窗口属性，并非公开 API。
    /// 找不到属性或不在同一线程时不更改目标；安装失败时立即恢复原目标。
    /// </summary>
    internal sealed class MediaFileDropTracker : IDisposable
    {
        private const string TargetProperty = "OleDropTargetInterface";
        private readonly PowerPoint.Application application;
        private readonly Timer timer;
        private readonly Dictionary<IntPtr, Registration> targets = new Dictionary<IntPtr, Registration>();
        private readonly HashSet<IntPtr> failedWindows = new HashSet<IntPtr>();
        // 极少数原生撤销/恢复失败时保留引用，不能让仍注册的代理引用已释放的 COM 对象。
        private static readonly List<Registration> unrestoredTargets = new List<Registration>();
        private readonly uint uiThread = GetCurrentThreadId();
        private bool disposed;
        private bool refreshing;

        internal MediaFileDropTracker(PowerPoint.Application application)
        {
            this.application = application ?? throw new ArgumentNullException(nameof(application));
            timer = new Timer { Interval = 1000 };
            timer.Tick += OnTimer;
            RefreshTargets();
            timer.Start();
        }

        private void OnTimer(object sender, EventArgs e) => RefreshTargets();

        private void RefreshTargets()
        {
            if (disposed || refreshing || targets.Values.Any(target => target.Proxy.IsDragging)) return;
            refreshing = true;
            try
            {
                foreach (var pair in targets.ToArray())
                {
                    if (pair.Value.RestorePending)
                    {
                        if (RestoreOriginal(pair.Key, pair.Value))
                        {
                            pair.Value.Release();
                            targets.Remove(pair.Key);
                        }
                        continue;
                    }
                    if (!IsWindow(pair.Key) || GetProp(pair.Key, TargetProperty) != pair.Value.ProxyPointer)
                    {
                        // Office 已销毁窗口或自行更新目标，不覆盖它的新注册。
                        pair.Value.Release();
                        targets.Remove(pair.Key);
                    }
                }
                failedWindows.RemoveWhere(window => !IsWindow(window));
                foreach (PowerPoint.DocumentWindow document in application.Windows)
                {
                    IntPtr window = new IntPtr(document.HWND);
                    Attach(window);
                    EnumWindowsCallback callback = (child, parameter) => { Attach(child); return true; };
                    EnumChildWindows(window, callback, IntPtr.Zero);
                    GC.KeepAlive(callback);
                }
            }
            catch (Exception ex) { Debug.WriteLine($"发现文件拖放目标失败：{ex.Message}"); }
            finally { refreshing = false; }
        }

        private void Attach(IntPtr window)
        {
            if (targets.ContainsKey(window) || failedWindows.Contains(window) ||
                GetWindowThreadProcessId(window, out _) != uiThread) return;
            IntPtr originalPointer = GetProp(window, TargetProperty);
            if (originalPointer == IntPtr.Zero) return;
            object originalObject = null;
            Registration registration = null;
            bool revoked = false;
            try
            {
                // 独立 RCW 由代理持有，并随代理一起回收，不强制释放仍可能被调用的 COM 对象。
                originalObject = Marshal.GetUniqueObjectForIUnknown(originalPointer);
                var original = originalObject as IMediaFileDropTarget;
                if (original == null) return;
                var proxy = new MediaFileDropTarget(this, original, window);
                registration = new Registration(original, proxy);
                if (RevokeDragDrop(window) < 0) return;
                revoked = true;
                int result = RegisterDragDrop(window, proxy);
                if (result < 0)
                {
                    Marshal.ThrowExceptionForHR(result);
                }
                targets.Add(window, registration);
                registration = null;
                originalObject = null;
                revoked = false;
            }
            catch (Exception ex)
            {
                failedWindows.Add(window);
                Debug.WriteLine($"安装文件拖放代理失败：{ex.Message}");
            }
            finally
            {
                if (revoked && registration != null)
                {
                    if (!RestoreOriginal(window, registration))
                    {
                        registration.RestorePending = true;
                        targets[window] = registration;
                        registration = null;
                        originalObject = null;
                    }
                }
                if (registration != null) registration.Release();
            }
        }

        internal bool TryGetDestination(IntPtr targetWindow, MediaDropPoint point, out Destination destination)
        {
            destination = null;
            if (disposed) return false;
            try
            {
                // 先检查落点属于实际的目标窗口，避免侧栏、功能区覆盖幻灯片时误插入。
                IntPtr hitWindow = WindowFromPoint(point);
                if (hitWindow != targetWindow && !IsChild(targetWindow, hitWindow)) return false;
                foreach (PowerPoint.DocumentWindow document in application.Windows)
                {
                    IntPtr documentWindow = new IntPtr(document.HWND);
                    if (documentWindow != targetWindow && !IsChild(documentWindow, targetWindow)) continue;
                    if (document.View.Type != PowerPoint.PpViewType.ppViewNormal) continue;
                    var slide = document.View.Slide as PowerPoint.Slide;
                    if (slide == null) continue;
                    var presentation = document.Presentation;
                    float width = presentation.PageSetup.SlideWidth;
                    float height = presentation.PageSetup.SlideHeight;
                    int left = document.PointsToScreenPixelsX(0);
                    int top = document.PointsToScreenPixelsY(0);
                    int right = document.PointsToScreenPixelsX(width);
                    int bottom = document.PointsToScreenPixelsY(height);
                    if (right <= left || bottom <= top || point.X < left || point.X > right ||
                        point.Y < top || point.Y > bottom) continue;
                    destination = new Destination(document, slide, new PointF(
                        (point.X - left) * width / (right - left), (point.Y - top) * height / (bottom - top)));
                    return true;
                }
            }
            catch (Exception ex) { Debug.WriteLine($"识别文件拖放位置失败：{ex.Message}"); }
            return false;
        }

        internal bool InsertFiles(Destination destination, string[] paths)
        {
            try
            {
                // 拖入非当前演示文稿时先激活目标，保持插入和撤销都属于目标文稿。
                destination.Window.Activate();
                var failures = MediaSourceInsertion.Insert(application, destination.Slide, paths, destination.Position);
                if (failures.Count > 0)
                    MessageBox.Show("以下图片或视频拖入失败：\n\n" + string.Join(Environment.NewLine, failures),
                        "拖入图片或视频", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return failures.Count < paths.Length;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"拖入文件时出错：{ex.Message}", "拖入图片或视频", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return false;
            }
        }

        internal static string[] ReadSupportedFiles(IDataObject data)
        {
            var format = new FORMATETC { cfFormat = 15, dwAspect = DVASPECT.DVASPECT_CONTENT,
                lindex = -1, tymed = TYMED.TYMED_HGLOBAL };
            STGMEDIUM medium = default(STGMEDIUM);
            bool received = false;
            try
            {
                if (data == null || data.QueryGetData(ref format) != 0) return null;
                data.GetData(ref format, out medium);
                received = true;
                if (medium.tymed != TYMED.TYMED_HGLOBAL || medium.unionmember == IntPtr.Zero) return null;
                uint count = DragQueryFile(medium.unionmember, uint.MaxValue, null, 0);
                if (count == 0) return null;
                var files = new List<string>();
                for (uint index = 0; index < count; index++)
                {
                    uint length = DragQueryFile(medium.unionmember, index, null, 0);
                    if (length == 0 || length >= 32768) return null;
                    var buffer = new StringBuilder((int)length + 1);
                    if (DragQueryFile(medium.unionmember, index, buffer, (uint)buffer.Capacity) != length) return null;
                    string path = buffer.ToString();
                    if (!MediaSourceInsertion.IsSupportedMediaFile(path)) return null;
                    files.Add(path);
                }
                return files.ToArray();
            }
            catch (Exception ex) { Debug.WriteLine($"读取拖放文件失败：{ex.Message}"); return null; }
            finally { if (received) ReleaseStgMedium(ref medium); }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            timer.Stop();
            timer.Tick -= OnTimer;
            timer.Dispose();
            foreach (var pair in targets)
            {
                if (RestoreOriginal(pair.Key, pair.Value)) pair.Value.Release();
                else unrestoredTargets.Add(pair.Value);
            }
            targets.Clear();
        }

        private static bool RestoreOriginal(IntPtr window, Registration registration)
        {
            try
            {
                if (!IsWindow(window)) return true;
                IntPtr current = GetProp(window, TargetProperty);
                // Office 或其他插件已替换目标时，不撤销它的新注册。
                if (current != IntPtr.Zero && current != registration.ProxyPointer) return true;
                if (current == registration.ProxyPointer && RevokeDragDrop(window) < 0) return false;
                int result = RegisterDragDrop(window, registration.Original);
                if (result >= 0) return true;
                Debug.WriteLine($"恢复原拖放目标失败：0x{result:X8}");
            }
            catch (Exception ex) { Debug.WriteLine(ex); }
            return false;
        }

        internal sealed class Destination
        {
            internal PowerPoint.DocumentWindow Window { get; }
            internal PowerPoint.Slide Slide { get; }
            internal PointF Position { get; }
            internal Destination(PowerPoint.DocumentWindow window, PowerPoint.Slide slide, PointF position)
            { Window = window; Slide = slide; Position = position; }
        }

        private sealed class Registration
        {
            internal IMediaFileDropTarget Original { get; }
            internal MediaFileDropTarget Proxy { get; }
            internal IntPtr ProxyPointer { get; private set; }
            internal bool RestorePending { get; set; }
            internal Registration(IMediaFileDropTarget original, MediaFileDropTarget proxy)
            {
                Original = original;
                Proxy = proxy;
                ProxyPointer = Marshal.GetComInterfaceForObject(proxy, typeof(IMediaFileDropTarget));
            }
            internal void Release()
            {
                if (ProxyPointer == IntPtr.Zero) return;
                Marshal.Release(ProxyPointer);
                ProxyPointer = IntPtr.Zero;
            }
        }

        private delegate bool EnumWindowsCallback(IntPtr window, IntPtr parameter);
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr window, EnumWindowsCallback callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetPropW")]
        private static extern IntPtr GetProp(IntPtr window, string name);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")]
        private static extern bool IsChild(IntPtr parent, IntPtr child);
        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(MediaDropPoint point);
        [DllImport("ole32.dll")]
        private static extern int RegisterDragDrop(IntPtr window, [MarshalAs(UnmanagedType.Interface)] IMediaFileDropTarget target);
        [DllImport("ole32.dll")]
        private static extern int RevokeDragDrop(IntPtr window);
        [DllImport("ole32.dll")]
        private static extern void ReleaseStgMedium(ref STGMEDIUM medium);
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, EntryPoint = "DragQueryFileW")]
        private static extern uint DragQueryFile(IntPtr drop, uint index, StringBuilder path, uint size);
    }
}

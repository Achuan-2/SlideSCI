using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace SlideSCI
{
    /// <summary>
    /// OLE 拖放代理。始终将原来的目标保存在强引用中，普通拖放仍完整转发。
    /// 图片/视频文件仅在幻灯片区域由插件处理，向拖放源返回 COPY，绝不移动源文件。
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class MediaFileDropTarget : IMediaFileDropTarget
    {
        private readonly MediaFileDropTracker tracker;
        private readonly IMediaFileDropTarget original;
        private readonly IntPtr window;
        private string[] files;
        private uint allowedEffects;
        internal bool IsDragging { get; private set; }

        internal MediaFileDropTarget(MediaFileDropTracker tracker, IMediaFileDropTarget original, IntPtr window)
        {
            this.tracker = tracker;
            this.original = original;
            this.window = window;
        }

        public int DragEnter(IDataObject data, uint keyState, MediaDropPoint point, ref uint effect)
        {
            IsDragging = true;
            allowedEffects = effect;
            try
            {
                files = MediaFileDropTracker.ReadSupportedFiles(data);
                int result = original.DragEnter(data, keyState, point, ref effect);
                if (CanHandle(point)) { effect = 1; return 0; }
                return result;
            }
            catch (Exception ex) { ClearDrag(); return Fail(ex, ref effect); }
        }

        public int DragOver(uint keyState, MediaDropPoint point, ref uint effect)
        {
            try
            {
                int result = original.DragOver(keyState, point, ref effect);
                if (CanHandle(point)) { effect = 1; return 0; }
                return result;
            }
            catch (Exception ex) { return Fail(ex, ref effect); }
        }

        public int DragLeave()
        {
            try { return original.DragLeave(); }
            catch (Exception ex) { Debug.WriteLine(ex); return Marshal.GetHRForException(ex); }
            finally { ClearDrag(); }
        }

        public int Drop(IDataObject data, uint keyState, MediaDropPoint point, ref uint effect)
        {
            try
            {
                // Drop 时重新读取实际拖放数据，不能使用剪贴板里可能无关的路径。
                string[] droppedFiles = MediaFileDropTracker.ReadSupportedFiles(data);
                if ((allowedEffects & 1) != 0 && droppedFiles != null &&
                    tracker.TryGetDestination(window, point, out var destination))
                {
                    // 不调用原目标的 Drop，避免 Office 与插件各插入一次。
                    try { original.DragLeave(); }
                    catch (Exception ex) { Debug.WriteLine(ex); }
                    bool inserted = tracker.InsertFiles(destination, droppedFiles);
                    effect = inserted ? 1u : 0u;
                    return 0;
                }
                return original.Drop(data, keyState, point, ref effect);
            }
            catch (Exception ex) { return Fail(ex, ref effect); }
            finally { ClearDrag(); }
        }

        private bool CanHandle(MediaDropPoint point) => (allowedEffects & 1) != 0 && files != null &&
            tracker.TryGetDestination(window, point, out _);

        private void ClearDrag() { files = null; IsDragging = false; }

        private static int Fail(Exception exception, ref uint effect)
        {
            Debug.WriteLine($"文件拖放失败：{exception.Message}");
            effect = 0;
            return Marshal.GetHRForException(exception);
        }
    }

    // IDropTarget 的原生 ABI：POINTL 按值传递，HRESULT 原样返回。
    [ComImport]
    [ComVisible(true)]
    [Guid("00000122-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMediaFileDropTarget
    {
        [PreserveSig]
        int DragEnter([In, MarshalAs(UnmanagedType.Interface)] IDataObject data, uint keyState,
            MediaDropPoint point, ref uint effect);
        [PreserveSig]
        int DragOver(uint keyState, MediaDropPoint point, ref uint effect);
        [PreserveSig]
        int DragLeave();
        [PreserveSig]
        int Drop([In, MarshalAs(UnmanagedType.Interface)] IDataObject data, uint keyState,
            MediaDropPoint point, ref uint effect);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MediaDropPoint
    {
        public int X;
        public int Y;
    }
}

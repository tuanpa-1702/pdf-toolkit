// PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
// SPDX-License-Identifier: AGPL-3.0-only
// Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
// Phần mở rộng Explorer cho file .pdf: ảnh thu nhỏ (IThumbnailProvider) và khung xem trước (IPreviewHandler).
// Vẽ trang PDF bằng Windows.Data.Pdf có sẵn trong Windows 10/11 - không cần thư viện ngoài.
// Windows nạp DLL này vào tiến trình riêng (dllhost / prevhost), không chạy trong Explorer.
// Build: build_shell.bat. Đăng ký (HKCU, không cần admin): PDF Toolkit.exe --install
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Data.Pdf;
using Windows.Storage.Streams;

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyTitle("PDF Toolkit Explorer Extension")]
[assembly: AssemblyCopyright("Copyright (C) 2026 TuanPA - GNU AGPL v3")]

namespace PDFToolkitShell
{
    [ComImport, Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IInitializeWithStream
    {
        void Initialize(IStream pstream, uint grfMode);
    }

    [ComImport, Guid("e357fccd-a995-4576-b01f-234630154e96"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IThumbnailProvider
    {
        void GetThumbnail(uint cx, out IntPtr phbmp, out uint pdwAlpha);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    [StructLayout(LayoutKind.Sequential)]
    public struct MSG { public IntPtr hwnd; public int message; public IntPtr wParam, lParam; public int time, ptX, ptY; }

    [ComImport, Guid("8895b1c6-b41f-4c1c-a562-0d564250836f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPreviewHandler
    {
        void SetWindow(IntPtr hwnd, ref RECT prc);
        void SetRect(ref RECT prc);
        void DoPreview();
        void Unload();
        void SetFocus();
        void QueryFocus(out IntPtr phwnd);
        [PreserveSig] int TranslateAccelerator(ref MSG pmsg);
    }

    [ComImport, Guid("fc4801a3-2ba9-11cf-a229-00aa003d7352"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IObjectWithSite
    {
        void SetSite([MarshalAs(UnmanagedType.IUnknown)] object pUnkSite);
        void GetSite(ref Guid riid, [MarshalAs(UnmanagedType.IUnknown)] out object ppvSite);
    }

    [ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleWindow
    {
        void GetWindow(out IntPtr phwnd);
        void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
    }

    /// Đọc PDF từ dữ liệu trong bộ nhớ và vẽ trang thành ảnh (Windows.Data.Pdf, chạy trên luồng nền).
    public sealed class PdfPages : IDisposable
    {
        readonly PdfDocument doc;
        public const long MaxBytes = 512L * 1024 * 1024;

        public PdfPages(byte[] data)
        {
            doc = Task.Run(() =>
            {
                var ras = new InMemoryRandomAccessStream();
                var w = new DataWriter(ras);
                w.WriteBytes(data);
                Wait(w.StoreAsync());
                w.DetachStream();
                ras.Seek(0);
                return Wait(PdfDocument.LoadFromStreamAsync(ras));
            }).Result;
        }

        // Chờ thao tác bất đồng bộ của WinRT (luồng nền, không cần vòng lặp thông điệp).
        static T Wait<T>(Windows.Foundation.IAsyncOperation<T> op)
        {
            Wait((Windows.Foundation.IAsyncInfo)op);
            return op.GetResults();
        }

        static void Wait(Windows.Foundation.IAsyncAction op)
        {
            Wait((Windows.Foundation.IAsyncInfo)op);
            op.GetResults();
        }

        static void Wait(Windows.Foundation.IAsyncInfo op)
        {
            while (op.Status == Windows.Foundation.AsyncStatus.Started) System.Threading.Thread.Sleep(2);
            if (op.Status == Windows.Foundation.AsyncStatus.Error) throw op.ErrorCode;
            if (op.Status == Windows.Foundation.AsyncStatus.Canceled) throw new OperationCanceledException();
        }

        public int Count { get { return (int)doc.PageCount; } }

        /// Kích thước trang (đơn vị 1/96 inch, đã tính xoay trang).
        public SizeF PageSize(int i)
        {
            using (PdfPage p = doc.GetPage((uint)i)) return new SizeF((float)p.Size.Width, (float)p.Size.Height);
        }

        /// Vẽ trang i vừa khung w x h (giữ tỉ lệ), nền trắng.
        public Bitmap Render(int i, int w, int h)
        {
            return Task.Run(() =>
            {
                using (PdfPage p = doc.GetPage((uint)i))
                {
                    double s = Math.Min(w / p.Size.Width, h / p.Size.Height);
                    var opt = new PdfPageRenderOptions { DestinationWidth = (uint)Math.Max(1, Math.Round(p.Size.Width * s)),
                                                         DestinationHeight = (uint)Math.Max(1, Math.Round(p.Size.Height * s)) };
                    var outS = new InMemoryRandomAccessStream();
                    Wait(p.RenderToStreamAsync(outS, opt));         // ảnh PNG
                    var r = new DataReader(outS.GetInputStreamAt(0));
                    uint n = Wait(r.LoadAsync((uint)outS.Size));
                    var png = new byte[n];
                    r.ReadBytes(png);
                    using (var st = new MemoryStream(png))
                    using (var tmp = new Bitmap(st))
                        return new Bitmap(tmp);                    // tách khỏi stream
                }
            }).Result;
        }

        public void Dispose() { }

        public static byte[] ReadAll(IStream s)
        {
            var ms = new MemoryStream();
            var buf = new byte[1 << 16];
            IntPtr pRead = Marshal.AllocCoTaskMem(4);
            try
            {
                while (true)
                {
                    s.Read(buf, buf.Length, pRead);
                    int n = Marshal.ReadInt32(pRead);
                    if (n <= 0) break;
                    ms.Write(buf, 0, n);
                    if (ms.Length > MaxBytes) throw new IOException("PDF quá lớn");
                }
            }
            finally { Marshal.FreeCoTaskMem(pRead); }
            return ms.ToArray();
        }
    }

    // ====================================================================== ảnh thu nhỏ
    [ComVisible(true), Guid("3F0C5D2E-8B7A-4E61-9C24-6A1D5E8B7F31"), ClassInterface(ClassInterfaceType.None),
     ProgId("PDFToolkit.Thumbnail")]
    public class PdfThumbnailProvider : IInitializeWithStream, IThumbnailProvider
    {
        byte[] data;

        public void Initialize(IStream pstream, uint grfMode)
        {
            data = PdfPages.ReadAll(pstream);
            Marshal.ReleaseComObject(pstream);
        }

        public void GetThumbnail(uint cx, out IntPtr phbmp, out uint pdwAlpha)
        {
            phbmp = IntPtr.Zero;
            pdwAlpha = 0;
            if (data == null) throw new COMException("chưa có dữ liệu", unchecked((int)0x8000FFFF));
            using (Bitmap bmp = Thumbnail(data, (int)cx))
                phbmp = bmp.GetHbitmap(Color.White);
            pdwAlpha = 1;                                           // WTSAT_RGB
            data = null;
        }

        /// Trang đầu thu nhỏ, cạnh dài = size, có viền mảnh để trang trắng không lẫn vào nền Explorer.
        public static Bitmap Thumbnail(byte[] pdf, int size)
        {
            var pages = new PdfPages(pdf);
            if (pages.Count == 0) throw new COMException("PDF không có trang", unchecked((int)0x80004005));
            Bitmap bmp = pages.Render(0, size, size);
            using (var g = Graphics.FromImage(bmp))
            using (var p = new Pen(Color.FromArgb(160, 160, 160)))
                g.DrawRectangle(p, 0, 0, bmp.Width - 1, bmp.Height - 1);
            return bmp;
        }
    }

    // ====================================================================== khung xem trước
    [ComVisible(true), Guid("7A9E3C14-2D5B-4F86-B1E0-93C7A64D2B58"), ClassInterface(ClassInterfaceType.None),
     ProgId("PDFToolkit.Preview")]
    public class PdfPreviewHandler : IPreviewHandler, IInitializeWithStream, IObjectWithSite, IOleWindow
    {
        [DllImport("user32.dll")] static extern IntPtr SetParent(IntPtr child, IntPtr parent);
        [DllImport("user32.dll")] static extern IntPtr GetFocus();

        byte[] data;
        IntPtr parent;
        RECT rect;
        object site;
        PreviewView view;

        public void Initialize(IStream pstream, uint grfMode)
        {
            data = PdfPages.ReadAll(pstream);
            Marshal.ReleaseComObject(pstream);
        }

        public void SetWindow(IntPtr hwnd, ref RECT prc)
        {
            parent = hwnd;
            rect = prc;
            if (view != null)
            {
                SetParent(view.Handle, parent);
                Place();
            }
        }

        public void SetRect(ref RECT prc)
        {
            rect = prc;
            Place();
        }

        void Place()
        {
            if (view != null) view.Bounds = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
        }

        public void DoPreview()
        {
            if (data == null) return;
            view = new PreviewView();
            view.CreateControl();
            if (parent != IntPtr.Zero) SetParent(view.Handle, parent);
            Place();
            view.Load(data);
            view.Visible = true;
        }

        public void Unload()
        {
            if (view != null) { view.Dispose(); view = null; }
            data = null;
        }

        public void SetFocus() { if (view != null) view.Focus(); }
        public void QueryFocus(out IntPtr phwnd) { phwnd = GetFocus(); }
        public int TranslateAccelerator(ref MSG pmsg) { return 1; }      // S_FALSE: không xử lý phím tắt

        public void SetSite(object pUnkSite) { site = pUnkSite; }
        public void GetSite(ref Guid riid, out object ppvSite) { ppvSite = site; }
        public void GetWindow(out IntPtr phwnd) { phwnd = view != null ? view.Handle : parent; }
        public void ContextSensitiveHelp(bool fEnterMode) { }

        public PreviewView View { get { return view; } }
    }

    /// Danh sách trang cuộn dọc, vừa chiều rộng khung. Trang vẽ dần khi cuộn tới (file nhiều trang vẫn mở nhanh).
    public class PreviewView : ScrollableControl
    {
        const int Gap = 10, MaxPages = 300;
        PdfPages pages;
        readonly List<SizeF> sizes = new List<SizeF>();
        readonly Dictionary<int, Bitmap> cache = new Dictionary<int, Bitmap>();
        string error;
        readonly Timer timer = new Timer { Interval = 30 };

        public PreviewView()
        {
            DoubleBuffered = true;
            AutoScroll = true;
            BackColor = Color.FromArgb(229, 231, 235);
            Visible = false;
            timer.Tick += (s, e) => RenderNext();
        }

        public void Load(byte[] data)
        {
            try
            {
                pages = new PdfPages(data);
                for (int i = 0; i < Math.Min(pages.Count, MaxPages); i++) sizes.Add(pages.PageSize(i));
            }
            catch (Exception)
            {
                error = "Không xem trước được file này (có mật khẩu hoặc bị hỏng).\nCannot preview this file (password-protected or damaged).";
            }
            Relayout();
        }

        public int PageCount { get { return sizes.Count; } }

        int PageWidth { get { return Math.Max(50, ClientSize.Width - 2 * Gap); } }

        Rectangle PageRect(int i)
        {
            int w = PageWidth, y = Gap;
            for (int k = 0; k < i; k++) y += (int)(w * sizes[k].Height / sizes[k].Width) + Gap;
            return new Rectangle(Gap, y, w, (int)(w * sizes[i].Height / sizes[i].Width));
        }

        void Relayout()
        {
            if (sizes.Count == 0) { AutoScrollMinSize = Size.Empty; Invalidate(); return; }
            Rectangle last = PageRect(sizes.Count - 1);
            AutoScrollMinSize = new Size(0, last.Bottom + Gap);
            foreach (var b in cache.Values) b.Dispose();
            cache.Clear();
            Invalidate();
            timer.Start();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (pages != null) Relayout();
        }

        protected override void OnScroll(ScrollEventArgs se) { base.OnScroll(se); Invalidate(); timer.Start(); }
        protected override void OnMouseWheel(MouseEventArgs e) { base.OnMouseWheel(e); Invalidate(); timer.Start(); }

        /// Vẽ 1 trang còn thiếu trong vùng đang thấy (mỗi nhịp timer 1 trang để giao diện không bị đứng).
        void RenderNext()
        {
            if (pages == null || IsDisposed) { timer.Stop(); return; }
            var visible = new Rectangle(0, -AutoScrollPosition.Y, ClientSize.Width, ClientSize.Height);
            for (int i = 0; i < sizes.Count; i++)
            {
                Rectangle r = PageRect(i);
                if (r.Top > visible.Bottom + ClientSize.Height) break;    // vẽ trước thêm 1 màn hình
                if (r.Bottom < visible.Top || cache.ContainsKey(i)) continue;
                try { cache[i] = pages.Render(i, r.Width, r.Height); }
                catch { cache[i] = null; }
                Invalidate();
                return;
            }
            timer.Stop();
        }

        /// Vẽ toàn bộ phần đang thấy vào ảnh (kiểm thử, không cần hiện cửa sổ).
        public Bitmap Snapshot()
        {
            while (timer.Enabled) RenderNext();
            var bmp = new Bitmap(Math.Max(1, ClientSize.Width), Math.Max(1, ClientSize.Height));
            using (var g = Graphics.FromImage(bmp)) Draw(g);
            return bmp;
        }

        protected override void OnPaint(PaintEventArgs e) { Draw(e.Graphics); }

        void Draw(Graphics g)
        {
            g.Clear(BackColor);
            if (error != null)
            {
                using (var f = new Font("Segoe UI", 9.5f))
                    g.DrawString(error, f, Brushes.DimGray, new RectangleF(Gap, Gap, ClientSize.Width - 2 * Gap, 200));
                return;
            }
            g.TranslateTransform(0, AutoScrollPosition.Y);
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            int top = -AutoScrollPosition.Y, bottom = top + ClientSize.Height;
            using (var border = new Pen(Color.FromArgb(203, 213, 225)))
                for (int i = 0; i < sizes.Count; i++)
                {
                    Rectangle r = PageRect(i);
                    if (r.Top > bottom) break;
                    if (r.Bottom < top) continue;
                    Bitmap b;
                    g.FillRectangle(Brushes.White, r);
                    if (cache.TryGetValue(i, out b) && b != null) g.DrawImage(b, r);
                    g.DrawRectangle(border, r.X - 1, r.Y - 1, r.Width + 1, r.Height + 1);
                }
            if (pages != null && pages.Count > sizes.Count)
            {
                Rectangle last = PageRect(sizes.Count - 1);
                using (var f = new Font("Segoe UI", 9f))
                    g.DrawString("+" + (pages.Count - sizes.Count) + " ...", f, Brushes.DimGray, Gap, last.Bottom + 2);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                timer.Dispose();
                foreach (var b in cache.Values) if (b != null) b.Dispose();
                cache.Clear();
            }
            base.Dispose(disposing);
        }
    }
}

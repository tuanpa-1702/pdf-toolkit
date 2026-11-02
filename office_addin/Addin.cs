// PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
// SPDX-License-Identifier: AGPL-3.0-only
// Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
// Add-in PDF Toolkit cho Word / Excel / PowerPoint / Outlook (COM add-in, không cần quyền admin). Phần Outlook: Outlook.cs.
// Tab "PDF Toolkit" trên ribbon:
//   Tạo PDF (Tạo PDF, Tuỳ chọn) | Tạo & gửi email (đính kèm email, trộn thư ra PDF) |
//   Xem xét & nhận xét (gửi để nhận xét, nhập nhận xét từ PDF...) | Xử lý PDF (ký số, mật khẩu, nén...) | Ứng dụng
// Build: build_addin.bat (dùng csc.exe có sẵn của .NET Framework 4). Đăng ký: PDF Toolkit.exe --install
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyVersion("1.4.0.0")]
[assembly: AssemblyTitle("PDF Toolkit Office Add-in")]
[assembly: AssemblyCopyright("Copyright (C) 2026 TuanPA - GNU AGPL v3")]

namespace PDFToolkitAddin
{
    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        // custom là SAFEARRAY(VARIANT) - phải khai báo rõ, mặc định CLR marshal Array thành _Array và làm Office crash
        [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] int connectMode,
                                      [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst,
                                      [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(2)] void OnDisconnection([In] int removeMode,
                                         [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(3)] void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(4)] void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
        [DispId(5)] void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    [ComImport, Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonExtensibility
    {
        [DispId(1)] string GetCustomUI(string ribbonID);
    }

    [ComImport, Guid("000C0395-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRibbonControl
    {
        [DispId(1)] string Id { get; }
        [DispId(2)] object Context { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
        [DispId(3)] string Tag { get; }
    }

    // Đổi Image -> IPictureDisp cho ribbon
    internal class PictureConverter : AxHost
    {
        private PictureConverter() : base("") { }
        public static object ToPicture(Image img) { return GetIPictureDispFromPicture(img); }
    }

    [ComVisible(true), Guid("6E2B7C1A-4F0D-4E8B-9A57-2D3C8F1B5A61"), ProgId("PDFToolkit.OfficeAddin"),
     ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private dynamic app;
        private string host = "";            // word | excel | powerpoint

        public void OnConnection(object application, int connectMode, object addInInst, ref Array custom)
        {
            app = application;
            try
            {
                string name = ((string)app.Name).ToLowerInvariant();
                host = name.Contains("word") ? "word" : name.Contains("excel") ? "excel" :
                       name.Contains("powerpoint") ? "powerpoint" : name.Contains("outlook") ? "outlook" : "";
            }
            catch { host = ""; }
            try { ((dynamic)addInInst).Object = this; } catch { }     // cho script gọi ExportActive()
        }

        /// Xuất tài liệu đang mở ra PDF theo Tuỳ chọn (không mở gì thêm) - dùng cho script / kiểm thử tự động.
        public string ExportActive() { return Export(false, false); }

        /// Nhập nhận xét từ file PDF vào tài liệu Word đang mở - dùng cho script / kiểm thử tự động.
        public string ImportCommentsFrom(string pdf) { return Comments.Import(this, app, pdf, true); }

        /// Trộn thư tài liệu Word đang mở ra từng file PDF (không email, không hộp thoại) - kiểm thử tự động.
        public int MergeToPdf(string nameTemplate, string folder) { return MailMergeDialog.RunQuiet(this, app, nameTemplate, folder); }

        public void OnDisconnection(int removeMode, ref Array custom) { app = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // ------------------------------------------------------------------ ribbon
        public string GetCustomUI(string ribbonID)
        {
            if (host == "") return "";
            if (host == "outlook") return OutlookUI.Xml(ribbonID);
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
  <ribbon><tabs><tab id='pdfToolkitTab' getLabel='GetLabel'>
    <group id='grpCreate' getLabel='GetLabel'>
      <button id='btnCreate' size='large' getLabel='GetLabel' getScreentip='GetTip' getImage='GetImage' onAction='OnAction'/>
      <button id='btnPrefs' size='large' getLabel='GetLabel' getScreentip='GetTip' getImage='GetImage' onAction='OnAction'/>
    </group>
    <group id='grpEmail' getLabel='GetLabel'>
      <button id='btnAttach' size='large' getLabel='GetLabel' getScreentip='GetTip' getImage='GetImage' onAction='OnAction'/>
      <button id='btnMailMerge' size='large' getLabel='GetLabel' getScreentip='GetTip' getImage='GetImage' onAction='OnAction' getVisible='GetVisible'/>
    </group>
    <group id='grpReview' getLabel='GetLabel'>
      <button id='btnReview' size='large' getLabel='GetLabel' getScreentip='GetTip' getImage='GetImage' onAction='OnAction'/>
      <menu id='mnuComments' size='large' getLabel='GetLabel' getScreentip='GetTip' getImage='GetImage' getVisible='GetVisible'>
        <button id='btnImportComments' getLabel='GetLabel' getScreentip='GetTip' getImage='GetImage' onAction='OnAction'/>
        <button id='btnExportMarkup' getLabel='GetLabel' getScreentip='GetTip' getImage='GetImage' onAction='OnAction'/>
        <menuSeparator id='sepComments'/>
        <button id='btnAcceptAll' getLabel='GetLabel' getScreentip='GetTip' onAction='OnAction'/>
        <button id='btnRejectAll' getLabel='GetLabel' getScreentip='GetTip' onAction='OnAction'/>
        <button id='btnDeleteComments' getLabel='GetLabel' getScreentip='GetTip' onAction='OnAction'/>
      </menu>
    </group>
    <group id='grpTools' getLabel='GetLabel'>
      <button id='btnSign' size='large' getLabel='GetLabel' getScreentip='GetTip' imageMso='SignatureLineInsert' onAction='OnAction'/>
      <button id='btnProtect' getLabel='GetLabel' getScreentip='GetTip' imageMso='FileDocumentEncrypt' onAction='OnAction'/>
      <button id='btnCompress' getLabel='GetLabel' getScreentip='GetTip' imageMso='PicturesCompress' onAction='OnAction'/>
      <button id='btnWatermark' getLabel='GetLabel' getScreentip='GetTip' imageMso='WordArtInsertGallery' onAction='OnAction'/>
      <button id='btnMerge' getLabel='GetLabel' getScreentip='GetTip' imageMso='FileSaveAsPdfOrXps' onAction='OnAction'/>
    </group>
    <group id='grpApp' getLabel='GetLabel'>
      <button id='btnOpen' size='large' getLabel='GetLabel' getScreentip='GetTip' getImage='GetImage' onAction='OnAction'/>
    </group>
  </tab></tabs></ribbon>
</customUI>";
        }

        private object ribbon;                                  // IRibbonUI (gọi kiểu late-binding)

        public void OnRibbonLoad(object ribbonUI) { ribbon = ribbonUI; }

        /// Chọn tab PDF Toolkit trên ribbon (kiểm thử / hướng dẫn sử dụng).
        public void ShowTab()
        {
            if (ribbon != null) ((dynamic)ribbon).ActivateTab("pdfToolkitTab");
        }

        internal static bool English()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\TuanPA\PDFToolkit"))
                    return k != null && Convert.ToString(k.GetValue("lang", "vi")) == "en";
            }
            catch { return false; }
        }

        internal static string L(string vi, string en) { return English() ? en : vi; }

        private static readonly Dictionary<string, string[]> Text = new Dictionary<string, string[]>
        {
            // id: { nhãn VI, nhãn EN, gợi ý VI, gợi ý EN }
            { "pdfToolkitTab", new[] { "PDF Toolkit", "PDF Toolkit", "", "" } },
            { "grpCreate", new[] { "Tạo PDF", "Create PDF", "", "" } },
            { "grpEmail", new[] { "Tạo & gửi email", "Create and Email", "", "" } },
            { "grpReview", new[] { "Xem xét & nhận xét", "Review and Comment", "", "" } },
            { "grpTools", new[] { "Xử lý PDF", "PDF tools", "", "" } },
            { "grpApp", new[] { "Ứng dụng", "App", "", "" } },
            { "btnCreate", new[] { "Tạo PDF", "Create PDF",
                "Tạo PDF từ tài liệu đang mở theo Tuỳ chọn (nơi lưu, phạm vi, bookmark...) rồi mở bằng PDF Toolkit",
                "Create a PDF from the open document using your Preferences, then open it in PDF Toolkit" } },
            { "btnPrefs", new[] { "Tuỳ chọn", "Preferences",
                "Nơi lưu PDF, phạm vi (cả tài liệu / trang / vùng chọn), bookmark, PDF/A, kèm nhận xét, việc làm sau khi tạo",
                "Output folder, range, bookmarks, PDF/A, comments, what to do after creating" } },
            { "btnAttach", new[] { "Tạo và đính kèm email", "Create and Attach to Email",
                "Tạo PDF rồi mở thư mới trong Outlook có sẵn file đính kèm", "Create a PDF and attach it to a new Outlook email" } },
            { "btnMailMerge", new[] { "Trộn thư", "Mail Merge",
                "Trộn thư ra PDF: mỗi người nhận một file PDF, có thể gửi email riêng cho từng người qua Outlook",
                "Merge to PDF: one PDF per recipient, optionally emailed to each one through Outlook" } },
            { "btnReview", new[] { "Tạo và gửi để nhận xét", "Create and Send for Commenting",
                "Tạo PDF (kèm nhận xét hiện có) và soạn email mời người khác xem, ghi chú trên PDF rồi gửi lại",
                "Create a PDF and email it to reviewers so they can add comments and send it back" } },
            { "mnuComments", new[] { "Nhận xét PDF", "PDF Comments", "Nhập nhận xét từ PDF, xử lý thay đổi và nhận xét", "Import PDF comments, handle changes and comments" } },
            { "btnImportComments", new[] { "Nhập nhận xét từ PDF...", "Import Comments from PDF...",
                "Đưa các ghi chú / tô sáng / gạch xoá trong file PDF đã được nhận xét vào tài liệu Word này (dạng Comment)",
                "Bring notes, highlights and strike-outs from a reviewed PDF into this Word document as comments" } },
            { "btnExportMarkup", new[] { "Tạo PDF kèm nhận xét & thay đổi", "Create PDF with Comments and Changes",
                "Tạo PDF có hiện nhận xét và Track Changes", "Create a PDF that shows comments and tracked changes" } },
            { "btnAcceptAll", new[] { "Chấp nhận tất cả thay đổi", "Accept All Changes", "Chấp nhận mọi thay đổi đang theo dõi (Track Changes)", "Accept all tracked changes" } },
            { "btnRejectAll", new[] { "Từ chối tất cả thay đổi", "Reject All Changes", "Từ chối mọi thay đổi đang theo dõi", "Reject all tracked changes" } },
            { "btnDeleteComments", new[] { "Xoá tất cả nhận xét", "Delete All Comments", "Xoá mọi Comment trong tài liệu", "Delete every comment in the document" } },
            { "btnSign", new[] { "PDF + Ký số", "PDF + Sign",
                "Tạo PDF rồi ký số bằng USB Token", "Create a PDF, then sign it with a USB token" } },
            { "btnProtect", new[] { "PDF + Mật khẩu", "PDF + Password",
                "Tạo PDF rồi đặt mật khẩu / giới hạn quyền", "Create a PDF, then set a password / permissions" } },
            { "btnCompress", new[] { "PDF + Nén", "PDF + Compress",
                "Tạo PDF rồi nén nhỏ để gửi email", "Create a PDF, then compress it for email" } },
            { "btnWatermark", new[] { "PDF + Watermark", "PDF + Watermark",
                "Tạo PDF rồi chèn chữ mờ (MẬT, BẢN NHÁP...)", "Create a PDF, then add a watermark" } },
            { "btnMerge", new[] { "PDF + Gộp file", "PDF + Merge",
                "Tạo PDF rồi gộp với các file PDF khác", "Create a PDF, then merge it with other PDFs" } },
            { "btnOpen", new[] { "Mở PDF Toolkit", "Open PDF Toolkit", "Mở trang chủ PDF Toolkit", "Open the PDF Toolkit home page" } },
            // Outlook
            { "olGrpMail", new[] { "PDF Toolkit", "PDF Toolkit", "", "" } },
            { "olGrpRead", new[] { "PDF Toolkit", "PDF Toolkit", "", "" } },
            { "olGrpCompose", new[] { "PDF Toolkit", "PDF Toolkit", "", "" } },
            { "olGrpInline", new[] { "PDF Toolkit", "PDF Toolkit", "", "" } },
            { "olMailToPdf", new[] { "Email → PDF", "Email to PDF",
                "Lưu email đang chọn thành PDF (giữ định dạng thư, nhúng file đính kèm); chọn nhiều thư thì gộp được thành 1 file",
                "Save the selected emails as PDF (keeps formatting, embeds attachments); several emails can be combined into one file" } },
            { "olAttToPdf", new[] { "Lưu đính kèm thành PDF", "Save Attachments as PDF",
                "Chuyển file đính kèm Word / Excel / PowerPoint / ảnh sang PDF và lưu vào thư mục (PDF giữ nguyên)",
                "Convert Word / Excel / PowerPoint / image attachments to PDF and save them to a folder (PDFs are copied as is)" } },
            { "olCtxMail", new[] { "Lưu thành PDF (PDF Toolkit)", "Save as PDF (PDF Toolkit)", "", "" } },
            { "olCtxMails", new[] { "Lưu thành PDF (PDF Toolkit)", "Save as PDF (PDF Toolkit)", "", "" } },
            { "olCtxAttOpen", new[] { "Mở bằng PDF Toolkit", "Open with PDF Toolkit", "", "" } },
            { "olCtxAttSave", new[] { "Lưu thành PDF...", "Save as PDF...", "", "" } },
            { "olAttachAsPdf", new[] { "Đính kèm dạng PDF", "Attach as PDF",
                "Chuyển file Word / Excel / PowerPoint / ảnh đang đính kèm sang PDF trước khi gửi",
                "Convert attached Word / Excel / PowerPoint / image files to PDF before sending" } },
            { "olCompressAtt", new[] { "Nén PDF đính kèm", "Compress PDF Attachments",
                "Nén nhỏ các file PDF đang đính kèm để gửi email nhẹ hơn", "Shrink the attached PDF files so the email is smaller" } },
            { "olProtectAtt", new[] { "PDF đính kèm + Mật khẩu", "Password-protect PDF Attachments",
                "Đặt mật khẩu mở file (AES-256) cho các PDF đang đính kèm", "Add an open password (AES-256) to the attached PDFs" } },
        };

        private static string T(string id, int idx)
        {
            if (id.EndsWith("_i")) id = id.Substring(0, id.Length - 2);       // nút trên tab trả lời trong khung đọc
            string[] v;
            if (!Text.TryGetValue(id, out v)) return id;
            return v[idx + 1];                      // ribbon luôn dùng tên tiếng Anh (giống các add-in Office khác)
        }

        public string GetLabel(IRibbonControl control) { return T(control.Id, 0); }
        public string GetTip(IRibbonControl control) { return T(control.Id, 2); }

        public bool GetVisible(IRibbonControl control)
        {
            return host == "word";                  // trộn thư, nhận xét / Track Changes: chỉ Word
        }

        public object GetImage(IRibbonControl control)
        {
            try
            {
                if (control.Id == "btnOpen")
                {
                    Icon ic = Icon.ExtractAssociatedIcon(ExePath());
                    using (Bitmap bmp = new Bitmap(ic.ToBitmap(), 32, 32))
                        return PictureConverter.ToPicture(new Bitmap(bmp));
                }
                string id = control.Id.EndsWith("_i") ? control.Id.Substring(0, control.Id.Length - 2) : control.Id;
                return PictureConverter.ToPicture(Icons.Get(id));
            }
            catch { return null; }
        }

        // ------------------------------------------------------------------ thao tác
        public void OnAction(IRibbonControl control)
        {
            if (control.Id.StartsWith("ol"))
            {
                OutlookOps.Do(control.Id, control.Context);
                return;
            }
            try
            {
                string id = control.Id;
                switch (id)
                {
                    case "btnOpen": Run(""); return;
                    case "btnPrefs":
                        using (var d = new PrefsDialog(Prefs.Load(), host)) d.ShowDialog();
                        return;
                    case "btnMailMerge": MailMergeDialog.Start(this, app); return;
                    case "btnImportComments": Comments.Import(this, app, null, false); return;
                    case "btnAcceptAll": if (HasDoc()) app.ActiveDocument.Revisions.AcceptAll(); return;
                    case "btnRejectAll": if (HasDoc()) app.ActiveDocument.Revisions.RejectAll(); return;
                    case "btnDeleteComments":
                        if (HasDoc() && MessageBox.Show(L("Xoá tất cả nhận xét trong tài liệu này?", "Delete all comments in this document?"),
                                                        "PDF Toolkit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            app.ActiveDocument.DeleteAllComments();
                        return;
                }
                bool markup = id == "btnReview" || id == "btnExportMarkup";
                string pdf = Export(markup, true);
                if (pdf == null) return;
                string q = "\"" + pdf + "\"";
                string name = Path.GetFileNameWithoutExtension(pdf);
                switch (id)
                {
                    case "btnCreate":
                    case "btnExportMarkup": After(pdf, Prefs.Load().After); break;
                    case "btnAttach": Mail.Compose(null, name, "", new[] { pdf }, false); break;
                    case "btnReview":
                        Mail.Compose(null, L("Đề nghị xem và nhận xét: ", "Please review and comment: ") + name,
                                     L("Kính gửi anh/chị,\n\nNhờ anh/chị xem file PDF đính kèm và ghi nhận xét trực tiếp trên file " +
                                       "(công cụ Ghi chú / Tô sáng / Gạch xoá trong Adobe Reader, Microsoft Edge hoặc PDF Toolkit), " +
                                       "lưu lại rồi gửi trả file PDF qua email này.\n\nCảm ơn anh/chị.",
                                       "Hello,\n\nPlease review the attached PDF and add your comments directly in the file " +
                                       "(Note / Highlight / Strikethrough tools in Adobe Reader, Microsoft Edge or PDF Toolkit), " +
                                       "save it and send the PDF back by replying to this email.\n\nThank you."),
                                     new[] { pdf }, false);
                        break;
                    case "btnSign": Run("--sign " + q); break;
                    case "btnProtect": Run("--open-tool protect " + q); break;
                    case "btnCompress": Run("--open-tool compress " + q); break;
                    case "btnWatermark": Run("--open-tool watermark " + q); break;
                    case "btnMerge": Run("--open-tool merge " + q); break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(Inner(ex).Message, "PDF Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private bool HasDoc()
        {
            try { return app.Documents.Count > 0; } catch { return false; }
        }

        internal static void After(string pdf, int mode)
        {
            switch (mode)
            {
                case 0: Run("\"" + pdf + "\""); break;
                case 1: Process.Start(new ProcessStartInfo(pdf) { UseShellExecute = true }); break;
                case 2: Process.Start("explorer.exe", "/select,\"" + pdf + "\""); break;
                default:
                    MessageBox.Show(L("Đã tạo PDF:\n", "PDF created:\n") + pdf, "PDF Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    break;
            }
        }

        internal static Exception Inner(Exception e)
        {
            while (e is TargetInvocationException && e.InnerException != null) e = e.InnerException;
            return e;
        }

        internal static string ExePath()
        {
            string env = Environment.GetEnvironmentVariable("PDFTOOLKIT_EXE");     // kiểm thử: chỉ định exe / script
            if (!string.IsNullOrEmpty(env) && File.Exists(env)) return env;
            string dir = Path.GetDirectoryName(new Uri(Assembly.GetExecutingAssembly().CodeBase).LocalPath);
            string exe = Path.Combine(dir, "PDF Toolkit.exe");
            if (File.Exists(exe)) return exe;
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\TuanPA\PDFToolkit\Install"))
            {
                if (k != null)
                {
                    string p = Convert.ToString(k.GetValue("Exe", ""));
                    if (File.Exists(p)) return p;
                }
            }
            throw new FileNotFoundException(L("Chưa cài PDF Toolkit. Mở PDF Toolkit.exe và chọn Cài đặt vào máy.",
                                              "PDF Toolkit is not installed. Run PDF Toolkit.exe and choose Install."));
        }

        internal static void Run(string args)
        {
            Process.Start(new ProcessStartInfo(ExePath(), args) { UseShellExecute = false });
        }

        internal static string Unique(string path)
        {
            if (!File.Exists(path)) return path;
            string dir = Path.GetDirectoryName(path), stem = Path.GetFileNameWithoutExtension(path);
            for (int i = 2; ; i++)
            {
                string p = Path.Combine(dir, stem + " (" + i + ").pdf");
                if (!File.Exists(p)) return p;
            }
        }

        /// Đường dẫn file PDF theo Tuỳ chọn (null nếu người dùng huỷ hộp thoại Lưu).
        private static string Target(string fullPath, string name, Prefs p, bool interactive)
        {
            string stem = Path.GetFileNameWithoutExtension(name);
            foreach (char c in Path.GetInvalidFileNameChars()) stem = stem.Replace(c, '_');
            string dir = "";
            try { dir = Path.GetDirectoryName(fullPath); } catch { }
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir) || dir.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                dir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);    // chưa lưu / file OneDrive web
            if (p.Folder == 1 && Directory.Exists(p.FolderPath)) dir = p.FolderPath;
            if (p.Folder == 2 && interactive)
            {
                using (var d = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", InitialDirectory = dir, FileName = stem + ".pdf",
                                                    Title = "PDF Toolkit" })
                    return d.ShowDialog() == DialogResult.OK ? d.FileName : null;
            }
            return Unique(Path.Combine(dir, stem + ".pdf"));
        }

        /// Xuất tài liệu đang mở ra PDF theo Tuỳ chọn, trả về đường dẫn (null nếu không có tài liệu / huỷ).
        private string Export(bool markup, bool interactive)
        {
            Prefs p = Prefs.Load();
            Cursor old = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                if (host == "word")
                {
                    if (app.Documents.Count == 0) return null;
                    dynamic doc = app.ActiveDocument;
                    string pdf = Target((string)doc.FullName, (string)doc.Name, p, interactive);
                    if (pdf == null) return null;
                    if (markup) p.Markup = true;
                    ExportWord(doc, pdf, p, false);
                    return pdf;
                }
                if (host == "excel")
                {
                    if (app.Workbooks.Count == 0) return null;
                    dynamic wb = app.ActiveWorkbook;
                    string pdf = Target((string)wb.FullName, (string)wb.Name, p, interactive);
                    if (pdf == null) return null;
                    dynamic what = p.Scope == 1 ? app.ActiveSheet : p.Scope == 2 ? app.Selection : wb;
                    // 0 = xlTypePDF; Quality 0 = chuẩn, 1 = nhỏ nhất
                    what.ExportAsFixedFormat(Type: 0, Filename: pdf, Quality: p.Screen ? 1 : 0, IncludeDocProperties: p.Props,
                                             IgnorePrintAreas: false, OpenAfterPublish: false);
                    return pdf;
                }
                if (host == "powerpoint")
                {
                    if (app.Presentations.Count == 0) return null;
                    dynamic pres;
                    try { pres = app.ActivePresentation; }
                    catch { pres = app.Presentations.Item(app.Presentations.Count); }   // không có cửa sổ đang chọn
                    string pdf = Target((string)pres.FullName, (string)pres.Name, p, interactive);
                    if (pdf == null) return null;
                    int slide = 0;
                    if (p.Scope == 1)
                        try { slide = (int)app.ActiveWindow.View.Slide.SlideIndex; } catch { slide = 0; }
                    ExportPowerPoint(pres, pdf, slide);
                    return pdf;
                }
                return null;
            }
            finally { Cursor.Current = old; }
        }

        internal void ExportWord(dynamic doc, string pdf, Prefs p, bool whole)
        {
            int range = 0;                                  // wdExportAllDocument
            if (!whole && p.Scope == 2)
            {
                try { if ((int)app.Selection.Type != 1) range = 1; } catch { }   // 1 = wdSelectionIP (chưa chọn gì)
            }
            else if (!whole && p.Scope == 1) range = 2;     // wdExportCurrentPage
            doc.ExportAsFixedFormat(OutputFileName: pdf, ExportFormat: 17, OpenAfterExport: false,
                                    OptimizeFor: p.Screen ? 1 : 0, Range: range,
                                    Item: p.Markup ? 7 : 0,  // 7 = wdExportDocumentWithMarkup
                                    IncludeDocProps: p.Props, KeepIRM: true,
                                    CreateBookmarks: p.Bookmarks ? 1 : 0, DocStructureTags: p.Tags,
                                    BitmapMissingFonts: true, UseISO19005_1: p.PdfA);
        }

        private static void ExportPowerPoint(dynamic pres, string pdf, int slide)
        {
            if (slide <= 0)
            {
                // ExportAsFixedFormat của PowerPoint có tham số PrintRange (object) không gọi được kiểu late-binding
                pres.SaveCopyAs(pdf, 32);                   // 32 = ppSaveAsPDF, không đổi file đang mở
                return;
            }
            // chỉ 1 slide: lưu bản sao tạm, xoá các slide khác rồi xuất
            string tmp = Path.Combine(Path.GetTempPath(), "pdftk_" + Guid.NewGuid().ToString("N") + ".pptx");
            pres.SaveCopyAs(tmp);
            dynamic app = pres.Application;
            dynamic copy = app.Presentations.Open(tmp, false, false, false);   // ReadOnly, Untitled, WithWindow
            try
            {
                for (int i = (int)copy.Slides.Count; i >= 1; i--)
                    if (i != slide) copy.Slides.Item(i).Delete();
                copy.SaveCopyAs(pdf, 32);
            }
            finally
            {
                copy.Close();
                try { File.Delete(tmp); } catch { }
            }
        }
    }
}

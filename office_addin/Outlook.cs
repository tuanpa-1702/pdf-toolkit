// PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
// SPDX-License-Identifier: AGPL-3.0-only
// Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
// Add-in cho Outlook (bản cổ điển):
//   Đọc thư: Email → PDF (1 hoặc nhiều thư, gộp được, nhúng file đính kèm), lưu file đính kèm thành PDF
//   Soạn thư: chuyển đính kèm Word / Excel / PowerPoint / ảnh sang PDF, nén PDF đính kèm, đặt mật khẩu PDF đính kèm
//   Chuột phải: thư -> Lưu thành PDF; file đính kèm -> Mở bằng PDF Toolkit / Lưu thành PDF
// Việc xử lý PDF chạy bằng "PDF Toolkit.exe --tool ..." (tiến trình riêng), Email → PDF dùng Word (giữ đúng định dạng thư).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PDFToolkitAddin
{
    public static class OutlookUI
    {
        const string Head = "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>";
        const string Btn = "getLabel='GetLabel' getScreentip='GetTip' getImage='GetImage' onAction='OnAction'";

        static string ReadGroup(string id)
        {
            return "<group id='" + id + "' getLabel='GetLabel'>" +
                   "<button id='olMailToPdf' size='large' " + Btn + "/>" +
                   "<button id='olAttToPdf' size='large' " + Btn + "/>" +
                   "<button id='btnOpen' size='large' " + Btn + "/></group>";
        }

        static string ComposeGroup(string id, string suffix)
        {
            return "<group id='" + id + "' getLabel='GetLabel'>" +
                   "<button id='olAttachAsPdf" + suffix + "' size='large' " + Btn + "/>" +
                   "<button id='olCompressAtt" + suffix + "' size='large' " + Btn + "/>" +
                   "<button id='olProtectAtt" + suffix + "' size='large' " + Btn + "/></group>";
        }

        /// Ribbon theo từng loại cửa sổ Outlook; loại khác (lịch, danh bạ...) không thêm gì.
        public static string Xml(string ribbonID)
        {
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                    return Head + "<ribbon><tabs>" +
                           "<tab idMso='TabMail'>" + ReadGroup("olGrpMail") + "</tab>" +
                           "</tabs><contextualTabs><tabSet idMso='TabComposeTools'><tab idMso='TabMessage'>" +   // trả lời ngay trong khung đọc
                           ComposeGroup("olGrpInline", "_i") + "</tab></tabSet></contextualTabs></ribbon><contextMenus>" +
                           "<contextMenu idMso='ContextMenuMailItem'><button id='olCtxMail' " + Btn + "/></contextMenu>" +
                           "<contextMenu idMso='ContextMenuMultipleItems'><button id='olCtxMails' " + Btn + "/></contextMenu>" +
                           "<contextMenu idMso='ContextMenuAttachments'><button id='olCtxAttOpen' " + Btn + "/>" +
                           "<button id='olCtxAttSave' " + Btn + "/></contextMenu>" +
                           "</contextMenus></customUI>";
                case "Microsoft.Outlook.Mail.Read":
                    return Head + "<ribbon><tabs><tab idMso='TabReadMessage'>" + ReadGroup("olGrpRead") +
                           "</tab></tabs></ribbon></customUI>";
                case "Microsoft.Outlook.Mail.Compose":
                    return Head + "<ribbon><tabs><tab idMso='TabNewMailMessage'>" + ComposeGroup("olGrpCompose", "") +
                           "</tab></tabs></ribbon></customUI>";
            }
            return "";
        }
    }

    /// Thao tác của add-in Outlook. Đối tượng Outlook dùng kiểu dynamic (late-binding) - kiểm thử được bằng đối tượng giả.
    public static class OutlookOps
    {
        static readonly string[] OfficeExt = { ".doc", ".docx", ".docm", ".rtf", ".odt", ".dot", ".dotx", ".xls", ".xlsx", ".xlsm",
                                               ".xlsb", ".ods", ".ppt", ".pptx", ".pptm", ".odp", ".pps", ".ppsx" };
        static readonly string[] ImageExt = { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp" };
        const int ClsExplorer = 34, ClsInspector = 35, ClsMail = 43;
        static bool busy;

        /// Cho kiểm thử: hộp thoại chọn nơi lưu / mật khẩu / hỏi gộp được thay bằng hàm trả lời sẵn.
        public static Func<string, string, string> AskSaveFile = DefaultAskSaveFile;
        public static Func<string, string> AskFolder = DefaultAskFolder;
        public static Func<int, DialogResult> AskMerge = DefaultAskMerge;
        public static Func<string> AskPassword = PasswordDialog.Ask;
        public static Action<string, MessageBoxIcon> Notify = (m, icon) => MessageBox.Show(m, "PDF Toolkit", MessageBoxButtons.OK, icon);
        public static Action<string> Open = pdf => Connect.After(pdf, Prefs.Load().After);
        public static Action<string> Reveal = p => Process.Start("explorer.exe", "/select,\"" + p + "\"");
        public static bool UseWord = true, ShowBusy = true;

        static string L(string vi, string en) { return Connect.L(vi, en); }

        public static void Do(string id, object context)
        {
            if (busy) return;
            busy = true;
            try
            {
                string op = id.EndsWith("_i") ? id.Substring(0, id.Length - 2) : id;
                switch (op)
                {
                    case "olMailToPdf": case "olCtxMail": case "olCtxMails": MailToPdf(Items(context)); break;
                    case "olAttToPdf": SaveAttachments(Items(context).SelectMany(m => Attachments(m, false)).ToList()); break;
                    case "olCtxAttOpen": OpenAttachments(Selected(context)); break;
                    case "olCtxAttSave": SaveAttachments(Selected(context)); break;
                    case "olAttachAsPdf": ComposeOp(Draft(context), "pdf"); break;
                    case "olCompressAtt": ComposeOp(Draft(context), "compress"); break;
                    case "olProtectAtt": ComposeOp(Draft(context), "protect"); break;
                }
            }
            catch (Exception ex)
            {
                Notify(Connect.Inner(ex).Message, MessageBoxIcon.Warning);
            }
            finally { busy = false; Busy.Hide(); }
        }

        // ------------------------------------------------------------------ lấy thư / file đính kèm
        static int Cls(dynamic o)
        {
            try { return (int)o.Class; } catch { return -1; }
        }

        /// Thư đang chọn (cửa sổ chính / menu chuột phải) hoặc thư đang mở (cửa sổ đọc thư).
        public static List<object> Items(object context)
        {
            var list = new List<object>();
            dynamic ctx = context;
            int c = Cls(ctx);
            if (c == ClsInspector) { list.Add(ctx.CurrentItem); return list; }
            dynamic sel = c == ClsExplorer ? ctx.Selection : ctx;
            int n = 0;
            try { n = (int)sel.Count; } catch { }
            for (int i = 1; i <= n; i++)
            {
                dynamic it = sel.Item(i);
                if (Cls(it) == ClsMail || HasSaveAs(it)) list.Add(it);
            }
            return list;
        }

        static bool HasSaveAs(dynamic it)
        {
            int c = Cls(it);
            return c == 43 || c == 46 || c == 53 || c == 54 || c == 55 || c == 56 || c == 57;   // thư, báo cáo, thư mời họp...
        }

        /// Thư đang soạn: cửa sổ soạn thư riêng hoặc trả lời ngay trong khung đọc.
        static dynamic Draft(object context)
        {
            dynamic ctx = context;
            int c = Cls(ctx);
            if (c == ClsInspector) return ctx.CurrentItem;
            if (c == ClsExplorer) return ctx.ActiveInlineResponse;
            return ctx;
        }

        static List<object> Selected(object context)
        {
            var list = new List<object>();
            dynamic sel = context;
            int n = 0;
            try { n = (int)sel.Count; } catch { }
            for (int i = 1; i <= n; i++) list.Add(sel.Item(i));
            return list;
        }

        /// File đính kèm thật của thư (bỏ ảnh nhúng trong nội dung thư). convertibleOnly: chỉ PDF / Office / ảnh.
        public static List<object> Attachments(object item, bool convertibleOnly)
        {
            dynamic mail = item;
            var list = new List<object>();
            dynamic atts;
            try { atts = mail.Attachments; } catch { return list; }
            int n = (int)atts.Count;
            for (int i = 1; i <= n; i++)
            {
                dynamic a = atts.Item(i);
                if (IsInline(a)) continue;
                if (convertibleOnly && Kind(Name(a)) == "") continue;
                list.Add(a);
            }
            return list;
        }

        static bool IsInline(dynamic a)
        {
            try { if ((int)a.Type != 1) return (int)a.Type == 6; } catch { }           // 1 = file, 5 = thư đính kèm, 6 = OLE
            try
            {
                dynamic pa = a.PropertyAccessor;
                object hidden = null;
                try { hidden = pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"); } catch { }
                if (hidden is bool && (bool)hidden) return true;
                string cid = "";
                try { cid = Convert.ToString(pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")); } catch { }
                if (cid != "")
                {
                    string body = "";
                    try { body = Convert.ToString(a.Parent.HTMLBody); } catch { }
                    return body.IndexOf("cid:" + cid, StringComparison.OrdinalIgnoreCase) >= 0;
                }
            }
            catch { }
            return false;
        }

        static string Name(object o)
        {
            dynamic a = o;
            string n = "";
            try { n = Convert.ToString(a.FileName); } catch { }
            if (string.IsNullOrEmpty(n)) try { n = Convert.ToString(a.DisplayName); } catch { }
            return string.IsNullOrEmpty(n) ? "file" : n;
        }

        /// "pdf" / "office" / "image" / "" (không chuyển được)
        public static string Kind(string name)
        {
            string ext = Path.GetExtension(name).ToLowerInvariant();
            if (ext == ".pdf") return "pdf";
            if (OfficeExt.Contains(ext)) return "office";
            if (ImageExt.Contains(ext)) return "image";
            return "";
        }

        static string SaveAtt(dynamic a, string dir)
        {
            Directory.CreateDirectory(dir);
            string p = Path.Combine(dir, Safe(Name(a)));
            a.SaveAsFile(p);
            return p;
        }

        // ------------------------------------------------------------------ Email → PDF
        static void MailToPdf(List<object> items)
        {
            if (items.Count == 0)
            {
                Notify(L("Hãy chọn email cần chuyển sang PDF.", "Select the emails to convert to PDF."), MessageBoxIcon.Information);
                return;
            }
            string target = null, folder;
            bool merge = false;
            if (items.Count == 1)
            {
                target = AskSaveFile(Subject(items[0]) + ".pdf", LastFolder());
                if (target == null) return;
                folder = Path.GetDirectoryName(target);
            }
            else
            {
                DialogResult r = AskMerge(items.Count);
                if (r == DialogResult.Cancel) return;
                merge = r == DialogResult.Yes;
                folder = AskFolder(LastFolder());
                if (folder == null) return;
            }
            LastFolder(folder);
            string tmp = TempDir();
            var pdfs = new List<string>();
            dynamic word = null;
            try
            {
                int k = 0;
                foreach (object m in items)
                {
                    k++;
                    Busy.Show(L("Đang chuyển email sang PDF", "Converting emails to PDF") + (items.Count > 1 ? " (" + k + "/" + items.Count + ")" : "") + "...");
                    pdfs.Add(OneMail(m, Path.Combine(tmp, k.ToString()), ref word));
                }
            }
            finally
            {
                if (word != null) try { word.Quit(0); } catch { }
            }
            var results = new List<string>();
            if (merge && pdfs.Count > 1)
            {
                Busy.Show(L("Đang gộp PDF...", "Merging PDFs..."));
                string merged = Tool.Run("merge", pdfs, null, null).First();
                string dst = Connect.Unique(Path.Combine(folder, Safe(Subject(items[0])) + " (+" + (pdfs.Count - 1) + ").pdf"));
                File.Copy(merged, dst);
                results.Add(dst);
            }
            else if (target != null)
            {
                File.Copy(pdfs[0], target, true);
                results.Add(target);
            }
            else
            {
                foreach (string p in pdfs)
                {
                    string dst = Connect.Unique(Path.Combine(folder, Path.GetFileName(p)));
                    File.Copy(p, dst);
                    results.Add(dst);
                }
            }
            TryDelete(tmp);
            Busy.Hide();
            if (results.Count == 1) Open(results[0]);
            else Reveal(results[0]);
        }

        static string Subject(object item)
        {
            string s = "";
            try { s = Convert.ToString(((dynamic)item).Subject); } catch { }
            s = Safe(s).Trim();
            if (s.Length > 80) s = s.Substring(0, 80).Trim();
            return s == "" ? "Email" : s;
        }

        /// Chuyển 1 thư sang PDF trong thư mục dir, nhúng các file đính kèm. Trả về đường dẫn PDF.
        static string OneMail(object mo, string dir, ref dynamic word)
        {
            dynamic m = mo;
            Directory.CreateDirectory(dir);
            string stem = Subject(m);
            string pdf = Path.Combine(dir, stem + ".pdf");
            bool ok = false;
            if (UseWord)
            {
                string mht = Path.Combine(dir, stem + ".mht");
                m.SaveAs(mht, 10);                                   // 10 = olMHTML: có tiêu đề thư (Từ, Gửi lúc, Đến...) và ảnh
                ok = WordToPdf(mht, pdf, ref word);
            }
            if (!ok)                                                 // không có Word -> HTML rồi dàn trang bằng PDF Toolkit
            {
                string html = Path.Combine(dir, stem + ".html");
                m.SaveAs(html, 5);                                   // 5 = olHTML
                File.Copy(Tool.Run("text_to_pdf", new[] { html }, null, null).First(), pdf, true);
            }
            var atts = Attachments(m, false);
            if (atts.Count > 0)
            {
                var files = new List<string>();
                int i = 0;
                foreach (dynamic a in atts)
                    try { files.Add(SaveAtt(a, Path.Combine(dir, "att" + (++i)))); } catch { }
                if (files.Count > 0)
                    File.Copy(Tool.Run("add_attach", new[] { pdf }, new Dictionary<string, object> { { "attach", files } }, null).First(), pdf, true);
            }
            return pdf;
        }

        static bool WordToPdf(string src, string pdf, ref dynamic word)
        {
            try
            {
                if (word == null)
                {
                    Type t = Type.GetTypeFromProgID("Word.Application");
                    if (t == null) return false;
                    word = Activator.CreateInstance(t);              // phiên Word riêng, chạy ẩn
                    word.Visible = false;
                    word.DisplayAlerts = 0;
                }
                dynamic d = word.Documents.Open(FileName: src, ConfirmConversions: false, ReadOnly: true, AddToRecentFiles: false,
                                                Visible: false);
                try
                {
                    try { d.ActiveWindow.View.Type = 3; } catch { }  // 3 = wdPrintView: dàn trang như khi in
                    d.ExportAsFixedFormat(OutputFileName: pdf, ExportFormat: 17, OpenAfterExport: false, OptimizeFor: 0,
                                          CreateBookmarks: 0, DocStructureTags: true);
                }
                finally { d.Close(0); }
                return File.Exists(pdf);
            }
            catch { return false; }
        }

        // ------------------------------------------------------------------ file đính kèm (thư đã nhận)
        static void OpenAttachments(List<object> atts)
        {
            var ok = atts.Where(a => Kind(Name(a)) != "").ToList();
            if (ok.Count == 0)
            {
                Notify(L("Chỉ mở được file PDF, Word, Excel, PowerPoint hoặc ảnh.", "Only PDF, Word, Excel, PowerPoint or image files can be opened."),
                       MessageBoxIcon.Information);
                return;
            }
            Busy.Show(L("Đang mở file đính kèm...", "Opening attachments..."));
            foreach (string pdf in ToPdf(ok, TempDir()))
                Open(pdf);
        }

        static void SaveAttachments(List<object> atts)
        {
            var ok = atts.Where(a => Kind(Name(a)) != "").ToList();
            if (ok.Count == 0)
            {
                Notify(L("Không có file đính kèm nào chuyển được sang PDF (PDF, Word, Excel, PowerPoint, ảnh).",
                         "There are no attachments that can be saved as PDF (PDF, Word, Excel, PowerPoint, images)."), MessageBoxIcon.Information);
                return;
            }
            string folder = AskFolder(LastFolder());
            if (folder == null) return;
            LastFolder(folder);
            Busy.Show(L("Đang chuyển file đính kèm sang PDF...", "Converting attachments to PDF..."));
            string tmp = TempDir();
            var saved = new List<string>();
            foreach (string pdf in ToPdf(ok, tmp))
            {
                string dst = Connect.Unique(Path.Combine(folder, Path.GetFileName(pdf)));
                File.Copy(pdf, dst);
                saved.Add(dst);
            }
            TryDelete(tmp);
            Busy.Hide();
            if (saved.Count == 1) Open(saved[0]);
            else if (saved.Count > 1) Reveal(saved[0]);
        }

        /// Lưu file đính kèm ra thư mục tạm và chuyển sang PDF (PDF giữ nguyên). Trả về PDF theo thứ tự file đính kèm.
        static List<string> ToPdf(List<object> atts, string tmp)
        {
            var files = new List<string>();
            int i = 0;
            foreach (dynamic a in atts) files.Add(SaveAtt(a, Path.Combine(tmp, (++i).ToString())));
            return ConvertFiles(files, tmp);
        }

        /// PDF giữ nguyên, Office -> office_to_pdf, ảnh -> images_to_pdf (mỗi ảnh 1 file). Tên PDF = tên file gốc.
        static List<string> ConvertFiles(List<string> files, string tmp)
        {
            var result = new string[files.Count];
            foreach (string kind in new[] { "office", "image" })
            {
                var idx = Enumerable.Range(0, files.Count).Where(j => Kind(files[j]) == kind).ToList();
                if (idx.Count == 0) continue;
                var opts = kind == "image" ? new Dictionary<string, object> { { "merge", false }, { "size", "A4" }, { "margin", 10 } } : null;
                var outs = Tool.Run(kind == "office" ? "office_to_pdf" : "images_to_pdf", idx.Select(j => files[j]), opts, null);
                Tool.Match(idx.Select(j => files[j]).ToList(), outs, (j, o) => result[idx[j]] = o);
            }
            var list = new List<string>();
            for (int j = 0; j < files.Count; j++)
            {
                string src = Kind(files[j]) == "pdf" ? files[j] : result[j];
                if (src == null) continue;
                string dir = Path.Combine(tmp, "pdf" + j);
                Directory.CreateDirectory(dir);
                string dst = Path.Combine(dir, Path.GetFileNameWithoutExtension(files[j]) + ".pdf");
                File.Copy(src, dst, true);
                list.Add(dst);
            }
            return list;
        }

        // ------------------------------------------------------------------ thư đang soạn
        static void ComposeOp(dynamic mail, string op)
        {
            if (mail == null) return;
            string want = op == "pdf" ? null : "pdf";
            List<object> all = Attachments((object)mail, true);
            var atts = all.Where(a => want == null ? Kind(Name(a)) != "pdf" : Kind(Name(a)) == "pdf")
                                              .ToList();
            if (atts.Count == 0)
            {
                Notify(op == "pdf"
                    ? L("Thư chưa có file Word, Excel, PowerPoint hoặc ảnh đính kèm để chuyển sang PDF.",
                        "This email has no Word, Excel, PowerPoint or image attachments to convert to PDF.")
                    : L("Thư chưa có file PDF đính kèm.", "This email has no PDF attachments."), MessageBoxIcon.Information);
                return;
            }
            string pw = null;
            if (op == "protect")
            {
                pw = AskPassword();
                if (string.IsNullOrEmpty(pw)) return;
            }
            Busy.Show(op == "pdf" ? L("Đang chuyển file đính kèm sang PDF...", "Converting attachments to PDF...")
                    : op == "compress" ? L("Đang nén PDF đính kèm...", "Compressing PDF attachments...")
                    : L("Đang đặt mật khẩu cho PDF đính kèm...", "Protecting PDF attachments..."));
            string tmp = TempDir();
            var files = new List<string>();
            int i = 0;
            foreach (dynamic a in atts) files.Add(SaveAtt(a, Path.Combine(tmp, (++i).ToString())));
            var outs = new string[files.Count];
            if (op == "pdf")
            {
                var pdfs = ConvertFiles(files, tmp);
                for (int j = 0, k = 0; j < files.Count && k < pdfs.Count; j++)          // ConvertFiles bỏ qua file lỗi -> khớp theo tên
                    if (Path.GetFileNameWithoutExtension(pdfs[k]) == Path.GetFileNameWithoutExtension(files[j])) outs[j] = pdfs[k++];
            }
            else
            {
                var opts = op == "compress" ? new Dictionary<string, object> { { "mode", "medium" } }
                                            : new Dictionary<string, object> { { "p_print", true }, { "p_copy", true }, { "p_modify", true },
                                                                               { "p_annot", true } };
                var env = op == "protect" ? new Dictionary<string, string> { { "user_pw", pw } } : null;
                var res = Tool.Run(op, files, opts, env);
                Tool.Match(files, res, (j, o) => outs[j] = o);
            }
            long before = 0, after = 0;
            int done = 0;
            var report = new StringBuilder();
            for (int j = 0; j < files.Count; j++)
            {
                if (outs[j] == null) continue;
                long b = new FileInfo(files[j]).Length, n = new FileInfo(outs[j]).Length;
                if (op == "compress" && n >= b)
                {
                    report.AppendLine("• " + Path.GetFileName(files[j]) + ": " + L("đã gọn sẵn, giữ nguyên", "already compact, unchanged"));
                    continue;
                }
                string name = op == "pdf" ? Path.GetFileNameWithoutExtension(files[j]) + ".pdf" : Path.GetFileName(files[j]);
                string dir = Path.Combine(tmp, "new" + j);
                Directory.CreateDirectory(dir);
                string dst = Path.Combine(dir, name);
                File.Copy(outs[j], dst, true);
                dynamic a = atts[j];
                a.Delete();
                mail.Attachments.Add(dst, 1);                         // 1 = olByValue
                before += b; after += n; done++;
                if (op == "compress") report.AppendLine("• " + name + ": " + Size(b) + " → " + Size(n));
            }
            TryDelete(tmp);
            Busy.Hide();
            string msg;
            if (op == "pdf")
                msg = L("Đã chuyển " + done + "/" + files.Count + " file đính kèm sang PDF.", "Converted " + done + "/" + files.Count + " attachments to PDF.");
            else if (op == "compress")
                msg = L("Đã nén " + done + "/" + files.Count + " file PDF", "Compressed " + done + "/" + files.Count + " PDF files") +
                      (done > 0 ? ": " + Size(before) + " → " + Size(after) : "") + "\n\n" + report;
            else
                msg = L("Đã đặt mật khẩu cho " + done + "/" + files.Count + " file PDF đính kèm.\n\nHãy gửi mật khẩu cho người nhận qua kênh khác (tin nhắn, điện thoại) - không ghi trong thư này.",
                        "Protected " + done + "/" + files.Count + " PDF attachments with a password.\n\nSend the password to recipients another way (chat, phone) - not in this email.");
            Notify(msg.TrimEnd(), done == files.Count || op == "compress" ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }

        static string Size(long b)
        {
            return b >= 1048576 ? (b / 1048576.0).ToString("0.0") + " MB" : Math.Max(1, b / 1024) + " KB";
        }

        // ------------------------------------------------------------------ tiện ích
        internal static string Safe(string s)
        {
            s = s ?? "";
            foreach (char c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
            return s;
        }

        internal static string TempDir()
        {
            string d = Path.Combine(Path.GetTempPath(), "PDFToolkit_mail", Guid.NewGuid().ToString("N").Substring(0, 12));
            Directory.CreateDirectory(d);
            return d;
        }

        static void TryDelete(string dir)
        {
            try { Directory.Delete(dir, true); } catch { }
        }

        const string Key = @"Software\TuanPA\PDFToolkit\Addin";

        static string LastFolder()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(Key))
                {
                    string p = k == null ? "" : Convert.ToString(k.GetValue("MailFolder", ""));
                    if (Directory.Exists(p)) return p;
                }
            }
            catch { }
            return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        static void LastFolder(string p)
        {
            try { using (var k = Registry.CurrentUser.CreateSubKey(Key)) k.SetValue("MailFolder", p); } catch { }
        }

        static string DefaultAskSaveFile(string name, string dir)
        {
            using (var d = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf", InitialDirectory = dir, FileName = name, Title = "PDF Toolkit" })
                return d.ShowDialog() == DialogResult.OK ? d.FileName : null;
        }

        static string DefaultAskFolder(string dir)
        {
            using (var d = new FolderBrowserDialog { SelectedPath = dir, Description = L("Chọn thư mục lưu PDF", "Choose a folder for the PDFs") })
                return d.ShowDialog() == DialogResult.OK ? d.SelectedPath : null;
        }

        static DialogResult DefaultAskMerge(int n)
        {
            return MessageBox.Show(L("Gộp " + n + " email thành 1 file PDF?\n\nYes: gộp thành 1 file\nNo: mỗi email 1 file PDF",
                                     "Combine " + n + " emails into one PDF?\n\nYes: one combined file\nNo: one PDF per email"),
                                   "PDF Toolkit", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        }
    }

    /// Chạy "PDF Toolkit.exe --tool <id>" ẩn, chờ xong (giao diện Outlook vẫn vẽ lại), trả về danh sách file kết quả.
    public static class Tool
    {
        public static int TimeoutMs = 15 * 60 * 1000;

        public static List<string> Run(string tool, IEnumerable<string> files, IDictionary<string, object> opts,
                                       IDictionary<string, string> secretOpts)
        {
            string outDir = OutlookOps.TempDir();
            var json = new JavaScriptSerializer();
            var args = new StringBuilder("--tool " + tool + " --out " + Q(outDir));
            if (opts != null)
                foreach (var kv in opts) args.Append(" --opt ").Append(Q(kv.Key + "=" + json.Serialize(kv.Value)));
            var psi = new ProcessStartInfo(Connect.ExePath()) { UseShellExecute = false, CreateNoWindow = true };
            if (secretOpts != null)                                  // mật khẩu: qua biến môi trường, không lộ trên dòng lệnh
            {
                int i = 0;
                foreach (var kv in secretOpts)
                {
                    string env = "PDFTK_SECRET_" + (++i);
                    psi.EnvironmentVariables[env] = kv.Value;
                    args.Append(" --opt-env ").Append(Q(kv.Key + "=" + env));
                }
            }
            foreach (string f in files) args.Append(' ').Append(Q(f));
            psi.Arguments = args.ToString();
            using (var p = Process.Start(psi))
            {
                var sw = Stopwatch.StartNew();
                while (!p.WaitForExit(100))
                {
                    Application.DoEvents();
                    if (sw.ElapsedMilliseconds > TimeoutMs)
                    {
                        try { p.Kill(); } catch { }
                        throw new TimeoutException(Connect.L("PDF Toolkit xử lý quá lâu - đã dừng.", "PDF Toolkit took too long - stopped."));
                    }
                }
            }
            string log = Path.Combine(outDir, "pdf_toolkit_cli.log");
            var lines = File.Exists(log) ? File.ReadAllLines(log, Encoding.UTF8) : new string[0];
            var outs = lines.Where(l => l.StartsWith("OK ")).Select(l => l.Substring(3)).Where(File.Exists).ToList();
            if (outs.Count == 0)
            {
                string err = string.Join("\n", lines.Where(l => l.StartsWith("ERROR ")).Select(l => l.Substring(6)));
                throw new InvalidOperationException(Connect.L("PDF Toolkit không xử lý được file", "PDF Toolkit could not process the file") +
                                                    (err != "" ? ":\n" + err : "."));
            }
            return outs;
        }

        /// Ghép file kết quả với file gốc theo tên (kết quả = tên gốc + hậu tố). set(chỉ số file gốc, file kết quả).
        public static void Match(List<string> inputs, List<string> outputs, Action<int, string> set)
        {
            var left = new List<string>(outputs);
            foreach (int j in Enumerable.Range(0, inputs.Count).OrderByDescending(j => Path.GetFileNameWithoutExtension(inputs[j]).Length))
            {
                string stem = Path.GetFileNameWithoutExtension(inputs[j]);
                string o = left.FirstOrDefault(x => Path.GetFileName(x).StartsWith(stem, StringComparison.OrdinalIgnoreCase));
                if (o == null && inputs.Count == 1 && left.Count == 1) o = left[0];
                if (o == null) continue;
                left.Remove(o);
                set(j, o);
            }
        }

        /// Đặt 1 tham số dòng lệnh trong ngoặc kép theo quy tắc của Windows (CommandLineToArgvW).
        public static string Q(string s)
        {
            if (s.Length > 0 && s.IndexOfAny(new[] { ' ', '\t', '"' }) < 0) return s;
            var sb = new StringBuilder("\"");
            int bs = 0;
            foreach (char c in s)
            {
                if (c == '\\') { bs++; continue; }
                if (c == '"') { sb.Append('\\', bs * 2 + 1); sb.Append('"'); }
                else { sb.Append('\\', bs); sb.Append(c); }
                bs = 0;
            }
            sb.Append('\\', bs * 2);
            sb.Append('"');
            return sb.ToString();
        }
    }

    /// Cửa sổ nhỏ "Đang xử lý..." hiện trong lúc chờ PDF Toolkit / Word.
    internal static class Busy
    {
        static Form form;
        static Label label;

        public static void Show(string text)
        {
            if (!OutlookOps.ShowBusy) return;
            if (form == null || form.IsDisposed)
            {
                form = new Form { FormBorderStyle = FormBorderStyle.FixedToolWindow, ShowInTaskbar = false, TopMost = true,
                                  StartPosition = FormStartPosition.CenterScreen, Text = "PDF Toolkit", ControlBox = false,
                                  ClientSize = new Size(380, 86), Font = new Font("Segoe UI", 9.5f) };
                label = new Label { Left = 16, Top = 14, Width = 348, Height = 22 };
                var bar = new ProgressBar { Left = 16, Top = 46, Width = 348, Height = 16, Style = ProgressBarStyle.Marquee };
                form.Controls.Add(label);
                form.Controls.Add(bar);
            }
            label.Text = text;
            if (!form.Visible) form.Show();
            form.Refresh();
            Application.DoEvents();
        }

        public static void Hide()
        {
            if (form != null && !form.IsDisposed) { form.Close(); form.Dispose(); }
            form = null;
        }
    }

    /// Hỏi mật khẩu mở file (nhập 2 lần).
    internal class PasswordDialog : Form
    {
        readonly TextBox pw1 = new TextBox { UseSystemPasswordChar = true }, pw2 = new TextBox { UseSystemPasswordChar = true };

        PasswordDialog()
        {
            bool en = Connect.English();
            Func<string, string, string> L = (vi, e) => en ? e : vi;
            Text = L("Đặt mật khẩu cho PDF đính kèm", "Protect PDF attachments");
            Font = new Font("Segoe UI", 9.5f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(400, 210);
            var info = new Label { Left = 16, Top = 12, Width = 368, Height = 40,
                                   Text = L("Người nhận phải nhập mật khẩu mới mở được file (mã hoá AES-256).",
                                            "Recipients must enter the password to open the files (AES-256 encryption).") };
            var l1 = new Label { Left = 16, Top = 58, Width = 130, Text = L("Mật khẩu:", "Password:") };
            var l2 = new Label { Left = 16, Top = 92, Width = 130, Text = L("Nhập lại:", "Confirm:") };
            pw1.SetBounds(150, 55, 234, 26);
            pw2.SetBounds(150, 89, 234, 26);
            var show = new CheckBox { Left = 150, Top = 120, Width = 234, Text = L("Hiện mật khẩu", "Show password") };
            show.CheckedChanged += (s, e) => { pw1.UseSystemPasswordChar = pw2.UseSystemPasswordChar = !show.Checked; };
            var ok = new Button { Text = "OK", Left = 214, Top = 166, Width = 80 };
            var cancel = new Button { Text = L("Huỷ", "Cancel"), Left = 304, Top = 166, Width = 80, DialogResult = DialogResult.Cancel };
            ok.Click += (s, e) =>
            {
                if (pw1.Text == "") { MessageBox.Show(this, L("Chưa nhập mật khẩu.", "Enter a password."), Text); return; }
                if (pw1.Text != pw2.Text) { MessageBox.Show(this, L("Hai lần nhập không khớp.", "The passwords do not match."), Text); return; }
                DialogResult = DialogResult.OK;
            };
            Controls.AddRange(new Control[] { info, l1, l2, pw1, pw2, show, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;
        }

        public static string Ask()
        {
            using (var d = new PasswordDialog())
                return d.ShowDialog() == DialogResult.OK ? d.pw1.Text : null;
        }
    }
}

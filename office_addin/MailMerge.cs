// PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
// SPDX-License-Identifier: AGPL-3.0-only
// Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
// Trộn thư (Mail Merge) của Word -> mỗi bản ghi một file PDF, tuỳ chọn gửi email riêng cho từng người qua Outlook.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace PDFToolkitAddin
{
    internal class MailMergeDialog : Form
    {
        readonly Connect owner;
        readonly dynamic app, doc;
        readonly List<string> fields = new List<string>();
        readonly TextBox nameTpl = new TextBox(), folder = new TextBox(), subject = new TextBox(), body = new TextBox();
        readonly ComboBox insertField = new ComboBox(), toField = new ComboBox();
        readonly CheckBox email = new CheckBox();
        readonly RadioButton draft = new RadioButton(), sendNow = new RadioButton();
        readonly ProgressBar bar = new ProgressBar();
        readonly Label status = new Label();
        readonly Button run = new Button(), close = new Button();
        bool running, cancel, quiet;
        int lastDone;

        public static int RunQuiet(Connect owner, dynamic app, string nameTemplate, string folder)
        {
            using (var d = new MailMergeDialog(owner, app, app.ActiveDocument))
            {
                d.quiet = true;
                d.nameTpl.Text = nameTemplate; d.folder.Text = folder; d.email.Checked = false;
                d.Run();
                return d.lastDone;
            }
        }
        static bool En { get { return Connect.English(); } }
        static string L(string vi, string en) { return En ? en : vi; }

        public static void Start(Connect owner, dynamic app)
        {
            if (app.Documents.Count == 0) return;
            dynamic doc = app.ActiveDocument;
            int type = -1;
            try { type = (int)doc.MailMerge.MainDocumentType; } catch { }
            bool hasData = false;
            try { hasData = (int)doc.MailMerge.State >= 1 && doc.MailMerge.DataSource.FieldNames.Count > 0; } catch { }
            if (type == -1 || !hasData)                     // -1 = wdNotAMergeDocument
            {
                MessageBox.Show(L("Tài liệu này chưa thiết lập trộn thư.\n\nVào Mailings > Start Mail Merge > Letters, chọn danh sách người nhận " +
                                  "(Select Recipients) và chèn các trường (Insert Merge Field), rồi bấm lại nút này.",
                                  "This document is not set up for mail merge.\n\nUse Mailings > Start Mail Merge > Letters, Select Recipients " +
                                  "and Insert Merge Field, then click this button again."),
                                "PDF Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            using (var d = new MailMergeDialog(owner, app, doc)) d.ShowDialog();
        }

        MailMergeDialog(Connect owner, dynamic app, dynamic doc)
        {
            this.owner = owner; this.app = app; this.doc = doc;
            foreach (dynamic f in doc.MailMerge.DataSource.FieldNames) fields.Add((string)f.Name);
            Text = L("Trộn thư ra PDF - PDF Toolkit", "Mail merge to PDF - PDF Toolkit");
            Font = new Font("Segoe UI", 9.5f);
            FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen; AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(600, 560);

            var t = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(14), ColumnCount = 2, AutoSize = false };
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150));
            t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Controls.Add(t);

            string first = fields.Count > 0 ? fields[0] : "";
            nameTpl.Text = "{" + first + "}"; nameTpl.Dock = DockStyle.Fill;
            Row(t, L("Tên file PDF:", "PDF file name:"), nameTpl);
            insertField.DropDownStyle = ComboBoxStyle.DropDownList; insertField.Width = 260;
            insertField.Items.Add(L("Chèn trường vào tên file...", "Insert a field into the name..."));
            foreach (string f in fields) insertField.Items.Add(f);
            insertField.SelectedIndex = 0;
            insertField.SelectedIndexChanged += (s, e) =>
            {
                if (insertField.SelectedIndex > 0) { nameTpl.SelectedText = "{" + insertField.SelectedItem + "}"; insertField.SelectedIndex = 0; }
            };
            Row(t, "", insertField);

            string dir = "";
            try { dir = Path.GetDirectoryName((string)doc.FullName); } catch { }
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) dir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            folder.Text = Path.Combine(dir, "PDF " + Path.GetFileNameWithoutExtension((string)doc.Name));
            var fr = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Dock = DockStyle.Fill };
            folder.Width = 330;
            var browse = new Button { Text = L("Chọn...", "Browse..."), AutoSize = true };
            browse.Click += (s, e) =>
            {
                using (var d = new FolderBrowserDialog()) if (d.ShowDialog(this) == DialogResult.OK) folder.Text = d.SelectedPath;
            };
            fr.Controls.Add(folder); fr.Controls.Add(browse);
            Row(t, L("Lưu vào thư mục:", "Save to folder:"), fr);

            email.Text = L("Gửi email riêng cho từng người, đính kèm PDF của họ (Outlook)", "Email each recipient their own PDF (Outlook)");
            email.AutoSize = true;
            t.Controls.Add(email, 0, t.RowCount); t.SetColumnSpan(email, 2); t.RowCount++;
            toField.DropDownStyle = ComboBoxStyle.DropDownList; toField.Width = 260;
            foreach (string f in fields) toField.Items.Add(f);
            int guess = fields.FindIndex(f => Regex.IsMatch(f, "mail", RegexOptions.IgnoreCase));
            if (toField.Items.Count > 0) toField.SelectedIndex = Math.Max(0, guess);
            Row(t, L("Trường email:", "Email field:"), toField);
            subject.Text = Path.GetFileNameWithoutExtension((string)doc.Name); subject.Dock = DockStyle.Fill;
            Row(t, L("Tiêu đề:", "Subject:"), subject);
            body.Multiline = true; body.Height = 110; body.Dock = DockStyle.Fill; body.ScrollBars = ScrollBars.Vertical;
            body.Text = L("Kính gửi {" + first + "},\r\n\r\nVui lòng xem file đính kèm.\r\n\r\nTrân trọng.",
                          "Dear {" + first + "},\r\n\r\nPlease find the attached file.\r\n\r\nBest regards.");
            Row(t, L("Nội dung:", "Message:"), body);
            var hint = new Label { Text = L("Dùng {Tên trường} trong tên file, tiêu đề và nội dung để chèn dữ liệu của từng người.",
                                            "Use {Field name} in the file name, subject and message to insert each recipient's data."),
                                   AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(420, 0) };
            Row(t, "", hint);
            var rr = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            draft.Text = L("Mở thư nháp để xem trước", "Open drafts to review"); draft.AutoSize = true; draft.Checked = true;
            sendNow.Text = L("Gửi ngay", "Send immediately"); sendNow.AutoSize = true;
            rr.Controls.Add(draft); rr.Controls.Add(sendNow);
            Row(t, L("Cách gửi:", "Sending:"), rr);
            EventHandler sync = (s, e) => { toField.Enabled = subject.Enabled = body.Enabled = draft.Enabled = sendNow.Enabled = email.Checked; };
            email.CheckedChanged += sync; sync(null, null);
            if (!Mail.Available()) { email.Enabled = false; email.Text += L(" - không tìm thấy Outlook", " - Outlook not found"); }

            bar.Dock = DockStyle.Fill; bar.Height = 18;
            t.Controls.Add(bar, 0, t.RowCount); t.SetColumnSpan(bar, 2); t.RowCount++;
            status.AutoSize = true;
            t.Controls.Add(status, 0, t.RowCount); t.SetColumnSpan(status, 2); t.RowCount++;

            var btns = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(10, 6, 10, 6) };
            close.Text = L("Đóng", "Close"); close.AutoSize = true;
            close.Click += (s, e) => { if (running) cancel = true; else Close(); };
            run.Text = L("Tạo PDF", "Create PDFs"); run.AutoSize = true; run.MinimumSize = new Size(100, 0);
            run.Click += (s, e) => Run();
            btns.Controls.Add(close); btns.Controls.Add(run);
            Controls.Add(btns);
            AcceptButton = run;
        }

        static void Row(TableLayoutPanel t, string label, Control c)
        {
            t.Controls.Add(new Label { Text = label, AutoSize = true, Margin = new Padding(0, 7, 6, 0) }, 0, t.RowCount);
            t.Controls.Add(c, 1, t.RowCount);
            t.RowCount++;
        }

        static string Fill(string tpl, Dictionary<string, string> v)
        {
            return Regex.Replace(tpl ?? "", @"\{([^{}]+)\}", m =>
            {
                string val; return v.TryGetValue(m.Groups[1].Value.Trim(), out val) ? val : m.Value;
            });
        }

        void Run()
        {
            string outDir = folder.Text.Trim();
            if (outDir == "") return;
            if (email.Checked && sendNow.Checked &&
                MessageBox.Show(this, L("Gửi email ngay cho tất cả người nhận trong danh sách?", "Send the emails to all recipients now?"),
                                "PDF Toolkit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            Directory.CreateDirectory(outDir);
            running = true; cancel = false; run.Enabled = false; close.Text = L("Dừng", "Stop");
            dynamic mm = doc.MailMerge, ds = mm.DataSource;
            int total = 0;
            try { total = (int)ds.RecordCount; } catch { }
            bar.Maximum = Math.Max(1, total);
            int done = 0, mailed = 0;
            var errors = new List<string>();
            Prefs prefs = Prefs.Load();
            try
            {
                app.ScreenUpdating = false;
                ds.ActiveRecord = -4;                          // wdFirstRecord
                int last = -999;
                while (!cancel)
                {
                    int rec = (int)ds.ActiveRecord;
                    if (rec == last) break;
                    last = rec;
                    var v = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (string f in fields)
                    {
                        try { v[f] = Convert.ToString(ds.DataFields.Item(f).Value).Trim(); } catch { v[f] = ""; }
                    }
                    string name = Fill(nameTpl.Text, v);
                    foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
                    if (name.Trim() == "") name = "record_" + rec;
                    string pdf = Connect.Unique(Path.Combine(outDir, name.Trim() + ".pdf"));
                    status.Text = L("Đang tạo: ", "Creating: ") + Path.GetFileName(pdf);
                    Application.DoEvents();
                    try
                    {
                        ds.FirstRecord = rec; ds.LastRecord = rec;
                        mm.Destination = 0;                    // wdSendToNewDocument
                        mm.Execute(false);
                        dynamic res = app.ActiveDocument;
                        try { owner.ExportWord(res, pdf, prefs, true); }
                        finally { res.Close(0); }
                        done++;
                        if (email.Checked && toField.SelectedItem != null)
                        {
                            string to = v[(string)toField.SelectedItem];
                            if (to.Contains("@"))
                            {
                                Mail.Compose(to, Fill(subject.Text, v), Fill(body.Text, v), new[] { pdf }, sendNow.Checked);
                                mailed++;
                            }
                            else errors.Add(name + ": " + L("không có địa chỉ email", "no email address"));
                        }
                    }
                    catch (Exception ex) { errors.Add(name + ": " + Connect.Inner(ex).Message); }
                    bar.Value = Math.Min(bar.Maximum, done + errors.Count);
                    Application.DoEvents();
                    ds.ActiveRecord = -2;                      // wdNextRecord
                }
            }
            finally
            {
                try { ds.FirstRecord = 1; ds.LastRecord = -16; } catch { }   // wdDefaultFirstRecord / wdDefaultLastRecord
                try { app.ScreenUpdating = true; } catch { }
                running = false; run.Enabled = true; close.Text = L("Đóng", "Close");
            }
            string msg = string.Format(L("Đã tạo {0} file PDF trong:\n{1}", "Created {0} PDF file(s) in:\n{1}"), done, outDir);
            if (email.Checked) msg += "\n" + string.Format(sendNow.Checked ? L("Đã gửi {0} email.", "Sent {0} email(s).")
                                                                           : L("Đã mở {0} thư nháp trong Outlook.", "Opened {0} draft(s) in Outlook."), mailed);
            if (cancel) msg += "\n" + L("(Đã dừng giữa chừng)", "(Stopped early)");
            if (errors.Count > 0) msg += "\n\n" + L("Lỗi:", "Errors:") + "\n" + string.Join("\n", errors.GetRange(0, Math.Min(10, errors.Count)));
            status.Text = "";
            lastDone = done;
            if (quiet) return;
            MessageBox.Show(this, msg, "PDF Toolkit", MessageBoxButtons.OK, errors.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            if (done > 0) System.Diagnostics.Process.Start("explorer.exe", "\"" + outDir + "\"");
        }
    }
}

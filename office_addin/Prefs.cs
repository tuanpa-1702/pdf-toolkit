// PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
// SPDX-License-Identifier: AGPL-3.0-only
// Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
// Tuỳ chọn tạo PDF của add-in (lưu ở HKCU\Software\TuanPA\PDFToolkit\Addin) + hộp thoại chỉnh.
using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PDFToolkitAddin
{
    internal class Prefs
    {
        const string Key = @"Software\TuanPA\PDFToolkit\Addin";
        public int Folder;          // 0 = cạnh file gốc, 1 = thư mục chọn sẵn, 2 = hỏi mỗi lần
        public string FolderPath = "";
        public int After;           // 0 = mở bằng PDF Toolkit, 1 = mở bằng trình xem PDF mặc định, 2 = mở thư mục, 3 = không làm gì
        public bool Bookmarks = true, Tags = true, Props = true, PdfA, Markup;
        public bool Screen;         // tối ưu cho màn hình (file nhỏ) thay vì in
        public int Scope;           // Word: 0 cả tài liệu / 1 trang hiện tại / 2 vùng chọn; Excel: 0 cả workbook / 1 sheet / 2 vùng chọn;
                                    // PowerPoint: 0 tất cả slide / 1 slide hiện tại

        public static Prefs Load()
        {
            var p = new Prefs();
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(Key))
                {
                    if (k == null) return p;
                    p.Folder = I(k, "Folder", 0); p.FolderPath = Convert.ToString(k.GetValue("FolderPath", ""));
                    p.After = I(k, "After", 0); p.Bookmarks = I(k, "Bookmarks", 1) == 1; p.Tags = I(k, "Tags", 1) == 1;
                    p.Props = I(k, "Props", 1) == 1; p.PdfA = I(k, "PdfA", 0) == 1; p.Markup = I(k, "Markup", 0) == 1;
                    p.Screen = I(k, "Screen", 0) == 1; p.Scope = I(k, "Scope", 0);
                }
            }
            catch { }
            return p;
        }

        static int I(RegistryKey k, string n, int d) { try { return Convert.ToInt32(k.GetValue(n, d)); } catch { return d; } }

        public void Save()
        {
            using (var k = Registry.CurrentUser.CreateSubKey(Key))
            {
                k.SetValue("Folder", Folder); k.SetValue("FolderPath", FolderPath ?? ""); k.SetValue("After", After);
                k.SetValue("Bookmarks", Bookmarks ? 1 : 0); k.SetValue("Tags", Tags ? 1 : 0); k.SetValue("Props", Props ? 1 : 0);
                k.SetValue("PdfA", PdfA ? 1 : 0); k.SetValue("Markup", Markup ? 1 : 0); k.SetValue("Screen", Screen ? 1 : 0);
                k.SetValue("Scope", Scope);
            }
        }
    }

    internal class PrefsDialog : Form
    {
        readonly Prefs p;
        readonly RadioButton[] folder = new RadioButton[3];
        readonly TextBox folderPath = new TextBox();
        readonly ComboBox after = new ComboBox(), scope = new ComboBox();
        readonly CheckBox bookmarks = new CheckBox(), tags = new CheckBox(), props = new CheckBox(), pdfa = new CheckBox(),
                          markup = new CheckBox(), screen = new CheckBox();

        public PrefsDialog(Prefs prefs, string host)
        {
            p = prefs;
            bool en = Connect.English();
            Func<string, string, string> L = (vi, e) => en ? e : vi;
            Text = L("Tuỳ chọn tạo PDF - PDF Toolkit", "PDF creation preferences - PDF Toolkit");
            Font = new Font("Segoe UI", 9.5f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(560, 470);

            var tbl = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(14), ColumnCount = 1, AutoScroll = true };
            Controls.Add(tbl);

            var gFolder = Group(L("Nơi lưu file PDF", "Where to save the PDF"));
            var fl = Flow(gFolder);
            folder[0] = Radio(fl, L("Cùng thư mục với file gốc (chưa lưu thì vào Documents)", "Next to the original file (Documents if unsaved)"));
            folder[1] = Radio(fl, L("Luôn lưu vào thư mục:", "Always save to this folder:"));
            var row = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(18, 0, 0, 4) };
            folderPath.Width = 380; folderPath.Text = p.FolderPath;
            var browse = new Button { Text = L("Chọn...", "Browse..."), AutoSize = true };
            browse.Click += (s, e) =>
            {
                using (var d = new FolderBrowserDialog { SelectedPath = folderPath.Text })
                    if (d.ShowDialog(this) == DialogResult.OK) { folderPath.Text = d.SelectedPath; folder[1].Checked = true; }
            };
            row.Controls.Add(folderPath); row.Controls.Add(browse); fl.Controls.Add(row);
            folder[2] = Radio(fl, L("Hỏi tên và nơi lưu mỗi lần", "Ask for a file name each time"));
            folder[Math.Max(0, Math.Min(2, p.Folder))].Checked = true;
            tbl.Controls.Add(gFolder);

            var gOut = Group(L("Nội dung PDF", "PDF content"));
            var ol = Flow(gOut);
            var sr = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            sr.Controls.Add(new Label { Text = L("Phạm vi:", "Range:"), AutoSize = true, Margin = new Padding(0, 6, 6, 0) });
            scope.DropDownStyle = ComboBoxStyle.DropDownList; scope.Width = 300;
            if (host == "word") scope.Items.AddRange(new object[] { L("Cả tài liệu", "Entire document"), L("Trang hiện tại", "Current page"), L("Vùng đang chọn", "Selection") });
            else if (host == "excel") scope.Items.AddRange(new object[] { L("Cả workbook", "Entire workbook"), L("Sheet đang mở", "Active sheet"), L("Vùng đang chọn", "Selection") });
            else scope.Items.AddRange(new object[] { L("Tất cả slide", "All slides"), L("Slide hiện tại", "Current slide") });
            scope.SelectedIndex = Math.Min(p.Scope, scope.Items.Count - 1);
            sr.Controls.Add(scope); ol.Controls.Add(sr);
            if (host == "word")
            {
                Check(ol, bookmarks, L("Tạo bookmark từ các Heading", "Create bookmarks from headings"), p.Bookmarks);
                Check(ol, markup, L("Kèm nhận xét và theo dõi thay đổi (Track Changes)", "Include comments and tracked changes"), p.Markup);
            }
            Check(ol, tags, L("Gắn thẻ cấu trúc (hỗ trợ trình đọc màn hình, PDF → Word chính xác hơn)", "Structure tags for accessibility"), p.Tags);
            Check(ol, props, L("Kèm thuộc tính tài liệu (tiêu đề, tác giả...)", "Include document properties"), p.Props);
            Check(ol, pdfa, L("PDF/A - chuẩn lưu trữ lâu dài", "PDF/A - archival standard"), p.PdfA);
            Check(ol, screen, L("Tối ưu dung lượng (xem trên màn hình, gửi email)", "Minimum size (on-screen viewing, email)"), p.Screen);
            if (host == "powerpoint") pdfa.Visible = props.Visible = false;
            tbl.Controls.Add(gOut);

            var gAfter = Group(L("Sau khi tạo PDF (nút Tạo PDF)", "After creating the PDF (Create PDF button)"));
            var al = Flow(gAfter);
            after.DropDownStyle = ComboBoxStyle.DropDownList; after.Width = 420;
            after.Items.AddRange(new object[] { L("Mở bằng PDF Toolkit", "Open in PDF Toolkit"), L("Mở bằng trình xem PDF mặc định", "Open in the default PDF viewer"),
                                                L("Mở thư mục chứa file", "Open the containing folder"), L("Không làm gì", "Do nothing") });
            after.SelectedIndex = Math.Max(0, Math.Min(3, p.After));
            al.Controls.Add(after);
            tbl.Controls.Add(gAfter);

            var btns = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, Height = 44, Padding = new Padding(10, 6, 10, 6) };
            var cancel = new Button { Text = L("Huỷ", "Cancel"), DialogResult = DialogResult.Cancel, AutoSize = true };
            var ok = new Button { Text = "OK", AutoSize = true, MinimumSize = new Size(90, 0) };
            ok.Click += (s, e) => { Apply(); DialogResult = DialogResult.OK; };
            btns.Controls.Add(cancel); btns.Controls.Add(ok);
            Controls.Add(btns);
            AcceptButton = ok; CancelButton = cancel;
        }

        void Apply()
        {
            p.Folder = folder[1].Checked ? 1 : folder[2].Checked ? 2 : 0;
            p.FolderPath = folderPath.Text.Trim();
            p.After = after.SelectedIndex; p.Scope = Math.Max(0, scope.SelectedIndex);
            p.Bookmarks = bookmarks.Checked; p.Markup = markup.Checked; p.Tags = tags.Checked; p.Props = props.Checked;
            p.PdfA = pdfa.Checked; p.Screen = screen.Checked;
            p.Save();
        }

        static GroupBox Group(string t) { return new GroupBox { Text = t, AutoSize = true, Dock = DockStyle.Top, Padding = new Padding(8), Margin = new Padding(0, 0, 0, 10) }; }
        static FlowLayoutPanel Flow(GroupBox g)
        {
            var f = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Dock = DockStyle.Fill, WrapContents = false };
            g.Controls.Add(f); return f;
        }
        static RadioButton Radio(FlowLayoutPanel f, string t) { var r = new RadioButton { Text = t, AutoSize = true }; f.Controls.Add(r); return r; }
        static void Check(FlowLayoutPanel f, CheckBox c, string t, bool v) { c.Text = t; c.AutoSize = true; c.Checked = v; f.Controls.Add(c); }
    }
}

// PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
// SPDX-License-Identifier: AGPL-3.0-only
// Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
// Nhập nhận xét từ file PDF (đã được người khác ghi chú / tô sáng / gạch xoá) vào tài liệu Word dưới dạng Comment.
// Đọc PDF bằng "PDF Toolkit.exe --export-comments <pdf> <json>", rồi tìm đoạn chữ tương ứng trong Word để gắn Comment.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace PDFToolkitAddin
{
    internal static class Comments
    {
        static string L(string vi, string en) { return Connect.L(vi, en); }

        /// quiet=true: không hỏi / không hiện thông báo (kiểm thử tự động). Trả về câu tóm tắt kết quả.
        public static string Import(Connect owner, dynamic app, string pdf, bool quiet)
        {
            if (app.Documents.Count == 0) return null;
            dynamic doc = app.ActiveDocument;
            if (pdf == null)
            {
                string dir = "", stem = "";
                try { dir = Path.GetDirectoryName((string)doc.FullName); stem = Path.GetFileNameWithoutExtension((string)doc.Name); } catch { }
                using (var d = new OpenFileDialog { Filter = "PDF (*.pdf)|*.pdf", Title = L("Chọn file PDF đã được nhận xét", "Choose the reviewed PDF"),
                                                    InitialDirectory = Directory.Exists(dir) ? dir : "", FileName = stem + ".pdf" })
                {
                    if (d.ShowDialog() != DialogResult.OK) return null;
                    pdf = d.FileName;
                }
            }
            string json = Path.Combine(Path.GetTempPath(), "pdftk_comments_" + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                try { app.StatusBar = L("PDF Toolkit: đang đọc nhận xét trong PDF...", "PDF Toolkit: reading PDF comments..."); } catch { }
                var psi = new ProcessStartInfo(Connect.ExePath(), "--export-comments \"" + pdf + "\" \"" + json + "\"")
                          { UseShellExecute = false, CreateNoWindow = true };
                using (var p = Process.Start(psi))
                    if (!p.WaitForExit(180000)) { try { p.Kill(); } catch { } throw new TimeoutException("PDF Toolkit --export-comments"); }
                if (!File.Exists(json)) throw new IOException(L("PDF Toolkit không trả kết quả.", "PDF Toolkit returned no result."));
                var root = (Dictionary<string, object>)new JavaScriptSerializer { MaxJsonLength = int.MaxValue }
                           .DeserializeObject(File.ReadAllText(json, System.Text.Encoding.UTF8));
                if (!(bool)root["ok"]) throw new IOException(Convert.ToString(root["error"]));
                var items = (object[])root["comments"];
                if (items.Length == 0)
                {
                    string none = L("File PDF này không có nhận xét nào.", "This PDF has no comments.");
                    if (!quiet) MessageBox.Show(none, "PDF Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return none;
                }
                if (!quiet && MessageBox.Show(string.Format(L("Tìm thấy {0} nhận xét trong PDF. Nhập vào tài liệu Word này?",
                                                              "Found {0} comment(s) in the PDF. Import them into this Word document?"), items.Length),
                                              "PDF Toolkit", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    return null;

                int placed = 0, loose = 0;
                bool track = false;
                try { track = (bool)doc.TrackRevisions; doc.TrackRevisions = false; } catch { }
                try
                {
                    foreach (Dictionary<string, object> c in items)
                    {
                        string kind = S(c, "kind"), author = S(c, "author"), content = S(c, "content"), anchor = S(c, "anchor");
                        int page = Convert.ToInt32(c.ContainsKey("page") ? c["page"] : 0);
                        string prefix = kind == "strikeout" ? L("[Đề nghị xoá] ", "[Delete] ") :
                                        kind == "insert" ? L("[Đề nghị chèn] ", "[Insert] ") :
                                        kind == "highlight" && content == "" ? L("[Tô sáng]", "[Highlight]") : "";
                        string text = prefix + content;
                        var reps = c.ContainsKey("replies") ? (object[])c["replies"] : new object[0];
                        foreach (Dictionary<string, object> r in reps)
                            text += "\n↳ " + (S(r, "author") != "" ? S(r, "author") + ": " : "") + S(r, "content");
                        dynamic range = Find(doc, anchor);
                        if (range == null)
                        {
                            // không tìm thấy đoạn chữ -> gắn vào đầu tài liệu, ghi rõ trang và đoạn trích
                            range = doc.Range(0, 0);
                            text = string.Format(L("[Trang {0}] ", "[Page {0}] "), page) +
                                   (anchor != "" ? "\"" + Cut(anchor, 80) + "\" - " : "") + text;
                            loose++;
                        }
                        else placed++;
                        dynamic cm = doc.Comments.Add(range, text.Trim());
                        if (author != "")
                        {
                            try { cm.Author = author; cm.Initial = Initials(author); } catch { }
                        }
                    }
                }
                finally { try { doc.TrackRevisions = track; } catch { } }
                string msg = string.Format(L("Đã nhập {0} nhận xét: {1} gắn đúng đoạn chữ, {2} gắn ở đầu tài liệu (không tìm thấy đoạn chữ tương ứng).",
                                             "Imported {0} comment(s): {1} attached to the matching text, {2} placed at the top (text not found)."),
                                           placed + loose, placed, loose);
                if (!quiet) MessageBox.Show(msg, "PDF Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return msg;
            }
            finally
            {
                Cursor.Current = Cursors.Default;
                try { app.StatusBar = ""; } catch { }
                try { File.Delete(json); } catch { }
            }
        }

        static string S(Dictionary<string, object> d, string k)
        {
            object v;
            return d.TryGetValue(k, out v) && v != null ? Convert.ToString(v).Trim() : "";
        }

        static string Cut(string s, int n) { return s.Length <= n ? s : s.Substring(0, n) + "..."; }

        static string Initials(string name)
        {
            string r = "";
            foreach (string w in name.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) r += char.ToUpper(w[0]);
            return r.Length > 4 ? r.Substring(r.Length - 4) : r;
        }

        /// Tìm đoạn chữ trong Word: thử nguyên văn, rồi 8 từ đầu, rồi 4 từ đầu (chữ trong PDF hay bị ngắt dòng / gạch nối khác).
        static dynamic Find(dynamic doc, string anchor)
        {
            anchor = (anchor ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            if (anchor.Length < 2) return null;
            string[] words = anchor.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var tries = new List<string> { Cut2(anchor, 200) };
            if (words.Length > 8) tries.Add(string.Join(" ", words, 0, 8));
            if (words.Length > 4) tries.Add(string.Join(" ", words, 0, 4));
            foreach (string t in tries)
            {
                dynamic r = doc.Content;
                dynamic f = r.Find;
                f.ClearFormatting();
                bool ok = f.Execute(FindText: t.Replace("^", "^^"), MatchCase: false, MatchWholeWord: false,
                                    MatchWildcards: false, Forward: true, Wrap: 0);
                if (ok) return r;
            }
            return null;
        }

        static string Cut2(string s, int n)
        {
            if (s.Length <= n) return s;
            int i = s.LastIndexOf(' ', n);
            return s.Substring(0, i > 20 ? i : n);
        }
    }
}

// PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
// SPDX-License-Identifier: AGPL-3.0-only
// Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
// Soạn email kèm file PDF qua Outlook (bản cổ điển, có COM). Không có Outlook -> mở thư mục chứa file.
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace PDFToolkitAddin
{
    internal static class Mail
    {
        static dynamic outlook;

        public static bool Available()
        {
            return Type.GetTypeFromProgID("Outlook.Application") != null;
        }

        static dynamic Outlook()
        {
            if (outlook != null)
            {
                try { var _ = outlook.Name; return outlook; } catch { outlook = null; }
            }
            Type t = Type.GetTypeFromProgID("Outlook.Application");
            if (t == null) return null;
            outlook = Activator.CreateInstance(t);           // Outlook chỉ chạy 1 phiên -> gắn vào phiên đang mở
            return outlook;
        }

        /// Tạo thư mới kèm file. send=false -> mở cửa sổ soạn thư để người dùng xem / sửa rồi tự bấm Gửi.
        public static bool Compose(string to, string subject, string body, string[] files, bool send)
        {
            dynamic ol = Outlook();
            if (ol == null)
            {
                Fallback(files);
                return false;
            }
            dynamic m = ol.CreateItem(0);                    // 0 = olMailItem
            if (!string.IsNullOrEmpty(to)) m.To = to;
            m.Subject = subject ?? "";
            foreach (string f in files) m.Attachments.Add(f);
            if (send)
            {
                m.Body = body ?? "";
                m.Send();
            }
            else
            {
                m.Display(false);                            // hiện trước để Outlook chèn chữ ký mặc định
                if (!string.IsNullOrEmpty(body))
                {
                    string html = System.Net.WebUtility.HtmlEncode(body).Replace("\r\n", "<br>").Replace("\n", "<br>");
                    string sig = Convert.ToString(m.HTMLBody);
                    int i = sig.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                    i = i < 0 ? -1 : sig.IndexOf('>', i);
                    m.HTMLBody = i < 0 ? html + sig : sig.Substring(0, i + 1) + "<div style='font-family:Calibri,Arial;font-size:11pt'>"
                                                       + html + "</div><br>" + sig.Substring(i + 1);
                }
            }
            return true;
        }

        static void Fallback(string[] files)
        {
            bool en = Connect.English();
            MessageBox.Show(en ? "Classic Outlook was not found. The PDF folder will be opened so you can attach the file manually."
                               : "Không tìm thấy Outlook (bản cổ điển). Thư mục chứa PDF sẽ được mở để bạn đính kèm thủ công.",
                            "PDF Toolkit", MessageBoxButtons.OK, MessageBoxIcon.Information);
            if (files.Length > 0 && File.Exists(files[0]))
                Process.Start("explorer.exe", "/select,\"" + files[0] + "\"");
        }
    }
}

// PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
// SPDX-License-Identifier: AGPL-3.0-only
// Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
// Icon ribbon vẽ bằng GDI+: trang tài liệu có dải đỏ "PDF" + biểu tượng phụ (sao, bánh răng, phong bì...).
using System;
using System.Drawing;
using System.Drawing.Drawing2D;

namespace PDFToolkitAddin
{
    internal static class Icons
    {
        static readonly Color Red = Color.FromArgb(220, 38, 38);
        static readonly Color Ink = Color.FromArgb(71, 85, 105);
        static readonly Color Gold = Color.FromArgb(245, 158, 11);
        static readonly Color Blue = Color.FromArgb(37, 99, 235);
        static readonly Color Green = Color.FromArgb(22, 163, 74);

        public static Bitmap Get(string id)
        {
            var bmp = new Bitmap(64, 64);                 // vẽ 64px, Office tự thu nhỏ (nét hơn trên màn hình HiDPI)
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                g.ScaleTransform(2f, 2f);                 // toạ độ dưới đây theo lưới 32x32
                Page(g);
                switch (id)
                {
                    case "btnCreate": Star(g, 22, 22, 8, Gold); break;
                    case "btnPrefs": Prefs(g); break;
                    case "btnAttach": Envelope(g, 15, 18, 15, 11); Star(g, 26, 16, 5, Gold); break;
                    case "btnMailMerge": Envelope(g, 14, 20, 13, 9); Envelope(g, 17, 17, 13, 9); Arrow(g); break;
                    case "btnReview": Bubble(g, 17, 6, 13, 10, Gold); Star(g, 21, 24, 6, Gold); break;
                    case "mnuComments": Bubble(g, 15, 15, 15, 12, Gold); break;
                    case "btnImportComments": Bubble(g, 15, 15, 15, 12, Blue); break;
                    case "btnExportMarkup": Bubble(g, 15, 15, 15, 12, Green); break;
                    case "olMailToPdf": case "olCtxMail": case "olCtxMails": Envelope(g, 14, 17, 16, 12); Arrow(g); break;
                    case "olAttToPdf": case "olCtxAttSave": Clip(g, 22, 13); Arrow(g); break;
                    case "olAttachAsPdf": Clip(g, 22, 13); Star(g, 25, 26, 5, Gold); break;
                    case "olCompressAtt": Squeeze(g); break;
                    case "olProtectAtt": Lock(g, 17, 17); break;
                }
            }
            return bmp;
        }

        static void Page(Graphics g)
        {
            var pts = new[] { new PointF(6, 2), new PointF(20, 2), new PointF(26, 8), new PointF(26, 30), new PointF(6, 30) };
            using (var b = new SolidBrush(Color.White)) g.FillPolygon(b, pts);
            using (var p = new Pen(Color.FromArgb(148, 163, 184), 1f)) g.DrawPolygon(p, pts);
            using (var p = new Pen(Color.FromArgb(148, 163, 184), 1f)) g.DrawLines(p, new[] { new PointF(20, 2), new PointF(20, 8), new PointF(26, 8) });
            using (var b = new SolidBrush(Red)) g.FillRectangle(b, 2, 5, 15, 6);      // dải đỏ kiểu Acrobat
            using (var b = new SolidBrush(Color.FromArgb(120, 0, 0))) g.FillRectangle(b, 2, 10, 15, 1);
        }

        static void Star(Graphics g, float cx, float cy, float r, Color c)
        {
            var pts = new PointF[16];
            for (int i = 0; i < 16; i++)
            {
                double a = Math.PI * 2 * i / 16 - Math.PI / 2;
                float rr = i % 2 == 0 ? r : r * 0.55f;
                pts[i] = new PointF(cx + (float)(Math.Cos(a) * rr), cy + (float)(Math.Sin(a) * rr));
            }
            using (var b = new SolidBrush(c)) g.FillPolygon(b, pts);
            using (var p = new Pen(Color.FromArgb(180, 83, 9), 0.8f)) g.DrawPolygon(p, pts);
        }

        static void Prefs(Graphics g)
        {
            using (var b = new SolidBrush(Color.FromArgb(241, 245, 249))) g.FillRectangle(b, 16, 17, 14, 13);
            using (var p = new Pen(Ink, 1.2f))
            {
                g.DrawRectangle(p, 16, 17, 14, 13);
                for (int i = 0; i < 3; i++)
                {
                    float y = 20.5f + i * 3.5f;
                    g.DrawEllipse(p, 18, y - 1, 2, 2);
                    g.DrawLine(p, 22, y, 28, y);
                }
            }
        }

        static void Envelope(Graphics g, float x, float y, float w, float h)
        {
            using (var b = new SolidBrush(Color.FromArgb(248, 250, 252))) g.FillRectangle(b, x, y, w, h);
            using (var p = new Pen(Ink, 1.1f))
            {
                g.DrawRectangle(p, x, y, w, h);
                g.DrawLines(p, new[] { new PointF(x, y), new PointF(x + w / 2, y + h * 0.6f), new PointF(x + w, y) });
            }
        }

        static void Arrow(Graphics g)
        {
            var pts = new[] { new PointF(24, 1), new PointF(29, 1), new PointF(29, 8), new PointF(31.5f, 8),
                              new PointF(26.5f, 14), new PointF(21.5f, 8), new PointF(24, 8) };
            using (var b = new SolidBrush(Green)) g.FillPolygon(b, pts);
        }

        static void Clip(Graphics g, float x, float y)            // kẹp giấy (file đính kèm)
        {
            using (var p = new Pen(Ink, 1.6f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            using (var path = new GraphicsPath())
            {
                path.AddLine(x + 6, y + 4, x + 6, y + 13);
                path.AddArc(x, y + 10, 6, 6, 0, 180);
                path.AddLine(x, y + 13, x, y + 2);
                path.AddArc(x, y - 2, 4, 4, 180, 180);
                path.AddLine(x + 4, y, x + 4, y + 11);
                g.DrawPath(p, path);
            }
        }

        static void Squeeze(Graphics g)                           // hai mũi tên ép vào giữa (nén)
        {
            using (var b = new SolidBrush(Blue))
            {
                g.FillPolygon(b, new[] { new PointF(16, 13), new PointF(30, 13), new PointF(23, 20) });
                g.FillPolygon(b, new[] { new PointF(16, 30), new PointF(30, 30), new PointF(23, 23) });
            }
            using (var p = new Pen(Ink, 1.2f)) g.DrawLine(p, 15, 21.5f, 31, 21.5f);
        }

        static void Lock(Graphics g, float x, float y)
        {
            using (var p = new Pen(Ink, 1.8f)) g.DrawArc(p, x + 2.5f, y, 9, 10, 180, 180);
            using (var b = new SolidBrush(Gold)) g.FillRectangle(b, x, y + 5, 14, 10);
            using (var p = new Pen(Color.FromArgb(180, 83, 9), 0.9f)) g.DrawRectangle(p, x, y + 5, 14, 10);
            using (var b = new SolidBrush(Color.FromArgb(120, 53, 15))) g.FillEllipse(b, x + 5.5f, y + 8, 3, 3);
        }

        static void Bubble(Graphics g, float x, float y, float w, float h, Color c)
        {
            using (var path = new GraphicsPath())
            {
                path.AddArc(x, y, 6, 6, 180, 90);
                path.AddArc(x + w - 6, y, 6, 6, 270, 90);
                path.AddArc(x + w - 6, y + h - 6, 6, 6, 0, 90);
                path.AddLine(x + w * 0.45f, y + h, x + w * 0.25f, y + h + 4);
                path.AddLine(x + w * 0.25f, y + h + 4, x + w * 0.25f, y + h);
                path.AddArc(x, y + h - 6, 6, 6, 90, 90);
                path.CloseFigure();
                using (var b = new SolidBrush(Color.FromArgb(254, 243, 199))) g.FillPath(b, path);
                using (var p = new Pen(c, 1.3f)) g.DrawPath(p, path);
            }
            using (var p = new Pen(c, 1.1f))
                for (int i = 0; i < 2; i++)
                    g.DrawLine(p, x + 3, y + 3.5f + i * 3, x + w - 3, y + 3.5f + i * 3);
        }
    }
}

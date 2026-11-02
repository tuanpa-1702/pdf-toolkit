# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Bộ công cụ PDF xử lý theo lô (trang chủ dạng lưới công cụ của PDF Toolkit).

Mỗi công cụ khai báo: nhóm, tên, mô tả, loại file đầu vào, các tuỳ chọn (Opt) và hàm chạy run(ctx).
Phần giao diện (tools_ui.py) tự dựng form tuỳ chọn từ danh sách Opt, rồi chạy run() trong luồng nền.
Toàn bộ chạy offline: PyMuPDF, Pillow, Microsoft Office (chuyển Word/Excel/PowerPoint), pdf2docx, openpyxl.
"""
import html as _html
import io
import json
import os
import re
import secrets
import zipfile
from dataclasses import dataclass, field
from datetime import datetime

import pymupdf

import pdf_core

PDF_EXT = (".pdf",)
IMAGE_EXT = (".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp", ".psd", ".jxr", ".pnm")
EBOOK_EXT = (".epub", ".mobi", ".fb2", ".cbz", ".xps", ".oxps", ".svg")
WORD_EXT = (".doc", ".docx", ".docm", ".rtf", ".odt", ".wpd", ".wps", ".dot", ".dotx")
EXCEL_EXT = (".xls", ".xlsx", ".xlsm", ".xlsb", ".ods", ".csv")
PPT_EXT = (".ppt", ".pptx", ".pptm", ".odp", ".pps", ".ppsx")
TEXT_EXT = (".txt", ".md", ".markdown", ".html", ".htm", ".json", ".xml", ".log", ".csv")
MAIL_EXT = (".eml",)
OPENABLE_EXT = PDF_EXT + IMAGE_EXT + EBOOK_EXT     # mở thẳng được bằng MuPDF (ảnh / ebook -> PDF)

PAPERS = {"A4": "a4", "A3": "a3", "A5": "a5", "Letter": "letter", "Legal": "legal"}
MM = 72 / 25.4


class ToolError(Exception):
    pass


class Cancelled(Exception):
    pass


# ---------------------------------------------------------------------------
#  Khai báo công cụ
# ---------------------------------------------------------------------------
@dataclass
class Opt:
    key: str
    label: str
    kind: str = "text"          # text|textarea|int|float|bool|choice|color|file|files|pages|password|form|info
    default: object = None
    choices: list = None        # [(giá trị, nhãn)]
    min: float = 0
    max: float = 100000
    step: float = 1
    suffix: str = ""
    tip: str = ""
    placeholder: str = ""
    show_if: tuple = None       # (key, (giá trị...)) - chỉ hiện khi tuỳ chọn khác có giá trị này
    prefill: object = None      # hàm(path) -> giá trị, nạp từ file đầu tiên trong danh sách
    filter: str = ""            # bộ lọc cho kind=file/files


@dataclass
class Tool:
    id: str
    group: str
    name: str
    desc: str
    icon: int
    run: object
    opts: list = field(default_factory=list)
    accept: tuple = PDF_EXT
    output: str = "each"        # each: mỗi file 1 kết quả | single: gộp 1 kết quả | none: chỉ báo cáo
    min_files: int = 1
    max_files: int = 0          # 0 = không giới hạn
    popular: bool = False
    keywords: str = ""
    note: str = ""


GROUPS = [
    ("organize", "Sắp xếp & Quản lý trang"),
    ("stamp", "Đóng dấu & Đánh số"),
    ("to_pdf", "Chuyển đổi sang PDF"),
    ("from_pdf", "Chuyển đổi từ PDF"),
    ("secure", "Bảo mật, Form & Tối ưu"),
]

TOOLS = []


def tool(**kw):
    def deco(fn):
        TOOLS.append(Tool(run=fn, **kw))
        return fn
    return deco


def get_tool(tid):
    return next(t for t in TOOLS if t.id == tid)


# ---------------------------------------------------------------------------
#  Ngữ cảnh chạy
# ---------------------------------------------------------------------------
class Ctx:
    def __init__(self, files, opts, out_dir=None, progress=None, log=None, cancel_flag=None):
        self.files = list(files)
        self.opts = dict(opts)
        self.out_dir = out_dir
        self._progress = progress or (lambda f, m: None)
        self._log = log or print
        self._cancel = cancel_flag or (lambda: False)
        self.outputs = []

    def __getitem__(self, k):
        return self.opts.get(k)

    def progress(self, frac, msg=""):
        if self._cancel():
            raise Cancelled()
        self._progress(max(0.0, min(1.0, frac)), msg)

    def log(self, msg):
        self._log(msg)

    def out(self, src, suffix, ext=".pdf", name=None, folder=False):
        """Đường dẫn kết quả không trùng file có sẵn (cùng thư mục file gốc nếu không chọn thư mục)."""
        d = self.out_dir or os.path.dirname(os.path.abspath(src))
        stem = name or (os.path.splitext(os.path.basename(src))[0] + suffix)
        stem = re.sub(r'[<>:"/\\|?*]', "_", stem).strip() or "ket_qua"
        p = os.path.join(d, stem + ("" if folder else ext))
        k = 2
        while os.path.exists(p):
            p = os.path.join(d, f"{stem} ({k})" + ("" if folder else ext))
            k += 1
        if folder:
            os.makedirs(p)
        self.outputs.append(p)
        return p

    def each(self, fn):
        """Chạy fn(path, i) cho từng file; lỗi ở một file không dừng cả lô."""
        n = len(self.files)
        errors = 0
        for i, f in enumerate(self.files):
            self.progress(i / n, os.path.basename(f))
            try:
                fn(f, i)
            except (Cancelled, KeyboardInterrupt):
                raise
            except Exception as e:
                errors += 1
                self.log(f"✘ {os.path.basename(f)}: {e}")
        self.progress(1, "Xong")
        if errors and errors == n:
            raise ToolError("Không xử lý được file nào.")


# ---------------------------------------------------------------------------
#  Tiện ích chung
# ---------------------------------------------------------------------------
def open_doc(path, password=None):
    ext = os.path.splitext(path)[1].lower()
    if ext in (".psd", ".jxr") or ext not in OPENABLE_EXT and ext not in PDF_EXT:
        return image_doc_pillow(path)
    doc = pymupdf.open(path)
    if doc.needs_pass:
        if not (password and doc.authenticate(password)):
            doc.close()
            raise ToolError("file có mật khẩu - dùng công cụ 'Mở khoá PDF' trước (hoặc nhập mật khẩu)")
    if not doc.is_pdf:
        if doc.is_reflowable:
            doc.layout(pymupdf.paper_rect("a4"), fontsize=11)
        pdf = pymupdf.open("pdf", doc.convert_to_pdf())
        doc.close()
        return pdf
    return doc


def image_doc_pillow(path):
    """Ảnh MuPDF không đọc được (PSD...) -> PDF 1 trang qua Pillow."""
    from PIL import Image
    im = Image.open(path)
    im.load()
    if im.mode not in ("RGB", "RGBA", "L"):
        im = im.convert("RGBA" if "A" in im.getbands() else "RGB")
    buf = io.BytesIO()
    im.save(buf, "PNG")
    doc = pymupdf.open()
    w, h = im.size
    dpi = im.info.get("dpi", (96, 96))[0] or 96
    page = doc.new_page(width=w * 72 / dpi, height=h * 72 / dpi)
    page.insert_image(page.rect, stream=buf.getvalue())
    return doc


def save_doc(doc, path, fonts=False, **kw):
    if fonts:
        try:
            doc.subset_fonts()
        except Exception:
            pass
    doc.save(path, garbage=kw.pop("garbage", 3), deflate=True, **kw)


def normalize_rotation(doc, pages=None):
    """Bỏ thuộc tính xoay (nội dung giữ nguyên hướng hiển thị) để toạ độ hiển thị = toạ độ trang."""
    for i in (pages if pages is not None else range(doc.page_count)):
        p = doc[i]
        if p.rotation:
            p.remove_rotation()


def parse_pages(spec, n, default_all=True):
    """'1-3, 5, 8-' -> [0,1,2,4,7..n-1]. Hỗ trợ 'le' (lẻ), 'chan' (chẵn), 'cuoi' (trang cuối)."""
    spec = (spec or "").strip().lower()
    if not spec:
        return list(range(n)) if default_all else []
    out = []
    for part in re.split(r"[,;]\s*|\s+", spec):
        if not part:
            continue
        if part in ("le", "lẻ", "odd"):
            out += list(range(0, n, 2))
            continue
        if part in ("chan", "chẵn", "even"):
            out += list(range(1, n, 2))
            continue
        part = part.replace("cuối", str(n)).replace("cuoi", str(n)).replace("end", str(n))
        m = re.fullmatch(r"(\d*)\s*-\s*(\d*)", part)
        if m:
            a = int(m.group(1) or 1)
            b = int(m.group(2) or n)
            if a > b:
                out += list(range(min(a, n), b - 1, -1))
            else:
                out += list(range(a, min(b, n) + 1))
        elif part.isdigit():
            out.append(int(part))
        else:
            raise ToolError(f"Không hiểu phạm vi trang '{part}' (ví dụ đúng: 1-3, 5, 8-)")
    res = [p - 1 for p in out if 1 <= p <= n]
    if not res:
        raise ToolError(f"Phạm vi trang '{spec}' không có trang nào (file có {n} trang)")
    return res


def rgb(hexstr, default=(0, 0, 0)):
    h = (hexstr or "").lstrip("#")
    if len(h) != 6:
        return default
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def paper(name, orientation="auto", like=None):
    """Khổ giấy: name trong PAPERS hoặc 'same' (= like). orientation: auto | portrait | landscape."""
    if name == "same" and like is not None:
        r = pymupdf.Rect(0, 0, like.width, like.height)
    else:
        r = pymupdf.paper_rect(PAPERS.get(name, "a4"))
    if (orientation == "landscape" and r.width < r.height) or (orientation == "portrait" and r.width > r.height):
        r = pymupdf.Rect(0, 0, r.height, r.width)
    return r


PAPER_CHOICES = [("A4", "A4"), ("A3", "A3"), ("A5", "A5"), ("Letter", "Letter"), ("Legal", "Legal")]


_ARCH = None


def archive():
    global _ARCH
    if _ARCH is None:
        _ARCH = pdf_core._archive()
    return _ARCH


def html_text(page, rect, text, size=10, color="#000000", align="left", family="Arial", bold=False,
              opacity=1.0, overlay=True):
    """Viết 1 khối chữ tiếng Việt (HTML) vào vùng `rect` (toạ độ hiển thị, trang đã bỏ xoay)."""
    css = pdf_core.font_css([family]) + (
        f"\n* {{font-family:'{family}'; font-size:{size:g}pt; color:{color}; margin:0; padding:0;"
        f" line-height:1.15; {'font-weight:bold;' if bold else ''}}} p {{text-align:{align}; margin:0}}")
    body = "".join(f"<p>{_html.escape(ln) or '&nbsp;'}</p>" for ln in str(text).split("\n"))
    page.insert_htmlbox(pymupdf.Rect(rect), body, css=css, archive=archive(), scale_low=0,
                        opacity=opacity, overlay=overlay)


def fill_vars(template, page_no, total, filename, extra=None):
    now = datetime.now()
    s = (template.replace("{trang}", str(page_no)).replace("{tong}", str(total))
         .replace("{ngay}", now.strftime("%d/%m/%Y")).replace("{gio}", now.strftime("%H:%M"))
         .replace("{ten_file}", os.path.splitext(os.path.basename(filename))[0]))
    for k, v in (extra or {}).items():
        s = s.replace("{" + k + "}", str(v))
    return s


POSITIONS = [("bc", "Dưới - giữa"), ("br", "Dưới - phải"), ("bl", "Dưới - trái"),
             ("tc", "Trên - giữa"), ("tr", "Trên - phải"), ("tl", "Trên - trái")]


def pos_rect(page_rect, pos, size, margin_pt, lines=1):
    """Vùng chữ 1 dòng tại vị trí (tl/tc/tr/bl/bc/br) + căn lề tương ứng."""
    W, H = page_rect.width, page_rect.height
    h = size * 1.45 * lines + 2
    y0 = margin_pt if pos[0] == "t" else H - margin_pt - h
    align = {"l": "left", "c": "center", "r": "right"}[pos[1]]
    return pymupdf.Rect(margin_pt, y0, W - margin_pt, y0 + h), align


def page_count_text(n):
    return f"{n} trang"


def pages_opt(label="Trang áp dụng", tip=None):
    return Opt("pages", label, "pages", "", placeholder="để trống = tất cả; vd: 1-3, 5, 8-  hoặc  le / chan",
               tip=tip or "Số trang tính từ 1. 'le' = trang lẻ, 'chan' = trang chẵn, '8-' = từ trang 8 đến cuối")


PASSWORD_OPT = Opt("password", "Mật khẩu mở file (nếu có)", "password", "",
                   tip="Chỉ cần khi file PDF đầu vào có mật khẩu mở")


def out_name_opt(default=""):
    return Opt("out_name", "Tên file kết quả", "text", default, placeholder="để trống = tự đặt tên")


def simple_page_tool(ctx, suffix, fn, fonts=False):
    """Mở từng file, gọi fn(doc, path) để sửa, lưu thành file mới."""
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        fn(doc, path)
        save_doc(doc, ctx.out(path, suffix), fonts=fonts)
        doc.close()
    ctx.each(one)


# ===========================================================================
#  1. SẮP XẾP & QUẢN LÝ TRANG
# ===========================================================================
@tool(id="merge", group="organize", name="Gộp PDF", icon=0xE8F1, popular=True, output="single", min_files=2,
      accept=OPENABLE_EXT, keywords="merge ghep noi",
      desc="Gộp nhiều file PDF / ảnh thành 1 file. Kéo thả để sắp thứ tự.",
      opts=[Opt("bookmarks", "Tạo bookmark theo tên từng file", "bool", True),
            out_name_opt("Gop_PDF")])
def t_merge(ctx):
    out = pymupdf.open()
    toc = []
    n = len(ctx.files)
    for i, f in enumerate(ctx.files):
        ctx.progress(i / n, os.path.basename(f))
        src = open_doc(f, ctx["password"])
        start = out.page_count
        if ctx["bookmarks"]:
            toc.append([1, os.path.splitext(os.path.basename(f))[0], start + 1])
        for lvl, title, pno, *_ in src.get_toc():
            toc.append([lvl + (1 if ctx["bookmarks"] else 0), title, pno + start])
        out.insert_pdf(src)
        src.close()
    if toc:
        try:
            out.set_toc(toc)
        except Exception:
            pass
    save_doc(out, ctx.out(ctx.files[0], "", name=ctx["out_name"] or "Gop_PDF"))
    ctx.log(f"Đã gộp {n} file, {out.page_count} trang.")


@tool(id="split", group="organize", name="Tách PDF", icon=0xE8C6, popular=True, keywords="split chia",
      desc="Tách 1 file thành nhiều file: theo khoảng trang, mỗi N trang, từng trang, hoặc theo bookmark.",
      opts=[Opt("mode", "Cách tách", "choice", "ranges",
                [("ranges", "Theo khoảng trang"), ("every", "Mỗi N trang một file"), ("each", "Mỗi trang một file"),
                 ("bookmarks", "Theo bookmark cấp 1")]),
            Opt("ranges", "Các khoảng trang", "text", "", placeholder="vd: 1-3, 4-10, 11-  (mỗi khoảng = 1 file)",
                show_if=("mode", ("ranges",))),
            Opt("n", "Số trang mỗi file", "int", 2, min=1, max=10000, show_if=("mode", ("every",))),
            PASSWORD_OPT])
def t_split(ctx):
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        n = doc.page_count
        mode = ctx["mode"]
        parts = []                      # [(tên, [trang])]
        if mode == "each":
            parts = [(f"_trang_{p + 1}", [p]) for p in range(n)]
        elif mode == "every":
            k = max(1, int(ctx["n"]))
            parts = [(f"_trang_{a + 1}-{min(a + k, n)}", list(range(a, min(a + k, n)))) for a in range(0, n, k)]
        elif mode == "bookmarks":
            marks = [(t, p - 1) for lvl, t, p, *_ in doc.get_toc() if lvl == 1 and p >= 1]
            if not marks:
                raise ToolError("file không có bookmark")
            marks.sort(key=lambda m: m[1])
            for j, (t, p) in enumerate(marks):
                end = marks[j + 1][1] if j + 1 < len(marks) else n
                if end > p:
                    parts.append((f"_{j + 1:02d}_{t[:60]}", list(range(p, end))))
        else:
            specs = [s for s in re.split(r"[,;]", ctx["ranges"] or "") if s.strip()]
            if not specs:
                raise ToolError("chưa nhập khoảng trang")
            for s in specs:
                parts.append((f"_trang_{s.strip().replace(' ', '')}", parse_pages(s, n)))
        folder = ctx.out(path, "_tach", folder=True)
        stem = os.path.splitext(os.path.basename(path))[0]
        for name, pages in parts:
            new = pymupdf.open()
            for p in pages:
                new.insert_pdf(doc, from_page=p, to_page=p)
            fn = re.sub(r'[<>:"/\\|?*]', "_", stem + name) + ".pdf"
            save_doc(new, os.path.join(folder, fn))
            new.close()
        ctx.log(f"{os.path.basename(path)}: {len(parts)} file → {folder}")
        doc.close()
    ctx.each(one)


@tool(id="extract_pages", group="organize", name="Trích trang", icon=0xE8A7,
      desc="Lấy các trang chọn ra thành file mới.",
      opts=[Opt("pages", "Trang cần lấy", "pages", "", placeholder="vd: 1-3, 5, 8-"), PASSWORD_OPT])
def t_extract(ctx):
    if not (ctx["pages"] or "").strip():
        raise ToolError("Hãy nhập các trang cần lấy.")

    def fn(doc, path):
        doc.select(parse_pages(ctx["pages"], doc.page_count))
    simple_page_tool(ctx, "_trich", fn)


@tool(id="delete_pages", group="organize", name="Xoá trang", icon=0xE74D,
      desc="Xoá các trang chỉ định khỏi file.",
      opts=[Opt("pages", "Trang cần xoá", "pages", "", placeholder="vd: 2, 5-7, cuoi"), PASSWORD_OPT])
def t_delete(ctx):
    def fn(doc, path):
        rm = set(parse_pages(ctx["pages"], doc.page_count, default_all=False))
        if not rm:
            raise ToolError("chưa nhập trang cần xoá")
        if len(rm) >= doc.page_count:
            raise ToolError("không thể xoá tất cả các trang")
        doc.select([p for p in range(doc.page_count) if p not in rm])
    simple_page_tool(ctx, "_xoa_trang", fn)


@tool(id="rotate", group="organize", name="Xoay trang", icon=0xE7AD,
      desc="Xoay tất cả hoặc một số trang 90° / 180° / 270°.",
      opts=[Opt("angle", "Góc xoay", "choice", 90, [(90, "90° phải"), (270, "90° trái"), (180, "180°")]),
            Opt("which", "Chỉ xoay", "choice", "all", [("all", "Mọi trang trong phạm vi"),
                                                       ("portrait", "Trang dọc"), ("landscape", "Trang ngang")]),
            pages_opt(), PASSWORD_OPT])
def t_rotate(ctx):
    def fn(doc, path):
        for p in parse_pages(ctx["pages"], doc.page_count):
            pg = doc[p]
            land = pg.rect.width > pg.rect.height
            if ctx["which"] == "portrait" and land or ctx["which"] == "landscape" and not land:
                continue
            pg.set_rotation((pg.rotation + int(ctx["angle"])) % 360)
    simple_page_tool(ctx, "_xoay", fn)


@tool(id="reverse", group="organize", name="Đảo ngược thứ tự trang", icon=0xE8CB,
      desc="Trang cuối thành trang đầu (hay dùng cho bản scan bị ngược).", opts=[PASSWORD_OPT])
def t_reverse(ctx):
    simple_page_tool(ctx, "_dao", lambda doc, p: doc.select(list(range(doc.page_count - 1, -1, -1))))


@tool(id="alternate_merge", group="organize", name="Gộp xen kẽ (scan 2 mặt)", icon=0xE8AB, output="single",
      min_files=2, max_files=2, keywords="duplex collate odd even",
      desc="Ghép file mặt trước (trang lẻ) và file mặt sau (trang chẵn) khi scan 2 mặt bằng máy 1 mặt.",
      opts=[Opt("reverse_b", "File thứ 2 có thứ tự ngược (scan lật cả xấp)", "bool", True),
            out_name_opt(), PASSWORD_OPT],
      note="File 1 = mặt trước, file 2 = mặt sau.")
def t_alternate(ctx):
    a = open_doc(ctx.files[0], ctx["password"])
    b = open_doc(ctx.files[1], ctx["password"])
    bp = list(range(b.page_count))
    if ctx["reverse_b"]:
        bp.reverse()
    out = pymupdf.open()
    for i in range(max(a.page_count, len(bp))):
        if i < a.page_count:
            out.insert_pdf(a, from_page=i, to_page=i)
        if i < len(bp):
            out.insert_pdf(b, from_page=bp[i], to_page=bp[i])
    save_doc(out, ctx.out(ctx.files[0], "_xen_ke", name=ctx["out_name"] or None))
    ctx.log(f"Kết quả {out.page_count} trang.")


@tool(id="blank_page", group="organize", name="Thêm trang trắng", icon=0xE8A5,
      desc="Chèn trang trắng vào đầu, cuối, sau một trang hoặc xen sau mỗi trang.",
      opts=[Opt("where", "Vị trí", "choice", "end", [("end", "Cuối file"), ("start", "Đầu file"),
                                                     ("after", "Sau trang số"), ("every", "Sau mỗi trang")]),
            Opt("after", "Sau trang", "int", 1, min=1, max=100000, show_if=("where", ("after",))),
            Opt("count", "Số trang trắng", "int", 1, min=1, max=100), PASSWORD_OPT])
def t_blank(ctx):
    def fn(doc, path):
        n = doc.page_count
        cnt = int(ctx["count"])
        w = ctx["where"]
        positions = {"end": [n], "start": [0], "after": [min(int(ctx["after"]), n)],
                     "every": list(range(n, 0, -1))}[w]
        for pos in positions:
            ref = doc[min(max(pos - 1, 0), n - 1)].rect
            for _ in range(cnt):
                doc.new_page(pno=pos if pos < doc.page_count else -1, width=ref.width, height=ref.height)
    simple_page_tool(ctx, "_them_trang", fn)


@tool(id="divide", group="organize", name="Chia đôi trang", icon=0xE8A1, keywords="split page half",
      desc="Cắt mỗi trang thành 2 (hoặc nhiều) trang - dùng cho bản scan 2 trang sách trên 1 mặt.",
      opts=[Opt("dir", "Chia theo", "choice", "v", [("v", "Chiều dọc (trái | phải)"), ("h", "Chiều ngang (trên / dưới)")]),
            Opt("parts", "Số phần", "int", 2, min=2, max=10), pages_opt(), PASSWORD_OPT])
def t_divide(ctx):
    def one(path, i):
        src = open_doc(path, ctx["password"])
        normalize_rotation(src)
        sel = set(parse_pages(ctx["pages"], src.page_count))
        out = pymupdf.open()
        k = int(ctx["parts"])
        for p in range(src.page_count):
            r = src[p].rect
            if p not in sel:
                out.insert_pdf(src, from_page=p, to_page=p)
                continue
            for j in range(k):
                if ctx["dir"] == "v":
                    clip = pymupdf.Rect(r.x0 + r.width * j / k, r.y0, r.x0 + r.width * (j + 1) / k, r.y1)
                else:
                    clip = pymupdf.Rect(r.x0, r.y0 + r.height * j / k, r.x1, r.y0 + r.height * (j + 1) / k)
                np_ = out.new_page(width=clip.width, height=clip.height)
                np_.show_pdf_page(np_.rect, src, p, clip=clip)
        save_doc(out, ctx.out(path, "_chia"))
    ctx.each(one)


NUP = {2: (2, 1, "landscape"), 4: (2, 2, "portrait"), 6: (3, 2, "landscape"), 8: (4, 2, "landscape"),
       9: (3, 3, "portrait"), 16: (4, 4, "portrait")}


@tool(id="nup", group="organize", name="In nhiều trang trên 1 tờ (N-up)", icon=0xF0E2, keywords="n-up in tiet kiem",
      desc="In 2, 4, 6, 9, 16 trang trên một tờ giấy để tiết kiệm giấy.",
      opts=[Opt("n", "Số trang / tờ", "choice", 2, [(k, f"{k} trang") for k in NUP]),
            Opt("paper", "Khổ giấy", "choice", "A4", PAPER_CHOICES),
            Opt("order", "Thứ tự", "choice", "row", [("row", "Ngang trước (trái→phải)"), ("col", "Dọc trước (trên→dưới)")]),
            Opt("margin", "Lề", "float", 5, min=0, max=50, suffix=" mm"),
            Opt("border", "Kẻ khung mỗi trang", "bool", False), PASSWORD_OPT])
def t_nup(ctx):
    def one(path, i):
        src = open_doc(path, ctx["password"])
        normalize_rotation(src)
        cols, rows, orient = NUP[int(ctx["n"])]
        sheet = paper(ctx["paper"], orient)
        m = float(ctx["margin"]) * MM
        cw, ch = (sheet.width - 2 * m) / cols, (sheet.height - 2 * m) / rows
        out = pymupdf.open()
        per = cols * rows
        for s in range(0, src.page_count, per):
            pg = out.new_page(width=sheet.width, height=sheet.height)
            for k in range(per):
                if s + k >= src.page_count:
                    break
                c, r = (k % cols, k // cols) if ctx["order"] == "row" else (k // rows, k % rows)
                cell = pymupdf.Rect(m + c * cw, m + r * ch, m + (c + 1) * cw, m + (r + 1) * ch)
                inner = cell + (2, 2, -2, -2)
                pg.show_pdf_page(inner, src, s + k)
                if ctx["border"]:
                    pg.draw_rect(inner, color=(0.6, 0.6, 0.6), width=0.5)
        save_doc(out, ctx.out(path, f"_{ctx['n']}up"))
    ctx.each(one)


@tool(id="booklet", group="organize", name="In kiểu sách gập (booklet)", icon=0xE82D, keywords="booklet saddle",
      desc="Sắp trang để in 2 mặt, gập đôi thành cuốn sách nhỏ (2 trang A5 trên 1 tờ A4 ngang).",
      opts=[Opt("paper", "Khổ tờ in", "choice", "A4", PAPER_CHOICES), PASSWORD_OPT],
      note="In 2 mặt, lật theo cạnh ngắn, rồi gập đôi.")
def t_booklet(ctx):
    def one(path, i):
        src = open_doc(path, ctx["password"])
        normalize_rotation(src)
        n = src.page_count
        total = (n + 3) // 4 * 4
        sheet = paper(ctx["paper"], "landscape")
        half = sheet.width / 2
        out = pymupdf.open()
        for s in range(total // 4):
            for left, right in ((total - 2 * s, 1 + 2 * s), (2 + 2 * s, total - 1 - 2 * s)):
                pg = out.new_page(width=sheet.width, height=sheet.height)
                for k, num in enumerate((left, right)):
                    if num <= n:
                        cell = pymupdf.Rect(k * half, 0, (k + 1) * half, sheet.height) + (8, 8, -8, -8)
                        pg.show_pdf_page(cell, src, num - 1)
        save_doc(out, ctx.out(path, "_booklet"))
        ctx.log(f"{os.path.basename(path)}: {n} trang → {out.page_count} mặt in ({out.page_count // 2} tờ)")
    ctx.each(one)


@tool(id="poster", group="organize", name="In khổ lớn ghép nhiều tờ (poster)", icon=0xF0E3, keywords="tile poster",
      desc="Phóng to 1 trang ra nhiều tờ (vd 2x2 tờ A4) để in ghép thành poster.",
      opts=[Opt("cols", "Số cột", "int", 2, min=1, max=10), Opt("rows", "Số hàng", "int", 2, min=1, max=10),
            Opt("paper", "Khổ tờ in", "choice", "A4", PAPER_CHOICES), pages_opt(), PASSWORD_OPT])
def t_poster(ctx):
    def one(path, i):
        src = open_doc(path, ctx["password"])
        normalize_rotation(src)
        cols, rows = int(ctx["cols"]), int(ctx["rows"])
        out = pymupdf.open()
        for p in parse_pages(ctx["pages"], src.page_count):
            r = src[p].rect
            tw, th = r.width / cols, r.height / rows
            sheet = paper(ctx["paper"], "landscape" if tw > th else "portrait")
            for rr in range(rows):
                for cc in range(cols):
                    clip = pymupdf.Rect(r.x0 + cc * tw, r.y0 + rr * th, r.x0 + (cc + 1) * tw, r.y0 + (rr + 1) * th)
                    pg = out.new_page(width=sheet.width, height=sheet.height)
                    pg.show_pdf_page(sheet + (10, 10, -10, -10), src, p, clip=clip)
        save_doc(out, ctx.out(path, "_poster"))
    ctx.each(one)


@tool(id="crop", group="organize", name="Cắt lề trang", icon=0xE7A8, keywords="crop trim margin",
      desc="Cắt bớt lề theo số mm, hoặc tự cắt sát nội dung (bỏ khoảng trắng thừa).",
      opts=[Opt("mode", "Cách cắt", "choice", "auto", [("auto", "Tự động sát nội dung"), ("manual", "Theo số mm")]),
            Opt("pad", "Chừa lề quanh nội dung", "float", 5, min=0, max=100, suffix=" mm", show_if=("mode", ("auto",))),
            Opt("top", "Cắt trên", "float", 10, min=0, max=500, suffix=" mm", show_if=("mode", ("manual",))),
            Opt("bottom", "Cắt dưới", "float", 10, min=0, max=500, suffix=" mm", show_if=("mode", ("manual",))),
            Opt("left", "Cắt trái", "float", 10, min=0, max=500, suffix=" mm", show_if=("mode", ("manual",))),
            Opt("right", "Cắt phải", "float", 10, min=0, max=500, suffix=" mm", show_if=("mode", ("manual",))),
            pages_opt(), PASSWORD_OPT])
def t_crop(ctx):
    def fn(doc, path):
        normalize_rotation(doc)
        for p in parse_pages(ctx["pages"], doc.page_count):
            pg = doc[p]
            r = pg.rect
            if ctx["mode"] == "manual":
                new = pymupdf.Rect(r.x0 + ctx["left"] * MM, r.y0 + ctx["top"] * MM,
                                   r.x1 - ctx["right"] * MM, r.y1 - ctx["bottom"] * MM)
            else:
                box = pymupdf.Rect()
                for kind, bb in pg.get_bboxlog():
                    bb = pymupdf.Rect(bb) & r
                    if bb.is_empty or bb.get_area() > 0.9 * r.get_area():
                        continue                    # bỏ nền phủ cả trang
                    box |= bb
                if box.is_empty:
                    continue
                pad = ctx["pad"] * MM
                new = (box + (-pad, -pad, pad, pad)) & r
            if new.width > 20 and new.height > 20:
                pg.set_cropbox(new + (pg.cropbox.x0, pg.cropbox.y0, pg.cropbox.x0, pg.cropbox.y0))
    simple_page_tool(ctx, "_cat_le", fn)


@tool(id="fix_size", group="organize", name="Đồng bộ khổ giấy", icon=0xE740, keywords="resize a4 page size",
      desc="Đưa mọi trang về cùng một khổ (vd A4), nội dung tự co giãn vừa trang.",
      opts=[Opt("paper", "Khổ giấy", "choice", "A4", PAPER_CHOICES),
            Opt("orient", "Hướng", "choice", "auto", [("auto", "Giữ hướng từng trang"), ("portrait", "Dọc"),
                                                      ("landscape", "Ngang")]),
            Opt("margin", "Lề", "float", 0, min=0, max=50, suffix=" mm"), PASSWORD_OPT])
def t_fix_size(ctx):
    def one(path, i):
        src = open_doc(path, ctx["password"])
        normalize_rotation(src)
        out = pymupdf.open()
        m = ctx["margin"] * MM
        for p in range(src.page_count):
            r = src[p].rect
            orient = ctx["orient"] if ctx["orient"] != "auto" else ("landscape" if r.width > r.height else "portrait")
            sheet = paper(ctx["paper"], orient)
            pg = out.new_page(width=sheet.width, height=sheet.height)
            pg.show_pdf_page(sheet + (m, m, -m, -m), src, p)
        save_doc(out, ctx.out(path, f"_{ctx['paper']}"))
    ctx.each(one)


@tool(id="margins", group="organize", name="Thêm lề trang", icon=0xE9A6, keywords="margin le them khoang trang",
      desc="Nới rộng trang thêm lề trắng (vd để đóng gáy, ghi chú bên lề) - nội dung, ghi chú, liên kết giữ nguyên.",
      opts=[Opt("top", "Lề trên", "float", 10, min=0, max=300, suffix=" mm"),
            Opt("bottom", "Lề dưới", "float", 10, min=0, max=300, suffix=" mm"),
            Opt("left", "Lề trái", "float", 10, min=0, max=300, suffix=" mm"),
            Opt("right", "Lề phải", "float", 10, min=0, max=300, suffix=" mm"),
            pages_opt(), PASSWORD_OPT])
def t_margins(ctx):
    def fn(doc, path):
        normalize_rotation(doc)
        for p in parse_pages(ctx["pages"], doc.page_count):
            pg = doc[p]
            t, b, l, r = (ctx[k] * MM for k in ("top", "bottom", "left", "right"))
            # vùng trang theo toạ độ gốc PDF (y hướng lên): CropBox nếu có, không thì MediaBox (có thể kế thừa)
            box = []
            for key in ("CropBox", "MediaBox"):
                typ, val = doc.xref_get_key(pg.xref, key)
                if typ == "array":
                    box = [float(v) for v in re.findall(r"[-+]?\d*\.?\d+", val)]
                    break
            if len(box) != 4:
                box = list(pg.mediabox)
            nb = f"[{box[0] - l:g} {box[1] - b:g} {box[2] + r:g} {box[3] + t:g}]"
            doc.xref_set_key(pg.xref, "MediaBox", nb)
            doc.xref_set_key(pg.xref, "CropBox", nb)
    simple_page_tool(ctx, "_them_le", fn)


def _is_blank(page, threshold):
    if page.get_text("text").strip():
        return False
    pix = page.get_pixmap(dpi=25, colorspace=pymupdf.csGRAY, alpha=False)
    data = pix.samples
    try:
        import numpy as np
        dark = int((np.frombuffer(data, dtype=np.uint8) < 200).sum())
    except ImportError:
        dark = sum(1 for b in data if b < 200)
    return dark / max(1, len(data)) * 100 < threshold


@tool(id="remove_blank", group="organize", name="Xoá trang trắng", icon=0xE894, keywords="blank empty",
      desc="Tự phát hiện và xoá trang trắng (kể cả trang scan gần như trắng).",
      opts=[Opt("threshold", "Độ nhạy (% điểm tối tối đa)", "float", 0.5, min=0, max=10, step=0.1,
                tip="Trang có ít hơn tỉ lệ điểm tối này được coi là trắng. Tăng nếu trang scan có nhiều chấm bẩn"),
            PASSWORD_OPT])
def t_remove_blank(ctx):
    def fn(doc, path):
        keep = [p for p in range(doc.page_count) if not _is_blank(doc[p], float(ctx["threshold"]))]
        removed = doc.page_count - len(keep)
        if not keep:
            raise ToolError("tất cả các trang đều trắng")
        doc.select(keep)
        ctx.log(f"{os.path.basename(path)}: xoá {removed} trang trắng")
    simple_page_tool(ctx, "_bo_trang_trang", fn)


def _toc_text(path):
    try:
        d = pymupdf.open(path)
        t = "\n".join(f"{lvl} | {title} | {p}" for lvl, title, p, *_ in d.get_toc())
        d.close()
        return t
    except Exception:
        return ""


@tool(id="bookmarks", group="organize", name="Sửa mục lục điều hướng (bookmark)", icon=0xE8A4,
      keywords="bookmark outline toc",
      desc="Xem / sửa / thêm bookmark. Mỗi dòng: cấp | tiêu đề | trang.",
      opts=[Opt("toc", "Bookmark", "textarea", "", prefill=_toc_text,
                placeholder="1 | Chương 1 | 1\n2 | Mục 1.1 | 3\n1 | Chương 2 | 10"),
            PASSWORD_OPT])
def t_bookmarks(ctx):
    toc = []
    for ln in (ctx["toc"] or "").splitlines():
        if not ln.strip():
            continue
        parts = [x.strip() for x in ln.split("|")]
        if len(parts) != 3 or not parts[0].isdigit() or not parts[2].isdigit():
            raise ToolError(f"Dòng bookmark sai định dạng: '{ln}' (đúng: cấp | tiêu đề | trang)")
        toc.append([int(parts[0]), parts[1], int(parts[2])])
    if toc and toc[0][0] != 1:
        raise ToolError("Bookmark đầu tiên phải là cấp 1")

    def fn(doc, path):
        doc.set_toc(toc)
    simple_page_tool(ctx, "_bookmark", fn)


@tool(id="toc_page", group="organize", name="Tạo trang mục lục", icon=0xE8FD, keywords="table of contents",
      desc="Chèn trang mục lục (có số trang, bấm để nhảy tới) từ bookmark của file.",
      opts=[Opt("title", "Tiêu đề", "text", "MỤC LỤC"), Opt("size", "Cỡ chữ", "float", 11, min=6, max=24),
            Opt("max_level", "Hiện tới cấp", "int", 2, min=1, max=6), PASSWORD_OPT])
def t_toc_page(ctx):
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        toc = [t for t in doc.get_toc() if t[0] <= int(ctx["max_level"])]
        if not toc:
            raise ToolError("file không có bookmark - dùng công cụ 'Sửa bookmark' để thêm trước")
        size = float(ctx["size"])
        r0 = doc[0].rect
        line_h = size * 1.6
        per_page = int((r0.height - 140) // line_h)
        n_toc = (len(toc) + per_page - 1) // per_page
        for k in range(n_toc):                          # chèn trang mục lục vào đầu
            doc.new_page(pno=k, width=r0.width, height=r0.height)
        for k in range(n_toc):
            pg = doc[k]
            y = 60
            if k == 0:
                html_text(pg, pymupdf.Rect(60, y, r0.width - 60, y + 40), ctx["title"], size * 1.6, bold=True,
                          align="center")
                y += 50
            for lvl, title, pno, *_ in toc[k * per_page:(k + 1) * per_page]:
                target = pno - 1 + n_toc
                indent = 60 + (lvl - 1) * 18
                row = pymupdf.Rect(indent, y, r0.width - 60, y + line_h)
                html_text(pg, pymupdf.Rect(row.x0, row.y0, row.x1 - 40, row.y1), title, size, bold=lvl == 1)
                html_text(pg, pymupdf.Rect(row.x1 - 40, row.y0, row.x1, row.y1), str(target + 1), size, align="right")
                pg.insert_link({"kind": pymupdf.LINK_GOTO, "from": row, "page": target, "to": pymupdf.Point(0, 0)})
                y += line_h
        doc.set_toc([[lvl, t, p + n_toc] for lvl, t, p, *_ in doc.get_toc()])
        save_doc(doc, ctx.out(path, "_muc_luc"), fonts=True)
    ctx.each(one)


@tool(id="add_attach", group="organize", name="Đính kèm file vào PDF", icon=0xE723, keywords="attachment embed",
      desc="Nhúng file bất kỳ (Excel, Word, ảnh...) vào bên trong PDF.",
      opts=[Opt("attach", "File đính kèm", "files", [], filter="Tất cả (*.*)"), PASSWORD_OPT])
def t_add_attach(ctx):
    if not ctx["attach"]:
        raise ToolError("Chưa chọn file đính kèm.")

    def fn(doc, path):
        names = set(doc.embfile_names())
        for f in ctx["attach"]:
            name = os.path.basename(f)
            k = 2
            while name in names:
                stem, ext = os.path.splitext(os.path.basename(f))
                name = f"{stem} ({k}){ext}"
                k += 1
            names.add(name)
            with open(f, "rb") as fh:
                doc.embfile_add(name, fh.read(), filename=os.path.basename(f))
    simple_page_tool(ctx, "_dinh_kem", fn)


@tool(id="extract_attach", group="organize", name="Xuất file đính kèm ra thư mục", icon=0xE896, keywords="attachment",
      desc="Lưu các file được nhúng trong PDF (và file đính kèm trong ghi chú) ra thư mục.", opts=[PASSWORD_OPT])
def t_extract_attach(ctx):
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        items = [(n, doc.embfile_get(n)) for n in doc.embfile_names()]
        for pg in doc:
            for a in pg.annots(types=[pymupdf.PDF_ANNOT_FILE_ATTACHMENT]) or []:
                info = a.file_info
                items.append((info.get("filename") or f"dinh_kem_{len(items)}", a.get_file()))
        if not items:
            ctx.log(f"{os.path.basename(path)}: không có file đính kèm")
            return
        folder = ctx.out(path, "_dinh_kem", folder=True)
        for name, data in items:
            with open(os.path.join(folder, re.sub(r'[<>:"/\\|?*]', "_", name)), "wb") as fh:
                fh.write(data)
        ctx.log(f"{os.path.basename(path)}: {len(items)} file → {folder}")
    ctx.each(one)


@tool(id="overlay", group="organize", name="Lồng mẫu nền / tiêu đề thư (letterhead)", icon=0xE81E,
      keywords="overlay letterhead template background",
      desc="Đặt trang của một file PDF khác (giấy tiêu đề công ty, khung mẫu...) lên trên hoặc dưới mọi trang.",
      opts=[Opt("overlay_file", "File PDF mẫu", "file", "", filter="PDF (*.pdf)"),
            Opt("mode", "Vị trí", "choice", "under", [("under", "Làm nền (nằm dưới nội dung)"), ("over", "Đè lên trên")]),
            Opt("repeat", "Trang mẫu", "choice", "first", [("first", "Dùng trang 1 của mẫu cho mọi trang"),
                                                         ("match", "Trang N của mẫu cho trang N")]),
            pages_opt(), PASSWORD_OPT])
def t_overlay(ctx):
    if not ctx["overlay_file"]:
        raise ToolError("Chưa chọn file PDF mẫu.")
    tpl = open_doc(ctx["overlay_file"])

    def fn(doc, path):
        normalize_rotation(doc)
        for p in parse_pages(ctx["pages"], doc.page_count):
            k = 0 if ctx["repeat"] == "first" else p
            if k >= tpl.page_count:
                continue
            doc[p].show_pdf_page(doc[p].rect, tpl, k, overlay=ctx["mode"] == "over", keep_proportion=False)
    simple_page_tool(ctx, "_chong_trang", fn)


@tool(id="compare", group="organize", name="So sánh 2 file PDF", icon=0xE8F4, output="single", min_files=2,
      max_files=2, keywords="compare diff khac nhau",
      desc="Đặt 2 bản cạnh nhau, tô đỏ chữ bị xoá ở bản 1 và tô xanh chữ mới thêm ở bản 2.",
      opts=[out_name_opt(), PASSWORD_OPT], note="File 1 = bản cũ, file 2 = bản mới.")
def t_compare(ctx):
    import difflib
    a = open_doc(ctx.files[0], ctx["password"])
    b = open_doc(ctx.files[1], ctx["password"])
    normalize_rotation(a)
    normalize_rotation(b)
    out = pymupdf.open()
    n = max(a.page_count, b.page_count)
    changed_pages, n_del, n_add = [], 0, 0
    for p in range(n):
        ctx.progress(p / n, f"Trang {p + 1}")
        wa = a[p].get_text("words") if p < a.page_count else []
        wb = b[p].get_text("words") if p < b.page_count else []
        sm = difflib.SequenceMatcher(None, [w[4] for w in wa], [w[4] for w in wb], autojunk=False)
        dels, adds = [], []
        for op, i1, i2, j1, j2 in sm.get_opcodes():
            if op in ("delete", "replace"):
                dels += [pymupdf.Rect(w[:4]) for w in wa[i1:i2]]
            if op in ("insert", "replace"):
                adds += [pymupdf.Rect(w[:4]) for w in wb[j1:j2]]
        if dels or adds or (p >= a.page_count) != (p >= b.page_count):
            changed_pages.append(p + 1)
        n_del += len(dels)
        n_add += len(adds)
        for d, rects, col in ((a, dels, (1, 0.2, 0.2)), (b, adds, (0.1, 0.8, 0.2))):
            if p < d.page_count:
                for r in rects:
                    d[p].draw_rect(r, color=None, fill=col, fill_opacity=0.35, overlay=True)
        ra = a[p].rect if p < a.page_count else b[p].rect
        rb = b[p].rect if p < b.page_count else ra
        gap = 20
        pg = out.new_page(width=ra.width + rb.width + gap, height=max(ra.height, rb.height) + 30)
        html_text(pg, pymupdf.Rect(0, 4, ra.width, 24), f"Bản 1 - trang {p + 1}", 10, "#b91c1c", "center")
        html_text(pg, pymupdf.Rect(ra.width + gap, 4, pg.rect.width, 24), f"Bản 2 - trang {p + 1}", 10, "#15803d",
                  "center")
        if p < a.page_count:
            pg.show_pdf_page(pymupdf.Rect(0, 30, ra.width, 30 + ra.height), a, p)
        if p < b.page_count:
            pg.show_pdf_page(pymupdf.Rect(ra.width + gap, 30, ra.width + gap + rb.width, 30 + rb.height), b, p)
        pg.draw_line((ra.width + gap / 2, 0), (ra.width + gap / 2, pg.rect.height), color=(0.7, 0.7, 0.7), width=0.5)
    save_doc(out, ctx.out(ctx.files[0], "_so_sanh", name=ctx["out_name"] or None), fonts=True)
    if changed_pages:
        ctx.log(f"Khác nhau ở {len(changed_pages)} trang: {', '.join(map(str, changed_pages[:40]))}"
                f"{'...' if len(changed_pages) > 40 else ''}  (bị xoá {n_del} từ, thêm {n_add} từ)")
    else:
        ctx.log("Hai file giống nhau về nội dung chữ.")


# ===========================================================================
#  2. ĐÓNG DẤU & ĐÁNH SỐ
# ===========================================================================
FONT_CHOICES = [("Arial", "Arial"), ("Times New Roman", "Times New Roman"), ("Tahoma", "Tahoma"),
                ("Calibri", "Calibri"), ("Verdana", "Verdana"), ("Courier New", "Courier New")]


def _font_path(family, bold=False):
    files = pdf_core.vn_fonts().get(family) or pdf_core.vn_fonts()["Arial"]
    return files.get("b" if bold else "r") or files.get("r")


@tool(id="watermark", group="stamp", name="Thêm watermark", icon=0xE7B8, popular=True, keywords="watermark mo",
      desc="Chèn chữ mờ (MẬT, BẢN NHÁP, tên công ty...) hoặc logo mờ lên các trang.",
      opts=[Opt("type", "Loại", "choice", "text", [("text", "Chữ"), ("image", "Ảnh / logo")]),
            Opt("text", "Nội dung", "text", "BẢN NHÁP", show_if=("type", ("text",))),
            Opt("font", "Font", "choice", "Arial", FONT_CHOICES, show_if=("type", ("text",))),
            Opt("size", "Cỡ chữ", "float", 60, min=6, max=400, show_if=("type", ("text",))),
            Opt("bold", "In đậm", "bool", True, show_if=("type", ("text",))),
            Opt("color", "Màu", "color", "#dc2626", show_if=("type", ("text",))),
            Opt("image", "File ảnh", "file", "", filter="Ảnh (*.png *.jpg *.jpeg *.bmp)", show_if=("type", ("image",))),
            Opt("scale", "Độ rộng ảnh (% trang)", "float", 50, min=5, max=100, show_if=("type", ("image",))),
            Opt("opacity", "Độ đậm", "float", 25, min=1, max=100, suffix=" %"),
            Opt("angle", "Góc nghiêng", "float", 45, min=-180, max=180, suffix=" °"),
            Opt("layout", "Bố cục", "choice", "center", [("center", "Một lần giữa trang"), ("tile", "Lặp kín trang")]),
            Opt("layer", "Lớp", "choice", "over", [("over", "Đè lên nội dung"), ("under", "Nằm dưới nội dung")]),
            pages_opt(), PASSWORD_OPT])
def t_watermark(ctx):
    op = float(ctx["opacity"]) / 100
    angle = float(ctx["angle"])
    if ctx["type"] == "image":
        if not ctx["image"]:
            raise ToolError("Chưa chọn file ảnh.")
        from PIL import Image
        im = Image.open(ctx["image"]).convert("RGBA")
        a = im.getchannel("A").point(lambda v: int(v * op))
        im.putalpha(a)
        if angle:
            im = im.rotate(angle, expand=True, resample=Image.BICUBIC)
        buf = io.BytesIO()
        im.save(buf, "PNG")
        img_bytes, ratio = buf.getvalue(), im.height / im.width
    else:
        text = ctx["text"] or ""
        if not text.strip():
            raise ToolError("Chưa nhập nội dung watermark.")
        font = pymupdf.Font(fontfile=_font_path(ctx["font"], ctx["bold"]))
        size = float(ctx["size"])

    def fn(doc, path):
        normalize_rotation(doc)
        for p in parse_pages(ctx["pages"], doc.page_count):
            pg = doc[p]
            r = pg.rect
            c = pymupdf.Point(r.x0 + r.width / 2, r.y0 + r.height / 2)
            over = ctx["layer"] == "over"
            if ctx["type"] == "image":
                w = r.width * float(ctx["scale"]) / 100
                h = w * ratio
                if ctx["layout"] == "tile":
                    centers = [pymupdf.Point(x, y) for y in _steps(r.y0, r.y1, h * 1.6) for x in _steps(r.x0, r.x1, w * 1.4)]
                else:
                    centers = [c]
                for cc in centers:
                    pg.insert_image(pymupdf.Rect(cc.x - w / 2, cc.y - h / 2, cc.x + w / 2, cc.y + h / 2),
                                    stream=img_bytes, overlay=over)
                continue
            tw = pymupdf.TextWriter(r, opacity=op, color=rgb(ctx["color"]))
            w = font.text_length(text, fontsize=size)
            if ctx["layout"] == "tile":
                diag = (r.width ** 2 + r.height ** 2) ** 0.5
                centers = [pymupdf.Point(x, y) for y in _steps(c.y - diag / 2, c.y + diag / 2, size * 3.2)
                           for x in _steps(c.x - diag / 2, c.x + diag / 2, w + size * 2)]
            else:
                centers = [c]
            for cc in centers:
                tw.append(pymupdf.Point(cc.x - w / 2, cc.y + size * 0.35), text, font=font, fontsize=size)
            tw.write_text(pg, morph=(c, pymupdf.Matrix(angle)), overlay=over)
    simple_page_tool(ctx, "_watermark", fn, fonts=True)


def _steps(a, b, step):
    n = max(1, int((b - a) // step) + 1)
    off = a + ((b - a) - (n - 1) * step) / 2
    return [off + k * step for k in range(n)]


@tool(id="header_footer", group="stamp", name="Đầu trang / chân trang", icon=0xE8BC, keywords="header footer tieu de chan trang",
      desc="Thêm dòng chữ đầu trang / cuối trang. Dùng được {trang}, {tong}, {ngay}, {gio}, {ten_file}.",
      opts=[Opt("hl", "Đầu trang - trái", "text", ""), Opt("hc", "Đầu trang - giữa", "text", ""),
            Opt("hr", "Đầu trang - phải", "text", ""),
            Opt("fl", "Cuối trang - trái", "text", "{ten_file}"), Opt("fc", "Cuối trang - giữa", "text", ""),
            Opt("fr", "Cuối trang - phải", "text", "Trang {trang}/{tong}"),
            Opt("font", "Font", "choice", "Arial", FONT_CHOICES), Opt("size", "Cỡ chữ", "float", 9, min=5, max=48),
            Opt("color", "Màu chữ", "color", "#404040"), Opt("margin", "Cách mép giấy", "float", 10, min=2, max=60, suffix=" mm"),
            Opt("line", "Kẻ đường phân cách", "bool", False),
            pages_opt(), PASSWORD_OPT])
def t_header_footer(ctx):
    fields = [(k, ctx[k]) for k in ("hl", "hc", "hr", "fl", "fc", "fr") if (ctx[k] or "").strip()]
    if not fields:
        raise ToolError("Chưa nhập nội dung header / footer.")

    def fn(doc, path):
        normalize_rotation(doc)
        total = doc.page_count
        size, m = float(ctx["size"]), float(ctx["margin"]) * MM
        for p in parse_pages(ctx["pages"], total):
            pg = doc[p]
            for k, tpl in fields:
                rect, align = pos_rect(pg.rect, ("t" if k[0] == "h" else "b") + {"l": "l", "c": "c", "r": "r"}[k[1]],
                                       size, m)
                html_text(pg, rect, fill_vars(tpl, p + 1, total, path), size, ctx["color"], align, ctx["font"])
            if ctx["line"]:
                for top in {k[0] for k, _ in fields}:
                    y = m + size * 1.6 if top == "h" else pg.rect.height - m - size * 1.6
                    pg.draw_line((m, y), (pg.rect.width - m, y), color=rgb(ctx["color"]), width=0.5)
    simple_page_tool(ctx, "_header_footer", fn, fonts=True)


@tool(id="page_numbers", group="stamp", name="Đánh số trang", icon=0xE8EF, popular=True, keywords="page number so trang",
      desc="Đánh số trang với nhiều kiểu: 1, Trang 1, Trang 1/10, - 1 -...",
      opts=[Opt("fmt", "Kiểu", "choice", "{n}", [("{n}", "1"), ("Trang {n}", "Trang 1"), ("Trang {n}/{tong}", "Trang 1/10"),
                                              ("{n}/{tong}", "1/10"), ("- {n} -", "- 1 -"), ("Page {n} of {tong}", "Page 1 of 10"),
                                              ("custom", "Tuỳ chỉnh...")]),
            Opt("custom", "Mẫu tuỳ chỉnh", "text", "Trang {n}", show_if=("fmt", ("custom",)),
                tip="{n} = số trang, {tong} = tổng số trang"),
            Opt("pos", "Vị trí", "choice", "bc", POSITIONS),
            Opt("start_page", "Bắt đầu đánh từ trang", "int", 1, min=1, max=100000,
                tip="vd 2 = bỏ qua trang bìa"),
            Opt("start_num", "Số bắt đầu", "int", 1, min=0, max=1000000),
            Opt("font", "Font", "choice", "Arial", FONT_CHOICES), Opt("size", "Cỡ chữ", "float", 10, min=5, max=48),
            Opt("color", "Màu chữ", "color", "#000000"), Opt("margin", "Cách mép giấy", "float", 10, min=2, max=60, suffix=" mm"),
            Opt("mirror", "Đổi bên trang chẵn / lẻ (in 2 mặt)", "bool", False), PASSWORD_OPT])
def t_page_numbers(ctx):
    fmt = ctx["custom"] if ctx["fmt"] == "custom" else ctx["fmt"]

    def fn(doc, path):
        normalize_rotation(doc)
        first = int(ctx["start_page"]) - 1
        total_numbered = doc.page_count - first + int(ctx["start_num"]) - 1
        for p in range(first, doc.page_count):
            pg = doc[p]
            n = p - first + int(ctx["start_num"])
            pos = ctx["pos"]
            if ctx["mirror"] and n % 2 == 0 and pos[1] in "lr":
                pos = pos[0] + ("r" if pos[1] == "l" else "l")
            rect, align = pos_rect(pg.rect, pos, float(ctx["size"]), float(ctx["margin"]) * MM)
            text = fmt.replace("{n}", str(n)).replace("{tong}", str(total_numbered))
            html_text(pg, rect, text, float(ctx["size"]), ctx["color"], align, ctx["font"])
    simple_page_tool(ctx, "_danh_so", fn, fonts=True)


@tool(id="bates", group="stamp", name="Đánh số Bates (hồ sơ)", icon=0xE8EC, keywords="bates numbering ho so",
      desc="Đánh mã hồ sơ liên tục qua nhiều file (vd HS-000001, HS-000002...).",
      opts=[Opt("prefix", "Tiền tố", "text", "HS-"), Opt("suffix", "Hậu tố", "text", ""),
            Opt("start", "Số bắt đầu", "int", 1, min=0, max=10 ** 9), Opt("digits", "Số chữ số", "int", 6, min=1, max=12),
            Opt("continuous", "Đánh liên tục qua các file trong danh sách", "bool", True),
            Opt("pos", "Vị trí", "choice", "br", POSITIONS), Opt("size", "Cỡ chữ", "float", 9, min=5, max=36),
            Opt("color", "Màu chữ", "color", "#000000"), Opt("margin", "Cách mép giấy", "float", 8, min=2, max=60, suffix=" mm"),
            PASSWORD_OPT])
def t_bates(ctx):
    counter = [int(ctx["start"])]

    def fn(doc, path):
        normalize_rotation(doc)
        if not ctx["continuous"]:
            counter[0] = int(ctx["start"])
        first = counter[0]
        for pg in doc:
            code = f"{ctx['prefix']}{counter[0]:0{int(ctx['digits'])}d}{ctx['suffix']}"
            rect, align = pos_rect(pg.rect, ctx["pos"], float(ctx["size"]), float(ctx["margin"]) * MM)
            html_text(pg, rect, code, float(ctx["size"]), ctx["color"], align, "Arial", bold=True)
            counter[0] += 1
        ctx.log(f"{os.path.basename(path)}: {ctx['prefix']}{first:0{int(ctx['digits'])}d} → "
                f"{ctx['prefix']}{counter[0] - 1:0{int(ctx['digits'])}d}")
    simple_page_tool(ctx, "_bates", fn, fonts=True)


STAMPS = ["ĐÃ DUYỆT", "ĐÃ KIỂM TRA", "BẢN SAO", "BẢN CHÍNH", "MẬT", "KHẨN", "BẢN NHÁP", "ĐÃ THANH TOÁN",
          "ĐÃ NHẬN", "HUỶ", "APPROVED", "CONFIDENTIAL", "COPY", "DRAFT"]


@tool(id="stamp", group="stamp", name="Đóng dấu chữ", icon=0xE8C8, keywords="stamp con dau da duyet mat",
      desc="Đóng dấu khung chữ như ĐÃ DUYỆT, BẢN SAO, MẬT... kèm ngày (tuỳ chọn).",
      opts=[Opt("preset", "Dấu", "choice", "ĐÃ DUYỆT", [(s, s) for s in STAMPS] + [("custom", "Tuỳ chỉnh...")]),
            Opt("custom", "Nội dung dấu", "text", "", show_if=("preset", ("custom",))),
            Opt("sub", "Dòng phụ", "text", "{ngay}", tip="vd: {ngay}, tên người duyệt... để trống = không có"),
            Opt("color", "Màu", "color", "#dc2626"), Opt("size", "Cỡ chữ", "float", 22, min=8, max=120),
            Opt("pos", "Vị trí", "choice", "tr", [("tr", "Trên - phải"), ("tl", "Trên - trái"), ("tc", "Trên - giữa"),
                                                  ("cc", "Giữa trang"), ("br", "Dưới - phải"), ("bl", "Dưới - trái")]),
            Opt("which", "Trang", "choice", "first", [("first", "Trang đầu"), ("all", "Mọi trang"), ("last", "Trang cuối")]),
            Opt("opacity", "Độ đậm", "float", 85, min=10, max=100, suffix=" %"), PASSWORD_OPT])
def t_stamp(ctx):
    text = ctx["custom"] if ctx["preset"] == "custom" else ctx["preset"]
    if not (text or "").strip():
        raise ToolError("Chưa nhập nội dung dấu.")
    size = float(ctx["size"])
    font = pymupdf.Font(fontfile=_font_path("Arial", True))
    col = rgb(ctx["color"])

    def fn(doc, path):
        normalize_rotation(doc)
        n = doc.page_count
        pages = {"first": [0], "all": range(n), "last": [n - 1]}[ctx["which"]]
        sub = fill_vars(ctx["sub"] or "", 1, n, path)
        sub_size = size * 0.45
        w = max(font.text_length(text, size), font.text_length(sub, sub_size) if sub else 0) + size * 1.2
        h = size * 1.5 + (sub_size * 1.5 if sub else 0) + size * 0.3
        for p in pages:
            pg = doc[p]
            r = pg.rect
            m = 30
            x = {"l": m, "c": (r.width - w) / 2, "r": r.width - m - w}[ctx["pos"][1]]
            y = {"t": m, "c": (r.height - h) / 2, "b": r.height - m - h}[ctx["pos"][0]]
            box = pymupdf.Rect(x, y, x + w, y + h)
            op = float(ctx["opacity"]) / 100
            pg.draw_rect(box, color=col, width=size / 9, radius=0.12, stroke_opacity=op)
            pg.draw_rect(box + (size / 6, size / 6, -size / 6, -size / 6), color=col, width=size / 30, radius=0.1,
                         stroke_opacity=op)
            html_text(pg, pymupdf.Rect(box.x0, box.y0 + size * 0.2, box.x1, box.y0 + size * 1.7), text, size,
                      ctx["color"], "center", bold=True, opacity=op)
            if sub:
                html_text(pg, pymupdf.Rect(box.x0, box.y0 + size * 1.55, box.x1, box.y1), sub, sub_size, ctx["color"],
                          "center", bold=True, opacity=op)
    simple_page_tool(ctx, "_dong_dau", fn, fonts=True)


@tool(id="bg_color", group="stamp", name="Đổi màu nền trang", icon=0xE790, keywords="background color",
      desc="Tô màu nền cho trang (nằm dưới nội dung).",
      opts=[Opt("color", "Màu nền", "color", "#fff7d6"), pages_opt(), PASSWORD_OPT])
def t_bg(ctx):
    def fn(doc, path):
        normalize_rotation(doc)
        for p in parse_pages(ctx["pages"], doc.page_count):
            doc[p].draw_rect(doc[p].rect, color=None, fill=rgb(ctx["color"]), overlay=False)
    simple_page_tool(ctx, "_mau_nen", fn)


@tool(id="greyscale", group="stamp", name="Chuyển đen trắng", icon=0xE7B3, keywords="grayscale greyscale den trang",
      desc="Chuyển chữ, hình vẽ và ảnh sang thang xám (giữ chữ chọn được).", opts=[PASSWORD_OPT])
def t_grey(ctx):
    simple_page_tool(ctx, "_den_trang", lambda doc, p: doc.recolor(1))


def _raster_doc(src, dpi, fn=None, quality=80, gray=False):
    """Chuyển từng trang thành ảnh (fn(PIL image) -> PIL image để xử lý thêm)."""
    out = pymupdf.open()
    for pg in src:
        pix = pg.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY if gray else pymupdf.csRGB, alpha=False)
        if fn:
            im = fn(pix.pil_image())
            buf = io.BytesIO()
            im.save(buf, "JPEG", quality=quality)
            data = buf.getvalue()
        else:
            data = pix.tobytes("jpg", jpg_quality=quality)
        np_ = out.new_page(width=pg.rect.width, height=pg.rect.height)
        np_.insert_image(np_.rect, stream=data)
    return out


@tool(id="invert", group="stamp", name="Đảo màu (nền tối)", icon=0xE793, keywords="invert dark mode",
      desc="Đảo màu trang (nền đen chữ trắng) để đọc ban đêm. Trang sẽ thành ảnh.",
      opts=[Opt("dpi", "Độ phân giải", "int", 150, min=72, max=300, suffix=" dpi"), PASSWORD_OPT])
def t_invert(ctx):
    from PIL import ImageOps

    def one(path, i):
        src = open_doc(path, ctx["password"])
        save_doc(_raster_doc(src, int(ctx["dpi"]), lambda im: ImageOps.invert(im.convert("RGB"))), ctx.out(path, "_dao_mau"))
    ctx.each(one)


@tool(id="scan_effect", group="stamp", name="Tạo hiệu ứng bản scan", icon=0xE8B9, keywords="scanner effect scan",
      desc="Làm file trông như được scan từ giấy (xám, hơi nghiêng, có hạt).",
      opts=[Opt("dpi", "Độ phân giải", "int", 150, min=72, max=300, suffix=" dpi"),
            Opt("tilt", "Độ nghiêng tối đa", "float", 0.6, min=0, max=3, step=0.1, suffix=" °"),
            Opt("color", "Giữ màu", "bool", False), PASSWORD_OPT])
def t_scan(ctx):
    import random
    from PIL import Image, ImageEnhance, ImageFilter

    def eff(im):
        im = im.convert("RGB" if ctx["color"] else "L")
        a = random.uniform(-float(ctx["tilt"]), float(ctx["tilt"]))
        im = im.rotate(a, resample=Image.BICUBIC, fillcolor=255 if im.mode == "L" else (255, 255, 255))
        noise = Image.effect_noise(im.size, 18).convert(im.mode)
        im = Image.blend(im, noise, 0.06)
        im = ImageEnhance.Contrast(im).enhance(1.15).filter(ImageFilter.GaussianBlur(0.4))
        return im

    def one(path, i):
        src = open_doc(path, ctx["password"])
        save_doc(_raster_doc(src, int(ctx["dpi"]), eff, quality=70), ctx.out(path, "_scan"))
    ctx.each(one)


# ===========================================================================
#  3. CHUYỂN ĐỔI SANG PDF
# ===========================================================================
@tool(id="images_to_pdf", group="to_pdf", name="Ảnh → PDF", icon=0xEB9F, popular=True, accept=IMAGE_EXT,
      keywords="jpg png anh image tiff webp heic",
      desc="Chuyển ảnh JPG, PNG, TIFF, WEBP, BMP, PSD... thành PDF. Kéo thả để sắp thứ tự.",
      opts=[Opt("merge", "Gộp tất cả ảnh vào 1 file", "bool", True),
            Opt("size", "Khổ trang", "choice", "image", [("image", "Theo kích thước ảnh")] + PAPER_CHOICES),
            Opt("orient", "Hướng trang", "choice", "auto", [("auto", "Tự động theo ảnh"), ("portrait", "Dọc"),
                                                          ("landscape", "Ngang")]),
            Opt("margin", "Lề", "float", 0, min=0, max=50, suffix=" mm"),
            out_name_opt()])
def t_images(ctx):
    m = float(ctx["margin"]) * MM
    out = pymupdf.open()

    def add(path):
        src = open_doc(path)
        for sp in src:
            ir = sp.rect
            if ctx["size"] == "image":
                pr = pymupdf.Rect(0, 0, ir.width + 2 * m, ir.height + 2 * m)
            else:
                orient = ctx["orient"] if ctx["orient"] != "auto" else ("landscape" if ir.width > ir.height else "portrait")
                pr = paper(ctx["size"], orient)
            pg = out.new_page(width=pr.width, height=pr.height)
            pg.show_pdf_page(pr + (m, m, -m, -m), src, sp.number)
        src.close()

    if ctx["merge"]:
        n = len(ctx.files)
        for i, f in enumerate(ctx.files):
            ctx.progress(i / n, os.path.basename(f))
            try:
                add(f)
            except Exception as e:
                ctx.log(f"✘ {os.path.basename(f)}: {e}")
        if not out.page_count:
            raise ToolError("Không đọc được ảnh nào.")
        save_doc(out, ctx.out(ctx.files[0], "", name=ctx["out_name"] or "Anh_sang_PDF"))
    else:
        def one(path, i):
            nonlocal out
            out = pymupdf.open()
            add(path)
            save_doc(out, ctx.out(path, ""))
        ctx.each(one)


def _office_kind(path):
    ext = os.path.splitext(path)[1].lower()
    return "word" if ext in WORD_EXT else "excel" if ext in EXCEL_EXT else "ppt" if ext in PPT_EXT else None


@tool(id="office_to_pdf", group="to_pdf", name="Word / Excel / PowerPoint → PDF", icon=0xE8A5, popular=True,
      accept=WORD_EXT + EXCEL_EXT + PPT_EXT, keywords="word excel powerpoint docx xlsx pptx office",
      desc="Chuyển file Office sang PDF bằng Microsoft Office trên máy (giữ đúng định dạng nhất).",
      opts=[Opt("fit_width", "Excel: co vừa chiều ngang 1 trang giấy", "bool", True),
            Opt("landscape", "Excel: in trang ngang", "bool", False),
            Opt("all_sheets", "Excel: xuất tất cả sheet", "bool", True)],
      note="Cần Microsoft Office trên máy. File ODT/ODS/ODP, RTF, CSV cũng được.")
def t_office(ctx):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    apps = {}
    progids = {"word": "Word.Application", "excel": "Excel.Application", "ppt": "PowerPoint.Application"}

    def app(kind):
        if kind not in apps:
            try:
                a = win32com.client.DispatchEx(progids[kind])
            except Exception:
                raise ToolError(f"máy chưa cài Microsoft {dict(word='Word', excel='Excel', ppt='PowerPoint')[kind]}")
            try:
                if kind == "word":
                    a.Visible = False
                    a.DisplayAlerts = 0
                elif kind == "excel":
                    a.Visible = False
                    a.DisplayAlerts = False
                    a.AskToUpdateLinks = False
                else:
                    a.DisplayAlerts = 1
            except Exception:
                pass
            apps[kind] = a
        return apps[kind]

    def one(path, i):
        path = os.path.abspath(path)
        out = os.path.abspath(ctx.out(path, ""))
        kind = _office_kind(path)
        if kind == "word":
            d = app("word").Documents.Open(path, False, True, False, "__khong_co_mat_khau__")
            try:
                d.ExportAsFixedFormat(out, 17)
            finally:
                d.Close(0)
        elif kind == "excel":
            xl = app("excel")
            wb = xl.Workbooks.Open(path, 0, True, None, "__khong_co_mat_khau__")
            try:
                if ctx["fit_width"] or ctx["landscape"]:
                    try:
                        xl.PrintCommunication = False
                    except Exception:
                        pass
                    for ws in wb.Worksheets:
                        ps = ws.PageSetup
                        if ctx["fit_width"]:
                            ps.Zoom = False
                            ps.FitToPagesWide = 1
                            ps.FitToPagesTall = False
                        if ctx["landscape"]:
                            ps.Orientation = 2
                    try:
                        xl.PrintCommunication = True
                    except Exception:
                        pass
                target = wb if ctx["all_sheets"] else wb.ActiveSheet
                target.ExportAsFixedFormat(0, out)
            finally:
                wb.Close(False)
        elif kind == "ppt":
            pres = app("ppt").Presentations.Open(path, True, False, False)
            try:
                pres.SaveAs(out, 32)
            finally:
                pres.Close()
        else:
            raise ToolError("không phải file Office")
        if not os.path.exists(out):
            raise ToolError("Office không xuất được PDF")

    try:
        ctx.each(one)
    finally:
        for a in apps.values():
            try:
                a.Quit()
            except Exception:
                pass
        apps.clear()
        pythoncom.CoUninitialize()


@tool(id="ebook_to_pdf", group="to_pdf", name="Ebook / XPS / SVG → PDF", icon=0xE736, accept=EBOOK_EXT,
      keywords="epub mobi fb2 cbz xps svg ebook",
      desc="Chuyển sách điện tử EPUB, MOBI, FB2, truyện tranh CBZ, file XPS / SVG sang PDF.",
      opts=[Opt("size", "Khổ trang (sách)", "choice", "A5", PAPER_CHOICES),
            Opt("font", "Cỡ chữ (sách)", "float", 11, min=6, max=30)])
def t_ebook(ctx):
    def one(path, i):
        doc = pymupdf.open(path)
        if doc.is_reflowable:
            doc.layout(pymupdf.paper_rect(PAPERS[ctx["size"]]), fontsize=float(ctx["font"]))
        pdf = pymupdf.open("pdf", doc.convert_to_pdf())
        toc = doc.get_toc()
        if toc:
            try:
                pdf.set_toc(toc)
            except Exception:
                pass
        save_doc(pdf, ctx.out(path, ""))
    ctx.each(one)


DOC_CSS = """
body {font-family: '%(f)s'; font-size: %(s)gpt; line-height: 1.4; color: #111;}
h1 {font-size: 1.8em; margin: 0.6em 0 0.3em 0;} h2 {font-size: 1.45em; margin: 0.6em 0 0.3em 0;}
h3 {font-size: 1.2em; margin: 0.5em 0 0.2em 0;} p {margin: 0 0 0.5em 0;}
table {border-collapse: collapse; margin: 0.4em 0;} th, td {border: 1px solid #999; padding: 2px 5px; text-align: left;}
th {background-color: #e8eef7; font-weight: bold;}
pre, code {font-family: 'Consolas'; font-size: 0.9em;} pre {white-space: pre-wrap; background-color: #f5f5f5; padding: 4px;}
blockquote {margin-left: 1.5em; color: #555;} a {color: #0563c1;}
"""


def html_to_pdf(body, out_path, family="Arial", size=11, paper_name="A4", margin_mm=18, landscape=False,
                extra_dirs=(), extra_files=()):
    """Dàn trang HTML thành PDF nhiều trang (MuPDF Story), nhúng font tiếng Việt."""
    fams = [family, "Consolas", "Arial"]
    css = pdf_core.font_css(fams) + DOC_CSS % {"f": family, "s": size}
    arch = pdf_core._archive()
    for d in extra_dirs:
        if d and os.path.isdir(d):
            arch.add(d)
    for data, name in extra_files:
        arch.add(data, name)
    story = pymupdf.Story(body, user_css=css, archive=arch)
    box = paper(paper_name, "landscape" if landscape else "portrait")
    m = margin_mm * MM
    buf = io.BytesIO()
    writer = pymupdf.DocumentWriter(buf)
    more = True
    while more:
        dev = writer.begin_page(box)
        more, _ = story.place(box + (m, m, -m, -m))
        story.draw(dev)
        writer.end_page()
    writer.close()
    doc = pymupdf.open("pdf", buf.getvalue())
    save_doc(doc, out_path, fonts=True)
    return doc.page_count


def _read_text(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-16", "cp1258", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeError:
            continue
    return raw.decode("utf-8", "replace")


def _csv_html(text):
    import csv
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    rows = list(csv.reader(io.StringIO(text), dialect))
    if not rows:
        return "<p></p>"
    head = "".join(f"<th>{_html.escape(c)}</th>" for c in rows[0])
    body = "".join("<tr>" + "".join(f"<td>{_html.escape(c)}</td>" for c in r) + "</tr>" for r in rows[1:])
    return f"<table><tr>{head}</tr>{body}</table>"


@tool(id="text_to_pdf", group="to_pdf", name="Văn bản / Markdown / HTML → PDF", icon=0xE8A5, accept=TEXT_EXT,
      keywords="txt markdown md html csv json xml",
      desc="Chuyển file TXT, Markdown, HTML, CSV (thành bảng), JSON, XML sang PDF có định dạng.",
      opts=[Opt("font", "Font", "choice", "Arial", FONT_CHOICES), Opt("size", "Cỡ chữ", "float", 11, min=6, max=30),
            Opt("paper", "Khổ giấy", "choice", "A4", PAPER_CHOICES), Opt("landscape", "Trang ngang", "bool", False),
            Opt("margin", "Lề", "float", 18, min=0, max=50, suffix=" mm")])
def t_text(ctx):
    import markdown

    def one(path, i):
        ext = os.path.splitext(path)[1].lower()
        text = _read_text(path)
        if ext in (".md", ".markdown"):
            body = markdown.markdown(text, extensions=["tables", "fenced_code", "sane_lists", "nl2br"])
        elif ext in (".html", ".htm"):
            body = re.sub(r"(?is)<script.*?</script>", "", text)
        elif ext == ".csv":
            body = _csv_html(text)
        elif ext == ".json":
            try:
                text = json.dumps(json.loads(text), ensure_ascii=False, indent=2)
            except ValueError:
                pass
            body = f"<pre>{_html.escape(text)}</pre>"
        elif ext == ".xml":
            body = f"<pre>{_html.escape(text)}</pre>"
        else:
            body = "".join(f"<p>{_html.escape(ln) or '&nbsp;'}</p>" for ln in text.splitlines()) \
                if len(text) < 3_000_000 else f"<pre>{_html.escape(text)}</pre>"
            body = f'<div style="white-space: pre-wrap">{body}</div>'
        n = html_to_pdf(body, ctx.out(path, ""), ctx["font"], float(ctx["size"]), ctx["paper"], float(ctx["margin"]),
                        ctx["landscape"], extra_dirs=[os.path.dirname(path)] if ext in (".html", ".htm", ".md") else ())
        ctx.log(f"{os.path.basename(path)}: {n} trang")
    ctx.each(one)


@tool(id="email_to_pdf", group="to_pdf", name="Email (.eml) → PDF", icon=0xE715, accept=MAIL_EXT, keywords="email eml mail",
      desc="Chuyển email đã lưu (.eml) sang PDF: người gửi, người nhận, ngày, nội dung, ảnh trong thư.",
      opts=[Opt("attach", "Nhúng file đính kèm vào PDF", "bool", True)],
      note="Email trong Outlook: kéo thư ra Desktop sẽ được file .msg; dùng 'Lưu thành' > .eml hoặc Mail Toolkit để xuất.")
def t_email(ctx):
    from email import policy
    from email.parser import BytesParser

    def one(path, i):
        with open(path, "rb") as fh:
            msg = BytesParser(policy=policy.default).parse(fh)
        files, k = [], 0
        cid_map = {}
        for part in msg.walk():
            cid = part.get("Content-ID")
            if cid and part.get_content_maintype() == "image":
                k += 1
                name = f"cid_{k}.{part.get_content_subtype()}"
                files.append((part.get_payload(decode=True), name))
                cid_map[cid.strip("<>")] = name
        body_part = msg.get_body(preferencelist=("html", "plain"))
        if body_part is None:
            content = "<p>(Không có nội dung)</p>"
        elif body_part.get_content_subtype() == "html":
            content = body_part.get_content()
            content = re.sub(r"(?is)<(script|style|head)\b.*?</\1>", "", content)
            for cid, name in cid_map.items():
                content = content.replace("cid:" + cid, name)
            content = re.sub(r'(?is)<img\b[^>]*src=["\'](?:https?:|cid:)[^>]*>', "", content)
        else:
            content = f'<pre>{_html.escape(body_part.get_content())}</pre>'
        atts = [(p.get_filename() or "dinh_kem", p.get_payload(decode=True)) for p in msg.iter_attachments()]
        rows = "".join(f"<tr><th>{k}</th><td>{_html.escape(str(msg.get(h, '')))}</td></tr>"
                       for k, h in (("Từ", "From"), ("Đến", "To"), ("Cc", "Cc"), ("Ngày", "Date"), ("Tiêu đề", "Subject"))
                       if msg.get(h))
        if atts:
            rows += f"<tr><th>Đính kèm</th><td>{_html.escape(', '.join(a[0] for a in atts))}</td></tr>"
        body = f"<table style='width:100%'>{rows}</table><hr/>{content}"
        out = ctx.out(path, "")
        try:
            html_to_pdf(body, out, "Arial", 10.5, extra_files=files)
        except Exception:
            text = body_part.get_content() if body_part is not None else ""
            text = re.sub(r"<[^>]+>", " ", text) if "<" in text else text
            html_to_pdf(f"<table>{rows}</table><hr/><pre>{_html.escape(text)}</pre>", out, "Arial", 10.5)
        if ctx["attach"] and atts:
            d = pymupdf.open(out)
            for name, data in atts:
                if data:
                    try:
                        d.embfile_add(name, data, filename=name)
                    except Exception:
                        pass
            d.saveIncr()
            d.close()
    ctx.each(one)


# ===========================================================================
#  4. CHUYỂN ĐỔI TỪ PDF
# ===========================================================================
@tool(id="pdf_to_images", group="from_pdf", name="PDF → Ảnh", icon=0xE91B, popular=True,
      keywords="jpg png image anh tiff webp",
      desc="Xuất từng trang thành ảnh JPG, PNG, WEBP, TIFF, BMP.",
      opts=[Opt("fmt", "Định dạng", "choice", "jpg", [("jpg", "JPG"), ("png", "PNG"), ("webp", "WEBP"),
                                                     ("tiff", "TIFF"), ("bmp", "BMP")]),
            Opt("dpi", "Độ phân giải", "int", 150, min=36, max=600, suffix=" dpi",
                tip="96 = xem màn hình, 150 = thường dùng, 300 = in ấn"),
            Opt("quality", "Chất lượng JPG / WEBP", "int", 85, min=10, max=100, show_if=("fmt", ("jpg", "webp"))),
            Opt("zip", "Đóng gói thành 1 file ZIP", "bool", False), pages_opt(), PASSWORD_OPT])
def t_to_images(ctx):
    fmt = ctx["fmt"]
    q = int(ctx["quality"])

    def img_bytes(pix):
        if fmt == "png":
            return pix.tobytes("png")
        if fmt == "jpg":
            return pix.tobytes("jpg", jpg_quality=q)
        buf = io.BytesIO()
        pix.pil_save(buf, format={"webp": "WEBP", "tiff": "TIFF", "bmp": "BMP"}[fmt], quality=q)
        return buf.getvalue()

    def one(path, i):
        doc = open_doc(path, ctx["password"])
        pages = parse_pages(ctx["pages"], doc.page_count)
        stem = os.path.splitext(os.path.basename(path))[0]
        if ctx["zip"]:
            zf = zipfile.ZipFile(ctx.out(path, "_anh", ".zip"), "w", zipfile.ZIP_STORED)
        else:
            folder = ctx.out(path, "_anh", folder=True)
        for k, p in enumerate(pages):
            ctx.progress((i + k / len(pages)) / len(ctx.files), f"{os.path.basename(path)} - trang {p + 1}")
            pix = doc[p].get_pixmap(dpi=int(ctx["dpi"]), alpha=False)
            name = f"{stem}_trang_{p + 1:03d}.{fmt}"
            data = img_bytes(pix)
            if ctx["zip"]:
                zf.writestr(name, data)
            else:
                with open(os.path.join(folder, name), "wb") as fh:
                    fh.write(data)
        if ctx["zip"]:
            zf.close()
        ctx.log(f"{os.path.basename(path)}: {len(pages)} ảnh")
    ctx.each(one)


@tool(id="pdf_to_word", group="from_pdf", name="PDF → Word", icon=0xE8A5, popular=True, keywords="docx word",
      desc="Chuyển PDF sang Word (.docx) để sửa: giữ đoạn văn, bảng, ảnh, font.",
      opts=[pages_opt("Trang cần chuyển"), PASSWORD_OPT],
      note="PDF scan (chỉ có ảnh) sẽ ra Word chứa ảnh. File phức tạp có thể lệch bố cục.")
def t_to_word(ctx):
    import logging
    import cv2_lite
    cv2_lite.install()                  # bản exe không kèm OpenCV: pdf2docx dùng bản thay thế nhỏ
    from pdf2docx import Converter
    logging.getLogger().setLevel(logging.ERROR)

    def one(path, i):
        doc = open_doc(path, ctx["password"])
        pages = parse_pages(ctx["pages"], doc.page_count)
        src = path
        tmp = None
        if doc.needs_pass or doc.is_encrypted or os.path.splitext(path)[1].lower() != ".pdf":
            import tempfile
            fd, tmp = tempfile.mkstemp(suffix=".pdf")
            os.close(fd)
            doc.save(tmp)
            src = tmp
        doc.close()
        try:
            cv = Converter(src)
            cv.convert(ctx.out(path, "", ".docx"), pages=pages)
            cv.close()
        finally:
            if tmp:
                os.remove(tmp)
    ctx.each(one)


def _page_without_text(doc, pno):
    """Bản sao 1 trang đã xoá chữ ngang (giữ ảnh, hình vẽ, nền, chữ xoay dọc) - làm nền cho slide có chữ sửa được."""
    tmp = pymupdf.open()
    tmp.insert_pdf(doc, from_page=pno, to_page=pno)
    pg = tmp[0]
    for b in pg.get_text("dict")["blocks"]:
        for ln in b.get("lines", []):
            if abs(ln["dir"][1]) > 0.1:
                continue
            for sp in ln["spans"]:
                if sp["text"].strip():
                    pg.add_redact_annot(pymupdf.Rect(sp["bbox"]))
    pg.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE, graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
    return tmp


def _ppt_font(name):
    """Tên font trong PDF (vd ABCDEF+Arial-BoldMT) -> tên font Office."""
    n = name.split("+")[-1]
    n = re.sub(r"(?i)[-,](bold|italic|oblique|regular|black|light|medium|semibold).*$", "", n)
    n = re.sub(r"(?i)(MT|PS|PSMT)$", "", n)
    known = {"helvetica": "Arial", "helv": "Arial", "arialmt": "Arial", "times": "Times New Roman",
             "timesnewroman": "Times New Roman", "timesnewromanps": "Times New Roman", "courier": "Courier New",
             "couriernew": "Courier New", "tiro": "Times New Roman", "cour": "Courier New",
             "nimbussans": "Arial", "nimbussansl": "Arial", "nimbusroman": "Times New Roman",
             "nimbusromno9l": "Times New Roman", "nimbusmono": "Courier New", "nimbusmonops": "Courier New",
             "liberationsans": "Arial", "liberationserif": "Times New Roman", "liberationmono": "Courier New",
             "dejavusans": "Arial", "notosans": "Arial", "notoserif": "Times New Roman"}
    key = re.sub(r"[\s_-]", "", n).lower()
    return known.get(key, n.replace("_", " ") or "Arial")


@tool(id="pdf_to_ppt", group="from_pdf", name="PDF → PowerPoint", icon=0xE8A5, keywords="pptx powerpoint slide trinh chieu",
      desc="Chuyển mỗi trang PDF thành 1 slide PowerPoint (.pptx): chữ sửa được trên nền ảnh, hoặc giữ nguyên dạng ảnh.",
      opts=[Opt("mode", "Kiểu slide", "choice", "editable",
                [("editable", "Chữ sửa được (nền = hình ảnh / hình vẽ của trang)"),
                 ("image", "Mỗi trang là 1 ảnh (giống bản gốc 100%, không sửa chữ)")]),
            Opt("dpi", "Độ nét ảnh", "int", 150, min=72, max=300, suffix=" dpi"),
            pages_opt("Trang cần chuyển"), PASSWORD_OPT],
      note="PDF scan: hãy OCR trước để có chữ sửa được. Font không có trên máy sẽ được PowerPoint thay bằng font tương tự.")
def t_to_ppt(ctx):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
    from pptx.util import Emu, Pt
    EMU = 12700                                       # 1 pt = 12700 EMU

    def one(path, i):
        doc = open_doc(path, ctx["password"])
        pages = parse_pages(ctx["pages"], doc.page_count)
        prs = Presentation()
        first = doc[pages[0]].rect
        prs.slide_width, prs.slide_height = Emu(int(first.width * EMU)), Emu(int(first.height * EMU))
        blank = prs.slide_layouts[6]
        zoom = ctx["dpi"] / 72
        for k, pno in enumerate(pages):
            ctx.progress((i + k / len(pages)) / len(ctx.files), f"{os.path.basename(path)} - trang {pno + 1}")
            page = doc[pno]
            slide = prs.slides.add_slide(blank)
            # trang khác khổ slide đầu: co vừa, căn giữa
            s = min(prs.slide_width / (page.rect.width * EMU), prs.slide_height / (page.rect.height * EMU))
            ox = (prs.slide_width - page.rect.width * EMU * s) / 2
            oy = (prs.slide_height - page.rect.height * EMU * s) / 2

            def emu(v, off=0.0):
                return Emu(int(off + v * EMU * s))
            # trang đặt xoay (/Rotate): hướng chữ khác hướng slide -> giữ dạng ảnh cho đúng bố cục
            editable = ctx["mode"] == "editable" and page.rotation == 0 and page.get_text("text").strip()
            src = _page_without_text(doc, pno) if editable else None
            pix = (src[0] if src else page).get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), alpha=False)
            if src:
                src.close()
            slide.shapes.add_picture(io.BytesIO(pix.tobytes("png" if editable else "jpg")), emu(0, ox), emu(0, oy),
                                     emu(page.rect.width), emu(page.rect.height))
            if not editable:
                continue
            m = page.rotation_matrix
            for b in page.get_text("dict")["blocks"]:
                for ln in b.get("lines", []):
                    spans = [sp for sp in ln["spans"] if sp["text"].strip()]
                    if not spans or abs(ln["dir"][1]) > 0.1:      # chữ xoay dọc: giữ trong ảnh nền
                        continue
                    r = pymupdf.Rect(ln["bbox"]) * m
                    r.normalize()
                    tb = slide.shapes.add_textbox(emu(r.x0, ox), emu(r.y0, oy), emu(r.width * 1.06 + 2),
                                                  emu(r.height))
                    tf = tb.text_frame
                    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
                    tf.word_wrap = False
                    tf.auto_size = MSO_AUTO_SIZE.NONE
                    tf.vertical_anchor = MSO_ANCHOR.TOP
                    para = tf.paragraphs[0]
                    for sp in ln["spans"]:
                        run = para.add_run()
                        run.text = sp["text"]
                        f = run.font
                        f.size = Pt(max(1.0, round(sp["size"] * s * 2) / 2))
                        f.name = _ppt_font(sp["font"])
                        f.bold = bool(sp["flags"] & 16) or "bold" in sp["font"].lower()
                        f.italic = bool(sp["flags"] & 2) or "italic" in sp["font"].lower()
                        c = sp["color"]
                        f.color.rgb = RGBColor((c >> 16) & 255, (c >> 8) & 255, c & 255)
        doc.close()
        prs.save(ctx.out(path, "", ".pptx"))
        ctx.log(f"{os.path.basename(path)}: {len(pages)} slide")
    ctx.each(one)


@tool(id="pdf_to_excel", group="from_pdf", name="PDF → Excel", icon=0xE80A, keywords="xlsx excel table bang csv",
      desc="Lấy các bảng trong PDF ra Excel (mỗi bảng 1 sheet) hoặc CSV.",
      opts=[Opt("fmt", "Định dạng", "choice", "xlsx", [("xlsx", "Excel (.xlsx)"), ("csv", "CSV")]),
            Opt("fallback", "Không có bảng thì lấy toàn bộ chữ (mỗi dòng 1 hàng)", "bool", True),
            pages_opt(), PASSWORD_OPT])
def t_to_excel(ctx):
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        tables = []
        for p in parse_pages(ctx["pages"], doc.page_count):
            for k, t in enumerate(doc[p].find_tables().tables):
                rows = [[(c or "").replace("\n", " ").strip() for c in r] for r in t.extract()]
                if rows:
                    tables.append((f"Trang {p + 1} - Bang {k + 1}", rows))
        if not tables:
            if not ctx["fallback"]:
                raise ToolError("không tìm thấy bảng")
            rows = []
            for p in parse_pages(ctx["pages"], doc.page_count):
                for b in doc[p].get_text("blocks"):
                    for ln in b[4].splitlines():
                        if ln.strip():
                            rows.append([ln.strip()])
            tables = [("Noi dung", rows)]
            ctx.log(f"{os.path.basename(path)}: không thấy bảng - xuất toàn bộ chữ")
        if ctx["fmt"] == "csv":
            import csv
            with open(ctx.out(path, "", ".csv"), "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                for name, rows in tables:
                    w.writerow([f"# {name}"])
                    w.writerows(rows)
                    w.writerow([])
        else:
            from openpyxl import Workbook
            from openpyxl.styles import Font
            wb = Workbook()
            wb.remove(wb.active)
            for name, rows in tables:
                ws = wb.create_sheet(name[:31])
                for r in rows:
                    ws.append([_num(c) for c in r])
                for c in ws[1]:
                    c.font = Font(bold=True)
                for col in ws.columns:
                    ws.column_dimensions[col[0].column_letter].width = min(60, max(8, max(len(str(c.value or "")) for c in col) + 2))
            wb.save(ctx.out(path, "", ".xlsx"))
        ctx.log(f"{os.path.basename(path)}: {len(tables)} bảng")
    ctx.each(one)


def _num(s):
    """'1.234.567' / '1,234.5' -> số nếu chắc chắn là số, ngược lại giữ chữ."""
    t = (s or "").strip()
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t):
        return int(t.replace(".", ""))
    if re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", t):
        return float(t.replace(",", "")) if "." in t else int(t.replace(",", ""))
    if re.fullmatch(r"-?\d+", t) and not (len(t) > 1 and t.startswith("0")):
        return int(t)
    if re.fullmatch(r"-?\d+\.\d+", t):
        return float(t)
    return s


def _page_markdown(page):
    """Chữ của trang -> Markdown đơn giản (tiêu đề theo cỡ chữ, đậm / nghiêng, bảng)."""
    tabs = page.find_tables().tables
    tab_rects = [pymupdf.Rect(t.bbox) for t in tabs]
    d = page.get_text("dict", flags=pymupdf.TEXT_PRESERVE_WHITESPACE)
    sizes = {}
    for b in d["blocks"]:
        for ln in b.get("lines", []):
            for s in ln["spans"]:
                sizes[round(s["size"])] = sizes.get(round(s["size"]), 0) + len(s["text"])
    body = max(sizes, key=sizes.get) if sizes else 11
    items = [(pymupdf.Rect(t.bbox).y0, t.to_markdown()) for t in tabs]
    for b in d["blocks"]:
        if b.get("type") != 0 or any(pymupdf.Rect(b["bbox"]).intersects(r) for r in tab_rects):
            continue
        lines = []
        for ln in b["lines"]:
            parts = []
            for s in ln["spans"]:
                t = s["text"]
                if not t.strip():
                    parts.append(t)
                    continue
                bold, ital = s["flags"] & 16, s["flags"] & 2
                core = t.strip()
                if bold:
                    core = f"**{core}**"
                if ital:
                    core = f"*{core}*"
                parts.append(t[:len(t) - len(t.lstrip())] + core + t[len(t.rstrip()):])
            lines.append("".join(parts).strip())
        text = " ".join(l for l in lines if l)
        if not text:
            continue
        size = max(s["size"] for ln in b["lines"] for s in ln["spans"])
        plain = text.replace("**", "")
        if size >= body * 1.6:
            text = "# " + plain
        elif size >= body * 1.3:
            text = "## " + plain
        elif size >= body * 1.12:
            text = "### " + plain
        elif re.match(r"^[•●▪◦\-–]\s*", text):
            text = "- " + re.sub(r"^[•●▪◦\-–]\s*", "", text)
        items.append((b["bbox"][1], text))
    items.sort(key=lambda x: x[0])
    return "\n\n".join(t for _, t in items)


@tool(id="pdf_to_text", group="from_pdf", name="PDF → Văn bản / Markdown", icon=0xE8D2, keywords="txt markdown md html json ai",
      desc="Lấy chữ trong PDF ra TXT, Markdown (dùng cho ChatGPT / AI), HTML hoặc JSON.",
      opts=[Opt("fmt", "Định dạng", "choice", "txt", [("txt", "Text (.txt)"), ("md", "Markdown (.md)"),
                                                     ("html", "HTML (.html)"), ("json", "JSON (.json)")]),
            Opt("page_marks", "Ghi dấu ngắt trang", "bool", True), pages_opt(), PASSWORD_OPT],
      note="PDF scan không có lớp chữ sẽ ra file rỗng (cần OCR).")
def t_to_text(ctx):
    fmt = ctx["fmt"]

    def one(path, i):
        doc = open_doc(path, ctx["password"])
        pages = parse_pages(ctx["pages"], doc.page_count)
        if fmt == "json":
            data = [json.loads(doc[p].get_text("json")) for p in pages]
            for d in data:
                for b in d.get("blocks", []):
                    b.pop("image", None)
            text = json.dumps({"file": os.path.basename(path), "pages": data}, ensure_ascii=False, indent=1)
        elif fmt == "html":
            text = "<!DOCTYPE html><html><head><meta charset='utf-8'></head><body>" + \
                   "\n<hr/>\n".join(doc[p].get_text("xhtml") for p in pages) + "</body></html>"
        else:
            chunks = []
            for p in pages:
                t = _page_markdown(doc[p]) if fmt == "md" else doc[p].get_text("text")
                if ctx["page_marks"]:
                    t = (f"\n<!-- Trang {p + 1} -->\n\n" if fmt == "md" else f"\n===== Trang {p + 1} =====\n") + t
                chunks.append(t)
            text = "\n".join(chunks).strip() + "\n"
        with open(ctx.out(path, "", "." + fmt), "w", encoding="utf-8") as fh:
            fh.write(text)
        if not text.strip() or len(re.sub(r"\s|=+|Trang \d+|<!--.*?-->", "", text)) < 5:
            ctx.log(f"{os.path.basename(path)}: không có chữ - có thể là PDF scan")
    ctx.each(one)


@tool(id="pdf_to_svg", group="from_pdf", name="PDF → SVG", icon=0xE790, keywords="svg vector",
      desc="Xuất từng trang thành ảnh vector SVG (phóng to không vỡ).",
      opts=[Opt("text_path", "Chuyển chữ thành nét vẽ", "bool", False), pages_opt(), PASSWORD_OPT])
def t_to_svg(ctx):
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        folder = ctx.out(path, "_svg", folder=True)
        stem = os.path.splitext(os.path.basename(path))[0]
        for p in parse_pages(ctx["pages"], doc.page_count):
            with open(os.path.join(folder, f"{stem}_trang_{p + 1:03d}.svg"), "w", encoding="utf-8") as fh:
                fh.write(doc[p].get_svg_image(text_as_path=bool(ctx["text_path"])))
    ctx.each(one)


@tool(id="extract_images", group="from_pdf", name="Lấy ảnh gốc trong PDF", icon=0xE8B9, keywords="extract image",
      desc="Lưu nguyên gốc các ảnh nhúng trong PDF ra thư mục.",
      opts=[Opt("min_px", "Bỏ qua ảnh nhỏ hơn", "int", 50, min=0, max=5000, suffix=" px"), pages_opt(), PASSWORD_OPT])
def t_extract_images(ctx):
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        folder = ctx.out(path, "_anh_goc", folder=True)
        stem = os.path.splitext(os.path.basename(path))[0]
        seen, n = set(), 0
        for p in parse_pages(ctx["pages"], doc.page_count):
            for img in doc[p].get_images(full=True):
                xref = img[0]
                if xref in seen:
                    continue
                seen.add(xref)
                try:
                    d = doc.extract_image(xref)
                except Exception:
                    continue
                if not d or min(d["width"], d["height"]) < int(ctx["min_px"]):
                    continue
                n += 1
                with open(os.path.join(folder, f"{stem}_trang{p + 1}_{n:03d}.{d['ext']}"), "wb") as fh:
                    fh.write(d["image"])
        ctx.log(f"{os.path.basename(path)}: {n} ảnh")
    ctx.each(one)


# ===========================================================================
#  5. BẢO MẬT, FORM & TỐI ƯU
# ===========================================================================
@tool(id="protect", group="secure", name="Đặt mật khẩu / phân quyền", icon=0xE72E, popular=True,
      keywords="password protect encrypt mat khau khoa quyen",
      desc="Đặt mật khẩu mở file (mã hoá AES-256) và giới hạn in, sao chép, chỉnh sửa.",
      opts=[Opt("user_pw", "Mật khẩu mở file", "password", "", tip="Để trống nếu chỉ muốn giới hạn quyền"),
            Opt("owner_pw", "Mật khẩu chủ (đổi quyền)", "password", "",
                tip="Dùng để gỡ giới hạn sau này. Để trống = tự sinh ngẫu nhiên"),
            Opt("p_print", "Cho phép in", "bool", True), Opt("p_copy", "Cho phép sao chép chữ / ảnh", "bool", False),
            Opt("p_modify", "Cho phép chỉnh sửa", "bool", False), Opt("p_annot", "Cho phép ghi chú / điền form", "bool", True),
            PASSWORD_OPT])
def t_protect(ctx):
    _protect(ctx, ctx["user_pw"] or "", "_bao_mat")


def _protect(ctx, user_pw, suffix):
    perm = pymupdf.PDF_PERM_ACCESSIBILITY
    if ctx["p_print"]:
        perm |= pymupdf.PDF_PERM_PRINT | pymupdf.PDF_PERM_PRINT_HQ
    if ctx["p_copy"]:
        perm |= pymupdf.PDF_PERM_COPY
    if ctx["p_modify"]:
        perm |= pymupdf.PDF_PERM_MODIFY | pymupdf.PDF_PERM_ASSEMBLE
    if ctx["p_annot"]:
        perm |= pymupdf.PDF_PERM_ANNOTATE | pymupdf.PDF_PERM_FORM
    owner = ctx["owner_pw"] or secrets.token_urlsafe(16)
    if not user_pw and all(ctx[k] for k in ("p_print", "p_copy", "p_modify", "p_annot")):
        raise ToolError("Chưa đặt mật khẩu mở file và cũng không giới hạn quyền nào." if suffix == "_bao_mat"
                        else "Hãy bỏ chọn ít nhất một quyền cần chặn (in, sao chép, chỉnh sửa...).")

    def one(path, i):
        doc = open_doc(path, ctx["password"])
        doc.save(ctx.out(path, suffix), garbage=3, deflate=True, encryption=pymupdf.PDF_ENCRYPT_AES_256,
                 owner_pw=owner, user_pw=user_pw, permissions=perm)
    ctx.each(one)
    if not ctx["owner_pw"]:
        ctx.log("Mật khẩu chủ được tạo ngẫu nhiên (không lưu lại).")
    if not user_pw:
        ctx.log("Người nhận vẫn mở xem được bình thường; Adobe Reader, Foxit, Edge, Chrome... sẽ chặn các quyền "
                "đã bỏ chọn.")


@tool(id="restrict", group="secure", name="Chặn in / sao chép", icon=0xE8D7, popular=True,
      keywords="restrict permission no print no copy chan in chan sao chep khong cho in bao ve quyen",
      desc="Không cần mật khẩu: người nhận vẫn mở xem bình thường nhưng không in, không sao chép chữ / ảnh, không "
           "sửa được (tuỳ chọn).",
      opts=[Opt("p_print", "Cho phép in", "bool", False), Opt("p_copy", "Cho phép sao chép chữ / ảnh", "bool", False),
            Opt("p_modify", "Cho phép chỉnh sửa", "bool", False),
            Opt("p_annot", "Cho phép ghi chú / điền form", "bool", False),
            Opt("owner_pw", "Mật khẩu chủ (đổi quyền)", "password", "",
                tip="Không bắt buộc - dùng để gỡ giới hạn sau này. Để trống = tự sinh ngẫu nhiên"),
            PASSWORD_OPT])
def t_restrict(ctx):
    _protect(ctx, "", "_chan_in")


@tool(id="form_detect", group="secure", name="Tự tạo ô form từ chỗ trống", icon=0xE8A5,
      keywords="form field detect auto fillable tao o dien tu dong nhan dien bieu mau",
      desc="Tự nhận diện chỗ cần điền (dòng gạch dưới / chấm chấm, đường kẻ trống, ô vuông, khung trống) và tạo ô form "
           "để người nhận gõ trực tiếp vào PDF.",
      opts=[pages_opt(), PASSWORD_OPT])
def t_form_detect(ctx):
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        pd = pdf_core.PdfDocument()
        pd.doc = doc
        pages = parse_pages(ctx["pages"], doc.page_count) if ctx["pages"] else None
        counts = pd.detect_fields(pages)
        if not counts:
            ctx.log(f"{os.path.basename(path)}: không thấy chỗ trống để tạo ô form")
            return
        doc.save(ctx.out(path, "_form"), garbage=3, deflate=True)
        ctx.log(f"{os.path.basename(path)}: tạo {sum(counts.values())} ô "
                f"({counts.get('text', 0) + counts.get('multiline', 0)} ô chữ, {counts.get('check', 0)} ô đánh dấu)")
    ctx.each(one)


@tool(id="redact", group="secure", name="Xoá vĩnh viễn thông tin (redact)", icon=0xE8C6, popular=True,
      keywords="redact redaction che xoa thong tin nhay cam bi mat so dien thoai email cccd cmnd tai khoan an danh",
      desc="Tìm và xoá hẳn chữ / số nhạy cảm (tên, số điện thoại, email, CCCD, số tài khoản...) khỏi file - không khôi "
           "phục, không copy hay tìm lại được. Chỗ bị xoá tô đen (hoặc trắng).",
      opts=[Opt("terms", "Chữ cần xoá (mỗi dòng 1 cụm)", "textarea", "",
                placeholder="vd:\nNguyễn Văn A\nCông ty ABC\n0123456789"),
            Opt("p_email", "Mọi địa chỉ email", "bool", False),
            Opt("p_phone", "Mọi số điện thoại", "bool", False),
            Opt("p_id", "Mọi số CCCD / CMND (9 hoặc 12 số)", "bool", False),
            Opt("p_account", "Mọi số tài khoản / dãy số dài (10-19 số)", "bool", False),
            Opt("whole", "Chỉ khớp nguyên từ", "bool", False, tip="Bỏ chọn: 'An' khớp cả trong 'Anh', 'Toàn'..."),
            Opt("fill", "Tô chỗ đã xoá", "choice", "black", [("black", "Màu đen"), ("white", "Màu trắng")]),
            Opt("scrub", "Xoá cả dữ liệu ẩn (tác giả, metadata, chữ ẩn, file đính kèm, JavaScript)", "bool", True),
            PASSWORD_OPT])
def t_redact(ctx):
    terms = [t for t in (ctx["terms"] or "").splitlines() if t.strip()]
    pats = [pdf_core.REDACT_PRESETS[k][1] for k in ("email", "phone", "id", "account") if ctx["p_" + k]]
    if not terms and not pats:
        raise ToolError("Nhập chữ cần xoá hoặc chọn ít nhất một loại thông tin.")
    fill = pdf_core.redact_fill(ctx["fill"])

    def one(path, i):
        doc = open_doc(path, ctx["password"])
        n = 0
        for page in doc:
            for r in pdf_core.find_redact_rects(page, terms, pats, ctx["whole"]):
                page.add_redact_annot(r, fill=fill)
                n += 1
        if not n:
            ctx.log(f"{os.path.basename(path)}: không tìm thấy chữ cần xoá (file scan cần OCR trước)")
            return
        pdf_core.apply_redactions_doc(doc)
        if ctx["scrub"]:
            doc.scrub(metadata=True, xml_metadata=True, attached_files=True, embedded_files=True, javascript=True,
                      remove_links=False, reset_fields=False, reset_responses=True, hidden_text=True,
                      redactions=False, clean_pages=True)
        doc.save(ctx.out(path, "_da_xoa_thong_tin"), garbage=4, deflate=True)
        ctx.log(f"{os.path.basename(path)}: đã xoá {n} chỗ")
    ctx.each(one)


@tool(id="unlock", group="secure", name="Mở khoá / gỡ giới hạn", icon=0xE785, keywords="unlock decrypt remove password",
      desc="Bỏ mật khẩu mở file (cần biết mật khẩu) và gỡ giới hạn in / sao chép / sửa.",
      opts=[Opt("password", "Mật khẩu hiện tại", "password", "", tip="Không cần nếu file chỉ bị giới hạn quyền")])
def t_unlock(ctx):
    def one(path, i):
        doc = pymupdf.open(path)
        if not doc.is_encrypted and not doc.needs_pass:
            ctx.log(f"{os.path.basename(path)}: file không có mật khẩu / giới hạn")
        if doc.needs_pass and not doc.authenticate(ctx["password"] or ""):
            raise ToolError("sai mật khẩu")
        doc.save(ctx.out(path, "_mo_khoa"), garbage=3, deflate=True, encryption=pymupdf.PDF_ENCRYPT_NONE)
    ctx.each(one)


META_KEYS = [("title", "Tiêu đề"), ("author", "Tác giả"), ("subject", "Chủ đề"), ("keywords", "Từ khoá"),
             ("creator", "Ứng dụng tạo"), ("producer", "Phần mềm PDF")]


def _meta_prefill(key):
    def f(path):
        try:
            d = pymupdf.open(path)
            v = (d.metadata or {}).get(key) or ""
            d.close()
            return v
        except Exception:
            return ""
    return f


@tool(id="metadata", group="secure", name="Xem / sửa thông tin file", icon=0xE946, keywords="metadata properties author title",
      desc="Xem và sửa Tiêu đề, Tác giả, Chủ đề, Từ khoá... (nạp sẵn từ file đầu tiên).",
      opts=[Opt(k, lbl, "text", "", prefill=_meta_prefill(k)) for k, lbl in META_KEYS] + [PASSWORD_OPT])
def t_metadata(ctx):
    def fn(doc, path):
        m = dict(doc.metadata or {})
        for k, _ in META_KEYS:
            m[k] = ctx[k] or ""
        m["modDate"] = pymupdf.get_pdf_now()
        doc.set_metadata({k: v for k, v in m.items() if k in ("title", "author", "subject", "keywords", "creator",
                                                               "producer", "creationDate", "modDate")})
    simple_page_tool(ctx, "_thong_tin", fn)


@tool(id="sanitize", group="secure", name="Xoá dữ liệu ẩn", icon=0xE7EF, keywords="sanitize metadata javascript clean",
      desc="Xoá thông tin tác giả, JavaScript, file nhúng, chữ ẩn, liên kết... trước khi gửi ra ngoài.",
      opts=[Opt("metadata", "Thông tin file (tác giả, phần mềm, ngày tạo...)", "bool", True),
            Opt("javascript", "JavaScript", "bool", True), Opt("embedded", "File nhúng / đính kèm", "bool", True),
            Opt("hidden", "Chữ ẩn (không nhìn thấy)", "bool", True), Opt("links", "Liên kết", "bool", False),
            Opt("annots", "Ghi chú / bình luận", "bool", False), Opt("thumbs", "Ảnh thu nhỏ trang", "bool", True),
            PASSWORD_OPT])
def t_sanitize(ctx):
    def fn(doc, path):
        doc.scrub(metadata=ctx["metadata"], xml_metadata=ctx["metadata"], javascript=ctx["javascript"],
                  embedded_files=ctx["embedded"], attached_files=ctx["embedded"], hidden_text=ctx["hidden"],
                  remove_links=ctx["links"], thumbnails=ctx["thumbs"], reset_fields=False, reset_responses=False,
                  redactions=False, clean_pages=True)
        if ctx["annots"]:
            for pg in doc:
                for a in list(pg.annots() or []):
                    pg.delete_annot(a)
    simple_page_tool(ctx, "_sach", fn)


@tool(id="flatten", group="secure", name="Cố định ghi chú & form vào trang", icon=0xE8B0, keywords="flatten form annotation",
      desc="Biến ô form, ghi chú, chữ ký thành nội dung cố định của trang (không sửa được nữa).",
      opts=[Opt("annots", "Ghi chú / hình vẽ", "bool", True), Opt("widgets", "Ô form", "bool", True), PASSWORD_OPT])
def t_flatten(ctx):
    simple_page_tool(ctx, "_phang", lambda doc, p: doc.bake(annots=bool(ctx["annots"]), widgets=bool(ctx["widgets"])))


@tool(id="remove_annots", group="secure", name="Xoá ghi chú / tô sáng", icon=0xED62, keywords="remove annotation comment",
      desc="Xoá các ghi chú, tô sáng, hình vẽ, bình luận khỏi file.",
      opts=[Opt("keep_links", "Giữ liên kết", "bool", True), Opt("keep_forms", "Giữ ô form", "bool", True), pages_opt(),
            PASSWORD_OPT])
def t_remove_annots(ctx):
    def fn(doc, path):
        n = 0
        for p in parse_pages(ctx["pages"], doc.page_count):
            pg = doc[p]
            for a in list(pg.annots() or []):
                pg.delete_annot(a)
                n += 1
            if not ctx["keep_links"]:
                for l in pg.get_links():
                    pg.delete_link(l)
            if not ctx["keep_forms"]:
                for w in list(pg.widgets() or []):
                    pg.delete_widget(w)
        ctx.log(f"{os.path.basename(path)}: xoá {n} ghi chú")
    simple_page_tool(ctx, "_bo_ghi_chu", fn)


def _form_fields(path):
    """[(tên, loại, giá trị, [lựa chọn])] của file đầu tiên."""
    out = []
    try:
        d = pymupdf.open(path)
        seen = set()
        for pg in d:
            for w in pg.widgets() or []:
                if w.field_name in seen or w.field_type == pymupdf.PDF_WIDGET_TYPE_SIGNATURE:
                    continue
                seen.add(w.field_name)
                kind = {pymupdf.PDF_WIDGET_TYPE_CHECKBOX: "check", pymupdf.PDF_WIDGET_TYPE_RADIOBUTTON: "radio",
                        pymupdf.PDF_WIDGET_TYPE_COMBOBOX: "choice", pymupdf.PDF_WIDGET_TYPE_LISTBOX: "choice"
                        }.get(w.field_type, "text")
                choices = [c if isinstance(c, str) else c[0] for c in (w.choice_values or [])]
                if kind == "radio":
                    choices = sorted({s for x in pg.widgets() if x.field_name == w.field_name
                                      for s in (x.button_states().get("normal") or [])} - {"Off"})
                out.append((w.field_name, kind, w.field_value, choices))
        d.close()
    except Exception:
        pass
    return out


@tool(id="form_fill", group="secure", name="Điền form PDF", icon=0xE70F, keywords="form fill dien",
      desc="Điền các ô form trong PDF (ô chữ, ô chọn, danh sách). Có thể làm phẳng sau khi điền.",
      opts=[Opt("fields", "Các ô form", "form", {}, prefill=_form_fields),
            Opt("flatten", "Làm phẳng sau khi điền (không sửa được nữa)", "bool", False), PASSWORD_OPT],
      note="Áp dụng cùng giá trị cho mọi file trong danh sách (cùng mẫu form).")
def t_form_fill(ctx):
    values = ctx["fields"] or {}
    if not values:
        raise ToolError("File không có ô form, hoặc chưa nhập giá trị.")

    def fn(doc, path):
        n = 0
        for pg in doc:
            for w in pg.widgets() or []:
                if w.field_name not in values:
                    continue
                v = values[w.field_name]
                if w.field_type == pymupdf.PDF_WIDGET_TYPE_CHECKBOX:
                    w.field_value = w.on_state() if v in (True, "true", "1", "x", "X", "on", "Yes", "có") else "Off"
                elif w.field_type == pymupdf.PDF_WIDGET_TYPE_RADIOBUTTON:
                    w.field_value = v if v in w.button_states().get("normal", []) else "Off"
                else:
                    w.field_value = str(v)
                w.update()
                n += 1
        if ctx["flatten"]:
            doc.bake(annots=False, widgets=True)
        ctx.log(f"{os.path.basename(path)}: điền {n} ô")
    simple_page_tool(ctx, "_da_dien", fn)


COMPRESS_STEPS = [None, (220, 150, 80), (170, 120, 70), (130, 100, 60), (110, 85, 50), (90, 72, 40), (75, 60, 30)]


def compress_bytes(data, step):
    doc = pymupdf.open("pdf", data)
    if step:
        thr, target, q = step
        doc.rewrite_images(dpi_threshold=thr, dpi_target=target, quality=q, lossy=True, lossless=True,
                           bitonal=True, color=True, gray=True)
    try:
        doc.subset_fonts()
    except Exception:
        pass
    out = doc.tobytes(garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True, use_objstms=1)
    doc.close()
    return out


@tool(id="compress", group="secure", name="Nén PDF", icon=0xF012, popular=True, keywords="compress giam dung luong email",
      desc="Giảm dung lượng: chọn mức nén, hoặc nhập dung lượng mong muốn (vd dưới 1 MB để gửi email).",
      opts=[Opt("mode", "Mức nén", "choice", "medium", [("low", "Nhẹ - giữ nguyên chất lượng ảnh"),
                                                       ("medium", "Trung bình - ảnh 150 dpi (khuyến nghị)"),
                                                       ("high", "Mạnh - ảnh 96 dpi"),
                                                       ("target", "Theo dung lượng mong muốn")]),
            Opt("target_kb", "Dung lượng tối đa", "int", 1024, min=50, max=500000, suffix=" KB",
                show_if=("mode", ("target",)), tip="vd 1024 KB = 1 MB, 200 KB cho cổng dịch vụ công"),
            PASSWORD_OPT])
def t_compress(ctx):
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        data = doc.tobytes()
        before = os.path.getsize(path)
        doc.close()
        mode = ctx["mode"]
        if mode == "target":
            limit = int(ctx["target_kb"]) * 1024
            best = None
            for step in COMPRESS_STEPS:
                ctx.progress(i / len(ctx.files), f"{os.path.basename(path)}: thử mức {COMPRESS_STEPS.index(step) + 1}")
                out = compress_bytes(data, step)
                if best is None or len(out) < len(best):
                    best = out
                if len(out) <= limit:
                    break
            out = best
            if len(out) > limit:
                ctx.log(f"⚠ {os.path.basename(path)}: chỉ nén được còn {len(out) / 1024:,.0f} KB (chưa đạt "
                        f"{limit / 1024:,.0f} KB) - file nhiều chữ / vector khó nén hơn")
        else:
            out = compress_bytes(data, {"low": None, "medium": (200, 150, 75), "high": (120, 96, 55)}[mode])
        if len(out) >= before and mode != "target":
            ctx.log(f"{os.path.basename(path)}: file đã tối ưu, không nhỏ hơn được")
        with open(ctx.out(path, "_nen"), "wb") as fh:
            fh.write(out)
        ctx.log(f"{os.path.basename(path)}: {before / 1024:,.0f} KB → {len(out) / 1024:,.0f} KB "
                f"(giảm {max(0, (1 - len(out) / before) * 100):.0f}%)")
    ctx.each(one)


@tool(id="repair", group="secure", name="Sửa file PDF lỗi", icon=0xE90F, keywords="repair fix corrupt loi",
      desc="Thử khôi phục file PDF bị hỏng, không mở được hoặc báo lỗi trong Adobe Reader.", opts=[PASSWORD_OPT])
def t_repair(ctx):
    def one(path, i):
        pymupdf.TOOLS.mupdf_warnings()             # xoá cảnh báo cũ
        doc = pymupdf.open(path)
        if doc.needs_pass and not doc.authenticate(ctx["password"] or ""):
            raise ToolError("file có mật khẩu")
        n = doc.page_count
        doc.save(ctx.out(path, "_da_sua"), garbage=4, deflate=True, clean=True)
        warn = pymupdf.TOOLS.mupdf_warnings()
        ctx.log(f"{os.path.basename(path)}: {n} trang" + (f" - đã sửa: {warn.splitlines()[0][:120]}" if warn else
                                                          " - không phát hiện lỗi"))
    ctx.each(one)


@tool(id="rasterize", group="secure", name="Chuyển trang thành ảnh", icon=0xEB9F, keywords="rasterize flatten image",
      desc="Biến mọi trang thành ảnh: không chọn / sao chép / sửa chữ được nữa (giống bản scan sạch).",
      opts=[Opt("dpi", "Độ phân giải", "int", 150, min=72, max=400, suffix=" dpi"),
            Opt("quality", "Chất lượng JPG", "int", 80, min=20, max=100),
            Opt("gray", "Đen trắng", "bool", False), PASSWORD_OPT])
def t_rasterize(ctx):
    def one(path, i):
        src = open_doc(path, ctx["password"])
        save_doc(_raster_doc(src, int(ctx["dpi"]), quality=int(ctx["quality"]), gray=ctx["gray"]),
                 ctx.out(path, "_anh_hoa"))
    ctx.each(one)


@tool(id="verify_sig", group="secure", name="Kiểm tra chữ ký số", icon=0xEB95, output="none",
      keywords="verify signature chu ky so",
      desc="Kiểm tra các chữ ký số trong file: người ký, thời điểm, file có bị sửa sau khi ký không.")
def t_verify(ctx):
    import pdf_sign

    def one(path, i):
        with open(path, "rb") as fh:
            res = pdf_sign.verify_pdf(fh.read())
        if not res:
            ctx.log(f"{os.path.basename(path)}: không có chữ ký số")
        for r in res:
            when = r["time"].astimezone().strftime("%d/%m/%Y %H:%M") if r["time"] else "?"
            state = ("✔ toàn vẹn" if r["intact"] else "✘ ĐÃ BỊ SỬA SAU KHI KÝ") if not r["error"] else "Lỗi: " + r["error"][:80]
            ctx.log(f"{os.path.basename(path)} | {r['field']} | {r['signer']} | {when} | {state}")
    ctx.each(one)


# ===========================================================================
#  OCR (nhận dạng chữ trên bản scan) - Tesseract tích hợp trong MuPDF + mô hình tessdata_best
# ===========================================================================
def tessdata_dir():
    import sys
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    d = os.path.join(base, "tessdata")
    if not os.path.exists(os.path.join(d, "vie.traineddata")):
        raise ToolError("thiếu dữ liệu OCR (thư mục tessdata)")
    return d


_OCR_FIX = str.maketrans({"\xad": "-", "‐": "-", "‑": "-", "ﬁ": "fi", "ﬂ": "fl"})


def _ocr_words(page, lang, dpi, tessdata):
    """[(Rect, từ)] nhận dạng được trên trang.

    Mỗi span của MuPDF ứng với đúng 1 từ của Tesseract. Không dùng get_text('words') vì MuPDF đôi khi
    bỏ mất dấu cách giữa 2 từ (nhất là sau chữ có dấu: 'CHỦ NGHĨA' -> 'CHỦNGHĨA')."""
    return [w for ln in _ocr_lines(page, lang, dpi, tessdata) for w in ln]


def _ocr_lines(page, lang, dpi, tessdata):
    """Như _ocr_words nhưng giữ nhóm theo dòng: [[(Rect, từ), ...], ...] - dùng để dựng lại bố cục."""
    tp = page.get_textpage_ocr(flags=0, language=lang, dpi=dpi, full=True, tessdata=tessdata)
    lines = []
    for b in page.get_text("rawdict", textpage=tp)["blocks"]:
        for ln in b.get("lines", []):
            words = []
            for s in ln["spans"]:
                chars = [c for c in s["chars"] if c["c"].strip()]
                if not chars:
                    continue
                r = pymupdf.Rect(chars[0]["bbox"])
                for c in chars[1:]:
                    r |= pymupdf.Rect(c["bbox"])
                words.append((r, "".join(c["c"] for c in chars).translate(_OCR_FIX)))
            if words:
                lines.append(words)
    return lines


def _word_score(words):
    """Điểm 'giống chữ thật' để đoán hướng trang: số ký tự chữ / số trong các từ dài >= 2."""
    return sum(len(t) for _, t in words if len(t) >= 2 and sum(c.isalnum() for c in t) >= len(t) * 0.7)


def _orient_page(page, lang, tessdata):
    """Thử xoay 0/90/180/270 ở độ phân giải thấp, chọn hướng đọc được nhiều chữ nhất."""
    best, best_rot = -1, 0
    base = page.rotation
    for rot in (0, 180, 90, 270):
        page.set_rotation((base + rot) % 360)
        try:
            s = _word_score(_ocr_words(page, lang, 100, tessdata))
        except Exception:
            s = 0
        if s > best:
            best, best_rot = s, rot
        if rot == 180 and best > 150:               # đủ rõ, khỏi thử chiều ngang
            break
    page.set_rotation((base + best_rot) % 360)
    return best_rot


def _add_text_layer(page, words, font):
    """Chèn lớp chữ vô hình (render_mode 3) đúng vị trí từng từ - tìm kiếm / bôi đen / copy được."""
    tw = pymupdf.TextWriter(page.rect)
    for r, t in words:
        h = r.height
        fs = max(2.0, h * 0.86)
        w = font.text_length(t, fontsize=fs)
        if w > 0:
            k = r.width / w
            if 0.55 < k < 1.8:
                fs *= k                                 # khớp bề ngang để vùng chọn trùng chữ
        try:
            # dấu cách thật sau mỗi từ -> chữ trích ra / copy có khoảng trắng đúng
            tw.append(pymupdf.Point(r.x0, r.y1 - h * 0.2), t + " ", font=font, fontsize=fs)
        except Exception:
            pass
    tw.write_text(page, render_mode=3)


def _ocr_job(png, w, h, lang, dpi, tess, orient):
    """Chạy trong TIẾN TRÌNH CON: Tesseract trong MuPDF giữ GIL suốt lúc nhận dạng, chạy trong luồng của app
    sẽ làm treo giao diện (Not Responding). Trả về (góc cần xoay thêm, [(rect, từ)]) theo toạ độ trang
    sau khi xoay + remove_rotation."""
    doc = pymupdf.open()
    pg = doc.new_page(width=w, height=h)
    pg.insert_image(pg.rect, stream=png)
    rot = _orient_page(pg, lang, tess) if orient else 0
    pg.remove_rotation()
    return rot, [(tuple(r), t) for r, t in _ocr_words(pg, lang, dpi, tess)]


def _remove_rules(png):
    """Xoá đường kẻ ngang / dọc dài (kẻ bảng, gạch chân) khỏi ảnh xám trước khi OCR: Tesseract đọc ô bảng có
    đường kẻ thường ra chữ rác."""
    import io
    import numpy as np
    from PIL import Image
    import imgops
    try:
        img = np.array(Image.open(io.BytesIO(png)).convert("L"))
    except Exception:
        return png
    ink = imgops.adaptive_mean_inv(img, 25, 15)
    H, W = ink.shape
    rules = imgops.dilate3(imgops.open_line(ink, max(40, W // 25), 1) | imgops.open_line(ink, max(40, H // 35), 0))
    if not rules.any():
        return png
    img[rules] = 255
    buf = io.BytesIO()
    Image.fromarray(img).save(buf, "PNG")
    return buf.getvalue()


def _ocr_lines_job(png, w, h, lang, dpi, tess, orient):
    """Như _ocr_job nhưng trả về từ theo dòng: (góc xoay, [[(rect, từ), ...], ...])."""
    doc = pymupdf.open()
    pg = doc.new_page(width=w, height=h)
    pg.insert_image(pg.rect, stream=_remove_rules(png))
    rot = _orient_page(pg, lang, tess) if orient else 0
    pg.remove_rotation()
    return rot, [[(tuple(r), t) for r, t in ln] for ln in _ocr_lines(pg, lang, dpi, tess)]


def _run_ocr_jobs(items, make_args, tick, job=None):
    """Nhận dạng song song nhiều trang ở tiến trình con. make_args(item) -> tham số _ocr_job (render ảnh
    lần lượt để không giữ quá nhiều ảnh trong bộ nhớ). Trả về {item: kết quả}."""
    job = job or _ocr_job
    from concurrent.futures import FIRST_COMPLETED, ProcessPoolExecutor, wait
    results = {}
    if not items:
        return results
    workers = max(1, min(4, (os.cpu_count() or 2) - 1, len(items)))
    try:
        ex = ProcessPoolExecutor(max_workers=workers)
    except Exception:                                   # không tạo được tiến trình con -> chạy tại chỗ
        for k, it in enumerate(items):
            tick(k)
            results[it] = job(*make_args(it))
        return results
    queue = iter(items)
    pending = {}

    def submit():
        it = next(queue, None)
        if it is not None:
            pending[ex.submit(job, *make_args(it))] = it

    try:
        for _ in range(workers * 2):
            submit()
        tick(0)
        while pending:
            done, _ = wait(pending, timeout=0.3, return_when=FIRST_COMPLETED)
            for f in done:
                results[pending.pop(f)] = f.result()
                submit()
            tick(len(results))                          # cũng là chỗ kiểm tra nút Dừng
    finally:
        ex.shutdown(wait=not pending, cancel_futures=True)
    return results


@tool(id="ocr", group="secure", name="OCR - nhận dạng chữ bản scan", icon=0xE8D3, popular=True,
      accept=PDF_EXT + IMAGE_EXT, keywords="ocr scan nhan dang chu tim kiem searchable",
      desc="Biến PDF / ảnh scan thành PDF tìm kiếm, bôi đen, copy được chữ (tiếng Việt + tiếng Anh), giữ nguyên hình scan.",
      opts=[Opt("lang", "Ngôn ngữ", "choice", "vie", [("vie", "Tiếng Việt (đọc được cả chữ tiếng Anh xen kẽ)"),
                                                     ("eng", "Tiếng Anh"),
                                                     ("eng+vie", "Chủ yếu tiếng Anh, có ít tiếng Việt")],
                tip="Văn bản tiếng Việt nên chọn 'Tiếng Việt': trộn thêm mô hình tiếng Anh dễ làm mất dấu"),
            Opt("dpi", "Độ phân giải nhận dạng", "int", 300, min=150, max=600, suffix=" dpi",
                tip="300 dpi là tối ưu. Chữ rất nhỏ (< 8pt) dùng 400"),
            Opt("mode", "Trang đã có chữ", "choice", "skip", [("skip", "Bỏ qua (chỉ OCR trang scan)"),
                                                              ("force", "OCR lại tất cả các trang")]),
            Opt("rotate", "Tự xoay trang scan bị ngược / nằm ngang", "bool", True),
            Opt("txt", "Xuất thêm file .txt", "bool", False),
            pages_opt(), PASSWORD_OPT],
      note="Chạy offline. Khoảng 3-8 giây / trang tuỳ máy.")
def t_ocr(ctx):
    tess = tessdata_dir()
    lang = ctx["lang"]
    # Font cho lớp chữ vô hình: tessdata\ocr_font.ttf (kèm theo app). Arial / Times... dùng chung 1 glyph cho
    # "-" và gạch mềm, cho dấu cách và dấu cách cứng -> copy chữ ra bị sai ký tự.
    ff = os.path.join(tess, "ocr_font.ttf")
    font = pymupdf.Font(fontfile=ff if os.path.exists(ff) else _font_path("Arial"))

    dpi = int(ctx["dpi"])

    def one(path, fi):
        doc = open_doc(path, ctx["password"])
        pages = parse_pages(ctx["pages"], doc.page_count)
        skipped, rotated = 0, 0
        todo = []
        for p in pages:
            pg = doc[p]
            if ctx["mode"] == "skip" and len(pg.get_text("text").strip()) > 20:
                skipped += 1
                continue
            if ctx["mode"] == "force":
                # bỏ lớp chữ cũ (thường là OCR kém chất lượng) trước khi nhận dạng lại
                pg.add_redact_annot(pg.rect)
                pg.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                                    graphics=pymupdf.PDF_REDACT_LINE_ART_NONE, text=pymupdf.PDF_REDACT_TEXT_REMOVE)
            todo.append(p)

        def job(p):
            pg = doc[p]
            png = pg.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY).tobytes("png")
            return (png, pg.rect.width, pg.rect.height, lang, dpi, tess, bool(ctx["rotate"]))

        def tick(n_done):
            ctx.progress((fi + n_done / max(1, len(todo))) / len(ctx.files),
                         f"{os.path.basename(path)} - đã nhận dạng {n_done}/{len(todo)} trang")

        results = _run_ocr_jobs(todo, job, tick)
        for p in todo:
            rot, words = results[p]
            pg = doc[p]
            if rot:
                rotated += 1
                pg.set_rotation((pg.rotation + rot) % 360)
            pg.remove_rotation()
            if words:
                _add_text_layer(pg, [(pymupdf.Rect(r), t) for r, t in words], font)
        done = len(todo)
        texts = [doc[p].get_text("text") for p in pages]
        out = ctx.out(path, "_ocr")
        save_doc(doc, out, fonts=True)
        if ctx["txt"]:
            with open(os.path.splitext(out)[0] + ".txt", "w", encoding="utf-8") as fh:
                fh.write("\n\n".join(f"===== Trang {p + 1} =====\n{t}" for p, t in zip(pages, texts)))
            ctx.outputs.append(os.path.splitext(out)[0] + ".txt")
        ctx.log(f"{os.path.basename(path)}: nhận dạng {done} trang" + (f", bỏ qua {skipped} trang đã có chữ" if skipped else "")
                + (f", tự xoay {rotated} trang" if rotated else ""))
    ctx.each(one)


def _analyse_scan(pix, lines, erase, k):
    """Đọc ảnh trang (RGB) tại các ô chữ OCR (k = số điểm ảnh / 1 pt).
    - Bỏ 'từ' giả: OCR đọc nhầm logo / con dấu / hình thành chữ (ô quá đậm, quá cao, không có chữ cái) -> giữ
      nguyên trên nền, không xoá.
    - erase=True: tô ô chữ thật bằng màu nền xung quanh (logo, con dấu, kẻ bảng còn lại trên ảnh nền).
    Trả về (ảnh numpy, các dòng đã lọc, [(màu, độ dày nét, đường chân chữ, đỉnh nét chữ)] từng dòng - đơn vị pt).
    Ô chữ Tesseract có chung chiều cao cả dòng (dòng có logo bên cạnh thì rất cao) nên đường chân chữ / cỡ chữ
    lấy theo nét mực thật: hàng điểm ảnh cuối còn >= 25% mực của hàng đậm nhất (bỏ phần đuôi g, p, y)."""
    import statistics
    import numpy as np
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)[:, :, :3].copy()
    wl = np.array([0.299, 0.587, 0.114])
    lum = arr @ wl
    H, W = lum.shape
    heights = [r.height for ln in lines for r, t in ln if sum(ch.isalnum() for ch in t) >= 2]
    med_h = statistics.median(heights) if heights else 0
    out_lines, info = [], []
    for ln in lines:
        cols, wsum, dark, wpx, runs, nstarts = np.zeros(3), 0, 0, 0, 0, 0
        kept = []
        ly0 = max(0, int(min(r.y0 for r, _ in ln) * k) - 1)
        profile = np.zeros(max(1, min(H, int(max(r.y1 for r, _ in ln) * k + 0.999) + 1) - ly0))
        for r, t in ln:
            x0, y0 = max(0, int(r.x0 * k) - 1), max(0, int(r.y0 * k) - 1)
            x1, y1 = min(W, int(r.x1 * k + 0.999) + 1), min(H, int(r.y1 * k + 0.999) + 1)
            if x1 - x0 < 2 or y1 - y0 < 2:
                continue
            # nền = trung vị viền ngoài ô chữ (rộng thêm 2 px)
            X0, Y0, X1, Y1 = max(0, x0 - 2), max(0, y0 - 2), min(W, x1 + 2), min(H, y1 + 2)
            ring = np.concatenate([arr[Y0, X0:X1], arr[Y1 - 1, X0:X1], arr[Y0:Y1, X0], arr[Y0:Y1, X1 - 1]])
            bg = np.median(ring, axis=0)
            m = lum[y0:y1, x0:x1] < float(bg @ wl) - 45      # điểm ảnh thuộc nét chữ
            n = int(m.sum())
            dens = n / m.size
            alnum = sum(ch.isalnum() for ch in t)
            fake = (dens > 0.5 or n < 4                      # khối đặc (logo, ô tô màu) / không có nét
                    or (alnum == 0 and len(t) <= 2 and med_h and r.height > med_h * 1.3)
                    or (med_h and r.height > med_h * 2.6 and len(ln) <= 3))
            if fake:
                continue
            kept.append((r, t))
            cols += arr[y0:y1, x0:x1][m].mean(axis=0) * n
            wsum += n
            dark += n
            wpx += x1 - x0
            seg = m.sum(axis=1)
            a = y0 - ly0
            profile[a:a + len(seg)] += seg[:max(0, len(profile) - a)]
            # độ dày nét = độ dài trung bình các đoạn mực theo chiều ngang (để nhận chữ đậm)
            starts = int((m[:, 1:] & ~m[:, :-1]).sum() + m[:, 0].sum())
            runs += n
            nstarts += starts
            if erase:                                    # rộng thêm 1 px để không sót viền mờ của chữ
                arr[max(0, y0 - 1):min(H, y1 + 1), max(0, x0 - 1):min(W, x1 + 1)] = bg.astype(np.uint8)
        if not kept:
            continue
        c = cols / wsum if wsum else np.zeros(3)
        # chữ gần đen / xám đậm không rõ màu -> đen tuyền (bản scan làm chữ đen thành xám)
        if c.max() < 100 or (c.max() - c.min() < 40 and c.mean() < 140):
            c = np.zeros(3)
        rows = np.nonzero(profile >= profile.max() * 0.25)[0]
        top_px = ly0 + int(np.nonzero(profile > 0)[0][0])
        base_px = ly0 + int(rows[-1]) + 1
        out_lines.append(kept)
        stroke = max(0.0, runs / nstarts - 0.8) / k if nstarts else 0   # trừ phần nhoè ~0.8 px của bản scan
        info.append((tuple(round(float(v) / 255, 2) for v in c), stroke, base_px / k, top_px / k))
    return arr, out_lines, info


def _place_editable_text(page, lines, info, font, font_b, detect_bold):
    """Chèn chữ THẬT (nhìn thấy, sửa được) theo từng dòng: mỗi từ đúng vị trí, cỡ chữ chung cho cả dòng."""
    import statistics

    def size_for(ln, f, h):
        # cỡ chữ: khớp bề ngang các từ (chính xác hơn chiều cao vì có dấu tiếng Việt / chữ hoa); lấy mức thấp
        # trong dòng để từ không đè lên từ bên cạnh
        ks = []
        for r, t in ln:
            w1 = f.text_length(t, fontsize=1)
            if w1 > 0 and len(t) >= 2:
                ks.append(r.width / w1)
        if ks:
            ks.sort()
            fs = min(statistics.median(ks), ks[len(ks) // 4] * 1.08)
        else:
            fs = h
        return max(h * 0.6, min(fs, h * 1.7, 200))

    # chữ đậm: tỉ lệ độ dày nét / cỡ chữ lớn hơn hẳn mức chung của trang (đo ở 150 dpi: thường ~0.14,
    # đậm >= 0.17). Dòng quá ngắn (1 chữ số...) đo không tin cậy -> giữ chữ thường.
    ratios = [i[1] / size_for(ln, font, max(1.0, i[2] - i[3])) for ln, i in zip(lines, info)]
    norm = [r for r, ln in zip(ratios, lines) if len(ln) >= 3 and r > 0]
    med = statistics.median(norm) if norm else 0.14
    limit = max(med * 1.17, 0.12)
    writers = {}
    for ln, (color, _stroke, base, top), ratio in zip(lines, info, ratios):
        h = max(1.0, base - top)                         # chiều cao nét mực (đỉnh dấu / chữ hoa -> chân chữ)
        letters = sum(ch.isalpha() for _, t in ln for ch in t)
        bold = bool(detect_bold and ratio > limit and letters >= 3)
        f = font_b if bold else font
        fs = size_for(ln, f, h)
        tw = writers.setdefault(color, pymupdf.TextWriter(page.rect))
        for r, t in ln:
            try:
                tw.append(pymupdf.Point(r.x0, base), t + " ", font=f, fontsize=fs)
            except Exception:
                pass
    for color, tw in writers.items():
        tw.write_text(page, color=color)


@tool(id="ocr_editable", group="secure", name="OCR → PDF chỉnh sửa được", icon=0xE8AC, popular=True,
      accept=PDF_EXT + IMAGE_EXT, keywords="ocr scan editable sua chu text layout bo cuc convert",
      desc="Nhận dạng bản scan rồi dựng lại trang bằng CHỮ THẬT (giữ vị trí, cỡ, màu, chữ đậm) - sửa chữ trực tiếp "
           "trong trình sửa. Logo, con dấu, chữ ký, kẻ bảng được giữ lại trên nền.",
      opts=[Opt("lang", "Ngôn ngữ", "choice", "vie", [("vie", "Tiếng Việt (đọc được cả chữ tiếng Anh xen kẽ)"),
                                                     ("eng", "Tiếng Anh"),
                                                     ("eng+vie", "Chủ yếu tiếng Anh, có ít tiếng Việt")]),
            Opt("background", "Nền trang", "choice", "clean",
                [("clean", "Giữ hình ảnh của bản scan, xoá chữ cũ (logo, con dấu, bảng giữ nguyên)"),
                 ("white", "Nền trắng - chỉ giữ chữ (file nhẹ nhất)")]),
            Opt("font", "Font chữ", "choice", "Arial", FONT_CHOICES),
            Opt("bold", "Nhận biết chữ đậm", "bool", True),
            Opt("dpi", "Độ phân giải nhận dạng", "int", 300, min=150, max=600, suffix=" dpi"),
            Opt("bg_dpi", "Độ phân giải ảnh nền", "int", 150, min=72, max=300, suffix=" dpi",
                tip="Ảnh nền chỉ còn logo / con dấu / kẻ bảng - 150 dpi là đủ"),
            Opt("mode", "Trang đã có chữ", "choice", "skip", [("skip", "Giữ nguyên (chỉ dựng lại trang scan)"),
                                                              ("force", "Dựng lại tất cả các trang")]),
            Opt("rotate", "Tự xoay trang scan bị ngược / nằm ngang", "bool", True),
            pages_opt(), out_name_opt(), PASSWORD_OPT],
      note="Xem trước 2 trang đầu ở khung bên phải (Kết quả / File gốc). Chữ nhận dạng sai sửa trực tiếp trong trình sửa "
           "(chế độ Sửa file). Khoảng 3-8 giây / trang.")
def t_ocr_editable(ctx):
    tess = tessdata_dir()
    lang = ctx["lang"]
    dpi = int(ctx["dpi"])
    bg_dpi = int(ctx["bg_dpi"])
    family = ctx["font"] or "Arial"
    # Arial thường: dùng Liberation Sans kèm app (cùng số đo với Arial) - Arial dùng chung glyph cho '-' và gạch
    # mềm nên copy chữ ra bị sai ký tự
    ff, fb = os.path.join(tess, "ocr_font.ttf"), os.path.join(tess, "ocr_font_bold.ttf")
    lib = family == "Arial"
    font = pymupdf.Font(fontfile=ff if lib and os.path.exists(ff) else _font_path(family))
    font_b = pymupdf.Font(fontfile=fb if lib and os.path.exists(fb) else _font_path(family, bold=True))

    def one(path, fi):
        doc = open_doc(path, ctx["password"])
        pages = parse_pages(ctx["pages"], doc.page_count)
        todo = [p for p in pages
                if ctx["mode"] == "force" or len(doc[p].get_text("text").strip()) <= 20]

        def job(p):
            pg = doc[p]
            png = pg.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY).tobytes("png")
            return (png, pg.rect.width, pg.rect.height, lang, dpi, tess, bool(ctx["rotate"]))

        def tick(n_done):
            ctx.progress((fi + 0.85 * n_done / max(1, len(todo))) / len(ctx.files),
                         f"{os.path.basename(path)} - đã nhận dạng {n_done}/{len(todo)} trang")

        results = _run_ocr_jobs(todo, job, tick, job=_ocr_lines_job)
        out = pymupdf.open()
        n_words = 0
        for i in range(doc.page_count):
            if i not in results:
                out.insert_pdf(doc, from_page=i, to_page=i)
                continue
            rot, raw = results[i]
            pg = doc[i]
            if rot:
                pg.set_rotation((pg.rotation + rot) % 360)
            pg.remove_rotation()
            lines = [[(pymupdf.Rect(r), t) for r, t in ln] for ln in raw]
            pix = pg.get_pixmap(dpi=bg_dpi, colorspace=pymupdf.csRGB, alpha=False, annots=False)
            erase = ctx["background"] == "clean"
            arr, lines, info = _analyse_scan(pix, lines, erase, pix.width / pg.rect.width)
            n_words += sum(len(ln) for ln in lines)
            np_ = out.new_page(width=pg.rect.width, height=pg.rect.height)
            if erase:
                import numpy as np
                lum = arr.mean(axis=2)
                if (lum < 225).mean() > 0.0015:             # còn hình / kẻ bảng / con dấu -> giữ làm nền
                    bg = pymupdf.Pixmap(pymupdf.csRGB, arr.shape[1], arr.shape[0], np.ascontiguousarray(arr).tobytes(), 0)
                    np_.insert_image(np_.rect, stream=bg.tobytes("jpg", jpg_quality=80))
            _place_editable_text(np_, lines, info, font, font_b, bool(ctx["bold"]))
        toc = doc.get_toc(simple=False)
        if toc:
            try:
                out.set_toc(toc)
            except Exception:
                pass
        save_doc(out, ctx.out(path, "_sua_duoc", name=ctx["out_name"] or None), fonts=True)
        ctx.log(f"{os.path.basename(path)}: dựng lại {len(results)} trang thành chữ sửa được ({n_words} từ)"
                + (f", giữ nguyên {len(pages) - len(todo)} trang đã có chữ" if len(pages) > len(todo) else ""))
    ctx.each(one)


# ===========================================================================
#  BỔ SUNG v1.3
# ===========================================================================
def _rebuild_pages(doc, pages, place):
    """Tạo tài liệu mới: trang trong `pages` được vẽ lại bằng place(new_page, src, i), trang khác giữ nguyên."""
    src = pymupdf.open()
    src.insert_pdf(doc)
    out = pymupdf.open()
    for i in range(src.page_count):
        if i in pages:
            r = src[i].rect
            np_ = out.new_page(width=r.width, height=r.height)
            place(np_, src, i)
        else:
            out.insert_pdf(src, from_page=i, to_page=i)
    toc = doc.get_toc(simple=False)
    if toc:
        try:
            out.set_toc(toc)
        except Exception:
            pass
    return out


@tool(id="rotate_custom", group="organize", name="Xoay góc tuỳ ý", icon=0xE7AD,
      keywords="rotate custom angle xoay nghieng",
      desc="Xoay nội dung trang theo góc bất kỳ (vd 1.5°, 15°, 45°), khổ trang giữ nguyên, chữ vẫn chọn được.",
      opts=[Opt("angle", "Góc xoay", "float", 5.0, min=-180, max=180, step=0.5, suffix=" °",
                tip="Số dương = xoay ngược chiều kim đồng hồ"),
            pages_opt(), PASSWORD_OPT],
      note="Ghi chú / liên kết trên trang được xoay sẽ bị bỏ.")
def t_rotate_custom(ctx):
    ang = float(ctx["angle"])

    def one(path, i):
        doc = open_doc(path, ctx["password"])
        pages = set(parse_pages(ctx["pages"], doc.page_count))
        out = _rebuild_pages(doc, pages, lambda np_, src, k: np_.show_pdf_page(np_.rect, src, k, rotate=ang))
        save_doc(out, ctx.out(path, "_xoay_goc"))
    ctx.each(one)


def _skew_angle(page, max_deg=5.0):
    """Ước lượng độ nghiêng của bản scan (độ, dương = chữ dốc lên bên phải) bằng phương pháp chiếu dòng."""
    import numpy as np
    import imgops
    pix = page.get_pixmap(dpi=100, colorspace=pymupdf.csGRAY, alpha=False)
    a = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.stride)[:, :pix.width]
    bw = np.where(a <= imgops.otsu_threshold(a), 255, 0).astype(np.uint8)
    if bw.mean() < 0.5:                                     # trang gần như trắng
        return 0.0

    def score(deg):
        s = imgops.rotate(bw, deg).sum(axis=1, dtype=np.float64)
        return float(((s[1:] - s[:-1]) ** 2).sum())

    best = max(np.arange(-max_deg, max_deg + 1e-6, 0.25), key=score)
    fine = max(np.arange(best - 0.25, best + 0.25 + 1e-6, 0.05), key=score)
    return float(fine)


@tool(id="deskew", group="organize", name="Làm thẳng trang scan", icon=0xE7AD,
      keywords="deskew straighten scan nghieng thang",
      desc="Tự phát hiện và xoay thẳng các trang scan bị nghiêng (tới ±5°).",
      opts=[Opt("max", "Độ nghiêng tối đa cần dò", "float", 5.0, min=0.5, max=15, step=0.5, suffix=" °"),
            Opt("min", "Bỏ qua nếu lệch dưới", "float", 0.15, min=0, max=2, step=0.05, suffix=" °"),
            pages_opt(), PASSWORD_OPT],
      note="Phù hợp nhất với bản scan có chữ. Ghi chú / liên kết trên trang được xoay sẽ bị bỏ.")
def t_deskew(ctx):
    def one(path, fi):
        doc = open_doc(path, ctx["password"])
        pages = parse_pages(ctx["pages"], doc.page_count)
        angles = {}
        for k, p in enumerate(pages):
            ctx.progress((fi + k / len(pages)) / len(ctx.files), f"{os.path.basename(path)} - trang {p + 1}")
            a = _skew_angle(doc[p], float(ctx["max"]))
            if abs(a) >= float(ctx["min"]):
                angles[p] = a
        out = _rebuild_pages(doc, set(angles), lambda np_, src, k: np_.show_pdf_page(np_.rect, src, k,
                                                                                     rotate=angles[k]))
        save_doc(out, ctx.out(path, "_thang"))
        ctx.log(f"{os.path.basename(path)}: làm thẳng {len(angles)}/{len(pages)} trang"
                + (" (" + ", ".join(f"tr.{p + 1}: {a:+.2f}°" for p, a in list(angles.items())[:8]) + ")"
                   if angles else ""))
    ctx.each(one)


@tool(id="duplex_collate", group="organize", name="Sắp xếp bản scan 2 mặt (1 file)", icon=0xE8AB,
      keywords="duplex collate scan 2 mat",
      desc="File scan có nửa đầu là mặt trước, nửa sau là mặt sau → sắp lại đúng thứ tự 1, 2, 3, 4...",
      opts=[Opt("reverse", "Nửa sau có thứ tự ngược (lật cả xấp giấy để scan)", "bool", True), PASSWORD_OPT])
def t_duplex(ctx):
    def fn(doc, path):
        n = doc.page_count
        half = (n + 1) // 2
        front = list(range(half))
        back = list(range(half, n))
        if ctx["reverse"]:
            back.reverse()
        order = []
        for i in range(half):
            order.append(front[i])
            if i < len(back):
                order.append(back[i])
        doc.select(order)
    simple_page_tool(ctx, "_2_mat", fn)


@tool(id="combine_single_page", group="organize", name="Ghép thành 1 trang dài", icon=0xE8A9,
      keywords="combine single page long merge pages",
      desc="Nối tất cả các trang thành một trang dài liền mạch (dọc hoặc ngang) - tiện đọc trên màn hình.",
      opts=[Opt("dir", "Hướng nối", "choice", "v", [("v", "Từ trên xuống dưới"), ("h", "Từ trái sang phải")]),
            Opt("gap", "Khoảng cách giữa các trang", "float", 0, min=0, max=50, step=1, suffix=" mm"),
            Opt("line", "Kẻ vạch phân cách", "bool", True), pages_opt(), PASSWORD_OPT],
      note="PDF giới hạn khổ trang 5 m; nội dung dài hơn sẽ được thu nhỏ cho vừa.")
def t_combine_single(ctx):
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        pages = parse_pages(ctx["pages"], doc.page_count)
        rects = [doc[p].rect for p in pages]
        gap = float(ctx["gap"]) * MM
        vert = ctx["dir"] == "v"
        W = max(r.width for r in rects) if vert else sum(r.width for r in rects) + gap * (len(rects) - 1)
        H = sum(r.height for r in rects) + gap * (len(rects) - 1) if vert else max(r.height for r in rects)
        k = min(1.0, 14400 / max(W, H))
        out = pymupdf.open()
        pg = out.new_page(width=W * k, height=H * k)
        pos = 0.0
        for n, (p, r) in enumerate(zip(pages, rects)):
            if vert:
                box = pymupdf.Rect(0, pos, r.width, pos + r.height)
                pos += r.height + gap
            else:
                box = pymupdf.Rect(pos, 0, pos + r.width, r.height)
                pos += r.width + gap
            box = box * k
            pg.show_pdf_page(box, doc, p)
            if ctx["line"] and n:
                a, b = ((box.tl, box.tr) if vert else (box.tl, box.bl))
                pg.draw_line(a, b, color=(0.75, 0.75, 0.75), width=0.6)
        save_doc(out, ctx.out(path, "_1_trang"))
        ctx.log(f"{os.path.basename(path)}: {len(pages)} trang → 1 trang {W * k / MM:.0f} x {H * k / MM:.0f} mm")
    ctx.each(one)


LABEL_STYLES = {"1": "D", "i": "r", "I": "R", "a": "a", "A": "A", "-": ""}


def _labels_text(path):
    try:
        doc = pymupdf.open(path)
        rules = doc.get_page_labels()
    except Exception:
        return ""
    back = {v: k for k, v in LABEL_STYLES.items()}
    return "\n".join(f"{r['startpage'] + 1} | {back.get(r.get('style', ''), '1')} | {r.get('prefix', '')} | "
                     f"{r.get('firstpagenum', 1)}" for r in rules)


@tool(id="page_labels", group="organize", name="Đặt nhãn số trang (i, ii, A-1...)", icon=0xE8EC,
      keywords="page labels nhan trang roman",
      desc="Đặt số trang hiển thị trong trình đọc PDF: phần mở đầu i, ii, iii; phụ lục A-1, A-2...",
      opts=[Opt("rules", "Quy tắc", "textarea", "1 | i |  | 1\n5 | 1 |  | 1", prefill=_labels_text,
                tip="Mỗi dòng: từ trang | kiểu | tiền tố | số bắt đầu. Kiểu: 1 (1,2,3), i (i,ii), I (I,II), "
                    "a (a,b), A (A,B), - (chỉ tiền tố)",
                placeholder="1 | i |  | 1\n5 | 1 |  | 1\n20 | 1 | A- | 1"),
            Opt("clear", "Xoá hết nhãn trang (về 1, 2, 3...)", "bool", False), PASSWORD_OPT])
def t_page_labels(ctx):
    def fn(doc, path):
        if ctx["clear"]:
            doc.set_page_labels([])
            return
        rules = []
        for ln in (ctx["rules"] or "").splitlines():
            if not ln.strip():
                continue
            parts = [x.strip() for x in ln.split("|")] + ["", "", ""]
            if not parts[0].isdigit():
                raise ToolError(f"Dòng nhãn sai định dạng: '{ln}' (đúng: từ trang | kiểu | tiền tố | số bắt đầu)")
            style = LABEL_STYLES.get(parts[1] or "1")
            if style is None:
                raise ToolError(f"Kiểu nhãn không hợp lệ: '{parts[1]}' (dùng 1, i, I, a, A hoặc -)")
            start = max(1, min(int(parts[0]), doc.page_count)) - 1
            rules.append({"startpage": start, "style": style, "prefix": parts[2],
                          "firstpagenum": int(parts[3]) if parts[3].isdigit() else 1})
        if not rules:
            raise ToolError("chưa nhập quy tắc nhãn trang")
        doc.set_page_labels(sorted(rules, key=lambda r: r["startpage"]))
    simple_page_tool(ctx, "_nhan_trang", fn)


def _paper_name(w, h):
    a, b = sorted((w / MM, h / MM))
    for name, (x, y) in {"A0": (841, 1189), "A1": (594, 841), "A2": (420, 594), "A3": (297, 420), "A4": (210, 297),
                         "A5": (148, 210), "A6": (105, 148), "B4": (250, 353), "B5": (176, 250),
                         "Letter": (215.9, 279.4), "Legal": (215.9, 355.6), "Tabloid": (279.4, 431.8)}.items():
        if abs(a - x) < 3 and abs(b - y) < 3:
            return name
    return "Tuỳ chỉnh"


@tool(id="page_dimensions", group="organize", name="Xem kích thước trang", icon=0xE740, output="none",
      keywords="page size dimensions kho giay",
      desc="Liệt kê khổ giấy, kích thước (mm), hướng và góc xoay của từng trang; có thể xuất CSV.",
      opts=[Opt("csv", "Xuất thêm file CSV", "bool", False), PASSWORD_OPT])
def t_page_dims(ctx):
    def one(path, i):
        doc = open_doc(path, ctx["password"])
        rows = []
        groups = {}
        for p in doc:
            r = p.rect
            name = _paper_name(r.width, r.height)
            orient = "ngang" if r.width > r.height else "dọc"
            rows.append((p.number + 1, name, round(r.width / MM, 1), round(r.height / MM, 1), orient, p.rotation))
            groups.setdefault((name, orient, round(r.width / MM), round(r.height / MM)), []).append(p.number + 1)
        ctx.log(f"{os.path.basename(path)}: {doc.page_count} trang")
        for (name, orient, w, h), pg in groups.items():
            rng = _ranges(pg)
            ctx.log(f"   {name} {orient} {w} x {h} mm: {len(pg)} trang ({rng})")
        if ctx["csv"]:
            import csv
            out = ctx.out(path, "_kich_thuoc", ".csv")
            with open(out, "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(["Trang", "Khổ", "Rộng (mm)", "Cao (mm)", "Hướng", "Xoay"])
                w.writerows(rows)
    ctx.each(one)


def _ranges(nums):
    out, start, prev = [], None, None
    for n in nums:
        if start is None:
            start = prev = n
        elif n == prev + 1:
            prev = n
        else:
            out.append(f"{start}-{prev}" if prev != start else str(start))
            start = prev = n
    if start is not None:
        out.append(f"{start}-{prev}" if prev != start else str(start))
    s = ", ".join(out)
    return s if len(s) < 120 else s[:117] + "..."


def _attach_text(path):
    try:
        doc = pymupdf.open(path)
        return "\n".join(f"{n} | {doc.embfile_info(n).get('description', '')}" for n in doc.embfile_names())
    except Exception:
        return ""


@tool(id="edit_attach", group="organize", name="Sửa / xoá file đính kèm", icon=0xE723,
      keywords="attachment edit delete dinh kem",
      desc="Xem danh sách file nhúng trong PDF, xoá file không cần hoặc sửa mô tả.",
      opts=[Opt("items", "File đính kèm", "textarea", "", prefill=_attach_text,
                tip="Mỗi dòng: tên file | mô tả. Xoá dòng nào thì file đó bị xoá khỏi PDF.",
                placeholder="(file không có đính kèm)"),
            Opt("remove_all", "Xoá tất cả file đính kèm", "bool", False), PASSWORD_OPT],
      max_files=1)
def t_edit_attach(ctx):
    def fn(doc, path):
        names = doc.embfile_names()
        keep = {}
        if not ctx["remove_all"]:
            for ln in (ctx["items"] or "").splitlines():
                if ln.strip():
                    name, _, desc = ln.partition("|")
                    keep[name.strip()] = desc.strip()
        removed = 0
        for n in names:
            if n not in keep:
                doc.embfile_del(n)
                removed += 1
            elif keep[n] != (doc.embfile_info(n).get("description") or ""):
                doc.embfile_upd(n, desc=keep[n])
        if ctx["remove_all"]:
            for pg in doc:
                for a in list(pg.annots(types=[pymupdf.PDF_ANNOT_FILE_ATTACHMENT]) or []):
                    pg.delete_annot(a)
        ctx.log(f"{os.path.basename(path)}: xoá {removed} file, còn {len(names) - removed} file")
    simple_page_tool(ctx, "_dinh_kem", fn)


@tool(id="adjust_colors", group="stamp", name="Chỉnh sáng / tương phản", icon=0xE793,
      keywords="adjust colors brightness contrast saturation gamma scan",
      desc="Chỉnh độ sáng, tương phản, độ bão hoà, gamma, làm nét - làm rõ bản scan mờ. Trang sẽ thành ảnh.",
      opts=[Opt("brightness", "Độ sáng", "int", 100, min=20, max=300, suffix=" %"),
            Opt("contrast", "Tương phản", "int", 120, min=20, max=300, suffix=" %"),
            Opt("saturation", "Độ bão hoà màu", "int", 100, min=0, max=300, suffix=" %",
                tip="0% = đen trắng"),
            Opt("gamma", "Gamma", "float", 1.0, min=0.3, max=3, step=0.1, tip="< 1 làm tối vùng giữa, > 1 làm sáng"),
            Opt("sharpen", "Làm nét chữ", "bool", True),
            Opt("dpi", "Độ phân giải", "int", 200, min=72, max=400, suffix=" dpi"),
            Opt("quality", "Chất lượng JPEG", "int", 85, min=30, max=100), pages_opt(), PASSWORD_OPT])
def t_adjust(ctx):
    from PIL import ImageEnhance, ImageFilter

    def eff(im):
        im = im.convert("RGB")
        if ctx["brightness"] != 100:
            im = ImageEnhance.Brightness(im).enhance(ctx["brightness"] / 100)
        if ctx["contrast"] != 100:
            im = ImageEnhance.Contrast(im).enhance(ctx["contrast"] / 100)
        if ctx["saturation"] != 100:
            im = ImageEnhance.Color(im).enhance(ctx["saturation"] / 100)
        g = float(ctx["gamma"])
        if abs(g - 1) > 0.01:
            lut = [min(255, int(255 * (i / 255) ** (1 / g) + 0.5)) for i in range(256)] * 3
            im = im.point(lut)
        if ctx["sharpen"]:
            im = im.filter(ImageFilter.UnsharpMask(radius=1.5, percent=120, threshold=3))
        return im

    def one(path, i):
        src = open_doc(path, ctx["password"])
        pages = set(parse_pages(ctx["pages"], src.page_count))
        out = pymupdf.open()
        for pg in src:
            if pg.number not in pages:
                out.insert_pdf(src, from_page=pg.number, to_page=pg.number)
                continue
            ctx.progress((i + pg.number / src.page_count) / len(ctx.files), f"{os.path.basename(path)} - trang {pg.number + 1}")
            pix = pg.get_pixmap(dpi=int(ctx["dpi"]), alpha=False)
            buf = io.BytesIO()
            eff(pix.pil_image()).save(buf, "JPEG", quality=int(ctx["quality"]))
            np_ = out.new_page(width=pg.rect.width, height=pg.rect.height)
            np_.insert_image(np_.rect, stream=buf.getvalue())
        save_doc(out, ctx.out(path, "_chinh_mau"))
    ctx.each(one)


_COLOR_OPS = re.compile(rb"(?<![\w/])(?:[-+]?(?:\d+\.?\d*|\.\d+)\s+){1,4}(?:rg|RG|g|G|k|K|sc|SC)(?![\w])"
                        rb"|(?<![\w/])(?:(?:[-+]?(?:\d+\.?\d*|\.\d+)\s+){0,4}/[^\s/\[\]()<>{}%]+\s+)(?:scn|SCN)(?![\w])"
                        rb"|(?<![\w/])(?:[-+]?(?:\d+\.?\d*|\.\d+)\s+){1,4}(?:scn|SCN)(?![\w])"
                        rb"|/[^\s/\[\]()<>{}%]+\s+(?:cs|CS)(?![\w])")
_BT_ET = re.compile(rb"(?<![\w])BT(?![\w]).*?(?<![\w])ET(?![\w])", re.S)
_STRINGS = re.compile(rb"\((?:\\.|[^\\()]|\((?:\\.|[^\\()])*\))*\)|<[0-9A-Fa-f\s]*>", re.S)


def _recolor_text(stream, rgb_):
    """Đổi màu mọi khối chữ (BT...ET) trong 1 content stream: bỏ lệnh màu bên trong, đặt màu mới ở đầu khối,
    bọc q/Q để màu không lan sang hình vẽ phía sau."""
    col = (" ".join(f"{c:.3f}" for c in rgb_)).encode()
    head = b"q BT /DeviceRGB cs /DeviceRGB CS " + col + b" rg " + col + b" RG "

    def blk(m):
        body = m.group(0)[2:-2]
        # không đụng vào chuỗi chữ: tách chuỗi ra trước khi thay lệnh màu
        parts, last = [], 0
        for s in _STRINGS.finditer(body):
            parts.append(_COLOR_OPS.sub(b" ", body[last:s.start()]))
            parts.append(s.group(0))
            last = s.end()
        parts.append(_COLOR_OPS.sub(b" ", body[last:]))
        return head + b"".join(parts) + b" ET Q"
    return _BT_ET.sub(blk, stream)


@tool(id="text_color", group="stamp", name="Đổi màu chữ", icon=0xE8D3,
      keywords="text color mau chu den in",
      desc="Đổi màu toàn bộ chữ trong PDF (vd chữ xanh / xám nhạt → đen đậm để in rõ). Hình ảnh giữ nguyên.",
      opts=[Opt("color", "Màu chữ mới", "color", "#000000"), pages_opt(), PASSWORD_OPT])
def t_text_color(ctx):
    col = rgb(ctx["color"])

    def fn(doc, path):
        done = set()
        for p in parse_pages(ctx["pages"], doc.page_count):
            pg = doc[p]
            pg.clean_contents()
            xrefs = list(pg.get_contents())
            xrefs += [x[0] for x in pg.get_xobjects()]           # form XObject (chữ trong khối nhúng)
            for x in xrefs:
                if x in done:
                    continue
                done.add(x)
                try:
                    data = doc.xref_stream(x)
                except Exception:
                    continue
                if data and b"BT" in data:
                    doc.update_stream(x, _recolor_text(data, col))
    simple_page_tool(ctx, "_mau_chu", fn)


@tool(id="font_to_outline", group="secure", name="Chuyển chữ thành nét vẽ", icon=0xE8D2,
      keywords="font outline curves text to path in an",
      desc="Biến chữ thành đường vẽ (vector): hiển thị / in đúng mọi máy, không còn cần font, không copy chữ được.",
      opts=[pages_opt(), PASSWORD_OPT],
      note="Dùng khi gửi file in ấn hoặc không muốn người nhận chép / sửa chữ. Liên kết sẽ bị bỏ.")
def t_outline(ctx):
    def one(path, fi):
        doc = open_doc(path, ctx["password"])
        pages = set(parse_pages(ctx["pages"], doc.page_count))
        out = pymupdf.open()
        for pg in doc:
            if pg.number not in pages:
                out.insert_pdf(doc, from_page=pg.number, to_page=pg.number)
                continue
            ctx.progress((fi + pg.number / doc.page_count) / len(ctx.files),
                         f"{os.path.basename(path)} - trang {pg.number + 1}")
            svg = pg.get_svg_image(text_as_path=True)
            tmp = pymupdf.open("svg", svg.encode("utf-8"))
            one_page = pymupdf.open("pdf", tmp.convert_to_pdf())
            r = pg.rect
            np_ = out.new_page(width=r.width, height=r.height)
            np_.show_pdf_page(np_.rect, one_page, 0)
        save_doc(out, ctx.out(path, "_net_ve"))
    ctx.each(one)


def _layers_text(path):
    try:
        doc = pymupdf.open(path)
        import i18n
        on, off = ("show", "hide") if i18n.LANG == "en" else ("hiện", "ẩn")
        return "\n".join(f"{v['name']} | {on if v['on'] else off}" for v in doc.get_ocgs().values())
    except Exception:
        return ""


@tool(id="layers", group="secure", name="Lớp nội dung (layers)", icon=0xE81E,
      keywords="layers ocg lop an hien",
      desc="Xem các lớp (layer) trong PDF bản vẽ / thiết kế và đặt lớp nào hiện, lớp nào ẩn khi mở file.",
      opts=[Opt("items", "Các lớp", "textarea", "", prefill=_layers_text,
                tip="Mỗi dòng: tên lớp | hiện hoặc ẩn", placeholder="(file không có lớp)"),
            Opt("all_on", "Hiện tất cả các lớp", "bool", False), PASSWORD_OPT], max_files=1)
def t_layers(ctx):
    def fn(doc, path):
        ocgs = doc.get_ocgs()
        if not ocgs:
            raise ToolError("file không có lớp nội dung (layer)")
        want = {}
        for ln in (ctx["items"] or "").splitlines():
            name, _, state = ln.partition("|")
            if name.strip():
                want[name.strip()] = strip_vn(state.strip().lower()) not in ("an", "off", "hide", "0", "tat")
        on, off = [], []
        for x, v in ocgs.items():
            vis = True if ctx["all_on"] else want.get(v["name"], v["on"])
            (on if vis else off).append(x)
        doc.set_layer(-1, on=on, off=off)
        ctx.log(f"{os.path.basename(path)}: {len(on)} lớp hiện, {len(off)} lớp ẩn")
    simple_page_tool(ctx, "_lop", fn)


def strip_vn(s):
    import unicodedata
    s = s.replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


TSA_CHOICES = [("http://timestamp.digicert.com", "DigiCert"), ("http://timestamp.sectigo.com", "Sectigo"),
               ("http://time.certum.pl", "Certum")]


@tool(id="timestamp", group="secure", name="Gắn dấu thời gian (TSA)", icon=0xE823,
      keywords="timestamp tsa rfc3161 dau thoi gian",
      desc="Gắn dấu thời gian tin cậy (RFC 3161) chứng minh file tồn tại, không bị sửa từ thời điểm này. "
           "Không cần USB Token.",
      opts=[Opt("tsa", "Máy chủ dấu thời gian", "choice", TSA_CHOICES[0][0], TSA_CHOICES),
            PASSWORD_OPT],
      note="Cần Internet: chỉ gửi mã băm (SHA-256) của file tới máy chủ, không gửi nội dung file. "
           "Không sửa file sau khi gắn dấu.")
def t_timestamp(ctx):
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.sign import signers, timestamps

    def one(path, i):
        doc = open_doc(path, ctx["password"])
        if doc.is_encrypted or doc.needs_pass:
            raise ToolError("file có mật khẩu - mở khoá trước khi gắn dấu thời gian")
        data = doc.tobytes() if not doc.metadata.get("encryption") else None
        with open(path, "rb") as fh:
            raw = fh.read()
        try:
            w = IncrementalPdfFileWriter(io.BytesIO(raw))
        except Exception:
            w = IncrementalPdfFileWriter(io.BytesIO(data or doc.tobytes()))
        out = io.BytesIO()
        try:
            signers.PdfTimeStamper(timestamps.HTTPTimeStamper(ctx["tsa"], timeout=20)).timestamp_pdf(
                w, md_algorithm="sha256", output=out)
        except Exception as e:
            raise ToolError(f"không lấy được dấu thời gian từ máy chủ ({e})")
        dst = ctx.out(path, "_dau_thoi_gian")
        with open(dst, "wb") as fh:
            fh.write(out.getvalue())
        ctx.log(f"{os.path.basename(path)}: đã gắn dấu thời gian ({ctx['tsa']})")
    ctx.each(one)

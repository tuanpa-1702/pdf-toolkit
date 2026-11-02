# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Xử lý PDF cho PDF Toolkit (dựa trên PyMuPDF).

Tách riêng khỏi giao diện để có thể kiểm thử độc lập.
"""
import os
import re
from dataclasses import dataclass, field

import logging

import pymupdf

logging.getLogger("fontTools").setLevel(logging.CRITICAL)
log = logging.getLogger("pdf_core")

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"}

# ---------------------------------------------------------------------------
#  Font: dùng font Windows (Arial / Times / Courier) vì có đủ dấu tiếng Việt
# ---------------------------------------------------------------------------
FONT_DIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
FONT_FILES = {
    "sans":  ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    "serif": ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    "mono":  ("cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"),
}
FAMILY_NAMES = {"sans": "Arial", "serif": "Times New Roman", "mono": "Courier New"}

FLAG_ITALIC, FLAG_SERIF, FLAG_MONO, FLAG_BOLD = 2, 4, 8, 16


def font_family_of(fontname: str, flags: int) -> str:
    n = (fontname or "").lower()
    if flags & FLAG_MONO or any(k in n for k in ("courier", "mono", "consol")):
        return "mono"
    serif_names = ("times", "roman", "georgia", "garamond", "cambria", "minion", "palatino", "book antiqua")
    if any(k in n for k in serif_names) or (flags & FLAG_SERIF and "sans" not in n and "arial" not in n):
        return "serif"
    return "sans"


def is_bold(fontname: str, flags: int) -> bool:
    n = (fontname or "").lower()
    return bool(flags & FLAG_BOLD) or any(k in n for k in ("bold", "black", "heavy", "semibold", "demi"))


def is_italic(fontname: str, flags: int) -> bool:
    n = (fontname or "").lower()
    return bool(flags & FLAG_ITALIC) or "italic" in n or "oblique" in n


def font_file(family: str = "sans", bold: bool = False, italic: bool = False):
    files = FONT_FILES.get(family, FONT_FILES["sans"])
    path = os.path.join(FONT_DIR, files[(1 if bold else 0) + (2 if italic else 0)])
    if os.path.exists(path):
        return path
    fallback = os.path.join(FONT_DIR, "arial.ttf")
    return fallback if os.path.exists(fallback) else None


_font_cache = {}


def _font(path):
    if path not in _font_cache:
        _font_cache[path] = pymupdf.Font(fontfile=path) if path else pymupdf.Font("helv")
    return _font_cache[path]


def int_to_rgb(c: int):
    return ((c >> 16) & 255) / 255.0, ((c >> 8) & 255) / 255.0, (c & 255) / 255.0


# ---------------------------------------------------------------------------
#  Danh sách font Windows có đủ dấu tiếng Việt (cho trình soạn chữ định dạng)
# ---------------------------------------------------------------------------
USER_FONT_DIR = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")
_APP_DATA = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "PDFToolkit")
_FONT_CACHE_FILE = os.path.join(_APP_DATA, "fonts_v4.json")
# font tách từ file .ttc / bản tĩnh (Thường, Đậm) dựng từ font variable - tạo khi dùng lần đầu
FONT_EXTRACT_DIR = os.path.join(_APP_DATA, "fonts")
_VN_TEST = "ăâđêôơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợụủứừửữựỳỵỷỹĐ"
_fonts = None
_all_fonts = None


def _registry_font_files():
    import winreg
    files = set()
    key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    for root, base in ((winreg.HKEY_LOCAL_MACHINE, FONT_DIR), (winreg.HKEY_CURRENT_USER, USER_FONT_DIR)):
        try:
            k = winreg.OpenKey(root, key_path)
        except OSError:
            continue
        i = 0
        while True:
            try:
                _, val, _ = winreg.EnumValue(k, i)
            except OSError:
                break
            i += 1
            if isinstance(val, str) and val.lower().endswith((".ttf", ".otf", ".ttc")):
                files.add(val if os.path.isabs(val) else os.path.join(base, val))
        winreg.CloseKey(k)
    return sorted(f for f in files if os.path.exists(f))


def _face_info(f):
    """[(family, kiểu 'r'|'b'|'i'|'bi', đủ dấu tiếng Việt, là font variable)] của 1 mặt font, [] nếu không dùng được."""
    cmap = f.getBestCmap() or {}
    if not all(ord(c) in cmap for c in "AaZz09 "):           # font ký hiệu / chữ không phải Latin
        return []
    name = f["name"]
    # font variable: tên họ "thật" ở nameID 16 (vd Segoe UI Variable). Font tĩnh dùng nameID 1 để
    # Arial Narrow / Arial Black... không bị gộp vào Arial
    family = (name.getDebugName(16) if "fvar" in f else None) or name.getDebugName(1)
    if not family or family.startswith("."):
        return []
    vn = all(ord(c) in cmap for c in _VN_TEST)
    if "fvar" in f:                                          # variable: dựng bản Thường + Đậm khi dùng
        axes = {a.axisTag: (a.minValue, a.maxValue) for a in f["fvar"].axes}
        italic = bool(f["OS/2"].fsSelection & 0x01) if "OS/2" in f else False
        out = [(family, "i" if italic else "r", vn, True)]
        if "wght" in axes and axes["wght"][1] >= 600:
            out.append((family, "bi" if italic else "b", vn, True))
        return out
    sel = f["OS/2"].fsSelection if "OS/2" in f else 0
    sub = (name.getDebugName(17) or name.getDebugName(2) or "").lower()
    bold = bool(sel & 0x20) or "bold" in sub
    italic = bool(sel & 0x01) or "italic" in sub or "oblique" in sub
    return [(family, ("b" if bold else "") + ("i" if italic else "") or "r", vn, False)]


def _file_faces(path):
    """[(family, kiểu, đủ dấu TV, spec)] trong 1 file font. spec: đường dẫn, hoặc 'đường dẫn|chỉ số mặt|kiểu'
    với file .ttc / font variable (dựng ra file riêng khi dùng - xem _resolve_font)."""
    from fontTools.ttLib import TTCollection, TTFont
    if path.lower().endswith(".ttc"):
        coll = TTCollection(path, lazy=True)
        faces = list(enumerate(coll.fonts))
    else:
        faces = [(-1, TTFont(path, lazy=True))]
    out = []
    try:
        for idx, f in faces:
            for fam, style, vn, var in _face_info(f):
                spec = path if (idx < 0 and not var) else f"{path}|{max(idx, 0)}|{style if var else ''}"
                out.append((fam, style, vn, spec))
    finally:
        for _i, f in faces:
            f.close()
    return out


def _resolve_font(spec):
    """Đường dẫn file font dùng được cho MuPDF: tách mặt font khỏi .ttc, dựng bản tĩnh từ font variable."""
    if "|" not in spec:
        return spec
    import hashlib
    path, idx, style = spec.rsplit("|", 2)
    base = os.path.splitext(os.path.basename(path))[0]
    st = os.stat(path)
    tag = hashlib.md5(f"{spec}|{st.st_size}|{int(st.st_mtime)}".encode()).hexdigest()[:8]
    out = os.path.join(FONT_EXTRACT_DIR, f"ptk_{base}_{idx}{style}_{tag}.ttf")
    if os.path.exists(out):
        return out
    from fontTools.ttLib import TTFont
    f = TTFont(path, fontNumber=int(idx)) if path.lower().endswith(".ttc") else TTFont(path)
    if style:                                         # font variable -> bản tĩnh theo độ đậm
        from fontTools.varLib import instancer
        axes = {a.axisTag: a for a in f["fvar"].axes}
        loc = {}
        for tag_, a in axes.items():
            if tag_ == "wght":
                loc[tag_] = min(max(700 if "b" in style else 400, a.minValue), a.maxValue)
            else:
                loc[tag_] = a.defaultValue
        try:
            f = instancer.instantiateVariableFont(f, loc, inplace=False)
        except Exception:                             # vài font (Segoe UI Variable) lỗi bảng GPOS: bỏ kerning rồi dựng lại
            f.close()
            f = TTFont(path, fontNumber=int(idx)) if path.lower().endswith(".ttc") else TTFont(path)
            for t in ("GPOS", "GDEF"):
                if t in f:
                    del f[t]
            f = instancer.instantiateVariableFont(f, loc, inplace=False)
    os.makedirs(FONT_EXTRACT_DIR, exist_ok=True)
    tmp = out + ".tmp"
    f.save(tmp)
    os.replace(tmp, out)
    return out


class _Faces(dict):
    """{kiểu: spec} - trả về file font thật (tách / dựng khi lấy lần đầu)."""

    _bad = set()                                      # spec đã lỗi - không dựng lại mỗi lần

    def _get(self, k):
        spec = dict.__getitem__(self, k)
        if spec not in self._bad:
            try:
                return _resolve_font(spec)
            except Exception:
                log.exception("Không dựng được font %s", spec)
                self._bad.add(spec)
        return font_file("sans", "b" in k, "i" in k)

    def __getitem__(self, k):
        return self._get(k)

    def get(self, k, default=None):
        return self._get(k) if k in self else default

    def values(self):
        return [self._get(k) for k in self]

    def items(self):
        return [(k, self._get(k)) for k in self]


def _scan_fonts():
    """Quét font Windows 1 lần (có cache): {tên font: (_Faces, đủ dấu tiếng Việt)}."""
    import json
    cache = {}
    try:
        with open(_FONT_CACHE_FILE, encoding="utf-8") as fh:
            cache = json.load(fh)
    except Exception:
        pass
    new_cache, fams, vn_ok = {}, {}, {}
    for path in _registry_font_files():
        st = os.stat(path)
        key = f"{path}|{st.st_size}|{int(st.st_mtime)}"
        if key in cache:
            faces = cache[key]
        else:
            try:
                faces = _file_faces(path)
            except Exception:
                faces = []
        new_cache[key] = faces
        for fam, style, vn, spec in faces:
            fams.setdefault(fam, _Faces()).setdefault(style, spec)
            vn_ok[fam] = vn_ok.get(fam, True) and vn
    if new_cache != cache:
        try:
            os.makedirs(os.path.dirname(_FONT_CACHE_FILE), exist_ok=True)
            with open(_FONT_CACHE_FILE, "w", encoding="utf-8") as fh:
                json.dump(new_cache, fh)
        except OSError:
            pass
    for fam, faces in fams.items():                 # font chỉ có bản đậm / nghiêng: dùng tạm làm bản thường
        if "r" not in faces:
            dict.__setitem__(faces, "r", next(iter(dict.values(faces))))
    for k in FAMILY_NAMES.values():                 # luôn có 3 font cơ bản
        if k not in fams:
            fam = next(f for f, n in FAMILY_NAMES.items() if n == k)
            files = [os.path.join(FONT_DIR, x) for x in FONT_FILES[fam]]
            if os.path.exists(files[0]):
                fams[k] = _Faces({s: p for s, p in zip(("r", "b", "i", "bi"), files) if os.path.exists(p)})
                vn_ok[k] = True
    order = sorted(fams, key=str.lower)
    return {k: (fams[k], vn_ok.get(k, False)) for k in order}


def all_fonts():
    """Mọi font Windows dùng được: {tên: {'r','b','i','bi': file}} (kể cả font thiếu dấu tiếng Việt)."""
    global _all_fonts, _fonts
    if _all_fonts is None:
        scan = _scan_fonts()
        _all_fonts = {k: v for k, (v, _vn) in scan.items()}
        _fonts = {k: v for k, (v, vn) in scan.items() if vn}
    return _all_fonts


def vn_fonts():
    """Font có đủ dấu tiếng Việt: {tên font: {'r': file, 'b': file, 'i': file, 'bi': file}}."""
    all_fonts()
    return _fonts


def _norm_name(s):
    return re.sub(r"[^a-z0-9]", "", s.lower())


def guess_family(pdf_font: str, flags: int = 0) -> str:
    """Tìm font Windows tương ứng với tên font trong PDF (vd 'ABCDEF+TimesNewRomanPS-BoldMT')."""
    fonts = vn_fonts()
    n = pdf_font.split("+", 1)[-1]
    n = re.split(r"[-,]", n)[0]
    n = _norm_name(n)
    for suf in ("psmt", "mt", "ps", "bolditalic", "bold", "italic", "oblique", "regular", "semibold", "light"):
        if n.endswith(suf) and len(n) > len(suf) + 2:
            n = n[: -len(suf)]
    if n:
        names = {_norm_name(k): k for k in fonts}
        if n in names:
            return names[n]
        for nk, k in names.items():
            if nk.startswith(n) or n.startswith(nk):
                return k
    return FAMILY_NAMES[font_family_of(pdf_font, flags)]


def font_css(families):
    """CSS @font-face cho các font đã dùng (MuPDF tự nhúng, tự rút gọn font khi lưu)."""
    fonts = all_fonts()
    out = []
    for fam in sorted(set(families)):
        files = fonts.get(fam)
        if not files:
            continue
        reg = files.get("r") or next(iter(files.values()))
        for style, weight, ital in (("r", "normal", "normal"), ("b", "bold", "normal"),
                                    ("i", "normal", "italic"), ("bi", "bold", "italic")):
            path = files.get(style)
            if not path:
                continue
            out.append("@font-face {font-family: '%s'; src: url('%s'); font-weight: %s; font-style: %s;}"
                       % (fam.replace("'", ""), os.path.basename(path), weight, ital))
        if "r" not in files:
            out.append("@font-face {font-family: '%s'; src: url('%s');}" % (fam.replace("'", ""), os.path.basename(reg)))
    return "\n".join(out)


def _archive():
    arch = pymupdf.Archive(FONT_DIR)
    try:
        os.makedirs(FONT_EXTRACT_DIR, exist_ok=True)   # thư mục đọc theo yêu cầu: font dựng sau vẫn thấy
    except OSError:
        pass
    for d in (USER_FONT_DIR, FONT_EXTRACT_DIR):
        if os.path.isdir(d):
            arch.add(d)
    return arch


BASE_CSS = "p {margin: 0; padding: 0; line-height: 1.15;} body {margin: 0; padding: 0;}"


def _anchor_text(page, words, rects, near=False):
    """Chữ trên trang nằm dưới các vùng (tô sáng / gạch) hoặc dòng chữ gần vùng (ghi chú)."""
    picked = []
    for r in rects:
        r = pymupdf.Rect(r)
        if near:
            r = pymupdf.Rect(r.x0 - 4, r.y0 - 2, page.rect.x1, r.y1 + 2)
        for w in words:
            wr = pymupdf.Rect(w[:4])
            inter = wr & r
            if not inter.is_empty and inter.get_area() >= 0.5 * wr.get_area() and w not in picked:
                picked.append(w)
        if near and picked:
            break
    picked.sort(key=lambda w: (w[5], w[6], w[7]))          # block, line, word
    return " ".join(w[4] for w in picked)[:200]


def export_comments(pdf_path, out_json):
    """Ghi nhận xét / tô sáng / gạch xoá trong PDF ra JSON (add-in Word dùng để nhập nhận xét). Trả mã thoát."""
    import json
    kinds = {pymupdf.PDF_ANNOT_HIGHLIGHT: "highlight", pymupdf.PDF_ANNOT_UNDERLINE: "underline",
             pymupdf.PDF_ANNOT_SQUIGGLY: "underline", pymupdf.PDF_ANNOT_STRIKE_OUT: "strikeout",
             pymupdf.PDF_ANNOT_TEXT: "note", pymupdf.PDF_ANNOT_FREE_TEXT: "text", pymupdf.PDF_ANNOT_CARET: "insert"}
    out = []
    try:
        doc = pymupdf.open(pdf_path)
        for page in doc:
            words = page.get_text("words")
            annots = list(page.annots() or [])
            by_xref = {a.xref: a for a in annots}
            replies = {}
            for a in annots:
                irt = getattr(a, "irt_xref", 0)
                if irt and irt in by_xref:
                    replies.setdefault(irt, []).append(a)
            for a in annots:
                if getattr(a, "irt_xref", 0) in by_xref:
                    continue                                  # trả lời -> gộp vào nhận xét gốc
                t = a.type[0]
                if t in (pymupdf.PDF_ANNOT_POPUP, pymupdf.PDF_ANNOT_LINK, pymupdf.PDF_ANNOT_WIDGET):
                    continue
                info = a.info
                content = (info.get("content") or "").strip()
                kind = kinds.get(t, "shape")
                if kind in ("highlight", "underline", "strikeout"):
                    v = a.vertices or []
                    rects = [pymupdf.Quad(v[i:i + 4]).rect for i in range(0, len(v) - 3, 4)] or [a.rect]
                    anchor = _anchor_text(page, words, rects)
                else:
                    anchor = _anchor_text(page, words, [a.rect], near=True)
                if not content and kind not in ("highlight", "underline", "strikeout"):
                    continue
                reps = [{"author": (r.info.get("title") or ""), "content": (r.info.get("content") or "").strip()}
                        for r in replies.get(a.xref, [])]
                out.append({"page": page.number + 1, "kind": kind, "author": info.get("title") or "",
                            "content": content, "anchor": anchor, "date": info.get("modDate") or "",
                            "replies": reps})
        with open(out_json, "w", encoding="utf-8") as fh:
            json.dump({"ok": True, "comments": out}, fh, ensure_ascii=False)
        return 0
    except Exception as e:
        with open(out_json, "w", encoding="utf-8") as fh:
            json.dump({"ok": False, "error": str(e)}, fh, ensure_ascii=False)
        return 1


def disp_to_unrot(page, rect):
    """Vùng theo toạ độ hiển thị (đã xoay) -> toạ độ trang gốc."""
    r = pymupdf.Rect(rect) * page.derotation_matrix
    r.normalize()
    return r


def _merge_line_rects(rects):
    """Gộp ô các từ liền nhau trên cùng dòng thành 1 ô (gạch chân / tô sáng liền mạch, không đứt ở dấu cách)."""
    out = []
    for r in rects:
        r = pymupdf.Rect(r)
        if out:
            last = out[-1]
            same_line = abs(last.y0 - r.y0) < 0.4 * r.height and abs(last.y1 - r.y1) < 0.4 * r.height
            if same_line and -2 <= r.x0 - last.x1 <= max(r.height, 6):
                out[-1] = last | r
                continue
        out.append(r)
    return out


# ---------------------------------------------------------------------------
#  Dòng chữ trên trang
# ---------------------------------------------------------------------------
@dataclass
class TextLine:
    bbox: pymupdf.Rect
    text: str
    size: float
    font: str
    flags: int
    color: int
    baseline: float
    uniform: bool                       # cả dòng cùng một kiểu chữ
    spans: list = field(default_factory=list)
    chars: list = field(default_factory=list)   # [(bbox, span)] tương ứng từng ký tự của `text`
    block: int = -1                     # số thứ tự khối chữ (đoạn văn) chứa dòng - để gom dòng thành khối khi sửa

    def match_rect(self, start, end):
        r = pymupdf.Rect(self.chars[start][0])
        for bb, _ in self.chars[start + 1:end]:
            r |= bb
        return r

    @property
    def family(self):
        return font_family_of(self.font, self.flags)

    @property
    def bold(self):
        return is_bold(self.font, self.flags)

    @property
    def italic(self):
        return is_italic(self.font, self.flags)


def _style_key(s):
    return (s["font"], round(s["size"], 1), s["color"], s["flags"] & (FLAG_BOLD | FLAG_ITALIC))


# Chuẩn hoá ký tự hay gặp trong PDF (dấu cách cứng, gạch nối mềm...) để tìm kiếm khớp với chữ gõ vào
_CHAR_MAP = str.maketrans({"\xa0": " ", " ": " ", " ": " ", " ": " ", "\xad": "-",
                           "‐": "-", "‑": "-", "–": "-", "−": "-", "\t": " "})


def normalize(s: str) -> str:
    return s.translate(_CHAR_MAP)


def get_text_lines(page) -> list:
    """Các dòng chữ nằm ngang trên trang (toạ độ chưa xoay)."""
    out = []
    flags = pymupdf.TEXT_PRESERVE_WHITESPACE | pymupdf.TEXT_MEDIABOX_CLIP
    d = page.get_text("rawdict", flags=flags)
    for bi, b in enumerate(d["blocks"]):
        if b.get("type", 0) != 0:
            continue
        for ln in b["lines"]:
            if abs(ln["dir"][0] - 1) > 1e-3:      # bỏ qua chữ xoay/dọc
                continue
            chars = []
            for s in ln["spans"]:
                for ch in s["chars"]:
                    chars.append((normalize(ch["c"]), pymupdf.Rect(ch["bbox"]), s))
            # bỏ khoảng trắng đầu/cuối
            while chars and not chars[0][0].strip():
                chars.pop(0)
            while chars and not chars[-1][0].strip():
                chars.pop()
            if not chars:
                continue
            spans = []
            for _, _, s in chars:
                if not spans or spans[-1] is not s:
                    spans.append(s)
            spans = [s for s in spans if any(ch["c"].strip() for ch in s["chars"])]
            for s in spans:
                s["text"] = normalize("".join(ch["c"] for ch in s["chars"]))
            main = max(spans, key=lambda s: len(s["text"].strip()))
            bbox = pymupdf.Rect(spans[0]["bbox"])
            for s in spans[1:]:
                bbox |= pymupdf.Rect(s["bbox"])
            # mỗi ký tự của `text` ứng với đúng một phần tử của `chars`
            text = "".join((c or " ")[0] for c, _, _ in chars)
            out.append(TextLine(
                bbox=bbox, text=text, size=main["size"], font=main["font"], flags=main["flags"],
                color=main["color"], baseline=main["origin"][1],
                uniform=len({_style_key(s) for s in spans}) == 1, spans=spans,
                chars=[(bb, s) for _, bb, s in chars], block=bi))
    return out


def line_at(lines, x, y):
    """Dòng chứa điểm (x, y) - toạ độ chưa xoay."""
    best = None
    for i, ln in enumerate(lines):
        r = ln.bbox
        if r.x0 - 1 <= x <= r.x1 + 1 and r.y0 - 1 <= y <= r.y1 + 1:
            if best is None or r.get_area() < lines[best].bbox.get_area():
                best = i
    return best


def _redact(page, rects):
    """Xoá chữ trong các vùng, giữ nguyên ảnh, đường kẻ và nền."""
    for r in rects:
        r = pymupdf.Rect(r)
        h = r.height
        dx = 0.4 if r.width > 3 else 0
        rr = pymupdf.Rect(r.x0 + dx, r.y0 + h * 0.25, r.x1 - dx, r.y1 - h * 0.25)
        page.add_redact_annot(rr, fill=False)
    page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                          graphics=pymupdf.PDF_REDACT_LINE_ART_NONE,
                          text=pymupdf.PDF_REDACT_TEXT_REMOVE)


def _write(page, x, baseline, text, size, family, bold, italic, color, max_width=None, rotate=0):
    path = font_file(family, bold, italic)
    font = _font(path)
    if max_width:
        w = font.text_length(text, fontsize=size)
        if w > max_width > 0:
            size = max(size * 0.6, size * max_width / w)
    fname = "TK" + re.sub(r"\W", "", os.path.basename(path or "helv").split(".")[0])
    kw = dict(fontsize=size, fontname=fname, color=color, rotate=rotate)
    if path:
        kw["fontfile"] = path
    page.insert_text(pymupdf.Point(x, baseline), text, **kw)
    return size


# ---------------------------------------------------------------------------
#  Tài liệu
# ---------------------------------------------------------------------------
class PdfError(Exception):
    pass


# ---------------------------------------------------------------------------
#  Xoá vĩnh viễn thông tin (redaction): đánh dấu vùng -> áp dụng = xoá hẳn chữ / điểm ảnh / hình vẽ bên dưới
# ---------------------------------------------------------------------------
REDACT_PRESETS = {
    "email": ("Địa chỉ email", r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+"),
    "phone": ("Số điện thoại", r"(?<![\d.])(?:\+?84|0)(?:[\s.-]?\d){8,10}(?![\d.])"),
    "id": ("Số CCCD / CMND (9 hoặc 12 số)", r"(?<!\d)(?:\d{12}|\d{9})(?!\d)"),
    "account": ("Số tài khoản / dãy số dài (10-19 số, có thể cách nhau)", r"(?<!\d)\d(?:[\s-]?\d){9,18}(?!\d)"),
}


def find_redact_rects(page, terms=(), patterns=(), whole_word=False):
    """Vùng (toạ độ trang chưa xoay) chứa các chữ cần xoá: terms = cụm chữ (không phân biệt hoa thường),
    patterns = biểu thức chính quy. Ghép từ trong cùng dòng để tìm được cả cụm có dấu cách (vd '0912 345 678')."""
    rects = []
    words = page.get_text("words", sort=True)            # x0 y0 x1 y1 text block line wno
    lines = {}
    for w in words:
        lines.setdefault((w[5], w[6]), []).append(w)
    regs = [re.compile(p) for p in patterns if p]
    for t in terms:
        t = t.strip()
        if t:
            regs.append(re.compile((r"(?<!\w)" + re.escape(t) + r"(?!\w)") if whole_word else re.escape(t),
                                   re.IGNORECASE))
    if not regs:
        return rects
    for ws in lines.values():
        s, spans = "", []
        for w in ws:
            if s:
                s += " "
            spans.append((len(s), len(s) + len(w[4]), w))
            s += w[4]
        for rg in regs:
            for m in rg.finditer(s):
                a, b = m.span()
                if a == b:
                    continue
                hit = [w for (x, y, w) in spans if x < b and y > a]
                if not hit:
                    continue
                r = pymupdf.Rect(hit[0][:4])
                for w in hit[1:]:
                    r |= pymupdf.Rect(w[:4])
                rects.append(r)                           # xoá cả từ chứa phần khớp: thừa còn hơn sót ký tự
    return rects


def redact_fill(color):
    return {"black": (0, 0, 0), "white": (1, 1, 1)}.get(color, (0, 0, 0))


# ---------------------------------------------------------------------------
#  Tự nhận diện ô form: chỗ trống để điền (gạch dưới / chấm chấm, đường kẻ trống, ô vuông, khung trống)
# ---------------------------------------------------------------------------
_CHECK_GLYPHS = "☐□❑❒▢"


def _field_label(words, x0, y0, y1):
    """Nhãn bên trái chỗ trống (cùng dòng): 'Họ và tên:' -> 'Họ và tên'."""
    row = [w for w in words if w[2] <= x0 + 2 and min(w[3], y1) - max(w[1], y0) > 0.4 * (w[3] - w[1])]
    row.sort(key=lambda w: w[0])
    s = " ".join(w[4] for w in row[-6:])
    s = re.split(r"[_.…]{3,}", s)[-1] if re.search(r"[_.…]{3,}", s) else s
    s = re.sub(r"[_.…:;,\s]+$", "", s).strip(" -:()[]")
    return s[-40:].strip()


def _check_label(words, r):
    """Nhãn của ô đánh dấu: chữ ngay bên phải ô ('☐ Nam' -> 'Nam'), không có thì lấy bên trái."""
    row = []
    for w in words:
        if min(w[3], r.y1) - max(w[1], r.y0) <= 0.3 * (w[3] - w[1]):
            continue
        t = w[4]
        if w[0] <= r.x0 + 1 and t[:1] in _CHECK_GLYPHS:       # '☐Đồng' (không có dấu cách sau ô)
            t = t.lstrip(_CHECK_GLYPHS)
            if t:
                row.append((r.x1, w[1], w[2], w[3], t))
        elif w[0] >= r.x1 - 1 and w[0] - r.x1 < 60 and t not in _CHECK_GLYPHS:
            row.append(w)
    row.sort(key=lambda w: w[0])
    s, last = [], r.x1
    for w in row[:4]:
        if w[0] - last > 12:                           # khoảng trống lớn: sang ô / cột khác
            break
        s.append(w[4]); last = w[2]
    t = re.sub(r"[_.…:;,]+$", "", " ".join(s)).strip()
    return t or _field_label(words, r.x0, r.y0, r.y1)


def detect_form_fields(page):
    """Chỗ trống có thể điền trên trang: [dict(rect=Rect toạ độ trang chưa xoay, kind='text'|'multiline'|'check',
    label=str)]. Bỏ qua vùng đã có ô form."""
    words = page.get_text("words", sort=True)
    out = []

    def add(r, kind, label=""):
        r = pymupdf.Rect(r)
        if r.is_empty or r.width < 4:
            return
        for o in out:                                   # trùng vùng đã nhận diện
            inter = (o["rect"] & r)
            if not inter.is_empty and abs(inter) > 0.5 * min(abs(o["rect"]), abs(r)):
                return
        out.append(dict(rect=r, kind=kind, label=label))

    def has_text(r, pad=0):
        rr = pymupdf.Rect(r) + (pad, pad, -pad, -pad)
        for w in words:
            wr = pymupdf.Rect(w[:4])
            if (wr & rr).is_valid and abs(wr & rr) > 0.3 * abs(wr) and not re.fullmatch(r"[_.…" + _CHECK_GLYPHS + "]+", w[4]):
                return True
        return False

    # 1) chuỗi gạch dưới / chấm chấm trong chữ (theo từng ký tự)
    raw = page.get_text("rawdict")
    for b in raw["blocks"]:
        for ln in b.get("lines", []):
            if abs(ln["dir"][1]) > 0.05:
                continue
            chars = [c for sp in ln["spans"] for c in sp["chars"]]
            i = 0
            while i < len(chars):
                ch = chars[i]["c"]
                if ch in "_.…" or ch in _CHECK_GLYPHS:
                    j = i
                    while j + 1 < len(chars) and chars[j + 1]["c"] == ch:
                        j += 1
                    n = j - i + 1
                    if ch in _CHECK_GLYPHS:
                        for k in range(i, j + 1):
                            cb = pymupdf.Rect(chars[k]["bbox"])
                            s = max(cb.width, cb.height * 0.7)
                            c = (cb.tl + cb.br) / 2
                            add(pymupdf.Rect(c.x - s / 2, c.y - s / 2, c.x + s / 2, c.y + s / 2), "check",
                                _check_label(words, cb))
                    elif (ch == "_" and n >= 4) or (ch == "." and n >= 5) or (ch == "…" and n >= 2):
                        r = pymupdf.Rect(chars[i]["bbox"]) | pymupdf.Rect(chars[j]["bbox"])
                        h = max(14.0, min(22.0, r.height * 1.25))
                        add(pymupdf.Rect(r.x0, r.y1 - h, r.x1, r.y1 + 1), "text", _field_label(words, r.x0, r.y0, r.y1))
                    i = j + 1
                else:
                    i += 1

    # 2) hình vẽ: đường kẻ ngang trống, ô vuông nhỏ, khung trống
    lines, boxes = [], []
    for d in page.get_drawings():
        for it in d["items"]:
            if it[0] == "l":
                p1, p2 = it[1], it[2]
                if abs(p1.y - p2.y) < 0.6 and abs(p1.x - p2.x) >= 40:
                    lines.append(pymupdf.Rect(min(p1.x, p2.x), p1.y, max(p1.x, p2.x), p1.y))
            elif it[0] == "re":
                r = pymupdf.Rect(it[1])
                if r.height <= 1.6 and r.width >= 40:                  # đường kẻ vẽ bằng hình chữ nhật mảnh
                    lines.append(pymupdf.Rect(r.x0, r.y1, r.x1, r.y1))
                else:
                    boxes.append(r)
            elif it[0] == "qu":
                boxes.append(it[1].rect)
    for r in boxes:
        w, h = r.width, r.height
        if 5.5 <= w <= 18 and 5.5 <= h <= 18 and abs(w - h) <= 2.5:
            if not has_text(r, 1):
                add(r, "check", _check_label(words, r))
        elif w >= 40 and 11 <= h <= 90 and not has_text(r, 1):
            inner = any(o != r and r.contains(o) and o.width < w - 2 for o in boxes)
            if not inner:
                add(r + (1.5, 1.5, -1.5, -1.5), "multiline" if h > 32 else "text",
                    _field_label(words, r.x0, r.y0, r.y1))
    for l in lines:
        area = pymupdf.Rect(l.x0, l.y0 - 16, l.x1, l.y0 - 0.5)
        if has_text(area):
            continue                                   # chữ nằm trên đường kẻ: gạch chân / kẻ bảng có nội dung
        if any(abs(b.y0 - l.y0) < 1 or abs(b.y1 - l.y0) < 1 for b in boxes if b.x0 <= l.x0 + 1 and b.x1 >= l.x1 - 1):
            continue                                   # cạnh của khung
        add(pymupdf.Rect(l.x0, l.y0 - 15, l.x1, l.y0 - 0.5), "text", _field_label(words, l.x0, l.y0 - 14, l.y0))

    # bỏ vùng đã có ô form
    have = [w.rect for w in page.widgets() or []]
    return [f for f in out if not any(abs(f["rect"] & h) > 0.3 * abs(f["rect"]) for h in have if (f["rect"] & h).is_valid)]


def field_name_from_label(label, used, n):
    s = re.sub(r"\s+", " ", label or "").strip()
    s = re.sub(r"[^\w ]", "", s).strip()[:30]
    base = s or f"O_{n}"
    name, k = base, 2
    while name in used:
        name = f"{base}_{k}"
        k += 1
    used.add(name)
    return name


def scan_images_doc(images, dpi):
    """Ảnh quét (đường dẫn) -> PDF, khổ trang đúng kích thước thật theo dpi. Ảnh màu / xám nén JPEG (file nhỏ),
    ảnh đen trắng giữ PNG (nét chữ)."""
    doc = pymupdf.open()
    for path in images:
        pix = pymupdf.Pixmap(path)
        if pix.alpha:
            pix = pymupdf.Pixmap(pix, 0)
        if pix.colorspace and pix.colorspace.n not in (1, 3):
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        bw = pix.n == 1 and len(set(pix.samples[::97])) <= 2      # chỉ 2 mức (đen / trắng)
        data = pix.tobytes("png") if bw else pix.tobytes("jpg", jpg_quality=85)
        page = doc.new_page(width=pix.width * 72 / dpi, height=pix.height * 72 / dpi)
        page.insert_image(page.rect, stream=data)
    return doc


def apply_redactions_doc(doc):
    """Áp dụng mọi vùng đánh dấu xoá trong file. Trả về số vùng đã xoá."""
    n = 0
    for page in doc:
        k = sum(1 for _ in (page.annots(types=[pymupdf.PDF_ANNOT_REDACT]) or []))
        if k:
            page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_PIXELS,
                                  graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,   # giữ kẻ bảng đi qua
                                  text=pymupdf.PDF_REDACT_TEXT_REMOVE)
            n += k
    return n


# ---------------------------------------------------------------------------
#  Tách nền ảnh con dấu (chèn như tem dán: nền trong suốt)
# ---------------------------------------------------------------------------
STAMP_INKS = {"red": 0.0, "blue": 220.0, "purple": 280.0, "green": 130.0}


def _hue_sat(rgb):
    """rgb uint8 (h, w, 3) -> (hue độ 0-360, độ bão hoà 0-1, độ sáng 0-255)."""
    import numpy as np
    f = rgb.astype(np.float32)
    mx, mn = f.max(axis=2), f.min(axis=2)
    d = mx - mn
    sat = np.where(mx > 0, d / np.maximum(mx, 1), 0)
    r, g, b = f[..., 0], f[..., 1], f[..., 2]
    dd = np.maximum(d, 1e-6)
    h = np.where(mx == r, ((g - b) / dd) % 6, np.where(mx == g, (b - r) / dd + 2, (r - g) / dd + 4)) * 60
    return h, sat, mx


def detect_stamp_ink(rgb):
    """Đoán màu mực con dấu: 'red' / 'blue' / 'purple' / 'green', hoặc 'dark' nếu ảnh gần như không có màu."""
    import numpy as np
    h, s, v = _hue_sat(rgb)
    colored = (s > 0.3) & (v > 50)
    if colored.mean() < 0.004:
        return "dark"
    hist = np.histogram(h[colored], bins=36, range=(0, 360))[0].astype(float)
    hist = hist + np.roll(hist, 1) + np.roll(hist, -1)            # gộp bin kề (đỏ nằm 2 phía 0°/360°)
    peak = (int(hist.argmax()) * 10 + 5) % 360
    return min(STAMP_INKS, key=lambda k: min(abs(peak - STAMP_INKS[k]), 360 - abs(peak - STAMP_INKS[k])))


def extract_stamp(rgb, ink="auto", uniform=True, crop=True, despeckle=True, max_side=1400):
    """Tách con dấu khỏi nền -> RGBA (numpy uint8).
    ink: auto | red | blue | purple | green | dark | white (chỉ xoá nền trắng).
    - Mực màu: giữ điểm ảnh cùng sắc độ với mực -> bỏ được chữ đen, chữ ký, kẻ bảng, nền giấy ngả vàng / hoa văn.
    - dark: con dấu đen -> độ trong suốt theo độ tối so với màu giấy.
    uniform: tô lại 1 màu mực (đậm, tươi như dấu mới); crop: cắt sát con dấu; despeckle: bỏ chấm lẻ."""
    import numpy as np
    H, W = rgb.shape[:2]
    if max(H, W) > max_side:                                     # ảnh scan quá lớn -> thu nhỏ cho nhanh
        k = max_side / max(H, W)
        pm = pymupdf.Pixmap(pymupdf.csRGB, W, H, np.ascontiguousarray(rgb).tobytes(), 0)
        pm = pymupdf.Pixmap(pm, int(W * k), int(H * k), None)
        rgb = np.frombuffer(pm.samples, np.uint8).reshape(pm.height, pm.width, pm.n)[..., :3]
        H, W = rgb.shape[:2]
    if ink == "auto":
        ink = detect_stamp_ink(rgb)
    f = rgb.astype(np.float32)
    lum = f @ np.array([0.299, 0.587, 0.114], np.float32)
    border = np.concatenate([lum[0], lum[-1], lum[:, 0], lum[:, -1]])
    paper = float(np.percentile(border, 75)) if border.size else 255.0
    if ink in STAMP_INKS:
        h, s, v = _hue_sat(rgb)
        dh = np.abs(h - STAMP_INKS[ink])
        dh = np.minimum(dh, 360 - dh)
        w_hue = np.clip(1 - (dh - 25) / 15, 0, 1)                   # trong ±25° = 1, ngoài ±40° = 0 (giấy vàng ~45°)
        w_sat = np.clip((s - 0.12) / 0.28, 0, 1)
        w_val = np.clip((v - 35) / 40, 0, 1)                        # quá tối (gần đen) -> không phải mực dấu
        alpha = w_hue * w_sat * w_val
    elif ink == "dark":
        alpha = np.clip((paper - lum - 30) / 90, 0, 1)
    else:                                                          # white: chỉ xoá nền trắng / giấy
        alpha = np.clip((paper - 12 - lum) / 70, 0, 1)
    a8 = (alpha * 255).astype(np.uint8)
    a8[a8 < 45] = 0                                                # vệt mờ của nền giấy / nhiễu scan
    if despeckle:
        import imgops
        a8 = imgops.median3(a8)
        # bỏ đốm nhỏ không dính với con dấu (bụi scan, dấu chấm của chữ màu)
        n, lab, areas = imgops.label_components(a8 > 60)
        if n > 1:
            keep = np.zeros(n, bool)
            keep[1:] = areas[1:] >= max(12, areas[1:].max() * 0.002)
            a8 = np.where(keep[lab] | (a8 <= 60), a8, 0).astype(np.uint8)
    out = np.zeros((H, W, 4), np.uint8)
    strong = a8 > 150
    if uniform and strong.any():
        col = np.median(rgb[strong].reshape(-1, 3), axis=0)
        if ink in STAMP_INKS:                                      # mực tươi hơn bản scan (scan làm nhạt màu)
            col = np.clip((col - col.min()) * 1.25 + col.min() * 0.6, 0, 255)
        else:
            col = np.minimum(col, 60)
        out[..., :3] = col.astype(np.uint8)
        a8 = np.clip(a8.astype(np.float32) * 1.35, 0, 255).astype(np.uint8)   # nét đậm, liền hơn
    else:
        out[..., :3] = rgb
    out[..., 3] = a8
    if crop:
        ys, xs = np.nonzero(a8 > 90)
        if len(xs):
            pad = 6
            y0, y1 = max(0, ys.min() - pad), min(H, ys.max() + 1 + pad)
            x0, x1 = max(0, xs.min() - pad), min(W, xs.max() + 1 + pad)
            out = out[y0:y1, x0:x1]
    return np.ascontiguousarray(out), ink


def rgba_to_png(rgba):
    import numpy as np
    h, w = rgba.shape[:2]
    pm = pymupdf.Pixmap(pymupdf.csRGB, w, h, np.ascontiguousarray(rgba).tobytes(), 1)
    return pm.tobytes("png")


class PdfDocument:
    MAX_UNDO = 25
    MAX_UNDO_BYTES = 600 * 1024 * 1024

    def __init__(self):
        self.doc = None
        self.path = None
        self.file_size = 0
        self.modified = False
        self.new_fonts = False
        self._undo = []
        self._redo = []

    # ----------------------------- mở / lưu -----------------------------
    @staticmethod
    def _load(path, password=None):
        ext = os.path.splitext(path)[1].lower()
        if ext in IMAGE_EXTS:
            img = pymupdf.open(path)
            data = img.convert_to_pdf()
            img.close()
            return pymupdf.open("pdf", data)
        with open(path, "rb") as f:            # mở từ bộ nhớ -> không khoá file, ghi đè được
            data = f.read()
        doc = pymupdf.open("pdf", data)
        if doc.needs_pass:
            if not password or not doc.authenticate(password):
                doc.close()
                raise PdfError("NEED_PASSWORD")
        return doc

    def open(self, path, password=None):
        doc = self._load(path, password)
        if self.doc:
            self.doc.close()
        self.doc = doc
        self.path = path
        self.file_size = os.path.getsize(path)
        self.modified = False
        self._undo.clear()
        self._redo.clear()

    def new_from_scan(self, images, dpi):
        """Tài liệu mới (chưa lưu) từ các ảnh quét."""
        doc = scan_images_doc(images, dpi)
        if self.doc:
            self.doc.close()
        self.doc = doc
        self.path = None
        self.file_size = 0
        self.modified = True
        self._undo.clear()
        self._redo.clear()

    def insert_scan(self, images, dpi, at=None):
        src = scan_images_doc(images, dpi)
        self.snapshot()
        pos = self.page_count if at is None else at
        self.doc.insert_pdf(src, start_at=pos)
        n = src.page_count
        src.close()
        return n

    def new_blank(self):
        """Tài liệu trống 1 trang A4 (chưa có đường dẫn - lưu lần đầu sẽ hỏi chỗ lưu)."""
        if self.doc:
            self.doc.close()
        self.doc = pymupdf.open()
        self.doc.new_page(width=595, height=842)
        self.path = None
        self.file_size = 0
        self.modified = False
        self._undo.clear()
        self._redo.clear()

    def save(self, path=None):
        path = path or self.path
        if not path:
            raise PdfError("Chưa có đường dẫn lưu file.")
        if self.new_fonts:                      # font Windows nhúng khi thêm chữ -> chỉ giữ các ký tự đã dùng
            try:
                self.doc.subset_fonts()
            except Exception:
                pass
            self.new_fonts = False
        data = self.doc.tobytes(garbage=3, deflate=True, use_objstms=1)
        tmp = path + ".tmp"
        with open(tmp, "wb") as f:
            f.write(data)
        os.replace(tmp, path)
        self.path = path
        self.file_size = len(data)
        self.modified = False

    @property
    def page_count(self):
        return self.doc.page_count if self.doc else 0

    # ----------------------------- hoàn tác -----------------------------
    def snapshot(self, fonts=False):
        self.new_fonts |= fonts
        self._undo.append(self.doc.tobytes())
        total = sum(len(b) for b in self._undo)
        while len(self._undo) > self.MAX_UNDO or (total > self.MAX_UNDO_BYTES and len(self._undo) > 1):
            total -= len(self._undo.pop(0))
        self._redo.clear()
        self.modified = True

    def _restore(self, src, dst):
        if not src:
            return False
        dst.append(self.doc.tobytes())
        self.doc.close()
        self.doc = pymupdf.open("pdf", src.pop())
        self.modified = True
        return True

    def undo(self):
        return self._restore(self._undo, self._redo)

    def redo(self):
        return self._restore(self._redo, self._undo)

    @property
    def can_undo(self):
        return bool(self._undo)

    @property
    def can_redo(self):
        return bool(self._redo)

    # ----------------------------- thao tác trang -----------------------------
    def delete_pages(self, indices):
        indices = sorted(set(indices))
        if len(indices) >= self.page_count:
            raise PdfError("Không thể xoá tất cả các trang.")
        self.snapshot()
        self.doc.delete_pages(indices)

    def reorder(self, order):
        if sorted(order) != list(range(self.page_count)):
            raise PdfError("Thứ tự trang không hợp lệ.")
        self.snapshot()
        self.doc.select(list(order))

    def move_pages(self, indices, delta):
        """Dịch các trang đã chọn lên (-1) / xuống (+1). Trả về vị trí mới."""
        n = self.page_count
        order = list(range(n))
        sel = sorted(set(indices), reverse=(delta > 0))
        for i in sel:
            pos = order.index(i)
            new = pos + delta
            if 0 <= new < n and order[new] not in sel:
                order[pos], order[new] = order[new], order[pos]
        if order == list(range(n)):
            return list(indices)
        self.reorder(order)
        return sorted(order.index(i) for i in indices)

    def rotate(self, indices, degrees):
        self.snapshot()
        for i in indices:
            p = self.doc[i]
            p.set_rotation((p.rotation + degrees) % 360)

    def insert_files(self, paths, at=None, passwords=None):
        """Chèn file PDF / ảnh vào vị trí `at` (mặc định cuối). Trả về số trang đã chèn."""
        loaded = [self._load(p, (passwords or {}).get(p)) for p in paths]
        self.snapshot()
        pos = self.page_count if at is None else at
        total = 0
        for src in loaded:
            self.doc.insert_pdf(src, start_at=pos)
            pos += src.page_count
            total += src.page_count
            src.close()
        return total

    def extract(self, indices, out_path):
        new = pymupdf.open()
        for i in sorted(indices):
            new.insert_pdf(self.doc, from_page=i, to_page=i)
        new.save(out_path, garbage=3, deflate=True)
        new.close()

    # ----------------------------- nén -----------------------------
    COMPRESS_LEVELS = {
        "low":    None,                                              # chỉ tối ưu cấu trúc
        "medium": dict(dpi_threshold=200, dpi_target=150, quality=75),
        "high":   dict(dpi_threshold=120, dpi_target=96, quality=55),
    }

    def compressed_bytes(self):
        return self.doc.tobytes(garbage=4, deflate=True, deflate_images=True, deflate_fonts=True,
                                clean=True, use_objstms=1)

    def compress(self, level="medium"):
        """Nén ảnh + tối ưu cấu trúc. Trả về (kích thước trước, kích thước sau) theo byte."""
        before = len(self.doc.tobytes(garbage=1, deflate=True)) if self.modified or not self.file_size else self.file_size
        self.snapshot()
        params = self.COMPRESS_LEVELS[level]
        if params:
            self.doc.rewrite_images(lossy=True, lossless=True, bitonal=True, color=True, gray=True, **params)
        try:
            self.doc.subset_fonts()
        except Exception:
            pass
        data = self.compressed_bytes()
        self.doc.close()
        self.doc = pymupdf.open("pdf", data)
        return before, len(data)

    # ----------------------------- chữ định dạng (HTML) -----------------------------
    @staticmethod
    def _first_baseline(size, rect, html, css, scale_low):
        """Vị trí baseline dòng đầu khi đặt `html` vào `rect` (thử trên trang tạm)."""
        tmp = pymupdf.open()
        p = tmp.new_page(width=size[0], height=size[1])
        p.insert_htmlbox(rect, html, css=css, archive=_archive(), scale_low=scale_low)
        for b in p.get_text("dict")["blocks"]:
            for ln in b.get("lines", []):
                for s in ln["spans"]:
                    if s["text"].strip():
                        tmp.close()
                        return s["origin"][1]
        tmp.close()
        return None

    def move_line(self, pno, line: TextLine, html, families, dx, dy):
        """Dời dòng chữ (dx, dy theo toạ độ trang chưa xoay): xoá chữ ở chỗ cũ, đặt lại cùng nội dung / định dạng
        ở chỗ mới (baseline dời đúng dy)."""
        import dataclasses
        bb = pymupdf.Rect(line.bbox) + (dx, dy, dx, dy)
        bb.x1 += max(bb.width * 0.5, 2 * bb.height)       # nới bề ngang: font thay thế hơi rộng hơn không bị xuống dòng
        dest = dataclasses.replace(line, bbox=bb, baseline=line.baseline + dy)
        self.edit_line_rich(pno, line, html, families, "left", fit=True, dest=dest)

    def edit_line_rich(self, pno, line: TextLine, html, families, align="left", fit=False, dest=None):
        """Thay dòng chữ bằng nội dung có định dạng, giữ nguyên vị trí baseline dòng đầu.
        dest: đặt nội dung mới ở vị trí của dòng khác (dời dòng) - chữ cũ vẫn xoá ở `line`."""
        self.snapshot(fonts=True)
        page = self.doc[pno]
        _redact(page, [line.bbox])
        if not html:
            return
        self._place_line(page, dest or line, html, families, align, fit)

    def move_lines(self, pno, lines, htmls, dx, dy):
        """Dời cả khối chữ (nhiều dòng): xoá hết các dòng cũ một lượt rồi đặt lại từng dòng ở chỗ mới (giữ nguyên
        cách ngắt dòng / khoảng cách dòng). htmls: [(html, fonts)] tương ứng từng dòng."""
        import dataclasses
        self.snapshot(fonts=True)
        page = self.doc[pno]
        _redact(page, [ln.bbox for ln in lines])
        for ln, (body, fams) in zip(lines, htmls):
            bb = pymupdf.Rect(ln.bbox) + (dx, dy, dx, dy)
            bb.x1 += max(bb.width * 0.5, 2 * bb.height)
            self._place_line(page, dataclasses.replace(ln, bbox=bb, baseline=ln.baseline + dy), body, fams,
                             "left", True)

    def delete_lines(self, pno, lines):
        self.snapshot(fonts=True)
        _redact(self.doc[pno], [ln.bbox for ln in lines])

    def edit_block_rich(self, pno, lines, html, families):
        """Thay cả khối chữ bằng nội dung mới: chữ chảy lại trong bề ngang của khối (xuống dòng tự động),
        baseline dòng đầu giữ như cũ."""
        self.snapshot(fonts=True)
        page = self.doc[pno]
        _redact(page, [ln.bbox for ln in lines])
        if not html:
            return
        css = font_css(families) + "\n" + BASE_CSS
        bounds = disp_to_unrot(page, page.rect)
        box = pymupdf.Rect(lines[0].bbox)
        for ln in lines[1:]:
            box |= ln.bbox
        first = min(lines, key=lambda ln: ln.bbox.y0)
        rect = pymupdf.Rect(box.x0, first.bbox.y0 - first.bbox.height, box.x1 + 2, bounds.y1)
        got = self._first_baseline((bounds.width, bounds.height), rect, html, css, 1)
        if got is not None:
            rect.y0 += first.baseline - got
            rect.y1 += first.baseline - got
        page.insert_htmlbox(rect, html, css=css, archive=_archive(), scale_low=1)

    def _place_line(self, page, line, html, families, align="left", fit=False):
        """Đặt nội dung HTML vào vị trí dòng `line` (baseline dòng đầu trùng baseline của line)."""
        css = font_css(families) + "\n" + BASE_CSS
        bounds = disp_to_unrot(page, page.rect)
        bb = line.bbox
        top = bb.y0 - bb.height
        if fit:
            rect = pymupdf.Rect(bb.x0, top, bb.x1, bb.y1 + bb.height * 0.6)
        elif align == "center":
            c = (bb.x0 + bb.x1) / 2
            hw = max(bb.width / 2, min(c - bounds.x0, bounds.x1 - c) - 2)
            rect = pymupdf.Rect(c - hw, top, c + hw, bounds.y1)
        elif align == "right":
            rect = pymupdf.Rect(bounds.x0 + 2, top, bb.x1, bounds.y1)
        else:
            rect = pymupdf.Rect(bb.x0, top, max(bb.x1, bounds.x1 - 2), bounds.y1)
        scale_low = 0 if fit else 1
        got = self._first_baseline((bounds.width, bounds.height), rect, html, css, scale_low)
        if got is not None:
            rect.y0 += line.baseline - got
            rect.y1 += line.baseline - got
        page.insert_htmlbox(rect, html, css=css, archive=_archive(), scale_low=scale_low)

    def add_rich_text(self, pno, disp_rect, html, families):
        """Thêm khung chữ có định dạng. `disp_rect` theo toạ độ hiển thị; chữ chảy xuống dưới nếu dài."""
        page = self.doc[pno]
        pr = page.rect
        r = pymupdf.Rect(disp_rect)
        if r.width < 8:                           # chỉ bấm chuột -> rộng đến lề phải
            r.x1 = pr.x1 - 20 if pr.x1 - 20 > r.x0 + 40 else pr.x1
        r.y1 = pr.y1                              # chiều cao: đến cuối trang
        css = font_css(families) + "\n" + BASE_CSS
        self.snapshot(fonts=True)
        spare, scale = page.insert_htmlbox(disp_to_unrot(page, r), html, css=css, archive=_archive(),
                                           rotate=page.rotation, scale_low=0)
        return scale

    # ----------------------------- chèn -----------------------------
    MAX_IMAGE_PX = 3000

    @classmethod
    def _shrink_image(cls, path=None, stream=None):
        """Ảnh quá lớn (ảnh điện thoại 40-50 MP) -> thu về cạnh dài 3000 px: đủ nét cho trang A4, chèn / lưu
        không bị đứng app. Giữ nền trong suốt (PNG) nếu có. Trả về bytes hoặc None nếu không cần thu nhỏ."""
        try:
            pix = pymupdf.Pixmap(path) if path else pymupdf.Pixmap(stream)
        except Exception:
            return None
        long_side = max(pix.width, pix.height)
        if long_side <= cls.MAX_IMAGE_PX:
            return None
        k = cls.MAX_IMAGE_PX / long_side
        if pix.colorspace and pix.colorspace.n not in (1, 3):     # CMYK... -> RGB
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        small = pymupdf.Pixmap(pix, int(pix.width * k), int(pix.height * k), None)
        return small.tobytes("png") if small.alpha else small.tobytes("jpg", jpg_quality=88)

    def insert_image(self, pno, disp_rect, path=None, stream=None):
        page = self.doc[pno]
        small = self._shrink_image(path, stream)
        if small is not None:
            path, stream = None, small
        r = pymupdf.Rect(disp_rect)
        if r.width < 8 or r.height < 8:           # chỉ bấm chuột -> kích thước gốc (tối đa 60% trang)
            src = pymupdf.open(path) if path else pymupdf.open(stream=stream)
            w, h = src[0].rect.width, src[0].rect.height
            src.close()
            k = min(1.0, page.rect.width * 0.6 / w, page.rect.height * 0.6 / h)
            r = pymupdf.Rect(r.x0, r.y0, r.x0 + w * k, r.y0 + h * k)
        self.snapshot()
        page.insert_image(disp_to_unrot(page, r), filename=path, stream=stream, keep_proportion=True,
                          rotate=(-page.rotation) % 360)

    def insert_image_stamp(self, pno, disp_rect, path=None, stream=None, author=""):
        """Chèn ảnh dạng con dấu (Stamp annotation): nằm trên trang như ghi chú - kéo / ẩn / xoá / cố định được,
        không trộn vào nội dung trang. Giữ tỉ lệ ảnh trong vùng."""
        page = self.doc[pno]
        small = self._shrink_image(path, stream)
        if small is not None:
            path, stream = None, small
        pix = pymupdf.Pixmap(path) if path else pymupdf.Pixmap(stream)
        r = pymupdf.Rect(disp_rect)
        if r.width < 8 or r.height < 8:           # chỉ bấm chuột -> kích thước gốc (tối đa 60% trang)
            k = min(1.0, page.rect.width * 0.6 / pix.width, page.rect.height * 0.6 / pix.height)
            r = pymupdf.Rect(r.x0, r.y0, r.x0 + pix.width * k, r.y0 + pix.height * k)
        else:                                     # vừa khung, giữ tỉ lệ, căn giữa
            k = min(r.width / pix.width, r.height / pix.height)
            w, h = pix.width * k, pix.height * k
            r = pymupdf.Rect(r.x0 + (r.width - w) / 2, r.y0 + (r.height - h) / 2,
                             r.x0 + (r.width + w) / 2, r.y0 + (r.height + h) / 2)
        self.snapshot()
        a = page.add_stamp_annot(disp_to_unrot(page, r), stamp=pix)
        now = pymupdf.get_pdf_now()
        a.set_info(title=author, creationDate=now, modDate=now)
        a.update()

    FIELD_TYPES = {"text": pymupdf.PDF_WIDGET_TYPE_TEXT, "multiline": pymupdf.PDF_WIDGET_TYPE_TEXT,
                   "check": pymupdf.PDF_WIDGET_TYPE_CHECKBOX, "combo": pymupdf.PDF_WIDGET_TYPE_COMBOBOX,
                   "list": pymupdf.PDF_WIDGET_TYPE_LISTBOX}

    def field_names(self):
        names = set()
        for pg in self.doc:
            for w in pg.widgets() or []:
                names.add(w.field_name)
        return names

    def add_field(self, pno, disp_rect, kind, name, value="", choices=(), fontsize=11, required=False):
        """Tạo ô form (text / multiline / check / combo / list) tại vùng hiển thị."""
        page = self.doc[pno]
        names = self.field_names()
        base, k = name or "O_form", 2
        while name in names or not name:
            name = f"{base}_{k}"
            k += 1
        self.snapshot()
        self._add_widget(page, disp_to_unrot(page, pymupdf.Rect(disp_rect)), kind, name, value, choices, fontsize,
                         required)
        return name

    def page_fields(self, pno):
        """Ô form trên trang: [(vùng hiển thị, xref, tên)]."""
        page = self.doc[pno]
        out = []
        for w in page.widgets() or []:
            r = w.rect * page.rotation_matrix
            r.normalize()
            if not r.is_empty:
                out.append((r, w.xref, w.field_name))
        return out

    def _widget(self, pno, xref):
        page = self.doc[pno]
        for w in page.widgets() or []:
            if w.xref == xref:
                return page, w
        raise PdfError("Không tìm thấy ô form này trên trang (có thể đã bị xoá).")

    def delete_field(self, pno, xref):
        self._widget(pno, xref)
        self.snapshot()
        page, w = self._widget(pno, xref)
        page.delete_widget(w)

    def move_field(self, pno, xref, old_disp, new_disp):
        self._widget(pno, xref)
        self.snapshot()
        page, w = self._widget(pno, xref)
        w.rect = disp_to_unrot(page, pymupdf.Rect(new_disp))
        w.update()
        return True

    def rename_field(self, pno, xref, name):
        if name in self.field_names():
            raise PdfError(f"Đã có ô form tên '{name}'.")
        self._widget(pno, xref)
        self.snapshot()
        page, w = self._widget(pno, xref)
        w.field_name = name
        w.update()

    def detect_fields(self, pages=None):
        """Tự nhận diện chỗ trống để điền và tạo ô form. Trả về {kind: số ô}."""
        pages = range(self.page_count) if pages is None else pages
        found = {pno: detect_form_fields(self.doc[pno]) for pno in pages}
        if not any(found.values()):
            return {}
        self.snapshot()
        used, counts, n = self.field_names(), {}, 0
        for pno, fields in found.items():
            page = self.doc[pno]
            for f in fields:
                n += 1
                name = field_name_from_label(f["label"], used, n)
                fs = 0 if f["kind"] == "multiline" else 10
                self._add_widget(page, f["rect"], f["kind"], name, fontsize=fs)
                counts[f["kind"]] = counts.get(f["kind"], 0) + 1
        return counts

    def _add_widget(self, page, rect, kind, name, value="", choices=(), fontsize=11, required=False):
        w = pymupdf.Widget()
        w.field_type = self.FIELD_TYPES[kind]
        w.field_name = name
        w.rect = rect
        w.border_color = (0.45, 0.55, 0.7)
        w.border_width = 0.6
        w.fill_color = (0.95, 0.97, 1.0)
        w.text_font = "Helv"
        w.text_fontsize = fontsize
        if kind == "multiline":
            w.field_flags |= pymupdf.PDF_TX_FIELD_IS_MULTILINE
        if required:
            w.field_flags |= pymupdf.PDF_FIELD_IS_REQUIRED
        if kind in ("combo", "list"):
            w.choice_values = list(choices)
            if value and value in choices:
                w.field_value = value
        elif kind == "check":
            w.field_value = bool(value)
        elif value:
            w.field_value = value
        page.add_widget(w)

    # Chữ hiển thị của liên kết là 1 khối riêng (Form XObject, vẽ bằng "q cm /Tên Do Q"); liên kết ghi tên khối
    # trong khoá /PTKText -> kéo / xoá liên kết thì chữ đi theo.
    LINK_TEXT_KEY = "PTKText"

    def _html_block(self, page, disp_rect, html, families):
        """Dựng chữ HTML thành khối đặt tại góc trên trái vùng (rộng theo vùng kéo, hoặc tới lề phải nếu chỉ bấm).
        Trả về (vùng hiển thị thật của chữ, tên khối trong tài nguyên trang)."""
        r = pymupdf.Rect(disp_rect)
        pr = page.rect
        w = r.width if r.width >= 8 else max(60, pr.x1 - r.x0 - 10)
        tmp = pymupdf.open()
        tp = tmp.new_page(width=w, height=max(60, pr.height))
        css = font_css(families) + "\n" + BASE_CSS
        tp.insert_htmlbox(tp.rect, html, css=css, archive=_archive(), scale_low=0)
        bb = pymupdf.Rect()
        for b in tp.get_text("dict")["blocks"]:
            bb |= pymupdf.Rect(b["bbox"])
        for dr in tp.get_drawings():                 # gạch chân, nền chữ
            bb |= dr["rect"]
        if bb.is_empty:
            raise PdfError("Chữ hiển thị trống.")
        bb = (bb + (-1, -1, 1, 1)) & tp.rect
        target = pymupdf.Rect(r.x0 + bb.x0, r.y0 + bb.y0, r.x0 + bb.x1, r.y0 + bb.y1)
        names = {x[1] for x in page.get_xobjects()}
        page.show_pdf_page(disp_to_unrot(page, target), tmp, 0, clip=bb, rotate=page.rotation)
        tmp.close()
        name = next((x[1] for x in page.get_xobjects() if x[1] not in names and x[2] == 0), None)
        if name:                                     # thêm ma trận cm riêng -> di chuyển / co giãn được về sau
            for cx in reversed(page.get_contents()):
                s = self.doc.xref_stream(cx).decode("latin-1")
                k = s.rfind(f"/{name} Do")
                if k >= 0:
                    self.doc.update_stream(cx, (s[:k] + "1 0 0 1 0 0 cm " + s[k:]).encode("latin-1"))
                    break
        return target, name

    def add_link(self, pno, disp_rect, uri=None, target_page=None, html=None, families=()):
        page = self.doc[pno]
        self.snapshot(fonts=bool(html))
        r = pymupdf.Rect(disp_rect)
        name = None
        if html:
            r, name = self._html_block(page, r, html, families)
        un = disp_to_unrot(page, r)
        if target_page is not None:
            page.insert_link({"kind": pymupdf.LINK_GOTO, "from": un, "page": target_page, "to": pymupdf.Point(0, 0)})
        else:
            page.insert_link({"kind": pymupdf.LINK_URI, "from": un, "uri": uri})
        if name:
            page = self.doc.reload_page(page)
            links = page.get_links()
            if links:
                self.doc.xref_set_key(links[-1]["xref"], self.LINK_TEXT_KEY, "/" + name)

    def page_links(self, pno):
        """[(vùng hiển thị, xref, {uri | page}, có chữ đi kèm)] - liên kết kéo / bấm được trên trang."""
        out = []
        for l in self.doc[pno].get_links():
            if not l.get("xref"):
                continue
            r = pymupdf.Rect(l["from"])
            r.normalize()
            if r.is_empty:
                continue
            target = {"uri": l["uri"]} if l.get("uri") else {"page": l["page"]} if l.get("page", -1) >= 0 else {}
            if l.get("kind") == pymupdf.LINK_LAUNCH and l.get("file"):
                target = {"file": l["file"]}
            has_text = self.doc.xref_get_key(l["xref"], self.LINK_TEXT_KEY)[0] == "name"
            out.append((r, l["xref"], target, has_text))
        return out

    def _link_text_op(self, page, xref):
        """Lệnh vẽ chữ đi kèm liên kết: (xref stream, nội dung, match cm, match Do) hoặc None."""
        typ, val = self.doc.xref_get_key(xref, self.LINK_TEXT_KEY)
        if typ != "name":
            return None
        rx = re.compile(r"((?:" + self._NUM + r"\s+){6})cm\s*(" + re.escape(val) + r"\s+Do\b)")
        for cx in page.get_contents():
            s = self.doc.xref_stream(cx).decode("latin-1")
            m = rx.search(s)
            if m:
                return cx, s, m
        return None

    def delete_link(self, pno, xref):
        page = self.doc[pno]
        lnk = next((l for l in page.get_links() if l.get("xref") == xref), None)
        if lnk is None:
            return False
        self.snapshot()
        op = self._link_text_op(page, xref)
        if op:
            cx, s, m = op
            self.doc.update_stream(cx, (s[:m.start(2)] + s[m.end(2):]).encode("latin-1"))
        page.delete_link(lnk)
        return True

    def set_link_uri(self, pno, xref, uri):
        page = self.doc[pno]
        lnk = next((l for l in page.get_links() if l.get("xref") == xref), None)
        if lnk is None:
            return False
        self.snapshot()
        lnk["kind"], lnk["uri"] = pymupdf.LINK_URI, uri
        page.update_link(lnk)
        return True

    def add_highlight(self, pno, disp_rect, color=(1, 0.92, 0.23)):
        """Tô sáng các từ nằm trong vùng (hoặc cả vùng nếu không có chữ). Trả về số từ."""
        page = self.doc[pno]
        un = disp_to_unrot(page, disp_rect)
        hits = []
        for w in page.get_text("words"):
            wr = pymupdf.Rect(w[:4])
            inter = wr & un
            if not inter.is_empty and inter.get_area() >= 0.5 * wr.get_area():
                hits.append(wr)
        self.snapshot()
        a = page.add_highlight_annot(quads=[h.quad for h in hits]) if hits else page.add_highlight_annot(un)
        a.set_colors(stroke=color)
        a.update()
        return len(hits)

    def add_note(self, pno, disp_point, text, author=""):
        page = self.doc[pno]
        self.snapshot()
        a = page.add_text_annot(pymupdf.Point(disp_point) * page.derotation_matrix, text, icon="Comment")
        now = pymupdf.get_pdf_now()
        a.set_info(title=author, content=text, creationDate=now, modDate=now)
        a.update()

    def add_bookmark(self, pno, title, disp_rect=None):
        """Thêm bookmark (mục lục) cấp 1 trỏ tới trang pno - tới đúng chỗ đầu vùng nếu có disp_rect."""
        page = self.doc[pno]
        self.snapshot()
        toc = self.doc.get_toc(simple=False)
        dest = {"kind": pymupdf.LINK_GOTO, "page": pno}
        if disp_rect is not None:
            r = disp_to_unrot(page, pymupdf.Rect(disp_rect))
            dest["to"] = pymupdf.Point(r.x0, r.y0)
        toc.append([1, title, pno + 1, dest])
        self.doc.set_toc(toc)

    def crop_to(self, pno, disp_rect):
        """Cắt trang: chỉ giữ vùng hiển thị disp_rect (đặt lại CropBox - nội dung ngoài vùng vẫn còn trong file)."""
        page = self.doc[pno]
        r = disp_to_unrot(page, pymupdf.Rect(disp_rect))
        cp = page.cropbox_position                       # toạ độ trang tính từ góc CropBox hiện tại -> theo MediaBox
        r = pymupdf.Rect(r.x0 + cp.x, r.y0 + cp.y, r.x1 + cp.x, r.y1 + cp.y) & page.mediabox
        if r.is_empty or r.width < 5 or r.height < 5:
            raise PdfError("Vùng chọn quá nhỏ để cắt trang.")
        self.snapshot()
        page.set_cropbox(r)

    # ----- ghi chú / chú thích có sẵn -----
    def _annot(self, pno, xref):
        page = self.doc[pno]                      # annot cần trang còn sống -> giữ tham chiếu trên annot
        if xref not in [x[0] for x in page.annot_xrefs()]:
            return None
        a = page.load_annot(xref)
        a._keep_page = page
        return a

    def annot_info(self, pno, xref):
        a = self._annot(pno, xref)
        if a is None:
            return None
        i = a.info
        d = i.get("modDate") or i.get("creationDate") or ""
        m = re.match(r"D:(\d{4})(\d{2})(\d{2})(\d{2})?(\d{2})?", d)
        date = f"{m[3]}/{m[2]}/{m[1]}" + (f" {m[4]}:{m[5]}" if m[4] and m[5] else "") if m else ""
        kinds = {pymupdf.PDF_ANNOT_TEXT: "Ghi chú", pymupdf.PDF_ANNOT_HIGHLIGHT: "Tô sáng",
                 pymupdf.PDF_ANNOT_FREE_TEXT: "Hộp chữ", pymupdf.PDF_ANNOT_SQUARE: "Hình chữ nhật",
                 pymupdf.PDF_ANNOT_CIRCLE: "Hình elip", pymupdf.PDF_ANNOT_LINE: "Đường thẳng"}
        return dict(title=i.get("title", ""), content=i.get("content", ""), date=date,
                    kind=kinds.get(a.type[0], a.type[1]))

    def set_annot_content(self, pno, xref, text):
        a = self._annot(pno, xref)
        if a is not None:
            self.snapshot()
            a = self._annot(pno, xref)
            a.set_info(content=text)
            a.update()

    # ----- ảnh / chữ ký tay / con dấu đã chèn: kéo để di chuyển, kéo góc để đổi cỡ -----
    _NUM = r"[-+]?(?:\d+\.?\d*|\.\d+)"
    _CM = re.compile(r"((?:" + _NUM + r"\s+){6})cm\b")

    def page_images(self, pno):
        """Ảnh di chuyển được trên trang: [(vùng hiển thị, xref)] (bỏ ảnh nền phủ gần kín trang - vd bản scan)."""
        page = self.doc[pno]
        area = abs(page.rect)
        out = []
        try:
            infos = page.get_image_info(xrefs=True)
        except Exception:
            return out
        for info in infos:
            if not info.get("xref"):
                continue                                  # ảnh inline: không sửa được
            r = pymupdf.Rect(info["bbox"]) * page.rotation_matrix
            r.normalize()
            if r.is_empty or abs(r) > area * 0.8:
                continue
            out.append((r, info["xref"]))
        return out

    def _find_image_op(self, page, xref, disp_rect):
        """Tìm lệnh vẽ ảnh trong nội dung trang: (xref stream, nội dung, match cm, match Do, ma trận cm).
        Chỉ nhận lệnh có ma trận cm cho đúng vị trí đang thấy (ảnh nằm trong nhóm biến đổi lồng nhau thì bỏ)."""
        names = {im[7] for im in page.get_images(full=True) if im[0] == xref and im[7]}
        if not names:
            raise PdfError("Ảnh này nằm trong một nhóm đối tượng - không di chuyển được.")
        want = disp_to_unrot(page, disp_rect)
        tm = page.transformation_matrix
        # 1 ảnh có thể có nhiều tên tài nguyên (vd sau khi nhân bản)
        do_re = re.compile(r"/(?:" + "|".join(re.escape(n) for n in names) + r")\s+Do\b")
        best = None
        for cx in page.get_contents():
            s = self.doc.xref_stream(cx).decode("latin-1")
            cms = list(self._CM.finditer(s))
            for d in do_re.finditer(s):
                prev = [m for m in cms if m.end() <= d.start()]
                if not prev:
                    continue
                cm = prev[-1]
                mat = pymupdf.Matrix(*[float(x) for x in cm.group(1).split()])
                r = pymupdf.Rect(0, 0, 1, 1) * mat * tm
                r.normalize()
                dist = sum(abs(a - b) for a, b in zip(r, want))
                if best is None or dist < best[0]:
                    best = (dist, cx, s, cm, d, mat)
        if best is None or best[0] > 3:
            raise PdfError("Không di chuyển được ảnh này (ảnh nằm trong nhóm đối tượng của file gốc).")
        return best[1:]

    @staticmethod
    def _fmt_matrix(m):
        return " ".join(f"{v:.6f}".rstrip("0").rstrip(".") or "0" for v in tuple(m))

    def _locate_image_do(self, page, xref, disp_rect):
        """Ảnh nằm trong nhóm biến đổi lồng nhau / trong Form XObject (file gốc từ Word, Acrobat, máy scan...):
        tìm đúng lệnh Do của lần vẽ này (trong nội dung trang hoặc trong form chứa nó) và ma trận biến đổi thật
        của ảnh (PDF user space). Trả về (xref stream, nội dung, match Do, ma trận ảnh T)."""
        want = disp_to_unrot(page, disp_rect)
        infos = [i for i in page.get_image_info(xrefs=True) if i.get("xref") == xref]
        if not infos:
            raise PdfError("Không tìm thấy ảnh này trên trang.")
        dist = [sum(abs(a - b) for a, b in zip(pymupdf.Rect(i["bbox"]), want)) for i in infos]
        k = dist.index(min(dist))
        if dist[k] > 3:
            raise PdfError("Không tìm thấy ảnh này trên trang.")
        # transform của MuPDF = lật ảnh (gốc trên-trái) x CTM thật x (PDF -> toạ độ MuPDF)
        flip = pymupdf.Matrix(1, 0, 0, -1, 0, 1)
        T = flip * pymupdf.Matrix(infos[k]["transform"]) * ~page.transformation_matrix
        if abs(T.a * T.d - T.b * T.c) < 1e-9:
            raise PdfError("Ảnh này có kích thước bằng 0 - không di chuyển được.")
        entries = [im for im in page.get_images(full=True) if im[0] == xref]
        refs = {im[9] for im in entries}
        names = {im[7] for im in entries if im[7]}
        if not names:
            raise PdfError("Không tìm thấy ảnh này trong nội dung trang.")
        do_re = re.compile(r"/(?:" + "|".join(re.escape(n) for n in names) + r")\s+Do\b")
        if len(refs) > 1 and 0 in refs:
            # vừa có lần vẽ thẳng trên trang (vd bản nhân bản) vừa có lần vẽ trong form: bỏ các lần vẽ thẳng
            # (đã xử lý được bằng _find_image_op) rồi tìm lần vẽ còn lại trong form
            tm = page.transformation_matrix
            direct = []
            for cx in page.get_contents():
                s = self.doc.xref_stream(cx).decode("latin-1")
                cms = list(self._CM.finditer(s))
                for d in do_re.finditer(s):
                    prev = [m for m in cms if m.end() <= d.start()]
                    if prev:
                        r = pymupdf.Rect(0, 0, 1, 1) * pymupdf.Matrix(*[float(v) for v in prev[-1].group(1).split()]) * tm
                        r.normalize()
                        direct.append(r)
            target = infos[k]

            def is_direct(i):
                return any(sum(abs(a - b) for a, b in zip(pymupdf.Rect(i["bbox"]), r)) < 1 for r in direct)
            n_page = sum(len(do_re.findall(self.doc.xref_stream(cx).decode("latin-1"))) for cx in page.get_contents())
            if len(direct) == n_page and not is_direct(target):
                infos = [i for i in infos if not is_direct(i)]
                k = infos.index(target)
                refs.discard(0)
        if len(refs) != 1:
            raise PdfError("Ảnh này nằm trong nhiều nhóm đối tượng khác nhau - không di chuyển được.")
        ref = refs.pop()
        if ref == 0:                                            # ảnh vẽ thẳng trong nội dung trang (lồng nhiều q/cm)
            # đọc lần lượt nội dung trang, tính ma trận thật ở từng lệnh vẽ ảnh rồi chọn lệnh khớp vị trí đang thấy
            # (không dựa vào số lần vẽ: ảnh bị cắt khuất / ngoài trang không có trong danh sách hiển thị)
            tm = page.transformation_matrix
            best, stack, ctm = None, [], pymupdf.Identity
            for cx in page.get_contents():
                s = self.doc.xref_stream(cx).decode("latin-1")
                hits, ctm = self._scan_do(s, names, stack, ctm)     # trạng thái q/cm nối tiếp giữa các stream
                for d, m in hits:
                    r = pymupdf.Rect(0, 0, 1, 1) * m * tm
                    r.normalize()
                    dist = sum(abs(a - b) for a, b in zip(r, want))
                    if best is None or dist < best[0]:
                        best = (dist, cx, s, d, m)
            if best is None or best[0] > 3:
                raise PdfError("Không xác định được lần vẽ ảnh này trong nội dung trang.")
            _dist, cx, s, d, m = best
            return cx, s, d, m
        # ảnh nằm trong Form XObject: chỉ sửa khi form được vẽ đúng 1 lần (trên trang này) và chứa 1 lần vẽ ảnh
        fx = self._own_form(page, ref)
        s = self.doc.xref_stream(fx).decode("latin-1")
        occ = list(do_re.finditer(s))
        if len(occ) != 1 or len(infos) != 1:
            raise PdfError("Ảnh này nằm trong nhóm đối tượng dùng nhiều lần - không di chuyển riêng được.")
        return fx, s, occ[0], T

    class _Span:
        """Vị trí '/Tên Do' trong nội dung (dùng như re.Match: start() / end())."""
        def __init__(self, a, b):
            self.a, self.b = a, b

        def start(self):
            return self.a

        def end(self):
            return self.b

    _TOK = re.compile(rb"%[^\r\n]*|/[^\s/\[\]()<>{}%]*|<<|>>|<[0-9A-Fa-f\s]*>|[\[\]{}]|"
                      rb"[-+]?(?:\d+\.?\d*|\.\d+)|[^\s/\[\]()<>{}%]+")

    def _scan_do(self, s, names, stack, ctm):
        """Đọc nội dung trang, theo dõi q / Q / cm: trả về ([(_Span của '/Tên Do', ma trận lúc vẽ)], ctm cuối).
        stack: ngăn xếp q (sửa tại chỗ - nối tiếp sang stream nội dung kế tiếp)."""
        b = s.encode("latin-1")
        hits, ops, i, n = [], [], 0, len(b)
        while i < n:
            c = b[i:i + 1]
            if c in b" \t\r\n\f\x00":
                i += 1
                continue
            if c == b"(":                                      # chuỗi: bỏ qua (có ngoặc lồng, ký tự thoát)
                depth, i = 1, i + 1
                while i < n and depth:
                    ch = b[i:i + 1]
                    if ch == b"\\":
                        i += 2
                        continue
                    depth += {b"(": 1, b")": -1}.get(ch, 0)
                    i += 1
                ops = []
                continue
            m = self._TOK.match(b, i)
            if not m:
                i += 1
                continue
            t, i0, i = m.group(), m.start(), m.end()
            if t[:1] == b"%":
                continue
            if t[:1] == b"/" or t[:1] in b"[]{}<>" or re.fullmatch(rb"[-+]?(?:\d+\.?\d*|\.\d+)", t):
                ops.append((t, i0))
                continue
            if t == b"q":
                stack.append(ctm)
            elif t == b"Q":
                ctm = stack.pop() if stack else ctm
            elif t == b"cm" and len(ops) >= 6:
                try:
                    ctm = pymupdf.Matrix(*[float(x) for x, _p in ops[-6:]]) * ctm
                except ValueError:
                    pass
            elif t == b"Do" and ops and ops[-1][0][:1] == b"/":
                if ops[-1][0][1:].decode("latin-1") in names:
                    hits.append((self._Span(ops[-1][1], i), ctm))
            elif t == b"ID":                                    # ảnh inline: bỏ qua dữ liệu tới EI
                e = re.compile(rb"\sEI(?=[\s/\[<(%]|$)").search(b, i + 1)
                i = e.end() if e else n
            ops = []
        return hits, ctm

    def _own_form(self, page, fx, xobjs=None):
        """Form XObject fx (có thể lồng trong form khác) chỉ dùng 1 lần, ở trang này? Nếu trang khác cũng dùng ->
        tách bản riêng cho trang này (cả chuỗi form cha) để dời ảnh ở trang này không xê dịch các trang khác."""
        xobjs = xobjs if xobjs is not None else page.get_xobjects()
        mine = [x for x in xobjs if x[0] == fx]
        if len(mine) != 1:
            raise PdfError("Ảnh này nằm trong nhóm đối tượng được vẽ nhiều lần trên trang - không di chuyển riêng được.")
        _x, name, invoker, _bbox = mine[0]
        others = any(fx in [x[0] for x in p.get_xobjects()] for p in self.doc if p.number != page.number)
        if not others:
            return fx
        new = self.doc.get_new_xref()
        self.doc.update_object(new, "<<>>")
        self.doc.xref_copy(fx, new)
        holder = page.xref if invoker == 0 else self._own_form(page, invoker, xobjs)
        self._set_xobject(holder, name, new)
        return new

    def _set_xobject(self, holder, name, target):
        """Trỏ tên XObject `name` trong /Resources của đối tượng holder (trang hoặc form) sang xref target.
        Bảng tài nguyên dùng chung (tham chiếu gián tiếp) thì tách bản riêng trước khi sửa."""
        def own(obj, key):
            v = self.doc.xref_get_key(obj, key)
            if v[0] != "xref":
                return obj, key + "/"
            nx = self.doc.get_new_xref()
            self.doc.update_object(nx, self.doc.xref_object(int(v[1].split()[0]), compressed=True))
            self.doc.xref_set_key(obj, key, f"{nx} 0 R")
            return nx, ""
        obj, path = own(holder, "Resources")
        obj, path2 = own(obj, path + "XObject")
        self.doc.xref_set_key(obj, path2 + name, f"{target} 0 R")

    def _image_shift(self, page, old_disp, new_disp):
        """Phép biến đổi (PDF user space) đưa vùng hiển thị cũ sang vùng mới."""
        o, n = disp_to_unrot(page, old_disp), disp_to_unrot(page, new_disp)
        f = (pymupdf.Matrix(1, 0, 0, 1, -o.x0, -o.y0)
             * pymupdf.Matrix(n.width / max(o.width, 1e-6), 0, 0, n.height / max(o.height, 1e-6), 0, 0)
             * pymupdf.Matrix(1, 0, 0, 1, n.x0, n.y0))
        tm = page.transformation_matrix
        return tm * f * ~tm

    _WRAP = re.compile(r"q\s+((?:" + _NUM + r"\s+){6})cm\s+$")

    def move_image(self, pno, xref, old_disp, new_disp):
        """Di chuyển / đổi cỡ ảnh: sửa ma trận cm ngay trước lệnh vẽ ảnh (giữ nguyên ảnh, nền trong suốt...).
        Ảnh nằm trong nhóm biến đổi lồng nhau / Form XObject: bọc lệnh vẽ bằng 'q X cm ... Q' với X tính từ
        ma trận thật của ảnh (lần dời sau sửa luôn X, không lồng thêm)."""
        page = self.doc[pno]
        F = self._image_shift(page, old_disp, new_disp)
        try:
            cx, s, cm, _d, mat = self._find_image_op(page, xref, old_disp)
        except PdfError:
            cx, s, d, T = self._locate_image_do(page, xref, old_disp)
            X = T * F * ~T                                      # X x T = T x F
            self.snapshot()
            w = self._WRAP.search(s[:d.start()][-200:])
            if w and s[d.end():].lstrip().startswith("Q"):      # đã bọc từ lần dời trước: gộp vào X cũ
                start = d.start() - len(s[:d.start()][-200:]) + w.start(1)
                end = d.start() - len(s[:d.start()][-200:]) + w.end(1)
                old_x = pymupdf.Matrix(*[float(v) for v in s[start:end].split()])
                s2 = s[:start] + self._fmt_matrix(X * old_x) + " " + s[end:]
            else:
                s2 = s[:d.start()] + f"q {self._fmt_matrix(X)} cm " + s[d.start():d.end()] + " Q" + s[d.end():]
            self.doc.update_stream(cx, s2.encode("latin-1"))
            return True
        new = mat * F
        nums = self._fmt_matrix(new) + " "
        self.snapshot()
        s2 = s[:cm.start(1)] + nums + s[cm.end(1):]
        self.doc.update_stream(cx, s2.encode("latin-1"))
        return True

    def duplicate_image(self, pno, xref, old_disp, new_disp):
        """Nhân bản ảnh (dùng lại cùng đối tượng ảnh - file không nặng thêm)."""
        page = self.doc[pno]
        try:
            self._find_image_op(page, xref, old_disp)          # chỉ ảnh di chuyển được
        except PdfError:
            self._locate_image_do(page, xref, old_disp)
        self.snapshot()
        page.insert_image(disp_to_unrot(page, new_disp), xref=xref, keep_proportion=False,
                          rotate=(-page.rotation) % 360)
        return True

    def delete_image(self, pno, xref, disp_rect):
        """Xoá đúng 1 lần vẽ ảnh này trên trang (ảnh dùng ở trang khác không bị ảnh hưởng)."""
        page = self.doc[pno]
        try:
            cx, s, _cm, d, _mat = self._find_image_op(page, xref, disp_rect)
        except PdfError:
            cx, s, d, _T = self._locate_image_do(page, xref, disp_rect)
        self.snapshot()
        self.doc.update_stream(cx, (s[:d.start()] + s[d.end():]).encode("latin-1"))
        return True

    def delete_annot(self, pno, xref):
        a = self._annot(pno, xref)
        if a is not None:
            self.snapshot()
            a = self._annot(pno, xref)
            a._keep_page.delete_annot(a)

    # ----- sao chép / dán vật trên trang (chữ, ảnh, khối chú thích) -----
    def image_png(self, xref):
        """Ảnh trong file (kèm nền trong suốt nếu có) dạng PNG."""
        pix = pymupdf.Pixmap(self.doc, xref)
        sm = self.doc.extract_image(xref).get("smask") or 0
        if pix.colorspace and pix.colorspace.n not in (1, 3):
            pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
        if sm and not pix.alpha:
            try:
                pix = pymupdf.Pixmap(pix, pymupdf.Pixmap(self.doc, sm))
            except Exception:
                pass
        return pix.tobytes("png")

    def annot_png(self, pno, xref, dpi=200):
        a = self._annot(pno, xref)
        if a is None:
            raise PdfError("Không tìm thấy khối này trên trang (có thể đã bị xoá).")
        page = a._keep_page
        k = dpi / 72
        return a.get_pixmap(matrix=pymupdf.Matrix(k, k).prerotate(page.rotation), alpha=True).tobytes("png")

    def copy_annot_object(self, xref):
        """Bản sao rời (chưa gắn vào trang) của khối chú thích - để dán lại kể cả sau khi khối gốc bị cắt / xoá."""
        nx = self.doc.get_new_xref()
        self.doc.update_object(nx, "<<>>")
        self.doc.xref_copy(xref, nx)
        for k in ("Popup", "IRT", "NM", "P", "Parent"):
            self.doc.xref_set_key(nx, k, "null")
        return nx

    def paste_annot(self, pno, src_xref, new_disp):
        """Dán bản sao khối chú thích (src_xref: bản sao rời từ copy_annot_object) vào trang, góc trên trái tại
        new_disp (toạ độ hiển thị, x0, y0). Trả về xref khối mới."""
        doc = self.doc
        if doc.xref_get_key(src_xref, "Subtype")[0] != "name":
            raise PdfError("Bản sao không còn hợp lệ.")
        self.snapshot()
        page = doc[pno]
        # tạo khối mới qua MuPDF (để danh sách chú thích của trang đang mở được cập nhật - không cần nạp lại
        # trang), rồi chép nội dung từ điển của bản sao vào khối đó
        nx = page.add_rect_annot(pymupdf.Rect(0, 0, 10, 10)).xref
        src_keys = set(doc.xref_get_keys(src_xref))
        for k in doc.xref_get_keys(nx):
            if k not in src_keys and k not in ("P", "Type"):
                doc.xref_set_key(nx, k, "null")
        for k in src_keys:
            if k in ("P", "Popup", "IRT", "NM", "Parent"):
                continue
            typ, val = doc.xref_get_key(src_xref, k)
            if typ != "null":
                doc.xref_set_key(nx, k, val)
        typ, val = doc.xref_get_key(nx, "Rect")
        q = [float(v) for v in re.findall(self._NUM, val)]
        old = pymupdf.Rect(q) * page.transformation_matrix * page.rotation_matrix
        old.normalize()
        x0, y0 = new_disp[0], new_disp[1]
        self._transform_annot(pno, nx, old, pymupdf.Rect(x0, y0, x0 + old.width, y0 + old.height))
        return nx

    def paste_lines(self, pno, items, dx, dy):
        """Dán các dòng chữ đã sao chép: items = [(TextLine gốc, (html, fonts))], dời (dx, dy) theo toạ độ trang
        chưa xoay so với chỗ gốc."""
        import dataclasses
        self.snapshot(fonts=True)
        page = self.doc[pno]
        for ln, (body, fams) in items:
            bb = pymupdf.Rect(ln.bbox) + (dx, dy, dx, dy)
            bb.x1 += max(bb.width * 0.5, 2 * bb.height)
            self._place_line(page, dataclasses.replace(ln, bbox=bb, baseline=ln.baseline + dy), body, fams,
                             "left", True)

    def paste_image(self, pno, disp_rect, png=None, xref=0):
        """Dán ảnh vào đúng vùng hiển thị: dùng lại đối tượng ảnh sẵn có (xref, cùng file) hoặc ảnh PNG."""
        page = self.doc[pno]
        self.snapshot()
        r = disp_to_unrot(page, pymupdf.Rect(disp_rect))
        if xref and self.doc.xref_get_key(xref, "Subtype")[1] == "/Image":
            page.insert_image(r, xref=xref, keep_proportion=False, rotate=(-page.rotation) % 360)
        else:
            page.insert_image(r, stream=png, keep_proportion=False, rotate=(-page.rotation) % 360)

    # ----- xoá vĩnh viễn thông tin (redaction) -----
    def mark_redactions(self, rects_by_page, fill="black", disp=False):
        """Đánh dấu vùng sẽ xoá: {pno: [Rect]} (disp=True: toạ độ hiển thị). Chưa xoá gì - xem lại rồi áp dụng."""
        n = 0
        if not any(rects_by_page.values()):
            return 0
        self.snapshot()
        col = redact_fill(fill)
        for pno, rects in rects_by_page.items():
            page = self.doc[pno]
            for r in rects:
                r = disp_to_unrot(page, pymupdf.Rect(r)) if disp else pymupdf.Rect(r)
                a = page.add_redact_annot(r, fill=col)
                a.set_colors(stroke=(0.86, 0.15, 0.15))
                a.update()
                n += 1
        return n

    def pending_redactions(self):
        """Số vùng đã đánh dấu nhưng chưa xoá vĩnh viễn."""
        if not self.doc:
            return 0
        return sum(1 for page in self.doc for a in (page.annots(types=[pymupdf.PDF_ANNOT_REDACT]) or []))

    def clear_redactions(self):
        if not self.pending_redactions():
            return 0
        self.snapshot()
        n = 0
        for page in self.doc:
            for a in list(page.annots(types=[pymupdf.PDF_ANNOT_REDACT]) or []):
                page.delete_annot(a)
                n += 1
        return n

    def apply_redactions(self):
        """Xoá hẳn mọi thứ dưới các vùng đã đánh dấu: chữ, điểm ảnh của ảnh, hình vẽ nằm trong vùng; tô màu đã chọn."""
        if not self.pending_redactions():
            return 0
        self.snapshot(fonts=True)
        return apply_redactions_doc(self.doc)

    def arrange(self, pno, kind, xref, disp_rect, front=True):
        """Đưa vật lên trên cùng / xuống dưới cùng. Ảnh: vẽ lại trên / dưới nội dung trang; khối chú thích: đổi
        thứ tự trong danh sách chú thích của trang."""
        page = self.doc[pno]
        if kind == "img":
            try:
                cx, s, _cm, d, _mat = self._find_image_op(page, xref, disp_rect)
            except PdfError:
                cx, s, d, _T = self._locate_image_do(page, xref, disp_rect)
            self.snapshot()
            self.doc.update_stream(cx, (s[:d.start()] + s[d.end():]).encode("latin-1"))
            page = self.doc[pno]
            page.insert_image(disp_to_unrot(page, pymupdf.Rect(disp_rect)), xref=xref, keep_proportion=False,
                              overlay=front, rotate=(-page.rotation) % 360)
            return True
        refs = [x[0] for x in page.annot_xrefs()]
        if xref not in refs:
            raise PdfError("Không tìm thấy khối này trên trang (có thể đã bị xoá).")
        refs.remove(xref)
        refs = refs + [xref] if front else [xref] + refs
        self.snapshot()
        self.doc.xref_set_key(page.xref, "Annots", "[" + " ".join(f"{x} 0 R" for x in refs) + "]")
        return True

    # ----- khối vẽ / chú thích đã chèn (hình chữ nhật, elip, đường, mũi tên, hộp chữ...): kéo để di chuyển, đổi cỡ -----
    MOVABLE_ANNOTS = {pymupdf.PDF_ANNOT_SQUARE, pymupdf.PDF_ANNOT_CIRCLE, pymupdf.PDF_ANNOT_LINE,
                      pymupdf.PDF_ANNOT_POLY_LINE, pymupdf.PDF_ANNOT_POLYGON, pymupdf.PDF_ANNOT_INK,
                      pymupdf.PDF_ANNOT_FREE_TEXT, pymupdf.PDF_ANNOT_STAMP, pymupdf.PDF_ANNOT_FILE_ATTACHMENT,
                      pymupdf.PDF_ANNOT_CARET}
    # có giao diện tự sinh lại được sau khi đổi toạ độ (Stamp có thể là ảnh riêng -> chỉ dời /Rect)
    _REGEN = MOVABLE_ANNOTS - {pymupdf.PDF_ANNOT_STAMP}

    def page_annots(self, pno):
        """Khối chú thích kéo được trên trang: [(vùng hiển thị, xref)]."""
        page = self.doc[pno]
        out = []
        for a in page.annots() or []:
            if a.type[0] in self.MOVABLE_ANNOTS:
                r = a.rect * page.rotation_matrix
                r.normalize()
                if not r.is_empty:
                    out.append((r, a.xref))
        return out

    def move_annot(self, pno, xref, old_disp, new_disp):
        """Dời / co giãn chú thích từ vùng hiển thị cũ sang vùng mới: biến đổi trực tiếp các toạ độ trong
        từ điển PDF (/Rect, /L, /Vertices, /InkList, /CL) rồi cho MuPDF vẽ lại giao diện."""
        page = self.doc[pno]
        if xref not in [x[0] for x in page.annot_xrefs()]:
            raise PdfError("Không tìm thấy khối này trên trang (có thể đã bị xoá).")
        self.snapshot()
        return self._transform_annot(pno, xref, old_disp, new_disp)

    def _transform_annot(self, pno, xref, old_disp, new_disp):
        page = self.doc[pno]
        o, n = disp_to_unrot(page, old_disp), disp_to_unrot(page, new_disp)
        if o.width < 0.5 or o.height < 0.5:
            sx = sy = 1.0
        else:
            sx, sy = n.width / o.width, n.height / o.height
        mu = pymupdf.Matrix(sx, 0, 0, sy, n.x0 - o.x0 * sx, n.y0 - o.y0 * sy)   # trong toạ độ trang chưa xoay
        t = page.transformation_matrix                                           # PDF (y lên) -> MuPDF
        mr = t * mu * ~t                                                         # cùng phép biến đổi trên toạ độ PDF

        def pts(vals):
            out = []
            for i in range(0, len(vals) - 1, 2):
                p = pymupdf.Point(vals[i], vals[i + 1]) * mr
                out += [p.x, p.y]
            return out

        def nums(s):
            return [float(x) for x in re.findall(self._NUM, s or "")]

        def arr(v):
            return "[" + " ".join(f"{x:.3f}" for x in v) + "]"

        doc = self.doc
        kind = doc.xref_get_key(xref, "Subtype")[1]
        typ, val = doc.xref_get_key(xref, "Rect")
        r = nums(val)
        if len(r) == 4:
            a = pymupdf.Point(r[0], r[1]) * mr
            b = pymupdf.Point(r[2], r[3]) * mr
            doc.xref_set_key(xref, "Rect", arr([min(a.x, b.x), min(a.y, b.y), max(a.x, b.x), max(a.y, b.y)]))
        for key in ("L", "Vertices", "CL"):
            typ, val = doc.xref_get_key(xref, key)
            if typ == "array":
                doc.xref_set_key(xref, key, arr(pts(nums(val))))
        typ, val = doc.xref_get_key(xref, "InkList")
        if typ == "array":
            strokes = re.findall(r"\[([^\[\]]*)\]", val)
            doc.xref_set_key(xref, "InkList", "[" + " ".join(arr(pts(nums(s))) for s in strokes) + "]")
        if kind == "/Link":                          # chữ hiển thị đi kèm liên kết: sửa ma trận cm của khối chữ
            op = self._link_text_op(page, xref)
            if op:
                cx, s, m = op
                mat = pymupdf.Matrix(*[float(x) for x in m.group(1).split()])
                new = mat * mr
                txt = " ".join(f"{v:.4f}".rstrip("0").rstrip(".") or "0" for v in tuple(new)) + " "
                self.doc.update_stream(cx, (s[:m.start(1)] + txt + s[m.end(1):]).encode("latin-1"))
        elif kind != "/Stamp":
            a = self._annot(pno, xref)
            if a is not None and a.type[0] in self._REGEN:
                a.update()
        return True

    def add_shape(self, pno, kind, p1, p2, stroke=(0.86, 0.15, 0.15), fill=None, width=1.5, arrow=False):
        """kind: rect | ellipse | line. p1, p2 theo toạ độ hiển thị."""
        page = self.doc[pno]
        self.snapshot()
        m = page.derotation_matrix
        if kind == "line":
            a = page.add_line_annot(pymupdf.Point(p1) * m, pymupdf.Point(p2) * m)
            if arrow:
                a.set_line_ends(pymupdf.PDF_ANNOT_LE_NONE, pymupdf.PDF_ANNOT_LE_OPEN_ARROW)
        else:
            r = disp_to_unrot(page, pymupdf.Rect(p1, p2))
            a = page.add_rect_annot(r) if kind == "rect" else page.add_circle_annot(r)
        a.set_border(width=width)
        if width <= 0 and fill:
            a.set_colors(stroke=fill, fill=fill)
        elif fill:
            a.set_colors(stroke=stroke, fill=fill)
        else:
            a.set_colors(stroke=stroke)
        a.update()

    # ----------------------------- chú thích (Comment) -----------------------------
    MARKUPS = {"highlight": "add_highlight_annot", "underline": "add_underline_annot",
               "strikeout": "add_strikeout_annot", "squiggly": "add_squiggly_annot"}
    MARKUP_COLORS = {"highlight": (1, 0.92, 0.23), "underline": (0.1, 0.45, 0.9),
                     "strikeout": (0.86, 0.15, 0.15), "squiggly": (0.86, 0.15, 0.15)}

    def words_in(self, pno, disp_rect):
        """Ô chữ (toạ độ trang chưa xoay) của các từ nằm >= 50% trong vùng hiển thị."""
        page = self.doc[pno]
        un = disp_to_unrot(page, disp_rect)
        hits = []
        for w in page.get_text("words"):
            wr = pymupdf.Rect(w[:4])
            inter = wr & un
            if not inter.is_empty and inter.get_area() >= 0.5 * wr.get_area():
                hits.append(wr)
        return hits

    def add_markup(self, pno, kind, word_rects=None, disp_rect=None, color=None, author="", note=""):
        """Tô sáng / gạch chân / gạch ngang / gạch sóng các từ (word_rects: toạ độ chưa xoay) hoặc theo vùng."""
        page = self.doc[pno]
        rects = list(word_rects or [])
        if not rects and disp_rect is not None:
            rects = self.words_in(pno, disp_rect)
        if not rects and disp_rect is None:
            raise PdfError("Chưa chọn chữ nào.")
        self.snapshot()
        fn = getattr(page, self.MARKUPS[kind])
        if rects:
            a = fn(quads=[r.quad for r in _merge_line_rects(rects)])
        else:
            a = fn(disp_to_unrot(page, pymupdf.Rect(disp_rect)))
        a.set_colors(stroke=color or self.MARKUP_COLORS[kind])
        now = pymupdf.get_pdf_now()
        a.set_info(title=author, content=note, creationDate=now, modDate=now)
        a.update()
        return len(rects)

    def add_caret(self, pno, disp_point, text, author=""):
        """Đề xuất chèn chữ (dấu ^ tại vị trí + nội dung)."""
        page = self.doc[pno]
        self.snapshot()
        a = page.add_caret_annot(pymupdf.Point(disp_point) * page.derotation_matrix)
        now = pymupdf.get_pdf_now()
        a.set_info(title=author, content=text, creationDate=now, modDate=now)
        a.set_colors(stroke=(0.1, 0.45, 0.9))
        a.update()

    def add_textbox(self, pno, disp_rect, text, fontsize=11, color=(0, 0, 0), fill=None, border=1.0,
                    callout=None, author=""):
        """Hộp chữ (FreeText): border=0 -> kiểu máy chữ (Typewriter). callout: [điểm mũi tên, điểm gập] (hiển thị)."""
        page = self.doc[pno]
        m = page.derotation_matrix
        r = disp_to_unrot(page, pymupdf.Rect(disp_rect))
        self.snapshot()
        kw = dict(fontsize=fontsize, fontname="helv", text_color=color, border_width=border,
                  rotate=page.rotation, fill_color=fill)
        if callout:
            kw["callout"] = [pymupdf.Point(p) * m for p in callout]
            kw["line_end"] = pymupdf.PDF_ANNOT_LE_OPEN_ARROW
        a = page.add_freetext_annot(r, text, **kw)
        now = pymupdf.get_pdf_now()
        a.set_info(title=author, creationDate=now, modDate=now)
        a.update()

    @staticmethod
    def textbox_size(text, fontsize):
        """Cỡ hộp chữ vừa nội dung (pt)."""
        lines = text.splitlines() or [""]
        w = max(pymupdf.get_text_length(ln, fontname="helv", fontsize=fontsize) for ln in lines)
        return w * 1.12 + fontsize + 6, len(lines) * fontsize * 1.35 + fontsize * 0.6 + 4

    def add_ink(self, pno, strokes, color=(0.86, 0.15, 0.15), width=1.5, opacity=1.0):
        """Nét vẽ tự do (Pencil / bút dạ quang). strokes: [[(x, y) hiển thị]]."""
        page = self.doc[pno]
        m = page.derotation_matrix
        pts = [[tuple(pymupdf.Point(p) * m) for p in s] for s in strokes if len(s) >= 2]
        if not pts:
            return
        self.snapshot()
        a = page.add_ink_annot(pts)
        a.set_border(width=width)
        a.set_colors(stroke=color)
        a.set_opacity(opacity)
        a.update()

    def add_poly(self, pno, points, kind="polygon", color=(0.86, 0.15, 0.15), width=1.5, fill=None, clouds=0):
        """Đa giác / đường gấp khúc / đám mây (clouds > 0). points: [(x, y) hiển thị]."""
        page = self.doc[pno]
        m = page.derotation_matrix
        pts = [pymupdf.Point(p) * m for p in points]
        if len(pts) < 2:
            return
        self.snapshot()
        a = page.add_polygon_annot(pts) if kind == "polygon" else page.add_polyline_annot(pts)
        a.set_border(width=width, clouds=clouds) if clouds else a.set_border(width=width)
        a.set_colors(stroke=color, fill=fill) if fill else a.set_colors(stroke=color)
        a.update()

    def add_file_annot(self, pno, disp_point, path, author=""):
        """Ghim file đính kèm tại 1 điểm trên trang (biểu tượng kẹp giấy)."""
        page = self.doc[pno]
        with open(path, "rb") as f:
            data = f.read()
        self.snapshot()
        name = os.path.basename(path)
        a = page.add_file_annot(pymupdf.Point(disp_point) * page.derotation_matrix, data, name,
                                ufilename=name, desc=name, icon="Paperclip")
        a.set_info(title=author, content=name)
        a.update()

    def annot_at(self, pno, disp_point):
        """xref chú thích trên cùng tại điểm (bỏ liên kết, ô form)."""
        page = self.doc[pno]
        pt = pymupdf.Point(disp_point)
        best = None
        for a in page.annots() or []:
            r = a.rect * page.rotation_matrix
            r.normalize()
            if (r + (-3, -3, 3, 3)).contains(pt):
                best = a.xref                         # annot sau nằm trên
        return best

    def list_annots(self):
        """Tất cả chú thích trong file (cho bảng danh sách ghi chú)."""
        out = []
        skip = {pymupdf.PDF_ANNOT_LINK, pymupdf.PDF_ANNOT_WIDGET, pymupdf.PDF_ANNOT_POPUP}
        for pno, page in enumerate(self.doc):
            for a in page.annots() or []:
                if a.type[0] in skip:
                    continue
                i = a.info
                text = (i.get("content") or "").strip()
                if a.type[0] in (pymupdf.PDF_ANNOT_HIGHLIGHT, pymupdf.PDF_ANNOT_UNDERLINE,
                                 pymupdf.PDF_ANNOT_STRIKE_OUT, pymupdf.PDF_ANNOT_SQUIGGLY) and not text:
                    try:
                        text = page.get_textbox(a.rect).strip()      # chữ được đánh dấu
                    except Exception:
                        pass
                d = i.get("modDate") or i.get("creationDate") or ""
                mm = re.match(r"D:(\d{4})(\d{2})(\d{2})(\d{2})?(\d{2})?", d)
                date = f"{mm[3]}/{mm[2]}/{mm[1]}" + (f" {mm[4]}:{mm[5]}" if mm[4] and mm[5] else "") if mm else ""
                out.append(dict(page=pno, xref=a.xref, kind=a.type[1], author=i.get("title", ""), date=date,
                                text=text.replace("\r", "\n")))
        return out

    def delete_annot_any(self, pno, xref):
        """Xoá chú thích bất kỳ theo xref (dùng cho danh sách ghi chú / cục tẩy)."""
        self.delete_annot(pno, xref)

    # ----- liên kết hàng loạt -----
    URL_RE = re.compile(r"(?i)\b((?:https?://|www\.)[^\s<>\"']+[^\s<>\"'.,;:!?)\]]|"
                        r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+)")

    def auto_links(self, pages=None):
        """Biến địa chỉ web / email dạng chữ thành liên kết bấm được. Trả về số liên kết đã tạo."""
        self.snapshot()
        n = 0
        for pno in pages if pages is not None else range(self.page_count):
            page = self.doc[pno]
            have = [pymupdf.Rect(l["from"]) * page.derotation_matrix for l in page.get_links()]
            for blk in page.get_text("dict")["blocks"]:
                for line in blk.get("lines", []):
                    for sp in line["spans"]:
                        for m in self.URL_RE.finditer(sp["text"]):
                            u = m.group(1)
                            hits = page.search_for(u, clip=pymupdf.Rect(sp["bbox"]) + (-1, -1, 1, 1))
                            for h in hits[:1]:
                                if any((h & x).get_area() > 0.5 * h.get_area() for x in have):
                                    continue
                                uri = u if "://" in u else ("mailto:" + u if "@" in u and "/" not in u else "https://" + u)
                                page.insert_link({"kind": pymupdf.LINK_URI, "from": h, "uri": uri})
                                have.append(h)
                                n += 1
        if not n:
            self._undo.pop()
        return n

    def remove_web_links(self, pages=None):
        """Xoá mọi liên kết web (giữ liên kết trong file). Trả về số đã xoá."""
        self.snapshot()
        n = 0
        for pno in pages if pages is not None else range(self.page_count):
            page = self.doc[pno]
            for l in page.get_links():
                if l.get("kind") == pymupdf.LINK_URI:
                    op = self._link_text_op(page, l["xref"]) if l.get("xref") else None
                    if op:
                        cx, s, mt = op
                        self.doc.update_stream(cx, (s[:mt.start(2)] + s[mt.end(2):]).encode("latin-1"))
                    page.delete_link(l)
                    n += 1
        if not n:
            self._undo.pop()
        return n

    # ----- trang -----
    def duplicate_pages(self, pages):
        """Nhân bản các trang chọn (bản sao nằm ngay sau trang gốc). Trả về chỉ số các bản sao."""
        self.snapshot()
        out = []
        for k, p in enumerate(sorted(pages)):
            src = p + k                               # trang trước đã chèn bản sao -> lệch
            self.doc.fullcopy_page(src, src + 1)
            out.append(src + 1)
        return out

    def swap_pages(self, a, b):
        if a == b:
            return
        self.snapshot()
        a, b = sorted((a, b))
        order = list(range(self.page_count))
        order[a], order[b] = order[b], order[a]
        self.doc.select(order)

    def replace_pages(self, pages, src_path, src_start=0, password=None):
        """Thay các trang chọn bằng các trang lấy từ file khác (bắt đầu từ trang src_start)."""
        src = self._load(src_path, password)
        try:
            pages = sorted(pages)
            if src_start + len(pages) > src.page_count:
                raise PdfError(f"File nguồn chỉ có {src.page_count} trang - không đủ {len(pages)} trang để thay.")
            self.snapshot()
            for k, p in enumerate(pages):
                self.doc.insert_pdf(src, from_page=src_start + k, to_page=src_start + k, start_at=p)
                self.doc.delete_page(p + 1)
        finally:
            src.close()

    def find_duplicate_pages(self):
        """Nhóm các trang giống hệt nhau (cùng chữ + gần như cùng hình ảnh). [[p1, p2, ...], ...]"""
        import hashlib
        groups = {}
        for p, page in enumerate(self.doc):
            pix = page.get_pixmap(matrix=pymupdf.Matrix(0.25, 0.25), colorspace=pymupdf.csGRAY, alpha=False)
            sig = bytes(v // 32 for v in pix.samples)            # bỏ khác biệt nhỏ do khử răng cưa
            key = hashlib.md5(page.get_text("text").strip().encode("utf-8") + sig
                              + f"{pix.width}x{pix.height}".encode()).hexdigest()
            groups.setdefault(key, []).append(p)
        return [g for g in groups.values() if len(g) > 1]

    # ----------------------------- chữ ký số -----------------------------
    def has_signatures(self):
        try:
            for p in self.doc:
                for _ in p.widgets(types=[pymupdf.PDF_WIDGET_TYPE_SIGNATURE]):
                    return True
        except Exception:
            pass
        return False

    def pdf_box(self, pno, disp_rect):
        """Vùng hiển thị -> toạ độ PDF chuẩn (gốc dưới-trái) dùng cho trường chữ ký."""
        page = self.doc[pno]
        r = disp_to_unrot(page, disp_rect) * ~page.transformation_matrix
        r.normalize()
        return r.x0, r.y0, r.x1, r.y1

    @staticmethod
    def _pattern(needle, match_case):
        return re.compile(re.escape(normalize(needle)), 0 if match_case else re.IGNORECASE)

    def find_text(self, needle, match_case=False, pages=None):
        """Đếm số lần xuất hiện theo trang: {pno: count}."""
        res = {}
        if not needle:
            return res
        pat = self._pattern(needle, match_case)
        for pno in (pages if pages is not None else range(self.page_count)):
            n = sum(len(pat.findall(ln.text)) for ln in get_text_lines(self.doc[pno]))
            if n:
                res[pno] = n
        return res

    def find_replace(self, needle, repl, match_case=False, fit=True, pages=None):
        """Tìm và thay thế. Trả về số chỗ đã thay."""
        if not needle:
            return 0
        pat = self._pattern(needle, match_case)
        targets = self.find_text(needle, match_case, pages)
        if not targets:
            return 0
        self.snapshot(fonts=True)
        count = 0
        for pno in targets:
            page = self.doc[pno]
            jobs = []
            for ln in get_text_lines(page):
                n = len(pat.findall(ln.text))
                if not n:
                    continue
                if ln.uniform:
                    # Cả dòng một kiểu chữ: viết lại cả dòng để phần sau tự dịch theo
                    jobs.append(("line", ln, pat.sub(lambda m: repl, ln.text)))
                    count += n
                else:
                    # Dòng nhiều kiểu chữ: chỉ thay đúng vùng tìm thấy, giữ kiểu chữ của vùng đó
                    for m in pat.finditer(ln.text):
                        jobs.append(("hit", ln.match_rect(m.start(), m.end()), ln.chars[m.start()][1]))
                        count += 1
            if not jobs:
                continue
            _redact(page, [j[1].bbox if j[0] == "line" else j[1] for j in jobs])
            for kind, obj, extra in jobs:
                if kind == "line":
                    if extra.strip():
                        _write(page, obj.bbox.x0, obj.baseline, extra, obj.size, obj.family, obj.bold, obj.italic,
                               int_to_rgb(obj.color), max_width=obj.bbox.width if fit else None)
                elif repl:
                    sp = extra
                    _write(page, obj.x0, sp["origin"][1], repl, sp["size"],
                           font_family_of(sp["font"], sp["flags"]), is_bold(sp["font"], sp["flags"]),
                           is_italic(sp["font"], sp["flags"]), int_to_rgb(sp["color"]),
                           max_width=obj.width if fit else None)
        return count

# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
"""Bộ lọc loại file cho hộp thoại chọn / lưu file của Windows (Explorer - giữ các thư mục người dùng đã pin).

Bộ lọc gộp nhiều định dạng (vd. "Word / Excel / PowerPoint → PDF (*.doc *.docx *.xls ... 21 đuôi)") rất dài và khó
chọn -> tách thêm mỗi nhóm một dòng: Tất cả định dạng hỗ trợ / Word / Excel / PowerPoint / ..."""
import re

from PySide6.QtWidgets import QFileDialog

import i18n

_orig = {}

KINDS = [
    ("PDF", ("pdf",)),
    ("Word", ("doc", "docx", "docm", "dot", "dotx", "dotm", "odt", "rtf", "wpd", "wps")),
    ("Excel", ("xls", "xlsx", "xlsm", "xlsb", "xlt", "xltx", "ods", "csv")),
    ("PowerPoint", ("ppt", "pptx", "pptm", "pps", "ppsx", "pot", "potx", "odp")),
    ("Ảnh", ("png", "jpg", "jpeg", "bmp", "gif", "tif", "tiff", "webp", "heic", "jfif", "ico", "svg")),
    ("Văn bản", ("txt", "md", "html", "htm", "xml", "json")),
    ("Email", ("eml", "msg")),
]


def split_filter(flt):
    """'Office (*.doc *.xls *.ppt ...)' -> 'Tất cả định dạng hỗ trợ (...);;Word (*.doc);;Excel (*.xls);;...'.
    Bộ lọc chỉ có 1 nhóm định dạng giữ nguyên."""
    out = []
    for part in (flt or "").split(";;"):
        part = part.strip()
        if not part:
            continue
        m = re.search(r"\(([^)]*)\)\s*$", part)
        pats = [p.lower() for p in m.group(1).split()] if m else []
        exts = {p[2:] for p in pats if p.startswith("*.")}
        kinds = [(k, [f"*.{e}" for e in es if e in exts]) for k, es in KINDS]
        kinds = [(k, ps) for k, ps in kinds if ps]
        if len(kinds) < 2:
            out.append(part)
            continue
        out.append(f"{i18n.tr('Tất cả định dạng hỗ trợ')} ({' '.join(pats)})")
        known = set()
        for k, ps in kinds:
            known.update(ps)
            out.append(f"{i18n.tr(k)} ({' '.join(ps)})")
        rest = [p for p in pats if p not in known]
        if rest:
            out.append(f"{i18n.tr('Khác')} ({' '.join(rest)})")
    return ";;".join(out)


def _wrap(name):
    fn = _orig[name]

    def call(*a, **kw):
        a = list(a)
        if "filter" in kw:
            kw["filter"] = split_filter(i18n.tr(kw["filter"] or ""))
        elif len(a) >= 4 and a[3]:
            a[3] = split_filter(i18n.tr(a[3]))
        return fn(*a, **kw)
    return call


def install():
    """Bọc QFileDialog.getOpenFileName(s) / getSaveFileName để tách bộ lọc (gọi sau i18n.install)."""
    if _orig:
        return
    for name in ("getOpenFileName", "getOpenFileNames", "getSaveFileName"):
        _orig[name] = getattr(QFileDialog, name)
        setattr(QFileDialog, name, staticmethod(_wrap(name)))

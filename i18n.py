# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Đa ngôn ngữ VI / EN.

Mã nguồn viết chữ tiếng Việt trực tiếp; chữ tiếng Việt chính là khoá tra từ điển (i18n_en.py).
install() bọc các hàm hiển thị chữ của Qt (setText, setToolTip, addItem, QMessageBox...) để tự dịch khi
ngôn ngữ là "en". Chữ ghép động (f-string) được dịch theo mẫu: "Trang {0}/{1}" -> "Page {0}/{1}".
Đổi ngôn ngữ: set_lang() rồi dựng lại cửa sổ (chữ chỉ được dịch lúc gán vào widget).
"""
import re

LANG = "vi"
EXACT = {}
PATTERNS = []          # [(regex đã biên dịch, mẫu tiếng Anh)]
ASCII_PATTERNS = []    # mẫu không có chữ có dấu (thử cả với chuỗi không dấu)
ML_PATTERNS = []       # mẫu nhiều dòng - thử trước khi tách chuỗi theo dòng
_VI = re.compile(r"[À-ỹĐđ]")
_EDGE = re.compile(r"^([\s✔✘→•·*\-–⚠]*)(.*?)([\s:.…]*)$", re.S)
_FILTER = re.compile(r"^(.*?) (\([*.;\w\s]+\))$", re.S)
_SEPS = ("  (", "  |  ", ": ", " | ", " · ", " — ", " → ", "; ", ", ", " - ")
_loaded = False


def _load():
    global _loaded
    if _loaded:
        return
    _loaded = True
    try:
        from i18n_en import EXACT as E, PATTERNS as P
    except ImportError:
        return
    EXACT.update(E)
    pats = []
    for vi, en in P.items():
        parts = re.split(r"(\{\d+\})", vi)
        literal = "".join(p for p in parts if not re.fullmatch(r"\{\d+\}", p))
        if sum(c.isalpha() for c in literal) < 3:          # mẫu quá chung chung -> bỏ
            continue
        rx = "".join("(.+?)" if re.fullmatch(r"\{\d+\}", p) else re.escape(p.replace("{{", "{").replace("}}", "}"))
                     for p in parts)
        pats.append((len(literal), re.compile(rx, re.S), en, "\n" in literal))
    pats.sort(key=lambda x: -x[0])                         # mẫu dài (cụ thể) thử trước
    PATTERNS[:] = [(rx, en) for _, rx, en, _ml in pats]
    ML_PATTERNS[:] = [(rx, en) for _, rx, en, ml in pats if ml]
    ASCII_PATTERNS[:] = [(rx, en) for rx, en in PATTERNS if not _VI.search(rx.pattern)]


def set_lang(lang):
    global LANG
    LANG = "en" if lang == "en" else "vi"
    if LANG == "en":
        _load()


_paused = 0


class paused:
    """with i18n.paused(): ... -> không dịch (dữ liệu người dùng: giá trị ô form PDF, tên file...)."""

    def __enter__(self):
        global _paused
        _paused += 1

    def __exit__(self, *a):
        global _paused
        _paused -= 1


def tr(s, _depth=0):
    """Dịch một chuỗi hiển thị (trả nguyên văn nếu đang dùng tiếng Việt hoặc không có bản dịch)."""
    if LANG == "vi" or _paused or not isinstance(s, str) or not s or _depth > 4:
        return s
    v = EXACT.get(s)
    if v is not None:
        return v
    if not _VI.search(s):
        for rx, en in ASCII_PATTERNS:                       # vd "Trang {0}/{1}" không có dấu
            m = rx.fullmatch(s)
            if m:
                return en.format(*m.groups())
        return s
    if "\n" in s:
        for rx, en in ML_PATTERNS:
            m = rx.fullmatch(s)
            if m:
                try:
                    return en.format(*[tr(g, _depth + 1) for g in m.groups()])
                except (IndexError, KeyError, ValueError):
                    break
        return "\n".join(tr(x, _depth + 1) for x in s.split("\n"))
    if ";;" in s:                                           # bộ lọc QFileDialog
        return ";;".join(tr(x, _depth + 1) for x in s.split(";;"))
    m = _EDGE.match(s)
    if m and (m[1] or m[3]):
        v = EXACT.get(m[2])
        if v is not None:
            return m[1] + v + m[3]
    for rx, en in PATTERNS:
        m = rx.fullmatch(s)
        if m:
            try:
                return en.format(*[tr(g, _depth + 1) for g in m.groups()])
            except (IndexError, KeyError, ValueError):
                break
    m = _FILTER.match(s)
    if m:
        return tr(m[1], _depth + 1) + " " + m[2]
    for sep in _SEPS:                                       # tách đôi ở dấu ngăn đầu tiên rồi dịch từng phần
        if sep in s:
            head, tail = s.split(sep, 1)
            out = tr(head, _depth + 1) + sep + tr(tail, _depth + 1)
            if out != s:
                return out
    return s


# ---------------------------------------------------------------------------
#  Gắn vào Qt
# ---------------------------------------------------------------------------
def _wrap_method(cls, name, idx=(0,), list_arg=False):
    orig = getattr(cls, name)

    def f(self, *a, **kw):
        a = list(a)
        for i in idx:
            if i < len(a):
                if list_arg and isinstance(a[i], (list, tuple)):
                    a[i] = [tr(x) for x in a[i]]
                elif isinstance(a[i], str):
                    a[i] = tr(a[i])
        return orig(self, *a, **kw)
    setattr(cls, name, f)


def _wrap_str_args(cls, name, static=False):
    """Dịch tham số str ĐẦU TIÊN (hàm có nhiều kiểu gọi: addItem(text, data) / addItem(icon, text, data),
    addAction, addTab, drawText...). Các str sau (vd userData của combobox) giữ nguyên."""
    orig = getattr(cls, name)

    def f(*a, **kw):
        a = list(a)
        for i, x in enumerate(a):
            if isinstance(x, str):
                a[i] = tr(x)
                break
        return orig(*a, **kw)
    setattr(cls, name, staticmethod(f) if static else f)


def _wrap_static(cls, name, idx):
    orig = getattr(cls, name)

    def f(*a, **kw):
        a = list(a)
        for i in idx:
            if i < len(a) and isinstance(a[i], str):
                a[i] = tr(a[i])
        for k in ("caption", "title", "text", "label", "filter"):
            if isinstance(kw.get(k), str):
                kw[k] = tr(kw[k])
        return orig(*a, **kw)
    setattr(cls, name, staticmethod(f))


def _tr_subclass(cls):
    def __init__(self, *a, **kw):
        cls.__init__(self, *[tr(x) if isinstance(x, str) else x for x in a], **kw)
    # cùng tên lớp gốc -> metaObject / stylesheet giữ nguyên tên
    return type(cls.__name__, (cls,), {"__init__": __init__})


_installed = False


def install(modules):
    """Bọc các hàm Qt hiển thị chữ + thay lớp widget có chữ trong constructor ở các module giao diện."""
    global _installed
    from PySide6.QtGui import QAction, QPainter
    from PySide6.QtWidgets import (QAbstractButton, QComboBox, QFileDialog, QFormLayout, QGroupBox, QInputDialog, QLabel,
                                   QLineEdit, QMenu, QMessageBox, QPlainTextEdit, QStatusBar, QTableWidget,
                                   QTabWidget, QTextEdit, QWidget, QCheckBox, QPushButton, QRadioButton)
    if not _installed:
        _installed = True
        _wrap_method(QWidget, "setToolTip")
        _wrap_method(QWidget, "setWindowTitle")
        _wrap_method(QLabel, "setText")
        _wrap_method(QAbstractButton, "setText")
        _wrap_method(QAction, "setText")
        _wrap_method(QAction, "setToolTip")
        _wrap_method(QAction, "setIconText")
        _wrap_method(QGroupBox, "setTitle")
        _wrap_str_args(QFormLayout, "addRow")
        _wrap_str_args(QFormLayout, "insertRow")
        _wrap_method(QMenu, "setTitle")
        for c in (QLineEdit, QPlainTextEdit, QTextEdit):
            _wrap_method(c, "setPlaceholderText")
        _wrap_method(QPlainTextEdit, "appendPlainText")
        _wrap_method(QStatusBar, "showMessage")
        _wrap_method(QComboBox, "addItems", list_arg=True)
        _wrap_method(QComboBox, "setItemText", idx=(1,))
        _wrap_str_args(QComboBox, "addItem")
        _wrap_str_args(QComboBox, "insertItem")
        _wrap_str_args(QTabWidget, "addTab")
        _wrap_str_args(QTabWidget, "insertTab")
        _wrap_method(QTabWidget, "setTabText", idx=(1,))
        _wrap_method(QTableWidget, "setHorizontalHeaderLabels", list_arg=True)
        _wrap_str_args(QMenu, "addAction")
        _wrap_str_args(QMenu, "addMenu")
        _wrap_str_args(QPainter, "drawText")
        _wrap_method(QMessageBox, "setText")
        _wrap_method(QMessageBox, "setInformativeText")
        for n in ("information", "warning", "critical", "question", "about"):
            _wrap_static(QMessageBox, n, (1, 2))
        for n in ("getText", "getMultiLineText", "getInt", "getDouble", "getItem"):
            _wrap_static(QInputDialog, n, (1, 2))
        for n, idx in (("getOpenFileName", (1, 3)), ("getOpenFileNames", (1, 3)), ("getSaveFileName", (1, 3)),
                       ("getExistingDirectory", (1,))):
            _wrap_static(QFileDialog, n, idx)
    subs = {c.__name__: _tr_subclass(c) for c in (QLabel, QPushButton, QCheckBox, QRadioButton, QGroupBox,
                                                    QAction, QMenu)}
    for m in modules:
        for name, sub in subs.items():
            if getattr(m, name, None) is not None:
                setattr(m, name, sub)

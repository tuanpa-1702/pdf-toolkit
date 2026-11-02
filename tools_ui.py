# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Giao diện bộ công cụ: trang chủ dạng lưới + màn hình chạy từng công cụ (tự dựng form từ pdf_tools.Opt)."""
import json
import os
import re
import shutil
import subprocess
import tempfile
import unicodedata

import pymupdf

from PySide6.QtCore import Qt, QRect, QSize, QSettings, QThread, QTimer, Signal
from PySide6.QtGui import QColor, QFont, QFontDatabase, QFontMetrics, QIcon, QImage, QKeySequence, QPainter, QPixmap, QShortcut
from PySide6.QtWidgets import (QAbstractItemView, QButtonGroup, QCheckBox, QColorDialog, QComboBox, QDialog, QDoubleSpinBox, QFileDialog,
                               QFormLayout, QFrame, QGraphicsDropShadowEffect, QHBoxLayout, QHeaderView, QLabel, QLayout, QLineEdit,
                               QListWidget, QListWidgetItem, QMenu, QMessageBox, QPlainTextEdit, QProgressBar, QPushButton,
                               QRadioButton, QScrollArea, QSizePolicy, QSpinBox, QSplitter, QTabWidget, QTableWidget,
                               QTableWidgetItem, QToolButton, QVBoxLayout, QWidget)

import i18n
import installer
import pdf_tools as T

ACCENT = "#2563eb"
GROUP_COLORS = {"organize": "#2563eb", "stamp": "#d97706", "to_pdf": "#059669", "from_pdf": "#7c3aed",
                "secure": "#dc2626", "editor": "#0891b2"}

_icon_font = None


def glyph_icon(code, color="#1e293b", mirror=False, size=64):
    """Icon vẽ từ font biểu tượng có sẵn của Windows (Segoe Fluent Icons / MDL2 Assets)."""
    global _icon_font
    if _icon_font is None:
        fams = QFontDatabase.families()
        _icon_font = next((f for f in ("Segoe Fluent Icons", "Segoe MDL2 Assets") if f in fams), "Segoe UI Symbol")
    pm = QPixmap(size, size)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.TextAntialiasing)
    f = QFont(_icon_font)
    f.setPixelSize(int(size * 0.62))
    p.setFont(f)
    p.setPen(QColor(color))
    p.drawText(pm.rect(), Qt.AlignCenter, chr(code))
    p.end()
    if mirror:
        from PySide6.QtGui import QTransform
        pm = pm.transformed(QTransform().scale(-1, 1))
    return QIcon(pm)


def badge_pixmap(code, color, size=40):
    """Ô vuông bo góc màu nhạt + biểu tượng màu đậm (icon của thẻ công cụ)."""
    pm = QPixmap(size * 2, size * 2)
    pm.setDevicePixelRatio(2)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    c = QColor(color)
    bg = QColor(c)
    bg.setAlpha(28)
    p.setBrush(bg)
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(0, 0, size, size, 10, 10)
    if _icon_font is None:
        glyph_icon(0xE700)
    f = QFont(_icon_font)
    f.setPixelSize(int(size * 0.5))
    p.setFont(f)
    p.setPen(c)
    p.drawText(QRect(0, 0, size, size), Qt.AlignCenter, chr(code))
    p.end()
    return pm


def strip_accents(s):
    s = unicodedata.normalize("NFD", s.lower()).replace("đ", "d")
    return "".join(ch for ch in s if unicodedata.category(ch) != "Mn")


def open_path(path):
    try:
        os.startfile(path)
    except OSError:
        pass


def reveal(path):
    if os.path.isdir(path):
        open_path(path)
    else:
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])


# ===========================================================================
#  Lưới thẻ: số cột theo độ rộng cửa sổ, các cột rộng bằng nhau lấp đầy hàng
# ===========================================================================
class GridFlow(QLayout):
    def __init__(self, parent=None, min_w=270, max_cols=5, spacing=14):
        super().__init__(parent)
        self.items = []
        self.min_w = min_w
        self.max_cols = max_cols
        self.setSpacing(spacing)
        self.setContentsMargins(0, 0, 0, 0)

    def addItem(self, item):
        self.items.append(item)

    def count(self):
        return len(self.items)

    def itemAt(self, i):
        return self.items[i] if 0 <= i < len(self.items) else None

    def takeAt(self, i):
        return self.items.pop(i) if 0 <= i < len(self.items) else None

    def expandingDirections(self):
        return Qt.Orientations(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, w):
        return self._do(QRect(0, 0, w, 0), True)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._do(rect, False)

    def sizeHint(self):
        return QSize(self.min_w, self.heightForWidth(self.min_w))

    def minimumSize(self):
        return QSize(self.min_w, 0)

    def _do(self, rect, test):
        """Chia cột đều; các thẻ cùng hàng cao bằng thẻ cao nhất (thẻ có chữ xuống dòng)."""
        sp = self.spacing()
        shown = [it for it in self.items if not (it.widget() and it.widget().isHidden())]
        if not shown:
            return 0
        cols = max(1, min(self.max_cols, (rect.width() + sp) // (self.min_w + sp)))
        cw = (rect.width() - sp * (cols - 1)) / cols
        y = rect.y()
        for r in range(0, len(shown), cols):
            row = shown[r:r + cols]
            h = max((it.heightForWidth(int(cw)) if it.hasHeightForWidth() else it.sizeHint().height())
                    for it in row)
            h = max(h, 76)
            if not test:
                for c, it in enumerate(row):
                    x = rect.x() + round(c * (cw + sp))
                    it.setGeometry(QRect(x, y, int(cw), h))
            y += h + sp
        return y - sp - rect.y()


class ToolCard(QFrame):
    clicked = Signal(str)

    def __init__(self, tid, name, desc, icon, color):
        super().__init__()
        self.tid = tid
        self.search = strip_accents(f"{name} {desc} {i18n.tr(name)} {i18n.tr(desc)}")
        self.setObjectName("card")
        self.setCursor(Qt.PointingHandCursor)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(16, 14, 16, 14)
        lay.setSpacing(14)
        ic = QLabel()
        ic.setPixmap(badge_pixmap(icon, color))
        ic.setFixedSize(40, 40)
        lay.addWidget(ic, 0, Qt.AlignTop)
        col = QVBoxLayout()
        col.setSpacing(4)
        t = QLabel(name)
        t.setWordWrap(True)                         # tên / mô tả dài: xuống dòng, thẻ tự cao lên
        t.setStyleSheet("font-weight:600; font-size:10.5pt; color:#0f172a; background:transparent")
        d = QLabel(desc)
        d.setWordWrap(True)
        d.setStyleSheet("color:#64748b; font-size:9pt; background:transparent")
        d.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        col.addWidget(t)
        col.addWidget(d, 1)
        lay.addLayout(col, 1)
        self.setToolTip(desc)
        sp = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        sp.setHeightForWidth(True)                  # rộng theo cột của lưới, cao theo số dòng chữ
        self.setSizePolicy(sp)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.LeftButton and self.rect().contains(e.position().toPoint()):
            self.clicked.emit(self.tid)


def version_row(ver, author):
    """Dòng 'V1.3  by TuanPA' đặt dưới tên app (trang chủ, hộp thoại Hỗ trợ)."""
    row = QHBoxLayout()
    text = " ".join(x for x in (ver, f"by {author}" if author else "") if x)
    if text:
        v = QLabel()
        v.setObjectName("versionLabel")
        with i18n.paused():
            v.setText(text)
        row.addWidget(v, 0, Qt.AlignVCenter)
    row.addStretch()
    return row


EDITOR_CARDS = [
    ("editor:edit", "Sửa PDF", "Xem, sửa chữ kiểu Word, chèn ảnh / liên kết / ghi chú / hình vẽ", 0xE70F),
    ("editor:pages", "Sắp xếp trang trực quan", "Kéo thả thumbnail để đổi thứ tự, xoá, xoay, gộp trang", 0xF0E2),
    ("editor:sign", "Ký số PDF", "Ký bằng USB Token (VNPT-CA, Viettel-CA...), đặt vị trí chữ ký", 0xEE56),
    ("editor:scan", "Máy scan → PDF", "Quét tài liệu giấy thành PDF: mặt kính hoặc khay nạp giấy tự động", 0xE8FE),
    ("verify_sig", "Kiểm tra chữ ký số", "Người ký, thời điểm, file có bị sửa sau khi ký không", 0xEB95),
]


class HomePage(QWidget):
    toolChosen = Signal(str)
    langChosen = Signal(str)
    aboutRequested = Signal(int)        # 0 = thông tin & hỗ trợ, 1 = giấy phép bên thứ ba
    resumeRequested = Signal(int)       # quay lại tab tài liệu (-1 = tab đang chọn)
    supportMail = Signal()

    def __init__(self, title, subtitle=""):
        super().__init__()
        self.cards = []
        self.sections = []
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.cat = "all"                               # danh mục đang lọc
        head = QWidget()
        head.setObjectName("homeHead")
        hl = QHBoxLayout(head)
        hl.setContentsMargins(28, 16, 28, 14)
        tl = QVBoxLayout()
        tl.setSpacing(4)
        row = QHBoxLayout()
        row.setSpacing(10)
        m = re.match(r"(.*?)\s+(V[\d.]+)(?:\s+by\s+(.+))?$", title)
        name, ver, author = m.groups() if m else (title, "", None)
        h1 = QLabel()
        h1.setObjectName("homeTitle")
        with i18n.paused():                            # tên app / phiên bản: không dịch
            h1.setText(name)
        row.addWidget(h1)
        row.addStretch()
        tl.addLayout(row)
        tl.addLayout(version_row(ver, author))         # dòng 2: phiên bản + tác giả, nhỏ hơn tên app
        self.sub_labels = []
        if subtitle:
            sr = QHBoxLayout()
            sr.setSpacing(6)
            dot = QLabel("●")
            dot.setStyleSheet("color:#22c55e; font-size:8pt")
            s = QLabel(subtitle)
            s.setStyleSheet("color:#64748b")
            sr.addWidget(dot)
            sr.addWidget(s)
            sr.addStretch()
            tl.addLayout(sr)
            self.sub_labels = [dot, s]
        hl.addLayout(tl, 1)
        self.head, self.head_lay = head, hl
        self.search = QLineEdit()
        self.search.setPlaceholderText("Tìm nhanh công cụ...  (vd: gộp, nén, word, mật khẩu)")
        self.search.setClearButtonEnabled(True)
        self.search.setMinimumWidth(220)
        self.search.setMaximumWidth(440)
        self.search.setMinimumHeight(38)
        self.search.addAction(glyph_icon(0xE721, "#64748b"), QLineEdit.LeadingPosition)
        kbd = QLabel("Ctrl + K")
        kbd.setObjectName("kbd")
        kl = QHBoxLayout(self.search)
        kl.setContentsMargins(0, 0, 10, 0)              # chừa viền phải + để ô gọn giữa chiều cao
        kl.addStretch()
        kl.addWidget(kbd, 0, Qt.AlignVCenter)
        self.search.setTextMargins(0, 0, 72, 0)
        self.search.textChanged.connect(lambda t: kbd.setVisible(not t))   # nhường chỗ cho nút xoá
        self.search.textChanged.connect(self.filter)
        QShortcut(QKeySequence("Ctrl+K"), self, self._focus_search, context=Qt.WidgetWithChildrenShortcut)
        hl.addWidget(self.search, 1, Qt.AlignVCenter)
        hl.addSpacing(12)
        # quay lại file đang mở (chỉ hiện khi có tab tài liệu); mũi tên bên phải = chọn tab
        self.resume_btn = _CompactToolButton()
        self.resume_btn.setObjectName("resumeBtn")
        self.resume_btn.setIcon(glyph_icon(0xE8A5, "#ffffff"))
        self.resume_btn.setIconSize(QSize(18, 18))
        self.resume_btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.resume_btn.setPopupMode(QToolButton.MenuButtonPopup)
        self.resume_btn.setCursor(Qt.PointingHandCursor)
        self.resume_btn.setToolTip("Quay lại file đang làm việc (Esc)")
        self.resume_btn.clicked.connect(lambda: self.resumeRequested.emit(-1))
        self.resume_btn.setMenu(QMenu(self.resume_btn))
        self.resume_btn.hide()
        QShortcut(QKeySequence("Esc"), self, lambda: self.resume_btn.isVisible() and self.resumeRequested.emit(-1),
                  context=Qt.WidgetWithChildrenShortcut)
        hl.addWidget(self.resume_btn, 0, Qt.AlignVCenter)
        hl.addSpacing(8)
        lang = QFrame()
        lang.setObjectName("langSwitch")
        lang.setToolTip("Ngôn ngữ giao diện")
        ll = QHBoxLayout(lang)
        ll.setContentsMargins(3, 3, 3, 3)
        ll.setSpacing(2)
        self.lang_btns = {}
        grp = QButtonGroup(self)
        for code, text in (("vi", "VI"), ("en", "EN")):
            b = QToolButton()
            b.setText(text)
            b.setCheckable(True)
            b.setObjectName("langBtn")
            b.setCursor(Qt.PointingHandCursor)
            b.setToolTip("Tiếng Việt" if code == "vi" else "English")
            b.clicked.connect(lambda _=False, c=code: self.langChosen.emit(c))
            grp.addButton(b)
            ll.addWidget(b)
            self.lang_btns[code] = b
        self.set_lang_buttons(i18n.LANG)
        hl.addWidget(lang, 0, Qt.AlignVCenter)
        hl.addSpacing(8)
        self.setup_btn = _CompactToolButton()          # menu Cài đặt (gắn menu ở MainWindow)
        self.setup_btn.setObjectName("setupBtn")
        self.setup_btn.setIcon(glyph_icon(0xE713, "#334155"))
        self.setup_btn.setIconSize(QSize(18, 18))
        self.setup_btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.setup_btn.setPopupMode(QToolButton.InstantPopup)
        self.setup_btn.setCursor(Qt.PointingHandCursor)
        hl.addWidget(self.setup_btn, 0, Qt.AlignVCenter)
        hl.addSpacing(8)
        self.about_btn = _CompactToolButton()          # Thông tin & hỗ trợ
        self.about_btn.setObjectName("setupBtn")
        self.about_btn.setIcon(glyph_icon(0xE946, "#334155"))
        self.about_btn.setIconSize(QSize(18, 18))
        self.about_btn.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        self.about_btn.setText("Hỗ trợ")
        self.about_btn.setToolTip("Thông tin phiên bản, liên hệ hỗ trợ, giấy phép bên thứ ba")
        self.about_btn.setCursor(Qt.PointingHandCursor)
        self.about_btn.clicked.connect(lambda: self.aboutRequested.emit(0))
        hl.addWidget(self.about_btn, 0, Qt.AlignVCenter)
        outer.addWidget(head)

        self.scroll = scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        body = QWidget()
        body.setObjectName("homeBody")
        self.body_lay = QVBoxLayout(body)
        self.body_lay.setContentsMargins(28, 8, 28, 28)
        self.body_lay.setSpacing(6)
        scroll.setWidget(body)
        # thanh cuộn dọc chiếm chỗ bên phải -> bớt lề phải tương ứng để mép phải các thẻ thẳng hàng với
        # thanh tiêu đề / thẻ Mở PDF và đều với lề trái
        vsb = scroll.verticalScrollBar()
        vsb.rangeChanged.connect(lambda lo, hi: self.body_lay.setContentsMargins(
            28, 8, max(4, 28 - vsb.sizeHint().width()) if hi > lo else 28, 28))

        self.open_card = OpenCard()
        self.open_card.openClicked.connect(lambda: self.toolChosen.emit("editor:open"))
        self.open_card.recentClicked.connect(lambda p: self.toolChosen.emit("recent:" + p))
        oc = QWidget()
        oc.setObjectName("homeTop")
        ol = QVBoxLayout(oc)
        ol.setContentsMargins(28, 16, 28, 0)
        ol.setSpacing(12)
        ol.addWidget(self.open_card)
        cat_scroll = QScrollArea()                     # thanh danh mục: cuộn ngang nếu cửa sổ hẹp
        cat_scroll.setObjectName("catScroll")
        cat_scroll.setWidgetResizable(True)
        cat_scroll.setFrameShape(QFrame.NoFrame)
        cat_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        cats = QWidget()
        self.cat_lay = QHBoxLayout(cats)
        self.cat_lay.setContentsMargins(0, 0, 0, 10)
        self.cat_lay.setSpacing(4)
        cat_scroll.setWidget(cats)
        ol.addWidget(cat_scroll)
        outer.addWidget(oc)
        outer.addWidget(scroll, 1)

        self._section("Chỉnh sửa & Ký số", "editor",
                      [(tid, n, d, ic) for tid, n, d, ic in EDITOR_CARDS], max_cols=5)
        pop = [t for t in T.TOOLS if t.popular and t.id != "verify_sig"]
        self._section("Công cụ phổ biến", "popular", [(t.id, t.name, t.desc, t.icon, t.group) for t in pop])
        for gid, gname in T.GROUPS:
            self._section(gname, gid, [(t.id, t.name, t.desc, t.icon) for t in T.TOOLS
                                       if t.group == gid and t.id != "verify_sig"])
        self.empty = QLabel("Không tìm thấy công cụ phù hợp.")
        self.empty.setStyleSheet("color:#94a3b8; font-size:12pt; padding:30px")
        self.empty.hide()
        self.body_lay.addWidget(self.empty)
        self.body_lay.addStretch()

        total = len({c.tid for s in self.sections for c in s[3]})
        self.cat_btns = {}
        grp = QButtonGroup(self)
        order = [("all", f"Tất cả công cụ ({total})")] + [(s[0], s[5]) for s in self.sections]
        order[1], order[2] = order[2], order[1]        # Phổ biến trước, rồi Chỉnh sửa & Ký số
        for gid, text in order:
            b = QToolButton()
            b.setObjectName("catBtn")
            b.setText(i18n.tr(text).replace("&", "&&"))     # "&" của nút = phím tắt -> hiện đúng ký tự
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda _=False, g=gid: self.set_category(g))
            grp.addButton(b)
            self.cat_lay.addWidget(b)
            self.cat_btns[gid] = b
        self.cat_lay.addStretch()
        cats.adjustSize()
        cat_scroll.setFixedHeight(cats.sizeHint().height() + 8)
        self.cat_btns["all"].setChecked(True)
        self.filter("")

    def set_lang_buttons(self, lang):
        self.lang_btns.get(lang, self.lang_btns["vi"]).setChecked(True)

    def set_open_docs(self, names, active):
        """Nút "Quay lại file đang mở": names = tên các tab, active = tab đang chọn (-1: không có)."""
        b = self.resume_btn
        b.setVisible(bool(names))
        if not names:
            self._fit_header()
            return
        cur = names[active] if 0 <= active < len(names) else names[0]
        short = cur if len(cur) <= 28 else cur[:25] + "..."
        back = i18n.tr("Quay lại")
        tip = i18n.tr("Quay lại file đang làm việc (Esc)")
        with i18n.paused():                              # tên file: không dịch
            label = back + "  " + short
            if len(names) > 1:
                label += f"  (+{len(names) - 1})"
            b.full_text = label
            b.short_text = back + (f" ({len(names)})" if len(names) > 1 else "")
            b.setText(label)
            b.setToolTip(tip + "\n" + cur)
            m = b.menu()
            m.clear()
            for i, n in enumerate(names):
                a = m.addAction(n)
                a.setCheckable(True)
                a.setChecked(i == active)
                a.triggered.connect(lambda _=False, i=i: self.resumeRequested.emit(i))
        self._fit_header()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._fit_header()

    def _fit_header(self):
        """Thanh tiêu đề không được rộng hơn cửa sổ: thiếu chỗ thì thu gọn dần (nút Quay lại rút gọn ->
        Cài đặt / Hỗ trợ chỉ còn icon -> ẩn dòng mô tả). Các nút này không ép độ rộng tối thiểu (xem
        _CompactToolButton) nên cửa sổ hẹp vẫn nhận đúng độ rộng để tính lại ở đây."""
        if getattr(self, "_fitting", False) or not hasattr(self, "head"):
            return
        self._fitting = True
        try:
            r = self.resume_btn
            levels = (
                lambda: (r.setText(getattr(r, "full_text", r.text())),
                         [b.setToolButtonStyle(Qt.ToolButtonTextBesideIcon) for b in (self.setup_btn, self.about_btn)],
                         [x.setVisible(True) for x in self.sub_labels]),
                lambda: r.setText(getattr(r, "short_text", r.text())),
                lambda: [b.setToolButtonStyle(Qt.ToolButtonIconOnly) for b in (self.setup_btn, self.about_btn)],
                lambda: [x.setVisible(False) for x in self.sub_labels],
            )
            avail = self.width()
            for apply in levels:                    # mức sau giữ thay đổi của mức trước
                apply()
                self.head_lay.invalidate()
                if self.head_lay.sizeHint().width() <= avail:
                    break
        finally:
            self._fitting = False

    def _focus_search(self):
        self.search.setFocus()
        self.search.selectAll()

    def set_category(self, gid):
        """Bấm 1 danh mục: chỉ hiện nhóm đó (tìm kiếm luôn tìm trong tất cả)."""
        self.cat = gid
        self.cat_btns[gid].setChecked(True)
        if self.search.text():
            self.search.clear()                        # -> filter("")
        else:
            self.filter("")
        self.scroll.verticalScrollBar().setValue(0)

    def _section(self, title, gid, items, max_cols=5):
        hdr = QWidget()
        h = QHBoxLayout(hdr)
        h.setContentsMargins(0, 16, 0, 4)
        h.setSpacing(10)
        bar = QFrame()
        bar.setFixedSize(4, 18)
        bar.setStyleSheet(f"background:{ACCENT}; border-radius:2px")
        lbl = QLabel(title)
        f = QFont(lbl.font())
        f.setPointSizeF(11.5)
        f.setBold(True)
        f.setCapitalization(QFont.AllUppercase)
        f.setLetterSpacing(QFont.AbsoluteSpacing, 0.6)
        lbl.setFont(f)
        lbl.setStyleSheet("color:#0f172a")
        cnt = QLabel()
        cnt.setStyleSheet("color:#94a3b8")
        h.addWidget(bar, 0, Qt.AlignVCenter)
        h.addWidget(lbl)
        h.addStretch()
        h.addWidget(cnt)
        wrap = QWidget()
        flow = GridFlow(wrap, max_cols=max_cols)
        cards = []
        for it in items:
            tid, name, desc, icon = it[:4]
            color = GROUP_COLORS.get(it[4] if len(it) > 4 else gid, ACCENT)
            c = ToolCard(tid, name, desc, icon, color)
            kw = T.get_tool(tid).keywords if not tid.startswith("editor:") else ""
            c.search += " " + strip_accents(kw)
            c.clicked.connect(self.toolChosen)
            flow.addWidget(c)
            cards.append(c)
        self.body_lay.addWidget(hdr)
        self.body_lay.addWidget(wrap)
        self.sections.append((gid, hdr, wrap, cards, cnt, title))

    def filter(self, text):
        words = strip_accents(text).split()
        if words and self.cat != "all":                # gõ tìm -> tìm trong tất cả công cụ
            self.cat = "all"
            self.cat_btns["all"].setChecked(True)
        any_shown = False
        for gid, hdr, wrap, cards, cnt, _t in self.sections:
            in_cat = self.cat in ("all", gid)
            hide_dup = bool(words) and gid == "popular"          # khi tìm kiếm, tránh hiện trùng
            shown = 0
            for c in cards:
                ok = in_cat and not hide_dup and all(w in c.search for w in words)
                c.setVisible(ok)
                shown += ok
            hdr.setVisible(shown > 0)
            wrap.setVisible(shown > 0)
            cnt.setText(f"{shown} tính năng")
            wrap.layout().invalidate()
            any_shown |= shown > 0
        self.empty.setVisible(not any_shown)


# ===========================================================================
#  Danh sách file đầu vào (kéo thả, sắp thứ tự)
# ===========================================================================
class FileList(QListWidget):
    changed = Signal()

    def __init__(self):
        super().__init__()
        self.accept_ext = ()
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.model().rowsMoved.connect(lambda *a: self.changed.emit())

    def paths(self):
        return [self.item(i).data(Qt.UserRole) for i in range(self.count())]

    def add_paths(self, paths):
        have = set(self.paths())
        bad = []
        for p in paths:
            if os.path.isdir(p):
                self.add_paths([os.path.join(p, f) for f in sorted(os.listdir(p))])
                continue
            if self.accept_ext and os.path.splitext(p)[1].lower() not in self.accept_ext:
                bad.append(os.path.basename(p))
                continue
            if p in have:
                continue
            it = QListWidgetItem(os.path.basename(p))
            it.setData(Qt.UserRole, p)
            it.setToolTip(p)
            try:
                it.setText(f"{os.path.basename(p)}   ({os.path.getsize(p) / 1024 / 1024:.1f} MB)")
            except OSError:
                pass
            self.addItem(it)
            have.add(p)
        self.changed.emit()
        return bad

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            super().dragEnterEvent(e)

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            super().dragMoveEvent(e)

    def dropEvent(self, e):
        if e.mimeData().hasUrls():
            self.add_paths([u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()])
            e.acceptProposedAction()
        else:
            super().dropEvent(e)

    def paintEvent(self, e):
        super().paintEvent(e)
        if self.count() == 0:
            p = QPainter(self.viewport())
            p.setPen(QColor("#94a3b8"))
            f = p.font()
            f.setPointSize(11)
            p.setFont(f)
            p.drawText(self.viewport().rect(), Qt.AlignCenter, "Kéo thả file vào đây\nhoặc bấm  Thêm file")
            p.end()


# ===========================================================================
#  Ô nhập cho từng loại tuỳ chọn
# ===========================================================================
class ColorPick(QPushButton):
    changed = Signal()

    def __init__(self, hexstr):
        super().__init__()
        self.setFixedWidth(90)
        self.set(hexstr or "#000000")
        self.clicked.connect(self.pick)

    def set(self, h):
        self.hex = h
        fg = "#fff" if QColor(h).lightness() < 140 else "#000"
        self.setText(h.upper())
        self.setStyleSheet(f"background:{h}; color:{fg}; border:1px solid #94a3b8; border-radius:4px; padding:3px")

    def pick(self):
        c = QColorDialog.getColor(QColor(self.hex), self, "Chọn màu")
        if c.isValid():
            self.set(c.name())
            self.changed.emit()


class FileField(QWidget):
    changed = Signal()

    def __init__(self, filt, multi=False):
        super().__init__()
        self.filt, self.multi, self.files = filt or "Tất cả (*.*)", multi, []
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        self.edit = QLineEdit()
        self.edit.setReadOnly(True)
        self.edit.setPlaceholderText("(chưa chọn)")
        b = QPushButton("Chọn...")
        b.clicked.connect(self.pick)
        c = QPushButton("Bỏ")
        c.clicked.connect(lambda: self.set([]))
        lay.addWidget(self.edit, 1)
        lay.addWidget(b)
        lay.addWidget(c)

    def pick(self):
        if self.multi:
            paths, _ = QFileDialog.getOpenFileNames(self, "Chọn file", "", self.filt)
            if paths:
                self.set(self.files + paths)
        else:
            p, _ = QFileDialog.getOpenFileName(self, "Chọn file", "", self.filt)
            if p:
                self.set([p])

    def set(self, files):
        self.files = list(files)
        self.edit.setText("; ".join(os.path.basename(f) for f in self.files))
        self.edit.setToolTip("\n".join(self.files))
        self.changed.emit()

    def value(self):
        return self.files if self.multi else (self.files[0] if self.files else "")


class FormTable(QTableWidget):
    """Bảng điền form: tên ô | giá trị."""

    def __init__(self):
        super().__init__(0, 2)
        self.setHorizontalHeaderLabels(["Ô form", "Giá trị"])
        self.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.horizontalHeader().setStretchLastSection(True)
        self.verticalHeader().setVisible(False)
        self.setMinimumHeight(180)
        self.fields = []

    def load(self, fields):
        self.fields = fields or []
        self.setRowCount(len(self.fields))
        for r, (name, kind, value, choices) in enumerate(self.fields):
            it = QTableWidgetItem(name)
            it.setFlags(Qt.ItemIsEnabled)
            self.setItem(r, 0, it)
            with i18n.paused():                     # giá trị trong form PDF: không dịch
                if kind == "check":
                    w = QCheckBox()
                    w.setChecked(value not in (None, "", "Off", False))
                elif kind in ("choice", "radio"):
                    w = QComboBox()
                    w.setEditable(kind == "choice")
                    w.addItems([""] + list(choices))
                    w.setCurrentText(str(value or ""))
                else:
                    w = QLineEdit(str(value or ""))
            self.setCellWidget(r, 1, w)

    def value(self):
        out = {}
        for r, (name, kind, value, choices) in enumerate(self.fields):
            w = self.cellWidget(r, 1)
            if isinstance(w, QCheckBox):
                out[name] = w.isChecked()
            elif isinstance(w, QComboBox):
                out[name] = w.currentText()
            else:
                out[name] = w.text()
        return out


class SafeThread(QThread):
    """QThread tự giữ tham chiếu tới khi chạy xong hẳn.
    Tín hiệu "xong" phát ra từ bên trong run() - lúc nhận được, luồng vẫn còn chạy nốt. Nếu nơi nhận bỏ tham chiếu
    (self.worker = None, đóng hộp thoại...) Python huỷ QThread đang chạy -> Qt dừng cả app (0xc0000409)."""
    _alive = set()

    def __init__(self):
        super().__init__()
        SafeThread._alive.add(self)
        self.finished.connect(self._release)

    def _release(self):
        SafeThread._alive.discard(self)


class Worker(SafeThread):
    progress = Signal(float, str)
    log = Signal(str)
    done = Signal(list, str)

    def __init__(self, tool, files, opts, out_dir):
        super().__init__()
        self.tool, self.files, self.opts, self.out_dir = tool, files, opts, out_dir
        self.cancelled = False

    def run(self):
        ctx = T.Ctx(self.files, self.opts, self.out_dir, progress=lambda f, m: self.progress.emit(f, m),
                    log=self.log.emit, cancel_flag=lambda: self.cancelled)
        err = ""
        try:
            self.tool.run(ctx)
        except T.Cancelled:
            err = "Đã dừng."
        except Exception as e:
            err = str(e) or type(e).__name__
        self.done.emit([p for p in ctx.outputs if os.path.exists(p)], err)


class OptionsForm(QScrollArea):
    """Form tuỳ chọn tự dựng từ danh sách pdf_tools.Opt (dùng chung cho màn hình công cụ và hộp thoại)."""
    changed = Signal()

    def __init__(self):
        super().__init__()
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.tool = None
        self.widgets = {}
        self.touched = set()

    def load(self, tool, saved=None, hide=()):
        self.tool = tool
        self.touched = set()
        saved = saved or {}
        host = QWidget()
        form = QFormLayout(host)
        form.setLabelAlignment(Qt.AlignRight | Qt.AlignVCenter)
        form.setFieldGrowthPolicy(QFormLayout.ExpandingFieldsGrow)
        self.form = form
        self.widgets = {}
        for o in tool.opts:
            if o.key in hide:
                continue
            val = saved.get(o.key, o.default) if o.kind not in ("password", "form", "files") else o.default
            w = self._make_widget(o, val)
            self.widgets[o.key] = (o, w)
            form.addRow("" if o.kind == "bool" else o.label + ":", w)
            if o.tip:
                w.setToolTip(o.tip)
        if not self.widgets:
            form.addRow(QLabel("Công cụ này không có tuỳ chọn."))
        self.setWidget(host)
        self._update_visibility()

    def prefill(self, paths):
        for key, (o, w) in self.widgets.items():
            if o.prefill and key not in self.touched:
                v = o.prefill(paths[0]) if paths else ([] if o.kind == "form" else "")
                if o.kind == "form":
                    w.load(v)
                elif o.kind == "textarea":
                    w.setPlainText(v or "")
                else:
                    w.setText(v or "")

    def savable(self):
        return {k: v for k, v in self.values().items()
                if self.widgets[k][0].kind not in ("password", "form", "files")}

    def _make_widget(self, o, val):
        k = o.kind
        touch = lambda *_: self.touched.add(o.key)
        if k in ("text", "pages", "password"):
            w = QLineEdit("" if val is None else str(val))
            w.setPlaceholderText(o.placeholder)
            if k == "password":
                w.setEchoMode(QLineEdit.Password)
            w.textEdited.connect(touch)
        elif k == "textarea":
            w = QPlainTextEdit(str(val or ""))
            w.setPlaceholderText(o.placeholder)
            w.setMinimumHeight(160)
            w.textChanged.connect(lambda: self.touched.add(o.key) if w.hasFocus() else None)
        elif k == "int":
            w = QSpinBox()
            w.setRange(int(o.min), int(o.max))
            w.setValue(int(val if val is not None else 0))
            w.setSuffix(o.suffix)
            w.setMaximumWidth(160)
        elif k == "float":
            w = QDoubleSpinBox()
            w.setRange(o.min, o.max)
            w.setDecimals(1 if o.step >= 0.1 else 2)
            w.setSingleStep(o.step if o.step != 1 else 1)
            w.setValue(float(val if val is not None else 0))
            w.setSuffix(o.suffix)
            w.setMaximumWidth(160)
        elif k == "bool":
            w = QCheckBox(o.label)
            w.setChecked(bool(val))
            w.toggled.connect(self._update_visibility)
        elif k == "choice":
            w = QComboBox()
            for v, lbl in o.choices:
                w.addItem(lbl, v)
            i = next((j for j, (v, _) in enumerate(o.choices) if v == val), 0)
            w.setCurrentIndex(i)
            w.currentIndexChanged.connect(self._update_visibility)
        elif k == "color":
            w = ColorPick(val)
        elif k in ("file", "files"):
            w = FileField(o.filter, multi=k == "files")
            if val and k == "file" and os.path.exists(str(val)):
                w.set([val])
        elif k == "form":
            w = FormTable()
        else:
            w = QLabel(str(val or ""))
        sig = {"text": "textChanged", "pages": "textChanged", "password": "textChanged", "textarea": "textChanged",
               "int": "valueChanged", "float": "valueChanged", "bool": "toggled", "choice": "currentIndexChanged",
               "color": "changed", "file": "changed", "files": "changed"}.get(k)
        if sig:
            getattr(w, sig).connect(lambda *_: self.changed.emit())
        return w

    def _value(self, o, w):
        k = o.kind
        if k in ("text", "pages", "password"):
            return w.text()
        if k == "textarea":
            return w.toPlainText()
        if k in ("int", "float"):
            return w.value()
        if k == "bool":
            return w.isChecked()
        if k == "choice":
            return w.currentData()
        if k == "color":
            return w.hex
        if k in ("file", "files", "form"):
            return w.value()
        return None

    def values(self):
        return {key: self._value(o, w) for key, (o, w) in self.widgets.items()}

    def _update_visibility(self, *_):
        if not self.tool:
            return
        vals = self.values()
        for key, (o, w) in self.widgets.items():
            if o.show_if:
                dep, allowed = o.show_if
                self.form.setRowVisible(w, vals.get(dep) in allowed)



# ===========================================================================
#  Xem trước: trang của file gốc và kết quả (chạy thử công cụ trên vài trang đầu)
# ===========================================================================
# Công cụ có thể chạy thử nhanh để xem trước kết quả (OCR, chuyển đổi... quá chậm hoặc không ra PDF)
PREVIEW_TOOLS = {"watermark", "header_footer", "page_numbers", "bates", "stamp", "bg_color", "greyscale", "invert",
                 "scan_effect", "rasterize", "rotate", "reverse", "delete_pages", "extract_pages", "blank_page",
                 "divide", "nup", "booklet", "poster", "crop", "fix_size", "remove_blank", "toc_page", "flatten",
                 "remove_annots", "form_fill", "compress", "text_to_pdf", "email_to_pdf", "ebook_to_pdf",
                 "rotate_custom", "deskew", "duplex_collate", "combine_single_page", "adjust_colors", "text_color",
                 "font_to_outline", "layers", "page_labels", "ocr_editable"}
PREVIEW_PAGES = 10
PREVIEW_PAGE_LIMIT = {"ocr_editable": 2}      # công cụ chậm: chỉ chạy thử vài trang đầu


class PreviewJob(SafeThread):
    done = Signal(int, object, str)        # lượt, bytes PDF kết quả (hoặc None), lỗi

    def __init__(self, seq, tool, src, opts):
        super().__init__()
        self.seq, self.tool, self.src, self.opts = seq, tool, src, dict(opts)
        self.cancelled = False

    def run(self):
        tmp = tempfile.mkdtemp(prefix="pdftoolkit_prev_")
        logs = []
        try:
            src = self.src
            try:
                d = pymupdf.open(src)
                if d.needs_pass:
                    d.authenticate(self.opts.get("password") or "")
                limit = PREVIEW_PAGE_LIMIT.get(self.tool.id, PREVIEW_PAGES)
                if d.is_pdf and not d.needs_pass and d.page_count > limit:
                    sub = pymupdf.open()
                    sub.insert_pdf(d, to_page=limit - 1)
                    src = os.path.join(tmp, os.path.basename(self.src))
                    sub.save(src)
                d.close()
            except Exception:
                pass
            out = os.path.join(tmp, "out")
            os.makedirs(out)
            ctx = T.Ctx([src], self.opts, out, log=logs.append, cancel_flag=lambda: self.cancelled)
            self.tool.run(ctx)
            pdfs = [x for x in ctx.outputs if x.lower().endswith(".pdf") and os.path.isfile(x)]
            if not pdfs:
                raise T.ToolError(logs[-1].lstrip("✘ ") if logs else "không có kết quả PDF để xem trước")
            with open(pdfs[0], "rb") as fh:
                self.done.emit(self.seq, fh.read(), "")
        except T.Cancelled:
            pass
        except Exception as e:
            msg = next((m.lstrip("✘ ") for m in logs if m.startswith("✘")), "") or str(e) or type(e).__name__
            self.done.emit(self.seq, None, msg)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


class PreviewPane(QFrame):
    """Khung xem trước bên phải: chuyển giữa 'Kết quả' và 'File gốc', lật trang."""

    def __init__(self):
        super().__init__()
        self.setObjectName("previewPane")
        self.setMinimumWidth(260)
        self.tool = None
        self.src_path = None
        self.src_doc = None
        self.res_doc = None
        self.opts = {}
        self.job = None
        self.seq = 0
        self.pending = False
        self.mode = "src"
        self.page = {"src": 0, "res": 0}
        v = QVBoxLayout(self)
        v.setContentsMargins(10, 8, 10, 8)
        v.setSpacing(6)
        head = QHBoxLayout()
        t = QLabel("Xem trước")
        t.setStyleSheet("font-weight:600; background:transparent")
        head.addWidget(t)
        head.addStretch()
        self.mode_grp = QButtonGroup(self)
        self.b_res = QToolButton()
        self.b_res.setText("Kết quả")
        self.b_src = QToolButton()
        self.b_src.setText("File gốc")
        for b, m in ((self.b_res, "res"), (self.b_src, "src")):
            b.setCheckable(True)
            b.setObjectName("segBtn")
            b.clicked.connect(lambda _=False, m=m: self.set_mode(m))
            self.mode_grp.addButton(b)
            head.addWidget(b)
        v.addLayout(head)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setObjectName("previewScroll")
        self.view = QLabel()
        self.view.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.view.setWordWrap(True)
        self.view.setObjectName("previewView")
        self.scroll.setWidget(self.view)
        v.addWidget(self.scroll, 1)

        nav = QHBoxLayout()
        self.b_prev = QToolButton()
        self.b_prev.setIcon(glyph_icon(0xE76B))
        self.b_prev.setAutoRaise(True)
        self.b_prev.clicked.connect(lambda: self.turn(-1))
        self.b_next = QToolButton()
        self.b_next.setIcon(glyph_icon(0xE76C))
        self.b_next.setAutoRaise(True)
        self.b_next.clicked.connect(lambda: self.turn(1))
        self.page_lbl = QLabel()
        self.page_lbl.setStyleSheet("background:transparent")
        self.info = QLabel()
        self.info.setWordWrap(True)
        self.info.setStyleSheet("color:#64748b; background:transparent")
        nav.addWidget(self.b_prev)
        nav.addWidget(self.page_lbl)
        nav.addWidget(self.b_next)
        nav.addSpacing(8)
        nav.addWidget(self.info, 1)
        v.addLayout(nav)

        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(450)
        self.timer.timeout.connect(self._start_job)
        self.rtimer = QTimer(self)
        self.rtimer.setSingleShot(True)
        self.rtimer.setInterval(80)
        self.rtimer.timeout.connect(self._render)

    # ----- dữ liệu vào -----
    def has_result(self):
        return bool(self.tool and self.tool.id in PREVIEW_TOOLS)

    def set_tool(self, tool):
        self.tool = tool
        self.res_doc = None
        self.src_path = None
        self.src_doc = None
        self.page = {"src": 0, "res": 0}
        self.b_res.setVisible(self.has_result())
        self.set_mode("res" if self.has_result() else "src", render=False)
        self._render()

    def set_source(self, path, opts):
        self.opts = dict(opts)
        if path != self.src_path:
            self.src_path = path
            self.src_doc = None
            self.res_doc = None
            self.page = {"src": 0, "res": 0}
            if path:
                try:
                    d = pymupdf.open(path)
                    if d.needs_pass:
                        d.authenticate(self.opts.get("password") or "")
                    self.src_doc = None if d.needs_pass else d
                except Exception:
                    self.src_doc = None
        self.schedule()
        self._render()

    def options_changed(self, opts):
        self.opts = dict(opts)
        self.schedule()

    def schedule(self):
        if self.has_result() and self.src_path:
            self.timer.start()

    # ----- chạy thử công cụ -----
    def _start_job(self):
        if self.job and self.job.isRunning():
            self.job.cancelled = True
            self.pending = True
            return
        if not (self.has_result() and self.src_path):
            return
        self.seq += 1
        self.info.setText("Đang tạo bản xem trước...")
        self.job = PreviewJob(self.seq, self.tool, self.src_path, self.opts)
        self.job.done.connect(self._job_done)
        self.job.finished.connect(self._job_finished)
        self.job.start()

    def _job_finished(self):
        if self.pending:
            self.pending = False
            self._start_job()

    def _job_done(self, seq, data, err):
        if seq != self.seq or self.pending:
            return
        if err:
            self.res_doc = None
            self.res_err = err
        else:
            self.res_doc = pymupdf.open("pdf", data)
            self.res_size = len(data)
            self.res_err = ""
        if self.mode == "res":
            self._render()

    def stop(self):
        self.timer.stop()
        if self.job and self.job.isRunning():
            self.job.cancelled = True
            self.job.wait(3000)

    # ----- hiển thị -----
    def set_mode(self, m, render=True):
        self.mode = m
        (self.b_res if m == "res" else self.b_src).setChecked(True)
        if render:
            self._render()

    def turn(self, d):
        doc = self.res_doc if self.mode == "res" else self.src_doc
        if doc:
            self.page[self.mode] = max(0, min(doc.page_count - 1, self.page[self.mode] + d))
            self._render()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self.rtimer.start()

    def _placeholder(self, text):
        self.view.setPixmap(QPixmap())
        self.view.setText(text)
        self.view.setAlignment(Qt.AlignCenter)
        self.page_lbl.setText("")
        self.b_prev.setEnabled(False)
        self.b_next.setEnabled(False)

    def _render(self):
        res = self.mode == "res"
        doc = self.res_doc if res else self.src_doc
        if not self.src_path:
            self._placeholder("Thêm file để xem trước")
            self.info.setText("")
            return
        if doc is None:
            if res and (self.job and self.job.isRunning() or self.timer.isActive()):
                self._placeholder("Đang tạo bản xem trước...")
            elif res and getattr(self, "res_err", ""):
                self._placeholder("Không xem trước được:\n" + self.res_err)
                self.info.setText("")
            else:
                self._placeholder("Không xem trước được loại file này")
            return
        i = self.page[self.mode] = max(0, min(doc.page_count - 1, self.page[self.mode]))
        pg = doc[i]
        avail = max(200, self.scroll.viewport().width() - 16)
        dpr = self.devicePixelRatioF()
        z = min(avail / max(1, pg.rect.width), 3.0) * dpr
        try:
            pix = pg.get_pixmap(matrix=pymupdf.Matrix(z, z), alpha=False)
            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
            pm = QPixmap.fromImage(img)
            pm.setDevicePixelRatio(dpr)
            self.view.setText("")
            self.view.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
            self.view.setPixmap(pm)
        except Exception as e:
            self._placeholder(f"Lỗi hiển thị trang: {e}")
            return
        self.page_lbl.setText(f"{i + 1} / {doc.page_count}")
        self.b_prev.setEnabled(i > 0)
        self.b_next.setEnabled(i < doc.page_count - 1)
        if res:
            more = self.src_doc is not None and self.src_doc.page_count > PREVIEW_PAGES
            self.info.setText(f"Kết quả {doc.page_count} trang · {self.res_size / 1024:,.0f} KB"
                              + (f"  (chạy thử trên {PREVIEW_PAGES} trang đầu)" if more else ""))
        else:
            self.info.setText(f"File gốc {doc.page_count} trang")


# ===========================================================================
#  Màn hình một công cụ
# ===========================================================================
class ToolPage(QWidget):
    back = Signal()
    openInEditor = Signal(str)

    def __init__(self):
        super().__init__()
        self.settings = QSettings(installer.ORG, "PDFToolkit")
        self.tool = None
        self.worker = None
        self.outputs = []
        outer = QVBoxLayout(self)
        outer.setContentsMargins(20, 12, 20, 12)

        head = QHBoxLayout()
        b = QToolButton()
        b.setText("  Tất cả công cụ")
        b.setIcon(glyph_icon(0xE72B))
        b.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        b.setAutoRaise(True)
        b.clicked.connect(self._back)
        head.addWidget(b)
        head.addStretch()
        outer.addLayout(head)

        tl = QHBoxLayout()
        self.icon = QLabel()
        self.title = QLabel()
        self.title.setStyleSheet("font-size:17pt; font-weight:700; color:#0f172a")
        self.desc = QLabel()
        self.desc.setWordWrap(True)
        self.desc.setStyleSheet("color:#475569")
        tcol = QVBoxLayout()
        tcol.addWidget(self.title)
        tcol.addWidget(self.desc)
        tl.addWidget(self.icon, 0, Qt.AlignTop)
        tl.addSpacing(8)
        tl.addLayout(tcol, 1)
        outer.addLayout(tl)

        split = QSplitter()
        # --- trái: file ---
        left = QWidget()
        ll = QVBoxLayout(left)
        ll.setContentsMargins(0, 8, 8, 0)
        self.files_label = QLabel()
        self.files_label.setStyleSheet("font-weight:600")
        ll.addWidget(self.files_label)
        self.files = FileList()
        self.files.changed.connect(self._files_changed)
        ll.addWidget(self.files, 1)
        row = QHBoxLayout()
        for text, fn in (("Thêm file", self.add_files), ("Bỏ chọn", self.remove_sel), ("Xoá hết", self.clear_files)):
            bb = QPushButton(text)
            bb.clicked.connect(fn)
            row.addWidget(bb)
        for code, d in ((0xE74A, -1), (0xE74B, 1)):
            bb = QToolButton()
            bb.setIcon(glyph_icon(code))
            bb.setToolTip("Đưa file đang chọn lên" if d < 0 else "Đưa file đang chọn xuống")
            bb.clicked.connect(lambda _=False, d=d: self.move_sel(d))
            row.addWidget(bb)
        ll.addLayout(row)
        self.note = QLabel()
        self.note.setWordWrap(True)
        self.note.setStyleSheet("color:#b45309")
        ll.addWidget(self.note)
        split.addWidget(left)

        # --- phải: tuỳ chọn + chạy ---
        right = QWidget()
        rl = QVBoxLayout(right)
        rl.setContentsMargins(8, 8, 0, 0)
        lbl = QLabel("Tuỳ chọn")
        lbl.setStyleSheet("font-weight:600")
        rl.addWidget(lbl)
        self.opts = OptionsForm()
        rl.addWidget(self.opts, 1)

        outbox = QHBoxLayout()
        self.r_same = QRadioButton("Lưu cạnh file gốc")
        self.r_dir = QRadioButton("Lưu vào thư mục:")
        self.out_dir = QLineEdit(self.settings.value("tools/out_dir", ""))
        self.out_dir.setPlaceholderText("chọn thư mục...")
        bdir = QPushButton("Chọn...")
        bdir.clicked.connect(self.pick_dir)
        (self.r_dir if self.settings.value("tools/out_mode", "same") == "dir" else self.r_same).setChecked(True)
        # 2 dòng (không phải 1) để cột tuỳ chọn không ép cửa sổ rộng hơn màn hình laptop
        outbox.addWidget(self.r_same)
        outbox.addStretch()
        rl.addLayout(outbox)
        dirbox = QHBoxLayout()
        dirbox.addWidget(self.r_dir)
        dirbox.addWidget(self.out_dir, 1)
        dirbox.addWidget(bdir)
        rl.addLayout(dirbox)

        runrow = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress.setRange(0, 1000)
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(8)
        self.status = QLabel("")
        self.status.setStyleSheet("color:#475569")
        self.run_btn = QPushButton("  Chạy")
        self.run_btn.setIcon(glyph_icon(0xE768, "#ffffff"))
        self.run_btn.setObjectName("primary")
        self.run_btn.setMinimumHeight(38)
        self.run_btn.setMinimumWidth(150)
        self.run_btn.clicked.connect(self.run_or_stop)
        pcol = QVBoxLayout()
        pcol.addWidget(self.status)
        pcol.addWidget(self.progress)
        runrow.addLayout(pcol, 1)
        runrow.addWidget(self.run_btn)
        rl.addLayout(runrow)

        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumBlockCount(2000)
        self.log.setPlaceholderText("Kết quả sẽ hiện ở đây")
        self.log.setMaximumHeight(150)
        rl.addWidget(self.log)
        res = QHBoxLayout()
        self.b_folder = QPushButton("Mở thư mục kết quả")
        self.b_folder.clicked.connect(lambda: self.outputs and reveal(self.outputs[0]))
        self.b_open = QPushButton("Mở file kết quả")
        self.b_open.clicked.connect(lambda: self.outputs and open_path(self.outputs[0]))
        self.b_edit = QPushButton("Mở trong trình sửa")
        self.b_edit.clicked.connect(self._open_editor)
        for x in (self.b_folder, self.b_open, self.b_edit):
            x.setEnabled(False)
            res.addWidget(x)
        res.addStretch()
        rl.addLayout(res)
        split.addWidget(right)
        self.preview = PreviewPane()
        split.addWidget(self.preview)
        split.setStretchFactor(1, 1)
        split.setStretchFactor(2, 1)
        split.setSizes([300, 640, 560])
        outer.addWidget(split, 1)
        self.files.currentItemChanged.connect(lambda *_: self._update_preview())
        self.opts.changed.connect(lambda: self.preview.options_changed(self.opts.values()))

    # ------------------------------------------------------------------ dựng form
    def load_tool(self, tid, files=None):
        t = T.get_tool(tid)
        self.tool = t
        self.icon.setPixmap(badge_pixmap(t.icon, GROUP_COLORS.get(t.group, ACCENT), 48))
        self.title.setText(t.name)
        self.desc.setText(t.desc)
        self.note.setText(t.note)
        self.note.setVisible(bool(t.note))
        self.files.accept_ext = tuple(t.accept)
        lim = (f"đúng {t.min_files}" if t.min_files == t.max_files else
               f"tối thiểu {t.min_files}" if t.min_files > 1 else "một hoặc nhiều")
        exts = ", ".join(e.lstrip(".").upper() for e in t.accept[:10]) + ("..." if len(t.accept) > 10 else "")
        self.files_label.setText(f"File đầu vào ({lim} file: {exts})")
        self.files.clear()
        self.log.clear()
        self.status.clear()
        self.progress.setValue(0)
        self.outputs = []
        self._result_buttons()
        saved = {}
        try:
            saved = json.loads(self.settings.value(f"tools/opts/{tid}", "{}"))
        except ValueError:
            pass

        self.opts.load(t, saved)
        self.preview.set_tool(t)
        if files:
            self.files.add_paths(files)

    # ------------------------------------------------------------------ file
    def _files_changed(self):
        n = self.files.count()
        self.opts.prefill(self.files.paths())
        self.status.setText(f"{n} file" if n else "")
        self._update_preview()

    def _update_preview(self):
        it = self.files.currentItem() or (self.files.item(0) if self.files.count() else None)
        self.preview.set_source(it.data(Qt.UserRole) if it else None, self.opts.values())

    def add_files(self):
        t = self.tool
        filt = "Các file hỗ trợ (" + " ".join("*" + e for e in t.accept) + ");;Tất cả (*.*)"
        paths, _ = QFileDialog.getOpenFileNames(self, "Chọn file", self.settings.value("tools/last_dir", ""), filt)
        if paths:
            self.settings.setValue("tools/last_dir", os.path.dirname(paths[0]))
            bad = self.files.add_paths(paths)
            if bad:
                QMessageBox.information(self, t.name, "Bỏ qua file không đúng loại:\n" + "\n".join(bad[:15]))

    def remove_sel(self):
        for it in self.files.selectedItems():
            self.files.takeItem(self.files.row(it))
        self.files.changed.emit()

    def clear_files(self):
        self.files.clear()
        self.files.changed.emit()

    def move_sel(self, d):
        rows = sorted((self.files.row(i) for i in self.files.selectedItems()), reverse=d > 0)
        for r in rows:
            nr = r + d
            if 0 <= nr < self.files.count():
                it = self.files.takeItem(r)
                self.files.insertItem(nr, it)
                it.setSelected(True)
        self.files.changed.emit()

    def pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Thư mục lưu kết quả", self.out_dir.text())
        if d:
            self.out_dir.setText(d)
            self.r_dir.setChecked(True)

    # ------------------------------------------------------------------ chạy
    def run_or_stop(self):
        if self.worker and self.worker.isRunning():
            self.worker.cancelled = True
            self.status.setText("Đang dừng...")
            return
        t = self.tool
        files = self.files.paths()
        if len(files) < t.min_files:
            QMessageBox.information(self, t.name, f"Cần ít nhất {t.min_files} file." if t.min_files > 1
                                    else "Hãy thêm file cần xử lý.")
            return
        if t.max_files and len(files) > t.max_files:
            QMessageBox.information(self, t.name, f"Công cụ này dùng đúng {t.max_files} file.")
            return
        out_dir = None
        if self.r_dir.isChecked():
            out_dir = self.out_dir.text().strip()
            if not out_dir or not os.path.isdir(out_dir):
                QMessageBox.information(self, t.name, "Hãy chọn thư mục lưu kết quả.")
                return
        vals = self.opts.values()
        self.settings.setValue(f"tools/opts/{t.id}", json.dumps(self.opts.savable(), ensure_ascii=False))
        self.settings.setValue("tools/out_mode", "dir" if out_dir else "same")
        self.settings.setValue("tools/out_dir", self.out_dir.text())
        self.log.clear()
        self.outputs = []
        self._result_buttons()
        self.worker = Worker(t, files, vals, out_dir)
        self.worker.progress.connect(self._on_progress)
        self.worker.log.connect(self.log.appendPlainText)
        self.worker.done.connect(self._on_done)
        self.run_btn.setText("  Dừng")
        self.run_btn.setIcon(glyph_icon(0xE71A, "#ffffff"))
        self.status.setText("Đang xử lý...")
        self.worker.start()

    def _on_progress(self, frac, msg):
        self.progress.setValue(int(frac * 1000))
        if msg:
            self.status.setText(msg)

    def _on_done(self, outputs, err):
        self.run_btn.setText("  Chạy")
        self.run_btn.setIcon(glyph_icon(0xE768, "#ffffff"))
        self.outputs = outputs
        self._result_buttons()
        if err:
            self.status.setText("✘ " + err)
            self.status.setStyleSheet("color:#b91c1c")
            self.log.appendPlainText("✘ " + err)
        else:
            self.progress.setValue(1000)
            self.status.setStyleSheet("color:#15803d")
            self.status.setText(f"✔ Xong - {len(outputs)} kết quả" if outputs else "✔ Xong")
            for p in outputs:
                self.log.appendPlainText("→ " + p)
        self.worker = None

    def _result_buttons(self):
        has = bool(self.outputs)
        self.b_folder.setEnabled(has)
        self.b_open.setEnabled(has and os.path.isfile(self.outputs[0]))
        self.b_edit.setEnabled(has and os.path.isfile(self.outputs[0]) and self.outputs[0].lower().endswith(".pdf"))
        if not has:
            self.status.setStyleSheet("color:#475569")

    def _open_editor(self):
        if self.outputs:
            self.openInEditor.emit(self.outputs[0])

    def busy(self):
        return bool(self.worker and self.worker.isRunning())

    def _back(self):
        if self.busy():
            QMessageBox.information(self, self.tool.name, "Công cụ đang chạy - bấm Dừng trước.")
            return
        self.preview.stop()
        self.back.emit()


TOOLS_STYLE = """
#homeHead { background:#ffffff; border-bottom:1px solid #e2e8f0; }
#homeBody, #homeTop { background:#f8fafc; }
#homeTitle { font-size:20pt; font-weight:700; color:#0f172a; }
#versionLabel { font-size:11pt; font-weight:700; color:#2563eb; }
#authorTag { background:#eff6ff; color:#1d4ed8; border:1px solid #bfdbfe; border-radius:9px; padding:1px 8px;
             font-size:8.5pt; font-weight:600; }
#kbd { background:#f1f5f9; color:#64748b; border:1px solid #e2e8f0; border-radius:4px; padding:0 5px;
       font-family:Consolas; font-size:8pt; }
#catScroll, #catScroll > QWidget > QWidget { background:transparent; }
#catScroll { border-bottom:1px solid #e2e8f0; }
#catScroll QScrollBar:horizontal { height:6px; background:transparent; }
#catScroll QScrollBar::handle:horizontal { background:#cbd5e1; border-radius:3px; min-width:40px; }
#catScroll QScrollBar::add-line, #catScroll QScrollBar::sub-line { width:0; }
QToolButton#catBtn { padding:6px 14px; border:none; border-radius:15px; color:#334155; background:transparent; }
QToolButton#catBtn:hover:!checked { background:#e2e8f0; }
QToolButton#catBtn:checked { background:#2563eb; color:white; font-weight:600; }
#card { background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; }
#card:hover { border:1px solid #93c5fd; background:#f8fbff; }
QPushButton#primary { background:#2563eb; color:white; border:none; border-radius:6px; font-weight:600; padding:6px 18px; }
QPushButton#primary:hover { background:#1d4ed8; }
QProgressBar { background:#e2e8f0; border:none; border-radius:4px; }
QProgressBar::chunk { background:#2563eb; border-radius:4px; }
"""


# ===========================================================================
#  Ô "Mở PDF" lớn trên trang chủ + file mở gần đây
# ===========================================================================
class OpenCard(QFrame):
    openClicked = Signal()
    recentClicked = Signal(str)

    def __init__(self):
        super().__init__()
        self.setObjectName("openCard")
        self.lay = lay = QHBoxLayout(self)
        lay.setContentsMargins(18, 16, 22, 16)
        lay.setSpacing(18)
        btn = QPushButton("   Mở PDF")
        btn.setObjectName("openBig")
        btn.setIcon(glyph_icon(0xE710, "#ffffff"))
        btn.setIconSize(QSize(22, 22))
        f = QFont(btn.font())
        f.setPointSizeF(14)
        f.setBold(True)                 # sizeHint không tính font 14pt của QSS -> tự đo để không cụt chữ ("Open PD")
        btn.setMinimumSize(max(150, QFontMetrics(f).horizontalAdvance(btn.text()) + 22 + 44 + 16), 58)
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(self.openClicked)
        lay.addWidget(btn, 0, Qt.AlignVCenter)
        self.col = col = QVBoxLayout()
        col.setSpacing(4)
        t = QLabel("Mở file để xem, sửa chữ, chèn, sắp xếp trang, OCR, ký số...")
        t.setStyleSheet("font-size:11.5pt; font-weight:600; color:#0f172a; background:transparent")
        hr = QHBoxLayout()
        hr.setSpacing(6)
        h = QLabel("Hoặc kéo thả file PDF / ảnh vào cửa sổ này.")
        h.setStyleSheet("color:#64748b; background:transparent")
        kbd = QLabel("Ctrl + O")
        kbd.setObjectName("kbd")
        hr.addWidget(h)
        hr.addWidget(kbd)
        hr.addStretch()
        col.addWidget(t)
        col.addLayout(hr)
        lay.addLayout(col, 1)
        self.btn, self.title, self.hint = btn, t, (h, kbd)
        self._chips, self._prefix = [], []
        # file gần đây: bên phải, xuống dưới khi hẹp, hẹp nữa thì ẩn bớt chip. Không góp vào độ rộng tối thiểu
        # của thẻ - nếu không, chip tên dài ép cả trang chủ rộng hơn màn hình (tràn nút Hỗ trợ, thẻ công cụ)
        self.recent_w = _ShrinkRow()
        self.recent_w.setStyleSheet("background:transparent")
        self.recent_box = QHBoxLayout(self.recent_w)
        self.recent_box.setContentsMargins(0, 0, 0, 0)
        self.recent_box.setSpacing(8)
        lay.addWidget(self.recent_w, 0, Qt.AlignVCenter)
        self._below = False

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._relayout()

    @staticmethod
    def _size_text(p):
        n = os.path.getsize(p)
        return f"{n / 1048576:.1f} MB" if n >= 1048576 else f"{max(1, round(n / 1024))} KB"

    def set_recent(self, paths):
        while self.recent_box.count():
            it = self.recent_box.takeAt(0)
            if it.widget():
                it.widget().deleteLater()
        self._chips, self._prefix = [], []
        paths = [p for p in paths if os.path.exists(p)][:3]           # 3 file gần nhất
        if paths:
            lbl = QLabel("Gần đây:")
            lbl.setStyleSheet("color:#64748b; background:transparent")
            ic = QLabel()
            ic.setPixmap(glyph_icon(0xE81C, "#94a3b8").pixmap(16, 16))
            self.recent_box.addWidget(ic)
            self.recent_box.addWidget(lbl)
            self._prefix = [ic, lbl]
        for p in paths:
            name = os.path.basename(p)
            b = QPushButton()
            b.setObjectName("recent")
            b.setIcon(glyph_icon(0xEA90, "#dc2626"))
            b.setToolTip(p)
            b.setCursor(Qt.PointingHandCursor)
            with i18n.paused():                                         # tên file: không dịch
                b.setText(f"{name if len(name) <= 26 else name[:23] + '...'}   {self._size_text(p)}")
            b.clicked.connect(lambda _=False, p=p: self.recentClicked.emit(p))
            self.recent_box.addWidget(b)
            self._chips.append(b)
        self._relayout(force=True)

    def _recent_width(self, n):
        """Độ rộng hàng file gần đây khi hiện n chip đầu."""
        items = self._prefix + self._chips[:n]
        if not items or n == 0:
            return 0
        sp = self.recent_box.spacing()
        return sum(w.sizeHint().width() for w in items) + sp * (len(items) - 1)

    def _relayout(self, force=False):
        """File gần đây bên phải; không đủ chỗ thì xuống dưới dòng hướng dẫn; vẫn thiếu thì ẩn bớt chip cuối."""
        m = self.lay.contentsMargins()
        inner = self.width() - m.left() - m.right()
        text_w = max(self.title.sizeHint().width(),
                     sum(x.sizeHint().width() for x in self.hint) + 6)
        left = self.btn.minimumWidth() + self.lay.spacing()
        side = inner - left - text_w - self.lay.spacing()              # chỗ trống bên phải dòng hướng dẫn
        n_all = len(self._chips)
        below = n_all > 0 and side < self._recent_width(n_all)
        avail = (inner - left) if below else side
        n = n_all
        while n > 1 and self._recent_width(n) > avail:
            n -= 1
        for i, c in enumerate(self._chips):
            c.setVisible(i < n)
        if below == self._below and not force:
            return
        self.lay.removeWidget(self.recent_w)
        self.col.removeWidget(self.recent_w)
        if below:
            self.col.addWidget(self.recent_w, 0, Qt.AlignLeft)
        else:
            self.lay.addWidget(self.recent_w, 0, Qt.AlignVCenter)
        self._below = below


class _CompactToolButton(QToolButton):
    """Nút trên thanh tiêu đề trang chủ: độ rộng tối thiểu = chỉ icon, để không ép cửa sổ rộng hơn màn hình.
    HomePage._fit_header đổi sang chỉ icon / chữ ngắn trước khi thiếu chỗ thật."""

    def __init__(self, *a):
        super().__init__(*a)
        # QToolButton mặc định Fixed (min = sizeHint) -> cho co lại tới minimumSizeHint
        self.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)

    def minimumSizeHint(self):
        return QSize(44, super().minimumSizeHint().height())


class _ShrinkRow(QWidget):
    """Hàng widget không ép độ rộng tối thiểu cho layout cha (thiếu chỗ thì cha ẩn bớt phần tử)."""

    def minimumSizeHint(self):
        return QSize(0, super().minimumSizeHint().height())


# ===========================================================================
#  Ribbon (thanh công cụ chia tab như Word)
# ===========================================================================
def _tinted(icon, color, size):
    pm = icon.pixmap(size, size)
    if pm.isNull():
        return pm
    p = QPainter(pm)
    p.setCompositionMode(QPainter.CompositionMode_SourceIn)
    p.fillRect(pm.rect(), QColor(color))
    p.end()
    return pm


def tint_icon(icon, color, size=64, badge=False, on_color=None):
    """Tô lại icon bằng màu của nhóm. badge=True: đặt icon trên ô vuông bo góc nền nhạt (giống thẻ ở trang chủ).
    on_color: màu icon khi nút đang bật (nút chế độ đang chọn nền đậm -> icon trắng)."""
    out = QIcon()
    for mode, col in ((QIcon.Normal, color), (QIcon.Disabled, "#cbd5e1")):
        if not badge:
            pm = _tinted(icon, col, size)
            if pm.isNull():
                return icon
            out.addPixmap(pm, mode)
            if on_color and mode == QIcon.Normal:
                out.addPixmap(_tinted(icon, on_color, size), QIcon.Normal, QIcon.On)
            continue
        S = 96
        pm = QPixmap(S, S)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        bg = QColor(col)
        bg.setAlpha(30 if mode == QIcon.Normal else 18)
        p.setPen(Qt.NoPen)
        p.setBrush(bg)
        p.drawRoundedRect(QRect(0, 0, S, S), 24, 24)
        g = _tinted(icon, col, 64)
        if not g.isNull():
            p.drawPixmap(QRect(16, 16, 64, 64), g)
        p.end()
        out.addPixmap(pm, mode)
    return out


RIBBON_ICON = "#3f3f46"          # một màu icon dịu cho cả thanh công cụ (không tô mỗi nhóm một màu)
RIBBON_ACCENT = "#6d28d9"        # nút đang bật: nền tím, icon trắng


class RibbonGroup(QFrame):
    """Một nhóm trên thanh kính nổi: các nút icon + nhãn bên dưới. Hiện tối đa `limit` nút (theo thứ tự thêm),
    phần còn lại gom vào nút '<tên nhóm> ▾' (menu xổ xuống) để thanh gọn. limit=0: cả nhóm là 1 nút xổ xuống."""

    def __init__(self, title, color, limit=3):
        super().__init__()
        self.setObjectName("ribbonGroup")
        self.title = title
        self.color = RIBBON_ICON
        self.limit = limit
        self.shown = 0
        self.more = None
        self.row = QHBoxLayout(self)
        self.row.setContentsMargins(2, 0, 2, 0)
        self.row.setSpacing(2)

    def showEvent(self, e):
        # cỡ chữ của nút lấy từ stylesheet (chỉ có sau khi polish) -> lúc hiện mới đo được bề rộng thật của nhãn;
        # khoá bề rộng tối thiểu để layout không bóp nút thành "Com...PDF"
        for b in self.findChildren(QToolButton):
            b.setMinimumWidth(max(54, b.sizeHint().width()))
        super().showEvent(e)

    def _bind(self, b, action, text, size, badge):
        """Nút dùng chung QAction nhưng có chữ / icon riêng (Qt đặt lại chúng mỗi khi action đổi trạng thái)."""
        b.setDefaultAction(action)
        cache = {}

        def apply(b=b, a=action):
            if text:
                b.setText(text)
            key = a.icon().cacheKey()
            if key not in cache:
                cache[key] = tint_icon(a.icon(), self.color, size, badge, "#ffffff" if a.isCheckable() else None)
            b.setIcon(cache[key])
        apply()
        action.changed.connect(apply)

    @staticmethod
    def _wrap(text, limit=9):
        """Nhãn dài -> 2 dòng, ngắt ở dấu cách gần giữa nhất (nút ribbon không tự xuống dòng)."""
        if len(text) <= limit or " " not in text:
            return text
        mid = len(text) / 2
        i = min((k for k, c in enumerate(text) if c == " "), key=lambda k: abs(k - mid))
        return text[:i] + "\n" + text[i + 1:]

    @staticmethod
    def _button():
        b = QToolButton()
        b.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
        b.setIconSize(QSize(22, 22))
        b.setAutoRaise(True)
        b.setMinimumWidth(54)
        # rộng theo chữ, không bị layout bóp lại -> không còn "Com...PDF"; thiếu chỗ thì thanh cuộn ngang
        b.setSizePolicy(QSizePolicy.Fixed, b.sizePolicy().verticalPolicy())
        b.setObjectName("ribbonBig")
        return b

    def _add(self, b):
        i = self.row.indexOf(self.more) if self.more is not None else -1
        self.row.insertWidget(i if i >= 0 else self.row.count(), b)    # nút '▾' luôn ở cuối nhóm

    def big(self, action, text=None):
        if self.shown >= self.limit:
            self._to_menu(action)
            return None
        b = self._button()
        self._bind(b, action, self._wrap(i18n.tr(text or action.text())), 44, False)
        self._add(b)
        self.shown += 1
        return b

    def small(self, actions):
        for a in actions:
            if self.shown < self.limit:
                b = self._button()
                self._bind(b, a, self._wrap(i18n.tr(a.text())), 44, False)
                self._add(b)
                self.shown += 1
            else:
                self._to_menu(a)

    def set_icon(self, code):
        """Icon riêng cho nút xổ xuống của nhóm (glyph Segoe)."""
        if self.more is not None:
            self.more.setIcon(tint_icon(glyph_icon(code), self.color, 44))

    def _to_menu(self, a):
        m = self._menu()
        if self.shown == 0 and not m.actions():          # cả nhóm là 1 nút xổ xuống: icon = icon việc đầu tiên
            self.more.setIcon(tint_icon(a.icon(), self.color, 44))
        m.addAction(a)

    def _menu(self):
        if self.more is None:
            b = self._button()
            b.setPopupMode(QToolButton.InstantPopup)
            b.setText(self._wrap(i18n.tr(self.title)).replace("&", "&&"))
            b.setIcon(tint_icon(glyph_icon(0xE70D), self.color, 44))
            b.setToolTip(i18n.tr("Thêm công cụ: ") + i18n.tr(self.title))
            b.setMenu(QMenu(b))
            self.row.addWidget(b)
            self.more = b
        return self.more.menu()


class RibbonTab(QWidget):
    """1 tab: thanh kính nổi bo tròn ở giữa, các nhóm ngăn bằng vạch mảnh."""

    def __init__(self, color):
        super().__init__()
        self.setObjectName("ribbonTab")
        self.color = color
        out = QHBoxLayout(self)
        out.setContentsMargins(14, 4, 14, 12)          # chừa chỗ cho bóng đổ của thanh
        out.addStretch()
        self.bar = QFrame()
        self.bar.setObjectName("ribbonBar")
        sh = QGraphicsDropShadowEffect(self.bar)
        sh.setBlurRadius(26)
        sh.setOffset(0, 6)
        sh.setColor(QColor(90, 60, 160, 46))
        self.bar.setGraphicsEffect(sh)
        self.lay = QHBoxLayout(self.bar)
        self.lay.setContentsMargins(8, 5, 8, 5)
        self.lay.setSpacing(4)
        out.addWidget(self.bar)
        out.addStretch()

    def group(self, title, color=None, limit=3):
        if self.lay.count():
            sep = QFrame()
            sep.setObjectName("ribbonSep")
            sep.setFixedWidth(1)
            self.lay.addWidget(sep)
        g = RibbonGroup(title, color or self.color, limit)
        self.lay.addWidget(g)
        return g


class Ribbon(QTabWidget):
    FULL, COLLAPSED = 136, 46

    def __init__(self):
        super().__init__()
        self.setObjectName("ribbon")
        self.setDocumentMode(False)               # False: dải tab gọn, căn giữa (::tab-bar alignment)
        self.tabBar().setDrawBase(False)          # bỏ đường kẻ ngang thừa phía trên
        self.tabBar().setExpanding(False)         # dải tab gọn theo nội dung (viên thuốc ở giữa)
        self.tabBar().setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(self.FULL)
        self.collapsed = False
        self.b_collapse = QToolButton()
        self.b_collapse.setAutoRaise(True)
        self.b_collapse.setObjectName("ribbonCollapse")
        self.b_collapse.clicked.connect(self.toggle)
        self._collapse_icon()
        # góc phải rộng bằng góc trái (thanh truy cập nhanh) -> dải tab nằm đúng giữa màn hình
        self.right_box = QWidget()
        rl = QHBoxLayout(self.right_box)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(0)
        self._pad = QWidget()                     # ô đệm: QTabWidget chừa chỗ theo sizeHint của góc
        rl.addWidget(self._pad)
        rl.addWidget(self.b_collapse)
        self.setCornerWidget(self.right_box, Qt.TopRightCorner)
        self.tabBar().tabBarDoubleClicked.connect(lambda _: self.toggle())
        self.tabBar().tabBarClicked.connect(lambda _: self.collapsed and self.toggle())

    def _collapse_icon(self):
        self.b_collapse.setIcon(glyph_icon(0xE70D if self.collapsed else 0xE70E, "#475569"))
        self.b_collapse.setToolTip("Hiện thanh công cụ" if self.collapsed else
                                   "Thu gọn thanh công cụ (bấm đúp vào tên tab cũng được)")

    def balance_corners(self):
        left = self.cornerWidget(Qt.TopLeftCorner)
        bw = self.b_collapse.sizeHint().width()
        w = max(left.sizeHint().width() if left else 0, bw + 16)
        self._pad.setFixedWidth(w - bw)

    def showEvent(self, e):
        self.balance_corners()                    # cỡ nút chỉ đúng sau khi áp stylesheet
        super().showEvent(e)

    def toggle(self):
        self.collapsed = not self.collapsed
        self.setFixedHeight(self.COLLAPSED if self.collapsed else self.FULL)
        self._collapse_icon()

    def tab(self, title, color=ACCENT):
        t = RibbonTab(color)
        sc = RibbonScroll()                       # cửa sổ hẹp: cuộn ngang thay vì bóp nhỏ nút, cắt chữ
        sc.setWidget(t)
        self.addTab(sc, title)
        return t


class RibbonScroll(QScrollArea):
    """Vùng cuộn ngang của 1 tab ribbon (cửa sổ hẹp): thanh cuộn ngang + lăn chuột trên ribbon = cuộn ngang."""

    def __init__(self):
        super().__init__()
        self.setObjectName("ribbonScroll")
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

    def wheelEvent(self, e):
        sb = self.horizontalScrollBar()
        if sb.maximum() > 0 and e.angleDelta().y():
            sb.setValue(sb.value() - e.angleDelta().y())
            e.accept()
            return
        super().wheelEvent(e)


# ===========================================================================
#  Hộp thoại chạy 1 công cụ trên file đang mở trong trình sửa
# ===========================================================================
class ToolDialog(QDialog):
    """in_place=True: kết quả thay vào tài liệu đang sửa (hoàn tác được). False: xuất file ra out_dir.
    src_path: 1 file hoặc danh sách file. pdf_actions: [(mã, nhãn)] - nút hiện ra sau khi xuất xong file PDF
    (vd. mở trong tab mới / chèn vào file đang mở); bấm nút -> accept() với result_action, result_paths."""

    def __init__(self, parent, tool, src_path, in_place, out_dir, pdf_actions=(), show_sources=False):
        super().__init__(parent)
        self.srcs = list(src_path) if isinstance(src_path, (list, tuple)) else [src_path]
        self.tool, self.src, self.in_place, self.out_dir = tool, self.srcs[0], in_place, out_dir
        self.result_path = None
        self.result_paths = []
        self.result_action = None
        self.pdf_actions = list(pdf_actions)
        self.outputs = []
        self.worker = None
        self.settings = QSettings(installer.ORG, "PDFToolkit")
        self.setWindowTitle(tool.name)
        main = QHBoxLayout(self)
        left = QWidget()
        lay = QVBoxLayout(left)
        lay.setContentsMargins(0, 0, 0, 0)
        main.addWidget(left, 5)
        self.preview = None
        if tool.id in PREVIEW_TOOLS:
            self.preview = PreviewPane()
            main.addWidget(self.preview, 6)
            self.resize(1180, 700)
        else:
            self.resize(640, 560)
        head = QHBoxLayout()
        ic = QLabel()
        ic.setPixmap(badge_pixmap(tool.icon, GROUP_COLORS.get(tool.group, ACCENT), 44))
        head.addWidget(ic, 0, Qt.AlignTop)
        col = QVBoxLayout()
        t = QLabel(tool.name)
        t.setStyleSheet("font-size:14pt; font-weight:700; color:#0f172a")
        d = QLabel(tool.desc + (f"\n{tool.note}" if tool.note else ""))
        d.setWordWrap(True)
        d.setStyleSheet("color:#475569")
        col.addWidget(t)
        col.addWidget(d)
        head.addLayout(col, 1)
        lay.addLayout(head)
        self.opts = OptionsForm()
        saved = {}
        try:
            saved = json.loads(self.settings.value(f"tools/opts/{tool.id}", "{}"))
        except ValueError:
            pass
        if show_sources:
            names = [os.path.basename(p) for p in self.srcs]
            more = f"  (+{len(names) - 6})" if len(names) > 6 else ""
            src_lbl = QLabel()
            src_lbl.setWordWrap(True)
            src_lbl.setStyleSheet("background:#f1f5f9; border:1px solid #e2e8f0; border-radius:8px; "
                                  "padding:6px 10px; color:#334155")
            with i18n.paused():
                src_lbl.setText(i18n.tr("File nguồn:") + " " + ", ".join(names[:6]) + more)
            lay.addWidget(src_lbl)
        self.opts.load(tool, saved, hide={"password"})
        self.opts.prefill(self.srcs)
        if self.preview:
            self.preview.set_tool(tool)
            self.preview.set_source(self.src, self.opts.values())
            self.opts.changed.connect(lambda: self.preview.options_changed(self.opts.values()))
        lay.addWidget(self.opts, 1)
        if not in_place and tool.output != "none":
            row = QHBoxLayout()
            row.addWidget(QLabel("Lưu kết quả vào:"))
            self.dir_edit = QLineEdit(out_dir)
            row.addWidget(self.dir_edit, 1)
            b = QPushButton("Chọn...")
            b.clicked.connect(self._pick_dir)
            row.addWidget(b)
            lay.addLayout(row)
        else:
            self.dir_edit = None
        self.status = QLabel("Kết quả áp dụng thẳng vào file đang sửa - có thể hoàn tác (Ctrl+Z)." if in_place else "")
        self.status.setStyleSheet("color:#475569")
        self.status.setWordWrap(True)
        lay.addWidget(self.status)
        self.progress = QProgressBar()
        self.progress.setRange(0, 1000)
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(8)
        lay.addWidget(self.progress)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(110)
        self.log.hide()
        lay.addWidget(self.log)
        btns = QHBoxLayout()
        self.b_folder = QPushButton("Mở thư mục kết quả")
        self.b_folder.clicked.connect(lambda: self.outputs and reveal(self.outputs[0]))
        self.b_folder.hide()
        btns.addWidget(self.b_folder)
        self.action_btns = []
        for code, label in self.pdf_actions:
            b = QPushButton(label)
            b.clicked.connect(lambda _=False, c=code: self._finish(c))
            b.hide()
            btns.addWidget(b)
            self.action_btns.append(b)
        btns.addStretch()
        self.run_btn = QPushButton("Áp dụng" if in_place else ("Kiểm tra" if tool.output == "none" else "Xuất file"))
        self.run_btn.setObjectName("primary")
        self.run_btn.setMinimumHeight(34)
        self.run_btn.clicked.connect(self.run)
        self.close_btn = QPushButton("Đóng")
        self.close_btn.clicked.connect(self.reject)
        btns.addWidget(self.run_btn)
        btns.addWidget(self.close_btn)
        lay.addLayout(btns)

    def _pick_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Thư mục lưu kết quả", self.dir_edit.text())
        if d:
            self.dir_edit.setText(d)

    def run(self):
        if self.worker:
            self.worker.cancelled = True
            return
        out_dir = self.dir_edit.text().strip() if self.dir_edit else self.out_dir
        if not os.path.isdir(out_dir):
            QMessageBox.information(self, self.tool.name, "Thư mục lưu kết quả không tồn tại.")
            return
        self.settings.setValue(f"tools/opts/{self.tool.id}", json.dumps(self.opts.savable(), ensure_ascii=False))
        self.log.clear()
        self.log.hide()
        for b in self.action_btns:
            b.hide()
        self.worker = Worker(self.tool, self.srcs, self.opts.values(), out_dir)
        self.worker.progress.connect(lambda f, m: (self.progress.setValue(int(f * 1000)), self.status.setText(m)))
        self.worker.log.connect(self._log)
        self.worker.done.connect(self._done)
        self.run_btn.setText("Dừng")
        self.close_btn.setEnabled(False)
        self.status.setText("Đang xử lý...")
        self.worker.start()

    def _log(self, m):
        self.log.show()
        self.log.appendPlainText(m)

    def _done(self, outputs, err):
        self.worker = None
        self.close_btn.setEnabled(True)
        self.outputs = outputs
        if err:
            self.run_btn.setText("Thử lại")
            self.status.setStyleSheet("color:#b91c1c")
            self.status.setText("✘ " + err)
            return
        self.progress.setValue(1000)
        if self.in_place:
            pdfs = [p for p in outputs if p.lower().endswith(".pdf") and os.path.isfile(p)]
            if pdfs:
                self.result_path = pdfs[0]
                self.accept()
                return
        self.status.setStyleSheet("color:#15803d")
        self.status.setText("✔ Xong" + (f" - {len(outputs)} kết quả" if outputs else ""))
        for p in outputs:
            self._log("→ " + p)
        self.b_folder.setVisible(bool(outputs))
        pdfs = self._pdf_outputs()
        for b in self.action_btns:
            b.setVisible(bool(pdfs))
        if pdfs and self.action_btns:
            self.action_btns[0].setObjectName("primary")
            self.action_btns[0].setStyle(self.action_btns[0].style())
            self.action_btns[0].setFocus()
        self.run_btn.setText("Chạy lại")
        self.run_btn.setObjectName("")
        self.run_btn.setStyle(self.run_btn.style())
        self.close_btn.setText("Đóng")

    def _pdf_outputs(self):
        return [p for p in self.outputs if p.lower().endswith(".pdf") and os.path.isfile(p)]

    def _finish(self, code):
        self.result_action = code
        self.result_paths = self._pdf_outputs()
        self.accept()

    def reject(self):
        if self.worker:
            return
        if self.preview:
            self.preview.stop()
        super().reject()

    def accept(self):
        if self.preview:
            self.preview.stop()
        super().accept()


TOOLS_STYLE += """
QToolButton#resumeBtn { padding:7px 30px 7px 14px; border:none; border-radius:10px; background:#2563eb; color:white;
                        font-weight:600; }
QToolButton#resumeBtn:hover { background:#1d4ed8; }
QToolButton#resumeBtn::menu-button { width:24px; border:none; border-left:1px solid #60a5fa;
                                     border-top-right-radius:10px; border-bottom-right-radius:10px; }
QToolButton#resumeBtn::menu-button:hover { background:#1e40af; }
#openCard { background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #eff6ff, stop:1 #f5f3ff);
            border:1px solid #bfdbfe; border-radius:14px; }
QPushButton#openBig { background:#2563eb; color:white; font-size:14pt; font-weight:700; border:none; border-radius:10px;
                      padding:8px 22px; }
QPushButton#openBig:hover { background:#1d4ed8; }
QPushButton#recent { background:white; border:1px solid #e2e8f0; border-radius:8px; padding:5px 12px; color:#1e293b; }
QPushButton#recent:hover { border-color:#93c5fd; }
#previewPane { background:#f1f5f9; border:1px solid #e2e8f0; border-radius:10px; }
#previewScroll, #previewView { background:transparent; color:#64748b; }
QToolButton#segBtn { padding:3px 12px; border:1px solid #cbd5e1; background:#ffffff; border-radius:6px; }
QToolButton#segBtn:checked { background:#2563eb; color:white; border-color:#2563eb; }
QToolButton#setupBtn { padding:7px 14px; border:1px solid #e2e8f0; border-radius:10px; background:#ffffff;
                       color:#334155; font-weight:600; }
QToolButton#setupBtn:hover { background:#f1f5f9; }
QToolButton#setupBtn[highlight="true"] { background:#2563eb; color:white; border-color:#2563eb; }
QToolButton#setupBtn::menu-indicator { image:none; width:0; }
#langSwitch { background:#f1f5f9; border:1px solid #e2e8f0; border-radius:10px; }
QToolButton#langBtn { padding:4px 12px; border-radius:7px; color:#475569; font-weight:600; border:none; }
QToolButton#langBtn:checked { background:#2563eb; color:white; }
QToolButton#langBtn:hover:!checked { background:#e2e8f0; }
#ribbonHost { background:qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #ece7fb, stop:0.5 #eef0fa, stop:1 #def5ee);
              border-bottom:1px solid #e4e1f0; }
#ribbon, #ribbonTab, #ribbonScroll, #ribbonScroll > QWidget { background:transparent; }
#ribbonScroll QScrollBar:horizontal { height:10px; background:#e2e8f0; border-radius:5px; margin:1px 32px; }
#ribbonScroll QScrollBar::handle:horizontal { background:#64748b; border-radius:4px; min-width:60px; margin:1px 0; }
#ribbonScroll QScrollBar::handle:horizontal:hover { background:#2563eb; }
QToolButton#ribbonArrow { background:#2563eb; border:none; border-radius:8px; }
QToolButton#ribbonArrow:hover { background:#1d4ed8; }
QToolButton#thumbsToggle { background:#ffffff; border:1px solid #cbd5e1; border-radius:6px; }
QToolButton#thumbsToggle:hover { background:#dbeafe; border-color:#2563eb; }
#ribbonScroll QScrollBar::add-line, #ribbonScroll QScrollBar::sub-line { width:0; }
#ribbon::pane { border:none; background:transparent; margin:0; }
#ribbon::tab-bar { alignment:center; }
#ribbon QTabBar { background:rgba(255,255,255,0.45); border:1px solid rgba(255,255,255,0.9); border-radius:17px; }
#ribbon QTabBar::tab { padding:5px 15px; margin:3px 1px; color:#3f3f46; background:transparent; border:none;
                       border-radius:13px; }
#ribbon QTabBar::tab:selected { background:#ffffff; color:#6d28d9; font-weight:600; }
#ribbon QTabBar::tab:hover:!selected { background:rgba(255,255,255,0.6); color:#18181b; }
#ribbonBar { background:rgba(255,255,255,0.58); border:1px solid rgba(255,255,255,0.95); border-radius:20px; }
#ribbonGroup { background:transparent; border:none; }
QFrame#ribbonSep { background:rgba(24,24,27,0.10); border:none; margin:9px 3px; }
QToolButton#ribbonBig { padding:4px 3px 3px 3px; font-size:8.5pt; border-radius:11px; color:#27272a; border:none; }
QToolButton#ribbonBig:hover { background:rgba(255,255,255,0.9); }
QToolButton#ribbonBig:pressed { background:rgba(109,40,217,0.14); }
QToolButton#ribbonBig:checked { background:#6d28d9; color:#ffffff; }
QToolButton#ribbonBig:checked:hover { background:#5b21b6; }
QToolButton#ribbonBig:disabled { color:#a1a1aa; }
QToolButton#ribbonBig::menu-indicator { image:none; width:0; }
QToolButton#ribbonCollapse { border-radius:8px; padding:4px; margin-right:12px; }
QToolButton#ribbonCollapse:hover { background:rgba(255,255,255,0.7); }
#quickBar QToolButton { border-radius:8px; padding:4px; }
#quickBar QToolButton:hover { background:rgba(255,255,255,0.7); }
#quickBar QToolButton:checked { background:rgba(109,40,217,0.16); }
"""

# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Thông tin & hỗ trợ: phiên bản, liên hệ IT, giấy phép bên thứ ba (đọc từ thư mục licenses\\ do
collect_licenses.py tạo lúc build)."""
import json, os, platform, re, shutil, sys

from PySide6.QtCore import QStandardPaths, Qt, QUrl, QUrlQuery
from PySide6.QtGui import QDesktopServices, QFont, QGuiApplication
from PySide6.QtWidgets import (QAbstractItemView, QDialog, QDialogButtonBox, QFileDialog, QHBoxLayout,
                               QHeaderView, QLabel, QLineEdit, QMessageBox, QPlainTextEdit, QPushButton,
                               QSplitter, QTableWidget,
                               QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget)

import i18n

SUPPORT_EMAIL = "tuanpa1702@gmail.com"


def licenses_dir():
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, "licenses")


def load_licenses():
    try:
        return json.load(open(os.path.join(licenses_dir(), "index.json"), encoding="utf-8"))
    except (OSError, ValueError):
        return []


def send_support_mail(version):
    """Mở chương trình email mặc định với thư soạn sẵn gửi IT."""
    url = QUrl("mailto:" + SUPPORT_EMAIL)
    q = QUrlQuery()
    q.addQueryItem("subject", f"[PDF Toolkit {version}] " + i18n.tr("Yêu cầu hỗ trợ"))
    body = (i18n.tr("Mô tả vấn đề:") + "\n\n\n" + "-" * 30 +
            f"\nPDF Toolkit {version}\nWindows {platform.version()}\n")
    q.addQueryItem("body", body)
    url.setQuery(q)
    QDesktopServices.openUrl(url)


class AboutDialog(QDialog):
    def __init__(self, parent, title, version, tab=0, install_info=""):
        super().__init__(parent)
        self.version = version
        self.setWindowTitle("Thông tin & hỗ trợ")
        self.resize(1000, 800)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 18, 24, 18)
        lay.setSpacing(12)
        lay.addLayout(self._header(title))
        tabs = QTabWidget()
        tabs.setDocumentMode(False)
        tabs.addTab(self._info_tab(version, install_info), "Thông tin && hỗ trợ")     # && = chữ &
        tabs.addTab(self._license_tab(), "Giấy phép bên thứ ba")
        tabs.setCurrentIndex(tab)
        lay.addWidget(tabs, 1)
        bb = QHBoxLayout()
        bb.addStretch()
        close = QPushButton("Đóng")
        close.setMinimumWidth(100)
        close.clicked.connect(self.reject)
        bb.addWidget(close)
        lay.addLayout(bb)

    @staticmethod
    def _header(title):
        """Tên app + phiên bản màu xanh + nhãn tác giả (giống trang chủ)."""
        m = re.match(r"(.*?)\s+(V[\d.]+)(?:\s+by\s+(.+))?$", title)
        name, ver, author = m.groups() if m else (title, "", None)
        col = QVBoxLayout()
        col.setSpacing(4)
        row = QHBoxLayout()
        row.setSpacing(10)
        h = QLabel()
        h.setObjectName("homeTitle")
        with i18n.paused():
            h.setText(name)
        row.addWidget(h)
        row.addStretch()
        col.addLayout(row)
        from tools_ui import version_row
        col.addLayout(version_row(ver, author))
        sr = QHBoxLayout()
        sr.setSpacing(6)
        dot = QLabel("●")
        dot.setStyleSheet("color:#22c55e; font-size:8pt")
        s = QLabel("Bộ công cụ PDF - chạy offline, file không rời khỏi máy")
        s.setStyleSheet("color:#64748b")
        sr.addWidget(dot)
        sr.addWidget(s)
        sr.addStretch()
        col.addLayout(sr)
        return col

    # ------------------------------------------------------------------ thông tin
    def _info_tab(self, version, install_info):
        w = QWidget()
        w.setObjectName("aboutPage")
        w.setStyleSheet("#aboutPage { background:transparent; }")
        v = QVBoxLayout(w)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(14)
        card = QWidget()
        card.setObjectName("infoCard")
        card.setStyleSheet("#infoCard { background:#f8fafc; border:1px solid #e2e8f0; border-radius:12px; }"
                           "#infoCard QLabel { background:transparent; }")
        cl = QVBoxLayout(card)
        cl.setContentsMargins(18, 8, 18, 8)
        cl.setSpacing(0)
        v.addWidget(card)
        rows = [("Phiên bản", version),
                ("Bản quyền", f"© 2026 TuanPA <{SUPPORT_EMAIL}>"),
                ("Giấy phép", "GNU AGPL v3 - phần mềm tự do: được dùng, sao chép, sửa đổi và phân phối lại theo "
                              "giấy phép. KHÔNG CÓ BẢO HÀNH. Mã nguồn đi kèm chương trình (tab Giấy phép)."),
                ("Dữ liệu", "Chạy offline - file không rời khỏi máy. Riêng dấu thời gian (TSA) chỉ gửi mã băm "
                            "của file tới máy chủ dấu thời gian.")]
        if install_info:
            rows.append(("Cài đặt", install_info))
        for i, (k, val) in enumerate(rows):
            if i:
                line = QWidget()
                line.setFixedHeight(1)
                line.setStyleSheet("background:#e2e8f0")
                cl.addWidget(line)
            r = QHBoxLayout()
            r.setContentsMargins(0, 9, 0, 9)
            a = QLabel(k)
            a.setFixedWidth(120)
            a.setStyleSheet("color:#64748b")
            a.setAlignment(Qt.AlignTop | Qt.AlignLeft)
            b = QLabel(val)
            b.setWordWrap(True)
            b.setStyleSheet("color:#0f172a")
            b.setTextInteractionFlags(Qt.TextSelectableByMouse)
            r.addWidget(a)
            r.addWidget(b, 1)
            cl.addLayout(r)

        box = QWidget()
        box.setObjectName("supportBox")
        box.setStyleSheet("#supportBox { background:#eff6ff; border:1px solid #bfdbfe; border-radius:14px; }")
        bl = QVBoxLayout(box)
        bl.setContentsMargins(22, 18, 22, 18)
        bl.setSpacing(8)
        t = QLabel("Cần hỗ trợ?")
        t.setStyleSheet("font-size:12pt; font-weight:700; color:#1e3a8a; background:transparent")
        bl.addWidget(t)
        d = QLabel("Gặp lỗi, cần thêm tính năng hoặc cần hướng dẫn sử dụng, vui lòng liên hệ qua email:")
        d.setWordWrap(True)
        d.setStyleSheet("background:transparent")
        bl.addWidget(d)
        mail = QLabel(f'<a href="mailto:{SUPPORT_EMAIL}" style="color:#1d4ed8; text-decoration:none">'
                      f'{SUPPORT_EMAIL}</a>')
        mail.setStyleSheet("font-size:14pt; font-weight:600; background:transparent")
        mail.setTextInteractionFlags(Qt.TextBrowserInteraction)
        mail.linkActivated.connect(lambda _: send_support_mail(version))
        bl.addWidget(mail)
        btns = QHBoxLayout()
        b1 = QPushButton("Gửi email hỗ trợ")
        b1.setObjectName("primary")
        b1.clicked.connect(lambda: send_support_mail(version))
        b2 = QPushButton("Sao chép địa chỉ email")
        b2.clicked.connect(lambda: (QGuiApplication.clipboard().setText(SUPPORT_EMAIL),
                                    b2.setText(i18n.tr("Đã sao chép"))))
        btns.addWidget(b1)
        btns.addWidget(b2)
        btns.addStretch()
        bl.addLayout(btns)
        tip = QLabel("Khi báo lỗi, nên gửi kèm ảnh chụp màn hình và file PDF gặp lỗi (nếu không chứa thông tin riêng tư, nhạy cảm).")
        tip.setWordWrap(True)
        tip.setStyleSheet("color:#475569; background:transparent")
        bl.addWidget(tip)
        v.addWidget(box)
        v.addStretch()
        return w

    # ------------------------------------------------------------------ giấy phép
    def _license_tab(self):
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(20, 18, 20, 18)
        v.setSpacing(10)
        tr = i18n.tr
        html = ("<b>PDF Toolkit</b> © 2026 TuanPA - " +
                tr("phần mềm tự do theo GNU AGPL v3, không có bảo hành. Mã nguồn đầy đủ đi kèm chương trình "
                   "(nút Lưu mã nguồn); mã nguồn các thư viện tải tại link chính thức trong danh sách dưới đây, "
                   "hoặc liên hệ") +
                f" {SUPPORT_EMAIL}.<br>" +
                tr("Chương trình dùng các thành phần mã nguồn mở dưới đây, mỗi thành phần thuộc bản quyền của tác "
                   "giả tương ứng và được dùng nguyên bản theo giấy phép ghi kèm.") +
                "<br>• <b>PyMuPDF / MuPDF</b>: GNU AGPL v3.  • <b>Qt / PySide6</b>: GNU LGPL v3 - " +
                tr("thư viện Qt nằm riêng trong thư mục _internal của bản cài và có thể thay bằng phiên bản "
                   "tương thích."))
        with i18n.paused():
            intro = QLabel(html)
        intro.setWordWrap(True)
        intro.setTextFormat(Qt.RichText)
        intro.setObjectName("licIntro")
        intro.setStyleSheet("#licIntro { background:#f8fafc; border:1px solid #e2e8f0; border-radius:12px; "
                            "padding:10px 16px; color:#334155; font-size:9pt; }")
        v.addWidget(intro)
        self.lic_filter = QLineEdit()
        self.lic_filter.setPlaceholderText("Lọc theo tên hoặc giấy phép...")
        self.lic_filter.setClearButtonEnabled(True)
        self.lic_filter.setMinimumHeight(34)
        try:
            from tools_ui import glyph_icon
            self.lic_filter.addAction(glyph_icon(0xE721, "#64748b"), QLineEdit.LeadingPosition)
        except ImportError:
            pass
        v.addWidget(self.lic_filter)

        self.items = load_licenses()
        sp = QSplitter(Qt.Vertical)
        sp.setHandleWidth(10)
        self.table = QTableWidget(len(self.items), 4)
        self.table.setHorizontalHeaderLabels([i18n.tr(x) for x in ("Thành phần", "Phiên bản", "Giấy phép",
                                                                    "Công dụng")])
        self.table.verticalHeader().hide()
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.setShowGrid(False)
        self.table.setAlternatingRowColors(False)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.Interactive)
        hh.setSectionResizeMode(3, QHeaderView.Stretch)
        self.table.setColumnWidth(2, 260)
        for r, e in enumerate(self.items):
            vals = (i18n.tr(e["name"]), e.get("version", ""), e.get("license", ""), i18n.tr(e.get("purpose", "")))
            for c, val in enumerate(vals):
                it = QTableWidgetItem(val)
                it.setToolTip(e.get("url", ""))
                self.table.setItem(r, c, it)
        sp.addWidget(self.table)
        self.text = QPlainTextEdit()
        self.text.setReadOnly(True)
        self.text.setFont(QFont("Consolas", 9))
        sp.addWidget(self.text)
        sp.setSizes([300, 260])
        v.addWidget(sp, 1)
        foot = QHBoxLayout()
        self.url_lbl = QLabel()
        self.url_lbl.setOpenExternalLinks(True)
        foot.addWidget(self.url_lbl, 1)
        if not self.items:
            self.text.setPlainText(i18n.tr("Không tìm thấy dữ liệu giấy phép (thư mục licenses)."))
        save_src = QPushButton("Lưu mã nguồn...")
        save_src.setToolTip("Lưu mã nguồn đầy đủ của PDF Toolkit (file zip) ra máy")
        save_src.clicked.connect(self._save_source)
        save_src.setEnabled(os.path.isfile(os.path.join(licenses_dir(), "PDF-Toolkit-source.zip")))
        foot.addWidget(save_src)
        open_dir = QPushButton("Mở thư mục giấy phép")
        open_dir.clicked.connect(lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(licenses_dir())))
        open_dir.setEnabled(os.path.isdir(licenses_dir()))
        foot.addWidget(open_dir)
        v.addLayout(foot)
        self.table.currentCellChanged.connect(lambda r, *_: self._show(r))
        self.lic_filter.textChanged.connect(self._filter)
        if self.items:
            self.table.selectRow(0)
        return w

    def _save_source(self):
        src = os.path.join(licenses_dir(), "PDF-Toolkit-source.zip")
        docs = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        dst, _ = QFileDialog.getSaveFileName(self, "Lưu mã nguồn PDF Toolkit",
                                             os.path.join(docs, f"PDF-Toolkit-{self.version}-source.zip"),
                                             "File zip (*.zip)")
        if not dst:
            return
        try:
            shutil.copyfile(src, dst)
        except OSError as e:
            QMessageBox.warning(self, "PDF Toolkit", str(e))
            return
        QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(dst)))

    def _show(self, row):
        if not 0 <= row < len(self.items):
            return
        e = self.items[row]
        try:
            txt = open(os.path.join(licenses_dir(), e["file"]), encoding="utf-8", errors="replace").read()
        except OSError:
            txt = f"{e['name']} - {e.get('license', '')}\n{e.get('url', '')}"
        with i18n.paused():                     # văn bản giấy phép giữ nguyên gốc
            self.text.setPlainText(txt)
            u = e.get("url", "")
            self.url_lbl.setText(f'<a href="{u}">{u}</a>' if u else "")

    def _filter(self, s):
        s = s.lower().strip()
        for r, e in enumerate(self.items):
            hay = " ".join((e["name"], e.get("license", ""), e.get("purpose", ""))).lower()
            self.table.setRowHidden(r, bool(s) and s not in hay)

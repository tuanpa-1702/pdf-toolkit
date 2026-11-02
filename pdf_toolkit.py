# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""PDF Toolkit - công cụ PDF nội bộ.

Chức năng: xem, xoá / xoay / sắp xếp trang (kéo thả), gộp file, tách trang, nén PDF,
sửa / thêm chữ có định dạng (font, đậm, nghiêng, màu...), tìm & thay thế,
chèn ảnh / liên kết / tô sáng / ghi chú / hình vẽ, ký số bằng USB Token. Có hoàn tác.
"""
import html
import json
import os
import re
import sys

import pymupdf
from PySide6.QtCore import Qt, QBuffer, QElapsedTimer, QObject, QSize, QThread, QTimer, QPoint, QPointF, QRect, QRectF, QSettings, Signal
from PySide6.QtGui import (QAction, QPageSize, QActionGroup, QColor, QFont, QIcon, QImage, QKeySequence, QPainter,
                           QPainterPath, QPalette, QPen, QPixmap, QBrush, QTextCharFormat, QTextCursor, QTransform)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QButtonGroup, QCheckBox, QColorDialog, QComboBox,
                               QDialog, QDialogButtonBox, QDoubleSpinBox, QFileDialog, QFormLayout, QHBoxLayout,
                               QInputDialog, QLabel, QLineEdit, QListView, QListWidget, QListWidgetItem, QMainWindow,
                               QMessageBox, QPushButton, QRadioButton, QScrollArea, QSpinBox, QSplitter,
                               QStackedWidget, QTableWidget, QTableWidgetItem, QTextEdit, QToolBar, QVBoxLayout, QWidget,
                               QToolButton, QMenu, QToolTip, QPlainTextEdit, QProgressDialog, QTabBar, QTabWidget,
                               QStyledItemDelegate, QStyleOptionViewItem, QGroupBox)

import about
import filedlg
import i18n
import installer
import pdf_core
import pdf_tools
import tools_ui
from pdf_core import PdfDocument, PdfError, get_text_lines, int_to_rgb

APP_NAME = "PDF Toolkit"
APP_VERSION = "1.3.0"
try:                        # cờ lúc build (pdf_toolkit.spec sinh ra): bản không ghi tác giả cạnh logo
    from build_flags import HIDE_AUTHOR
except ImportError:
    HIDE_AUTHOR = False
# chỉ bỏ chữ "by TuanPA" cạnh tên app; thông tin bản quyền (Hỗ trợ, NOTICE, LICENSE) giữ nguyên theo AGPL
APP_TITLE = f"{APP_NAME} V{APP_VERSION.rsplit('.', 1)[0]}" + ("" if HIDE_AUTHOR else " by TuanPA")
THUMB_SCALE = 0.28
OPEN_FILTER = "PDF / Ảnh (*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp);;PDF (*.pdf);;Tất cả (*.*)"


def fmt_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024 or unit == "GB":
            return f"{n:,.0f} {unit}" if unit in ("B", "KB") else f"{n:,.1f} {unit}"
        n /= 1024


_icon_font = None


def glyph_icon(code, color="#1e293b", mirror=False):
    """Icon vẽ từ font biểu tượng có sẵn của Windows (Segoe Fluent Icons / MDL2 Assets)."""
    global _icon_font
    if _icon_font is None:
        from PySide6.QtGui import QFontDatabase
        fams = QFontDatabase.families()
        _icon_font = next((f for f in ("Segoe Fluent Icons", "Segoe MDL2 Assets") if f in fams), "Segoe UI Symbol")
    pm = QPixmap(64, 64)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.TextAntialiasing)
    f = QFont(_icon_font)
    f.setPixelSize(40)
    p.setFont(f)
    p.setPen(QColor(color))
    p.drawText(pm.rect(), Qt.AlignCenter, chr(code))
    p.end()
    if mirror:
        from PySide6.QtGui import QTransform
        pm = pm.transformed(QTransform().scale(-1, 1))
    return QIcon(pm)


def pix_to_qimage(pix):
    fmt = (QImage.Format_Grayscale8 if pix.n == 1 else
           QImage.Format_RGBA8888 if pix.alpha else QImage.Format_RGB888)
    return QImage(pix.samples, pix.width, pix.height, pix.stride, fmt).copy()


# ===========================================================================
#  Danh sách trang thu nhỏ (kéo thả để sắp xếp)
# ===========================================================================
class _CenteredThumbDelegate(QStyledItemDelegate):
    """Ảnh trang nằm giữa ô (mặc định Qt đặt sát trái -> lề phải rộng hơn lề trái, nhất là trang ngang)."""

    def initStyleOption(self, opt, index):
        super().initStyleOption(opt, index)
        opt.decorationPosition = QStyleOptionViewItem.Top
        opt.decorationAlignment = Qt.AlignCenter
        opt.displayAlignment = Qt.AlignHCenter


class ThumbList(QListWidget):
    orderChanged = Signal()
    filesDropped = Signal(list, int)     # đường dẫn file, vị trí chèn

    def __init__(self):
        super().__init__()
        self.setViewMode(QListView.ListMode)
        self.setFlow(QListView.TopToBottom)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setIconSize(QSize(150, 200))
        self.setSpacing(6)
        self.setUniformItemSizes(True)
        self.setAcceptDrops(True)
        self.setResizeMode(QListView.Adjust)
        self.setItemDelegate(_CenteredThumbDelegate(self))

    def set_grid(self, grid: bool):
        self.setFlow(QListView.LeftToRight if grid else QListView.TopToBottom)
        self.setWrapping(grid)
        self.setIconSize(QSize(170, 230) if grid else QSize(150, 200))
        self.setSpacing(12 if grid else 6)

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
            paths = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
            item = self.itemAt(e.position().toPoint())
            at = self.row(item) + 1 if item else self.count()
            e.acceptProposedAction()
            self.filesDropped.emit(paths, at)
            return
        super().dropEvent(e)
        QTimer.singleShot(0, self.orderChanged.emit)


# ===========================================================================
#  Vùng hiển thị trang
# ===========================================================================
class PageCanvas(QWidget):
    """mode: view | edit (bấm dòng chữ) | rect (kéo khung) | line (kéo đường) | point (bấm 1 điểm)."""
    lineClicked = Signal(int)
    pointClicked = Signal(float, float)       # toạ độ hiển thị (pt)
    rectDrawn = Signal(QRectF)                # toạ độ hiển thị (pt)
    segmentDrawn = Signal(QPointF, QPointF)
    noteClicked = Signal(int)                 # xref của ghi chú được bấm
    objectMoved = Signal(int, QRectF)         # ảnh / chữ ký / con dấu thứ i được kéo tới vùng mới (pt hiển thị)
    objectMenu = Signal(int, QPoint)          # bấm chuột phải lên ảnh
    objectClicked = Signal(int)               # bấm (không kéo) lên vật có thể bấm - vd liên kết
    objectSelected = Signal(int)              # vật đang chọn (để căn chỉnh / xoá bằng thanh công cụ); -1 = bỏ chọn
    objectActivated = Signal(int)             # bấm đúp / Enter lên vật (vd khung chữ -> mở trình sửa chữ)
    NUDGE_MS = 450                            # dời bằng phím mũi tên: ngừng bấm bao lâu thì ghi vào file
    strokeDrawn = Signal(list)                # chế độ free: [(x, y) pt hiển thị] của 1 nét vẽ tay
    polyDrawn = Signal(list)                  # chế độ poly: các đỉnh (bấm đúp / Enter / chuột phải để kết thúc)
    selectionMenu = Signal(QPoint)            # chế độ select: chuột phải lên vùng chữ đã chọn
    selectionChanged = Signal(int)            # số từ đang chọn
    regionChanged = Signal(object)            # Chụp vùng: khung vùng chọn mới (QRectF pt hiển thị) / None = bỏ chọn
    regionMenu = Signal(QPoint)               # chuột phải trong khung vùng chọn
    pageMenu = Signal(QPointF, QPoint)        # chuột phải lên chỗ trống của trang (pt hiển thị, vị trí màn hình)
    HANDLE = 9                                # cỡ ô kéo đổi kích thước (góc dưới phải)
    _REGION_CURSORS = {(0, 0): Qt.SizeFDiagCursor, (2, 2): Qt.SizeFDiagCursor, (2, 0): Qt.SizeBDiagCursor,
                       (0, 2): Qt.SizeBDiagCursor, (1, 0): Qt.SizeVerCursor, (1, 2): Qt.SizeVerCursor,
                       (0, 1): Qt.SizeHorCursor, (2, 1): Qt.SizeHorCursor, "move": Qt.SizeAllCursor}

    def __init__(self):
        super().__init__()
        self.setMouseTracking(True)
        self.pixmap = None
        self.zoom = 1.0
        self.mode = "view"
        self.lines = []
        self.rects = []        # vùng các dòng chữ trên màn hình
        self.hover = -1
        self.notes = []        # [(vùng trên màn hình, xref, tiêu đề, nội dung)] - ghi chú / chú thích có nội dung
        self.note_hover = -1
        self.drag = None       # (điểm bắt đầu, điểm hiện tại) theo pixel màn hình
        self.objs = []         # ảnh kéo được: [vùng trên màn hình]
        self.obj_hover = -1
        self.obj_drag = None   # {i, kind: move|size, start, orig, cur}
        self.obj_sel = -1      # vật đang chọn (bấm vào vật)
        self.obj_resizable = []
        self._nudge_timer = QTimer(self)
        self._nudge_timer.setSingleShot(True)
        self._nudge_timer.setInterval(self.NUDGE_MS)
        self._nudge_timer.timeout.connect(self._commit_nudge)
        self.obj_ratio = []    # đổi cỡ giữ tỉ lệ? (ảnh: có, khối vẽ: không)
        self.obj_edge = False
        self.obj_tips = []
        self.words = []        # chế độ select: [(vùng màn hình, chữ, ô chưa xoay, (khối, dòng))] theo thứ tự đọc
        self.sel = None        # (từ đầu, từ cuối) đang chọn
        self.stroke = []       # chế độ free: nét đang vẽ (pixel màn hình)
        self.poly = []         # chế độ poly: các đỉnh đã bấm (pixel màn hình)
        self.poly_hover = None
        self.region = None     # Chụp vùng: khung vùng chọn còn giữ lại (pixel màn hình) - kéo / đổi cỡ / chuột phải
        self.region_drag = None
        self.setFocusPolicy(Qt.ClickFocus)
        self.setStyleSheet("background:#e5e7eb")

    # ----- chọn chữ -----
    # self.words = các ký tự theo thứ tự đọc; self.sel = (mốc neo, mốc hiện tại) - mốc là vị trí GIỮA 2 ký tự
    def _char_at(self, pos):
        """Ký tự chứa điểm (hoặc gần nhất): (chỉ số, điểm nằm ở nửa phải ký tự?) - (-1, False) nếu trang không có chữ."""
        best, bd = -1, None
        for i, (r, *_x) in enumerate(self.words):
            if r.adjusted(0, -2, 0, 2).contains(pos):
                return i, pos.x() > r.center().x()
            dx = max(r.left() - pos.x(), 0, pos.x() - r.right())
            dy = max(r.top() - pos.y(), 0, pos.y() - r.bottom())
            d = dx + dy * 3                                     # ưu tiên cùng dòng
            if bd is None or d < bd:
                best, bd = i, d
        if best < 0:
            return -1, False
        return best, pos.x() > self.words[best][0].center().x()

    def _boundary(self, pos):
        i, right = self._char_at(pos)
        return -1 if i < 0 else i + (1 if right else 0)

    def _word_span(self, i):
        """[đầu, cuối) của từ chứa ký tự i (cắt ở dấu cách / đầu dòng)."""
        key = self.words[i][3]
        a = b = i
        while a > 0 and self.words[a - 1][3] == key and not self.words[a - 1][1].isspace():
            a -= 1
        while b + 1 < len(self.words) and self.words[b + 1][3] == key and not self.words[b + 1][1].isspace():
            b += 1
        return a, b + 1

    def sel_range(self):
        if not self.sel or self.sel[0] < 0 or self.sel[1] < 0:
            return []
        a, b = sorted(self.sel)
        return list(range(a, b))

    def selected_text(self):
        out, last = [], None
        for i in self.sel_range():
            _r, ch, _u, key = self.words[i]
            if last is not None and key != last:
                out.append("\n")
            out.append(ch)
            last = key
        return "\n".join(s.strip() for s in "".join(out).split("\n"))

    def _sel_runs(self):
        """Các đoạn liền nhau trên cùng dòng của vùng chọn: [(ô màn hình, ô chưa xoay)]."""
        runs = []
        for i in self.sel_range():
            r, ch, u, key = self.words[i]
            if runs and runs[-1][2] == key:
                runs[-1][0] = runs[-1][0].united(r)
                runs[-1][1] = runs[-1][1] | u
            else:
                runs.append([QRectF(r), pymupdf.Rect(u), key])
        return [(r, u) for r, u, _k in runs]

    def selected_rects(self):
        """Ô các đoạn chữ đang chọn (toạ độ trang chưa xoay) - để tô sáng / gạch chân."""
        return [u for _r, u in self._sel_runs() if not u.is_empty]

    def select_all(self):
        if self.words:
            self.sel = (0, len(self.words))
            self.update()
            self.selectionChanged.emit(len(self.words))

    def clear_selection(self):
        if self.sel:
            self.sel = None
            self.update()
            self.selectionChanged.emit(0)

    def _commit_nudge(self):
        """Ngừng bấm phím mũi tên: ghi vị trí mới của vật vào file (như thả chuột sau khi kéo)."""
        d = self.obj_drag
        if not d or d.get("kind") != "nudge":
            return
        self.obj_drag = None
        self.update()
        c, z = d["cur"], self.zoom
        if (c.topLeft() - d["orig"].topLeft()).manhattanLength() > 0.01:
            self.objectMoved.emit(d["i"], QRectF((c.x() - 20) / z, (c.y() - 20) / z, c.width() / z, c.height() / z))

    # ----- Chụp vùng: khung vùng chọn giữ lại trên trang -----
    def set_region(self, r_pt):
        """Hiện khung vùng chọn (QRectF pt hiển thị) - None: bỏ."""
        z = self.zoom
        self.region = None if r_pt is None else QRectF(20 + r_pt.x() * z, 20 + r_pt.y() * z,
                                                       r_pt.width() * z, r_pt.height() * z)
        self.region_drag = None
        self.update()

    def clear_region(self):
        if self.region is not None:
            self.region = None
            self.region_drag = None
            self.update()
            self.regionChanged.emit(None)

    def _region_pt(self):
        r, z = self.region, self.zoom
        return QRectF((r.x() - 20) / z, (r.y() - 20) / z, r.width() / z, r.height() / z)

    def _page_area(self):
        dpr = self.pixmap.devicePixelRatio()
        return QRectF(20, 20, self.pixmap.width() / dpr, self.pixmap.height() / dpr)

    def _region_handles(self):
        r = self.region
        xs, ys = (r.left(), r.center().x(), r.right()), (r.top(), r.center().y(), r.bottom())
        return [(QPointF(xs[i], ys[j]), (i, j)) for j in range(3) for i in range(3) if (i, j) != (1, 1)]

    def _region_hit(self, pos):
        """Ô kéo (i, j) / 'move' (trong khung) / None (ngoài khung)."""
        if self.region is None:
            return None
        for q, ij in self._region_handles():
            if abs(pos.x() - q.x()) <= 6 and abs(pos.y() - q.y()) <= 6:
                return ij
        return "move" if self.region.adjusted(-3, -3, 3, 3).contains(pos) else None

    def _drag_region(self, pos):
        d = self.region_drag
        o, dl, hit, page = d["orig"], pos - d["start"], d["hit"], self._page_area()
        if hit == "move":                                       # dời cả khung, không ra ngoài trang
            r = o.translated(dl)
            r.moveLeft(min(max(r.left(), page.left()), page.right() - r.width()))
            r.moveTop(min(max(r.top(), page.top()), page.bottom() - r.height()))
        else:                                                   # kéo cạnh / góc
            i, j = hit
            l, t, rr, b = o.left(), o.top(), o.right(), o.bottom()
            l += dl.x() if i == 0 else 0
            rr += dl.x() if i == 2 else 0
            t += dl.y() if j == 0 else 0
            b += dl.y() if j == 2 else 0
            r = QRectF(QPointF(l, t), QPointF(rr, b)).normalized().intersected(page)
        if r.width() >= 4 and r.height() >= 4:
            self.region = r
            self.update()

    def keyPressEvent(self, e):
        k = e.key()
        if self.region is not None and k == Qt.Key_Escape:
            self.clear_region()
            e.accept()
            return
        if 0 <= self.obj_sel < len(self.objs) and not self.drag and not self.stroke and not self.poly:
            arrows = {Qt.Key_Left: (-1, 0), Qt.Key_Right: (1, 0), Qt.Key_Up: (0, -1), Qt.Key_Down: (0, 1)}
            if k in arrows and not (e.modifiers() & (Qt.ControlModifier | Qt.AltModifier)):
                step = 10 if e.modifiers() & Qt.ShiftModifier else 1          # pt
                dx, dy = arrows[k]
                d = self.obj_drag
                if not d or d.get("kind") != "nudge":
                    r = QRectF(self.objs[self.obj_sel])
                    d = self.obj_drag = {"i": self.obj_sel, "kind": "nudge", "start": None, "orig": r,
                                         "cur": QRectF(r)}
                d["cur"] = d["cur"].translated(dx * step * self.zoom, dy * step * self.zoom)
                self.update()
                self._nudge_timer.start()
                e.accept()
                return
            if k in (Qt.Key_Return, Qt.Key_Enter):
                self._nudge_timer.stop()
                self._commit_nudge()
                self.objectActivated.emit(self.obj_sel)
                e.accept()
                return
        if self.mode == "poly" and self.poly:
            if e.key() in (Qt.Key_Return, Qt.Key_Enter):
                self._finish_poly()
                return
            if e.key() == Qt.Key_Escape:
                self.poly = []
                self.update()
                return
            if e.key() == Qt.Key_Backspace:
                self.poly.pop()
                self.update()
                return
        super().keyPressEvent(e)

    def _finish_poly(self):
        pts = [self.to_pt(p) for p in self.poly]
        self.poly, self.poly_hover = [], None
        self.update()
        if len(pts) >= 2:
            self.polyDrawn.emit([(p.x(), p.y()) for p in pts])

    def set_objects(self, rects, keep_ratio=None, edge_only=False, tips=None, resizable=None):
        """rects: vùng hiển thị (pt) của ảnh / chữ ký / con dấu / khối vẽ / khung chữ kéo được trên trang.
        keep_ratio[i]: đổi cỡ giữ tỉ lệ (ảnh) hay tự do (khối vẽ). edge_only: chỉ nắm được ở viền khối
        (đang ở chế độ vẽ hình -> bấm bên trong khối vẫn vẽ hình mới được). resizable[i]: có ô đổi cỡ không."""
        self._nudge_timer.stop()
        z = self.zoom
        self.objs = [QRectF(20 + r.x0 * z, 20 + r.y0 * z, r.width * z, r.height * z) for r in rects]
        self.obj_ratio = list(keep_ratio) if keep_ratio is not None else [True] * len(self.objs)
        self.obj_resizable = list(resizable) if resizable is not None else [True] * len(self.objs)
        self.obj_edge = edge_only
        self.obj_tips = list(tips) if tips is not None else [None] * len(self.objs)   # có chữ -> bấm được (liên kết)
        self.obj_hover = -1
        self.obj_drag = None
        self.obj_sel = -1
        self.update()

    def select_object(self, i, notify=True):
        i = i if 0 <= i < len(self.objs) else -1
        if i != self.obj_sel:
            self.obj_sel = i
            self.update()
            if notify:
                self.objectSelected.emit(i)

    def _obj_at(self, pos):
        for i in reversed(range(len(self.objs))):               # vật vẽ sau nằm trên
            r = self.objs[i]
            if not r.adjusted(-4, -4, 4, 4).contains(pos):
                continue
            inner = r.adjusted(6, 6, -6, -6)
            if self.obj_edge and inner.width() > 6 and inner.height() > 6 and inner.contains(pos) \
                    and not self._on_handle(i, pos):
                continue
            return i
        return -1

    def _on_handle(self, i, pos):
        if i < len(self.obj_resizable) and not self.obj_resizable[i]:
            return False
        r = self.objs[i]
        return abs(pos.x() - r.right()) <= self.HANDLE and abs(pos.y() - r.bottom()) <= self.HANDLE

    def _base_cursor(self):
        return Qt.IBeamCursor if self.mode in ("edit", "select") else Qt.ArrowCursor if self.mode == "view" \
            else Qt.CrossCursor

    def show_page(self, page, zoom, mode):
        self.zoom, self.mode = zoom, mode
        self.drag = None
        self.region, self.region_drag = None, None              # trang chủ đặt lại khung (nếu còn) sau khi vẽ
        self.objs, self.obj_hover, self.obj_drag, self.obj_sel = [], -1, None, -1
        dpr = self.devicePixelRatioF()
        pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom * dpr, zoom * dpr), alpha=False)
        img = pix_to_qimage(pix)
        img.setDevicePixelRatio(dpr)
        self.pixmap = QPixmap.fromImage(img)
        self.setFixedSize(int(pix.width / dpr) + 40, int(pix.height / dpr) + 40)
        self.lines, self.rects = [], []
        if mode == "edit":
            self.lines = get_text_lines(page)
            m = page.rotation_matrix
            for ln in self.lines:
                r = ln.bbox * m
                self.rects.append(QRectF(20 + r.x0 * zoom, 20 + r.y0 * zoom, r.width * zoom, r.height * zoom))
        self.words, self.sel, self.stroke, self.poly, self.poly_hover = [], None, [], [], None
        if mode == "select":                       # chọn theo từng ký tự (như Word / Acrobat)
            m = page.rotation_matrix
            raw = page.get_text("rawdict", sort=True, flags=pymupdf.TEXT_PRESERVE_WHITESPACE)
            for bi, b in enumerate(raw["blocks"]):
                for li, ln in enumerate(b.get("lines", [])):
                    for sp in ln["spans"]:
                        for ch in sp["chars"]:
                            u = pymupdf.Rect(ch["bbox"])
                            if u.is_empty and ch["c"] != " ":
                                if self.words and self.words[-1][3] == (bi, li):   # dấu tổ hợp: gộp vào chữ trước
                                    r0, c0, u0, k0 = self.words[-1]
                                    self.words[-1] = (r0, c0 + ch["c"], u0, k0)
                                continue
                            r = u * m
                            r.normalize()
                            self.words.append((QRectF(20 + r.x0 * zoom, 20 + r.y0 * zoom, max(r.width * zoom, 1),
                                                      r.height * zoom), ch["c"], u, (bi, li)))
        self.hover = -1
        self.notes, self.note_hover = [], -1
        m = page.rotation_matrix
        for a in page.annots() or []:
            info = a.info
            if a.type[0] in PdfDocument.MOVABLE_ANNOTS:
                continue                          # khối vẽ: kéo để di chuyển, ghi chú xem qua menu chuột phải
            if a.type[0] == pymupdf.PDF_ANNOT_TEXT or (info.get("content") or "").strip():
                r = a.rect * m
                self.notes.append((QRectF(20 + r.x0 * zoom, 20 + r.y0 * zoom, max(r.width * zoom, 14),
                                          max(r.height * zoom, 14)), a.xref, info.get("title", ""),
                                   info.get("content", "")))
        self.setCursor(self._base_cursor())
        self.update()

    def to_pt(self, pos):
        w = self.pixmap.width() / self.pixmap.devicePixelRatio()
        h = self.pixmap.height() / self.pixmap.devicePixelRatio()
        x = min(max(pos.x() - 20, 0), w)
        y = min(max(pos.y() - 20, 0), h)
        return QPointF(x / self.zoom, y / self.zoom)

    def clear(self):
        self.pixmap = None
        self.lines, self.rects = [], []
        self.setFixedSize(10, 10)
        self.update()

    def paintEvent(self, e):
        if not self.pixmap:
            return
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w = self.pixmap.width() / self.pixmap.devicePixelRatio()
        h = self.pixmap.height() / self.pixmap.devicePixelRatio()
        p.fillRect(QRectF(23, 23, w, h), QColor(0, 0, 0, 40))          # bóng đổ
        p.drawPixmap(QPointF(20, 20), self.pixmap)
        if self.mode == "edit":
            for i, r in enumerate(self.rects):
                if i == self.hover:
                    p.fillRect(r.adjusted(-2, -2, 2, 2), QColor(37, 99, 235, 45))
                    p.setPen(QPen(QColor("#2563eb"), 1.5))
                else:
                    p.setPen(QPen(QColor(37, 99, 235, 90), 1, Qt.DashLine))
                p.drawRect(r.adjusted(-2, -2, 2, 2))
        if self.obj_drag:
            o, c = self.obj_drag["orig"], self.obj_drag["cur"]
            dpr = self.pixmap.devicePixelRatio()
            p.fillRect(o, QColor(255, 255, 255, 170))                   # chỗ cũ mờ đi
            src = QRectF((o.x() - 20) * dpr, (o.y() - 20) * dpr, o.width() * dpr, o.height() * dpr)
            p.setOpacity(0.85)
            p.drawPixmap(c, self.pixmap, src)
            p.setOpacity(1)
            p.setPen(QPen(QColor("#2563eb"), 1.5))
            p.setBrush(Qt.NoBrush)
            p.drawRect(c)
        else:
            if 0 <= self.obj_sel < len(self.objs):                       # vật đang chọn: viền liền + nền nhạt
                r = self.objs[self.obj_sel].adjusted(-2, -2, 2, 2)
                p.setPen(QPen(QColor("#2563eb"), 2))
                p.setBrush(QColor(37, 99, 235, 28))
                p.drawRect(r)
            for i in {self.obj_hover, self.obj_sel}:
                if not 0 <= i < len(self.objs):
                    continue
                r = self.objs[i]
                if i != self.obj_sel:
                    p.setPen(QPen(QColor("#2563eb"), 1.2, Qt.DashLine))
                    p.setBrush(Qt.NoBrush)
                    p.drawRect(r)
                if i < len(self.obj_resizable) and not self.obj_resizable[i]:
                    continue
                h = self.HANDLE
                p.setPen(QPen(QColor("#ffffff"), 1))
                p.setBrush(QColor("#2563eb"))
                p.drawRect(QRectF(r.right() - h / 2, r.bottom() - h / 2, h, h))
        if self.mode == "select" and self.sel:
            p.setPen(Qt.NoPen)
            p.setBrush(QColor(37, 99, 235, 70))
            for r, _u in self._sel_runs():
                p.drawRect(r.adjusted(0, -1, 0, 1))
        if len(self.stroke) > 1:
            p.setPen(QPen(QColor("#dc2626"), 2, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            p.setBrush(Qt.NoBrush)
            p.drawPolyline(self.stroke)
        if self.poly:
            p.setPen(QPen(QColor("#2563eb"), 1.5, Qt.DashLine))
            p.setBrush(Qt.NoBrush)
            pts = self.poly + ([self.poly_hover] if self.poly_hover is not None else [])
            p.drawPolyline(pts)
            p.setBrush(QColor("#2563eb"))
            for q in self.poly:
                p.drawEllipse(q, 3, 3)
        if self.drag:
            a, b = self.drag
            p.setPen(QPen(QColor("#2563eb"), 1.5, Qt.DashLine))
            if self.mode == "line":
                p.drawLine(a, b)
            else:
                p.setBrush(QColor(37, 99, 235, 30))
                p.drawRect(QRectF(a, b).normalized())
        if self.region is not None:                                    # khung Chụp vùng: viền đứt + 8 ô kéo
            p.setPen(QPen(QColor("#2563eb"), 1.5, Qt.DashLine))
            p.setBrush(QColor(37, 99, 235, 22))
            p.drawRect(self.region)
            p.setPen(QPen(QColor("#2563eb"), 1.2))
            p.setBrush(QColor("#ffffff"))
            for q, _ij in self._region_handles():
                p.drawRect(QRectF(q.x() - 4, q.y() - 4, 8, 8))

    def _hit(self, pos):
        for i, r in enumerate(self.rects):
            if r.adjusted(-2, -2, 2, 2).contains(pos):
                return i
        return -1

    def _note_at(self, pos):
        for i, (r, *_rest) in enumerate(self.notes):
            if r.adjusted(-3, -3, 3, 3).contains(pos):
                return i
        return -1

    def mouseMoveEvent(self, e):
        pos = e.position()
        if self.mode == "select" and self.sel is not None and e.buttons() & Qt.LeftButton:
            j = self._boundary(pos)
            if j >= 0 and self.sel[0] >= 0 and j != self.sel[1]:
                self.sel = (self.sel[0], j)
                self.update()
            return
        if self.mode == "free" and self.stroke:
            if (pos - self.stroke[-1]).manhattanLength() >= 2:
                self.stroke.append(pos)
                self.update()
            return
        if self.mode == "poly" and self.poly:
            self.poly_hover = pos
            self.update()
            return
        if self.region_drag:
            self._drag_region(pos)
            return
        if self.region is not None and not self.drag:
            hit = self._region_hit(pos)
            self.setCursor(self._REGION_CURSORS.get(hit, self._base_cursor()))
            if hit is not None:
                return
        if self.obj_drag and self.obj_drag["kind"] != "nudge":
            d = self.obj_drag
            o = d["orig"]
            delta = pos - d["start"]
            if d["kind"] == "move":
                d["cur"] = o.translated(delta)
            elif self.obj_ratio[d["i"]]:                                # đổi cỡ, giữ tỉ lệ ảnh
                k = max(12 / max(o.width(), 1), (o.width() + delta.x()) / max(o.width(), 1),
                        (o.height() + delta.y()) / max(o.height(), 1))
                d["cur"] = QRectF(o.x(), o.y(), o.width() * k, o.height() * k)
            else:                                                       # khối vẽ: kéo tự do 2 chiều
                d["cur"] = QRectF(o.x(), o.y(), max(4.0, o.width() + delta.x()), max(4.0, o.height() + delta.y()))
            self.update()
            return
        if not self.drag and self.objs:
            i = self._obj_at(pos)
            if i != self.obj_hover:
                self.obj_hover = i
                self.update()
                tip = self.obj_tips[i] if 0 <= i < len(self.obj_tips) else None
                if tip:
                    QToolTip.showText(e.globalPosition().toPoint(), tip, self)
                else:
                    QToolTip.hideText()
            if i >= 0:
                link = bool(self.obj_tips[i]) if i < len(self.obj_tips) else False
                self.setCursor(Qt.SizeFDiagCursor if self._on_handle(i, pos) else
                               Qt.PointingHandCursor if link else Qt.SizeAllCursor)
                return
            else:
                self.setCursor(self._base_cursor())
        if not self.drag and self.mode in ("view", "note", "edit"):
            n = self._note_at(e.position())
            if n != self.note_hover:
                self.note_hover = n
                if n >= 0:
                    _, _, title, content = self.notes[n]
                    tip = (f"<b>{html.escape(title)}</b><br>" if title else "") + \
                        html.escape(content[:600]).replace("\n", "<br>") + \
                        "<br><i style='color:#64748b'>Bấm để xem / sửa ghi chú</i>"
                    QToolTip.showText(e.globalPosition().toPoint(), tip, self)
                    self.setCursor(Qt.PointingHandCursor)
                else:
                    QToolTip.hideText()
                    self.setCursor(Qt.IBeamCursor if self.mode == "edit" else
                                   Qt.ArrowCursor if self.mode == "view" else Qt.CrossCursor)
            if n >= 0:
                return
        if self.mode == "edit":
            h = self._hit(e.position())
            if h != self.hover:
                self.hover = h
                self.update()
        elif self.drag:
            self.drag = (self.drag[0], e.position())
            self.update()

    def mousePressEvent(self, e):
        if not self.pixmap:
            return
        if self.obj_drag and self.obj_drag["kind"] == "nudge":           # đang dời bằng phím: ghi luôn
            self._nudge_timer.stop()
            self._commit_nudge()
        pos = e.position()
        if self.mode == "select":
            if e.button() == Qt.RightButton:
                i, _right = self._char_at(pos)
                on_char = i >= 0 and self.words[i][0].adjusted(0, -2, 0, 2).contains(pos)
                if on_char and i not in self.sel_range():               # chuột phải lên chữ chưa chọn: chọn từ đó
                    self.sel = self._word_span(i)
                    self.update()
                self.selectionMenu.emit(e.globalPosition().toPoint())
            elif e.button() == Qt.LeftButton:
                j = self._boundary(pos)
                if e.modifiers() & Qt.ShiftModifier and self.sel and self.sel[0] >= 0 and j >= 0:
                    self.sel = (self.sel[0], j)                         # Shift+bấm: nới vùng chọn
                else:
                    self.sel = (j, j)
                self.update()
            return
        if self.mode == "free" and e.button() == Qt.LeftButton:
            self.stroke = [pos]
            return
        if self.mode == "poly":
            if e.button() == Qt.RightButton:
                self._finish_poly()
            elif e.button() == Qt.LeftButton:
                self.poly.append(pos)
                self.update()
            return
        if self.region is not None:
            hit = self._region_hit(pos)
            if hit is not None:
                if e.button() == Qt.RightButton:
                    self.regionMenu.emit(e.globalPosition().toPoint())
                elif e.button() == Qt.LeftButton:
                    self.region_drag = {"hit": hit, "start": pos, "orig": QRectF(self.region)}
                return
            if e.button() != Qt.LeftButton:
                return
            self.clear_region()                                         # bấm ra ngoài khung: kéo vùng mới
        i = self._obj_at(pos) if self.objs else -1
        if e.button() == Qt.RightButton and i >= 0:
            if not (i < len(self.obj_tips) and self.obj_tips[i]):
                self.select_object(i)
            self.objectMenu.emit(i, e.globalPosition().toPoint())
            return
        if e.button() == Qt.RightButton:
            self.pageMenu.emit(self.to_pt(pos), e.globalPosition().toPoint())
            return
        if e.button() != Qt.LeftButton:
            return
        if i < 0 and self.obj_sel >= 0:                                 # bấm ra ngoài: bỏ chọn vật
            self.select_object(-1)
        if i >= 0 and self._note_at(pos) < 0:                           # kéo ảnh / chữ ký / con dấu
            self.obj_drag = {"i": i, "kind": "size" if self._on_handle(i, pos) else "move",
                             "start": pos, "orig": QRectF(self.objs[i]), "cur": QRectF(self.objs[i])}
            return
        if self.mode in ("view", "note", "edit"):
            n = self._note_at(pos)
            if n >= 0:
                self.noteClicked.emit(self.notes[n][1])
                return
        if self.mode == "edit":
            h = self._hit(pos)
            if h >= 0:
                self.lineClicked.emit(h)
        elif self.mode in ("rect", "line"):
            self.drag = (pos, pos)
        elif self.mode == "point":
            pt = self.to_pt(pos)
            self.pointClicked.emit(pt.x(), pt.y())

    def mouseDoubleClickEvent(self, e):
        if self.objs and self.mode not in ("poly", "free", "select") and e.button() == Qt.LeftButton:
            i = self._obj_at(e.position())
            if i >= 0 and not (i < len(self.obj_tips) and self.obj_tips[i]):
                self.select_object(i)
                self.objectActivated.emit(i)                          # bấm đúp: sửa (khung chữ...)
                return
        if self.mode == "poly" and self.poly:
            if len(self.poly) > 1 and (self.poly[-1] - e.position()).manhattanLength() < 6:
                self.poly.pop()                                         # điểm thừa do lần bấm thứ 2 của bấm đúp
            self._finish_poly()
            return
        if self.mode == "select":                                       # bấm đúp: chọn 1 từ
            i, _right = self._char_at(e.position())
            if i >= 0:
                self.sel = self._word_span(i)
                self.update()
                self.selectionChanged.emit(len(self.sel_range()))
            return
        if self.mode == "point":                                        # bấm đúp không đặt thêm ghi chú / dấu thứ 2
            return
        super().mouseDoubleClickEvent(e)

    def mouseReleaseEvent(self, e):
        if self.mode == "select" and e.button() == Qt.LeftButton:
            self.selectionChanged.emit(len(self.sel_range()))
            return
        if self.mode == "free" and e.button() == Qt.LeftButton and self.stroke:
            pts, self.stroke = self.stroke, []
            self.update()
            if len(pts) >= 2:
                self.strokeDrawn.emit([(q.x(), q.y()) for q in (self.to_pt(x) for x in pts)])
            return
        if e.button() == Qt.LeftButton and self.obj_drag and self.obj_drag["kind"] != "nudge":
            d, self.obj_drag = self.obj_drag, None
            self.update()
            c = d["cur"]
            link = d["i"] < len(self.obj_tips) and bool(self.obj_tips[d["i"]])
            if (c.topLeft() - d["orig"].topLeft()).manhattanLength() > 2 or abs(c.width() - d["orig"].width()) > 2 \
                    or abs(c.height() - d["orig"].height()) > 2:
                z = self.zoom
                if not link:
                    self.select_object(d["i"])
                self.objectMoved.emit(d["i"], QRectF((c.x() - 20) / z, (c.y() - 20) / z, c.width() / z, c.height() / z))
            elif d["kind"] == "move" and link:
                self.objectClicked.emit(d["i"])                         # bấm liên kết (không kéo) -> mở
            else:
                self.select_object(d["i"])                              # bấm (không kéo): chọn vật
            return
        if self.region_drag and e.button() == Qt.LeftButton:
            self.region_drag = None
            self.update()
            self.regionChanged.emit(self._region_pt())
            return
        if e.button() != Qt.LeftButton or not self.drag:
            return
        a, b = self.to_pt(self.drag[0]), self.to_pt(e.position())
        self.drag = None
        self.update()
        if self.mode == "line":
            self.segmentDrawn.emit(a, b)
        else:
            self.rectDrawn.emit(QRectF(a, b).normalized())


class MailJob(tools_ui.SafeThread):
    """Mở thư mới có file đính kèm qua Simple MAPI (MAPISendMailW + MAPI_DIALOG) ở luồng riêng:
    Outlook mở chậm không làm treo app. Kết quả: mã MAPI (0 = ok, 1 = người dùng huỷ)."""
    finished_with = Signal(int)

    def __init__(self, path, name):
        super().__init__()
        self.path, self.name = path, name

    def run(self):
        import ctypes
        from ctypes import wintypes as W

        class FileDesc(ctypes.Structure):
            _fields_ = [("ulReserved", W.ULONG), ("flFlags", W.ULONG), ("nPosition", W.ULONG),
                        ("lpszPathName", W.LPWSTR), ("lpszFileName", W.LPWSTR), ("lpFileType", ctypes.c_void_p)]

        class Message(ctypes.Structure):
            _fields_ = [("ulReserved", W.ULONG), ("lpszSubject", W.LPWSTR), ("lpszNoteText", W.LPWSTR),
                        ("lpszMessageType", W.LPWSTR), ("lpszDateReceived", W.LPWSTR),
                        ("lpszConversationID", W.LPWSTR), ("flFlags", W.ULONG), ("lpOriginator", ctypes.c_void_p),
                        ("nRecipCount", W.ULONG), ("lpRecips", ctypes.c_void_p),
                        ("nFileCount", W.ULONG), ("lpFiles", ctypes.POINTER(FileDesc))]
        code = -1
        try:
            ctypes.windll.ole32.CoInitializeEx(None, 2)            # STA - một số ứng dụng email cần
            fn = ctypes.windll.mapi32.MAPISendMailW
            fn.argtypes = [ctypes.c_size_t, ctypes.c_size_t, ctypes.POINTER(Message), W.ULONG, W.ULONG]
            fn.restype = W.ULONG
            f = FileDesc(0, 0, 0xFFFFFFFF, self.path, self.name, None)
            m = Message()
            m.lpszSubject = os.path.splitext(self.name)[0]
            m.nFileCount = 1
            m.lpFiles = ctypes.pointer(f)
            MAPI_LOGON_UI, MAPI_DIALOG = 0x1, 0x8     # hiện cửa sổ soạn thư - không tự gửi
            code = fn(0, 0, ctypes.byref(m), MAPI_LOGON_UI | MAPI_DIALOG, 0)
        except Exception:
            code = -1
        self.finished_with.emit(int(code))


class PageScroll(QScrollArea):
    zoomRequested = Signal(int)
    resized = Signal()                        # vùng xem đổi cỡ (kể cả lần đầu hiện ra) -> tính lại "vừa khung"

    def resizeEvent(self, e):
        super().resizeEvent(e)
        if e.size().width() != e.oldSize().width():
            self.resized.emit()

    def wheelEvent(self, e):
        if e.modifiers() & Qt.ControlModifier:
            self.zoomRequested.emit(1 if e.angleDelta().y() > 0 else -1)
            e.accept()
            return
        super().wheelEvent(e)


class PageStrip(QWidget):
    """Mọi trang xếp dọc nối tiếp (cuộn liền mạch như trình đọc PDF). Trang hiện tại là PageCanvas (sửa / chọn /
    vẽ được), đặt đúng vào chỗ của trang đó; các trang còn lại vẽ ảnh tĩnh, chỉ render khi lọt vào màn hình."""
    CACHE = 24

    def __init__(self, canvas):
        super().__init__()
        self.canvas = canvas
        canvas.setParent(self)
        self.doc = None
        self.base = []                    # cỡ trang (pt, đã tính xoay)
        self.dirty = True
        self.zoom = 0
        self.sizes, self.tops = [], []
        self.current = 0
        self.cache = {}                   # chỉ số trang -> QPixmap
        self.setStyleSheet("background:#e5e7eb")

    def sync(self, doc):
        """Đọc lại cỡ các trang khi đổi tài liệu / sau khi thêm, xoá, xoay trang."""
        if doc is not self.doc or self.dirty or len(self.base) != doc.page_count:
            self.doc, self.dirty = doc, False
            self.base = [(p.rect.width, p.rect.height) for p in doc]
            self.cache.clear()
            self.zoom = 0

    def invalidate(self, page=None):
        if page is None:
            self.dirty = True
        else:
            self.cache.pop(page, None)
            if self.doc is not None and 0 <= page < len(self.base):
                r = self.doc[page].rect
                if (r.width, r.height) != self.base[page]:
                    self.dirty = True
        self.update()

    def max_width(self):
        return max((w for w, _h in self.base), default=595)

    def layout_pages(self, zoom):
        """-> True nếu bố cục đổi (đổi zoom / cỡ trang)."""
        dpr = self.devicePixelRatioF()
        sizes = []
        for w, h in self.base:            # đúng cỡ mà PageCanvas.show_page tạo ra
            r = pymupdf.Rect(0, 0, w, h) * pymupdf.Matrix(zoom * dpr, zoom * dpr)
            sizes.append((int(r.irect.width / dpr) + 40, int(r.irect.height / dpr) + 40))
        if sizes == self.sizes and zoom == self.zoom:
            return False
        if zoom != self.zoom:
            self.cache.clear()
        self.zoom, self.sizes = zoom, sizes
        y, self.tops = 0, []
        for _w, h in sizes:
            self.tops.append(y)
            y += h
        self.setFixedSize(max((w for w, _h in sizes), default=10), max(y, 10))
        self.update()
        return True

    def page_rect(self, i):
        w, h = self.sizes[i]
        return QRect((self.width() - w) // 2, self.tops[i], w, h)

    def page_at(self, y):
        import bisect
        return max(0, min(len(self.tops) - 1, bisect.bisect_right(self.tops, y) - 1)) if self.tops else 0

    def place_canvas(self, i):
        self.current = i
        if 0 <= i < len(self.sizes):
            self.canvas.move(self.page_rect(i).topLeft())
            self.canvas.show()
            self.canvas.raise_()
        self.update()

    def _pixmap(self, i):
        pm = self.cache.get(i)
        if pm is None:
            dpr = self.devicePixelRatioF()
            pix = self.doc[i].get_pixmap(matrix=pymupdf.Matrix(self.zoom * dpr, self.zoom * dpr), alpha=False)
            img = pix_to_qimage(pix)
            img.setDevicePixelRatio(dpr)
            pm = QPixmap.fromImage(img)
            self.cache[i] = pm
            while len(self.cache) > self.CACHE:            # bỏ trang xa nhất so với trang đang xem
                far = max(self.cache, key=lambda k: abs(k - self.current))
                self.cache.pop(far)
        return pm

    def paintEvent(self, e):
        if self.doc is None or not self.sizes:
            return
        p = QPainter(self)
        clip = e.rect()
        for i in range(self.page_at(clip.top()), len(self.sizes)):
            r = self.page_rect(i)
            if r.top() > clip.bottom():
                break
            if i == self.current:
                continue
            pm = self._pixmap(i)
            w, h = pm.width() / pm.devicePixelRatio(), pm.height() / pm.devicePixelRatio()
            p.fillRect(QRectF(r.x() + 23, r.y() + 23, w, h), QColor(0, 0, 0, 40))     # bóng đổ (giống PageCanvas)
            p.drawPixmap(QPointF(r.x() + 20, r.y() + 20), pm)
        p.end()


# ===========================================================================
#  Hộp thoại
# ===========================================================================
class ColorButton(QPushButton):
    def __init__(self, color=(0, 0, 0)):
        super().__init__()
        self.setFixedWidth(60)
        self.set_color(color)
        self.clicked.connect(self.pick)

    def set_color(self, rgb):
        self.rgb = tuple(rgb)
        c = QColor.fromRgbF(*self.rgb)
        self.setStyleSheet(f"background:{c.name()}; border:1px solid #94a3b8; border-radius:4px; min-height:22px")

    def pick(self):
        c = QColorDialog.getColor(QColor.fromRgbF(*self.rgb), self, "Chọn màu chữ")
        if c.isValid():
            self.set_color((c.redF(), c.greenF(), c.blueF()))


def _qcolor(rgb):
    return QColor.fromRgbF(*rgb)


def align_icon(kind):
    """Icon căn lề vẽ tay (trái / giữa / phải / đều)."""
    pm = QPixmap(32, 32)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setPen(QPen(QColor("#1e293b"), 2.4))
    for k, w in enumerate((24, 16, 24, 12)):
        if kind == "justify":
            w = 24 if k < 3 else 14
        x = {"left": 4, "center": 16 - w / 2, "right": 28 - w, "justify": 4}[kind]
        y = 7 + k * 6
        p.drawLine(QPointF(x, y), QPointF(x + w, y))
    p.end()
    return QIcon(pm)


def color_icon(letter, rgb, bg=False):
    pm = QPixmap(32, 32)
    pm.fill(Qt.transparent)
    p = QPainter(pm)
    p.setRenderHint(QPainter.Antialiasing)
    f = QFont("Segoe UI", 1, QFont.Bold)
    f.setPixelSize(20 if not bg else 17)
    p.setFont(f)
    p.setPen(QColor("#1e293b"))
    p.drawText(QRectF(0, -2, 32, 26), Qt.AlignCenter, letter)
    p.fillRect(QRectF(4, 25, 24, 5), _qcolor(rgb) if rgb else QColor(0, 0, 0, 0))
    if not rgb:
        p.setPen(QPen(QColor("#94a3b8"), 1))
        p.drawRect(QRectF(4, 25, 24, 5))
    p.end()
    return QIcon(pm)


SIZES = ["8", "9", "10", "10.5", "11", "12", "13", "14", "16", "18", "20", "22", "24", "28", "32", "36", "48", "60", "72"]
ALIGN_FLAGS = {"left": Qt.AlignLeft, "center": Qt.AlignHCenter, "right": Qt.AlignRight, "justify": Qt.AlignJustify}


def _align_name(al):
    v = al.value if hasattr(al, "value") else int(al)
    if v & int(Qt.AlignHCenter.value):
        return "center"
    if v & int(Qt.AlignJustify.value):
        return "justify"
    if v & int(Qt.AlignRight.value):
        return "right"
    return "left"


ZOOM_STEPS = (1.0, 1.25, 1.5, 2.0, 2.5, 3.0, 4.0)


def auto_zoom(size):
    """Mức phóng để chữ trong ô soạn hiện cỡ ~13pt (chữ PDF nhỏ 6-8pt rất khó đọc khi soạn)."""
    want = 13.0 / max(1.0, size)
    return max((z for z in ZOOM_STEPS if z <= want + 0.01), default=1.0)


class RichTextEdit(QWidget):
    """Ô soạn chữ có thanh định dạng giống Word: font, cỡ, đậm/nghiêng/gạch, màu, nền, căn lề."""

    def __init__(self, family="Arial", size=12.0, color=(0, 0, 0)):
        super().__init__()
        self.fonts = pdf_core.all_fonts()
        vn = pdf_core.vn_fonts()
        self.default_family = family if family in self.fonts else "Arial"
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(4)
        bar = QHBoxLayout()
        bar.setSpacing(2)
        lay.addLayout(bar)

        self.font_box = QComboBox()
        self.font_box.setMinimumWidth(190)
        self.font_box.setMaxVisibleItems(18)
        with i18n.paused():                                  # tên font: không dịch
            for fam in vn:
                self.font_box.addItem(fam)
                self.font_box.setItemData(self.font_box.count() - 1, QFont(fam, 11), Qt.FontRole)
        others = [f for f in self.fonts if f not in vn]
        if others:
            self.font_box.insertSeparator(self.font_box.count())
            self.font_box.addItem("Font thiếu dấu tiếng Việt:")
            head = self.font_box.model().item(self.font_box.count() - 1)
            head.setEnabled(False)
            tip = "Font này thiếu một số chữ có dấu tiếng Việt - chữ thiếu sẽ hiện bằng font khác"
            with i18n.paused():
                for fam in others:
                    self.font_box.addItem(fam)
                    self.font_box.setItemData(self.font_box.count() - 1, QFont(fam, 11), Qt.FontRole)
            for k in range(self.font_box.count() - len(others), self.font_box.count()):
                self.font_box.setItemData(k, i18n.tr(tip), Qt.ToolTipRole)
        self.font_box.setToolTip("Font chữ Windows - nhóm đầu có đủ dấu tiếng Việt")
        self.font_box.activated.connect(lambda i: self._merge(families=self.font_box.itemText(i)))
        bar.addWidget(self.font_box)

        self.size_box = QComboBox()
        self.size_box.setEditable(True)
        self.size_box.addItems(SIZES)
        self.size_box.setFixedWidth(92)                 # đủ chỗ cho "10.5" khi gõ tay (QSS có padding)
        self.size_box.setToolTip("Cỡ chữ")
        self.size_box.activated.connect(lambda _: self._apply_size())
        self.size_box.lineEdit().editingFinished.connect(self._apply_size)
        bar.addWidget(self.size_box)
        bar.addSpacing(6)

        def tbtn(text, tip, checkable=True, icon=None, shortcut=None, font=None):
            b = QToolButton()
            b.setText(text)
            if icon:
                b.setIcon(icon)
                b.setIconSize(QSize(18, 18))
            if font:
                b.setFont(font)
            b.setCheckable(checkable)
            b.setToolTip(tip + (f"  ({shortcut})" if shortcut else ""))
            if shortcut:
                b.setShortcut(QKeySequence(shortcut))
            b.setFixedSize(30, 28)
            bar.addWidget(b)
            return b

        fb = QFont("Times New Roman", 12, QFont.Bold)
        fi = QFont("Times New Roman", 12); fi.setItalic(True)
        fu = QFont("Times New Roman", 12); fu.setUnderline(True)
        fs = QFont("Times New Roman", 12); fs.setStrikeOut(True)
        self.b_bold = tbtn("B", "In đậm", font=fb, shortcut="Ctrl+B")
        self.b_italic = tbtn("I", "In nghiêng", font=fi, shortcut="Ctrl+I")
        self.b_under = tbtn("U", "Gạch chân", font=fu, shortcut="Ctrl+U")
        self.b_strike = tbtn("S", "Gạch ngang", font=fs)
        self.b_sup = tbtn("x²", "Chỉ số trên", font=QFont("Segoe UI", 10))
        self.b_sub = tbtn("x₂", "Chỉ số dưới", font=QFont("Segoe UI", 10))
        self.b_bold.toggled.connect(lambda on: self._merge(bold=on))
        self.b_italic.toggled.connect(lambda on: self._merge(italic=on))
        self.b_under.toggled.connect(lambda on: self._merge(underline=on))
        self.b_strike.toggled.connect(lambda on: self._merge(strike=on))
        self.b_sup.toggled.connect(lambda on: self._merge(valign="sup" if on else "normal"))
        self.b_sub.toggled.connect(lambda on: self._merge(valign="sub" if on else "normal"))
        bar.addSpacing(6)

        self.color = tuple(color)
        self.bg = (1, 1, 0)
        self.b_color = tbtn("", "Màu chữ", checkable=False, icon=color_icon("A", self.color))
        self.b_color.setPopupMode(QToolButton.MenuButtonPopup)
        self.b_color.setFixedSize(42, 28)
        self.b_color.clicked.connect(lambda: self._merge(color=self.color))
        m = QMenu(self.b_color)
        for name, rgb in (("Đen", (0, 0, 0)), ("Đỏ", (0.8, 0, 0)), ("Xanh dương", (0, 0.27, 0.75)),
                          ("Xanh lá", (0, 0.5, 0.2)), ("Cam", (0.9, 0.45, 0)), ("Xám", (0.4, 0.4, 0.4)), ("Trắng", (1, 1, 1))):
            m.addAction(color_icon("", rgb), name, lambda c=rgb: self._set_color(c))
        m.addSeparator()
        m.addAction("Màu khác...", lambda: self._pick_color(False))
        self.b_color.setMenu(m)

        self.b_bg = tbtn("", "Tô nền chữ (highlight)", checkable=False, icon=color_icon("ab", self.bg, True))
        self.b_bg.setPopupMode(QToolButton.MenuButtonPopup)
        self.b_bg.setFixedSize(42, 28)
        self.b_bg.clicked.connect(lambda: self._merge(bg=self.bg))
        m = QMenu(self.b_bg)
        for name, rgb in (("Vàng", (1, 1, 0)), ("Xanh lá", (0.6, 1, 0.6)), ("Xanh da trời", (0.6, 0.9, 1)),
                          ("Hồng", (1, 0.75, 0.85)), ("Xám nhạt", (0.85, 0.85, 0.85))):
            m.addAction(color_icon("", rgb), name, lambda c=rgb: self._set_bg(c))
        m.addAction("Không tô nền", lambda: self._merge(bg=None))
        m.addSeparator()
        m.addAction("Màu khác...", lambda: self._pick_color(True))
        self.b_bg.setMenu(m)
        bar.addSpacing(6)

        self.align_btns = {}
        grp = QButtonGroup(self)
        grp.setExclusive(True)
        for key, tip, sc in (("left", "Căn trái", "Ctrl+L"), ("center", "Căn giữa", "Ctrl+E"),
                             ("right", "Căn phải", "Ctrl+R"), ("justify", "Căn đều hai bên", "Ctrl+J")):
            b = tbtn("", tip, icon=align_icon(key), shortcut=sc)
            grp.addButton(b)
            b.clicked.connect(lambda _=False, k=key: self.edit.setAlignment(ALIGN_FLAGS[k]))
            self.align_btns[key] = b
        bar.addStretch()
        self.zoom = 1.0
        self.zoom_box = QComboBox()
        self.zoom_box.setToolTip("Phóng to chữ trong ô soạn cho dễ đọc (không đổi cỡ chữ thật trên PDF)")
        for z in ZOOM_STEPS:
            self.zoom_box.addItem(f"{z * 100:.0f}%", z)
        self.zoom_box.setFixedWidth(100)                # đủ chỗ cho "150%" + mũi tên (QSS có padding)
        self.zoom_box.activated.connect(lambda i: self.set_zoom(self.zoom_box.itemData(i)))
        bar.addWidget(QLabel("Xem:"))
        bar.addWidget(self.zoom_box)

        self.edit = QTextEdit()
        self.edit.setAcceptRichText(True)
        self.edit.setStyleSheet("QTextEdit{background:white; border:1px solid #cbd5e1; border-radius:4px}")
        self.default_size = size
        self.zoom = auto_zoom(size)
        self._set_default_font()
        lay.addWidget(self.edit, 1)
        self._held, self._holding = None, False          # vùng chọn giữ lại sau khi đổi màu (xem _merge)
        self.edit.currentCharFormatChanged.connect(self._sync)
        self.edit.cursorPositionChanged.connect(self._sync_align)
        self.edit.cursorPositionChanged.connect(self._drop_held)
        self.edit.selectionChanged.connect(self._drop_held)
        fmt = QTextCharFormat()
        fmt.setFontFamilies([self.default_family])
        fmt.setFontPointSize(size * self.zoom)
        fmt.setForeground(_qcolor(self.color))
        self.edit.mergeCurrentCharFormat(fmt)
        self._sync_zoom_box()
        self._sync(self.edit.currentCharFormat())
        self._sync_align()

    # ----- thu phóng hiển thị: cỡ chữ trong ô soạn = cỡ thật x zoom -----
    def _pt(self, fmt):
        """Cỡ chữ thật (pt trên PDF) của một định dạng trong ô soạn."""
        return (fmt.fontPointSize() or self.default_size * self.zoom) / self.zoom

    def _set_default_font(self):
        f = QFont(self.default_family)
        f.setPointSizeF(self.default_size * self.zoom)
        self.edit.document().setDefaultFont(f)

    def _sync_zoom_box(self):
        i = min(range(len(ZOOM_STEPS)), key=lambda k: abs(ZOOM_STEPS[k] - self.zoom))
        self.zoom_box.setCurrentIndex(i)

    def set_zoom(self, z):
        if not z or abs(z - self.zoom) < 1e-6:
            return
        k = z / self.zoom
        doc = self.edit.document()
        cur = QTextCursor(doc)
        keep = self.edit.textCursor()
        pos, anchor = keep.position(), keep.anchor()
        b = doc.begin()
        while b.isValid():
            it = b.begin()
            while not it.atEnd():
                fr = it.fragment()
                if fr.isValid():
                    cur.setPosition(fr.position())
                    cur.setPosition(fr.position() + fr.length(), QTextCursor.KeepAnchor)
                    f = QTextCharFormat()
                    f.setFontPointSize((fr.charFormat().fontPointSize() or self.default_size * self.zoom) * k)
                    cur.mergeCharFormat(f)
                it += 1
            bf = b.charFormat()                      # dòng trống
            if not b.text():
                cur.setPosition(b.position())
                f = QTextCharFormat()
                f.setFontPointSize((bf.fontPointSize() or self.default_size * self.zoom) * k)
                cur.mergeBlockCharFormat(f)
            b = b.next()
        self.zoom = z
        self._set_default_font()
        keep.setPosition(anchor)
        keep.setPosition(pos, QTextCursor.KeepAnchor)
        self.edit.setTextCursor(keep)
        cf = self.edit.currentCharFormat()
        if cf.fontPointSize() and not keep.hasSelection():     # định dạng gõ tiếp tại con trỏ (khi có vùng
                                                                # chọn thì chữ đã được đổi cỡ ở vòng lặp trên)
            f = QTextCharFormat()
            f.setFontPointSize(cf.fontPointSize() * k)
            self.edit.mergeCurrentCharFormat(f)
        self._sync_zoom_box()
        self._sync(self.edit.currentCharFormat())

    # ----- áp dụng định dạng -----
    def _merge(self, families=None, size=None, bold=None, italic=None, underline=None, strike=None,
               color=None, bg=False, valign=None):
        if getattr(self, "_syncing", False):
            return
        fmt = QTextCharFormat()
        if families:
            fmt.setFontFamilies([families])
        if size:
            fmt.setFontPointSize(size * self.zoom)
        if bold is not None:
            fmt.setFontWeight(QFont.Bold if bold else QFont.Normal)
        if italic is not None:
            fmt.setFontItalic(italic)
        if underline is not None:
            fmt.setFontUnderline(underline)
        if strike is not None:
            fmt.setFontStrikeOut(strike)
        if color is not None:
            fmt.setForeground(_qcolor(color))
        if bg is not False:
            fmt.setBackground(_qcolor(bg) if bg else QBrush(Qt.NoBrush))
        if valign:
            fmt.setVerticalAlignment({"sup": QTextCharFormat.AlignSuperScript, "sub": QTextCharFormat.AlignSubScript,
                                      "normal": QTextCharFormat.AlignNormal}[valign])
        cur = self.edit.textCursor()
        if not cur.hasSelection() and self._held is not None:      # vùng vừa đổi màu (đã bỏ bôi đen để xem)
            cur = QTextCursor(self._held)
            cur.mergeCharFormat(fmt)
        else:
            self.edit.mergeCurrentCharFormat(fmt)
        if (color is not None or bg is not False) and cur.hasSelection():
            # chữ đang bôi đen luôn vẽ bằng màu chọn -> bỏ bôi đen để thấy ngay màu mới, nhưng vẫn nhớ vùng
            # chọn: bấm tiếp định dạng khác (đậm, cỡ, màu...) vẫn áp cho đúng đoạn đó tới khi di chuyển con trỏ
            self._held = QTextCursor(cur)
            nc = QTextCursor(cur)
            nc.setPosition(cur.selectionEnd())
            self._holding = True
            try:
                self.edit.setTextCursor(nc)
            finally:
                self._holding = False
        self.edit.setFocus()

    def _drop_held(self):
        if not self._holding:
            self._held = None

    def _apply_size(self):
        try:
            v = float(self.size_box.currentText().replace(",", "."))
        except ValueError:
            return
        if 1 <= v <= 400:
            self._merge(size=v)

    def _set_color(self, rgb):
        self.color = rgb
        self.b_color.setIcon(color_icon("A", rgb))
        self._merge(color=rgb)

    def _set_bg(self, rgb):
        self.bg = rgb
        self.b_bg.setIcon(color_icon("ab", rgb, True))
        self._merge(bg=rgb)

    def _pick_color(self, background):
        c = QColorDialog.getColor(_qcolor(self.bg if background else self.color), self,
                                  "Chọn màu nền" if background else "Chọn màu chữ")
        if c.isValid():
            (self._set_bg if background else self._set_color)((c.redF(), c.greenF(), c.blueF()))

    def _sync(self, fmt):
        self._syncing = True
        try:
            fams = fmt.fontFamilies() or [self.default_family]
            i = self.font_box.findText(fams[0])
            if i >= 0:
                self.font_box.setCurrentIndex(i)
            else:
                self.font_box.setEditText(fams[0])
            self.size_box.setEditText(f"{round(self._pt(fmt), 1):g}")
            self.b_bold.setChecked(fmt.fontWeight() >= 600)
            self.b_italic.setChecked(fmt.fontItalic())
            self.b_under.setChecked(fmt.fontUnderline())
            self.b_strike.setChecked(fmt.fontStrikeOut())
            self.b_sup.setChecked(fmt.verticalAlignment() == QTextCharFormat.AlignSuperScript)
            self.b_sub.setChecked(fmt.verticalAlignment() == QTextCharFormat.AlignSubScript)
        finally:
            self._syncing = False

    def _sync_align(self):
        self.align_btns[_align_name(self.edit.alignment())].setChecked(True)

    # ----- nạp / xuất -----
    def set_runs(self, runs, align="left"):
        """runs: [(text, family, size, bold, italic, rgb)] -> nạp vào ô soạn."""
        parts = []
        if runs:
            self.zoom = auto_zoom(max(r[2] for r in runs))
            self._set_default_font()
            self._sync_zoom_box()
        for text, fam, size, bold, italic, rgb in runs:
            st = (f"font-family:'{fam}'; font-size:{size * self.zoom:g}pt; font-weight:{700 if bold else 400}; "
                  f"font-style:{'italic' if italic else 'normal'}; color:{_qcolor(rgb).name()}")
            parts.append(f'<span style="{st}">{html.escape(text)}</span>')
        self.edit.setHtml(f'<p align="{align}" style="margin:0">{"".join(parts)}</p>')
        c = self.edit.textCursor()
        c.select(QTextCursor.Document)
        self.edit.setTextCursor(c)

    def is_empty(self):
        return not self.edit.toPlainText().strip()

    def first_align(self):
        return _align_name(self.edit.document().begin().blockFormat().alignment())

    def last_format(self):
        """Định dạng tại cuối nội dung (để nhớ cho lần thêm chữ sau)."""
        c = self.edit.textCursor()
        c.movePosition(QTextCursor.End)
        f = c.charFormat()
        fams = f.fontFamilies() or [self.default_family]
        fg = f.foreground().color() if f.foreground().style() != Qt.NoBrush else QColor("black")
        return dict(family=self._family(fams[0]), size=round(self._pt(f), 2),
                    color=(fg.redF(), fg.greenF(), fg.blueF()))

    def _family(self, name):
        return name if name in self.fonts else pdf_core.guess_family(name)

    def to_html(self):
        """Chuyển nội dung sang HTML cho MuPDF. Trả về (html, [font đã dùng])."""
        fams = set()
        out = []
        b = self.edit.document().begin()
        while b.isValid():
            text = b.text()
            align = _align_name(b.blockFormat().alignment())
            spans = []
            for fr in b.textFormats():
                seg = text[fr.start:fr.start + fr.length]
                if not seg:
                    continue
                f = fr.format
                fam = self._family((f.fontFamilies() or [self.default_family])[0])
                fams.add(fam)
                size = round(self._pt(f), 2)
                st = [f"font-family:'{fam}'"]
                va = f.verticalAlignment()
                if va in (QTextCharFormat.AlignSuperScript, QTextCharFormat.AlignSubScript):
                    st.append(f"font-size:{size * 0.65:.1f}pt")
                    st.append("vertical-align:" + ("super" if va == QTextCharFormat.AlignSuperScript else "sub"))
                else:
                    st.append(f"font-size:{size:g}pt")
                if f.fontWeight() >= 600:
                    st.append("font-weight:bold")
                if f.fontItalic():
                    st.append("font-style:italic")
                deco = [d for d, on in (("underline", f.fontUnderline()), ("line-through", f.fontStrikeOut())) if on]
                if deco:
                    st.append("text-decoration:" + " ".join(deco))
                if f.foreground().style() != Qt.NoBrush:
                    st.append("color:" + f.foreground().color().name())
                if f.background().style() != Qt.NoBrush:
                    st.append("background-color:" + f.background().color().name())
                body = html.escape(seg).replace(" ", "<br>").replace("\t", "    ")
                spans.append(f'<span style="{";".join(st)}">{body}</span>')
            if not spans:                       # dòng trống: giữ chiều cao theo cỡ chữ
                size = round(self._pt(b.charFormat()), 2)
                spans.append(f'<span style="font-size:{size:g}pt">&nbsp;</span>')
            out.append(f'<p style="text-align:{align};white-space:pre-wrap">{"".join(spans)}</p>')
            b = b.next()
        if not fams:
            fams.add(self.default_family)
        return "".join(out), sorted(fams)


class RichTextDialog(QDialog):
    last = dict(family="Arial", size=12.0, color=(0, 0, 0))

    @staticmethod
    def line_runs(line):
        """Các đoạn cùng kiểu chữ của 1 dòng PDF: [(chữ, font Windows, cỡ, đậm, nghiêng, màu)]."""
        runs = []
        for ch, (bb, sp) in zip(line.text, line.chars):
            if runs and runs[-1][0] is sp:
                runs[-1][1].append(ch)
            else:
                runs.append((sp, [ch]))
        return [("".join(chs), pdf_core.guess_family(sp["font"], sp["flags"]), round(sp["size"], 1),
                 pdf_core.is_bold(sp["font"], sp["flags"]), pdf_core.is_italic(sp["font"], sp["flags"]),
                 int_to_rgb(sp["color"])) for sp, chs in runs]

    def __init__(self, parent, title, line=None, runs=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(760, 360 if runs else 300 if line else 360)
        self.delete = False
        lay = QVBoxLayout(self)
        QApplication.setOverrideCursor(Qt.WaitCursor)      # lần đầu: quét font Windows
        try:
            self.rich = RichTextEdit(**self.last)
        finally:
            QApplication.restoreOverrideCursor()
        lay.addWidget(self.rich, 1)
        self.fit = None
        if line is not None:
            self.rich.set_runs(runs or self.line_runs(line))
            self.fit = QCheckBox("Co chữ cho vừa độ rộng dòng cũ nếu nội dung mới dài hơn")
            if runs:                                    # khối chữ: chữ tự xuống dòng trong bề ngang khối
                self.fit.hide()
            lay.addWidget(self.fit)
            info = QLabel(f"Font gốc trong PDF: {line.font}  →  dùng font Windows tương ứng. "
                          "Bôi đen một phần chữ để đổi định dạng riêng phần đó.")
            info.setWordWrap(True)
            info.setStyleSheet("color:#64748b")
            lay.addWidget(info)
        else:
            self.rich.edit.setPlaceholderText("Nhập nội dung... (Enter để xuống dòng, bôi đen để định dạng)")
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Áp dụng" if line is not None else "Thêm")
        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")
        if line is not None:
            d = btns.addButton("Xoá khối chữ này" if runs else "Xoá dòng này", QDialogButtonBox.DestructiveRole)
            d.clicked.connect(self._delete)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)
        self.rich.edit.setFocus()

    def _delete(self):
        self.delete = True
        self.accept()

    def accept(self):
        if not self.delete and not self.rich.is_empty():
            RichTextDialog.last = self.rich.last_format()
        super().accept()


class LinkDialog(QDialog):
    """Địa chỉ (web / email / trang) + chữ hiển thị soạn bằng bộ định dạng đầy đủ (font, cỡ, màu, đậm, gạch chân...)."""
    last = dict(family="Arial", size=11.0, color=(0.02, 0.39, 0.76))

    def __init__(self, parent, page_count, need_text):
        super().__init__(parent)
        self.setWindowTitle("Chèn liên kết")
        self.resize(720, 400)
        v = QVBoxLayout(self)
        lay = QFormLayout()
        self.r_url = QRadioButton("Địa chỉ web / email")
        self.r_page = QRadioButton("Trang trong file này")
        self.r_url.setChecked(True)
        self.url = QLineEdit()
        self.url.setPlaceholderText("vd: https://example.com  hoặc  ten@congty.com")
        self.page = QSpinBox()
        self.page.setRange(1, max(1, page_count))
        self.page.valueChanged.connect(lambda _: self.r_page.setChecked(True))
        self.url.textEdited.connect(lambda _: self.r_url.setChecked(True))
        lay.addRow(self.r_url, self.url)
        lay.addRow(self.r_page, self.page)
        v.addLayout(lay)
        cap = QLabel("Chữ hiển thị:" if need_text else
                     "Chữ hiển thị (để trống = biến vùng đã kéo trên chữ / ảnh có sẵn thành liên kết):")
        v.addWidget(cap)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self.rich = RichTextEdit(**self.last)
        finally:
            QApplication.restoreOverrideCursor()
        self.rich._merge(underline=True)               # mặc định kiểu liên kết: xanh, gạch chân
        self.rich.edit.setPlaceholderText("vd: Xem chi tiết tại đây")
        v.addWidget(self.rich, 1)
        if need_text:
            hint = QLabel("Bạn chỉ bấm một điểm nên cần nhập chữ hiển thị (hoặc kéo vùng trên chữ có sẵn).")
            hint.setWordWrap(True)
            hint.setStyleSheet("color:#b45309")
            v.addWidget(hint)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Chèn")
        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")
        btns.accepted.connect(self._ok)
        btns.rejected.connect(self.reject)
        v.addWidget(btns)
        self.need_text = need_text
        self.url.setFocus()

    @staticmethod
    def normalize(u):
        if "://" not in u and not u.lower().startswith("mailto:"):
            u = ("mailto:" + u) if "@" in u and "/" not in u else ("https://" + u)
        return u

    def target(self):
        """('uri', địa chỉ) hoặc ('page', chỉ số trang)."""
        if self.r_page.isChecked():
            return "page", self.page.value() - 1
        return "uri", self.normalize(self.url.text().strip())

    def html(self):
        """(html, fonts) của chữ hiển thị, hoặc (None, []) nếu để trống."""
        if self.rich.is_empty():
            return None, []
        return self.rich.to_html()

    def _ok(self):
        if self.r_url.isChecked() and not self.url.text().strip():
            QMessageBox.warning(self, self.windowTitle(), "Hãy nhập địa chỉ liên kết.")
            return
        if self.need_text and self.rich.is_empty():
            QMessageBox.warning(self, self.windowTitle(), "Hãy nhập chữ hiển thị cho liên kết.")
            return
        if not self.rich.is_empty():
            LinkDialog.last = self.rich.last_format()
        self.accept()


class ShapeDialog(QDialog):
    last = dict(stroke=(0.86, 0.15, 0.15), width=1.5, fill=None)

    def __init__(self, parent, kind):
        super().__init__(parent)
        self.setWindowTitle({"rect": "Hình chữ nhật", "ellipse": "Hình elip", "line": "Đường thẳng",
                             "arrow": "Mũi tên", "polygon": "Đa giác", "polyline": "Đường gấp khúc",
                             "cloud": "Đám mây", "pen": "Kiểu bút vẽ"}[kind])
        if kind in ("polygon", "cloud"):
            kind = "ellipse"                              # có tô nền như hình kín (không có nút che trắng riêng)
        lay = QFormLayout(self)
        v = self.last
        self.stroke = ColorButton(v["stroke"])
        self.width = QDoubleSpinBox(); self.width.setRange(0, 20); self.width.setSingleStep(0.5); self.width.setValue(v["width"])
        lay.addRow("Màu viền:", self.stroke)
        lay.addRow("Độ dày viền:", self.width)
        self.use_fill = QCheckBox("Tô nền")
        self.fill = ColorButton(v["fill"] or (1, 1, 1))
        self.use_fill.setChecked(bool(v["fill"]))
        if kind in ("rect", "ellipse"):
            row = QHBoxLayout(); row.addWidget(self.use_fill); row.addWidget(self.fill); row.addStretch()
            lay.addRow("", row)
            white = QPushButton("Che trắng (xoá vùng)")
            white.setToolTip("Nền trắng, không viền - dùng để che nội dung bên dưới")
            white.clicked.connect(self._whiteout)
            lay.addRow("", white)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Vẽ")
        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addRow(btns)

    def _whiteout(self):
        self.use_fill.setChecked(True)
        self.fill.set_color((1, 1, 1))
        self.width.setValue(0)
        self.accept()

    def values(self):
        v = dict(stroke=self.stroke.rgb, width=self.width.value(), fill=self.fill.rgb if self.use_fill.isChecked() else None)
        ShapeDialog.last = v
        return v


class TextBoxDialog(QDialog):
    """Hộp chữ / máy chữ / chú thích có mũi tên (FreeText) - chữ thường, cỡ, màu chữ, nền, viền."""
    last = dict(size=11.0, color=(0, 0, 0), fill=None, border=1.0)
    FILLS = [("Không nền", None), ("Vàng nhạt", (1, 0.98, 0.75)), ("Trắng", (1, 1, 1)), ("Xanh nhạt", (0.88, 0.95, 1))]

    def __init__(self, parent, title, border_default=None, text=""):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(460, 320)
        v = QVBoxLayout(self)
        self.text = QPlainTextEdit(text)
        self.text.setPlaceholderText("Nhập nội dung...")
        v.addWidget(self.text, 1)
        form = QFormLayout()
        L = self.last
        self.size = QDoubleSpinBox(); self.size.setRange(5, 72); self.size.setValue(L["size"])
        self.color = ColorButton(L["color"])
        self.fill = QComboBox()
        for name, _c in self.FILLS:
            self.fill.addItem(name)
        self.fill.setCurrentIndex(next((k for k, (_n, c) in enumerate(self.FILLS) if c == L["fill"]), 0))
        self.border = QDoubleSpinBox(); self.border.setRange(0, 6); self.border.setSingleStep(0.5)
        self.border.setValue(L["border"] if border_default is None else border_default)
        form.addRow("Cỡ chữ:", self.size)
        form.addRow("Màu chữ:", self.color)
        form.addRow("Nền:", self.fill)
        form.addRow("Viền (0 = không viền):", self.border)
        v.addLayout(form)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Thêm")
        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")
        btns.accepted.connect(self._ok)
        btns.rejected.connect(self.reject)
        v.addWidget(btns)
        self.text.setFocus()

    def _ok(self):
        if not self.text.toPlainText().strip():
            QMessageBox.warning(self, self.windowTitle(), "Hãy nhập nội dung.")
            return
        TextBoxDialog.last = dict(size=self.size.value(), color=self.color.rgb,
                                  fill=self.FILLS[self.fill.currentIndex()][1], border=self.border.value())
        self.accept()

    def values(self):
        return dict(text=self.text.toPlainText().rstrip(), fontsize=self.size.value(), color=self.color.rgb,
                    fill=self.FILLS[self.fill.currentIndex()][1], border=self.border.value())


class CommentsDialog(QDialog):
    """Danh sách mọi ghi chú / chú thích trong file: bấm đúp để tới trang, xoá, xuất bản tổng hợp."""
    COLS = ["Trang", "Loại", "Người tạo", "Ngày", "Nội dung"]
    KINDS = {"Text": "Ghi chú", "FreeText": "Hộp chữ", "Highlight": "Tô sáng", "Underline": "Gạch chân",
             "StrikeOut": "Gạch ngang", "Squiggly": "Gạch sóng", "Square": "Hình chữ nhật", "Circle": "Hình elip",
             "Line": "Đường / mũi tên", "Polygon": "Đa giác / đám mây", "PolyLine": "Đường gấp khúc",
             "Ink": "Nét vẽ tay", "Stamp": "Con dấu", "FileAttachment": "File đính kèm", "Caret": "Chèn chữ"}

    def __init__(self, win):
        super().__init__(win)
        self.win = win
        self.setWindowTitle("Danh sách ghi chú")
        self.resize(860, 480)
        v = QVBoxLayout(self)
        self.table = QTableWidget(0, len(self.COLS))
        self.table.setHorizontalHeaderLabels(self.COLS)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.cellDoubleClicked.connect(self._goto)
        v.addWidget(self.table, 1)
        row = QHBoxLayout()
        b_del = QPushButton("Xoá mục đã chọn")
        b_del.clicked.connect(self._delete)
        b_exp = QPushButton("Xuất bản tổng hợp...")
        b_exp.setToolTip("Lưu danh sách ghi chú ra file HTML (mở bằng trình duyệt / Word) hoặc CSV (Excel)")
        b_exp.clicked.connect(self._export)
        row.addWidget(b_del)
        row.addWidget(b_exp)
        row.addStretch()
        hint = QLabel("Bấm đúp một dòng để tới trang đó")
        hint.setStyleSheet("color:#64748b")
        row.addWidget(hint)
        b_close = QPushButton("Đóng")
        b_close.clicked.connect(self.accept)
        row.addWidget(b_close)
        v.addLayout(row)
        self.reload()

    def reload(self):
        self.items = self.win.pdf.list_annots()
        self.table.setRowCount(len(self.items))
        rows = [[str(a["page"] + 1), i18n.tr(self.KINDS.get(a["kind"], a["kind"])), a["author"], a["date"],
                 a["text"].replace("\n", " ")] for a in self.items]
        with i18n.paused():                              # nội dung ghi chú: không dịch
            for r, vals in enumerate(rows):
                for c, val in enumerate(vals):
                    self.table.setItem(r, c, QTableWidgetItem(val))
        self.table.resizeColumnsToContents()
        self.table.horizontalHeader().setStretchLastSection(True)

    def _goto(self, row, _col=0):
        if 0 <= row < len(self.items):
            self.win.goto(self.items[row]["page"])

    def _delete(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        if not rows:
            return
        for r in rows:
            a = self.items[r]
            self.win.pdf.delete_annot(a["page"], a["xref"])
        self.win.refresh_thumbs(sorted({self.items[r]["page"] for r in rows}))
        self.win.render_page()
        self.reload()

    def _export(self):
        base = os.path.splitext(self.win.pdf.path or "tai_lieu")[0] + "_ghi_chu"
        path, _ = QFileDialog.getSaveFileName(self, "Xuất danh sách ghi chú", base + ".html",
                                              "HTML (*.html);;CSV - Excel (*.csv)")
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                import csv
                with open(path, "w", newline="", encoding="utf-8-sig") as f:
                    w = csv.writer(f)
                    w.writerow([i18n.tr(c) for c in self.COLS])
                    for a in self.items:
                        w.writerow([a["page"] + 1, i18n.tr(self.KINDS.get(a["kind"], a["kind"])), a["author"],
                                    a["date"], a["text"]])
            else:
                name = html.escape(os.path.basename(self.win.pdf.path or ""))
                rows = "".join(
                    f"<tr><td>{a['page'] + 1}</td><td>{html.escape(i18n.tr(self.KINDS.get(a['kind'], a['kind'])))}</td>"
                    f"<td>{html.escape(a['author'])}</td><td>{html.escape(a['date'])}</td>"
                    f"<td>{html.escape(a['text']).replace(chr(10), '<br>')}</td></tr>" for a in self.items)
                head = "".join(f"<th>{html.escape(i18n.tr(c))}</th>" for c in self.COLS)
                with open(path, "w", encoding="utf-8") as f:
                    f.write(f"<!doctype html><meta charset='utf-8'><title>{name}</title>"
                            "<style>body{font-family:Segoe UI,Arial;margin:24px}table{border-collapse:collapse;width:100%}"
                            "td,th{border:1px solid #cbd5e1;padding:6px 8px;vertical-align:top;text-align:left}"
                            "th{background:#f1f5f9}</style>"
                            f"<h2>{html.escape(i18n.tr('Danh sách ghi chú'))} - {name}</h2>"
                            f"<table><tr>{head}</tr>{rows}</table>")
        except OSError as e:
            QMessageBox.warning(self, self.windowTitle(), str(e))
            return
        self.win.status.showMessage(f"Đã xuất {path}", 6000)


TSA_URLS = ["http://timestamp.digicert.com", "http://timestamp.sectigo.com", "http://time.certum.pl"]


class FieldDialog(QDialog):
    """Tạo ô form: loại, tên, giá trị mặc định, danh sách lựa chọn."""
    KINDS = [("text", "Ô chữ 1 dòng"), ("multiline", "Ô chữ nhiều dòng"), ("check", "Ô đánh dấu (checkbox)"),
             ("combo", "Danh sách thả xuống"), ("list", "Danh sách chọn")]

    def __init__(self, parent, existing=()):
        super().__init__(parent)
        self.setWindowTitle("Tạo ô form")
        self.resize(420, 360)
        f = QFormLayout(self)
        self.kind = QComboBox()
        for k, lbl in self.KINDS:
            self.kind.addItem(lbl, k)
        self.name = QLineEdit()
        n = 1
        while f"O_{n}" in existing:
            n += 1
        self.name.setText(f"O_{n}")
        self.value = QLineEdit()
        self.value.setPlaceholderText("để trống nếu không cần")
        self.checked = QCheckBox("Đánh dấu sẵn")
        self.choices = QPlainTextEdit()
        self.choices.setPlaceholderText("Mỗi dòng một lựa chọn")
        self.choices.setMaximumHeight(110)
        self.size = QDoubleSpinBox()
        self.size.setRange(0, 72)
        self.size.setValue(11)
        self.size.setSuffix(" pt")
        self.size.setToolTip("0 = tự co chữ vừa ô")
        self.required = QCheckBox("Bắt buộc điền")
        f.addRow("Loại ô:", self.kind)
        f.addRow("Tên ô:", self.name)
        f.addRow("Giá trị mặc định:", self.value)
        f.addRow("", self.checked)
        f.addRow("Các lựa chọn:", self.choices)
        f.addRow("Cỡ chữ:", self.size)
        f.addRow("", self.required)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        f.addRow(bb)
        self.form = f
        self.kind.currentIndexChanged.connect(self._sync)
        self._sync()

    def _sync(self):
        k = self.kind.currentData()
        self.form.setRowVisible(self.value, k not in ("check",))
        self.form.setRowVisible(self.checked, k == "check")
        self.form.setRowVisible(self.choices, k in ("combo", "list"))
        self.form.setRowVisible(self.size, k != "check")

    def _ok(self):
        if self.kind.currentData() in ("combo", "list") and not self._choices():
            QMessageBox.information(self, "Tạo ô form", "Hãy nhập ít nhất một lựa chọn.")
            return
        self.accept()

    def _choices(self):
        return [x.strip() for x in self.choices.toPlainText().splitlines() if x.strip()]

    def values(self):
        k = self.kind.currentData()
        return dict(kind=k, name=re.sub(r"[^\w.-]", "_", self.name.text().strip()),
                    value=self.checked.isChecked() if k == "check" else self.value.text(),
                    choices=self._choices(), fontsize=self.size.value(), required=self.required.isChecked())


class SignaturePad(QWidget):
    """Ô vẽ chữ ký bằng chuột / bút cảm ứng."""

    def __init__(self):
        super().__init__()
        self.setMinimumSize(520, 200)
        self.setCursor(Qt.CrossCursor)
        self.strokes = []
        self.color = QColor("#0b2a8a")
        self.pen_w = 2.6
        self.image = None                  # ảnh nạp từ file / chữ ký đã lưu (QImage)

    def clear(self):
        self.strokes, self.image = [], None
        self.update()

    def is_empty(self):
        return not self.strokes and self.image is None

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.image = None
            self.strokes.append([e.position()])
            self.update()

    def mouseMoveEvent(self, e):
        if self.strokes and e.buttons() & Qt.LeftButton:
            self.strokes[-1].append(e.position())
            self.update()

    def _path(self, k=1.0, dx=0, dy=0):
        path = QPainterPath()
        for st in self.strokes:
            pts = [QPointF((p.x() - dx) * k, (p.y() - dy) * k) for p in st]
            path.moveTo(pts[0])
            if len(pts) == 1:
                path.lineTo(pts[0] + QPointF(0.5, 0.5))
            for a, b in zip(pts, pts[1:]):                # làm mượt nét bằng đường cong qua trung điểm
                path.quadTo(a, (a + b) / 2)
            if len(pts) > 1:
                path.lineTo(pts[-1])
        return path

    def paintEvent(self, e):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        p.fillRect(self.rect(), QColor("#ffffff"))
        p.setPen(QPen(QColor("#cbd5e1"), 1, Qt.DashLine))
        y = self.height() * 0.72
        p.drawLine(QPointF(24, y), QPointF(self.rect().width() - 24, y))
        if self.image is not None:
            img = self.image.scaled(self.size() * 0.9, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            p.drawImage(QPointF((self.rect().width() - img.width()) / 2, (self.height() - img.height()) / 2), img)
        elif not self.strokes:
            p.setPen(QColor("#94a3b8"))
            p.drawText(self.rect(), Qt.AlignCenter, "Ký vào đây bằng chuột / bút cảm ứng")
        pen = QPen(self.color, self.pen_w, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        p.setPen(pen)
        p.setBrush(Qt.NoBrush)
        p.drawPath(self._path())

    def to_png(self):
        """PNG nền trong suốt, cắt sát nét ký, độ phân giải x4. Trả về (bytes, tỉ lệ rộng/cao)."""
        if self.image is not None:
            img = self.image
        else:
            br = self._path().boundingRect().adjusted(-self.pen_w * 2, -self.pen_w * 2, self.pen_w * 2, self.pen_w * 2)
            k = 4.0
            img = QImage(int(br.width() * k) + 1, int(br.height() * k) + 1, QImage.Format_ARGB32)
            img.fill(Qt.transparent)
            p = QPainter(img)
            p.setRenderHint(QPainter.Antialiasing)
            p.setPen(QPen(self.color, self.pen_w * k, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin))
            p.drawPath(self._path(k, br.x(), br.y()))
            p.end()
        buf = QBuffer()
        buf.open(QBuffer.WriteOnly)
        img.save(buf, "PNG")
        return bytes(buf.data()), img.width() / max(1, img.height())


class HandSignDialog(QDialog):
    """Vẽ chữ ký tay (hoặc nạp ảnh chữ ký), có thể lưu lại để dùng lần sau."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Chữ ký tay")
        self.settings = QSettings(installer.ORG, "PDFToolkit")
        self.png, self.ratio = None, 2.5
        v = QVBoxLayout(self)
        self.pad = SignaturePad()
        self.pad.setStyleSheet("border:1px solid #cbd5e1")
        v.addWidget(self.pad, 1)
        row = QHBoxLayout()
        row.addWidget(QLabel("Màu mực:"))
        self.color = QComboBox()
        for name, c in (("Xanh mực", "#0b2a8a"), ("Đen", "#111111"), ("Đỏ", "#b91c1c")):
            self.color.addItem(name, c)
        self.color.currentIndexChanged.connect(self._style)
        row.addWidget(self.color)
        row.addWidget(QLabel("Nét:"))
        self.pen = QDoubleSpinBox()
        self.pen.setRange(1, 8)
        self.pen.setSingleStep(0.5)
        self.pen.setValue(2.5)
        self.pen.valueChanged.connect(self._style)
        row.addWidget(self.pen)
        row.addStretch()
        b = QPushButton("Xoá")
        b.clicked.connect(self.pad.clear)
        row.addWidget(b)
        b = QPushButton("Dùng ảnh chữ ký...")
        b.clicked.connect(self._load_image)
        row.addWidget(b)
        self.b_saved = QPushButton("Chữ ký đã lưu")
        self.b_saved.clicked.connect(self._load_saved)
        self.b_saved.setEnabled(bool(self.settings.value("handsign/png", "")))
        row.addWidget(self.b_saved)
        v.addLayout(row)
        self.remember = QCheckBox("Lưu chữ ký này để dùng lại lần sau (chỉ lưu trên máy này)")
        self.remember.setChecked(True)
        v.addWidget(self.remember)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText("Chèn chữ ký")
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self._style()
        if self.settings.value("handsign/png", ""):
            self._load_saved()

    def _style(self):
        self.pad.color = QColor(self.color.currentData())
        self.pad.pen_w = self.pen.value()
        self.pad.update()

    def _load_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "Chọn ảnh chữ ký", "", "Ảnh (*.png *.jpg *.jpeg *.bmp *.webp)")
        if path:
            img = QImage(path)
            if not img.isNull():
                self.pad.strokes = []
                self.pad.image = img
                self.pad.update()

    def _load_saved(self):
        import base64
        data = self.settings.value("handsign/png", "")
        if data:
            img = QImage.fromData(base64.b64decode(data))
            if not img.isNull():
                self.pad.strokes = []
                self.pad.image = img
                self.pad.update()

    def _ok(self):
        if self.pad.is_empty():
            QMessageBox.information(self, "Chữ ký tay", "Hãy ký vào ô trắng hoặc chọn ảnh chữ ký.")
            return
        self.png, self.ratio = self.pad.to_png()
        if self.remember.isChecked():
            import base64
            self.settings.setValue("handsign/png", base64.b64encode(self.png).decode())
        self.accept()


STAMP_PRESETS = ["ĐÃ DUYỆT", "ĐÃ KIỂM TRA", "ĐÃ NHẬN", "ĐÃ THANH TOÁN", "BẢN SAO", "MẬT", "KHẨN", "HUỶ",
                 "APPROVED", "REVIEWED", "RECEIVED", "PAID", "COPY", "CONFIDENTIAL", "DRAFT", "VOID"]
STAMP_COLORS = (("Đỏ", "#c81e1e"), ("Xanh dương", "#1d4ed8"), ("Xanh lá", "#15803d"), ("Đen", "#1f2937"))


def render_stamp(text, sub, color, shape, tilt):
    """Vẽ con dấu thành ảnh PNG nền trong suốt. shape: rect | round. Trả về (png bytes, tỉ lệ rộng/cao)."""
    from PySide6.QtCore import QBuffer, QIODevice
    from PySide6.QtGui import QFontMetricsF, QTransform
    col = QColor(color)
    big = QFont("Arial", 64, QFont.Bold)
    small = QFont("Arial", 26, QFont.Bold)
    fm, fs = QFontMetricsF(big), QFontMetricsF(small)
    tw = fm.horizontalAdvance(text)
    sw = fs.horizontalAdvance(sub) if sub else 0
    if shape == "round":
        d = max(tw * 0.75, sw * 0.9, 240) + 120
        W, H = d, d
    else:
        W = max(tw, sw) + 90
        H = fm.height() + (fs.height() + 6 if sub else 0) + 50
    img = QImage(int(W) + 16, int(H) + 16, QImage.Format_ARGB32_Premultiplied)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)
    p.translate(8, 8)
    p.setBrush(Qt.NoBrush)
    if shape == "round":
        p.setPen(QPen(col, 10))
        p.drawEllipse(QRectF(5, 5, W - 10, H - 10))
        p.setPen(QPen(col, 3))
        p.drawEllipse(QRectF(24, 24, W - 48, H - 48))
        # chữ chính co lại cho vừa lòng con dấu
        k = min(1.0, (W - 110) / max(tw, 1))
        f2 = QFont(big)
        f2.setPointSizeF(64 * k)
        p.setFont(f2)
        p.setPen(col)
        hh = QFontMetricsF(f2).height()
        p.drawText(QRectF(0, H / 2 - hh / 2 - (fs.height() / 2 if sub else 0), W, hh), Qt.AlignCenter, text)
        if sub:
            p.setFont(small)
            p.drawText(QRectF(0, H / 2 + hh / 2 - fs.height() / 2 + 4, W, fs.height()), Qt.AlignCenter, sub)
    else:
        p.setPen(QPen(col, 8))
        p.drawRoundedRect(QRectF(4, 4, W - 8, H - 8), 18, 18)
        p.setPen(QPen(col, 2.5))
        p.drawRoundedRect(QRectF(16, 16, W - 32, H - 32), 10, 10)
        p.setPen(col)
        p.setFont(big)
        p.drawText(QRectF(0, 25, W, fm.height()), Qt.AlignCenter, text)
        if sub:
            p.setFont(small)
            p.drawText(QRectF(0, 25 + fm.height(), W, fs.height()), Qt.AlignCenter, sub)
    p.end()
    if tilt:
        img = img.transformed(QTransform().rotate(-tilt), Qt.SmoothTransformation)
    buf = QBuffer()
    buf.open(QIODevice.WriteOnly)
    img.save(buf, "PNG")
    return bytes(buf.data()), img.width() / max(1, img.height())


def stamps_dir():
    """Thư viện con dấu đã tách nền (chỉ trên máy này): %APPDATA%\\TuanPA\\PDFToolkit\\stamps"""
    d = os.path.join(installer._appdata(), "stamps")
    os.makedirs(d, exist_ok=True)
    return d


def qimage_rgb(img):
    """QImage -> numpy RGB (h, w, 3)."""
    import numpy as np
    img = img.convertToFormat(QImage.Format_RGB888)
    w, h = img.width(), img.height()
    a = np.frombuffer(img.constBits(), np.uint8).reshape(h, img.bytesPerLine())[:, :w * 3]
    return a.reshape(h, w, 3).copy()


def checker_pixmap(png, w, h):
    """Ảnh PNG trong suốt vẽ trên nền ô caro (thấy rõ phần đã tách nền)."""
    src = QPixmap()
    src.loadFromData(png)
    src = src.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    pm = QPixmap(src.size())
    p = QPainter(pm)
    s = 10
    for y in range(0, pm.height(), s):
        for x in range(0, pm.width(), s):
            p.fillRect(x, y, s, s, QColor("#ffffff") if (x // s + y // s) % 2 == 0 else QColor("#e5e7eb"))
    p.drawPixmap(0, 0, src)
    p.end()
    return pm


class StampDialog(QDialog):
    """Chèn con dấu (stamp) như tem dán nền trong suốt:
    - Con dấu chữ: mẫu có sẵn / chữ tuỳ ý, ngày + người đóng dấu, khung chữ nhật / tròn.
    - Từ ảnh con dấu: ảnh scan / chụp / dán từ clipboard -> tách nền (giữ đúng màu mực con dấu, bỏ chữ đen,
      chữ ký, nền giấy), cắt sát, có thể lưu vào thư viện.
    - Thư viện: con dấu đã lưu."""

    MODES = (("auto", "Tự động - giữ màu mực con dấu (bỏ chữ đen, chữ ký, nền giấy)"),
             ("red", "Mực đỏ"), ("blue", "Mực xanh"), ("purple", "Mực tím"),
             ("dark", "Mực đen / xám (bỏ nền giấy)"), ("white", "Chỉ xoá nền trắng"), ("none", "Giữ nguyên ảnh"))

    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle("Chèn con dấu (stamp)")
        self.resize(640, 640)
        self.settings = QSettings(installer.ORG, "PDFToolkit")
        self.png, self.ratio, self.shape = None, 3.0, "rect"
        self.src_img = None            # QImage ảnh con dấu gốc (tab Từ ảnh)
        self.lib_png = None
        v = QVBoxLayout(self)
        self.tabs = QTabWidget()
        v.addWidget(self.tabs)

        # --- tab 1: con dấu chữ ---
        t1 = QWidget()
        form = QFormLayout(t1)
        self.text = QComboBox()
        self.text.setEditable(True)
        with i18n.paused():                                  # chữ trên con dấu: giữ nguyên, không dịch
            self.text.addItems(STAMP_PRESETS)
            self.text.setCurrentText(self.settings.value("stamp/text", "ĐÃ DUYỆT"))
        form.addRow("Nội dung:", self.text)
        self.color = QComboBox()
        for name, c in STAMP_COLORS:
            self.color.addItem(name, c)
        self.color.setCurrentIndex(int(self.settings.value("stamp/color", 0)))
        form.addRow("Màu:", self.color)
        self.shape_box = QComboBox()
        self.shape_box.addItem("Khung chữ nhật", "rect")
        self.shape_box.addItem("Con dấu tròn", "round")
        self.shape_box.setCurrentIndex(int(self.settings.value("stamp/shape", 0)))
        form.addRow("Kiểu:", self.shape_box)
        self.date = QCheckBox("Thêm ngày giờ và người đóng dấu:")
        self.date.setChecked(self.settings.value("stamp/date", "true") == "true")
        self.who = QLineEdit(self.settings.value("stamp/who", "") or os.environ.get("USERNAME", ""))
        row = QHBoxLayout()
        row.addWidget(self.date)
        row.addWidget(self.who, 1)
        form.addRow("", row)
        self.tilt = QCheckBox("Đóng dấu hơi nghiêng (giống dấu thật)")
        self.tilt.setChecked(self.settings.value("stamp/tilt", "false") == "true")
        form.addRow("", self.tilt)
        self.tabs.addTab(t1, glyph_icon(0xE8D2), "Con dấu chữ")

        # --- tab 2: từ ảnh con dấu (tách nền) ---
        t2 = QWidget()
        f2 = QVBoxLayout(t2)
        row = QHBoxLayout()
        b = QPushButton("Chọn ảnh con dấu...")
        b.setIcon(glyph_icon(0xEB9F))
        b.clicked.connect(self._pick_image)
        row.addWidget(b)
        b = QPushButton("Dán ảnh (Ctrl+V)")
        b.setIcon(glyph_icon(0xE77F))
        b.setToolTip("Dán ảnh vừa chụp màn hình (Win+Shift+S) hoặc copy từ Word / trình duyệt")
        b.clicked.connect(self._paste)
        row.addWidget(b)
        row.addStretch()
        f2.addLayout(row)
        fm = QFormLayout()
        self.mode = QComboBox()
        for code, label in self.MODES:
            self.mode.addItem(label, code)
        self.mode.setCurrentIndex(max(0, self.mode.findData(self.settings.value("stamp/img_mode", "auto"))))
        fm.addRow("Tách nền:", self.mode)
        self.uniform = QCheckBox("Làm đậm, đồng màu mực (giống dấu mới đóng)")
        self.uniform.setChecked(self.settings.value("stamp/uniform", "true") == "true")
        fm.addRow("", self.uniform)
        self.crop = QCheckBox("Cắt sát con dấu")
        self.crop.setChecked(True)
        fm.addRow("", self.crop)
        f2.addLayout(fm)
        self.img_info = QLabel("Chọn hoặc dán ảnh có con dấu (ảnh scan, ảnh chụp văn bản đã đóng dấu...).")
        self.img_info.setWordWrap(True)
        self.img_info.setStyleSheet("color:#64748b")
        f2.addWidget(self.img_info)
        self.save_lib = QPushButton("Lưu vào thư viện con dấu")
        self.save_lib.setIcon(glyph_icon(0xE74E))
        self.save_lib.clicked.connect(self._save_to_lib)
        self.save_lib.setEnabled(False)
        f2.addWidget(self.save_lib, 0, Qt.AlignLeft)
        f2.addStretch()
        self.tabs.addTab(t2, glyph_icon(0xEB9F), "Từ ảnh con dấu (tách nền)")

        # --- tab 3: thư viện ---
        t3 = QWidget()
        f3 = QVBoxLayout(t3)
        self.lib = QListWidget()
        self.lib.setViewMode(QListView.IconMode)
        self.lib.setIconSize(QSize(120, 120))
        self.lib.setGridSize(QSize(140, 140))
        self.lib.setResizeMode(QListView.Adjust)
        self.lib.setMovement(QListView.Static)
        self.lib.currentRowChanged.connect(lambda _r: self._preview())
        self.lib.itemDoubleClicked.connect(lambda _i: self._ok())
        f3.addWidget(self.lib, 1)
        row = QHBoxLayout()
        b = QPushButton("Xoá khỏi thư viện")
        b.setIcon(glyph_icon(0xE74D))
        b.clicked.connect(self._lib_delete)
        row.addWidget(b)
        b = QPushButton("Mở thư mục thư viện")
        b.setIcon(glyph_icon(0xE838))
        b.clicked.connect(lambda: tools_ui.reveal(stamps_dir()))
        row.addWidget(b)
        row.addStretch()
        f3.addLayout(row)
        self.tabs.addTab(t3, glyph_icon(0xE8F1), "Thư viện con dấu")
        self._load_lib()

        self.view = QLabel()
        self.view.setAlignment(Qt.AlignCenter)
        self.view.setMinimumSize(420, 190)
        self.view.setStyleSheet("background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px")
        v.addWidget(self.view, 1)
        tip = QLabel("Bấm vào trang để đặt con dấu (hoặc kéo khung để chọn kích thước). Sau khi đặt: kéo để di "
                     "chuyển, kéo góc dưới phải để đổi cỡ, chuột phải để xoá / nhân bản.")
        tip.setWordWrap(True)
        tip.setStyleSheet("color:#64748b")
        v.addWidget(tip)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.button(QDialogButtonBox.Ok).setText("Chèn con dấu")
        bb.accepted.connect(self._ok)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.text.editTextChanged.connect(self._preview)
        for w in (self.color, self.shape_box, self.mode):
            w.currentIndexChanged.connect(self._preview)
        for w in (self.date, self.tilt, self.uniform, self.crop):
            w.toggled.connect(self._preview)
        self.who.textChanged.connect(self._preview)
        self.tabs.currentChanged.connect(lambda _i: self._preview())
        from PySide6.QtGui import QShortcut
        QShortcut(QKeySequence.Paste, self, activated=self._paste)
        self.tabs.setCurrentIndex(int(self.settings.value("stamp/tab", 0)))
        self._preview()

    # ----- nguồn ảnh -----
    def _pick_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "Chọn ảnh con dấu", self.settings.value("stamp/img_dir", ""),
                                              "Ảnh (*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff)")
        if path:
            self.settings.setValue("stamp/img_dir", os.path.dirname(path))
            self._set_image(QImage(path))

    def _paste(self):
        img = QApplication.clipboard().image()
        if img.isNull():
            QMessageBox.information(self, self.windowTitle(), "Clipboard không có ảnh. Hãy chụp vùng con dấu "
                                    "(Win+Shift+S) hoặc copy ảnh rồi bấm lại.")
            return
        self._set_image(img)

    def _set_image(self, img):
        if img.isNull():
            QMessageBox.warning(self, self.windowTitle(), "Không đọc được ảnh.")
            return
        self.src_img = img
        self.tabs.setCurrentIndex(1)
        self._preview()

    # ----- thư viện -----
    def _load_lib(self):
        self.lib.clear()
        d = stamps_dir()
        for name in sorted(os.listdir(d), key=lambda n: -os.path.getmtime(os.path.join(d, n))):
            if name.lower().endswith(".png"):
                p = os.path.join(d, name)
                it = QListWidgetItem(QIcon(checker_pixmap(open(p, "rb").read(), 120, 120)), "")
                it.setData(Qt.UserRole, p)
                it.setToolTip(name)
                self.lib.addItem(it)
        if self.lib.count():
            self.lib.setCurrentRow(0)

    def _save_to_lib(self):
        try:
            png, _r = self._make_image()
        except Exception as e:
            QMessageBox.warning(self, self.windowTitle(), str(e))
            return
        from datetime import datetime
        name, ok = QInputDialog.getText(self, "Lưu con dấu", "Tên con dấu:",
                                        text="con_dau_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
        if not ok or not name.strip():
            return
        safe = re.sub(r'[<>:"/\\|?*]', "_", name.strip())
        with open(os.path.join(stamps_dir(), safe + ".png"), "wb") as fh:
            fh.write(png)
        self._load_lib()
        self.status_msg("Đã lưu vào thư viện con dấu")

    def status_msg(self, m):
        self.img_info.setText(m)

    def _lib_delete(self):
        it = self.lib.currentItem()
        if not it:
            return
        if QMessageBox.question(self, self.windowTitle(), "Xoá con dấu này khỏi thư viện?") == QMessageBox.Yes:
            try:
                os.remove(it.data(Qt.UserRole))
            except OSError:
                pass
            self._load_lib()
            self._preview()

    # ----- tạo ảnh con dấu -----
    def _sub(self):
        if not self.date.isChecked():
            return ""
        from datetime import datetime
        who = self.who.text().strip()
        return datetime.now().strftime("%d/%m/%Y %H:%M") + (f" - {who}" if who else "")

    def _make_image(self):
        if self.src_img is None:
            raise PdfError("Chưa chọn ảnh con dấu.")
        import numpy as np
        mode = self.mode.currentData()
        rgb = qimage_rgb(self.src_img)
        if mode == "none":
            out = pdf_core.rgba_to_png(np.dstack([rgb, np.full(rgb.shape[:2], 255, np.uint8)]))
            return out, rgb.shape[1] / max(1, rgb.shape[0])
        rgba, ink = pdf_core.extract_stamp(rgb, mode, uniform=self.uniform.isChecked(), crop=self.crop.isChecked())
        if (rgba[..., 3] > 90).sum() < 30:
            raise PdfError("Không tìm thấy con dấu theo màu mực đã chọn - thử chế độ tách nền khác.")
        names = {"red": "đỏ", "blue": "xanh", "purple": "tím", "green": "xanh lá", "dark": "đen / xám", "white": ""}
        if mode == "auto":
            self.status_msg(i18n.tr("Nhận ra mực con dấu màu: ") + i18n.tr(names.get(ink, ink)))
        return pdf_core.rgba_to_png(rgba), rgba.shape[1] / max(1, rgba.shape[0])

    def _make(self):
        tab = self.tabs.currentIndex()
        if tab == 1:
            return self._make_image()
        if tab == 2:
            it = self.lib.currentItem()
            if not it:
                raise PdfError("Thư viện chưa có con dấu - tạo ở tab 'Từ ảnh con dấu'.")
            png = open(it.data(Qt.UserRole), "rb").read()
            pm = QPixmap()
            pm.loadFromData(png)
            return png, pm.width() / max(1, pm.height())
        text = self.text.currentText().strip() or "ĐÃ DUYỆT"
        return render_stamp(text, self._sub(), self.color.currentData(), self.shape_box.currentData(),
                            8 if self.tilt.isChecked() else 0)

    def _preview(self):
        self.save_lib.setEnabled(self.src_img is not None)
        try:
            png, _r = self._make()
        except Exception as e:
            self.view.setPixmap(QPixmap())
            self.view.setText(str(e))
            return
        self.view.setPixmap(checker_pixmap(png, 420, 180))

    def _ok(self):
        try:
            self.png, self.ratio = self._make()
        except Exception as e:
            QMessageBox.warning(self, self.windowTitle(), str(e))
            return
        tab = self.tabs.currentIndex()
        self.shape = self.shape_box.currentData() if tab == 0 else "image"
        s = self.settings
        s.setValue("stamp/tab", tab)
        with i18n.paused():
            s.setValue("stamp/text", self.text.currentText().strip())
        s.setValue("stamp/color", self.color.currentIndex())
        s.setValue("stamp/shape", self.shape_box.currentIndex())
        s.setValue("stamp/date", "true" if self.date.isChecked() else "false")
        s.setValue("stamp/who", self.who.text().strip())
        s.setValue("stamp/tilt", "true" if self.tilt.isChecked() else "false")
        s.setValue("stamp/img_mode", self.mode.currentData())
        s.setValue("stamp/uniform", "true" if self.uniform.isChecked() else "false")
        self.accept()


class SignDialog(QDialog):
    """Chọn chứng thư (USB Token) và thông tin hiển thị của chữ ký."""

    def __init__(self, parent, box_size=None):
        super().__init__(parent)
        self.setWindowTitle("Ký số")
        self.setMinimumWidth(620)
        self.box_size = box_size
        self.settings = QSettings(installer.ORG, "PDFToolkit")
        self.certs = []
        lay = QFormLayout(self)

        row = QHBoxLayout()
        self.cert_box = QComboBox()
        self.cert_box.setMinimumWidth(440)
        self.cert_box.currentIndexChanged.connect(self._cert_changed)
        refresh = QPushButton("Làm mới")
        refresh.setToolTip("Đọc lại danh sách chứng thư (sau khi cắm USB Token)")
        refresh.clicked.connect(self.load_certs)
        row.addWidget(self.cert_box, 1)
        row.addWidget(refresh)
        lay.addRow("Chứng thư số:", row)
        self.cert_info = QLabel()
        self.cert_info.setWordWrap(True)
        self.cert_info.setStyleSheet("color:#64748b")
        lay.addRow("", self.cert_info)

        self.reason = QComboBox()
        self.reason.setEditable(True)
        self.reason.addItems(["", "Tôi đồng ý với nội dung văn bản", "Phê duyệt", "Xác nhận", "Tôi là người lập văn bản"])
        self.reason.setEditText(self.settings.value("sign/reason", ""))
        self.location = QLineEdit(self.settings.value("sign/location", ""))
        self.location.setPlaceholderText("vd: Hà Nội")
        lay.addRow("Lý do ký:", self.reason)
        lay.addRow("Nơi ký:", self.location)

        if box_size:
            self.show_text = QCheckBox("Hiện thông tin người ký (tên, đơn vị, ngày giờ ký)")
            self.show_text.setChecked(self.settings.value("sign/show_text", "true") == "true")
            self.check_only = QCheckBox("Chỉ hiện dấu ✔ xác nhận đã ký (không chữ, không ảnh)")
            self.check_only.setChecked(self.settings.value("sign/check_only", "false") == "true")
            lay.addRow("", self.show_text)
            lay.addRow("", self.check_only)
            row = QHBoxLayout()
            self.image = QLineEdit(self.settings.value("sign/image", ""))
            self.image.setReadOnly(True)
            self.image.setPlaceholderText("(không có)")
            b_img = QPushButton("Chọn ảnh...")
            b_img.clicked.connect(self._pick_image)
            b_clear = QPushButton("Bỏ")
            b_clear.clicked.connect(lambda: (self.image.clear(), self._preview()))
            row.addWidget(self.image, 1); row.addWidget(b_img); row.addWidget(b_clear)
            lay.addRow("Ảnh chữ ký tay:", row)
            self._img_widgets = (self.image, b_img, b_clear)
            self.preview = QLabel()
            self.preview.setAlignment(Qt.AlignCenter)
            self.preview.setStyleSheet("background:#f1f5f9; border:1px solid #e2e8f0; padding:6px")
            self.preview.setMinimumHeight(110)
            lay.addRow("Xem trước:", self.preview)
            self.show_text.toggled.connect(self._preview)
            self.check_only.toggled.connect(self._check_only_changed)
            self._check_only_changed()
            self.reason.editTextChanged.connect(self._preview)
            self.location.textChanged.connect(self._preview)

        row = QHBoxLayout()
        self.use_tsa = QCheckBox("Gắn dấu thời gian (TSA)")
        self.use_tsa.setToolTip("Lấy thời điểm ký từ máy chủ tin cậy qua Internet - chữ ký vẫn hợp lệ sau khi chứng thư hết hạn")
        self.use_tsa.setChecked(self.settings.value("sign/tsa", "false") == "true")
        self.tsa_url = QComboBox()
        self.tsa_url.setEditable(True)
        self.tsa_url.addItems(TSA_URLS)
        self.tsa_url.setEditText(self.settings.value("sign/tsa_url", TSA_URLS[0]))
        row.addWidget(self.use_tsa); row.addWidget(self.tsa_url, 1)
        lay.addRow("", row)

        note = QLabel("Khi bấm Ký, driver của USB Token sẽ hỏi mã PIN. File đã ký được lưu thành file mới.\n"
                      "Sau khi ký, mọi chỉnh sửa và lưu đè sẽ làm chữ ký mất hiệu lực.")
        note.setStyleSheet("color:#64748b")
        lay.addRow("", note)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.ok = btns.button(QDialogButtonBox.Ok)
        self.ok.setText("Ký")
        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addRow(btns)
        self.load_certs()

    def load_certs(self):
        import pdf_sign
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            self.certs = pdf_sign.list_signing_certs(include_expired=True)
        finally:
            QApplication.restoreOverrideCursor()
        self.cert_box.blockSignals(True)
        self.cert_box.clear()
        last = self.settings.value("sign/cert", "")
        sel = 0
        for i, c in enumerate(self.certs):
            self.cert_box.addItem(("[HẾT HẠN] " if c.expired else "") + c.label())
            if c.serial == last:
                sel = i
        self.cert_box.blockSignals(False)
        if not self.certs:
            self.cert_box.addItem("(Không tìm thấy chứng thư số)")
            self.cert_info.setText("Hãy cắm USB Token, cài driver / phần mềm quản lý token của nhà cung cấp "
                                   "(VNPT-CA, Viettel-CA, BKAV-CA, FPT-CA...) rồi bấm Làm mới.")
            self.cert_info.setStyleSheet("color:#b45309")
        self.cert_box.setCurrentIndex(sel)
        self._cert_changed()

    def cert(self):
        i = self.cert_box.currentIndex()
        return self.certs[i] if 0 <= i < len(self.certs) else None

    def _cert_changed(self):
        c = self.cert()
        self.ok.setEnabled(bool(c) and not c.expired)
        if c:
            self.cert_info.setStyleSheet("color:#b91c1c" if c.expired else "color:#64748b")
            self.cert_info.setText(f"Cấp bởi: {c.issuer}  |  Số serial: {c.serial}  |  Hiệu lực: "
                                   f"{c.not_before:%d/%m/%Y} - {c.not_after:%d/%m/%Y}"
                                   + ("  — ĐÃ HẾT HẠN, không ký được" if c.expired else ""))
        self._preview()

    def _check_only_changed(self):
        on = self.check_only.isChecked()                # chỉ dấu ✔: tắt các tuỳ chọn chữ / ảnh
        self.show_text.setEnabled(not on)
        for wdg in self._img_widgets:
            wdg.setEnabled(not on)
        self._preview()

    def _pick_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "Chọn ảnh chữ ký tay", os.path.dirname(self.image.text()),
                                              "Ảnh (*.png *.jpg *.jpeg *.bmp)")
        if path:
            self.image.setText(path)
            self._preview()

    def image_path(self):
        if not self.box_size:
            return None
        p = self.image.text()
        return p if p and os.path.exists(p) else None

    def _preview(self):
        if not self.box_size or not hasattr(self, "preview"):
            return
        import pdf_sign
        c = self.cert()
        w, h = self.box_size
        if not c:
            self.preview.clear()
            return
        only = self.check_only.isChecked()
        lines = pdf_sign.stamp_lines(c, self.reason.currentText().strip(), self.location.text().strip()) \
            if self.show_text.isChecked() and not only else []
        try:
            doc = pdf_sign.appearance_doc(w, h, lines, None if only else self.image_path(), only)
            k = min(2.5, 520 / w, 180 / h)
            pix = doc[0].get_pixmap(matrix=pymupdf.Matrix(k, k), alpha=False)
            doc.close()
            self.preview.setPixmap(QPixmap.fromImage(pix_to_qimage(pix)))
        except Exception as e:
            self.preview.setText(f"Không xem trước được: {e}")

    def accept(self):
        s = self.settings
        c = self.cert()
        if c:
            s.setValue("sign/cert", c.serial)
        s.setValue("sign/reason", self.reason.currentText().strip())
        s.setValue("sign/location", self.location.text().strip())
        s.setValue("sign/tsa", "true" if self.use_tsa.isChecked() else "false")
        s.setValue("sign/tsa_url", self.tsa_url.currentText().strip())
        if self.box_size:
            s.setValue("sign/show_text", "true" if self.show_text.isChecked() else "false")
            s.setValue("sign/check_only", "true" if self.check_only.isChecked() else "false")
            s.setValue("sign/image", self.image.text())
        super().accept()


class InstallDialog(QDialog):
    """Cài PDF Toolkit cho tài khoản hiện tại (không cần quyền admin) + add-in Office."""

    def __init__(self, parent, office_only=False):
        super().__init__(parent)
        self.exe = None
        self.setWindowTitle("Cài đặt PDF Toolkit" if not office_only else "Add-in Office")
        self.resize(560, 0)
        v = QVBoxLayout(self)
        head = QLabel("Cài PDF Toolkit cho tài khoản Windows này - <b>không cần quyền administrator</b>.")
        head.setWordWrap(True)
        v.addWidget(head)
        cur = installer.installed_version()
        if cur:
            lb = QLabel(f"Đang cài bản {cur} - cài lại để cập nhật lên bản {APP_VERSION}.")
            lb.setStyleSheet("color:#64748b")
            v.addWidget(lb)
        row = QHBoxLayout()
        row.addWidget(QLabel("Thư mục cài:"))
        self.dir = QLineEdit(installer._get(installer.INFO_KEY, "Dir") or installer.default_dir())
        row.addWidget(self.dir, 1)
        b = QPushButton("Chọn...")
        b.clicked.connect(self._pick)
        row.addWidget(b)
        v.addLayout(row)
        self.c_start = QCheckBox("Shortcut trong Start Menu")
        self.c_desk = QCheckBox("Shortcut ngoài Desktop")
        self.c_open = QCheckBox('Thêm vào "Open with" cho file PDF (chọn làm mặc định trong Settings > Default apps)')
        self.c_office = QCheckBox("Add-in Office: tab PDF Toolkit trong Word, Excel, PowerPoint, Outlook")
        self.c_office.setToolTip("Xuất tài liệu ra PDF rồi ký số / đặt mật khẩu / nén / watermark / gộp bằng PDF Toolkit")
        self.c_ctx = QCheckBox('Menu chuột phải "PDF Toolkit" cho file PDF (nén, tách, chuyển Word / Excel, OCR, ký số...)')
        self.c_ctx.setToolTip('Windows 11: bấm "Show more options" hoặc Shift + chuột phải để thấy menu')
        self.c_expl = QCheckBox("Explorer: ảnh thu nhỏ trang đầu và khung xem trước (Preview pane) cho file PDF")
        self.c_expl.setToolTip("Khung xem trước chỉ thêm khi máy chưa có (vd Microsoft Edge đã có sẵn thì giữ của Edge)")
        for c in (self.c_start, self.c_desk, self.c_open, self.c_ctx, self.c_expl, self.c_office):
            c.setChecked(True)
            v.addWidget(c)
        if not installer.office_present():
            self.c_office.setChecked(False)
            self.c_office.setEnabled(False)
            self.c_office.setText(self.c_office.text() + "  (không tìm thấy Office)")
        if office_only and installer.installed_exe():
            for c in (self.c_start, self.c_desk, self.c_open):
                c.setChecked(bool(installer._get(installer.INFO_KEY, "Shortcuts")) if c is not self.c_open
                             else bool(installer._get(installer.INFO_KEY, "OpenWith")))
            ctx = installer._get(installer.INFO_KEY, "ContextMenu")
            self.c_ctx.setChecked(True if ctx is None else bool(ctx))
            expl = installer._get(installer.INFO_KEY, "Explorer")
            self.c_expl.setChecked(True if expl is None else bool(expl))
        from PySide6.QtWidgets import QProgressBar
        self.bar = QProgressBar()
        self.bar.setRange(0, 0)                         # chạy liên tục (không biết trước thời gian)
        self.bar.setTextVisible(False)
        self.bar.setFixedHeight(6)
        self.bar.hide()
        v.addWidget(self.bar)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(110)
        self.log.hide()
        v.addWidget(self.log)
        self.bb = bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.ok = bb.button(QDialogButtonBox.Ok)
        self.ok.setText("Cài đặt")
        bb.button(QDialogButtonBox.Cancel).setText("Huỷ")
        bb.accepted.connect(self._install)
        bb.rejected.connect(self.reject)
        v.addWidget(bb)
        self.job = None

    def _pick(self):
        d = QFileDialog.getExistingDirectory(self, "Thư mục cài đặt", self.dir.text())
        if d:
            self.dir.setText(os.path.join(os.path.normpath(d), installer.APP)
                             if os.path.basename(d) != installer.APP else os.path.normpath(d))

    def _install(self):
        """Cài trong luồng nền (chép ~200 MB mất 1-2 phút): cửa sổ vẫn vẽ lại, nhật ký hiện dần."""
        if self.job is not None:
            return
        self.log.clear()
        self.log.show()
        self.bar.show()
        self._set_busy(True)
        self.log.appendPlainText(i18n.tr("Đang cài đặt, vui lòng chờ (1-2 phút)..."))
        args = (self.dir.text().strip(), self.c_start.isChecked(), self.c_desk.isChecked(),
                self.c_open.isChecked(), self.c_office.isChecked(), APP_VERSION)
        self.job = InstallJob(args, self.c_ctx.isChecked(), self.c_expl.isChecked())
        self.job.log.connect(self.log.appendPlainText)
        self.job.done.connect(self._installed)
        self.job.start()

    def _set_busy(self, busy):
        for cls in (QCheckBox, QLineEdit, QPushButton):
            for w in self.findChildren(cls):
                w.setEnabled(not busy)
        if not busy and not installer.office_present():
            self.c_office.setEnabled(False)

    def _installed(self, exe, err):
        self.job = None
        self.bar.hide()
        self._set_busy(False)
        if err:
            self.log.appendPlainText("✘ " + err)
            return
        self.exe = exe
        QMessageBox.information(self, self.windowTitle(), self.log.toPlainText())
        self.accept()

    def reject(self):
        if self.job is not None:                        # đang cài: không cho đóng giữa chừng
            return
        super().reject()

    def closeEvent(self, e):
        if self.job is not None:
            e.ignore()
            return
        super().closeEvent(e)


class InstallJob(tools_ui.SafeThread):
    """Chạy installer.install trong luồng nền (COM của shortcut cần CoInitialize riêng cho luồng)."""
    log = Signal(str)
    done = Signal(str, str)                             # (exe đã cài, lỗi)

    def __init__(self, args, context_menu, explorer=True):
        super().__init__()
        self.args = args
        self.context_menu = context_menu
        self.explorer = explorer

    def run(self):
        exe, err = "", ""
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pythoncom = None
        try:
            exe = installer.install(*self.args, log=self.log.emit, context_menu=self.context_menu,
                                    explorer=self.explorer) or ""
        except Exception as e:
            err = str(e)
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        self.done.emit(exe, err)


class VerifyDialog(QDialog):
    def __init__(self, parent, results, title="Chữ ký số trong file"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(820, 300)
        lay = QVBoxLayout(self)
        if not results:
            lay.addWidget(QLabel("File này không có chữ ký số."))
        else:
            t = QTableWidget(len(results), 5)
            t.setHorizontalHeaderLabels(["Trường chữ ký", "Người ký", "Thời điểm ký", "Toàn vẹn", "Phạm vi"])
            t.verticalHeader().setVisible(False)
            t.setEditTriggers(QAbstractItemView.NoEditTriggers)
            t.horizontalHeader().setStretchLastSection(True)
            for i, r in enumerate(results):
                when = r["time"].astimezone().strftime("%d/%m/%Y %H:%M:%S") if r["time"] else ""
                if r["error"]:
                    ok_txt, ok_col = "Lỗi: " + r["error"][:80], "#b91c1c"
                elif r["intact"]:
                    ok_txt, ok_col = "✔ Nội dung không bị sửa", "#15803d"
                else:
                    ok_txt, ok_col = "✘ Nội dung đã bị sửa sau khi ký", "#b91c1c"
                cover = "Toàn bộ file" if r["covers_all"] else "Có thêm nội dung sau khi ký (vd: chữ ký khác)"
                for j, v in enumerate((r["field"], r["signer"], when, ok_txt, cover)):
                    it = QTableWidgetItem(v)
                    if j == 3:
                        it.setForeground(QColor(ok_col))
                    t.setItem(i, j, it)
            t.resizeColumnsToContents()
            lay.addWidget(t)
        note = QLabel("Chỉ kiểm tra tính toàn vẹn (file không bị sửa sau khi ký). Trạng thái thu hồi của chứng thư "
                      "cần kiểm tra bằng Adobe Acrobat Reader hoặc phần mềm của nhà cung cấp CA.")
        note.setWordWrap(True)
        note.setStyleSheet("color:#64748b")
        lay.addWidget(note)
        btns = QDialogButtonBox(QDialogButtonBox.Close)
        btns.button(QDialogButtonBox.Close).setText("Đóng")
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)


class FindReplaceDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.main = parent
        self.setWindowTitle("Tìm và thay thế")
        self.setMinimumWidth(480)
        lay = QFormLayout(self)
        self.find = QLineEdit(); self.repl = QLineEdit()
        self.case = QCheckBox("Phân biệt chữ hoa / thường")
        self.fit = QCheckBox("Co chữ cho vừa nếu nội dung mới dài hơn"); self.fit.setChecked(True)
        self.scope = QComboBox(); self.scope.addItems(["Tất cả các trang", "Chỉ trang đang xem", "Các trang đang chọn"])
        self.result = QLabel(""); self.result.setStyleSheet("color:#2563eb")
        lay.addRow("Tìm:", self.find)
        lay.addRow("Thay bằng:", self.repl)
        lay.addRow("Phạm vi:", self.scope)
        lay.addRow("", self.case)
        lay.addRow("", self.fit)
        lay.addRow("", self.result)
        row = QHBoxLayout()
        b_count = QPushButton("Đếm"); b_count.clicked.connect(self.count)
        b_repl = QPushButton("Thay tất cả"); b_repl.setDefault(True); b_repl.clicked.connect(self.replace)
        b_close = QPushButton("Đóng"); b_close.clicked.connect(self.close)
        row.addStretch(); row.addWidget(b_count); row.addWidget(b_repl); row.addWidget(b_close)
        lay.addRow(row)

    def _pages(self):
        return self.main.scope_pages(self.scope.currentIndex())

    def count(self):
        res = self.main.pdf.find_text(self.find.text(), self.case.isChecked(), self._pages())
        total = sum(res.values())
        pages = ", ".join(str(p + 1) for p in sorted(res))[:120]
        self.result.setText(f"Tìm thấy {total} chỗ" + (f" ở trang: {pages}" if total else ""))

    def replace(self):
        if not self.find.text():
            return
        n = self.main.run_edit(lambda: self.main.pdf.find_replace(
            self.find.text(), self.repl.text(), self.case.isChecked(), self.fit.isChecked(), self._pages()), all_pages=True)
        self.result.setText(f"Đã thay {n or 0} chỗ." + ("" if n else " (Không tìm thấy - chữ có thể là ảnh scan)"))


class ScanDialog(QDialog):
    """Tuỳ chọn quét: nguồn giấy, màu, độ phân giải."""

    def __init__(self, parent, insert=False):
        super().__init__(parent)
        self.setWindowTitle("Chèn trang từ máy scan" if insert else "Máy scan → PDF")
        self.setMinimumWidth(420)
        st = parent.settings
        lay = QFormLayout(self)
        self.source = QComboBox()
        self.source.addItem("Mặt kính (từng trang, hỏi quét tiếp)", "flatbed")
        self.source.addItem("Khay nạp giấy tự động - ADF (quét đến hết giấy)", "feeder")
        self.color = QComboBox()
        self.color.addItem("Màu", "color"); self.color.addItem("Xám", "gray")
        self.color.addItem("Đen trắng (văn bản, file nhỏ nhất)", "bw")
        self.dpi = QComboBox()
        for d, t in ((150, "150 dpi - nhẹ, đủ đọc"), (200, "200 dpi - khuyên dùng"), (300, "300 dpi - nét, để OCR / in"),
                     (600, "600 dpi - rất nét (chậm, file nặng)")):
            self.dpi.addItem(t, d)
        for cb, key, default in ((self.source, "scan_source", "flatbed"), (self.color, "scan_color", "color"),
                                 (self.dpi, "scan_dpi", 200)):
            v = st.value(key, default)
            i = cb.findData(int(v) if cb is self.dpi else v)
            cb.setCurrentIndex(max(0, i))
        lay.addRow("Nguồn giấy:", self.source)
        lay.addRow("Màu:", self.color)
        lay.addRow("Độ phân giải:", self.dpi)
        note = QLabel("Dùng máy scan có driver WIA của Windows (hầu hết máy scan / máy in đa năng). Sau khi quét có thể "
                      "bấm <b>OCR</b> để tìm và copy được chữ.")
        note.setWordWrap(True); note.setStyleSheet("color:#64748b")
        lay.addRow(note)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Quét")
        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addRow(btns)

    def values(self):
        return self.source.currentData(), self.color.currentData(), int(self.dpi.currentData())


class RedactFindDialog(QDialog):
    """Tìm chữ / mẫu thông tin nhạy cảm để đánh dấu xoá vĩnh viễn (chưa xoá ngay - xem lại khung đỏ rồi áp dụng)."""

    def __init__(self, parent, current=0):
        super().__init__(parent)
        self.main = parent
        self.setWindowTitle("Tìm & đánh dấu xoá")
        self.setMinimumWidth(500)
        lay = QFormLayout(self)
        self.terms = QPlainTextEdit()
        self.terms.setPlaceholderText("Mỗi dòng 1 cụm chữ, vd:\nNguyễn Văn A\nCông ty ABC")
        self.terms.setFixedHeight(90)
        lay.addRow("Chữ cần xoá:", self.terms)
        self.presets = {}
        box = QVBoxLayout()
        for k, (label, _pat) in pdf_core.REDACT_PRESETS.items():
            c = QCheckBox("Mọi " + label[0].lower() + label[1:])
            self.presets[k] = c
            box.addWidget(c)
        lay.addRow("Loại thông tin:", box)
        self.whole = QCheckBox("Chỉ khớp nguyên từ")
        lay.addRow("", self.whole)
        self.scope = QComboBox(); self.scope.addItems(["Tất cả các trang", "Chỉ trang đang xem", "Các trang đang chọn"])
        lay.addRow("Phạm vi:", self.scope)
        self.fill = QComboBox()
        self.fill.addItem("Màu đen", "black"); self.fill.addItem("Màu trắng", "white")
        self.fill.setCurrentIndex(max(0, self.fill.findData(parent.settings.value("redact_fill", "black"))))
        lay.addRow("Tô chỗ đã xoá:", self.fill)
        note = QLabel("Các chỗ tìm thấy được đánh dấu khung đỏ để xem lại; bấm <b>Xoá vĩnh viễn</b> để xoá hẳn.")
        note.setWordWrap(True); note.setStyleSheet("color:#64748b")
        lay.addRow(note)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Đánh dấu")
        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")
        btns.accepted.connect(self._ok)
        btns.rejected.connect(self.reject)
        lay.addRow(btns)

    def _ok(self):
        t, p = self.values()[:2]
        if not t and not p:
            QMessageBox.information(self, APP_NAME, "Nhập chữ cần xoá hoặc chọn ít nhất một loại thông tin.")
            return
        self.accept()

    def values(self):
        terms = [t for t in self.terms.toPlainText().splitlines() if t.strip()]
        pats = [pdf_core.REDACT_PRESETS[k][1] for k, c in self.presets.items() if c.isChecked()]
        return terms, pats, self.whole.isChecked(), self.fill.currentData(), self.main.scope_pages(
            self.scope.currentIndex())


class ImageInsertDialog(QDialog):
    """Chèn ảnh: chọn file, xem trước, chọn định dạng (nội dung trang / con dấu - stamp)."""
    FILTER = "Ảnh (*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff *.webp)"

    def __init__(self, parent, start_dir="", as_stamp=False):
        super().__init__(parent)
        self.setWindowTitle(i18n.tr("Chèn ảnh"))
        self.resize(620, 560)
        self.path = ""
        self.start_dir = start_dir
        v = QVBoxLayout(self)
        g = QGroupBox(i18n.tr("Nguồn ảnh"))
        h = QHBoxLayout(g)
        h.addWidget(QLabel(i18n.tr("Ảnh:")))
        self.src = QLineEdit()
        self.src.setReadOnly(True)
        self.src.setPlaceholderText(i18n.tr("Chưa chọn ảnh"))
        h.addWidget(self.src, 1)
        b = QPushButton(glyph_icon(0xE8E5), i18n.tr("Chọn ảnh..."))
        b.clicked.connect(self.browse)
        h.addWidget(b)
        v.addWidget(g)
        g = QGroupBox(i18n.tr("Xem trước"))
        pv = QVBoxLayout(g)
        self.preview = QLabel()
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setMinimumSize(560, 300)
        self.preview.setStyleSheet("background:#71717a; border-radius:6px; color:#e4e4e7")
        pv.addWidget(self.preview, 1)
        v.addWidget(g, 1)
        g = QGroupBox(i18n.tr("Định dạng"))
        ov = QVBoxLayout(g)
        self.r_content = QRadioButton(i18n.tr("Nội dung trang - ảnh nằm trong trang như ảnh gốc của file"))
        self.r_stamp = QRadioButton(i18n.tr("Con dấu (stamp) - nằm trên trang như ghi chú: ẩn / xoá / cố định "
                                            "được, không trộn vào nội dung"))
        (self.r_stamp if as_stamp else self.r_content).setChecked(True)
        ov.addWidget(self.r_content)
        ov.addWidget(self.r_stamp)
        v.addWidget(g)
        bb = QHBoxLayout()
        bb.addStretch()
        self.ok = QPushButton(i18n.tr("Chèn"))
        self.ok.setObjectName("primary")
        self.ok.setEnabled(False)
        self.ok.clicked.connect(self.accept)
        cancel = QPushButton(i18n.tr("Huỷ"))
        cancel.clicked.connect(self.reject)
        bb.addWidget(self.ok)
        bb.addWidget(cancel)
        v.addLayout(bb)
        self._pm = None

    @property
    def as_stamp(self):
        return self.r_stamp.isChecked()

    def browse(self):
        path, _ = QFileDialog.getOpenFileName(self, i18n.tr("Chọn ảnh để chèn"),
                                              os.path.dirname(self.path) if self.path else self.start_dir,
                                              self.FILTER)
        if path:
            self.set_path(path)

    def set_path(self, path):
        pm = QPixmap(path)
        if pm.isNull():
            QMessageBox.warning(self, APP_NAME, i18n.tr("Không đọc được ảnh:") + "\n" + path)
            return
        self.path, self._pm = path, pm
        self.src.setText(f"{os.path.basename(path)}  ({os.path.dirname(path)})")
        self.src.setCursorPosition(0)                   # hiện tên file ở đầu (không cuộn tới cuối đường dẫn)
        self.ok.setEnabled(True)
        self._show_preview()

    def _show_preview(self):
        if self._pm is not None:
            self.preview.setPixmap(self._pm.scaled(self.preview.size() - QSize(24, 24), Qt.KeepAspectRatio,
                                                   Qt.SmoothTransformation))

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._show_preview()


class TextResultDialog(QDialog):
    """Hiện đoạn chữ (vd kết quả OCR vùng chọn) để xem / sửa / sao chép."""

    def __init__(self, parent, title, text):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(560, 380)
        v = QVBoxLayout(self)
        self.edit = QPlainTextEdit(text or "")
        if not text:
            self.edit.setPlaceholderText(i18n.tr("Không nhận dạng được chữ nào trong vùng này."))
        v.addWidget(self.edit, 1)
        row = QHBoxLayout()
        row.addStretch()
        b_copy = QPushButton(i18n.tr("Sao chép"))
        b_copy.setObjectName("primary")
        b_copy.clicked.connect(self._copy)
        b_close = QPushButton(i18n.tr("Đóng"))
        b_close.clicked.connect(self.reject)
        row.addWidget(b_copy)
        row.addWidget(b_close)
        v.addLayout(row)

    def _copy(self):
        QApplication.clipboard().setText(self.edit.toPlainText())
        parent = self.parent()
        self.accept()
        if hasattr(parent, "toast"):
            parent.toast("Đã sao chép chữ vào bộ nhớ tạm", "Dán (Ctrl+V) vào Word, Outlook, Zalo...")


class CompressDialog(QDialog):
    def __init__(self, parent, current_size):
        super().__init__(parent)
        self.setWindowTitle("Nén PDF")
        self.setMinimumWidth(460)
        lay = QVBoxLayout(self)
        lay.addWidget(QLabel(f"Kích thước hiện tại: <b>{fmt_size(current_size)}</b>"))
        self.group = QButtonGroup(self)
        opts = [("low", "Nhẹ", "Chỉ tối ưu cấu trúc file, không giảm chất lượng ảnh."),
                ("medium", "Trung bình (khuyến nghị)", "Ảnh 150 dpi, JPEG 75% — phù hợp gửi email, in ấn thông thường."),
                ("high", "Mạnh", "Ảnh 96 dpi, JPEG 55% — file nhỏ nhất, ảnh có thể hơi mờ.")]
        for i, (key, title, desc) in enumerate(opts):
            rb = QRadioButton(title); rb.setProperty("level", key)
            if key == "medium":
                rb.setChecked(True)
            d = QLabel(desc); d.setStyleSheet("color:#64748b; margin-left:22px; margin-bottom:6px")
            self.group.addButton(rb, i)
            lay.addWidget(rb); lay.addWidget(d)
        note = QLabel("Chữ vẫn giữ nguyên (chọn, tìm kiếm được). Có thể hoàn tác (Ctrl+Z).")
        note.setStyleSheet("color:#64748b")
        lay.addWidget(note)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.button(QDialogButtonBox.Ok).setText("Nén")
        btns.button(QDialogButtonBox.Cancel).setText("Huỷ")
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def level(self):
        return self.group.checkedButton().property("level")


class DocTab:
    """Một tài liệu đang mở trong thanh tab: tài liệu (kèm lịch sử hoàn tác) + trang / zoom / vị trí cuộn."""

    def __init__(self, pdf):
        self.pdf = pdf
        self.current = 0
        self.zoom = 1.0
        self.fit_width = True
        self.scroll = 0


# ===========================================================================
#  Hộp thoại In có xem trước (hộp thoại in Windows 11 không xem trước được với ứng dụng desktop)
# ===========================================================================
class PrintDialog(QDialog):
    """Hộp thoại In kiểu Word: máy in, số bản, trang, 2 mặt, theo bộ, số trang / tờ, khổ giấy, hướng, thu phóng, màu
    - xem trước đúng như bản in."""
    NUPS = (1, 2, 4, 6, 9, 16)

    def __init__(self, win):
        super().__init__(win)
        from PySide6.QtPrintSupport import QPrinter, QPrinterInfo, QPrintPreviewWidget
        from PySide6.QtWidgets import QFrame
        self.win = win
        self.n = win.pdf.page_count
        self.setWindowTitle("In")
        self.resize(1120, 780)
        self.printer = QPrinter(QPrinter.HighResolution)
        with i18n.paused():
            self.printer.setDocName(os.path.basename(win.pdf.path or "PDF Toolkit"))
        self.settings = QSettings(installer.ORG, "PDFToolkit")
        lay = QHBoxLayout(self)
        side = QWidget()
        side.setObjectName("printSide")
        left = QVBoxLayout(side)
        left.setContentsMargins(0, 0, 8, 0)
        left.setSpacing(6)
        sc = QScrollArea()
        sc.setWidget(side)
        sc.setWidgetResizable(True)
        sc.setFrameShape(QFrame.NoFrame)
        sc.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        sc.setMinimumWidth(330)
        outer = QVBoxLayout()
        outer.addWidget(sc, 1)
        lay.addLayout(outer, 0)

        def head(text):
            l = QLabel(f"<b>{i18n.tr(text)}</b>")
            l.setStyleSheet("font-size:11pt; margin-top:6px")
            left.addWidget(l)

        def combo(label, items):
            if label:
                left.addWidget(QLabel(label))
            c = QComboBox()
            c.addItems(items)
            left.addWidget(c)
            return c

        # --- In + số bản ---
        top = QHBoxLayout()
        self.ok = QPushButton(glyph_icon(0xE749, "#ffffff"), "In")
        self.ok.setObjectName("primary")
        self.ok.setDefault(True)
        self.ok.setMinimumSize(110, 44)
        self.ok.clicked.connect(self.accept)
        top.addWidget(self.ok)
        top.addSpacing(10)
        top.addWidget(QLabel("Số bản:"))
        self.copies = QSpinBox()
        self.copies.setRange(1, 999)
        self.copies.setMinimumWidth(70)
        top.addWidget(self.copies)
        top.addStretch()
        left.addLayout(top)

        # --- máy in ---
        head("Máy in")
        self.printer_box = QComboBox()
        names = QPrinterInfo.availablePrinterNames()
        with i18n.paused():
            self.printer_box.addItems(names)
        default = QPrinterInfo.defaultPrinterName()
        if default in names:
            self.printer_box.setCurrentText(default)
        left.addWidget(self.printer_box)
        more = QPushButton("Thuộc tính máy in...")
        more.setToolTip("Mở hộp thoại in của Windows (khay giấy, chất lượng in... riêng của máy in) - bấm In ở đó sẽ in ngay")
        more.clicked.connect(self._windows_dialog)
        left.addWidget(more, 0, Qt.AlignRight)

        # --- cài đặt ---
        head("Tuỳ chọn in")
        self.which = combo(None, [f"In tất cả trang ({self.n})", f"Trang hiện tại ({win.current + 1})",
                                  f"Các trang đang chọn ({len(win.selected_pages())})", "Trang tuỳ chọn"])
        rr = QHBoxLayout()
        rr.addWidget(QLabel("Trang:"))
        self.rng = QLineEdit()
        self.rng.setPlaceholderText("vd: 1-3, 5, 8-")
        rr.addWidget(self.rng, 1)
        left.addLayout(rr)
        self.subset = combo(None, ["Mọi trang trong phạm vi", "Chỉ trang lẻ", "Chỉ trang chẵn"])
        self.duplex = combo(None, ["In 1 mặt", "In 2 mặt - lật theo cạnh dài", "In 2 mặt - lật theo cạnh ngắn"])
        self.collate = combo(None, ["Theo bộ  (1,2,3  1,2,3  1,2,3)", "Không theo bộ  (1,1,1  2,2,2  3,3,3)"])
        self.nup = combo(None, [i18n.tr("{0} trang / tờ").format(k) for k in self.NUPS])
        self.order = combo(None, ["Thứ tự: ngang rồi xuống", "Thứ tự: dọc rồi sang"])

        head("Khổ giấy và bố cục")
        self.paper = QComboBox()
        left.addWidget(self.paper)
        self.orient = combo(None, ["Hướng dọc", "Hướng ngang"])
        self.scale = combo(None, ["Vừa khổ giấy", "Kích thước thật (100%)", "Chỉ thu nhỏ trang quá khổ", "Tuỳ chỉnh"])
        pr_ = QHBoxLayout()
        pr_.addWidget(QLabel("Tỉ lệ:"))
        self.pct = QSpinBox()
        self.pct.setRange(10, 400)
        self.pct.setSuffix(" %")
        self.pct.setValue(100)
        pr_.addWidget(self.pct)
        pr_.addStretch()
        left.addLayout(pr_)
        self.autorot = QCheckBox("Tự xoay trang cho vừa giấy")
        self.autorot.setChecked(True)
        left.addWidget(self.autorot)
        self.center = QCheckBox("Căn giữa trang trên giấy")
        self.center.setChecked(True)
        left.addWidget(self.center)
        self.border = QCheckBox("Kẻ viền quanh mỗi trang (khi in nhiều trang / tờ)")
        left.addWidget(self.border)
        self.color = combo(None, ["In màu", "In đen trắng"])

        self.info = QLabel()
        self.info.setWordWrap(True)
        left.addSpacing(4)
        left.addWidget(self.info)
        left.addStretch()
        cancel = QPushButton("Huỷ")
        cancel.clicked.connect(self.reject)
        outer.addWidget(cancel, 0, Qt.AlignRight)

        # --- xem trước ---
        right = QVBoxLayout()
        lay.addLayout(right, 1)
        nav = QHBoxLayout()
        self.prev_btn = QToolButton()
        self.prev_btn.setIcon(glyph_icon(0xE76B))
        self.next_btn = QToolButton()
        self.next_btn.setIcon(glyph_icon(0xE76C))
        self.page_lbl = QLabel()
        nav.addStretch()
        nav.addWidget(self.prev_btn)
        nav.addWidget(self.page_lbl)
        nav.addWidget(self.next_btn)
        nav.addStretch()
        right.addLayout(nav)
        self.preview = QPrintPreviewWidget(self.printer)
        self.preview.setViewMode(QPrintPreviewWidget.SinglePageView)
        self.preview.setZoomMode(QPrintPreviewWidget.FitInView)
        self.preview.paintRequested.connect(self._paint)
        self.preview.previewChanged.connect(self._nav_update)
        right.addWidget(self.preview, 1)
        self.prev_btn.clicked.connect(lambda: self._go(-1))
        self.next_btn.clicked.connect(lambda: self._go(1))

        self._timer = QTimer(self)                    # gõ phạm vi / tỉ lệ: chờ ngừng gõ mới vẽ lại
        self._timer.setSingleShot(True)
        self._timer.setInterval(350)
        self._timer.timeout.connect(self._refresh)
        self.rng.textEdited.connect(lambda _: self.which.setCurrentIndex(3))
        self.rng.textChanged.connect(lambda _: self._timer.start())
        self.pct.valueChanged.connect(lambda _: (self.scale.setCurrentIndex(3), self._timer.start()))
        for c in (self.which, self.subset, self.collate, self.nup, self.order, self.scale, self.color):
            c.currentIndexChanged.connect(lambda _=0: self._refresh())
        for c in (self.autorot, self.center, self.border):
            c.toggled.connect(lambda _=False: self._refresh())
        self.duplex.currentIndexChanged.connect(lambda _: self._apply_duplex())
        self.orient.currentIndexChanged.connect(lambda _: self._apply_layout())
        self.paper.currentIndexChanged.connect(lambda _: self._apply_layout())
        self.printer_box.currentTextChanged.connect(self._printer_changed)
        self._restore()
        if not names:
            self.ok.setEnabled(False)
            self.info.setText("Máy tính chưa cài máy in nào.")
        else:
            self._printer_changed(self.printer_box.currentText())

    # ----- lưu / khôi phục lựa chọn lần trước (trừ phạm vi trang) -----
    _KEEP = ("subset", "duplex", "collate", "nup", "order", "orient", "scale", "color")

    def _restore(self):
        for k in self._KEEP:
            v = self.settings.value(f"print/{k}", None)
            if v is not None:
                try:
                    getattr(self, k).setCurrentIndex(int(v))
                except (TypeError, ValueError):
                    pass
        self.pct.blockSignals(True)
        self.pct.setValue(int(self.settings.value("print/pct", 100)))
        self.pct.blockSignals(False)
        self.autorot.setChecked(self.settings.value("print/autorot", True, type=bool))
        self.center.setChecked(self.settings.value("print/center", True, type=bool))
        self.border.setChecked(self.settings.value("print/border", False, type=bool))

    def _save(self):
        for k in self._KEEP:
            self.settings.setValue(f"print/{k}", getattr(self, k).currentIndex())
        self.settings.setValue("print/pct", self.pct.value())
        pid = self.paper.currentData()
        if pid is not None:
            self.settings.setValue("print/paper", int(pid))
        self.settings.setValue("print/autorot", self.autorot.isChecked())
        self.settings.setValue("print/center", self.center.isChecked())
        self.settings.setValue("print/border", self.border.isChecked())

    # ----- máy in -----
    def _printer_changed(self, name):
        from PySide6.QtPrintSupport import QPrinterInfo, QPrinter
        with i18n.paused():
            self.printer.setPrinterName(name)
        info = QPrinterInfo.printerInfo(name)
        # khổ giấy máy in hỗ trợ (giữ khổ đang chọn nếu có, mặc định khổ của máy in)
        cur = self.paper.currentData()
        self.paper.blockSignals(True)
        self.paper.clear()
        sizes = info.supportedPageSizes() or [QPageSize(QPageSize.A4)]
        with i18n.paused():
            for ps in sizes:
                self.paper.addItem(ps.name(), int(ps.id().value))
        saved = self.settings.value("print/paper", None)
        # ưu tiên: khổ đang chọn > khổ dùng lần trước > A4 > khổ mặc định của máy in
        i = -1
        for want in (cur, int(saved) if saved not in (None, "") else None, int(QPageSize.A4.value),
                     int(info.defaultPageSize().id().value)):
            if want is not None and i < 0:
                i = self.paper.findData(want)
        self.paper.setCurrentIndex(max(0, i))
        self.paper.blockSignals(False)
        modes = info.supportedDuplexModes()
        two = any(m in modes for m in (QPrinter.DuplexLongSide, QPrinter.DuplexShortSide, QPrinter.DuplexAuto))
        self.duplex.setEnabled(two)
        self.duplex.setToolTip("" if two else i18n.tr("Máy in này không hỗ trợ in 2 mặt tự động"))
        if not two:
            self.duplex.setCurrentIndex(0)
        self._apply_duplex(refresh=False)
        self._apply_layout()

    def _apply_duplex(self, refresh=True):
        from PySide6.QtPrintSupport import QPrinter
        self.printer.setDuplex((QPrinter.DuplexNone, QPrinter.DuplexLongSide,
                                QPrinter.DuplexShortSide)[self.duplex.currentIndex()])
        if refresh:
            self._refresh()

    def _apply_layout(self):
        from PySide6.QtGui import QPageLayout
        pid = self.paper.currentData()
        if pid is not None:
            self.printer.setPageSize(QPageSize(QPageSize.PageSizeId(pid)))
        self.printer.setPageOrientation(QPageLayout.Landscape if self.orient.currentIndex() == 1
                                        else QPageLayout.Portrait)
        self._refresh()

    def _refresh(self):
        self.pct.setEnabled(self.scale.currentIndex() == 3 and self.nup.currentIndex() == 0)
        self.scale.setEnabled(self.nup.currentIndex() == 0)
        self.order.setEnabled(self.nup.currentIndex() > 0)
        self.border.setEnabled(self.nup.currentIndex() > 0)
        self.rng.setEnabled(True)
        self.preview.updatePreview()

    # ----- trang sẽ in -----
    def pages(self):
        """Danh sách trang (0-based) sẽ in; ToolError nếu phạm vi gõ sai."""
        w = self.which.currentIndex()
        if w == 1:
            pages = [self.win.current]
        elif w == 2:
            pages = self.win.selected_pages()
        elif w == 3:
            pages = pdf_tools.parse_pages(self.rng.text(), self.n)
        else:
            pages = list(range(self.n))
        sub = self.subset.currentIndex()
        if sub == 1:
            pages = [p for p in pages if p % 2 == 0]          # trang lẻ (số trang 1, 3, 5...)
        elif sub == 2:
            pages = [p for p in pages if p % 2 == 1]
        if not pages:
            raise pdf_tools.ToolError(i18n.tr("Không có trang nào để in trong phạm vi đã chọn."))
        return pages

    def options(self):
        return dict(gray=self.color.currentIndex() == 1, nup=self.NUPS[self.nup.currentIndex()],
                    down_first=self.order.currentIndex() == 1, scale=self.scale.currentIndex(),
                    pct=self.pct.value(), autorot=self.autorot.isChecked(), center=self.center.isChecked(),
                    border=self.border.isChecked())

    def gray(self):
        return self.color.currentIndex() == 1

    def _paint(self, pr):
        try:
            pages = self.pages()
            err = None
        except pdf_tools.ToolError as e:
            pages, err = [], str(e)
        self.ok.setEnabled(bool(pages) and self.printer_box.count() > 0)
        if not err:
            o = self.options()
            sheets = -(-len(pages) // o["nup"])
            msg = i18n.tr("Sẽ in {0} trang trên {1} tờ").format(len(pages), sheets)
            if self.copies.value() > 1:
                msg += i18n.tr(" x {0} bản").format(self.copies.value())
            if self.duplex.currentIndex() > 0:
                msg += i18n.tr(" (2 mặt: {0} tờ giấy / bản)").format(-(-sheets // 2))
            self.info.setText(msg + ".")
        else:
            self.info.setText(err)
        self.info.setStyleSheet("color:#dc2626" if err else "color:#64748b")
        if pages:                                     # xem trước: dựng ảnh độ phân giải thấp cho nhanh
            self.win._render_pages(pr, pages, 96 if len(pages) <= 60 else 60, self.options())

    def _go(self, d):
        self.preview.setCurrentPage(min(max(1, self.preview.currentPage() + d), self.preview.pageCount()))
        self._nav_update()

    def _nav_update(self):
        n, k = self.preview.pageCount(), self.preview.currentPage()
        with i18n.paused():
            self.page_lbl.setText(f"  {k} / {n}  ")
        self.prev_btn.setEnabled(k > 1)
        self.next_btn.setEnabled(k < n)

    def _windows_dialog(self):
        """Hộp thoại in của Windows (tuỳ chọn riêng của máy in) - In ở đó là in luôn."""
        from PySide6.QtPrintSupport import QPrintDialog, QPrinter
        dlg = QPrintDialog(self.printer, self)
        dlg.setMinMax(1, self.n)
        dlg.setOption(QPrintDialog.PrintCurrentPage, True)
        if dlg.exec() != QDialog.Accepted:
            return
        rng = self.printer.printRange()
        if rng == QPrinter.PageRange:
            pages = list(range(max(1, self.printer.fromPage()) - 1, min(self.n, self.printer.toPage())))
        elif rng == QPrinter.CurrentPage:
            pages = [self.win.current]
        else:
            pages = list(range(self.n))
        self.result_pages = pages
        self.result_copies = 1                        # số bản / theo bộ: máy in tự xử lý theo hộp thoại Windows
        self._save()
        super().accept()

    def accept(self):
        try:
            self.result_pages = self.pages()
        except pdf_tools.ToolError as e:
            QMessageBox.warning(self, "In", str(e))
            return
        n, coll = self.copies.value(), self.collate.currentIndex() == 0
        if self.printer.supportsMultipleCopies():      # máy in tự in nhiều bản / theo bộ
            self.printer.setCopyCount(n)
            self.printer.setCollateCopies(coll)
            self.result_copies = 1
        else:                                         # tự lặp trang
            self.printer.setCopyCount(1)
            self.result_copies = n
        self.result_collate = coll
        self._save()
        super().accept()


# ===========================================================================
#  Cửa sổ chính
# ===========================================================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf = PdfDocument()
        self.current = 0
        self.zoom = 1.0
        self.fit_width = True
        self.mode = "view"
        self.page_objs = []
        self.sel_obj = None
        self.page_links = {}
        self.grid = False
        self.find_dialog = None
        self.settings = QSettings(installer.ORG, "PDFToolkit")
        self.resize(1360, 860)
        self.setAcceptDrops(True)

        # --- thumbnails + canvas ---
        self.thumbs = ThumbList()
        self.thumbs.currentRowChanged.connect(self.on_current_changed)
        self.thumbs.orderChanged.connect(self.on_reordered)
        self.thumbs.filesDropped.connect(self.on_files_dropped)
        self.thumbs.setContextMenuPolicy(Qt.CustomContextMenu)
        self.thumbs.customContextMenuRequested.connect(self.thumb_menu)

        self.canvas = PageCanvas()
        self.canvas.lineClicked.connect(self.on_line_clicked)
        self.canvas.pointClicked.connect(self.on_point_clicked)
        self.canvas.noteClicked.connect(self.on_note_clicked)
        self.canvas.objectMoved.connect(self.on_object_moved)
        self.canvas.objectMenu.connect(self.on_object_menu)
        self.canvas.objectClicked.connect(self.on_object_clicked)
        self.canvas.objectSelected.connect(self.on_object_selected)
        self.canvas.objectActivated.connect(self.on_object_activated)
        self.canvas.strokeDrawn.connect(self.on_stroke_drawn)
        self.canvas.polyDrawn.connect(self.on_poly_drawn)
        self.canvas.selectionMenu.connect(self.on_selection_menu)
        self.snap_region = None                                  # Chụp vùng: (trang, QRectF pt hiển thị)
        self.canvas.regionChanged.connect(self.on_region_changed)
        self.canvas.regionMenu.connect(self.on_region_menu)
        self.canvas.pageMenu.connect(self.on_page_menu)
        self.obj_clip = None                                     # vật đã Cắt / Sao chép trên trang (xem copy_object)
        self.canvas.selectionChanged.connect(
            lambda n: self.status.showMessage(f"Đã chọn {n} ký tự - Ctrl+C để sao chép, chuột phải để đánh dấu", 5000)
            if n else None)
        self.canvas.rectDrawn.connect(self.on_rect_drawn)
        self.canvas.segmentDrawn.connect(self.on_segment_drawn)
        self.strip = PageStrip(self.canvas)           # các trang xếp nối tiếp - cuộn liền mạch
        self.scroll = PageScroll()
        self.scroll.setWidget(self.strip)
        self.scroll.setAlignment(Qt.AlignHCenter)
        self.scroll.setStyleSheet("QScrollArea{background:#e5e7eb; border:none}")
        self.scroll.zoomRequested.connect(self.zoom_step)
        self._auto_scroll = False                       # đang tự cuộn tới trang (không đổi trang hiện tại theo cuộn)
        self.scroll.verticalScrollBar().valueChanged.connect(self._on_scrolled)
        self._fit_timer = QTimer(self)                  # gom nhiều lần đổi cỡ liên tiếp thành 1 lần vẽ lại
        self._fit_timer.setSingleShot(True)
        self._fit_timer.setInterval(40)
        self._fit_timer.timeout.connect(lambda: self.fit_width and self.pdf.doc and not self.grid and self.render_page())
        self.scroll.resized.connect(self._fit_timer.start)

        self.empty = QLabel("Kéo thả file PDF vào đây\nhoặc bấm  Mở  (Ctrl+O)")
        self.empty.setAlignment(Qt.AlignCenter)
        self.empty.setStyleSheet("color:#94a3b8; font-size:20px; background:#f1f5f9")

        self.splitter = QSplitter()
        self.splitter.addWidget(self.thumbs)
        self.splitter.addWidget(self.scroll)
        self.splitter.setStretchFactor(1, 1)
        self.splitter.setSizes([220, 1060])
        # nút thu gọn / mở bảng trang: nổi ở góc trên trái vùng xem trang (luôn thấy, kể cả khi đã thu gọn)
        self.thumbs_btn = QToolButton(self.scroll)
        self.thumbs_btn.setObjectName("thumbsToggle")
        self.thumbs_btn.setIconSize(QSize(14, 14))
        self.thumbs_btn.setCursor(Qt.PointingHandCursor)
        self.thumbs_btn.setFixedSize(26, 30)
        self.thumbs_btn.move(4, 6)
        self.thumbs_btn.clicked.connect(lambda: self.a_thumbs.trigger())   # a_thumbs tạo cùng các action
        self.thumbs_btn.raise_()
        if hasattr(self, "a_thumbs"):
            self._thumbs_btn_icon()
        self.editor_area = QStackedWidget()
        self.editor_area.addWidget(self.empty)
        self.editor_area.addWidget(self.splitter)

        # --- thanh tab tài liệu (mở nhiều PDF cùng lúc, như trình duyệt) ---
        self.docs = []                                   # [DocTab], cùng thứ tự với thanh tab
        self.tabbar = QTabBar()
        self.tabbar.setObjectName("docTabs")
        self.tabbar.setTabsClosable(False)              # nút ✕ riêng (phẳng như trình duyệt) - xem _add_tab
        self.tabbar.setMovable(True)
        self.tabbar.setExpanding(False)
        self.tabbar.setDocumentMode(True)
        self.tabbar.setElideMode(Qt.ElideMiddle)
        self.tabbar.setUsesScrollButtons(True)
        self.tabbar.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabbar.currentChanged.connect(self._activate_tab)
        self.tabbar.tabCloseRequested.connect(self.close_tab)
        self.tabbar.tabMoved.connect(self._tab_moved)
        self.tabbar.customContextMenuRequested.connect(self._tab_menu)
        self.tabbar.installEventFilter(self)             # bấm chuột giữa để đóng tab
        self._active = -1
        new_tab = QToolButton()
        new_tab.setObjectName("newTabBtn")
        new_tab.setIcon(glyph_icon(0xE710))
        new_tab.setToolTip("Mở thêm file PDF trong tab mới  (Ctrl+O)")
        new_tab.setAutoRaise(True)
        new_tab.clicked.connect(self.open_file)
        tabrow = QWidget()
        tabrow.setObjectName("docTabRow")
        tabrow.setAttribute(Qt.WA_StyledBackground, True)
        trl = QHBoxLayout(tabrow)
        trl.setContentsMargins(8, 4, 8, 0)
        trl.setSpacing(4)
        trl.addWidget(self.tabbar, 1)                   # tab xếp từ trái, chỉ hiện nút cuộn khi quá nhiều tab
        trl.addWidget(new_tab, 0, Qt.AlignVCenter)
        self.doc_area = QWidget()
        dal = QVBoxLayout(self.doc_area)
        dal.setContentsMargins(0, 0, 0, 0)
        dal.setSpacing(0)
        dal.addWidget(tabrow)
        dal.addWidget(self.editor_area, 1)

        # --- trang chủ (lưới công cụ) + màn hình công cụ + trình sửa ---
        self.home = tools_ui.HomePage(APP_TITLE, "Bộ công cụ PDF - chạy offline, file không rời khỏi máy")
        self.home.toolChosen.connect(self.open_tool)
        self.home.langChosen.connect(self.switch_language)
        self.home.aboutRequested.connect(self.show_about)
        self.home.resumeRequested.connect(self.resume_doc)
        self.home.supportMail.connect(lambda: about.send_support_mail(APP_VERSION))
        self._build_setup_menu()
        self.tool_page = tools_ui.ToolPage()
        self.tool_page.back.connect(self.show_home)
        self.tool_page.openInEditor.connect(self.open_in_editor)
        self.stack = QStackedWidget()
        for w in (self.home, self.tool_page, self.doc_area):
            self.stack.addWidget(w)
        self.setCentralWidget(self.stack)

        self._build_toolbar()
        self.status = self.statusBar()
        self.status.addPermanentWidget(self.zoom_bar)
        self.show_home()

    # ------------------------------------------------------------------ UI
    def _act(self, text, slot, shortcut=None, icon=None, tip=None, checkable=False, mirror=False):
        a = QAction(text, self)
        if icon is not None:
            a.setIcon(glyph_icon(icon, mirror=mirror))
        if shortcut:
            a.setShortcut(QKeySequence(shortcut))
        a.setToolTip(f"{tip or text}" + (f"  ({QKeySequence(shortcut).toString()})" if shortcut else ""))
        a.setCheckable(checkable)
        if slot:
            a.triggered.connect(slot)
        self.addAction(a)
        return a

    def _build_toolbar(self):
        tb = QToolBar("Ribbon")
        tb.setMovable(False)
        tb.setFloatable(False)
        tb.setStyleSheet("QToolBar{background:transparent; border:none; padding:0; spacing:0}")
        self.addToolBar(tb)
        self.toolbar = tb

        self.a_home = self._act("Trang chủ", self.show_home, "Alt+Home", 0xE80F, "Về trang chủ (tất cả công cụ)")
        self.a_open = self._act("Mở", self.open_file, "Ctrl+O", 0xE8E5)
        self.a_add = self._act("Gộp file", self.add_files, "Ctrl+M", 0xE710, "Chèn thêm PDF/ảnh vào sau trang đang chọn")
        self.a_save = self._act("Lưu", self.save, "Ctrl+S", 0xE74E)
        self.a_saveas = self._act("Lưu thành", self.save_as, "Ctrl+Shift+S", 0xE792)
        self.a_new = self._act("Tạo mới", self.new_document, "Ctrl+N", 0xE8A5, "Tạo tài liệu PDF trống (1 trang A4) trong tab mới")
        self.a_scan = self._act("Máy scan → PDF", lambda: self.scan_to_pdf(False), None, 0xE8FE,
                                "Quét tài liệu giấy từ máy scan thành PDF (mở trong tab mới)")
        self.a_scan_insert = self._act("Chèn trang từ máy scan", lambda: self.scan_to_pdf(True), None, 0xE8FE,
                                       "Quét thêm trang từ máy scan, chèn sau trang đang xem")
        self.a_print = self._act("In", self.print_doc, "Ctrl+P", 0xE749, "In tài liệu")
        self.a_email = self._act("Gửi email", self.email_doc, "Ctrl+Shift+E", 0xE715,
                                 "Mở thư mới (Outlook / ứng dụng email mặc định) đính kèm file PDF đang mở")
        self.a_undo = self._act("Hoàn tác", self.undo, "Ctrl+Z", 0xE7A7)
        self.a_redo = self._act("Làm lại", self.redo, "Ctrl+Y", 0xE7A6)
        self.a_del = self._act("Xoá trang", self.delete_pages, "Delete", 0xE74D, "Xoá các trang đang chọn")
        self.a_up = self._act("Lên", lambda: self.move_pages(-1), "Ctrl+Up", 0xE74A, "Đưa trang đang chọn lên trước")
        self.a_down = self._act("Xuống", lambda: self.move_pages(1), "Ctrl+Down", 0xE74B, "Đưa trang đang chọn xuống sau")
        self.a_rotl = self._act("Xoay trái", lambda: self.rotate(-90), "Ctrl+L", 0xE7AD, mirror=True)
        self.a_rotr = self._act("Xoay phải", lambda: self.rotate(90), "Ctrl+R", 0xE7AD)
        self.a_extract = self._act("Tách trang", self.extract, "Ctrl+E", 0xE8A7, "Lưu các trang đang chọn ra file mới")
        self.a_grid = self._act("Xem dạng lưới", self.toggle_grid, "Ctrl+G", 0xE8A9, "Xem dạng lưới để kéo thả sắp xếp nhiều trang", checkable=True)
        self.a_thumbs = self._act("Bảng trang", self.toggle_thumbs, "F4", 0xE90C,
                                  "Hiện / thu gọn bảng ảnh thu nhỏ các trang bên trái", checkable=True)
        self.a_thumbs.setChecked(True)
        if hasattr(self, "thumbs_btn"):
            self._thumbs_btn_icon()
        self.a_close_tab = self._act("Đóng tab", lambda: self.close_tab(self.tabbar.currentIndex()), "Ctrl+W", 0xE711,
                                     "Đóng file đang xem")
        self.a_next_tab = self._act("Tab sau", lambda: self._cycle_tab(1), "Ctrl+Tab")
        self.a_prev_tab = self._act("Tab trước", lambda: self._cycle_tab(-1), "Ctrl+Shift+Tab")
        grp = QActionGroup(self)
        self.tool_actions = {}

        def tool(key, text, icon, tip, shortcut=None):
            a = self._act(text, lambda: self.set_mode(key), shortcut, icon, tip, checkable=True)
            grp.addAction(a)
            self.tool_actions[key] = a
            return a

        self.a_view = tool("view", "Xem", 0xE7B3, "Chế độ xem", "Esc")
        self.a_edit = tool("edit", "Sửa file", 0xE70F, "Sửa mọi thành phần trong file: bấm chọn khung chữ / ảnh / chữ ký / "
                           "con dấu / khối vẽ, giữ chuột kéo hoặc phím mũi tên để di chuyển, bấm đúp để sửa chữ, "
                           "căn chỉnh, Delete để xoá", "Ctrl+T")
        self.a_addtext = tool("text", "Thêm chữ", 0xE8D2, "Kéo khung (hoặc bấm) trên trang để thêm chữ có định dạng", "Ctrl+Shift+T")
        self.a_find = self._act("Tìm && thay thế", self.find_replace, "Ctrl+H", 0xE721, "Tìm và thay thế chữ")
        tool("image", "Ảnh", 0xEB9F, "Chèn ảnh: kéo vùng đặt ảnh", "Ctrl+Shift+I")
        tool("link", "Liên kết", 0xE71B, "Chèn liên kết web / email / trang", "Ctrl+K")
        tool("highlight", "Tô sáng chữ", 0xE7E6, "Kéo qua đoạn chữ cần tô sáng")
        tool("note", "Ghi chú", 0xE90A, "Bấm vào vị trí để đặt ghi chú")
        tool("rect", "Hình chữ nhật", 0xE739, "Kéo để vẽ hình chữ nhật (hoặc che trắng)")
        tool("redact", "Đánh dấu xoá", 0xE8C6, "Kéo vùng cần xoá vĩnh viễn (chữ, ảnh, hình bên dưới) - xem lại rồi bấm "
             "Xoá vĩnh viễn")
        self.a_detect_fields = self._act("Tự tạo ô form", self.detect_fields, None, 0xE8A5,
                                         "Tự nhận diện chỗ trống cần điền (gạch dưới, chấm chấm, đường kẻ, ô vuông, "
                                         "khung trống) và tạo ô form trên mọi trang")
        self.a_redact_find = self._act("Tìm && đánh dấu...", self.redact_find, None, 0xE721,
                                       "Tìm chữ / số điện thoại / email / CCCD / số tài khoản để đánh dấu xoá")
        self.a_redact_apply = self._act("Xoá vĩnh viễn", self.redact_apply, None, 0xE74D,
                                        "Xoá hẳn mọi thứ trong các vùng đã đánh dấu - không khôi phục được sau khi lưu")
        self.a_redact_clear = self._act("Bỏ đánh dấu xoá", self.redact_clear, None, 0xE711,
                                        "Bỏ mọi vùng đánh dấu xoá chưa áp dụng")
        tool("ellipse", "Hình elip", 0xEA3A, "Kéo để vẽ hình elip")
        tool("line", "Đường thẳng", 0xE738, "Kéo để vẽ đường thẳng")
        tool("arrow", "Mũi tên", 0xF0AF, "Kéo để vẽ mũi tên")
        # --- nhận xét (Comment) ---
        tool("select", "Chọn chữ", 0xE8B0, "Kéo qua chữ để chọn - Ctrl+C sao chép, chuột phải để tô sáng / gạch chân...",
             "Ctrl+Shift+A")
        tool("snapshot", "Chụp vùng", 0xE722, "Kéo vùng trên trang để chụp thành ảnh (sao chép vào clipboard)",
             "Ctrl+Shift+C")
        tool("underline", "Gạch chân", 0xE8DC, "Kéo qua đoạn chữ cần gạch chân")
        tool("strikeout", "Gạch ngang", 0xEDE0, "Kéo qua đoạn chữ cần gạch ngang (đề xuất xoá)")
        tool("caret", "Đề xuất chèn chữ", 0xE8C5, "Bấm vào vị trí cần chèn chữ - nhập nội dung đề xuất")
        tool("textbox", "Khung chữ", 0xE8E9, "Kéo khung (hoặc bấm) để thêm hộp chữ có viền / nền")
        tool("typewriter", "Gõ chữ nhanh", 0xE765, "Bấm vào chỗ cần gõ chữ (không viền, không nền) - vd điền vào form giấy")
        tool("callout", "Chú thích có mũi tên", 0xE8BD, "Kéo từ điểm cần chỉ tới vị trí đặt hộp chữ")
        tool("pencil", "Bút vẽ", 0xEE56, "Kéo chuột để vẽ / ký tự do")
        tool("marker", "Bút dạ quang", 0xE7E6, "Tô dạ quang tự do (vùng không phải chữ)")
        tool("eraser", "Tẩy", 0xE75C, "Bấm vào ghi chú / nét vẽ / hình để xoá")
        tool("polygon", "Đa giác", 0xF158, "Bấm từng đỉnh, bấm đúp / Enter / chuột phải để kết thúc")
        tool("polyline", "Đường gấp khúc", 0xE7FD, "Bấm từng điểm, bấm đúp / Enter / chuột phải để kết thúc")
        tool("cloud", "Khoanh đám mây", 0xE753, "Kéo để vẽ khung đám mây (khoanh vùng cần sửa)")
        tool("attach", "Đính kèm file", 0xE723, "Bấm vào vị trí để ghim file (Excel, Word, ảnh...) lên trang")
        self.a_pen_style = self._act("Màu && độ dày nét...", self.pen_style, None, 0xE790, "Đổi màu / độ dày bút vẽ")
        self.a_comments = self._act("Danh sách ghi chú", self.show_comments, "Ctrl+Shift+L", 0xE8FD,
                                    "Xem tất cả ghi chú trong file, tới trang, xoá, xuất bản tổng hợp")
        self.a_copy = self._act("Sao chép chữ", self.copy_selection, "Ctrl+C", 0xE8C8)
        self.a_cut = self._act("Cắt", self.cut_object, "Ctrl+X", 0xE8C6, "Cắt vật đang chọn trên trang")
        self.a_paste = self._act("Dán", lambda: self.paste_object(), "Ctrl+V", 0xE77F,
                                 "Dán vật đã cắt / sao chép, hoặc ảnh / chữ từ bộ nhớ tạm, vào trang")
        self.a_selall = self._act("Chọn toàn bộ chữ trong trang", self.select_all_text, "Ctrl+A")
        self.a_autolinks = self._act("Tạo liên kết từ địa chỉ web", self.auto_links, None, 0xE71B,
                                     "Biến các địa chỉ web / email dạng chữ trong file thành liên kết bấm được")
        self.a_rmlinks = self._act("Xoá tất cả liên kết web", self.remove_web_links, None, 0xE8D8,
                                   "Xoá mọi liên kết tới trang web / email (giữ liên kết trong file)")
        self.a_dup = self._act("Nhân bản trang", self.duplicate_pages, "Ctrl+D", 0xE8C8,
                               "Tạo bản sao các trang đang chọn (ngay sau trang gốc)")
        self.a_swap = self._act("Đổi chỗ 2 trang", self.swap_pages, None, 0xE8AB,
                                "Đổi chỗ 2 trang đang chọn (chọn 2 trang ở danh sách bên trái)")
        self.a_replace = self._act("Thay trang...", self.replace_pages, None, 0xE8B5,
                                   "Thay các trang đang chọn bằng trang lấy từ file PDF khác")
        self.a_finddup = self._act("Tìm trang trùng", self.find_duplicates, None, 0xE721,
                                   "Tìm các trang giống hệt nhau (vd scan 2 lần) và chọn sẵn để xoá")
        tool("field", "Ô form", 0xE8A5, "Kéo vùng để tạo ô điền form (ô chữ, ô đánh dấu, danh sách chọn)")
        tool("handsign", "Chữ ký tay", 0xEC87, "Kéo vùng (hoặc bấm) để đặt chữ ký tay - vẽ bằng chuột hoặc dùng ảnh")
        tool("stamp", "Con dấu", 0xE8EC, "Chèn con dấu / stamp (ĐÃ DUYỆT, APPROVED, ảnh con dấu scan...) - "
             "bấm vào trang để đặt, kéo để di chuyển")
        self.a_sign = tool("sign", "Ký số (kéo vùng đặt chữ ký)...", 0xEE56, "Kéo vùng trên trang để đặt chữ ký số hiển thị")
        self.a_sign_hidden = self._act("Ký số không hiển thị...", self.sign_invisible, None, 0xE72E)
        self.a_verify = self._act("Kiểm tra chữ ký trong file...", self.verify_signatures, None, 0xEB95)
        self.a_view.setChecked(True)
        self.tool_actions["select"].setIcon(self._caret_icon())     # dấu nháy văn bản thay cho icon chuột
        # căn chỉnh vật đang chọn (bấm vào ảnh / chữ ký / con dấu / khối vẽ trên trang rồi bấm nút)
        self.align_acts = {}
        for how, text, tip in (("left", "Căn trái", "Căn vật đang chọn sát lề trái trang"),
                               ("hcenter", "Căn giữa", "Căn vật đang chọn vào giữa trang theo chiều ngang"),
                               ("right", "Căn phải", "Căn vật đang chọn sát lề phải trang"),
                               ("top", "Căn trên", "Căn vật đang chọn sát lề trên trang"),
                               ("vcenter", "Giữa dọc", "Căn vật đang chọn vào giữa trang theo chiều dọc"),
                               ("bottom", "Căn dưới", "Căn vật đang chọn sát lề dưới trang"),
                               ("center", "Giữa trang", "Đặt vật đang chọn vào chính giữa trang")):
            a = QAction(self._align_icon(how), text, self)
            a.setToolTip(i18n.tr(tip) + "\n" + i18n.tr("Bấm chọn ảnh / chữ ký / con dấu / khối vẽ trên trang trước"))
            a.triggered.connect(lambda _=False, h=how: self.align_selected(h))
            self.align_acts[how] = a
        self.a_align_margin = QAction("Chừa lề 1 cm", self)
        self.a_align_margin.setCheckable(True)
        self.a_align_margin.setChecked(self.settings.value("align_margin", True, type=bool))
        self.a_align_margin.setToolTip("Khi căn sát mép: chừa lề 1 cm (bỏ tích để căn sát mép trang)")
        self.a_align_margin.toggled.connect(lambda on: self.settings.setValue("align_margin", on))

        self.a_zout =self._act("Thu nhỏ", lambda: self.zoom_step(-1), "Ctrl+-", 0xE71F)
        self.a_zin = self._act("Phóng to", lambda: self.zoom_step(1), "Ctrl+=", 0xE8A3)
        self.a_fit = self._act("Vừa khung", self.zoom_fit, "Ctrl+0", 0xE9A6)
        self.a_prev = self._act("Trang trước", lambda: self.goto(self.current - 1), "PgUp", 0xE70E)
        self.a_next = self._act("Trang sau", lambda: self.goto(self.current + 1), "PgDown", 0xE70D)
        self.a_first = self._act("Trang đầu", lambda: self.goto(0), "Ctrl+Home")
        self.a_last = self._act("Trang cuối", lambda: self.goto(self.pdf.page_count - 1), "Ctrl+End")
        self.a_view.setText("Xem")
        # phím tắt phụ giống Office: F12 = Lưu thành, Ctrl+F4 = đóng, Ctrl++ = phóng to, Ctrl+Shift+Z = làm lại
        for a, keys in ((self.a_saveas, ("Ctrl+Shift+S", "F12")), (self.a_close_tab, ("Ctrl+W", "Ctrl+F4")),
                        (self.a_zin, ("Ctrl+=", "Ctrl++")), (self.a_redo, ("Ctrl+Y", "Ctrl+Shift+Z")),
                        (self.a_print, ("Ctrl+P", "Ctrl+Shift+F12"))):
            a.setShortcuts([QKeySequence(k) for k in keys])

        # --- công cụ của bộ pdf_tools chạy trên file đang mở ---
        self.tool_acts = {}

        def T(tid, text=None):
            if tid not in self.tool_acts:
                t = pdf_tools.get_tool(tid)
                a = QAction(glyph_icon(t.icon), text or t.name, self)
                a.setToolTip(f"{t.name}: {t.desc}")
                a.triggered.connect(lambda _=False, tid=tid: self.run_doc_tool(tid))
                self.tool_acts[tid] = a
            return self.tool_acts[tid]

        def P(tid, text):
            """Tạo PDF từ file khác / so sánh: chọn file ngay trong trình sửa (không rời màn hình đang xem)."""
            t = pdf_tools.get_tool(tid)
            a = QAction(glyph_icon(t.icon), text, self)
            a.setToolTip(f"{t.name}: {t.desc}")
            a.triggered.connect(lambda _=False, tid=tid: self.create_from_files(tid))
            return a

        ta = self.tool_actions
        self.a_compress = T("compress", "Nén PDF")
        rb = tools_ui.Ribbon()
        self.ribbon = rb

        tab = rb.tab("Trang chủ", "#2563eb")
        g = tab.group("Tệp")
        g.big(self.a_open)
        g.small([self.a_save, self.a_saveas, self.a_add])
        g.small([self.a_new, self.a_print, self.a_email])
        g = tab.group("Chế độ", "#0891b2", limit=4)
        g.big(self.a_view)
        g.big(ta["select"], "Chọn chữ")
        g.big(self.a_edit)
        g.big(self.a_addtext)
        self._align_group(tab)
        g = tab.group("Chèn nhanh", "#d97706")
        g.small([ta["image"], ta["handsign"], ta["stamp"]])
        g.small([ta["link"], ta["highlight"], ta["note"]])
        g.small([ta["rect"], self.a_find])
        g = tab.group("Trang")
        g.big(self.a_grid, "Xem dạng lưới")
        g.small([self.a_del, self.a_rotl, self.a_rotr])
        g = tab.group("Công cụ", "#059669")
        g.big(T("ocr"), "OCR")
        g.big(T("ocr_editable"), "OCR sửa được")
        g.big(self.a_compress, "Nén PDF")
        g.big(T("watermark"), "Watermark")
        g.big(T("page_numbers"), "Số trang")
        g = tab.group("Ký số", "#dc2626")
        g.big(self.a_sign, "Ký số")

        tab = rb.tab("Chỉnh sửa", "#0891b2")
        g = tab.group("Chữ")
        g.big(self.a_edit)
        g.big(self.a_addtext)
        g.big(self.a_find, "Tìm && thay thế")
        self._align_group(tab)
        g = tab.group("Đánh dấu")
        g.big(ta["highlight"], "Tô sáng")
        g.big(ta["note"], "Ghi chú")
        g = tab.group("Hình vẽ")
        g.small([ta["rect"], ta["ellipse"]])
        g.small([ta["line"], ta["arrow"]])
        g = tab.group("Ghi chú & form")
        g.big(ta["field"], "Tạo ô form")
        g.big(self.a_detect_fields, "Tự tạo ô form")
        g.small([T("remove_annots", "Xoá ghi chú"), T("flatten", "Cố định ghi chú"), T("form_fill", "Điền form")])

        tab = rb.tab("Chèn", "#d97706")
        g = tab.group("Nội dung")
        g.big(ta["image"], "Ảnh")
        g.big(self.a_addtext, "Thêm chữ")
        g = tab.group("Liên kết")
        g.big(ta["link"], "Liên kết")
        g.small([self.a_autolinks, self.a_rmlinks])
        g = tab.group("Trang")
        g.big(T("blank_page", "Trang trắng"))
        g.big(self.a_add, "Chèn file")
        g.big(self.a_scan_insert, "Trang quét")
        g = tab.group("Đóng dấu & đánh số")
        g.big(T("watermark"), "Watermark")
        g.small([T("header_footer", "Đầu / chân trang"), T("page_numbers", "Đánh số trang"), T("bates", "Số Bates")])
        g.small([T("stamp", "Đóng dấu nhiều trang")])
        g = tab.group("Chữ ký & con dấu")
        g.big(ta["handsign"], "Chữ ký tay")
        g.big(ta["stamp"], "Con dấu")
        g.big(ta["field"], "Ô form")
        g = tab.group("Khác")
        g.small([T("bg_color", "Màu nền trang"), T("toc_page", "Trang mục lục"), T("add_attach", "Đính kèm file")])
        g.small([T("edit_attach", "Sửa / xoá đính kèm"), T("page_labels", "Nhãn trang")])

        tab = rb.tab("Trang", "#2563eb")
        g = tab.group("Sắp xếp")
        g.big(self.a_grid, "Xem dạng lưới")
        g.small([self.a_up, self.a_down, T("reverse", "Đảo thứ tự")])
        g.small([self.a_dup, self.a_swap, self.a_replace])
        g.small([self.a_finddup, T("blank_page", "Trang trắng"), self.a_add])
        g = tab.group("Xoay & xoá")
        g.small([self.a_rotl, self.a_rotr, self.a_del])
        g.small([T("rotate_custom", "Xoay góc tuỳ ý"), T("deskew", "Làm thẳng bản scan"),
                 T("remove_blank", "Xoá trang trắng")])
        g = tab.group("Tách")
        g.big(self.a_extract, "Tách trang chọn")
        g.small([T("split", "Tách file..."), T("extract_pages", "Trích trang...")])
        g = tab.group("Bố cục")
        g.small([T("crop", "Cắt lề"), T("margins", "Thêm lề"), T("fix_size", "Đồng bộ khổ giấy")])
        g.small([T("divide", "Chia đôi trang")])
        g.small([T("nup", "Nhiều trang / 1 tờ"), T("booklet", "Sách gập"), T("poster", "In poster")])
        g.small([T("duplex_collate", "Sắp bản scan 2 mặt"), T("combine_single_page", "Ghép 1 trang dài"),
                 T("page_dimensions", "Kích thước trang")])
        g = tab.group("Điều hướng")
        g.small([T("bookmarks", "Bookmark"), T("toc_page", "Trang mục lục"), T("overlay", "Lồng mẫu nền")])
        g.small([T("page_labels", "Nhãn trang")])

        tab = rb.tab("Chuyển đổi", "#7c3aed")
        g = tab.group("Xuất ra")
        g.big(T("pdf_to_word"), "Word")
        g.big(T("pdf_to_excel"), "Excel")
        g.big(T("pdf_to_ppt"), "PowerPoint")
        g.big(T("pdf_to_images"), "Ảnh")
        g.big(T("pdf_to_text"), "Văn bản")
        g.small([T("pdf_to_svg", "SVG"), T("extract_images", "Lấy ảnh gốc"), T("extract_attach", "Xuất file đính kèm")])
        g = tab.group("Màu sắc")
        g.small([T("greyscale", "Đen trắng"), T("invert", "Đảo màu"), T("scan_effect", "Hiệu ứng scan")])
        g.small([T("adjust_colors", "Sáng / tương phản"), T("text_color", "Đổi màu chữ"),
                 T("rasterize", "Chuyển thành ảnh")])
        g = tab.group("Tạo PDF từ file khác")
        g.small([P("images_to_pdf", "Ảnh → PDF"), P("office_to_pdf", "Word / Excel / PPT → PDF"),
                 P("text_to_pdf", "Văn bản / Markdown → PDF")])
        g.small([P("ebook_to_pdf", "Ebook / XPS → PDF"), P("email_to_pdf", "Email → PDF"), self.a_scan])

        tab = rb.tab("OCR && Tối ưu", "#059669")
        g = tab.group("Nhận dạng chữ")
        g.big(T("ocr"), "OCR bản scan")
        g.big(T("ocr_editable"), "OCR → PDF sửa được")
        g = tab.group("Xử lý bản scan")
        g.big(T("deskew"), "Làm thẳng")
        g.big(T("adjust_colors"), "Sáng / tương phản")
        g = tab.group("Tối ưu")
        g.big(self.a_compress, "Nén PDF")
        g.big(T("repair"), "Sửa file lỗi")
        g = tab.group("Thông tin & dữ liệu ẩn")
        g.small([T("metadata", "Thông tin file"), T("sanitize", "Làm sạch dữ liệu ẩn"), T("remove_annots", "Xoá ghi chú")])

        tab = rb.tab("Bảo mật && Ký số", "#dc2626")
        g = tab.group("Ký số", "#dc2626")
        g.big(self.a_sign, "Ký số")
        g.small([self.a_sign_hidden, self.a_verify, ta["handsign"]])
        g.small([ta["stamp"]])
        g = tab.group("Dấu thời gian")
        g.big(T("timestamp"), "Dấu thời gian")
        g = tab.group("Xoá thông tin", "#dc2626")
        g.big(ta["redact"], "Đánh dấu xoá")
        g.big(self.a_redact_find, "Tìm && đánh dấu")
        g.big(self.a_redact_apply, "Xoá vĩnh viễn")
        g.small([self.a_redact_clear, T("redact", "Xoá thông tin nhiều file")])
        g = tab.group("Mật khẩu")
        g.big(T("protect"), "Đặt mật khẩu")
        g.big(T("restrict"), "Chặn in / sao chép")
        g.big(T("unlock"), "Mở khoá")
        g = tab.group("Form")
        g.big(self.a_detect_fields, "Tự tạo ô form")
        g.big(T("form_fill"), "Điền form")
        g.big(T("flatten"), "Cố định form")
        g = tab.group("Chống sao chép")
        g.small([T("font_to_outline", "Chữ → nét vẽ"), T("rasterize", "Chuyển thành ảnh"), T("layers", "Lớp nội dung")])
        g = tab.group("So sánh")
        g.big(P("compare", "So sánh 2 file"))

        tab = rb.tab("Xem PDF", "#475569")
        g = tab.group("Thu phóng")
        g.big(self.a_zin)
        g.big(self.a_zout)
        g.big(self.a_fit)
        g = tab.group("Trang")
        g.big(self.a_prev)
        g.big(self.a_next)
        g.big(self.a_grid, "Xem dạng lưới")
        g.big(self.a_thumbs, "Bảng trang")
        g = tab.group("Ứng dụng")
        g.big(self.a_home, "Tất cả công cụ")

        tab = rb.tab("Khác", "#d97706")                     # công cụ ghi chú / đánh dấu / vẽ
        g = tab.group("Công cụ")
        g.big(ta["select"], "Chọn chữ")
        g.big(ta["snapshot"], "Chụp vùng")
        g.big(self.a_comments, "Danh sách ghi chú")
        g = tab.group("Chữ")
        g.big(ta["textbox"], "Khung chữ")
        g.small([ta["typewriter"], ta["callout"], ta["note"]])
        g = tab.group("Đánh dấu chữ")
        g.small([ta["highlight"], ta["underline"], ta["strikeout"]])
        g.small([ta["caret"], ta["attach"]])
        g = tab.group("Vẽ")
        g.small([ta["line"], ta["arrow"], ta["rect"]])
        g.small([ta["ellipse"], ta["polygon"], ta["polyline"]])
        g.small([ta["cloud"], ta["pencil"], ta["marker"]])
        g.small([ta["eraser"], self.a_pen_style])
        g = tab.group("Con dấu")
        g.big(ta["stamp"], "Con dấu")
        g = tab.group("Quản lý")
        g.small([self.a_comments, T("flatten", "Cố định ghi chú"), T("remove_annots", "Xoá ghi chú")])

        # Thanh truy cập nhanh (góc trái, như Word)
        quick = QWidget()
        quick.setObjectName("quickBar")
        ql = QHBoxLayout(quick)
        ql.setContentsMargins(4, 0, 10, 0)
        ql.setSpacing(0)
        for a in (self.a_home, self.a_open, self.a_save, self.a_print, self.a_email, self.a_undo, self.a_redo,
                  None, self.a_view, ta["select"]):         # + 2 chế độ hay dùng: Xem, Chọn chữ
            if a is None:
                from PySide6.QtWidgets import QFrame
                sep = QFrame()
                sep.setFrameShape(QFrame.VLine)
                sep.setFixedSize(9, 18)
                sep.setStyleSheet("color:#d4d4d8")
                ql.addWidget(sep)
                continue
            b = QToolButton()
            b.setDefaultAction(a)
            b.setAutoRaise(True)
            b.setIconSize(QSize(18, 18))
            ql.addWidget(b)
        rb.setCornerWidget(quick, Qt.TopLeftCorner)
        host = QWidget()
        host.setObjectName("ribbonHost")
        host.setAttribute(Qt.WA_StyledBackground, True)
        hl = QVBoxLayout(host)
        hl.setContentsMargins(0, 0, 0, 0)
        hl.addWidget(rb)
        tb.addWidget(host)

        # Zoom ở góc phải thanh trạng thái
        self.zoom_bar = QWidget()
        zl = QHBoxLayout(self.zoom_bar)
        zl.setContentsMargins(0, 0, 6, 0)
        zl.setSpacing(2)
        for a in (self.a_zout, self.a_fit, self.a_zin):
            b = QToolButton()
            b.setDefaultAction(a)
            b.setIconSize(QSize(16, 16))
            b.setAutoRaise(True)
            zl.addWidget(b)
        self._act("", self.find_replace, "Ctrl+F")

        self.doc_actions = [self.a_add, self.a_save, self.a_saveas, self.a_del, self.a_up, self.a_down, self.a_rotl,
                            self.a_rotr, self.a_extract, self.a_grid, self.a_compress, self.a_find, self.a_zout,
                            self.a_zin, self.a_fit, self.a_sign_hidden, self.a_verify, self.a_prev, self.a_next,
                            self.a_close_tab, self.a_next_tab, self.a_prev_tab, self.a_print, self.a_email,
                            self.a_first, self.a_last, self.a_comments, self.a_autolinks, self.a_rmlinks,
                            self.a_dup, self.a_swap, self.a_replace, self.a_finddup, self.a_copy, self.a_selall,
                            self.a_cut, self.a_paste, self.a_redact_find, self.a_redact_apply, self.a_redact_clear,
                            self.a_scan_insert, self.a_detect_fields]
        self.doc_actions += list(self.tool_actions.values()) + list(self.tool_acts.values())

    TOOL_HINTS = {
        "view": ("view", ""),
        "edit": ("edit", "SỬA FILE: bấm để chọn khung chữ / ảnh / chữ ký / con dấu / khối vẽ - giữ chuột kéo hoặc phím "
                         "mũi tên để di chuyển (Shift: 10 pt), bấm đúp / Enter để sửa chữ, Delete để xoá, nhóm Căn chỉnh "
                         "để căn"),
        "text": ("rect", "THÊM CHỮ: kéo khung (hoặc bấm) tại vị trí cần thêm chữ"),
        "image": ("rect", "CHÈN ẢNH: kéo vùng đặt ảnh (hoặc bấm để giữ kích thước gốc) - kéo ảnh có sẵn để di chuyển"),
        "link": ("rect", "LIÊN KẾT: kéo vùng trên chữ / ảnh có sẵn, hoặc bấm để thêm chữ liên kết mới - "
                         "kéo liên kết đã chèn để di chuyển, chuột phải để sửa / xoá"),
        "highlight": ("rect", "TÔ SÁNG: kéo qua đoạn chữ cần tô"),
        "note": ("point", "GHI CHÚ: bấm vào vị trí đặt ghi chú"),
        "rect": ("rect", "HÌNH CHỮ NHẬT: kéo để vẽ - kéo viền hình đã vẽ để di chuyển, ô góc dưới phải để đổi cỡ"),
        "redact": ("rect", "ĐÁNH DẤU XOÁ: kéo vùng cần xoá vĩnh viễn (khung đỏ) - xem lại rồi bấm Xoá vĩnh viễn"),
        "ellipse": ("rect", "HÌNH ELIP: kéo để vẽ - kéo viền hình đã vẽ để di chuyển, ô góc dưới phải để đổi cỡ"),
        "line": ("line", "ĐƯỜNG THẲNG: kéo để vẽ - kéo viền khung đường đã vẽ để di chuyển, ô góc để đổi cỡ"),
        "arrow": ("line", "MŨI TÊN: kéo từ gốc đến đầu mũi tên - kéo viền khung mũi tên đã vẽ để di chuyển"),
        "select": ("select", "CHỌN CHỮ: kéo qua chữ (bấm đúp = cả từ, Shift+bấm = nới vùng chọn, Ctrl+A = cả trang) - Ctrl+C sao chép, "
                             "chuột phải để tô sáng / gạch chân / gạch ngang / tạo liên kết"),
        "snapshot": ("rect", "CHỤP VÙNG: kéo vùng cần chụp - ảnh vào bộ nhớ tạm; kéo khung / ô vuông để chỉnh, "
                             "chuột phải trong khung để lưu ảnh, in, OCR, cắt trang..."),
        "underline": ("rect", "GẠCH CHÂN: kéo qua đoạn chữ"),
        "strikeout": ("rect", "GẠCH NGANG: kéo qua đoạn chữ cần đề xuất xoá"),
        "caret": ("point", "ĐỀ XUẤT CHÈN CHỮ: bấm vào chỗ cần chèn rồi nhập nội dung - dấu ^ đánh dấu vị trí, "
                           "nội dung xem trong Danh sách ghi chú (Acrobat: rê chuột lên dấu ^)"),
        "textbox": ("rect", "HỘP CHỮ: kéo khung (hoặc bấm) để đặt hộp chữ - kéo viền hộp đã đặt để di chuyển"),
        "typewriter": ("rect", "MÁY CHỮ: bấm vào chỗ cần gõ chữ"),
        "callout": ("line", "CHÚ THÍCH MŨI TÊN: kéo từ điểm cần chỉ tới chỗ đặt hộp chữ"),
        "pencil": ("free", "BÚT VẼ: kéo chuột để vẽ tự do (Màu & nét bút ở tab Khác)"),
        "marker": ("free", "BÚT DẠ QUANG: kéo chuột để tô"),
        "eraser": ("point", "TẨY: bấm vào ghi chú / nét vẽ / hình cần xoá"),
        "polygon": ("poly", "ĐA GIÁC: bấm từng đỉnh - bấm đúp / Enter / chuột phải để kết thúc, Backspace bỏ điểm cuối"),
        "polyline": ("poly", "ĐƯỜNG GẤP KHÚC: bấm từng điểm - bấm đúp / Enter / chuột phải để kết thúc"),
        "cloud": ("rect", "ĐÁM MÂY: kéo để khoanh vùng"),
        "attach": ("point", "ĐÍNH KÈM FILE: bấm vào vị trí đặt biểu tượng file"),
        "sign": ("rect", "KÝ SỐ: kéo vùng đặt chữ ký (Esc để huỷ)"),
        "field": ("rect", "Ô FORM: kéo vùng đặt ô điền"),
        "handsign": ("rect", "CHỮ KÝ TAY: kéo vùng đặt chữ ký (hoặc bấm để dùng cỡ mặc định) - kéo chữ ký đã đặt để di chuyển"),
        "stamp": ("rect", "CON DẤU: bấm vào trang để đặt (hoặc kéo khung) - kéo con dấu đã đặt để di chuyển, góc dưới phải để đổi cỡ"),
    }

    def update_ui(self):
        # phím tắt của trình sửa (Delete, Ctrl+Z...) chỉ hoạt động khi đang ở màn hình sửa
        has = self.pdf.doc is not None and self.in_editor()
        for a in self.doc_actions:
            a.setEnabled(has)
        self.a_align_margin.setEnabled(has)
        self._update_align_ui()
        self.a_undo.setEnabled(has and self.pdf.can_undo)
        self.a_redo.setEnabled(has and self.pdf.can_redo)
        self._update_tabs()
        if not self.in_editor():
            page = self.stack.currentWidget()
            self.setWindowTitle(APP_TITLE if page is self.home else f"{self.tool_page.tool.name} — {APP_TITLE}")
            n_tools = len(pdf_tools.TOOLS) + sum(1 for c in tools_ui.EDITOR_CARDS if c[0].startswith("editor:"))
            self.status.showMessage(f"{n_tools} công cụ  |  Chạy offline - file không rời khỏi máy"
                                    if page is self.home else "Thêm file, chọn tuỳ chọn rồi bấm Chạy")
            return
        name = os.path.basename(self.pdf.path) if self.pdf.path else ""
        self.setWindowTitle(f"{'* ' if self.pdf.modified else ''}{name + ' — ' if name else ''}{APP_TITLE}")
        if has:
            sel = len(self.selected_pages())
            hint = self.TOOL_HINTS[self.mode][1]
            mode = f"  |  {hint}" if hint else ""
            self.status.showMessage(f"Trang {self.current + 1}/{self.pdf.page_count}"
                                    + (f"  |  Đang chọn {sel} trang" if sel > 1 else "")
                                    + f"  |  Zoom {self.zoom * 100:.0f}%  |  File: {fmt_size(self.pdf.file_size)}{mode}")
        else:
            self.status.showMessage("Chưa mở file")

    def show_error(self, msg):
        QMessageBox.warning(self, APP_NAME, str(msg))

    # ------------------------------------------------------------------ render
    def make_thumb(self, i):
        page = self.pdf.doc[i]
        pix = page.get_pixmap(matrix=pymupdf.Matrix(THUMB_SCALE, THUMB_SCALE), alpha=False)
        img = pix_to_qimage(pix)
        out = QPixmap(img.width() + 2, img.height() + 2)
        out.fill(QColor("#cbd5e1"))
        p = QPainter(out)
        p.drawImage(1, 1, img)
        label = str(i + 1)
        f = QFont("Segoe UI", 9, QFont.Bold)
        p.setFont(f)
        w = p.fontMetrics().horizontalAdvance(label) + 14
        r = QRectF((out.width() - w) / 2, out.height() - 24, w, 18)
        p.setRenderHint(QPainter.Antialiasing)
        p.setBrush(QColor(15, 23, 42, 200)); p.setPen(Qt.NoPen)
        p.drawRoundedRect(r, 9, 9)
        p.setPen(QColor("white")); p.drawText(r, Qt.AlignCenter, label)
        p.end()
        icon = QIcon(out)
        icon.addPixmap(out, QIcon.Selected)          # không để Qt phủ màu xanh lên trang đang chọn
        return icon

    def refresh_thumbs(self, select=None):
        self.strip.invalidate()                         # số trang / thứ tự / cỡ trang có thể đã đổi
        self.thumbs.blockSignals(True)
        self.thumbs.clear()
        ics = self.thumbs.iconSize()
        self.thumbs.setUniformItemSizes(self.grid)       # dạng danh sách: mỗi ô cao theo khổ trang (trang ngang
                                                         # không để trống nửa ô bên dưới)
        for i in range(self.pdf.page_count):
            it = QListWidgetItem(self.make_thumb(i), "")
            it.setData(Qt.UserRole, i)
            h = ics.height()
            if not self.grid:
                r = self.pdf.doc[i].rect
                h = min(ics.height(), round(r.height * min(ics.width() / r.width, ics.height() / r.height)))
            it.setSizeHint(QSize(ics.width() + 12, h + 12))
            self.thumbs.addItem(it)
            if i % 15 == 14:
                QApplication.processEvents()
        self.thumbs.blockSignals(False)
        sel = select if select is not None else [min(self.current, self.pdf.page_count - 1)]
        self.select_pages(sel)

    def refresh_thumb(self, i):
        self.strip.invalidate(i)
        it = self.thumbs.item(i)
        if it:
            it.setIcon(self.make_thumb(i))

    def select_pages(self, pages):
        pages = [p for p in pages if 0 <= p < self.pdf.page_count] or [0]
        self.thumbs.blockSignals(True)
        self.thumbs.clearSelection()
        self.thumbs.setCurrentRow(pages[0])
        for p in pages:
            self.thumbs.item(p).setSelected(True)
        self.thumbs.blockSignals(False)
        self.thumbs.scrollToItem(self.thumbs.item(pages[0]))
        self.current = pages[0]
        self.render_page()
        self._scroll_to_page(self.current, only_if_hidden=True)

    def selected_pages(self):
        rows = sorted(self.thumbs.row(i) for i in self.thumbs.selectedItems())
        return rows or ([self.current] if self.pdf.doc else [])

    def render_page(self):
        if not self.pdf.doc or self.grid:
            self.update_ui()
            return
        self.current = max(0, min(self.current, self.pdf.page_count - 1))
        page = self.pdf.doc[self.current]
        st = self.strip
        st.sync(self.pdf.doc)
        if self.fit_width:                              # vừa khung theo trang rộng nhất -> mọi trang cùng 1 zoom
            avail = max(200, self.scroll.viewport().width() - 60)
            self.zoom = max(0.2, min(4.0, avail / st.max_width()))
        # giữ nguyên chỗ đang đọc khi đổi zoom: nhớ vị trí tương đối trong trang hiện tại
        sb = self.scroll.verticalScrollBar()
        frac = None
        if st.sizes and self.current < len(st.sizes) and st.zoom and st.zoom != self.zoom:
            r = st.page_rect(self.current)
            frac = (sb.value() - r.top()) / max(1, r.height())
        changed = st.layout_pages(self.zoom)
        self.canvas.show_page(page, self.zoom, self.TOOL_HINTS[self.mode][0])
        st.place_canvas(self.current)
        if changed and frac is not None:
            r = st.page_rect(self.current)
            self._auto_scroll = True
            sb.setValue(int(r.top() + frac * r.height()))
            self._auto_scroll = False
        # vật kéo được: ảnh / chữ ký / con dấu ở chế độ Xem + khi chèn ảnh; khối vẽ (hình, đường, hộp chữ...)
        # ở chế độ Xem + khi vẽ hình (lúc vẽ hình chỉ nắm ở viền khối để vẫn vẽ được hình mới bên trong)
        objs = []
        if self.mode == "edit":                         # Sửa file: khung chữ cũng chọn / kéo / dời bằng phím được
            m = page.rotation_matrix
            groups = {}                                 # dòng cùng khối (đoạn văn) -> 1 khối chữ
            for idx, ln in enumerate(self.canvas.lines):
                groups.setdefault(ln.block if ln.block >= 0 else -1 - idx, []).append(idx)
            z = self.zoom
            boxes = []
            for idxs in groups.values():
                r = pymupdf.Rect(self.canvas.lines[idxs[0]].bbox)
                for k in idxs[1:]:
                    r |= self.canvas.lines[k].bbox
                r = r * m
                r.normalize()
                objs.append((r, "block", tuple(idxs)) if len(idxs) > 1 else (r, "text", idxs[0]))
                boxes.append(QRectF(20 + r.x0 * z, 20 + r.y0 * z, r.width * z, r.height * z))
            self.canvas.rects = boxes                   # khung đứt nét: cả khối chữ thay vì từng dòng
        if self.mode in self.MOVE_MODES:
            objs += [(r, "img", x) for r, x in self.pdf.page_images(self.current)]
        if self.mode in self.MOVE_MODES or self.mode in self.SHAPE_MODES:
            objs += [(r, "annot", x) for r, x in self.pdf.page_annots(self.current)]
        if self.mode in self.MOVE_MODES or self.mode == "field":
            objs += [(r, "field", x) for r, x, _n in self.pdf.page_fields(self.current)]
        self.page_links = {}
        if self.mode in self.MOVE_MODES or self.mode == "link":
            for r, x, target, has_text in self.pdf.page_links(self.current):
                objs.append((r, "link", x))
                self.page_links[x] = (target, has_text)
        self.page_objs = objs
        tips = [self._link_tip(self.page_links[x][0]) if k == "link" else None for _r, k, x in objs]
        ratio = [k == "img" or (k == "link" and self.page_links[x][1]) for _r, k, x in objs]
        self.canvas.set_objects([r for r, _k, _x in objs], ratio, edge_only=self.mode in self.SHAPE_MODES, tips=tips,
                                resizable=[k not in ("text", "block") for _r, k, _x in objs])
        self._restore_selection()
        s = self.snap_region                            # khung Chụp vùng: giữ khi vẽ lại cùng trang, cùng chế độ
        if s and self.mode == "snapshot" and s[0] == self.current:
            self.canvas.set_region(s[1])
        else:
            self.snap_region = None
        self.update_ui()

    MOVE_MODES = ("view", "edit", "image", "handsign", "stamp")
    SHAPE_MODES = ("rect", "ellipse", "line", "arrow", "textbox", "typewriter", "callout", "cloud")

    @staticmethod
    def _link_tip(t):
        if "uri" in t:
            return f"{t['uri']}\n" + i18n.tr("Bấm để mở - kéo để di chuyển")
        if "page" in t:
            return i18n.tr("Đi tới trang {0}").format(t["page"] + 1) + "\n" + i18n.tr("Bấm để mở - kéo để di chuyển")
        if "file" in t:
            return f"{t['file']}\n" + i18n.tr("Bấm để mở - kéo để di chuyển")
        return i18n.tr("Liên kết - kéo để di chuyển")

    def on_object_clicked(self, i):
        if not 0 <= i < len(self.page_objs) or self.page_objs[i][1] != "link":
            return
        self.open_link(self.page_links.get(self.page_objs[i][2], ({}, False))[0])

    def open_link(self, t):
        from PySide6.QtCore import QUrl
        from PySide6.QtGui import QDesktopServices
        if "page" in t:
            self.goto(t["page"])
        elif "uri" in t:
            u = t["uri"]
            if u.lower().startswith(("http://", "https://", "mailto:")) or QMessageBox.question(
                    self, APP_NAME, f"Mở liên kết này?\n\n{u}") == QMessageBox.Yes:
                QDesktopServices.openUrl(QUrl(u))
        elif "file" in t:                            # file ngoài: hỏi trước (có thể là chương trình)
            p = t["file"]
            if not os.path.isabs(p) and self.pdf.path:
                p = os.path.join(os.path.dirname(self.pdf.path), p)
            if QMessageBox.question(self, APP_NAME, f"Liên kết mở file:\n{p}\n\nMở?") == QMessageBox.Yes:
                QDesktopServices.openUrl(QUrl.fromLocalFile(p))

    def on_object_moved(self, i, r):
        if not 0 <= i < len(self.page_objs):
            return
        old, kind, xref = self.page_objs[i]
        new = pymupdf.Rect(r.left(), r.top(), r.right(), r.bottom())
        s = getattr(self, "sel_obj", None)
        if s and s[:3] == (self.current, kind, xref):
            self.sel_obj = (self.current, kind, xref, new)
        try:
            self._move_obj(kind, xref, old, new)
        except Exception as e:
            self.show_error(e)
        else:
            self.status.showMessage("Đã di chuyển  (Ctrl+Z để hoàn tác)", 4000)
        self.refresh_thumb(self.current)
        self.render_page()

    def _move_obj(self, kind, xref, old, new):
        """Dời / đổi cỡ 1 vật trên trang hiện tại từ vùng hiển thị old sang new."""
        pno = self.current
        if kind == "text":                              # khung chữ: xref = chỉ số dòng trong canvas.lines
            line = self.canvas.lines[xref]
            page = self.pdf.doc[pno]
            o, n = pdf_core.disp_to_unrot(page, old), pdf_core.disp_to_unrot(page, new)
            body, fams = self._line_html(line)
            return self.pdf.move_line(pno, line, body, fams, n.x0 - o.x0, n.y0 - o.y0)
        if kind == "block":                             # khối chữ: xref = các chỉ số dòng
            lines = [self.canvas.lines[k] for k in xref]
            page = self.pdf.doc[pno]
            o, n = pdf_core.disp_to_unrot(page, old), pdf_core.disp_to_unrot(page, new)
            return self.pdf.move_lines(pno, lines, [self._line_html(ln) for ln in lines], n.x0 - o.x0, n.y0 - o.y0)
        move = {"img": self.pdf.move_image, "field": self.pdf.move_field}.get(kind, self.pdf.move_annot)
        return move(pno, xref, old, new)

    def _line_html(self, line):
        """Nội dung + định dạng của 1 dòng chữ PDF dạng HTML (như khi mở trình sửa chữ, không sửa gì)."""
        if getattr(self, "_line_conv", None) is None:
            self._line_conv = RichTextEdit(**RichTextDialog.last)
        self._line_conv.set_runs(RichTextDialog.line_runs(line))
        return self._line_conv.to_html()

    def on_object_activated(self, i):
        if not 0 <= i < len(self.page_objs):
            return
        _r, kind, xref = self.page_objs[i]
        if kind == "text":
            self.on_line_clicked(xref)
        elif kind == "block":
            self.on_block_clicked(xref)
        elif kind == "annot":
            self.on_note_clicked(xref)

    ALIGN_MARGIN = 28.35                           # 1 cm - khoảng chừa khi căn sát mép trang

    def _align_group(self, tab):
        g = tab.group("Căn chỉnh", "#0891b2", limit=0)        # 1 nút xổ xuống: 7 kiểu căn + chừa lề
        A = self.align_acts
        g.small([A["left"], A["hcenter"], A["right"]])
        g.small([A["top"], A["vcenter"], A["bottom"]])
        g.small([A["center"], self.a_align_margin])
        g.set_icon(0xE8E3)                                  # icon căn giữa (Segoe)

    def _align_menu(self, parent, i):
        """Menu con căn vật (đã được chọn khi chuột phải) theo trang - dùng chung các nút trên thanh công cụ."""
        sub = QMenu("Căn chỉnh trên trang", parent)
        sub.setIcon(self._align_icon("center"))
        A = self.align_acts
        for it in (A["left"], A["hcenter"], A["right"], None, A["top"], A["vcenter"], A["bottom"], None, A["center"],
                   None, self.a_align_margin):
            if it is None:
                sub.addSeparator()
            else:
                sub.addAction(it)
        return sub

    # ----- vật đang chọn: (trang, loại, xref, vùng hiển thị) - giữ qua các lần vẽ lại trang -----
    def on_object_selected(self, i):
        if 0 <= i < len(self.page_objs) and self.page_objs[i][1] != "link":
            r, kind, xref = self.page_objs[i]
            self.sel_obj = (self.current, kind, xref, pymupdf.Rect(r))
        else:
            self.sel_obj = None
        self._update_align_ui()

    def _sel_index(self):
        """Chỉ số trong page_objs của vật đang chọn (-1 nếu không có / không còn)."""
        s = getattr(self, "sel_obj", None)
        if not s or s[0] != self.current:
            return -1
        _p, kind, xref, rect = s
        best, bd = -1, None
        for i, (r, k, x) in enumerate(self.page_objs):
            text = kind in ("text", "block")
            if (k == kind or (text and k in ("text", "block"))) and (x == xref or text):
                # chữ: chỉ số dòng đổi sau khi sửa (khối có thể thành dòng và ngược lại) -> so vị trí
                d = abs(r.x0 - rect.x0) + abs(r.y0 - rect.y0) + abs(r.x1 - rect.x1) + abs(r.y1 - rect.y1)
                if bd is None or d < bd:
                    best, bd = i, d
        if kind in ("text", "block") and (bd is None or bd > 20):
            return -1
        return best

    def _restore_selection(self):
        i = self._sel_index()
        if i < 0:
            self.sel_obj = None
        else:
            self.canvas.select_object(i, notify=False)
        self._update_align_ui()

    def _update_align_ui(self):
        ok = self.pdf.doc is not None and self.in_editor() and self._sel_index() >= 0
        for a in self.align_acts.values():
            a.setEnabled(ok)

    def clear_object_selection(self):
        self.sel_obj = None
        self.canvas.select_object(-1, notify=False)
        self._update_align_ui()

    def align_selected(self, how):
        i = self._sel_index()
        if i < 0:
            self.status.showMessage("Bấm chọn ảnh / chữ ký / con dấu / khối vẽ trên trang trước", 4000)
            return
        self.align_object(i, how)

    def delete_selected_object(self):
        i = self._sel_index()
        if i < 0:
            return False
        old, kind, xref = self.page_objs[i]
        self.sel_obj = None
        if kind == "text":
            line = self.canvas.lines[xref]
            self.run_edit(lambda: self.pdf.edit_line_rich(self.current, line, "", []))
        elif kind == "block":
            lines = [self.canvas.lines[k] for k in xref]
            self.run_edit(lambda: self.pdf.delete_lines(self.current, lines))
        elif kind == "annot":
            self.run_edit(lambda: self.pdf.delete_annot(self.current, xref))
        elif kind == "field":
            self.run_edit(lambda: self.pdf.delete_field(self.current, xref))
        else:
            self.run_edit(lambda: self.pdf.delete_image(self.current, xref, old))
        self.status.showMessage("Đã xoá  (Ctrl+Z để hoàn tác)", 4000)
        return True

    @staticmethod
    def _align_icon(how):
        """Icon căn: khung trang + vật (ô xanh) ở vị trí sau khi căn."""
        pm = QPixmap(64, 64)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        p.setPen(QPen(QColor("#1e293b"), 4))
        p.setBrush(Qt.NoBrush)
        p.drawRoundedRect(QRectF(8, 6, 48, 52), 4, 4)
        w, h = 22, 16
        x = {"left": 15, "right": 49 - w}.get(how, 32 - w / 2)
        y = {"top": 13, "bottom": 51 - h}.get(how, 32 - h / 2)
        p.setPen(Qt.NoPen)
        p.setBrush(QColor("#2563eb"))
        p.drawRect(QRectF(x, y, w, h))
        p.end()
        return QIcon(pm)

    @staticmethod
    def _caret_icon():
        """Icon Chọn chữ: chữ 'A' + dấu nháy soạn thảo (I-beam), không phải con trỏ chuột."""
        pm = QPixmap(64, 64)
        pm.fill(Qt.transparent)
        p = QPainter(pm)
        p.setRenderHint(QPainter.Antialiasing)
        f = QFont("Segoe UI", 1)
        f.setPixelSize(40)
        f.setBold(True)
        p.setFont(f)
        p.setPen(QColor("#1e293b"))
        p.drawText(QRectF(0, 8, 40, 50), Qt.AlignCenter, "A")
        p.setPen(QPen(QColor("#1e293b"), 4, Qt.SolidLine, Qt.RoundCap))
        p.drawLine(QPointF(50, 10), QPointF(50, 54))           # thân dấu nháy
        p.drawLine(QPointF(43, 10), QPointF(57, 10))           # chân trên
        p.drawLine(QPointF(43, 54), QPointF(57, 54))           # chân dưới
        p.end()
        return QIcon(pm)

    def align_object(self, i, how):
        if not 0 <= i < len(self.page_objs):
            return
        old, kind, xref = self.page_objs[i]
        pr = self.pdf.doc[self.current].rect                # khổ trang như đang hiển thị (đã tính xoay trang)
        mg = self.ALIGN_MARGIN if self.settings.value("align_margin", True, type=bool) else 0
        w, h = old.width, old.height
        x0, y0 = old.x0, old.y0
        if how == "left":
            x0 = pr.x0 + mg
        elif how == "right":
            x0 = pr.x1 - mg - w
        elif how in ("hcenter", "center"):
            x0 = pr.x0 + (pr.width - w) / 2
        if how == "top":
            y0 = pr.y0 + mg
        elif how == "bottom":
            y0 = pr.y1 - mg - h
        elif how in ("vcenter", "center"):
            y0 = pr.y0 + (pr.height - h) / 2
        if abs(x0 - old.x0) < 0.05 and abs(y0 - old.y0) < 0.05:
            self.status.showMessage("Vật đã ở đúng vị trí này", 3000)
            return
        new = pymupdf.Rect(x0, y0, x0 + w, y0 + h)
        s = getattr(self, "sel_obj", None)
        if s and s[:3] == (self.current, kind, xref):
            self.sel_obj = (self.current, kind, xref, new)          # giữ chọn vật sau khi căn
        try:
            self._move_obj(kind, xref, old, new)
        except Exception as e:
            self.show_error(e)
        else:
            self.status.showMessage("Đã căn chỉnh  (Ctrl+Z để hoàn tác)", 4000)
        self.refresh_thumb(self.current)
        self.render_page()

    # ----- Cắt / Sao chép / Dán vật trên trang (chữ, khối chữ, ảnh, chữ ký, con dấu, khối vẽ) -----
    @staticmethod
    def _clip_seq():
        """Số thứ tự bộ nhớ tạm Windows - đổi khi có chương trình khác sao chép (để biết nên dán gì)."""
        try:
            import ctypes
            return ctypes.windll.user32.GetClipboardSequenceNumber()
        except Exception:
            return None

    def _set_sys_clip(self, text=None, png=None):
        """Đưa bản sao lên bộ nhớ tạm của Windows (dán được sang Word, Outlook...)."""
        cb = QApplication.clipboard()
        if png:
            cb.setImage(QImage.fromData(png))
        elif text:
            cb.setText(text)
        return self._clip_seq()

    def copy_object(self, i=None):
        i = self._sel_index() if i is None else i
        if not 0 <= i < len(self.page_objs):
            self.status.showMessage("Bấm chọn khung chữ / ảnh / con dấu / khối vẽ trên trang trước", 4000)
            return False
        r, kind, xref = self.page_objs[i]
        if kind == "link":
            return False
        if kind == "field":
            self.status.showMessage("Ô form không sao chép được - dùng Tạo ô form để thêm ô mới", 4000)
            return False
        pno = self.current
        page = self.pdf.doc[pno]
        c = {"kind": kind, "pdf": self.pdf, "doc": self.pdf.doc, "page": pno, "rect": pymupdf.Rect(r), "n": 0}
        try:
            if kind in ("text", "block"):
                lines = [self.canvas.lines[k] for k in (xref if kind == "block" else (xref,))]
                c.update(kind="text", items=[(ln, self._line_html(ln)) for ln in lines],
                         unrot=pdf_core.disp_to_unrot(page, pymupdf.Rect(r)))
                text = "\n".join(ln.text for ln in sorted(lines, key=lambda ln: (ln.bbox.y0, ln.bbox.x0)))
                c["seq"] = self._set_sys_clip(text=text)
            elif kind == "img":
                c.update(xref=xref, png=self.pdf.image_png(xref))
                c["seq"] = self._set_sys_clip(png=c["png"])
            else:
                c.update(xref=self.pdf.copy_annot_object(xref), png=self.pdf.annot_png(pno, xref))
                c["seq"] = self._set_sys_clip(png=c["png"])
        except Exception as e:
            self.show_error(e)
            return False
        self.obj_clip = c
        self.status.showMessage("Đã sao chép - Ctrl+V để dán (sang trang / file khác cũng được)", 4000)
        return True

    def cut_object(self):
        if self._sel_index() < 0 or not self.copy_object():
            return
        self.obj_clip["n"] = -1                     # dán lần đầu: đúng chỗ cũ
        self.delete_selected_object()
        self.status.showMessage("Đã cắt - Ctrl+V để dán", 4000)

    def can_paste(self):
        c = self.obj_clip
        if c is not None and (c.get("seq") is None or c["seq"] == self._clip_seq()):
            return True
        md = QApplication.clipboard().mimeData()
        return md is not None and (md.hasImage() or bool(md.text().strip()))

    def paste_object(self, pos=None):
        """Dán vào trang hiện tại: tại pos (pt hiển thị, khi chuột phải) hoặc lệch nhẹ so với bản gốc."""
        if not self.pdf.doc or not self.in_editor():
            return
        c = self.obj_clip
        if c is None or (c.get("seq") is not None and c["seq"] != self._clip_seq()):
            return self._paste_system(pos)
        pno = self.current
        page = self.pdf.doc[pno]
        r, pr = c["rect"], page.rect
        if pos is not None:
            x0, y0 = pos.x(), pos.y()
        elif c["pdf"] is self.pdf and c["page"] == pno:
            c["n"] += 1
            x0, y0 = r.x0 + 12 * c["n"], r.y0 + 12 * c["n"]
        else:
            x0, y0 = r.x0, r.y0
        x0 = max(pr.x0, min(x0, pr.x1 - r.width))
        y0 = max(pr.y0, min(y0, pr.y1 - r.height))
        new = pymupdf.Rect(x0, y0, x0 + r.width, y0 + r.height)
        same = c["pdf"] is self.pdf and c["doc"] is self.pdf.doc
        kind = c["kind"]
        if kind == "text":
            n = pdf_core.disp_to_unrot(page, new)
            dx, dy = n.x0 - c["unrot"].x0, n.y0 - c["unrot"].y0
            self.sel_obj = (pno, "text", -1, new)
            self.run_edit(lambda: self.pdf.paste_lines(pno, c["items"], dx, dy))
        elif kind == "img":
            self.sel_obj = (pno, "img", c["xref"], new) if same else None
            self.run_edit(lambda: self.pdf.paste_image(pno, new, png=c["png"], xref=c["xref"] if same else 0))
        elif same:
            nx = self.run_edit(lambda: self.pdf.paste_annot(pno, c["xref"], (x0, y0)))
            if nx:
                self.sel_obj = (pno, "annot", nx, new)
                self._restore_selection()
        else:
            self.run_edit(lambda: self.pdf.insert_image_stamp(pno, new, stream=c["png"]))
        self.status.showMessage("Đã dán  (Ctrl+Z để hoàn tác)", 4000)

    def _paste_system(self, pos=None):
        """Dán ảnh / chữ từ bộ nhớ tạm của Windows (chép từ Word, trình duyệt, ảnh chụp màn hình...)."""
        cb = QApplication.clipboard()
        md = cb.mimeData()
        pno = self.current
        pr = self.pdf.doc[pno].rect
        img = cb.image() if md is not None and md.hasImage() else QImage()
        if not img.isNull():
            k = min(0.75, pr.width * 0.6 / img.width(), pr.height * 0.6 / img.height())   # 96 dpi -> pt
            w, h = img.width() * k, img.height() * k
            x0, y0 = (pos.x(), pos.y()) if pos is not None else (pr.x0 + (pr.width - w) / 2, pr.y0 + (pr.height - h) / 2)
            x0, y0 = max(pr.x0, min(x0, pr.x1 - w)), max(pr.y0, min(y0, pr.y1 - h))
            buf = QBuffer()
            buf.open(QBuffer.WriteOnly)
            img.save(buf, "PNG")
            png = bytes(buf.data())
            self.run_edit(lambda: self.pdf.insert_image(pno, (x0, y0, x0 + w, y0 + h), stream=png))
            self.status.showMessage("Đã dán ảnh  (Ctrl+Z để hoàn tác)", 4000)
            return
        text = md.text().strip() if md is not None else ""
        if not text:
            self.status.showMessage("Bộ nhớ tạm không có gì để dán", 4000)
            return
        conv = RichTextEdit(**RichTextDialog.last)
        conv.edit.setPlainText(text)
        body, fams = conv.to_html()
        x0, y0 = (pos.x(), pos.y()) if pos is not None else (pr.x0 + 56, pr.y0 + 56)
        self.run_edit(lambda: self.pdf.add_rich_text(pno, (x0, y0, x0, y0 + 20), body, fams))
        self.status.showMessage("Đã dán chữ  (Ctrl+Z để hoàn tác)", 4000)

    def duplicate_object(self, i):
        old, kind, xref = self.page_objs[i]
        if kind == "img":
            self.run_edit(lambda: self.pdf.duplicate_image(self.current, xref, old, old + (15, 15, 15, 15)))
            return
        keep = self.obj_clip
        if self.copy_object(i):
            self.obj_clip["seq"] = None
            self.paste_object()
        self.obj_clip = keep

    def save_object_image(self, i):
        old, kind, xref = self.page_objs[i]
        try:
            png = self.pdf.image_png(xref) if kind == "img" else self.pdf.annot_png(self.current, xref)
        except Exception as e:
            self.show_error(e)
            return
        base = os.path.splitext(os.path.basename(self.pdf.path or "trang"))[0]
        path, _ = QFileDialog.getSaveFileName(self, "Lưu thành ảnh",
                                              os.path.join(self._dir(), f"{base}_trang{self.current + 1}_anh.png"),
                                              "Ảnh PNG (*.png)")
        if not path:
            return
        with open(path, "wb") as fh:
            fh.write(png)
        self.toast("Đã lưu ảnh", os.path.basename(path))

    def on_page_menu(self, pt, gpos):
        """Chuột phải lên chỗ trống của trang: Dán + đổi chế độ."""
        if not self.pdf.doc or self.mode not in self.MOVE_MODES:
            return
        m = QMenu(self)
        a_paste = m.addAction(glyph_icon(0xE77F), "Dán")
        a_paste.setShortcut(QKeySequence("Ctrl+V"))
        a_paste.setShortcutVisibleInContextMenu(True)
        a_paste.setEnabled(self.can_paste())
        m.addSeparator()
        for key in ("view", "select", "edit", "snapshot"):
            if key in self.tool_actions:
                m.addAction(self.tool_actions[key])
        if m.exec(gpos) is a_paste:
            self.paste_object(pt)

    def on_object_menu(self, i, gpos):
        if not 0 <= i < len(self.page_objs):
            return
        old, kind, xref = self.page_objs[i]
        m = QMenu(self)

        def act(icon, text, key=None):
            a = m.addAction(glyph_icon(icon), text)
            if key:
                a.setShortcut(QKeySequence(key))
                a.setShortcutVisibleInContextMenu(True)
            return a

        a_note = a_edit = a_front = a_back = a_save = None
        if kind == "field":
            name = next((n for r, x, n in self.pdf.page_fields(self.current) if x == xref), "")
            m.addAction(f"Ô form: {name}").setEnabled(False)
            a_ren = act(0xE8AC, "Đổi tên ô...")
            a_del = act(0xE74D, "Xóa", "Delete")
            m.addSeparator()
            m.addMenu(self._align_menu(m, i))
            a_desel = act(0xE711, "Bỏ chọn", "Esc")
            act_ = m.exec(gpos)
            if act_ is a_del:
                self.delete_selected_object()
            elif act_ is a_desel:
                self.clear_object_selection()
            elif act_ is a_ren:
                new, ok = QInputDialog.getText(self, "Đổi tên ô form", "Tên ô (dùng khi xuất / nhập dữ liệu form):",
                                               text=name)
                if ok and new.strip() and new.strip() != name:
                    self.run_edit(lambda: self.pdf.rename_field(self.current, xref, new.strip()))
            return
        if kind != "link":
            if kind in ("text", "block"):
                a_edit = act(0xE70F, "Sửa chữ...", "Return")
                m.addSeparator()
            elif kind == "annot":
                a_note = act(0xE70F, "Xem / sửa ghi chú của khối")
                m.addSeparator()
            a_cut = act(0xE8C6, "Cắt", "Ctrl+X")
            a_copy = act(0xE8C8, "Sao chép", "Ctrl+C")
            a_paste = act(0xE77F, "Dán", "Ctrl+V")
            a_paste.setEnabled(self.can_paste())
            a_dup = act(0xE8C8, "Nhân bản")
            a_del = act(0xE74D, "Xóa", "Delete")
            m.addSeparator()
            if kind in ("img", "annot"):
                sub = m.addMenu(glyph_icon(0xE8B1), "Sắp xếp")
                a_front = sub.addAction(glyph_icon(0xE74A), "Đưa lên trên cùng")
                a_back = sub.addAction(glyph_icon(0xE74B), "Đưa xuống dưới cùng")
            m.addMenu(self._align_menu(m, i))
            m.addSeparator()
            if kind in ("img", "annot"):
                a_save = act(0xE74E, "Lưu thành ảnh...")
            a_desel = act(0xE711, "Bỏ chọn", "Esc")
            act_ = m.exec(gpos)
            if act_ is None:
                return
            if act_ is a_edit or act_ is a_note:
                self.on_object_activated(i)
            elif act_ is a_cut:
                self.cut_object()
            elif act_ is a_copy:
                self.copy_object(i)
            elif act_ is a_paste:
                self.paste_object()
            elif act_ is a_dup:
                self.duplicate_object(i)
            elif act_ is a_del:
                self.delete_selected_object()
            elif act_ in (a_front, a_back):
                self.run_edit(lambda: self.pdf.arrange(self.current, kind, xref, old, act_ is a_front))
            elif act_ is a_save:
                self.save_object_image(i)
            elif act_ is a_desel:
                self.clear_object_selection()
            return
        if kind == "link":
            target = self.page_links.get(xref, ({}, False))[0]
            a_open = m.addAction(glyph_icon(0xE8A7), "Mở liên kết")
            a_edit = m.addAction(glyph_icon(0xE70F), "Sửa địa chỉ liên kết...")
            a_edit.setEnabled("uri" in target)
            a_copy = m.addAction(glyph_icon(0xE8C8), "Sao chép địa chỉ")
            a_copy.setEnabled("uri" in target)
            a_rm = m.addAction(glyph_icon(0xE74D), "Xóa")
            m.addSeparator()
            m.addAction("Kéo để di chuyển, kéo ô vuông ở góc dưới phải để đổi kích thước").setEnabled(False)
            act = m.exec(gpos)
            if act is None:
                return
            if act is a_open:
                self.open_link(target)
            elif act is a_copy:
                QApplication.clipboard().setText(target.get("uri", ""))
            elif act is a_rm:
                self.run_edit(lambda: self.pdf.delete_link(self.current, xref))
            elif act is a_edit:
                u, ok = QInputDialog.getText(self, "Sửa liên kết", "Địa chỉ web / email:", text=target.get("uri", ""))
                if ok and u.strip():
                    u = LinkDialog.normalize(u.strip())
                    self.run_edit(lambda: self.pdf.set_link_uri(self.current, xref, u))

    # ------------------------------------------------------------------ file
    def confirm_discard(self):
        if not self.pdf.modified:
            return True
        r = QMessageBox.question(self, APP_NAME, "File đã thay đổi. Lưu lại trước khi tiếp tục?",
                                 QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)
        if r == QMessageBox.Save:
            return self.save()
        return r == QMessageBox.Discard

    def _ask_password(self, path):
        pw, ok = QInputDialog.getText(self, "File có mật khẩu", f"Nhập mật khẩu để mở:\n{os.path.basename(path)}",
                                      QLineEdit.Password)
        return pw if ok else None

    def confirm_discard_all(self):
        """Hỏi lưu cho mọi tab đang có thay đổi (trước khi thoát / dựng lại cửa sổ)."""
        if not self.docs:
            return self.confirm_discard()
        asked = []
        for i in range(len(self.docs)):
            if self.docs[i].pdf.modified:
                self.tabbar.setCurrentIndex(i)
                self.show_editor()
                if not self.confirm_discard():
                    return False
                asked.append(self.docs[i].pdf)
        for p in asked:                     # đã chọn "Không lưu" -> không hỏi lại
            p.modified = False
        return True

    def _discard_all(self):
        for d in self.docs:
            d.pdf.modified = False
        self.pdf.modified = False

    def _find_tab(self, path):
        key = os.path.normcase(os.path.abspath(path))
        for i, d in enumerate(self.docs):
            if d.pdf.path and os.path.normcase(os.path.abspath(d.pdf.path)) == key:
                return i
        return -1

    def open_path(self, path, new_tab=True):
        """Mở file trong tab mới (file đã mở rồi thì chuyển sang tab đó). new_tab=False: thay tài liệu tab hiện tại."""
        i = self._find_tab(path)
        if i >= 0:
            self.tabbar.setCurrentIndex(i)
            self.show_editor()
            return True
        target = PdfDocument() if (new_tab and self.docs) else self.pdf
        password = None
        while True:
            try:
                QApplication.setOverrideCursor(Qt.WaitCursor)
                target.open(path, password)
                break
            except PdfError as e:
                QApplication.restoreOverrideCursor()
                if str(e) == "NEED_PASSWORD":
                    password = self._ask_password(path)
                    if password is None:
                        return False
                    continue
                self.show_error(e)
                return False
            except Exception as e:
                QApplication.restoreOverrideCursor()
                self.show_error(f"Không mở được file:\n{e}")
                return False
            finally:
                if QApplication.overrideCursor():
                    QApplication.restoreOverrideCursor()
        self.editor_area.setCurrentWidget(self.splitter)
        if os.path.splitext(path)[1].lower() == ".pdf":
            self.add_recent(os.path.abspath(path))
        self.ribbon.setCurrentIndex(0)
        if target is self.pdf and self.docs:            # thay tài liệu của tab hiện tại
            tab = self.docs[self._active]
            tab.current, tab.fit_width, tab.scroll = 0, True, 0
            self.current, self.fit_width = 0, True
            self.show_editor()
            self.refresh_thumbs([0])
        else:                                           # tab mới
            self._add_tab(DocTab(target))
            self.show_editor()
        self.set_mode("select")                         # mở file: mặc định ở chế độ Chọn chữ
        self._update_tabs()
        return True

    def open_file(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Mở file PDF", self._dir(), OPEN_FILTER)
        for p in paths:
            self.open_path(p)

    def new_document(self):
        """Ctrl+N: tài liệu trống 1 trang A4 trong tab mới (lưu lần đầu sẽ hỏi chỗ lưu)."""
        target = PdfDocument() if self.docs else self.pdf
        target.new_blank()
        self.editor_area.setCurrentWidget(self.splitter)
        self.ribbon.setCurrentIndex(0)
        self._add_tab(DocTab(target))
        self.show_editor()
        self._update_tabs()

    # ------------------------------------------------------------------ in / gửi email
    def print_doc(self):
        """Ctrl+P: hộp thoại In có xem trước (máy in, số bản, trang, 2 mặt, theo bộ, số trang / tờ, khổ giấy,
        thu phóng, màu). Trang được dựng ảnh ở <= 300 dpi."""
        if not self.pdf.doc:
            return
        dlg = PrintDialog(self)
        if dlg.exec() != QDialog.Accepted or not dlg.result_pages:
            return
        try:
            self._print_pages(dlg.printer, dlg.result_pages, dlg.options(), getattr(dlg, "result_copies", 1),
                              getattr(dlg, "result_collate", True))
        except Exception as e:
            self.show_error(f"Lỗi khi in:\n{e}")

    def _print_pages(self, pr, pages, opts=None, copies=1, collate=True):
        prog = QProgressDialog("Đang gửi trang tới máy in...", "Huỷ", 0, len(pages), self)
        prog.setWindowTitle("In")
        prog.setWindowModality(Qt.WindowModal)
        prog.setMinimumDuration(400)
        try:
            done = self._render_pages(pr, pages, min(300, max(150, pr.resolution())), opts, prog, copies, collate)
        finally:
            prog.setValue(prog.maximum())
        if done:
            self.status.showMessage(f"Đã gửi {len(pages)} trang tới máy in {pr.printerName()}", 6000)

    _NUP_GRIDS = {1: [(1, 1)], 2: [(1, 2), (2, 1)], 4: [(2, 2)], 6: [(2, 3), (3, 2)], 9: [(3, 3)], 16: [(4, 4)]}

    @staticmethod
    def _fit(pw, ph, cw, ch, autorot):
        s = min(cw / pw, ch / ph)
        return max(s, min(cw / ph, ch / pw)) if autorot else s

    def _sheet_cells(self, area, nup, pw, ph, o):
        """Các ô trên 1 tờ giấy cho chế độ nhiều trang / tờ (chọn lưới cho trang to nhất)."""
        best = None
        for cols, rows in self._NUP_GRIDS.get(nup, [(1, 1)]):
            s = self._fit(pw, ph, area.width() / cols, area.height() / rows, o["autorot"])
            if best is None or s > best[0]:
                best = (s, cols, rows)
        _s, cols, rows = best
        cw, ch = area.width() / cols, area.height() / rows
        gap = 0 if nup == 1 else min(cw, ch) * 0.03
        cells = []
        for k in range(cols * rows):
            c, r = (k // rows, k % rows) if o["down_first"] else (k % cols, k // cols)
            cells.append(QRectF(area.x() + c * cw + gap / 2, area.y() + r * ch + gap / 2, cw - gap, ch - gap))
        return cells

    def _draw_print_page(self, p, pno, cell, dpi, o, res, nup):
        page = self.pdf.doc[pno]
        pw, ph = page.rect.width, page.rect.height
        rotate = o["autorot"] and abs(pw - ph) > 1 and (pw > ph) != (cell.width() > cell.height())
        w_pt, h_pt = (ph, pw) if rotate else (pw, ph)
        dw, dh = w_pt * res / 72, h_pt * res / 72                       # cỡ thật của trang trên giấy (pixel máy in)
        fit = min(cell.width() / dw, cell.height() / dh)
        mode = o["scale"] if nup == 1 else 0
        s = (fit, 1.0, min(1.0, fit), o["pct"] / 100)[mode]
        w, h = dw * s, dh * s
        if o["center"]:
            x, y = cell.x() + (cell.width() - w) / 2, cell.y() + (cell.height() - h) / 2
        else:
            x, y = cell.x(), cell.y()
        z = dpi * min(max(s, 0.25), 1.5) / 72
        pix = page.get_pixmap(matrix=pymupdf.Matrix(z, z), alpha=False,
                              colorspace=pymupdf.csGRAY if o["gray"] else pymupdf.csRGB)
        img = pix_to_qimage(pix)
        if rotate:
            img = img.transformed(QTransform().rotate(90), Qt.SmoothTransformation)
        target = QRectF(x, y, w, h)
        p.save()
        p.setClipRect(cell)
        p.drawImage(target, img)
        if o["border"] and nup > 1:
            p.setPen(QPen(QColor("#94a3b8"), max(1.0, res / 150)))
            p.setBrush(Qt.NoBrush)
            p.drawRect(target)
        p.restore()

    def _render_pages(self, pr, pages, dpi, opts=None, prog=None, copies=1, collate=True):
        """Vẽ các trang lên máy in / bản xem trước theo tuỳ chọn in. False nếu người dùng huỷ."""
        o = dict(gray=False, nup=1, down_first=False, scale=0, pct=100, autorot=True, center=True, border=False)
        if isinstance(opts, dict):
            o.update(opts)
        elif opts:
            o["gray"] = True
        nup = max(1, int(o["nup"]))
        sheets = [pages[i:i + nup] for i in range(0, len(pages), nup)]
        if copies > 1:                                                  # máy in không tự in nhiều bản
            sheets = sheets * copies if collate else [sh for sh in sheets for _ in range(copies)]
        if prog is not None:
            prog.setMaximum(len(sheets))
        p = QPainter()
        if not p.begin(pr):
            raise PdfError("Không mở được máy in.")
        res = pr.resolution()
        area = QRectF(pr.pageLayout().paintRectPixels(res))
        area.moveTo(0, 0)
        try:
            for k, sheet in enumerate(sheets):
                if prog is not None:
                    prog.setValue(k)
                    QApplication.processEvents()
                    if prog.wasCanceled():
                        pr.abort()
                        return False
                if k:
                    pr.newPage()
                first = self.pdf.doc[sheet[0]].rect
                cells = self._sheet_cells(area, nup, first.width, first.height, o)
                for j, pno in enumerate(sheet):
                    self._draw_print_page(p, pno, cells[j], dpi, o, res, nup)
        finally:
            p.end()
        return True

    def _mail_copy(self):
        """File để đính kèm: file gốc nếu đã lưu, không thì bản sao nội dung đang sửa trong thư mục tạm."""
        if self.pdf.path and not self.pdf.modified and self.pdf.path.lower().endswith(".pdf"):
            return self.pdf.path
        import tempfile
        name = os.path.splitext(os.path.basename(self.pdf.path or "tai_lieu"))[0] + ".pdf"
        d = os.path.join(tempfile.gettempdir(), "PDFToolkit_mail")
        os.makedirs(d, exist_ok=True)
        out = os.path.join(d, name)
        with open(out, "wb") as f:
            f.write(self.pdf.doc.tobytes(garbage=3, deflate=True))
        return out

    def email_doc(self):
        """Mở thư mới của Outlook / ứng dụng email mặc định với file PDF đính kèm (Simple MAPI - người dùng tự bấm Gửi)."""
        if not self.pdf.doc:
            return
        try:
            path = self._mail_copy()
        except Exception as e:
            self.show_error(f"Không tạo được file đính kèm:\n{e}")
            return
        self.status.showMessage("Đang mở thư mới trong ứng dụng email...", 8000)
        job = MailJob(path, os.path.basename(path))
        job.finished_with.connect(lambda code, p=path: self._mail_done(code, p))
        job.start()

    def _mail_done(self, code, path):
        if code in (0, 1):                        # 0 = đã mở thư / gửi, 1 = người dùng huỷ
            self.status.clearMessage()
            return
        r = QMessageBox.question(
            self, APP_NAME, f"Không mở được ứng dụng email (mã lỗi MAPI {code}).\n"
            "Máy cần cài Outlook (hoặc ứng dụng email mặc định có hỗ trợ MAPI).\n\n"
            "Mở thư mục chứa file để tự đính kèm?")
        if r == QMessageBox.Yes:
            import subprocess
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])

    # ------------------------------------------------------------------ tab tài liệu
    def _add_tab(self, tab):
        self._save_tab_state()
        self.docs.append(tab)
        self.tabbar.blockSignals(True)
        i = self.tabbar.addTab("")
        x = QToolButton()
        x.setObjectName("tabCloseBtn")
        x.setIcon(glyph_icon(0xE711, "#64748b"))
        x.setIconSize(QSize(10, 10))
        x.setAutoRaise(True)
        x.setToolTip("Đóng tab  (Ctrl+W)")
        x.clicked.connect(lambda _=False, b=x: self.close_tab(
            next((k for k in range(self.tabbar.count()) if self.tabbar.tabButton(k, QTabBar.RightSide) is b), -1)))
        self.tabbar.setTabButton(i, QTabBar.RightSide, x)
        self.tabbar.blockSignals(False)
        self._active = -1                               # tab cũ đã lưu trạng thái
        self.tabbar.setCurrentIndex(i)
        if self._active != i:                           # setCurrentIndex không phát tín hiệu khi đây là tab đầu tiên
            self._activate_tab(i)

    def _save_tab_state(self):
        if 0 <= self._active < len(self.docs):
            t = self.docs[self._active]
            t.current, t.zoom, t.fit_width = self.current, self.zoom, self.fit_width
            t.scroll = self.scroll.verticalScrollBar().value()

    def _activate_tab(self, i):
        if not 0 <= i < len(self.docs):
            return
        if i != self._active:
            self._save_tab_state()
        self._active = i
        t = self.docs[i]
        self.pdf = t.pdf
        self.current, self.zoom, self.fit_width = t.current, t.zoom, t.fit_width
        self.editor_area.setCurrentWidget(self.splitter if self.pdf.doc else self.empty)
        if self.pdf.doc:
            self.refresh_thumbs([min(t.current, self.pdf.page_count - 1)])
            pos = t.scroll
            QTimer.singleShot(0, lambda: self.scroll.verticalScrollBar().setValue(pos))
        self._update_tabs()
        self.update_ui()

    def _tab_moved(self, frm, to):
        self.docs.insert(to, self.docs.pop(frm))
        self._active = self.tabbar.currentIndex()

    def _cycle_tab(self, step):
        n = self.tabbar.count()
        if n > 1:
            self.tabbar.setCurrentIndex((self.tabbar.currentIndex() + step) % n)

    def close_tab(self, i):
        if not 0 <= i < len(self.docs):
            return False
        if i != self.tabbar.currentIndex():
            self.tabbar.setCurrentIndex(i)
        if not self.confirm_discard():
            return False
        tab = self.docs.pop(i)
        self._active = -1                               # không lưu trạng thái của tab đang đóng
        if not self.docs:
            self.tabbar.blockSignals(True)
            self.tabbar.removeTab(i)
            self.tabbar.blockSignals(False)
            self.pdf = PdfDocument()                    # không còn tab -> về trang chủ
            self.thumbs.clear()
            self.editor_area.setCurrentWidget(self.empty)
            self.show_home()
        else:
            self.tabbar.removeTab(i)                    # Qt chọn tab kế bên -> _activate_tab
            if self._active != self.tabbar.currentIndex():
                self._activate_tab(self.tabbar.currentIndex())
        if tab.pdf.doc:
            tab.pdf.doc.close()
        self._update_tabs()
        return True

    def _close_tabs(self, keep=None, right_of=None):
        for i in reversed(range(len(self.docs))):
            if (keep is not None and i != keep) or (right_of is not None and i > right_of):
                if not self.close_tab(i):
                    return
                if keep is not None and i < keep:
                    keep -= 1

    def _tab_menu(self, pos):
        i = self.tabbar.tabAt(pos)
        if i < 0:
            return
        path = self.docs[i].pdf.path
        m = QMenu(self)
        m.addAction(glyph_icon(0xE711), "Đóng", lambda: self.close_tab(i))
        m.addAction("Đóng các tab khác", lambda: self._close_tabs(keep=i)).setEnabled(len(self.docs) > 1)
        m.addAction("Đóng các tab bên phải", lambda: self._close_tabs(right_of=i)).setEnabled(i < len(self.docs) - 1)
        m.addSeparator()
        a = m.addAction(glyph_icon(0xE838), "Mở thư mục chứa file", lambda: tools_ui.reveal(path))
        a.setEnabled(bool(path and os.path.exists(path)))
        b = m.addAction(glyph_icon(0xE8C8), "Sao chép đường dẫn", lambda: QApplication.clipboard().setText(path or ""))
        b.setEnabled(bool(path))
        m.exec(self.tabbar.mapToGlobal(pos))

    def eventFilter(self, obj, ev):
        if obj is self.tabbar and ev.type() == ev.Type.MouseButtonRelease and ev.button() == Qt.MiddleButton:
            i = self.tabbar.tabAt(ev.position().toPoint())
            if i >= 0:
                self.close_tab(i)
                return True
        return super().eventFilter(obj, ev)

    def _update_tabs(self):
        if not hasattr(self, "docs"):
            return
        for i, d in enumerate(self.docs):
            name = os.path.basename(d.pdf.path) if d.pdf.path else i18n.tr("Tài liệu mới")
            if os.path.splitext(name)[1].lower() != ".pdf" and d.pdf.path:
                name = os.path.splitext(name)[0] + ".pdf"   # ảnh mở thành PDF
            text = ("● " if d.pdf.modified else "") + name
            with i18n.paused():                         # tên file: không dịch
                if self.tabbar.tabText(i) != text:
                    self.tabbar.setTabText(i, text)
                tip = d.pdf.path or name
                if self.tabbar.tabToolTip(i) != tip:
                    self.tabbar.setTabToolTip(i, tip)
        self.tabbar.parentWidget().setVisible(bool(self.docs))

    def _dir(self):
        return os.path.dirname(self.pdf.path) if self.pdf.path else os.path.expanduser("~\\Documents")

    def insert_paths(self, paths, at):
        passwords = {}
        for p in paths:
            if p.lower().endswith(".pdf"):
                try:
                    d = pymupdf.open(p)
                    if d.needs_pass:
                        pw = self._ask_password(p)
                        if pw is None:
                            return
                        passwords[p] = pw
                    d.close()
                except Exception as e:
                    self.show_error(f"Không đọc được {os.path.basename(p)}:\n{e}")
                    return
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            n = self.pdf.insert_files(paths, at, passwords)
        except Exception as e:
            self.show_error(f"Không gộp được file:\n{e}")
            return
        finally:
            QApplication.restoreOverrideCursor()
        self.refresh_thumbs(list(range(at, at + n)))
        self.status.showMessage(f"Đã chèn {n} trang từ {len(paths)} file", 5000)

    def add_files(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Chọn file để gộp", self._dir(), OPEN_FILTER)
        if paths:
            at = (max(self.selected_pages()) + 1) if self.pdf.page_count else 0
            self.insert_paths(paths, at)

    def on_files_dropped(self, paths, at):
        paths = [p for p in paths if os.path.splitext(p)[1].lower() in pdf_core.IMAGE_EXTS | {".pdf"}]
        if not paths:
            return
        if not self.pdf.doc:
            if self.open_path(paths[0]) and len(paths) > 1:
                self.insert_paths(paths[1:], self.pdf.page_count)
        else:
            self.insert_paths(paths, at)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()

    def dropEvent(self, e):
        paths = [u.toLocalFile() for u in e.mimeData().urls() if u.isLocalFile()]
        if self.stack.currentWidget() is self.tool_page:
            self.tool_page.files.add_paths(paths)
            return
        # thả vào cửa sổ (ngoài danh sách trang): mỗi file mở thành một tab; thả vào danh sách trang = chèn trang
        for p in paths:
            if os.path.splitext(p)[1].lower() in pdf_core.IMAGE_EXTS | {".pdf"}:
                self.open_path(p)

    # ------------------------------------------------------------------ màn hình
    def in_editor(self):
        return self.stack.currentWidget() is self.doc_area

    def show_home(self):
        rec = self.settings.value("recent", []) or []
        self.home.open_card.set_recent([rec] if isinstance(rec, str) else rec)
        self.home.set_open_docs([self.tabbar.tabText(i) for i in range(self.tabbar.count())],
                                self.tabbar.currentIndex())
        self.stack.setCurrentWidget(self.home)
        self.toolbar.hide()
        self.zoom_bar.hide()
        self.home.search.setFocus()
        self.update_ui()

    def resume_doc(self, i=-1):
        """Từ trang chủ quay lại tab tài liệu (i = -1: tab đang chọn)."""
        if not self.docs:
            return
        if 0 <= i < self.tabbar.count() and i != self.tabbar.currentIndex():
            self.tabbar.setCurrentIndex(i)
        self.show_editor()

    def show_editor(self):
        self.stack.setCurrentWidget(self.doc_area)
        self.toolbar.show()
        self.zoom_bar.show()
        self.update_ui()

    def open_tool(self, tid):
        if tid.startswith("recent:"):
            self.open_path(tid[7:])
            return
        if tid == "editor:open":
            self.open_file()
            return
        if tid == "editor:scan":
            self.scan_to_pdf(False)
            return
        if tid.startswith("editor:"):
            if not self.pdf.doc:
                path, _ = QFileDialog.getOpenFileName(self, "Mở file PDF", self._dir(), OPEN_FILTER)
                if not path or not self.open_path(path):
                    return
            self.show_editor()
            kind = tid.split(":", 1)[1]
            if kind == "pages" and not self.grid:
                self.a_grid.setChecked(True)
                self.toggle_grid()
            elif kind == "sign":
                self.set_mode("sign")
            elif kind == "edit" and self.grid:
                self.a_grid.setChecked(False)
                self.toggle_grid()
            return
        t = pdf_tools.get_tool(tid)
        files = []
        if self.pdf.path and not self.pdf.modified and os.path.splitext(self.pdf.path)[1].lower() in t.accept:
            files = [self.pdf.path]                 # dùng luôn file đang mở trong trình sửa
        self.tool_page.load_tool(tid, files)
        self.stack.setCurrentWidget(self.tool_page)
        self.toolbar.hide()
        self.zoom_bar.hide()
        self.update_ui()

    # Công cụ áp dụng thẳng vào tài liệu đang sửa (hoàn tác được); còn lại xuất ra file mới
    IN_PLACE_TOOLS = {"watermark", "header_footer", "page_numbers", "bates", "stamp", "bg_color", "greyscale", "invert",
                      "scan_effect", "rasterize", "crop", "fix_size", "divide", "nup", "booklet", "poster",
                      "remove_blank", "reverse", "blank_page", "bookmarks", "toc_page", "add_attach", "overlay",
                      "metadata", "sanitize", "flatten", "remove_annots", "form_fill", "ocr", "compress", "repair",
                      "delete_pages", "unlock", "rotate_custom", "deskew", "duplex_collate",
                      "combine_single_page", "page_labels", "edit_attach", "adjust_colors", "text_color",
                      "font_to_outline", "layers"}

    def run_doc_tool(self, tid):
        if not self.pdf.doc:
            return
        import shutil
        import tempfile
        t = pdf_tools.get_tool(tid)
        in_place = tid in self.IN_PLACE_TOOLS
        tmp = tempfile.mkdtemp(prefix="pdftoolkit_")
        stem = os.path.splitext(os.path.basename(self.pdf.path or "tai_lieu"))[0]
        src = os.path.join(tmp, stem + ".pdf")
        try:
            if self.pdf.new_fonts:
                try:
                    self.pdf.doc.subset_fonts()
                except Exception:
                    pass
                self.pdf.new_fonts = False
            with open(src, "wb") as fh:
                fh.write(self.pdf.doc.tobytes(garbage=1, deflate=True, encryption=pymupdf.PDF_ENCRYPT_NONE))
            out_dir = tmp if in_place else (os.path.dirname(self.pdf.path) if self.pdf.path
                                            else os.path.expanduser("~\\Documents"))
            acts = [] if in_place else [("tab", "Mở PDF kết quả trong tab mới")]
            dlg = tools_ui.ToolDialog(self, t, src, in_place, out_dir, pdf_actions=acts)
            ok = dlg.exec() == QDialog.Accepted
            if ok and dlg.result_action == "tab":
                for p in dlg.result_paths:
                    self.open_path(p)
            elif ok and dlg.result_path:
                with open(dlg.result_path, "rb") as fh:
                    data = fh.read()
                self.pdf.snapshot()
                self.pdf.doc.close()
                self.pdf.doc = pymupdf.open("pdf", data)
                self.refresh_thumbs([min(self.current, self.pdf.page_count - 1)])
                self.status.showMessage(f"Đã áp dụng: {t.name}  (Ctrl+Z để hoàn tác)", 6000)
        except Exception as e:
            self.show_error(f"{t.name}: {e}")
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def create_from_files(self, tid):
        """Nhóm 'Tạo PDF từ file khác' + 'So sánh 2 file' trong trình sửa.
        Chọn file nguồn -> chuyển đổi -> mở kết quả trong tab mới hoặc chèn vào sau trang đang xem."""
        import shutil
        import tempfile
        t = pdf_tools.get_tool(tid)
        exts = sorted(set(t.accept))
        flt = f"{t.name} ({' '.join('*' + e for e in exts)})"
        tmp = None
        if tid == "compare":
            if not self.pdf.doc:
                return
            other, _ = QFileDialog.getOpenFileName(self, "Chọn file để so sánh với file đang mở", self._dir(),
                                                   "PDF (*.pdf)")
            if not other:
                return
            tmp = tempfile.mkdtemp(prefix="pdftoolkit_")
            stem = os.path.splitext(os.path.basename(self.pdf.path or "tai_lieu"))[0]
            cur = os.path.join(tmp, stem + ".pdf")
            if self.pdf.path and not self.pdf.modified and self.pdf.path.lower().endswith(".pdf"):
                cur = self.pdf.path
            else:
                with open(cur, "wb") as fh:          # bản đang sửa (chưa lưu) cũng so sánh được
                    fh.write(self.pdf.doc.tobytes(garbage=1, deflate=True,
                                                  encryption=pymupdf.PDF_ENCRYPT_NONE))
            srcs = [cur, other]
            actions = [("tab", "Mở kết quả trong tab mới")]
        else:
            srcs, _ = QFileDialog.getOpenFileNames(self, t.name, self._dir(), flt)
            if not srcs:
                return
            actions = [("tab", "Mở trong tab mới")]
            if self.pdf.doc:
                actions.append(("insert", "Chèn vào sau trang đang xem"))
        out_dir = os.path.dirname(self.pdf.path) if (tid == "compare" and self.pdf.path) else os.path.dirname(srcs[0])
        if tid == "compare" and not self.pdf.path:
            out_dir = os.path.expanduser("~\\Documents")
        try:
            dlg = tools_ui.ToolDialog(self, t, srcs, False, out_dir, pdf_actions=actions, show_sources=True)
            if dlg.exec() != QDialog.Accepted or not dlg.result_paths:
                return
            if dlg.result_action == "insert" and self.pdf.doc:
                self.insert_paths(dlg.result_paths, self.current + 1)
            else:
                for p in dlg.result_paths:
                    self.open_path(p)
        except Exception as e:
            self.show_error(f"{t.name}: {e}")
        finally:
            if tmp:
                shutil.rmtree(tmp, ignore_errors=True)

    # ------------------------------------------------------------------ cài đặt
    def show_about(self, tab=0):
        exe = installer.installed_exe()
        if installer.is_installed_copy():
            info = f"Đã cài vào máy: {os.path.dirname(sys.executable)}"
        elif exe:
            info = f"Đang chạy bản không cài (bản đã cài: {os.path.dirname(exe)})"
        else:
            info = "Chưa cài vào máy (đang chạy trực tiếp file exe)" if installer.frozen() else ""
        about.AboutDialog(self, APP_TITLE, APP_VERSION, tab, info).exec()

    def _build_setup_menu(self):
        b = self.home.setup_btn
        m = QMenu(b)
        inst = installer.is_installed_copy()
        has = installer.installed_exe() is not None
        b.setText("Cài đặt" if inst or has else "Cài đặt vào máy")
        b.setProperty("highlight", "false" if inst or has else "true")
        b.setToolTip("Cài PDF Toolkit cho tài khoản này (không cần quyền admin), add-in Office")
        m.addAction(glyph_icon(0xE896), "Cài đặt / cập nhật vào máy...", self.install_dialog)
        m.addAction(glyph_icon(0xEA86), "Add-in Office (Word, Excel, PowerPoint)...",
                    lambda: self.install_dialog(office_only=True))
        m.addSeparator()
        a = m.addAction(glyph_icon(0xE74D), "Gỡ cài đặt...", self.uninstall_app)
        a.setEnabled(has)
        b.setMenu(m)

    def install_dialog(self, office_only=False):
        if not installer.frozen():
            QMessageBox.information(self, APP_NAME, "Chức năng cài đặt chỉ dùng được trong bản exe (PDF Toolkit.exe).")
            return
        dlg = InstallDialog(self, office_only)
        if dlg.exec() != QDialog.Accepted:
            return
        self._build_setup_menu()
        if dlg.exe and not installer.is_installed_copy():
            if QMessageBox.question(self, APP_NAME, "Đã cài xong. Mở bản đã cài ngay bây giờ?\n"
                                    "(Lần sau mở PDF Toolkit từ Start Menu / Desktop)") == QMessageBox.Yes:
                if self.confirm_discard_all():
                    import subprocess
                    subprocess.Popen([dlg.exe])
                    self._discard_all()
                    self.close()

    def uninstall_app(self):
        if QMessageBox.question(self, APP_NAME, "Gỡ PDF Toolkit khỏi máy?\nSẽ xoá shortcut, liên kết file PDF, "
                                "add-in Office, thư mục cài đặt và toàn bộ dữ liệu của app (file gần đây, cài đặt, "
                                "thư viện con dấu).\nFile PDF của bạn không bị ảnh hưởng.") != QMessageBox.Yes:
            return
        logs = []
        try:
            installer.uninstall(log=logs.append)
        except Exception as e:
            self.show_error(e)
            return
        self._build_setup_menu()
        QMessageBox.information(self, APP_NAME, "\n".join(logs))
        if installer.frozen() and "sẽ được xoá" in " ".join(logs):
            if self.confirm_discard_all():
                self._discard_all()
                self.close()

    def switch_language(self, lang):
        """Đổi ngôn ngữ: lưu lựa chọn rồi dựng lại cửa sổ (giữ file đang mở nếu có)."""
        on_home = self.stack.currentWidget() is self.home
        if lang == i18n.LANG or not self.confirm_discard_all():
            self.home.set_lang_buttons(i18n.LANG)
            return
        self.settings.setValue("lang", lang)
        i18n.set_lang(lang)
        paths = [d.pdf.path for d in self.docs
                 if d.pdf.doc is not None and d.pdf.path and os.path.exists(d.pdf.path)]
        active = self.pdf.path if self.pdf.path in paths else None
        w = MainWindow()
        w.setGeometry(self.geometry())
        if self.isMaximized():
            w.showMaximized()
        else:
            w.show()
        for p in paths:
            w.open_path(p)
        if active:
            w.open_path(active)                       # chuyển về tab đang xem
        if paths and on_home:
            w.show_home()
        _windows.append(w)
        self._discard_all()
        self.close()
        if self in _windows:
            _windows.remove(self)
        self.deleteLater()

    def add_recent(self, path):
        rec = [p for p in (self.settings.value("recent", []) or []) if p != path]
        if isinstance(rec, str):
            rec = [rec]
        self.settings.setValue("recent", ([path] + rec)[:8])

    def open_in_editor(self, path):
        self.open_path(path)

    def save(self):
        if not self.pdf.path or os.path.splitext(self.pdf.path)[1].lower() != ".pdf":
            return self.save_as()
        if not self.pdf.modified:
            self.status.showMessage("Không có thay đổi cần lưu", 4000)
            return True
        return self._save_to(self.pdf.path)

    def _signed_warning(self):
        """File có chữ ký số: lưu lại (ghi toàn bộ file) sẽ làm chữ ký mất hiệu lực."""
        if not self.pdf.has_signatures():
            return True
        return QMessageBox.question(
            self, APP_NAME, "File này có chữ ký số. Lưu thay đổi sẽ làm các chữ ký hiện có MẤT HIỆU LỰC.\n\n"
            "Vẫn lưu?") == QMessageBox.Yes

    def save_as(self):
        if self.pdf.path:
            default = os.path.splitext(self.pdf.path)[0] + "_moi.pdf"
        else:                                           # tài liệu mới (Ctrl+N)
            default = os.path.join(os.path.expanduser("~\\Documents"), "tai_lieu.pdf")
        path, _ = QFileDialog.getSaveFileName(self, "Lưu thành", default, "PDF (*.pdf)")
        return self._save_to(path) if path else False

    def _save_to(self, path):
        if not self._signed_warning():
            return False
        n = self.pdf.pending_redactions()
        if n:                                           # vùng đánh dấu chưa áp dụng: chữ bên dưới VẪN còn trong file
            r = QMessageBox.question(
                self, APP_NAME, f"Còn {n} vùng đánh dấu xoá chưa áp dụng - thông tin bên dưới VẪN còn trong file.\n\n"
                "Xoá vĩnh viễn các vùng này rồi lưu?", QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
            if r == QMessageBox.Cancel:
                return False
            if r == QMessageBox.Yes:
                self._redact_run(self.pdf.apply_redactions)
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            self.pdf.save(path)
        except PermissionError:
            self.show_error("Không ghi được file — file đang được mở bởi chương trình khác (Acrobat, trình duyệt...).\n"
                            "Hãy đóng file đó hoặc dùng 'Lưu thành'.")
            return False
        except Exception as e:
            self.show_error(f"Lỗi khi lưu:\n{e}")
            return False
        finally:
            QApplication.restoreOverrideCursor()
        self.status.showMessage(f"Đã lưu {path} ({fmt_size(self.pdf.file_size)})", 6000)
        self.update_ui()
        return True

    def showEvent(self, e):
        super().showEvent(e)
        if not getattr(self, "_caption_hidden", False):
            self._caption_hidden = True
            hide_caption_text(self)

    def closeEvent(self, e):
        if self.tool_page.busy():
            if QMessageBox.question(self, APP_NAME, "Công cụ đang chạy. Dừng và thoát?") != QMessageBox.Yes:
                e.ignore()
                return
            self.tool_page.worker.cancelled = True
            self.tool_page.worker.wait(5000)
        if self.confirm_discard_all():
            e.accept()
        else:
            e.ignore()

    # ------------------------------------------------------------------ thao tác trang
    def run_op(self, fn, select=None):
        try:
            res = fn()
        except Exception as e:
            self.show_error(e)
            return None
        self.refresh_thumbs(select if select is not None else (res if isinstance(res, list) else None))
        return res

    def delete_pages(self):
        # Delete khi đang chọn 1 vật trên trang: xoá vật đó, không xoá trang (trừ khi đang thao tác ở danh sách trang)
        if self.sender() is self.a_del and not self.thumbs.hasFocus() and self.delete_selected_object():
            return
        pages = self.selected_pages()
        if not pages:
            return
        if QMessageBox.question(self, APP_NAME, f"Xoá {len(pages)} trang: {', '.join(str(p + 1) for p in pages[:20])}"
                                f"{'...' if len(pages) > 20 else ''}?") != QMessageBox.Yes:
            return
        self.run_op(lambda: self.pdf.delete_pages(pages), [min(pages[0], self.pdf.page_count - len(pages) - 1)])

    def move_pages(self, delta):
        pages = self.selected_pages()
        self.run_op(lambda: self.pdf.move_pages(pages, delta))

    def on_reordered(self):
        order = [self.thumbs.item(i).data(Qt.UserRole) for i in range(self.thumbs.count())]
        if order == list(range(len(order))):
            return
        moved = [self.thumbs.row(i) for i in self.thumbs.selectedItems()]
        self.run_op(lambda: self.pdf.reorder(order), sorted(moved) or [0])

    def rotate(self, deg):
        pages = self.selected_pages()
        self.run_op(lambda: self.pdf.rotate(pages, deg), pages)

    def extract(self):
        pages = self.selected_pages()
        base = os.path.splitext(self.pdf.path or "tai_lieu")[0]
        name = f"{base}_trang_{pages[0] + 1}" + (f"-{pages[-1] + 1}" if len(pages) > 1 else "") + ".pdf"
        path, _ = QFileDialog.getSaveFileName(self, f"Tách {len(pages)} trang ra file mới", name, "PDF (*.pdf)")
        if path:
            try:
                self.pdf.extract(pages, path)
                self.status.showMessage(f"Đã lưu {len(pages)} trang vào {path}", 6000)
            except Exception as e:
                self.show_error(e)

    def thumb_menu(self, pos):
        if not self.pdf.doc:
            return
        m = QMenu(self)
        for a in (self.a_del, self.a_rotl, self.a_rotr, self.a_up, self.a_down, self.a_extract):
            m.addAction(a)
        m.addSeparator()
        m.addAction("Chèn file vào sau trang này...", self.add_files)
        m.exec(self.thumbs.mapToGlobal(pos))

    def toggle_grid(self):
        self.grid = self.a_grid.isChecked()
        if self.grid and not self.a_thumbs.isChecked():   # dạng lưới hiển thị ngay trong bảng trang
            self.a_thumbs.setChecked(True)
            self.toggle_thumbs()
        self.thumbs.set_grid(self.grid)
        self.scroll.setVisible(not self.grid)
        self.refresh_thumbs(self.selected_pages())

    def toggle_thumbs(self):
        """Thu gọn / mở bảng ảnh thu nhỏ bên trái (F4) - vùng xem trang rộng ra, trang tự vừa khung lại."""
        show = self.a_thumbs.isChecked()
        if not show and self.grid:                        # đang xem dạng lưới: thoát lưới trước
            self.a_grid.setChecked(False)
            self.toggle_grid()
        self.thumbs.setVisible(show)
        self._thumbs_btn_icon()

    def _thumbs_btn_icon(self):
        show = self.a_thumbs.isChecked()
        self.thumbs_btn.setIcon(tools_ui.glyph_icon(0xE76B if show else 0xE76C, "#334155"))
        self.thumbs_btn.setToolTip(i18n.tr("Thu gọn bảng trang (F4)") if show else i18n.tr("Mở bảng trang (F4)"))

    def undo(self):
        if self.pdf.undo():
            self.refresh_thumbs()

    def redo(self):
        if self.pdf.redo():
            self.refresh_thumbs()

    def on_current_changed(self, row):
        if row >= 0:
            self.current = row
            self.render_page()
            self._scroll_to_page(row)

    def goto(self, i):
        if self.pdf.doc and 0 <= i < self.pdf.page_count:
            self.select_pages([i])
            self._scroll_to_page(i)

    def _scroll_to_page(self, i, only_if_hidden=False):
        """Cuộn để trang i lên đầu vùng xem (only_if_hidden: chỉ khi trang đó đang khuất hẳn)."""
        st = self.strip
        if self.grid or not 0 <= i < len(st.sizes):
            return
        sb, vh = self.scroll.verticalScrollBar(), self.scroll.viewport().height()
        r = st.page_rect(i)
        if only_if_hidden and r.bottom() > sb.value() and r.top() < sb.value() + vh:
            return
        self._auto_scroll = True
        sb.setValue(r.top())
        self._auto_scroll = False

    def _on_scrolled(self, value):
        """Cuộn liền mạch: trang nằm ở 1/3 trên của vùng xem thành trang hiện tại (thanh trang thu nhỏ theo theo)."""
        st = self.strip
        if self._auto_scroll or self.grid or not self.pdf.doc or not st.sizes:
            return
        c = self.canvas
        if c.drag or c.obj_drag or c.stroke or c.poly or QApplication.mouseButtons() != Qt.NoButton:
            return                                      # đang kéo / vẽ trên trang hiện tại -> giữ nguyên
        i = st.page_at(value + self.scroll.viewport().height() // 3)
        if i == self.current or i >= self.pdf.page_count:
            return
        self.current = i
        self.thumbs.blockSignals(True)
        self.thumbs.clearSelection()
        self.thumbs.setCurrentRow(i)
        self.thumbs.blockSignals(False)
        self.thumbs.scrollToItem(self.thumbs.item(i))
        self.render_page()

    def scope_pages(self, idx):
        if idx == 1:
            return [self.current]
        if idx == 2:
            return self.selected_pages()
        return None

    # ------------------------------------------------------------------ nén
    def compress(self):
        size = self.pdf.file_size if not self.pdf.modified else len(self.pdf.doc.tobytes(garbage=1, deflate=True))
        dlg = CompressDialog(self, size)
        if dlg.exec() != QDialog.Accepted:
            return
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            self.status.showMessage("Đang nén...")
            QApplication.processEvents()
            before, after = self.pdf.compress(dlg.level())
        except Exception as e:
            self.show_error(f"Lỗi khi nén:\n{e}")
            return
        finally:
            QApplication.restoreOverrideCursor()
        self.refresh_thumbs()
        pct = (1 - after / before) * 100 if before else 0
        QMessageBox.information(self, "Nén PDF", f"Trước: {fmt_size(before)}\nSau:   {fmt_size(after)}\n"
                                f"Giảm {pct:.0f}%\n\nBấm Lưu / Lưu thành để ghi ra file.")

    # ------------------------------------------------------------------ sửa chữ
    def set_mode(self, mode):
        if self.grid and mode != "view":
            self.a_grid.setChecked(False)
            self.toggle_grid()
        if mode == self.mode:                           # bấm lại chế độ đang dùng (vd Esc): bỏ chọn vật
            self.sel_obj = None
        self.mode = mode
        self.tool_actions[mode].setChecked(True)
        self.render_page()

    def run_edit(self, fn, all_pages=False):
        try:
            res = fn()
        except Exception as e:
            self.show_error(e)
            return None
        if all_pages:
            self.refresh_thumbs(self.selected_pages())
        else:
            self.refresh_thumb(self.current)
            self.render_page()
        return res

    def on_line_clicked(self, idx):
        line = self.canvas.lines[idx]
        dlg = RichTextDialog(self, "Sửa dòng chữ", line)
        if dlg.exec() != QDialog.Accepted:
            return
        if dlg.delete or dlg.rich.is_empty():
            self.run_edit(lambda: self.pdf.edit_line_rich(self.current, line, "", []))
            return
        if not dlg.rich.edit.document().isModified() and not dlg.fit.isChecked():
            return
        body, fams = dlg.rich.to_html()
        self.run_edit(lambda: self.pdf.edit_line_rich(self.current, line, body, fams, dlg.rich.first_align(),
                                                      dlg.fit.isChecked()))

    def on_block_clicked(self, idxs):
        """Sửa cả khối chữ (đoạn văn nhiều dòng): nội dung các dòng nối lại, chữ chảy lại trong bề ngang khối."""
        lines = [self.canvas.lines[k] for k in idxs]
        runs = []
        for ln in lines:
            r = RichTextDialog.line_runs(ln)
            if runs and r:                              # nối dòng: thêm dấu cách (bỏ gạch nối cuối dòng)
                t = runs[-1][0]
                runs[-1] = (t[:-1] if t.endswith("-") else t + " ",) + runs[-1][1:]
            runs += r
        dlg = RichTextDialog(self, "Sửa khối chữ", lines[0], runs=runs)
        if dlg.exec() != QDialog.Accepted:
            return
        if dlg.delete or dlg.rich.is_empty():
            self.run_edit(lambda: self.pdf.delete_lines(self.current, lines))
            return
        if not dlg.rich.edit.document().isModified():
            return
        body, fams = dlg.rich.to_html()
        self.run_edit(lambda: self.pdf.edit_block_rich(self.current, lines, body, fams))

    # ----- quét từ máy scan -----
    def scan_to_pdf(self, insert=False):
        import scanner
        insert = insert and self.pdf.doc is not None and self.in_editor()
        dlg = ScanDialog(self, insert)
        if dlg.exec() != QDialog.Accepted:
            return
        source, color, dpi = dlg.values()
        for k, v in (("scan_source", source), ("scan_color", color), ("scan_dpi", dpi)):
            self.settings.setValue(k, v)

        def more(n):
            return QMessageBox.question(self, APP_NAME, f"Đã quét {n} trang.\n\nĐặt trang tiếp theo lên mặt kính rồi "
                                        "bấm Yes để quét tiếp, No để kết thúc.") == QMessageBox.Yes
        try:
            pages = scanner.scan_pages(source, color, dpi, more)
        except scanner.ScanError as e:
            self.show_error(e)
            return
        if not pages:
            return
        try:
            if insert:
                at = self.current + 1
                n = self.run_edit(lambda: self.pdf.insert_scan(pages, dpi, at), all_pages=True)
                if n:
                    self.refresh_thumbs([at])
                    self.status.showMessage(f"Đã chèn {n} trang quét sau trang {at}", 6000)
            else:
                target = PdfDocument() if self.docs else self.pdf
                target.new_from_scan(pages, dpi)
                self.editor_area.setCurrentWidget(self.splitter)
                self.ribbon.setCurrentIndex(0)
                self._add_tab(DocTab(target))
                self.show_editor()
                self._update_tabs()
                self.status.showMessage(f"Đã quét {len(pages)} trang - bấm Lưu để ghi ra file, OCR để nhận dạng chữ",
                                        8000)
        except Exception as e:
            self.show_error(f"Không tạo được PDF từ ảnh quét:\n{e}")
        finally:
            import shutil
            shutil.rmtree(os.path.dirname(pages[0]), ignore_errors=True)

    def detect_fields(self):
        if not self.pdf.doc:
            return
        counts = self._redact_run(self.pdf.detect_fields)
        if not counts:
            if counts is not None:
                QMessageBox.information(self, APP_NAME, "Không thấy chỗ trống nào để tạo ô form (dòng gạch dưới / chấm "
                                        "chấm, đường kẻ trống, ô vuông, khung trống).\n\nFile scan cần OCR trước, hoặc "
                                        "dùng nút Tạo ô form để kéo ô bằng tay.")
            return
        t = counts.get("text", 0) + counts.get("multiline", 0)
        self.toast("Đã tạo ô form", f"{t} ô chữ, {counts.get('check', 0)} ô đánh dấu - Ctrl+Z để hoàn tác")
        self.status.showMessage("Kiểm tra lại các ô; ô thừa: chế độ Tạo ô form, chuột phải để xoá", 8000)

    # ----- xoá vĩnh viễn thông tin (redaction) -----
    def _redact_run(self, fn):
        try:
            n = fn()
        except Exception as e:
            self.show_error(e)
            return None
        self.refresh_thumbs()
        return n

    def redact_find(self):
        if not self.pdf.doc:
            return
        dlg = RedactFindDialog(self, self.current)
        if dlg.exec() != QDialog.Accepted:
            return
        terms, pats, whole, fill, pages = dlg.values()
        if pages is None:                               # tất cả các trang
            pages = range(self.pdf.page_count)
        found = {}
        for pno in pages:
            rs = pdf_core.find_redact_rects(self.pdf.doc[pno], terms, pats, whole)
            if rs:
                found[pno] = rs
        n = sum(len(v) for v in found.values())
        if not n:
            QMessageBox.information(self, APP_NAME, "Không tìm thấy chữ cần xoá.\n\nFile scan (chữ là ảnh) cần OCR "
                                    "trước, hoặc kéo vùng bằng nút Đánh dấu xoá.")
            return
        self.settings.setValue("redact_fill", fill)
        self._redact_run(lambda: self.pdf.mark_redactions(found, fill))
        self.goto(min(found))
        self.status.showMessage(f"Đã đánh dấu {n} chỗ trên {len(found)} trang (khung đỏ) - xem lại rồi bấm "
                                "Xoá vĩnh viễn", 8000)

    def redact_apply(self):
        n = self.pdf.pending_redactions() if self.pdf.doc else 0
        if not n:
            self.status.showMessage("Chưa có vùng nào được đánh dấu xoá - dùng Đánh dấu xoá hoặc Tìm & đánh dấu", 5000)
            return
        if QMessageBox.question(self, APP_NAME, f"Xoá vĩnh viễn {n} vùng đã đánh dấu?\n\nChữ, ảnh và hình vẽ trong "
                                "vùng bị xoá hẳn khỏi file (không copy / tìm lại được). Trước khi lưu vẫn hoàn tác "
                                "được bằng Ctrl+Z.") != QMessageBox.Yes:
            return
        if self._redact_run(self.pdf.apply_redactions):
            self.toast("Đã xoá vĩnh viễn", f"{n} vùng - bấm Lưu để ghi ra file")

    def redact_clear(self):
        n = self._redact_run(self.pdf.clear_redactions) if self.pdf.doc else 0
        if n:
            self.status.showMessage(f"Đã bỏ {n} vùng đánh dấu xoá", 4000)

    def on_rect_drawn(self, r):
        rect = (r.left(), r.top(), r.right(), r.bottom())
        tiny = r.width() < 4 and r.height() < 4
        tool = self.mode
        if tool == "redact":
            if tiny:
                return
            fill = self.settings.value("redact_fill", "black")
            self.run_edit(lambda: self.pdf.mark_redactions({self.current: [rect]}, fill, disp=True))
            self.status.showMessage(f"{self.pdf.pending_redactions()} vùng đang chờ xoá - bấm Xoá vĩnh viễn để áp "
                                    "dụng", 6000)
            return
        if tool == "text":
            dlg = RichTextDialog(self, "Thêm chữ")
            if dlg.exec() != QDialog.Accepted or dlg.rich.is_empty():
                return
            body, fams = dlg.rich.to_html()
            scale = self.run_edit(lambda: self.pdf.add_rich_text(self.current, rect, body, fams))
            if scale is not None and scale < 0.99:
                self.status.showMessage(f"Khung không đủ chỗ - chữ đã được thu nhỏ còn {scale * 100:.0f}%", 6000)
        elif tool == "image":
            self.insert_image_dialog(rect)
        elif tool == "link":
            self.insert_link(rect, tiny)
        elif tool == "highlight":
            if tiny:
                return
            n = self.run_edit(lambda: self.pdf.add_highlight(self.current, rect))
            if n == 0:
                self.status.showMessage("Không có chữ trong vùng - đã tô sáng cả vùng", 5000)
        elif tool in ("rect", "ellipse"):
            if tiny:
                return
            dlg = ShapeDialog(self, tool)
            if dlg.exec() == QDialog.Accepted:
                v = dlg.values()
                self.run_edit(lambda: self.pdf.add_shape(self.current, tool, rect[:2], rect[2:], **v))
        elif tool == "field":
            if r.width() < 6 or r.height() < 6:
                return
            dlg = FieldDialog(self, self.pdf.field_names())
            if dlg.exec() == QDialog.Accepted:
                v = dlg.values()
                name = self.run_edit(lambda: self.pdf.add_field(self.current, rect, **v))
                if name:
                    self.status.showMessage(f"Đã tạo ô form '{name}'", 5000)
        elif tool == "handsign":
            dlg = HandSignDialog(self)
            if dlg.exec() == QDialog.Accepted and dlg.png:
                if tiny:                                 # chỉ bấm: chữ ký rộng khoảng 50 mm
                    w = 50 * 72 / 25.4
                    rect = (rect[0], rect[1], rect[0] + w, rect[1] + w / max(dlg.ratio, 0.1))
                self.run_edit(lambda: self.pdf.insert_image(self.current, rect, stream=dlg.png))
        elif tool == "stamp":
            dlg = StampDialog(self)
            if dlg.exec() == QDialog.Accepted and dlg.png:
                if tiny or r.width() < 20:                   # chỉ bấm: con dấu chữ ~55 mm, dấu tròn ~38 mm
                    w = (38 if dlg.shape in ("round", "image") else 55) * 72 / 25.4
                    h = w / max(dlg.ratio, 0.1)
                    rect = (rect[0] - w / 2, rect[1] - h / 2, rect[0] + w / 2, rect[1] + h / 2)
                self.run_edit(lambda: self.pdf.insert_image(self.current, rect, stream=dlg.png))
                self.status.showMessage("Đã đóng dấu - kéo để di chuyển, kéo góc dưới phải để đổi cỡ, "
                                        "chuột phải để xoá / nhân bản", 8000)
        elif tool == "snapshot":
            if tiny:
                return
            self.snap_region = (self.current, QRectF(r))          # giữ khung: kéo / đổi cỡ, chuột phải ra menu
            self.canvas.set_region(r)
            self.snapshot_region(rect)
        elif tool in ("underline", "strikeout"):
            if tiny:
                return
            author = os.environ.get("USERNAME", "")
            n = self.run_edit(lambda: self.pdf.add_markup(self.current, tool, disp_rect=rect, author=author))
            if n == 0:
                self.status.showMessage("Không có chữ trong vùng - đã đánh dấu cả vùng", 5000)
        elif tool in ("textbox", "typewriter"):
            dlg = TextBoxDialog(self, "Hộp chữ" if tool == "textbox" else "Máy chữ",
                                border_default=0 if tool == "typewriter" else None)
            if tool == "typewriter":
                dlg.fill.setCurrentIndex(0)
            if dlg.exec() != QDialog.Accepted:
                return
            v = dlg.values()
            if tiny or r.width() < 20 or r.height() < 10:                # chỉ bấm: hộp vừa nội dung
                w, h = PdfDocument.textbox_size(v["text"], v["fontsize"])
                rect = (rect[0], rect[1], rect[0] + w, rect[1] + h)
            author = os.environ.get("USERNAME", "")
            self.run_edit(lambda: self.pdf.add_textbox(self.current, rect, v["text"], v["fontsize"], v["color"],
                                                       v["fill"], v["border"], author=author))
        elif tool == "cloud":
            if tiny:
                return
            dlg = ShapeDialog(self, "cloud")
            if dlg.exec() == QDialog.Accepted:
                v = dlg.values()
                x0, y0, x1, y1 = rect
                self.run_edit(lambda: self.pdf.add_poly(self.current, [(x0, y0), (x1, y0), (x1, y1), (x0, y1)],
                                                        "polygon", v["stroke"], max(v["width"], 1), v["fill"], clouds=2))
        elif tool == "sign":
            if r.width() < 30 or r.height() < 15:
                self.show_error("Vùng chữ ký quá nhỏ - hãy kéo một khung rộng hơn (khoảng 6 x 2 cm).")
                return
            self.sign(rect)

    def on_segment_drawn(self, a, b):
        if (a - b).manhattanLength() < 4:
            return
        if self.mode == "callout":                     # a = điểm mũi tên chỉ vào, b = chỗ đặt hộp chữ
            dlg = TextBoxDialog(self, "Chú thích mũi tên", border_default=1)
            if dlg.exec() != QDialog.Accepted:
                return
            v = dlg.values()
            w, h = PdfDocument.textbox_size(v["text"], v["fontsize"])
            left = b.x() >= a.x()                       # hộp nằm phía nào của mũi tên
            x0 = b.x() if left else b.x() - w
            box = (x0, b.y() - h / 2, x0 + w, b.y() + h / 2)
            knee = (b.x() - 12 if left else b.x() + 12, b.y())
            end = (box[0] if left else box[2], b.y())
            author = os.environ.get("USERNAME", "")
            self.run_edit(lambda: self.pdf.add_textbox(self.current, box, v["text"], v["fontsize"], v["color"],
                                                       v["fill"], v["border"], callout=[(a.x(), a.y()), knee, end],
                                                       author=author))
            return
        dlg = ShapeDialog(self, self.mode)
        if dlg.exec() == QDialog.Accepted:
            v = dlg.values()
            v.pop("fill")
            self.run_edit(lambda: self.pdf.add_shape(self.current, "line", (a.x(), a.y()), (b.x(), b.y()),
                                                     arrow=self.mode == "arrow", **v))

    def on_point_clicked(self, x, y):
        author = os.environ.get("USERNAME", "")
        if self.mode == "eraser":
            xref = self.pdf.annot_at(self.current, (x, y))
            if xref:
                self.run_edit(lambda: self.pdf.delete_annot(self.current, xref))
            else:
                self.status.showMessage("Không có ghi chú / nét vẽ tại đây", 3000)
            return
        if self.mode == "caret":
            text, ok = QInputDialog.getMultiLineText(self, "Đề xuất chèn chữ", "Nội dung đề xuất chèn vào:")
            if ok and text.strip():
                self.run_edit(lambda: self.pdf.add_caret(self.current, (x, y), text.strip(), author))
            return
        if self.mode == "attach":
            path, _ = QFileDialog.getOpenFileName(self, "Chọn file để ghim lên trang", self._dir(), "Tất cả file (*)")
            if path:
                self.run_edit(lambda: self.pdf.add_file_annot(self.current, (x, y), path, author))
            return
        if self.mode != "note":
            return
        text, ok = QInputDialog.getMultiLineText(self, "Ghi chú", "Nội dung ghi chú:")
        if ok and text.strip():
            author = os.environ.get("USERNAME", "")
            self.run_edit(lambda: self.pdf.add_note(self.current, (x, y), text.strip(), author))

    # ------------------------------------------------------------------ nhận xét: bút, đa giác, chọn chữ
    def _pen(self):
        s = self.settings
        try:
            color = tuple(float(x) for x in str(s.value("pen/color", "0.86,0.15,0.15")).split(","))
        except ValueError:
            color = (0.86, 0.15, 0.15)
        return color, float(s.value("pen/width", 1.5))

    def pen_style(self):
        dlg = ShapeDialog(self, "pen")
        color, width = self._pen()
        dlg.stroke.set_color(color)
        dlg.width.setValue(width)
        dlg.width.setMinimum(0.5)
        if dlg.exec() == QDialog.Accepted:
            v = dlg.values()
            self.settings.setValue("pen/color", ",".join(f"{c:.3f}" for c in v["stroke"]))
            self.settings.setValue("pen/width", v["width"])

    def on_stroke_drawn(self, pts):
        if self.mode == "marker":                       # dạ quang: nét to, vàng, trong suốt
            self.run_edit(lambda: self.pdf.add_ink(self.current, [pts], (1, 0.85, 0.1), 10, 0.45))
        else:
            color, width = self._pen()
            self.run_edit(lambda: self.pdf.add_ink(self.current, [pts], color, width))

    def on_poly_drawn(self, pts):
        kind = "polygon" if self.mode == "polygon" else "polyline"
        if kind == "polygon" and len(pts) < 3:
            return
        dlg = ShapeDialog(self, kind)
        if dlg.exec() == QDialog.Accepted:
            v = dlg.values()
            self.run_edit(lambda: self.pdf.add_poly(self.current, pts, kind, v["stroke"], max(v["width"], 0.5),
                                                    v["fill"] if kind == "polygon" else None))

    def copy_selection(self):
        if self.mode in self.MOVE_MODES and self._sel_index() >= 0:
            self.copy_object()
            return
        if self.mode == "snapshot" and self.snap_region:
            self.snapshot_region(self._region_rect())
            return
        if self.mode == "select" and self.canvas.sel_range():
            QApplication.clipboard().setText(self.canvas.selected_text())
            self.status.showMessage("Đã sao chép chữ", 3000)

    def select_all_text(self):
        if self.mode != "select":
            self.set_mode("select")
        self.canvas.select_all()

    def on_selection_menu(self, gpos):
        if not self.canvas.sel_range():
            return
        m = QMenu(self)
        a_copy = m.addAction(glyph_icon(0xE8C8), "Sao chép")
        m.addSeparator()
        acts = {m.addAction(glyph_icon(icon), text): kind for kind, text, icon in (
            ("highlight", "Tô sáng", 0xE7E6), ("underline", "Gạch chân", 0xE8DC),
            ("strikeout", "Gạch ngang", 0xEDE0), ("squiggly", "Gạch sóng", 0xE8DC))}
        a_replace = m.addAction(glyph_icon(0xE8AC), "Đề xuất thay chữ...")
        a_note = m.addAction(glyph_icon(0xE90A), "Tô sáng kèm ghi chú...")
        m.addSeparator()
        a_link = m.addAction(glyph_icon(0xE71B), "Tạo liên kết...")
        a_all = m.addAction("Chọn hết chữ trang  (Ctrl+A)")
        act = m.exec(gpos)
        if act is None:
            return
        rects = self.canvas.selected_rects()
        author = os.environ.get("USERNAME", "")
        if act is a_copy:
            self.copy_selection()
        elif act in acts:
            self.run_edit(lambda: self.pdf.add_markup(self.current, acts[act], word_rects=rects, author=author))
        elif act is a_note:
            text, ok = QInputDialog.getMultiLineText(self, "Ghi chú", "Nội dung ghi chú cho đoạn chữ:")
            if ok:
                self.run_edit(lambda: self.pdf.add_markup(self.current, "highlight", word_rects=rects, author=author,
                                                          note=text.strip()))
        elif act is a_replace:
            text, ok = QInputDialog.getMultiLineText(self, "Đề xuất thay chữ",
                                                     f"Thay \"{self.canvas.selected_text()[:80]}\" bằng:")
            if ok and text.strip():
                self.run_edit(lambda: self.pdf.add_markup(self.current, "strikeout", word_rects=rects, author=author,
                                                          note="Thay bằng: " + text.strip()))
        elif act is a_link:
            page = self.pdf.doc[self.current]
            box = pymupdf.Rect()
            for r in rects:
                box |= r
            disp = box * page.rotation_matrix
            disp.normalize()
            self.insert_link(tuple(disp), False)
        elif act is a_all:
            self.canvas.select_all()

    def snapshot_region(self, rect):
        """Chụp vùng trang (kể cả ghi chú) ở ~200 dpi vào clipboard."""
        page = self.pdf.doc[self.current]
        clip = pdf_core.disp_to_unrot(page, pymupdf.Rect(rect))
        pix = page.get_pixmap(matrix=pymupdf.Matrix(200 / 72, 200 / 72).prerotate(page.rotation), clip=clip, alpha=False)
        QApplication.clipboard().setImage(pix_to_qimage(pix).copy())
        self.status.showMessage(f"Đã chụp vùng {pix.width}x{pix.height} px vào clipboard - dán (Ctrl+V) vào Word, "
                                "Outlook, Zalo...", 6000)
        self.toast("Đã lưu ảnh vào bộ nhớ tạm", "Dán (Ctrl+V) vào Word, Outlook, Zalo...")

    # ----- Chụp vùng: khung vùng chọn + menu chuột phải -----
    def _region_rect(self):
        r = self.snap_region[1]
        return (r.left(), r.top(), r.right(), r.bottom())

    def on_region_changed(self, r):
        self.snap_region = None if r is None else (self.current, QRectF(r))

    def on_region_menu(self, gpos):
        if not self.snap_region:
            return
        rect = self._region_rect()
        m = QMenu(self)
        m.addAction(glyph_icon(0xE8C8), "Sao chép ảnh vùng chọn\tCtrl+C", lambda: self.snapshot_region(rect))
        m.addAction(glyph_icon(0xE8C8), "Sao chép chữ trong vùng", lambda: self.region_copy_text(rect))
        m.addAction(glyph_icon(0xE74E), "Lưu vùng thành ảnh...", lambda: self.region_save_image(rect))
        m.addAction(glyph_icon(0xE749), "In vùng chọn...", lambda: self.region_print(rect))
        m.addSeparator()
        m.addAction(glyph_icon(0xE8A4), "Thêm bookmark...", lambda: self.region_bookmark(rect))
        m.addAction(glyph_icon(0xE90A), "Thêm ghi chú...", lambda: self.region_note(rect))
        m.addAction(glyph_icon(0xE71B), "Thêm liên kết...", lambda: self.insert_link(rect, False))
        m.addAction(glyph_icon(0xEB9F), "Chèn ảnh vào vùng...", lambda: self.region_image(rect))
        m.addAction(glyph_icon(0xE7E6), "Tô sáng vùng",
                    lambda: self.run_edit(lambda: self.pdf.add_highlight(self.current, rect)))
        m.addAction(glyph_icon(0xE739), "Che trắng vùng",
                    lambda: self.run_edit(lambda: self.pdf.add_shape(self.current, "rect", rect[:2], rect[2:],
                                                                     stroke=(1, 1, 1), fill=(1, 1, 1), width=0)))
        m.addAction(glyph_icon(0xE8C6), "Đánh dấu xoá vĩnh viễn",
                    lambda: self.run_edit(lambda: self.pdf.mark_redactions(
                        {self.current: [rect]}, self.settings.value("redact_fill", "black"), disp=True)))
        m.addSeparator()
        m.addAction(glyph_icon(0xE7A8), "Cắt trang theo vùng", lambda: self.region_crop(rect))
        m.addAction(glyph_icon(0xE8D2), "Nhận dạng chữ (OCR) trong vùng...", lambda: self.region_ocr(rect))
        m.addAction(glyph_icon(0xE71E), "Phóng to vùng", lambda: self.region_zoom(QRectF(self.snap_region[1])))
        m.addSeparator()
        m.addAction("Bỏ chọn vùng\tEsc", self.canvas.clear_region)
        m.exec(gpos)

    def _region_pix(self, rect, dpi=300):
        page = self.pdf.doc[self.current]
        clip = pdf_core.disp_to_unrot(page, pymupdf.Rect(rect))
        return page.get_pixmap(matrix=pymupdf.Matrix(dpi / 72, dpi / 72).prerotate(page.rotation), clip=clip,
                               alpha=False)

    def region_copy_text(self, rect):
        page = self.pdf.doc[self.current]
        text = page.get_text("text", clip=pdf_core.disp_to_unrot(page, pymupdf.Rect(rect)), sort=True).strip()
        if not text:
            QMessageBox.information(self, APP_NAME, i18n.tr("Vùng này không có chữ (có thể là ảnh scan) - "
                                                            "dùng 'Nhận dạng chữ (OCR) trong vùng'."))
            return
        QApplication.clipboard().setText(text)
        self.toast("Đã sao chép chữ vào bộ nhớ tạm", "Dán (Ctrl+V) vào Word, Outlook, Zalo...")

    def region_save_image(self, rect):
        base = os.path.splitext(os.path.basename(self.pdf.path or "trang"))[0]
        path, _ = QFileDialog.getSaveFileName(self, "Lưu vùng thành ảnh",
                                              os.path.join(self._dir(), f"{base}_trang{self.current + 1}.png"),
                                              "Ảnh PNG (*.png);;Ảnh JPEG (*.jpg)")
        if not path:
            return
        if not pix_to_qimage(self._region_pix(rect)).save(path, quality=92):
            self.show_error(PdfError(i18n.tr("Không lưu được ảnh:") + "\n" + path))
            return
        self.toast("Đã lưu ảnh", os.path.basename(path))

    def region_print(self, rect):
        from PySide6.QtPrintSupport import QPrintDialog, QPrinter
        pr = QPrinter(QPrinter.HighResolution)
        dlg = QPrintDialog(pr, self)
        dlg.setWindowTitle(i18n.tr("In vùng chọn"))
        if dlg.exec() != QDialog.Accepted:
            return
        self._print_region(pr, rect)

    def _print_region(self, pr, rect):
        """Vẽ vùng chọn (300 dpi) vừa khổ giấy, giữ tỉ lệ, căn giữa trên đầu trang giấy."""
        img = pix_to_qimage(self._region_pix(rect))
        p = QPainter(pr)
        area = pr.pageLayout().paintRectPixels(pr.resolution())
        area = QRectF(0, 0, area.width(), area.height())
        k = min(area.width() / img.width(), area.height() / img.height(), pr.resolution() / 300)
        w, h = img.width() * k, img.height() * k
        p.drawImage(QRectF((area.width() - w) / 2, 0, w, h), img)
        p.end()

    def region_bookmark(self, rect):
        page = self.pdf.doc[self.current]
        first = page.get_text("text", clip=pdf_core.disp_to_unrot(page, pymupdf.Rect(rect))).strip().split("\n")[0]
        title, ok = QInputDialog.getText(self, i18n.tr("Thêm bookmark"), i18n.tr("Tên bookmark:"),
                                         text=first[:80] or i18n.tr("Trang {0}").format(self.current + 1))
        if ok and title.strip():
            self.run_edit(lambda: self.pdf.add_bookmark(self.current, title.strip(), rect))
            self.toast("Đã thêm bookmark", title.strip())

    def region_note(self, rect):
        text, ok = QInputDialog.getMultiLineText(self, "Ghi chú", "Nội dung ghi chú:")
        if ok and text.strip():
            author = os.environ.get("USERNAME", "")
            self.run_edit(lambda: self.pdf.add_note(self.current, rect[:2], text.strip(), author))

    def region_image(self, rect):
        self.insert_image_dialog(rect)

    def insert_image_dialog(self, rect):
        """Hộp thoại chèn ảnh (chọn file + xem trước + định dạng: nội dung trang / con dấu) rồi chèn vào vùng."""
        dlg = ImageInsertDialog(self, self.settings.value("last_image_dir", ""),
                                self.settings.value("image_as_stamp", False, type=bool))
        QTimer.singleShot(0, dlg.browse)                 # mở ngay hộp chọn file (giống trước đây)
        if dlg.exec() != QDialog.Accepted or not dlg.path:
            return
        path, stamp = dlg.path, dlg.as_stamp
        self.settings.setValue("last_image_dir", os.path.dirname(path))
        self.settings.setValue("image_as_stamp", stamp)
        if stamp:
            author = os.environ.get("USERNAME", "")
            self.run_edit(lambda: self.pdf.insert_image_stamp(self.current, rect, path, author=author))
        else:
            self.run_edit(lambda: self.pdf.insert_image(self.current, rect, path))

    def region_crop(self, rect):
        if QMessageBox.question(self, APP_NAME, i18n.tr("Cắt trang {0} chỉ giữ lại vùng đang chọn?\n"
                                                        "(Ctrl+Z để hoàn tác)").format(self.current + 1)) \
                != QMessageBox.Yes:
            return
        self.snap_region = None
        self.run_edit(lambda: self.pdf.crop_to(self.current, rect))
        self.refresh_thumbs([self.current])

    def region_ocr(self, rect):
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.status.showMessage(i18n.tr("Đang nhận dạng chữ trong vùng..."))
        try:
            pix = self._region_pix(rect, 300)
            tmp = pymupdf.open()
            pg = tmp.new_page(width=pix.width * 72 / 300, height=pix.height * 72 / 300)
            pg.insert_image(pg.rect, pixmap=pix)
            tp = pg.get_textpage_ocr(flags=0, language="vie", dpi=300, full=True, tessdata=pdf_tools.tessdata_dir())
            text = pg.get_text("text", textpage=tp, sort=True).strip()
        except Exception as e:
            QApplication.restoreOverrideCursor()
            self.status.clearMessage()
            self.show_error(e)
            return
        QApplication.restoreOverrideCursor()
        self.status.clearMessage()
        TextResultDialog(self, i18n.tr("Chữ nhận dạng được (OCR)"), text).exec()

    def region_zoom(self, r):
        """Phóng to để vùng chọn vừa khung xem, cuộn tới vùng đó."""
        vp = self.scroll.viewport()
        self.zoom = max(0.2, min(4.0, (vp.width() - 80) / max(r.width(), 1), (vp.height() - 60) / max(r.height(), 1)))
        self.fit_width = False
        self.render_page()
        pr = self.strip.page_rect(self.current)
        self.scroll.verticalScrollBar().setValue(int(pr.top() + 20 + r.top() * self.zoom - 30))
        self.scroll.horizontalScrollBar().setValue(int(pr.left() + 20 + r.left() * self.zoom - 40))

    def toast(self, title, detail="", ms=2600):
        """Thẻ thông báo nổi ở giữa phía dưới vùng xem trang, tự mờ dần (không chặn thao tác)."""
        from PySide6.QtCore import QPropertyAnimation
        from PySide6.QtWidgets import QGraphicsOpacityEffect
        host = self.centralWidget()
        old = getattr(self, "_toast", None)
        if old is not None:
            try:
                old.deleteLater()                  # thẻ trước còn hiện (chụp liên tiếp): thay bằng thẻ mới
            except RuntimeError:
                pass                               # đã tự xoá sau khi mờ hết
        t = QLabel(host)
        self._toast = t
        t.setObjectName("toast")
        t.setTextFormat(Qt.RichText)
        t.setText(f"<span style='color:#16a34a; font-size:13pt'>✔</span>&nbsp; <b>{html.escape(i18n.tr(title))}</b>"
                  + (f"<br><span style='color:#52525b'>{html.escape(i18n.tr(detail))}</span>" if detail else ""))
        t.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        t.adjustSize()
        t.move((host.width() - t.width()) // 2, host.height() - t.height() - 48)
        eff = QGraphicsOpacityEffect(t)
        eff.setOpacity(1.0)
        t.setGraphicsEffect(eff)
        t.show()
        t.raise_()
        anim = QPropertyAnimation(eff, b"opacity", t)
        anim.setDuration(600)
        anim.setStartValue(1.0)
        anim.setEndValue(0.0)
        anim.finished.connect(t.deleteLater)
        QTimer.singleShot(ms, t, anim.start)

    def show_comments(self):
        if self.pdf.doc:
            CommentsDialog(self).exec()

    # ------------------------------------------------------------------ liên kết hàng loạt, thao tác trang
    def auto_links(self):
        n = self.run_edit(lambda: self.pdf.auto_links(), all_pages=True)
        if n is not None:
            QMessageBox.information(self, APP_NAME, f"Đã tạo {n} liên kết từ địa chỉ web / email trong file." if n else
                                    "Không tìm thấy địa chỉ web / email nào chưa có liên kết.")

    def remove_web_links(self):
        if QMessageBox.question(self, APP_NAME, "Xoá mọi liên kết tới trang web / email trong file?") != QMessageBox.Yes:
            return
        n = self.run_edit(lambda: self.pdf.remove_web_links(), all_pages=True)
        if n is not None:
            self.status.showMessage(f"Đã xoá {n} liên kết web", 5000)

    def _reload_pages(self, select=None):
        self.refresh_thumbs(select or [self.current])
        self.update_ui()

    def duplicate_pages(self):
        pages = self.selected_pages()
        try:
            new = self.pdf.duplicate_pages(pages)
        except Exception as e:
            self.show_error(e)
            return
        self._reload_pages(new)
        self.status.showMessage(f"Đã nhân bản {len(new)} trang  (Ctrl+Z để hoàn tác)", 5000)

    def swap_pages(self):
        pages = self.selected_pages()
        if len(pages) != 2:
            QMessageBox.information(self, APP_NAME, "Hãy chọn đúng 2 trang ở danh sách trang bên trái "
                                    "(giữ Ctrl rồi bấm) để hoán đổi.")
            return
        try:
            self.pdf.swap_pages(*pages)
        except Exception as e:
            self.show_error(e)
            return
        self._reload_pages(pages)

    def replace_pages(self):
        pages = self.selected_pages()
        path, _ = QFileDialog.getOpenFileName(self, "Chọn file PDF chứa trang thay thế", self._dir(), "PDF (*.pdf)")
        if not path:
            return
        start, ok = QInputDialog.getInt(self, "Thay trang",
                                        f"Thay {len(pages)} trang đang chọn bằng các trang của file mới,\n"
                                        "bắt đầu từ trang số:", 1, 1, 100000)
        if not ok:
            return
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            self.pdf.replace_pages(pages, path, start - 1)
        except PdfError as e:
            QApplication.restoreOverrideCursor()
            if str(e) == "NEED_PASSWORD":
                pw = self._ask_password(path)
                if pw is None:
                    return
                try:
                    self.pdf.replace_pages(pages, path, start - 1, pw)
                except Exception as e2:
                    self.show_error(e2)
                    return
            else:
                self.show_error(e)
                return
        except Exception as e:
            self.show_error(e)
            return
        finally:
            if QApplication.overrideCursor():
                QApplication.restoreOverrideCursor()
        self._reload_pages(pages)
        self.status.showMessage(f"Đã thay {len(pages)} trang  (Ctrl+Z để hoàn tác)", 5000)

    def find_duplicates(self):
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            groups = self.pdf.find_duplicate_pages()
        finally:
            QApplication.restoreOverrideCursor()
        if not groups:
            QMessageBox.information(self, APP_NAME, "Không có trang nào trùng nhau.")
            return
        extra = sorted(p for g in groups for p in g[1:])
        desc = "\n".join("Trang " + ", ".join(str(p + 1) for p in g) for g in groups[:15])
        if QMessageBox.question(self, APP_NAME, f"Các nhóm trang giống hệt nhau:\n{desc}\n\n"
                                f"Chọn sẵn {len(extra)} trang bản sao (giữ trang đầu mỗi nhóm) để bạn xem / xoá?"
                                ) == QMessageBox.Yes:
            self.select_pages(extra)
            self.status.showMessage("Đã chọn các trang trùng - bấm Delete để xoá", 8000)

    def on_note_clicked(self, xref):
        info = self.pdf.annot_info(self.current, xref)
        if info is None:
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("Ghi chú")
        dlg.resize(460, 320)
        v = QVBoxLayout(dlg)
        meta = " · ".join(x for x in (info["title"], info["date"], info["kind"]) if x)
        if meta:
            lb = QLabel(meta)
            lb.setStyleSheet("color:#64748b")
            v.addWidget(lb)
        ed = QPlainTextEdit(info["content"])
        v.addWidget(ed, 1)
        row = QHBoxLayout()
        b_del = QPushButton("Xoá ghi chú")
        row.addWidget(b_del)
        row.addStretch()
        b_save = QPushButton("Lưu")
        b_save.setDefault(True)
        b_close = QPushButton("Đóng")
        row.addWidget(b_save)
        row.addWidget(b_close)
        v.addLayout(row)
        act = {}
        b_del.clicked.connect(lambda: (act.update(a="del"), dlg.accept()))
        b_save.clicked.connect(lambda: (act.update(a="save"), dlg.accept()))
        b_close.clicked.connect(dlg.reject)
        if dlg.exec() != QDialog.Accepted:
            return
        if act.get("a") == "del":
            self.run_edit(lambda: self.pdf.delete_annot(self.current, xref))
        elif ed.toPlainText() != info["content"]:
            self.run_edit(lambda: self.pdf.set_annot_content(self.current, xref, ed.toPlainText()))

    def insert_link(self, rect, tiny):
        dlg = LinkDialog(self, self.pdf.page_count, need_text=tiny)
        if dlg.exec() != QDialog.Accepted:
            return
        kind, target = dlg.target()
        body, fams = dlg.html()                        # vùng liên kết = đúng vùng chữ hiển thị (pdf_core)
        self.run_edit(lambda: self.pdf.add_link(self.current, rect, uri=target if kind == "uri" else None,
                                                target_page=target if kind == "page" else None,
                                                html=body, families=fams))

    # ------------------------------------------------------------------ ký số
    def _signing_bytes(self):
        """Nội dung để ký: file gốc nếu chưa sửa (giữ nguyên chữ ký cũ), ngược lại bản đang sửa."""
        if not self.pdf.modified and self.pdf.path and self.pdf.path.lower().endswith(".pdf"):
            with open(self.pdf.path, "rb") as f:
                return f.read()
        if self.pdf.new_fonts:
            try:
                self.pdf.doc.subset_fonts()
            except Exception:
                pass
            self.pdf.new_fonts = False
        return self.pdf.doc.tobytes(garbage=3, deflate=True)

    def sign_invisible(self):
        self.sign(None)

    def sign(self, rect):
        import pdf_sign
        page = self.pdf.doc[self.current]
        box = self.pdf.pdf_box(self.current, rect) if rect else None
        dlg = SignDialog(self, (box[2] - box[0], box[3] - box[1]) if box else None)
        if dlg.exec() != QDialog.Accepted:
            return
        if self.pdf.modified and self.pdf.has_signatures():
            if QMessageBox.question(self, APP_NAME, "File đã có chữ ký số và đã bị chỉnh sửa trong phiên này - "
                                    "các chữ ký cũ sẽ mất hiệu lực. Vẫn ký?") != QMessageBox.Yes:
                return
        base = os.path.splitext(self.pdf.path or "tai_lieu")[0]
        out_path, _ = QFileDialog.getSaveFileName(self, "Lưu file đã ký", base + "_da_ky.pdf", "PDF (*.pdf)")
        if not out_path:
            return
        wc = dlg.cert()
        kw = dict(page_index=self.current, box=box, reason=dlg.reason.currentText().strip(),
                  location=dlg.location.text().strip(),
                  image_path=dlg.image_path(), show_text=dlg.show_text.isChecked() if box else True,
                  check_only=dlg.check_only.isChecked() if box else False,
                  tsa_url=dlg.tsa_url.currentText().strip() if dlg.use_tsa.isChecked() else None)
        if box and page.rotation:
            if QMessageBox.question(self, APP_NAME, "Trang này đang được xoay nên hình chữ ký cũng sẽ hiển thị xoay theo "
                                    "trang. Vẫn ký?") != QMessageBox.Yes:
                return
        pdf_sign.PARENT_HWND = int(self.winId())
        data = self._signing_bytes()
        while True:
            try:
                QApplication.setOverrideCursor(Qt.WaitCursor)
                self.status.showMessage("Đang ký... (nhập mã PIN của USB Token nếu được hỏi)")
                QApplication.processEvents()
                signed = pdf_sign.sign_pdf(data, wc, **kw)
                break
            except pdf_sign.SignError as e:
                QApplication.restoreOverrideCursor()
                self.show_error(f"Không ký được:\n{e}")
                return
            except Exception as e:
                QApplication.restoreOverrideCursor()
                if kw["tsa_url"] and ("timestamp" in str(e).lower() or "http" in str(e).lower()
                                      or "connect" in str(e).lower() or "tsa" in str(e).lower()):
                    if QMessageBox.question(self, APP_NAME, f"Không lấy được dấu thời gian từ máy chủ TSA:\n{e}\n\n"
                                            "Ký không có dấu thời gian?") == QMessageBox.Yes:
                        kw["tsa_url"] = None
                        continue
                    return
                self.show_error(f"Lỗi khi ký:\n{e}")
                return
            finally:
                if QApplication.overrideCursor():
                    QApplication.restoreOverrideCursor()
        try:
            with open(out_path, "wb") as f:
                f.write(signed)
        except PermissionError:
            self.show_error("Không ghi được file — file đang được mở bởi chương trình khác.")
            return
        self.pdf.modified = False                     # nội dung đã nằm trong file đã ký
        self.mode = "view"
        self.tool_actions["view"].setChecked(True)
        self.open_path(out_path)
        res = pdf_sign.verify_pdf(signed)
        mine = res[-1] if res else None
        if mine and mine["intact"]:
            QMessageBox.information(self, "Ký số", f"Đã ký số thành công.\n\nNgười ký: {mine['signer']}\n"
                                    f"File: {out_path}\n\nLưu ý: không sửa / lưu đè file này nữa, nếu không chữ ký sẽ mất hiệu lực.")
        else:
            VerifyDialog(self, res, "Kết quả ký").exec()

    def verify_signatures(self):
        import pdf_sign
        try:
            QApplication.setOverrideCursor(Qt.WaitCursor)
            res = pdf_sign.verify_pdf(self._signing_bytes())
        except Exception as e:
            QApplication.restoreOverrideCursor()
            self.show_error(f"Không kiểm tra được:\n{e}")
            return
        QApplication.restoreOverrideCursor()
        note = "  (file đang có thay đổi chưa lưu - kết quả tính trên bản đang sửa)" if self.pdf.modified else ""
        VerifyDialog(self, res, "Chữ ký số trong file" + note).exec()

    def find_replace(self):
        if not self.pdf.doc or not self.in_editor():
            return
        if not self.find_dialog:
            self.find_dialog = FindReplaceDialog(self)
        self.find_dialog.show()
        self.find_dialog.raise_()
        self.find_dialog.find.setFocus()

    # ------------------------------------------------------------------ zoom
    def zoom_step(self, d):
        if not self.pdf.doc:
            return
        steps = [0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1.0, 1.25, 1.5, 1.75, 2.0, 2.5, 3.0, 4.0]
        cur = min(range(len(steps)), key=lambda i: abs(steps[i] - self.zoom))
        self.zoom = steps[max(0, min(len(steps) - 1, cur + d))]
        self.fit_width = False
        self.render_page()

    def zoom_fit(self):
        self.fit_width = True
        self.render_page()


STYLE = """
QMainWindow { background:#f8fafc; }
QToolBar { background:#ffffff; border-bottom:1px solid #e2e8f0; spacing:2px; padding:4px; }
QToolButton { padding:4px 6px; border-radius:6px; color:#1e293b; }
QToolButton:hover { background:#eff6ff; }
QToolButton:checked { background:#dbeafe; color:#1d4ed8; }
QToolButton:disabled { color:#94a3b8; }
QListWidget { background:#f1f5f9; border:none; border-right:1px solid #e2e8f0; outline:0; }
QListWidget::item { border-radius:6px; padding:4px; }
QListWidget::item:selected { background:#dbeafe; border:1px solid #93c5fd; color:#0f172a; }
QListWidget::item:hover:!selected { background:#e2e8f0; }
QStatusBar { background:#ffffff; border-top:1px solid #e2e8f0; color:#334155; }
QDialog { background:#f8fafc; }
QPushButton { padding:6px 14px; background:#ffffff; border:1px solid #e2e8f0; border-radius:8px; color:#1e293b; }
QPushButton:hover { background:#f1f5f9; border-color:#cbd5e1; }
QPushButton:pressed { background:#e2e8f0; }
QPushButton:disabled { color:#94a3b8; background:#f8fafc; border-color:#e2e8f0; }
QLineEdit, QPlainTextEdit, QTextEdit, QSpinBox, QDoubleSpinBox, QComboBox {
    background:#ffffff; border:1px solid #e2e8f0; border-radius:8px; padding:4px 8px; color:#0f172a;
    selection-background-color:#bfdbfe; selection-color:#0f172a; }
QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {
    border-color:#60a5fa; }
QLineEdit:disabled, QSpinBox:disabled, QDoubleSpinBox:disabled, QComboBox:disabled { background:#f1f5f9; color:#94a3b8; }
QComboBox { padding-right:26px; }
QComboBox:hover { border-color:#cbd5e1; }
QComboBox::drop-down { border:none; width:26px; subcontrol-origin:padding; subcontrol-position:center right; }
QComboBox QAbstractItemView { background:#ffffff; border:1px solid #e2e8f0; padding:4px; outline:0;
    selection-background-color:#eff6ff; selection-color:#1d4ed8; }
QSpinBox, QDoubleSpinBox { padding-right:22px; }
QSpinBox::up-button, QDoubleSpinBox::up-button { subcontrol-origin:border; subcontrol-position:top right;
    width:20px; border:none; border-top-right-radius:8px; }
QSpinBox::down-button, QDoubleSpinBox::down-button { subcontrol-origin:border; subcontrol-position:bottom right;
    width:20px; border:none; border-bottom-right-radius:8px; }
QSpinBox::up-button:hover, QSpinBox::down-button:hover, QDoubleSpinBox::up-button:hover,
QDoubleSpinBox::down-button:hover { background:#f1f5f9; }
QScrollBar:vertical { background:transparent; width:12px; margin:2px; }
QScrollBar::handle:vertical { background:#cbd5e1; border-radius:4px; min-height:36px; margin:0 2px; }
QScrollBar:horizontal { background:transparent; height:12px; margin:2px; }
QScrollBar::handle:horizontal { background:#cbd5e1; border-radius:4px; min-width:36px; margin:2px 0; }
QScrollBar::handle:hover { background:#94a3b8; }
QScrollBar::add-line, QScrollBar::sub-line { width:0; height:0; border:none; background:none; }
QScrollBar::add-page, QScrollBar::sub-page { background:none; }
QMenu { background:#ffffff; border:1px solid #e2e8f0; padding:6px; }
QMenu::item { padding:6px 24px 6px 10px; border-radius:6px; color:#1e293b; }
QMenu::item:selected { background:#eff6ff; color:#1d4ed8; }
QMenu::item:disabled { color:#94a3b8; }
QMenu::separator { height:1px; background:#e2e8f0; margin:4px 8px; }
QMenu::icon { padding-left:8px; }
QToolTip { background:#ffffff; color:#1f2937; border:1px solid #d4d4d8; padding:5px 8px; }
QLabel#toast { background:#ffffff; color:#18181b; border:1px solid #d4d4d8; border-radius:12px; padding:10px 18px; }
QHeaderView::section { background:#f8fafc; color:#475569; font-weight:600; border:none;
    border-bottom:1px solid #e2e8f0; border-right:1px solid #f1f5f9; padding:6px 8px; }
QTableWidget, QTableView { background:#ffffff; border:1px solid #e2e8f0; border-radius:8px; gridline-color:#f1f5f9;
    selection-background-color:#eff6ff; selection-color:#1d4ed8; }
QDialog QTabWidget::pane { border:1px solid #e2e8f0; border-radius:12px; background:#ffffff; top:-1px; }
QDialog QTabBar::tab { background:transparent; color:#475569; padding:7px 16px; margin:0 4px 8px 0;
    border:none; border-radius:15px; }
QDialog QTabBar::tab:hover:!selected { background:#e2e8f0; }
QDialog QTabBar::tab:selected { background:#2563eb; color:#ffffff; font-weight:600; }
QWidget#docTabRow { background:#e2e8f0; border-bottom:1px solid #cbd5e1; }
QTabBar#docTabs { background:transparent; }
QTabBar#docTabs::tab { background:transparent; color:#475569; border:1px solid transparent; border-bottom:none;
    border-top-left-radius:8px; border-top-right-radius:8px; padding:6px 8px 6px 12px; margin-right:2px;
    min-width:90px; max-width:240px; }
QTabBar#docTabs::tab:hover:!selected { background:#f1f5f9; }
QTabBar#docTabs::tab:selected { background:#ffffff; color:#0f172a; border-color:#cbd5e1; margin-bottom:-1px; }
QToolButton#tabCloseBtn { border-radius:4px; padding:3px; margin-left:4px; }
QToolButton#tabCloseBtn:hover { background:#e2e8f0; }
QToolButton#newTabBtn { border-radius:14px; padding:4px; }
QToolButton#newTabBtn:hover { background:#cbd5e1; }
"""


def hide_caption_text(win):
    """Không vẽ chữ tiêu đề trên thanh tiêu đề Windows (giữ icon; taskbar / Alt+Tab vẫn hiện tên app)."""
    try:
        import ctypes
        from ctypes import wintypes

        class WTA_OPTIONS(ctypes.Structure):
            _fields_ = [("dwFlags", wintypes.DWORD), ("dwMask", wintypes.DWORD)]
        WTNCA_NODRAWCAPTION = 0x1
        opt = WTA_OPTIONS(WTNCA_NODRAWCAPTION, WTNCA_NODRAWCAPTION)
        ctypes.windll.uxtheme.SetWindowThemeAttribute(wintypes.HWND(int(win.winId())), 1,       # WTA_NONCLIENT
                                                      ctypes.byref(opt), ctypes.sizeof(opt))
    except Exception:
        pass


def prepare_style_assets(app):
    """Mũi tên của ô chọn / ô số (stylesheet có viền thì Qt không tự vẽ mũi tên): vẽ từ font icon Windows ra
    file PNG trong thư mục tạm rồi gắn vào stylesheet."""
    import tempfile
    # tooltip / thẻ nổi: Windows có lúc vẽ nền theo palette chứ không theo stylesheet -> đặt palette cùng màu
    # (nền trắng, chữ tối) để chữ không bao giờ lệch màu với nền
    pal = app.palette()
    for grp in (QPalette.Active, QPalette.Inactive, QPalette.Disabled):
        pal.setColor(grp, QPalette.ToolTipBase, QColor("#ffffff"))
        pal.setColor(grp, QPalette.ToolTipText, QColor("#1f2937"))
    app.setPalette(pal)
    QToolTip.setPalette(pal)
    d = os.path.join(tempfile.gettempdir(), "PDFToolkit_ui")
    paths = {}
    try:
        os.makedirs(d, exist_ok=True)
        for name, code, color in (("down", 0xE70D, "#64748b"), ("up", 0xE70E, "#64748b"),
                                  ("down_off", 0xE70D, "#cbd5e1"), ("up_off", 0xE70E, "#cbd5e1")):
            p = os.path.join(d, f"chevron_{name}.png")
            tools_ui.glyph_icon(code, color, size=32).pixmap(32, 32).save(p)
            paths[name] = p.replace("\\", "/")
    except OSError:
        return
    app.setStyleSheet(app.styleSheet() + f"""
QComboBox::down-arrow {{ image:url("{paths['down']}"); width:11px; height:11px; }}
QComboBox::down-arrow:disabled {{ image:url("{paths['down_off']}"); }}
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow {{ image:url("{paths['up']}"); width:9px; height:9px; }}
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow {{ image:url("{paths['down']}"); width:9px; height:9px; }}
QSpinBox::up-arrow:disabled, QDoubleSpinBox::up-arrow:disabled {{ image:url("{paths['up_off']}"); }}
QSpinBox::down-arrow:disabled, QDoubleSpinBox::down-arrow:disabled {{ image:url("{paths['down_off']}"); }}
""")


def run_cli(argv):
    """PDF Toolkit.exe --tool <id> --out <thư mục> [--opt key=value ...] [--opt-env key=BIẾN ...] file1 file2 ...
    Chạy 1 công cụ không cần giao diện (cho IT tự động hoá). Nhật ký ghi vào <thư mục>\\pdf_toolkit_cli.log."""
    import argparse
    ap = argparse.ArgumentParser(prog="PDF Toolkit")
    ap.add_argument("--tool", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--opt", action="append", default=[])
    ap.add_argument("--opt-env", action="append", default=[])    # key=TÊN_BIẾN: giá trị (mật khẩu) lấy từ biến môi trường
    ap.add_argument("files", nargs="+")
    a = ap.parse_args(argv)
    os.makedirs(a.out, exist_ok=True)
    lines = []
    code = 0
    try:
        t = pdf_tools.get_tool(a.tool)
        opts = {o.key: o.default for o in t.opts}
        for kv in a.opt:
            k, _, v = kv.partition("=")
            try:
                opts[k] = json.loads(v)
            except ValueError:
                opts[k] = v
        for kv in a.opt_env:
            k, _, v = kv.partition("=")
            opts[k] = os.environ.get(v, "")
        ctx = pdf_tools.Ctx(a.files, opts, a.out, log=lines.append)
        t.run(ctx)
        lines += ["OK " + p for p in ctx.outputs]
    except StopIteration:
        lines.append(f"ERROR không có công cụ '{a.tool}'. Có: " + ", ".join(t.id for t in pdf_tools.TOOLS))
        code = 2
    except Exception as e:
        lines.append(f"ERROR {type(e).__name__}: {e}")
        code = 1
    with open(os.path.join(a.out, "pdf_toolkit_cli.log"), "a", encoding="utf-8") as fh:
        fh.write("\n".join([f"# {a.tool} {' '.join(a.files)}"] + lines) + "\n")
    return code


_windows = []          # giữ tham chiếu cửa sổ (đổi ngôn ngữ tạo cửa sổ mới)


def setup_i18n():
    i18n.set_lang(QSettings(installer.ORG, "PDFToolkit").value("lang", "vi"))
    i18n.install([sys.modules[__name__], tools_ui, about])
    filedlg.install()               # hộp thoại Windows: tách bộ lọc gộp thành Word / Excel / PowerPoint...


def collect_ctx_files(key, files, wait_ms=1500):
    """Menu chuột phải với nhiều file: Windows chạy 1 tiến trình cho mỗi file. Tiến trình đầu tiên mở
    'cổng' (named pipe) chờ các tiến trình sau gửi đường dẫn sang rồi xử lý tất cả trong 1 cửa sổ.
    Trả về danh sách file cần xử lý, hoặc None nếu đã chuyển file cho tiến trình khác (thoát luôn)."""
    import ctypes
    import time
    from PySide6.QtCore import QElapsedTimer
    from PySide6.QtNetwork import QLocalServer, QLocalSocket
    name = f"PDFToolkit_ctx_{os.environ.get('USERNAME', '')}_{key}"
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateMutexW.restype = ctypes.c_void_p
    # named pipe trên Windows cho nhiều server trùng tên -> dùng mutex để chọn đúng 1 tiến trình gom file
    mutex = k32.CreateMutexW(None, True, "Local\\" + name)
    owner = bool(mutex) and ctypes.get_last_error() != 183             # 183 = ERROR_ALREADY_EXISTS
    if not owner:
        if mutex:
            k32.CloseHandle(ctypes.c_void_p(mutex))
        for _ in range(60):                                             # chờ tiến trình gom mở cổng (tối đa ~6 s)
            s = QLocalSocket()
            s.connectToServer(name)
            if s.waitForConnected(100):
                s.write(("\n".join(files) + "\n\0").encode("utf-8"))      # \0 = hết danh sách
                s.waitForBytesWritten(2000)
                ack = b""
                while b"ok" not in ack and s.waitForReadyRead(3000):
                    ack += bytes(s.readAll())
                s.disconnectFromServer()
                if b"ok" in ack:
                    return None
                break                                                   # cổng đang đóng: tự xử lý file của mình
            time.sleep(0.1)
        return list(files)                                              # không gặp ai: tự xử lý
    srv = QLocalServer()
    if not srv.listen(name):
        k32.CloseHandle(ctypes.c_void_p(mutex))
        return list(files)
    got = list(files)

    def receive(c):
        data = b""
        while b"\0" not in data and (c.bytesAvailable() or c.waitForReadyRead(2000)):
            data += bytes(c.readAll())
        if b"\0" in data:                              # nhận đủ -> báo "ok" để tiến trình kia thoát
            got.extend(p for p in data.split(b"\0")[0].decode("utf-8", "replace").splitlines() if p)
            c.write(b"ok")
            c.waitForBytesWritten(1000)
        c.disconnectFromServer()

    t = QElapsedTimer()
    t.start()
    deadline = wait_ms
    while t.elapsed() < deadline:
        srv.waitForNewConnection(50)                   # PySide6 trả về tuple -> kiểm tra bằng hasPendingConnections
        while srv.hasPendingConnections():
            receive(srv.nextPendingConnection())
            deadline = max(deadline, t.elapsed() + 800)   # còn file đang tới: chờ thêm
    srv.close()
    while srv.hasPendingConnections():                 # kết nối tới đúng lúc đóng cổng
        receive(srv.nextPendingConnection())
    k32.CloseHandle(ctypes.c_void_p(mutex))
    seen, out = set(), []
    for p in got:
        k = os.path.normcase(os.path.abspath(p))
        if k not in seen and os.path.exists(p):
            seen.add(k)
            out.append(p)
    # các tiến trình gửi tới theo thứ tự ngẫu nhiên -> xếp theo tên file (số trong tên theo giá trị: 2 < 10)
    return sorted(out, key=lambda p: [int(x) if x.isdigit() else x.lower()
                                      for x in re.split(r"(\d+)", os.path.basename(p))])


def _main_pipe_name():
    return f"PDFToolkit_main_{os.environ.get('USERNAME', '')}"


def _forwardable(args, ctx_files):
    """Lệnh chỉ mở file (Explorer, Open with, menu chuột phải Open / Sign, add-in Office) -> gửi sang cửa sổ
    đang chạy để mở thành tab mới. Các công cụ (gộp, nén...) vẫn mở cửa sổ riêng để không chen vào việc đang làm."""
    if ctx_files is not None:
        return args[1] in ("open", "sign") and bool(ctx_files)
    if len(args) >= 2 and args[0] == "--sign":
        return os.path.exists(args[1])
    return bool(args) and not args[0].startswith("--") and os.path.exists(args[0])


def forward_to_running(args, ctx_files, wait_ms=300):
    """Gửi lệnh mở file sang phiên PDF Toolkit đang chạy. True = phiên kia đã nhận (tiến trình này thoát).
    wait_ms: phiên kia có thể đang khởi động (đã giữ cổng nhưng chưa mở) -> thử kết nối lại đến hết thời gian này."""
    import ctypes
    import json
    import time
    from PySide6.QtNetwork import QLocalSocket
    end = time.monotonic() + wait_ms / 1000
    while True:
        s = QLocalSocket()
        s.connectToServer(_main_pipe_name())
        if s.waitForConnected(300):
            break
        if time.monotonic() >= end:
            return False
        time.sleep(0.2)
    try:                            # cho phép cửa sổ kia lên trước (Windows chặn nếu không có quyền foreground)
        ctypes.windll.user32.AllowSetForegroundWindow(-1)
    except Exception:
        pass
    msg = {"args": [os.path.abspath(a) if os.path.exists(a) else a for a in args], "ctx": ctx_files}
    s.write(json.dumps(msg).encode("utf-8") + b"\0")
    s.waitForBytesWritten(2000)
    ack = b""
    while b"ok" not in ack and s.waitForReadyRead(20000):     # phiên kia có thể còn đang dựng cửa sổ
        ack += bytes(s.readAll())
    s.disconnectFromServer()
    return b"ok" in ack


class InstanceServer(QObject):
    """Cổng của phiên đang chạy: nhận lệnh mở file từ các lần mở sau và mở thành tab (như trình duyệt).
    Mutex chọn đúng 1 phiên giữ cổng; phiên đó đóng thì lần mở kế tiếp tự nhận cổng."""

    def __init__(self, parent=None):
        super().__init__(parent)
        import ctypes
        from PySide6.QtNetwork import QLocalServer
        self._buf = {}
        self.srv = None
        self.pending = []                               # lệnh tới khi cửa sổ chưa dựng xong
        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.CreateMutexW.restype = ctypes.c_void_p
        self._mutex = k32.CreateMutexW(None, True, "Local\\" + _main_pipe_name())
        self.other = bool(self._mutex) and ctypes.get_last_error() == 183   # đã có phiên khác giữ cổng
        if not self._mutex or self.other:
            return
        srv = QLocalServer(self)
        if srv.listen(_main_pipe_name()):
            srv.newConnection.connect(self._accept)
            self.srv = srv

    def _accept(self):
        while self.srv.hasPendingConnections():
            c = self.srv.nextPendingConnection()
            self._buf[c] = b""
            c.readyRead.connect(lambda c=c: self._read(c))
            c.disconnected.connect(lambda c=c: (self._buf.pop(c, None), c.deleteLater()))
            if c.bytesAvailable():
                self._read(c)

    def _read(self, c):
        import json
        if c not in self._buf:
            return
        self._buf[c] += bytes(c.readAll())
        if b"\0" not in self._buf[c]:
            return
        raw = self._buf.pop(c).split(b"\0")[0]
        c.write(b"ok")                                  # báo nhận trước, mở file sau (mở file có thể lâu)
        c.flush()
        try:
            msg = json.loads(raw.decode("utf-8"))
        except Exception:
            return
        QTimer.singleShot(0, lambda: self._handle(msg.get("args") or [], msg.get("ctx")))

    def flush(self):
        while self.pending:
            self._handle(*self.pending.pop(0))

    def _handle(self, args, ctx_files):
        if not _windows:
            self.pending.append((args, ctx_files))
            return
        w = _windows[-1]
        if w.isMinimized():
            w.showNormal()
        w.raise_()
        w.activateWindow()
        dispatch_args(w, args, ctx_files)


def dispatch_args(w, args, ctx_files):
    """Thực hiện lệnh dòng lệnh trên cửa sổ w (lúc khởi động hoặc nhận từ phiên khác)."""
    if ctx_files is not None:
        key = args[1]
        if key in ("open", "sign"):
            for p in ctx_files:
                w.open_path(p)
            if key == "sign" and w.pdf.doc is not None:
                w.set_mode("sign")
        elif ctx_files:
            w.open_tool(key)
            if w.stack.currentWidget() is w.tool_page:
                w.tool_page.files.add_paths(ctx_files)
    elif len(args) >= 2 and args[0] == "--sign" and os.path.exists(args[1]):          # từ add-in Office
        w.open_path(args[1])
        if w.pdf.doc is not None:
            w.set_mode("sign")
    elif len(args) >= 2 and args[0] == "--open-tool":
        w.open_tool(args[1])
        files = [p for p in args[2:] if os.path.exists(p)]
        if files and w.stack.currentWidget() is w.tool_page:
            w.tool_page.files.add_paths(files)
    elif args and os.path.exists(args[0]):
        if w.open_path(args[0]) and len(args) > 1:
            w.insert_paths([p for p in args[1:] if os.path.exists(p)], w.pdf.page_count)


def run_setup_cli(argv):
    """--install [--dir X] [--no-desktop] [--no-startmenu] [--no-openwith] [--no-contextmenu] [--no-explorer] [--no-office] [--quiet]
    --uninstall [--quiet]   (cài cho người dùng hiện tại, không cần quyền admin - dùng được trong script IT)"""
    quiet = "--quiet" in argv
    logs = []
    try:
        if argv[0] == "--uninstall":
            installer.uninstall(log=logs.append, only_self=True)
        else:
            d = argv[argv.index("--dir") + 1] if "--dir" in argv else None
            installer.install(d, "--no-startmenu" not in argv, "--no-desktop" not in argv,
                              "--no-openwith" not in argv,
                              "--no-office" not in argv and installer.office_present(), APP_VERSION, log=logs.append,
                              context_menu="--no-contextmenu" not in argv, explorer="--no-explorer" not in argv)
        code = 0
    except Exception as e:
        logs.append("LỖI: " + str(e))
        code = 1
    if not quiet:
        app = QApplication.instance() or QApplication(sys.argv)
        app.setStyle("Fusion")
        setup_i18n()
        (QMessageBox.information if code == 0 else QMessageBox.warning)(None, APP_NAME, "\n".join(logs))
    return code


def main():
    if len(sys.argv) > 1 and sys.argv[1] == "--tool":
        sys.exit(run_cli(sys.argv[1:]))
    if len(sys.argv) > 1 and sys.argv[1] in ("--install", "--uninstall"):
        sys.exit(run_setup_cli(sys.argv[1:]))
    if len(sys.argv) > 3 and sys.argv[1] == "--export-comments":       # add-in Word: nhập nhận xét từ PDF
        sys.exit(pdf_core.export_comments(sys.argv[2], sys.argv[3]))
    os.environ.setdefault("QT_ENABLE_HIGHDPI_SCALING", "1")
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 10))
    app.setStyleSheet(STYLE + tools_ui.TOOLS_STYLE)
    prepare_style_assets(app)
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    icon = os.path.join(base, "app.ico")
    if os.path.exists(icon):
        app.setWindowIcon(QIcon(icon))
    try:                    # taskbar hiện icon của app thay vì icon python
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("TuanPA.PDFToolkit")
    except Exception:
        pass
    args = sys.argv[1:]
    ctx_files = None
    if len(args) >= 2 and args[0] == "--ctx":               # menu chuột phải trên file PDF
        ctx_files = collect_ctx_files(args[1], args[2:])
        if ctx_files is None:                                # đã chuyển file cho cửa sổ đang gom
            sys.exit(0)
    # giữ cổng ngay khi khởi động: các lần mở file ngay sau đó (cửa sổ này chưa dựng xong) vẫn gửi được sang đây
    server = InstanceServer(app)
    if server.other and _forwardable(args, ctx_files) and forward_to_running(args, ctx_files, 20000):
        sys.exit(0)                                          # đã mở thành tab trong cửa sổ đang chạy
    try:                    # cài đặt từ tên cũ NagaseVN (bản 1.3.0 trở về trước) -> chép sang tên mới
        installer.migrate_legacy()
    except Exception:
        pass
    setup_i18n()
    try:                    # bản đã cài mà mất / lệch đăng ký (vừa cập nhật exe, registry bị xoá) -> đăng ký lại
        installer.repair(APP_VERSION)
    except Exception:
        pass
    w = MainWindow()
    _windows.append(w)
    w.showMaximized()                                        # mở app luôn toàn màn hình (phóng to cửa sổ)
    dispatch_args(w, args, ctx_files)
    server.flush()                                           # file gửi tới trong lúc đang khởi động
    sys.exit(app.exec())


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()        # tiến trình con OCR trong bản exe
    main()

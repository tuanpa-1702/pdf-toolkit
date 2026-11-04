# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Giao diện Sáng / Tối / Theo Windows.

Chế độ tối đổi màu lúc chạy cho mọi stylesheet (của app và của từng widget): nền sáng -> xám than (không quá tối),
chữ / icon tối -> sáng, màu nhấn (xanh, tím, cam...) giữ nguyên sắc nhưng sáng hơn để nổi trên nền tối. Nhờ vậy
không phải viết lại hàng trăm stylesheet. Trang PDF vẫn trắng như Acrobat / Edge (đó là nội dung tài liệu).
"""
import colorsys
import re

MODES = ("light", "dark", "system")
MODE = "light"
DARK = False

# màu chính của chế độ tối
BG, BASE, RAISED, TEXT, MUTED = "#2b2f3a", "#30343f", "#363b47", "#eef0f4", "#a3a9b8"
ACCENT = "#7c3aed"


def system_dark():
    """Windows đang dùng giao diện tối cho ứng dụng? (Settings > Personalization > Colors)"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as k:
            return winreg.QueryValueEx(k, "AppsUseLightTheme")[0] == 0
    except OSError:
        return False


def set_mode(mode):
    global MODE, DARK
    MODE = mode if mode in MODES else "light"
    DARK = MODE == "dark" or (MODE == "system" and system_dark())
    return DARK


# ---------------------------------------------------------------------------
#  đổi màu
# ---------------------------------------------------------------------------
def _dark_rgb(r, g, b):
    h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    if s > 0.45 and 0.25 < l < 0.75:                # màu nhấn: giữ sắc, sáng hơn cho nền tối
        l = min(0.72, max(0.56, l + 0.09))
    else:                                           # nền sáng -> xám than, chữ tối -> chữ sáng
        l = 0.19 + (1 - l) * 0.84
        if l > 0.5:
            l = min(0.96, l + 0.05)
        else:
            s = min(s, 0.22)
    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return int(r2 * 255), int(g2 * 255), int(b2 * 255)


_HEX = re.compile(r"#([0-9a-fA-F]{6}|[0-9a-fA-F]{3})\b")
_RGB = re.compile(r"(rgba?)\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*((?:,\s*[\d.]+\s*)?)\)")


def css(text):
    """Stylesheet / HTML theo giao diện hiện tại (chế độ sáng: giữ nguyên)."""
    if not DARK or not text:
        return text

    def hexrep(m):
        v = m.group(1)
        if len(v) == 3:
            v = "".join(ch * 2 for ch in v)
        return "#%02x%02x%02x" % _dark_rgb(int(v[0:2], 16), int(v[2:4], 16), int(v[4:6], 16))

    def rgbrep(m):
        r, g, b = _dark_rgb(*(int(x) for x in m.group(2, 3, 4)))
        return f"{m.group(1)}({r}, {g}, {b}{m.group(5)})"
    text = _HEX.sub(hexrep, text)
    text = _RGB.sub(rgbrep, text)
    text = re.sub(r"(?<![-\w])white\b", BASE, text)
    text = re.sub(r"(?<![-\w])black\b", TEXT, text)
    return text


def icon_color(color):
    """Màu icon theo giao diện: chế độ tối -> icon mực tối thành sáng, icon màu sáng lên; icon trắng (trên nền
    màu nhấn) giữ nguyên."""
    if not DARK:
        return color
    from PySide6.QtGui import QColor
    q = QColor(color)
    if not q.isValid():
        return color
    h, l, s = colorsys.rgb_to_hls(q.redF(), q.greenF(), q.blueF())
    if l > 0.95:
        return q
    if s > 0.45 and 0.25 < l < 0.8:
        l = min(0.78, max(0.64, l + 0.12))
        r, g, b = (int(v * 255) for v in colorsys.hls_to_rgb(h, l, s))
    else:
        r, g, b = _dark_rgb(q.red(), q.green(), q.blue())
    return QColor(r, g, b, q.alpha())


# chỉnh tay vài chỗ mà đổi màu tự động chưa đủ tương phản
DARK_FIX = f"""
#ribbon QTabBar::tab {{ color:#e4e4e7; }}
#ribbon QTabBar::tab:selected {{ background:{ACCENT}; color:#ffffff; font-weight:600; }}
#ribbon QTabBar::tab:hover:!selected {{ background:rgba(255,255,255,0.10); color:#ffffff; }}
QToolButton#ribbonBig {{ color:#f4f4f5; }}
QToolButton#ribbonBig:hover {{ background:rgba(255,255,255,0.10); }}
QToolButton#ribbonBig:checked {{ background:{ACCENT}; color:#ffffff; }}
QToolButton#ribbonBig:disabled {{ color:#8b8f9c; }}
QToolButton#ribbonCollapse:hover {{ background:rgba(255,255,255,0.12); }}
#quickBar QToolButton:hover {{ background:rgba(255,255,255,0.12); }}
#quickBar QToolButton:checked {{ background:rgba(167,139,250,0.30); }}
#ribbonScroll QScrollBar:horizontal {{ background:#3a3f4c; }}
#ribbonScroll QScrollBar::handle:horizontal {{ background:#5b6273; }}
QToolTip {{ background:{RAISED}; color:{TEXT}; border:1px solid #4b5263; }}
"""


def app_css(base):
    return css(base) + (DARK_FIX if DARK else "")


_light_palette = None


def apply_palette(app):
    """Palette của Qt (ô nhập, tooltip, menu hệ thống...) theo giao diện. Gọi lại được khi đổi giao diện."""
    global _light_palette
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QPalette, QColor
    try:                    # Qt 6.8+: ép tông màu của app (thanh cuộn, menu hệ thống...) theo lựa chọn, không theo Windows
        app.styleHints().setColorScheme(Qt.ColorScheme.Dark if DARK else Qt.ColorScheme.Light)
    except AttributeError:
        pass
    if _light_palette is None:
        # palette sáng dựng tay - KHÔNG lấy app.palette() / standardPalette(): Windows đang để giao diện tối thì Qt
        # trả palette tối (chữ trắng) -> chế độ Sáng của app bị chữ trắng trên nền sáng
        _light_palette = QPalette()
        for role, col in ((QPalette.Window, "#f8fafc"), (QPalette.Base, "#ffffff"), (QPalette.AlternateBase, "#f1f5f9"),
                          (QPalette.Text, "#0f172a"), (QPalette.WindowText, "#0f172a"), (QPalette.Button, "#f1f5f9"),
                          (QPalette.ButtonText, "#0f172a"), (QPalette.BrightText, "#ffffff"),
                          (QPalette.ToolTipBase, "#ffffff"), (QPalette.ToolTipText, "#1f2937"),
                          (QPalette.Highlight, "#2563eb"), (QPalette.HighlightedText, "#ffffff"),
                          (QPalette.Link, "#2563eb"), (QPalette.PlaceholderText, "#94a3b8"), (QPalette.Light, "#ffffff"),
                          (QPalette.Midlight, "#e2e8f0"), (QPalette.Mid, "#cbd5e1"), (QPalette.Dark, "#94a3b8"),
                          (QPalette.Shadow, "#64748b")):
            _light_palette.setColor(role, QColor(col))
        for role in (QPalette.Text, QPalette.WindowText, QPalette.ButtonText):
            _light_palette.setColor(QPalette.Disabled, role, QColor("#94a3b8"))
    if not DARK:
        app.setPalette(_light_palette)
        return
    pal = QPalette(_light_palette)
    for role, col in ((QPalette.Window, BG), (QPalette.Base, BASE), (QPalette.AlternateBase, RAISED),
                      (QPalette.Text, TEXT), (QPalette.WindowText, TEXT), (QPalette.Button, RAISED),
                      (QPalette.ButtonText, TEXT), (QPalette.ToolTipBase, RAISED), (QPalette.ToolTipText, TEXT),
                      (QPalette.Highlight, "#4f8df7"), (QPalette.HighlightedText, "#ffffff"),
                      (QPalette.PlaceholderText, MUTED), (QPalette.Light, "#4b5263"), (QPalette.Midlight, "#414653"),
                      (QPalette.Mid, "#3a3f4c"), (QPalette.Dark, "#202430"), (QPalette.Shadow, "#14161c")):
        pal.setColor(role, QColor(col))
    for role in (QPalette.Text, QPalette.WindowText, QPalette.ButtonText):
        pal.setColor(QPalette.Disabled, role, QColor("#7b8191"))
    app.setPalette(pal)


_installed = False


def install_hooks():
    """Stylesheet / chữ có màu đặt riêng từng widget trong mã (setStyleSheet, QLabel HTML) cũng đổi theo giao
    diện. Gọi 1 lần trước khi dựng cửa sổ."""
    global _installed
    if _installed:
        return
    _installed = True
    from PySide6.QtWidgets import QWidget, QLabel
    set_ss, set_text = QWidget.setStyleSheet, QLabel.setText

    def set_style(self, sheet):
        set_ss(self, css(sheet or ""))

    def label_text(self, text):
        set_text(self, css(text) if DARK and text and "color" in text else text)
    QWidget.setStyleSheet = set_style
    QLabel.setText = label_text


def dark_title_bar(win):
    """Thanh tiêu đề Windows tối / sáng theo giao diện (Windows 10 2004+ / 11)."""
    try:
        import ctypes
        from ctypes import wintypes
        val = ctypes.c_int(1 if DARK else 0)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(wintypes.HWND(int(win.winId())), 20,    # DWMWA_USE_IMMERSIVE_DARK_MODE
                                                   ctypes.byref(val), ctypes.sizeof(val))
    except Exception:
        pass

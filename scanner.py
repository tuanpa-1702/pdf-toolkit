# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
"""Quét tài liệu từ máy scan qua WIA (Windows Image Acquisition - có sẵn trong Windows, không cần driver TWAIN).
Mặt kính: quét từng trang, hỏi quét tiếp; khay nạp giấy (ADF): quét liên tục đến khi hết giấy."""
import os
import tempfile

WIA_FORMAT_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
SCANNER = 1                                  # WiaDeviceType.ScannerDeviceType
P_DOC_HANDLING_SELECT = 3088                 # nguồn giấy: 1 = khay nạp (feeder), 2 = mặt kính (flatbed)
P_INTENT = 6146                              # 1 màu, 2 xám, 4 đen trắng (chữ)
P_XRES, P_YRES = 6147, 6148
INTENTS = {"color": 1, "gray": 2, "bw": 4}

# mã lỗi WIA thường gặp (HRESULT, dạng có dấu như pywin32 trả về)
E_PAPER_EMPTY = -2145320957                  # 0x80210003 hết giấy trong khay
E_OFFLINE = -2145320955                      # 0x80210005
E_BUSY = -2145320954                         # 0x80210006
E_PAPER_JAM = -2145320958                    # 0x80210002
E_NO_DEVICE = -2145320939                    # 0x80210015 không có máy scan
E_CANCELLED = -2145320860                    # 0x80210064 người dùng huỷ


class ScanError(Exception):
    pass


def _hresult(e):
    """HRESULT thật của lỗi COM (lỗi WIA nằm trong excepinfo)."""
    try:
        ex = e.excepinfo
        if ex and ex[5]:
            return ex[5]
    except Exception:
        pass
    return getattr(e, "hresult", None)


def _set(props, pid, value):
    """Đặt thuộc tính WIA nếu máy hỗ trợ (máy không có / không nhận giá trị thì bỏ qua)."""
    try:
        for p in props:
            if p.PropertyID == pid:
                p.Value = value
                return True
    except Exception:
        pass
    return False


MESSAGES = {
    E_OFFLINE: "Máy scan đang tắt hoặc chưa kết nối.",
    E_BUSY: "Máy scan đang bận - thử lại sau.",
    E_PAPER_JAM: "Kẹt giấy trong máy scan.",
    E_NO_DEVICE: "Không tìm thấy máy scan nào (WIA). Kiểm tra máy đã cắm / bật và đã cài driver.",
}


def scan_pages(source="flatbed", color="color", dpi=200, more=None, tmpdir=None):
    """Quét và trả về danh sách đường dẫn ảnh PNG (mỗi trang 1 ảnh). [] = người dùng huỷ ngay từ đầu.
    more(n): hỏi có quét tiếp sau trang thứ n không (chỉ dùng với mặt kính)."""
    try:
        import pythoncom
        import pywintypes
        import win32com.client
    except ImportError:
        raise ScanError("Thiếu thành phần pywin32 để dùng máy scan.")
    pythoncom.CoInitialize()
    try:
        dlg = win32com.client.Dispatch("WIA.CommonDialog")
    except pywintypes.com_error:
        raise ScanError("Windows trên máy này không có dịch vụ quét ảnh (WIA).")
    try:
        dev = dlg.ShowSelectDevice(SCANNER, False, False)
    except pywintypes.com_error as e:
        raise ScanError(MESSAGES.get(_hresult(e), f"Không mở được máy scan ({e})."))
    if dev is None:
        return []
    if source == "feeder" and not _set(dev.Properties, P_DOC_HANDLING_SELECT, 1):
        raise ScanError("Máy scan này không có khay nạp giấy tự động (ADF) - chọn Mặt kính.")
    if source == "flatbed":
        _set(dev.Properties, P_DOC_HANDLING_SELECT, 2)
    item = dev.Items(1)
    _set(item.Properties, P_INTENT, INTENTS.get(color, 1))
    _set(item.Properties, P_XRES, int(dpi))
    _set(item.Properties, P_YRES, int(dpi))
    tmpdir = tmpdir or tempfile.mkdtemp(prefix="pdftoolkit_scan_")
    pages = []
    while True:
        try:
            img = dlg.ShowTransfer(item, WIA_FORMAT_PNG, False)
        except pywintypes.com_error as e:
            code = _hresult(e)
            if code == E_PAPER_EMPTY:
                if pages:
                    break                            # khay đã hết giấy: xong
                raise ScanError("Khay nạp giấy đang trống - đặt giấy vào rồi quét lại.")
            if code == E_CANCELLED:
                break
            raise ScanError(MESSAGES.get(code, f"Lỗi khi quét: {e}"))
        if img is None:
            break
        p = os.path.join(tmpdir, f"page_{len(pages) + 1:03d}.png")
        if os.path.exists(p):
            os.remove(p)
        img.SaveFile(p)
        pages.append(p)
        if source == "feeder":
            continue
        if not (more and more(len(pages))):
            break
    return pages

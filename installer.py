# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Cài đặt PDF Toolkit cho người dùng hiện tại - KHÔNG cần quyền administrator.

- Chép exe vào %LOCALAPPDATA%\\Programs\\PDF Toolkit
- Shortcut Start Menu / Desktop
- "Open with" cho file .pdf (+ khai báo trong Default apps để chọn làm mặc định)
- Menu chuột phải "PDF Toolkit" cho file .pdf (nén, tách, chuyển Word / Excel, OCR, mật khẩu, ký số...)
- Mục gỡ cài đặt trong Settings > Apps
- Add-in Office (Word / Excel / PowerPoint / Outlook): COM add-in .NET đăng ký trong HKCU
- Explorer: ảnh thu nhỏ + khung xem trước file PDF (PDFToolkitShell.dll, chỉ thêm khung xem trước khi máy chưa có)
Mọi khoá registry đều nằm dưới HKEY_CURRENT_USER.
"""
import os
import shutil
import subprocess
import sys
import time
import winreg

APP = "PDF Toolkit"
EXE_NAME = "PDF Toolkit.exe"
DLL_NAME = "PDFToolkitAddin.dll"
ORG = "TuanPA"                                          # QSettings(ORG, "PDFToolkit"), %APPDATA%\TuanPA\PDFToolkit
PROGID = "TuanPA.PDFToolkit.pdf"
UNINSTALL_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\TuanPA.PDFToolkit"
INFO_KEY = r"Software\TuanPA\PDFToolkit\Install"
CAPS_KEY = r"Software\TuanPA\PDFToolkit\Capabilities"
# tên cũ (bản 1.3.0 trở về trước) - chuyển dữ liệu sang tên mới rồi dọn
LEGACY_ORG = "NagaseVN"
LEGACY_PROGID = "NagaseVN.PDFToolkit.pdf"
LEGACY_UNINSTALL_KEY = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\NagaseVN.PDFToolkit"
LEGACY_SETTINGS_KEY = r"Software\NagaseVN\PDFToolkit"
ADDIN_PROGID = "PDFToolkit.OfficeAddin"
ADDIN_CLSID = "{6E2B7C1A-4F0D-4E8B-9A57-2D3C8F1B5A61}"
ADDIN_CLASS = "PDFToolkitAddin.Connect"
ADDIN_VERSION = "1.4.0.0"
ADDIN_ASSEMBLY = f"PDFToolkitAddin, Version={ADDIN_VERSION}, Culture=neutral, PublicKeyToken=null"
OFFICE_APPS = ("Word", "Excel", "PowerPoint", "Outlook")
SHELL_DLL = "PDFToolkitShell.dll"
SHELL_VERSION = "1.0.0.0"
SHELL_ASSEMBLY = f"PDFToolkitShell, Version={SHELL_VERSION}, Culture=neutral, PublicKeyToken=null"
THUMB_CLSID = "{3F0C5D2E-8B7A-4E61-9C24-6A1D5E8B7F31}"
PREVIEW_CLSID = "{7A9E3C14-2D5B-4F86-B1E0-93C7A64D2B58}"
THUMB_IID = "{e357fccd-a995-4576-b01f-234630154e96}"          # IThumbnailProvider
PREVIEW_IID = "{8895b1c6-b41f-4c1c-a562-0d564250836f}"        # IPreviewHandler
PREVHOST_APPID = "{6d2b5079-2f0b-48dd-ab7f-97cec514d30b}"     # prevhost.exe 64-bit
HKCU = winreg.HKEY_CURRENT_USER


class InstallError(Exception):
    pass


def default_dir():
    return os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser(r"~\AppData\Local")), "Programs", APP)


def frozen():
    return bool(getattr(sys, "frozen", False))


def resource(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    for p in (os.path.join(base, name), os.path.join(base, "office_addin", name), os.path.join(base, "shell_ext", name)):
        if os.path.exists(p):
            return p
    return None


# ---------------------------------------------------------------------------
#  registry
# ---------------------------------------------------------------------------
def _set(path, values, view=0):
    """values: {tên: giá trị}; '' = giá trị mặc định. int -> DWORD, None -> REG_NONE rỗng."""
    with winreg.CreateKeyEx(HKCU, path, 0, winreg.KEY_WRITE | view) as k:
        for name, v in values.items():
            if v is None:
                winreg.SetValueEx(k, name, 0, winreg.REG_NONE, b"")
            elif isinstance(v, int):
                winreg.SetValueEx(k, name, 0, winreg.REG_DWORD, v)
            else:
                winreg.SetValueEx(k, name, 0, winreg.REG_SZ, str(v))


def _get(path, name="", view=0):
    try:
        with winreg.OpenKey(HKCU, path, 0, winreg.KEY_READ | view) as k:
            return winreg.QueryValueEx(k, name)[0]
    except OSError:
        return None


def _del_tree(path, view=0):
    try:
        with winreg.OpenKey(HKCU, path, 0, winreg.KEY_ALL_ACCESS | view) as k:
            while True:
                try:
                    sub = winreg.EnumKey(k, 0)
                except OSError:
                    break
                _del_tree(path + "\\" + sub, view)
        winreg.DeleteKeyEx(HKCU, path, view or winreg.KEY_WOW64_64KEY)
    except OSError:
        pass


def _del_value(path, name):
    try:
        with winreg.OpenKey(HKCU, path, 0, winreg.KEY_SET_VALUE) as k:
            winreg.DeleteValue(k, name)
    except OSError:
        pass


def _notify_assoc():
    try:
        import ctypes
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)     # SHCNE_ASSOCCHANGED
    except Exception:
        pass


# ---------------------------------------------------------------------------
#  trạng thái
# ---------------------------------------------------------------------------
def installed_exe():
    p = _get(INFO_KEY, "Exe")
    return p if p and os.path.exists(p) else None


def installed_version():
    return _get(INFO_KEY, "Version")


def _same(a, b):
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def is_installed_copy():
    """Đang chạy chính bản đã cài? (theo registry, hoặc đang chạy từ thư mục cài mặc định)"""
    if not frozen():
        return False
    exe = installed_exe()
    if exe:
        return _same(sys.executable, exe)
    return _same(sys.executable, os.path.join(default_dir(), EXE_NAME))


def repair(version, log=lambda s: None):
    """Bản đã cài nhưng mất / lệch thông tin đăng ký (registry bị xoá, vừa cập nhật exe...) -> đăng ký lại
    với các lựa chọn cũ. Trả về True nếu đã sửa."""
    if not is_installed_copy():
        return False
    exe = os.path.abspath(sys.executable)
    if _get(INFO_KEY, "Exe") and _get(INFO_KEY, "Version") == version and _get(UNINSTALL_KEY, "DisplayName") \
            and not _key_exists(LEGACY_SETTINGS_KEY):
        return False
    known = _get(INFO_KEY, "Exe") is not None
    start_menu = os.path.exists(os.path.join(_shell_folder("Programs"), APP + ".lnk"))
    desktop = os.path.exists(os.path.join(_shell_folder("Desktop"), APP + ".lnk"))
    open_with = bool(_get(INFO_KEY, "OpenWith")) if known else True
    office = bool(_get(INFO_KEY, "Office")) if known else office_present()
    ctx = _get(INFO_KEY, "ContextMenu")
    expl = _get(INFO_KEY, "Explorer")
    install(os.path.dirname(exe), start_menu or not known, desktop, open_with, office, version, log,
            context_menu=True if ctx is None else bool(ctx), explorer=True if expl is None else bool(expl))
    return True


def office_present():
    return any(_get(rf"Software\Classes\{a}.Application\CLSID") is not None or
               _has_hkcr(f"{a}.Application") for a in OFFICE_APPS)


def _has_hkcr(name):
    try:
        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, name).Close()
        return True
    except OSError:
        return False


def addin_registered():
    return _get(rf"Software\Microsoft\Office\Word\Addins\{ADDIN_PROGID}", "LoadBehavior") is not None


# ---------------------------------------------------------------------------
#  shortcut
# ---------------------------------------------------------------------------
def _shell_folder(name):
    import win32com.client
    return win32com.client.Dispatch("WScript.Shell").SpecialFolders(name)


def _shortcut(lnk, target, desc):
    import win32com.client
    sh = win32com.client.Dispatch("WScript.Shell")
    s = sh.CreateShortcut(lnk)
    s.TargetPath = target
    s.WorkingDirectory = os.path.dirname(target)
    s.IconLocation = target + ",0"
    s.Description = desc
    s.Save()


def _shortcut_paths():
    out = []
    for folder in ("Programs", "Desktop"):
        try:
            out.append(os.path.join(_shell_folder(folder), APP + ".lnk"))
        except Exception:
            pass
    return out


# ---------------------------------------------------------------------------
#  add-in Office
# ---------------------------------------------------------------------------
def _register_net_class(clsid, cls, assembly, version, dll, name, threading="Both", views=(winreg.KEY_WOW64_64KEY,),
                        extra=None):
    """Đăng ký 1 lớp COM .NET (mscoree.dll nạp assembly theo CodeBase) trong HKCU."""
    codebase = "file:///" + dll.replace("\\", "/")
    for view in views:
        base = rf"Software\Classes\CLSID\{clsid}"
        _del_tree(base, view)                                      # bỏ khoá phiên bản cũ
        _set(base, {"": name, **(extra or {})}, view)
        _set(base + r"\Implemented Categories\{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}", {}, view)
        inproc = {"": "mscoree.dll", "ThreadingModel": threading, "Class": cls, "Assembly": assembly,
                  "RuntimeVersion": "v4.0.30319", "CodeBase": codebase}
        _set(base + r"\InprocServer32", inproc, view)
        _set(base + rf"\InprocServer32\{version}", {k: v for k, v in inproc.items() if k not in ("", "ThreadingModel")}, view)


OUTLOOK_KEEP = r"Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"


def register_addin(dll):
    views = (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY)              # Office 64-bit và 32-bit
    _register_net_class(ADDIN_CLSID, ADDIN_CLASS, ADDIN_ASSEMBLY, ADDIN_VERSION, dll, ADDIN_PROGID, views=views)
    for view in views:
        _set(rf"Software\Classes\CLSID\{ADDIN_CLSID}\ProgId", {"": ADDIN_PROGID}, view)
    _set(rf"Software\Classes\{ADDIN_PROGID}", {"": ADDIN_CLASS})
    _set(rf"Software\Classes\{ADDIN_PROGID}\CLSID", {"": ADDIN_CLSID})
    for a in OFFICE_APPS:
        _set(rf"Software\Microsoft\Office\{a}\Addins\{ADDIN_PROGID}",
             {"FriendlyName": "PDF Toolkit", "Description": "Xuất PDF và mở bằng PDF Toolkit (ký số, mật khẩu, nén...)",
              "LoadBehavior": 3})
        _clear_disabled(a)
    _set(OUTLOOK_KEEP, {ADDIN_PROGID: 1})            # Outlook không tự tắt add-in khi khởi động chậm (lần đầu nạp .NET)


def _clear_disabled(app):
    """Nếu Office từng tự tắt add-in (Resiliency\\DisabledItems) thì bật lại."""
    for ver in ("16.0", "15.0"):
        path = rf"Software\Microsoft\Office\{ver}\{app}\Resiliency\DisabledItems"
        try:
            with winreg.OpenKey(HKCU, path, 0, winreg.KEY_ALL_ACCESS) as k:
                i, names = 0, []
                while True:
                    try:
                        n, v, _ = winreg.EnumValue(k, i)
                    except OSError:
                        break
                    if isinstance(v, bytes) and "pdftoolkitaddin" in v.decode("utf-16-le", "ignore").lower():
                        names.append(n)
                    i += 1
                for n in names:
                    winreg.DeleteValue(k, n)
        except OSError:
            pass


def unregister_addin():
    for a in OFFICE_APPS:
        _del_tree(rf"Software\Microsoft\Office\{a}\Addins\{ADDIN_PROGID}")
    _del_tree(rf"Software\Classes\{ADDIN_PROGID}")
    for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
        _del_tree(rf"Software\Classes\CLSID\{ADDIN_CLSID}", view)
    _del_value(OUTLOOK_KEEP, ADDIN_PROGID)


def _copy_dll(src, folder, name=DLL_NAME):
    """Chép DLL; nếu file cũ đang bị giữ (Office, Explorer đang xem trước...) thì chép sang tên mới."""
    dst = os.path.join(folder, name)
    try:
        shutil.copy2(src, dst)
        return dst
    except PermissionError:
        stem, ext = os.path.splitext(name)
        alt = os.path.join(folder, f"{stem}_{int(time.time())}{ext}")
        shutil.copy2(src, alt)
        return alt


# ---------------------------------------------------------------------------
#  Explorer: ảnh thu nhỏ + khung xem trước file PDF
# ---------------------------------------------------------------------------
PDF_SHELLEX = r"Software\Classes\.pdf\ShellEx"
PREVIEW_LIST = r"Software\Microsoft\Windows\CurrentVersion\PreviewHandlers"


def other_preview_handler():
    """Tên handler xem trước PDF sẵn có trên máy (vd của Microsoft Edge), None nếu không có."""
    for root in (winreg.HKEY_LOCAL_MACHINE,):
        try:
            with winreg.OpenKey(root, rf"{PDF_SHELLEX}\{PREVIEW_IID}") as k:
                clsid = winreg.QueryValueEx(k, "")[0]
        except OSError:
            continue
        if clsid and clsid.upper() != PREVIEW_CLSID.upper():
            try:
                with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}") as k:
                    return winreg.QueryValueEx(k, "")[0] or clsid
            except OSError:
                return clsid
    return None


def register_shell(dll, log=lambda s: None):
    """Ảnh thu nhỏ luôn đăng ký; khung xem trước chỉ đăng ký khi máy chưa có (Edge xem trước tốt hơn - giữ nguyên)."""
    _register_net_class(THUMB_CLSID, "PDFToolkitShell.PdfThumbnailProvider", SHELL_ASSEMBLY, SHELL_VERSION, dll,
                        "PDF Toolkit Thumbnail Provider", threading="Apartment")
    _set(rf"{PDF_SHELLEX}\{THUMB_IID}", {"": THUMB_CLSID})
    other = other_preview_handler()
    if other:
        _unregister_preview()
        log(f"Đã thêm ảnh thu nhỏ trang PDF trong Explorer (khung xem trước giữ của máy: {other})")
    else:
        _register_net_class(PREVIEW_CLSID, "PDFToolkitShell.PdfPreviewHandler", SHELL_ASSEMBLY, SHELL_VERSION, dll,
                            "PDF Toolkit Preview Handler", threading="Apartment",
                            extra={"AppID": PREVHOST_APPID, "DisplayName": "PDF Toolkit"})
        _set(PREVIEW_LIST, {PREVIEW_CLSID: "PDF Toolkit Preview Handler"})
        _set(rf"{PDF_SHELLEX}\{PREVIEW_IID}", {"": PREVIEW_CLSID})
        log("Đã thêm ảnh thu nhỏ và khung xem trước (Preview pane) file PDF trong Explorer")
    _notify_assoc()


def _unregister_preview():
    if (_get(rf"{PDF_SHELLEX}\{PREVIEW_IID}") or "").upper() == PREVIEW_CLSID.upper():
        _del_tree(rf"{PDF_SHELLEX}\{PREVIEW_IID}")
    _del_value(PREVIEW_LIST, PREVIEW_CLSID)
    _del_tree(rf"Software\Classes\CLSID\{PREVIEW_CLSID}", winreg.KEY_WOW64_64KEY)


def unregister_shell():
    """Chỉ gỡ phần của PDF Toolkit - handler của app khác (Edge, PDF-XChange...) tự hiện lại."""
    if (_get(rf"{PDF_SHELLEX}\{THUMB_IID}") or "").upper() == THUMB_CLSID.upper():
        _del_tree(rf"{PDF_SHELLEX}\{THUMB_IID}")
    _del_tree(rf"Software\Classes\CLSID\{THUMB_CLSID}", winreg.KEY_WOW64_64KEY)
    _unregister_preview()
    _notify_assoc()


# ---------------------------------------------------------------------------
#  cài / gỡ
# ---------------------------------------------------------------------------
INTERNAL = "_internal"


def _payload():
    """(exe khởi động, thư mục thư viện) để dựng bản cài dạng thư mục.
    Bản 1 file phải giải nén ~170 MB mỗi lần mở (phần mềm bảo mật quét lại hết -> rất chậm), còn bản thư mục
    chỉ giải nén một lần lúc cài. Bản 1 file mang theo exe khởi động của bản thư mục (launcher\\), thư viện thì
    lấy chính thư mục đang giải nén (_MEIPASS) - hai bản build từ cùng mã nên giống hệt nhau."""
    meipass = getattr(sys, "_MEIPASS", None)
    if not meipass:
        return sys.executable, None
    if _same(meipass, os.path.join(os.path.dirname(sys.executable), INTERNAL)):      # đang chạy bản thư mục
        return sys.executable, meipass
    launcher = os.path.join(meipass, "launcher", EXE_NAME)
    if os.path.isfile(launcher):
        return launcher, meipass
    return sys.executable, None                                                     # bản cũ: chỉ chép exe


def _copy_program(folder, exe):
    src_exe, src_lib = _payload()
    lib = os.path.join(folder, INTERNAL)
    stamp = str(int(time.time()))
    old_exe = old_lib = None
    try:                                    # đổi tên bản cũ trước: thất bại = bản cũ đang chạy, chưa hỏng gì
        if os.path.exists(exe):
            old_exe = exe + f".old{stamp}"
            os.replace(exe, old_exe)
        if os.path.isdir(lib):
            old_lib = lib + f".old{stamp}"
            os.replace(lib, old_lib)
    except PermissionError:
        if old_exe and not os.path.exists(exe):
            os.replace(old_exe, exe)
        raise InstallError("Bản PDF Toolkit đã cài đang mở - hãy đóng tất cả cửa sổ PDF Toolkit rồi thử lại.")
    shutil.copy2(src_exe, exe)
    if src_lib:
        shutil.copytree(src_lib, lib, ignore=lambda d, names: ["launcher"] if _same(d, src_lib) else [])
    for p in (old_exe, old_lib):
        if p:
            shutil.rmtree(p, ignore_errors=True) if os.path.isdir(p) else _try_remove(p)
    for f in os.listdir(folder):            # dọn bản cũ còn sót do lần trước đang bị khoá
        if ".old" in f and f.startswith((EXE_NAME, INTERNAL)):
            p = os.path.join(folder, f)
            shutil.rmtree(p, ignore_errors=True) if os.path.isdir(p) else _try_remove(p)


def _try_remove(p):
    try:
        os.remove(p)
    except OSError:
        pass
def install(folder=None, start_menu=True, desktop=True, open_with=True, office=True, version="", log=print,
            context_menu=True, explorer=True):
    if not frozen():
        raise InstallError("Chỉ cài được từ bản exe (PDF Toolkit.exe).")
    folder = os.path.abspath(folder or default_dir())
    os.makedirs(folder, exist_ok=True)
    exe = os.path.join(folder, EXE_NAME)
    if not _same(sys.executable, exe):
        _copy_program(folder, exe)
    log(f"Đã chép chương trình vào {folder}")

    old = _get(INFO_KEY, "Shortcuts") or ""
    for p in [x for x in old.split("|") if x]:
        try:
            os.remove(p)
        except OSError:
            pass
    links = []
    if start_menu:
        p = os.path.join(_shell_folder("Programs"), APP + ".lnk")
        _shortcut(p, exe, "Bộ công cụ PDF")
        links.append(p)
    if desktop:
        p = os.path.join(_shell_folder("Desktop"), APP + ".lnk")
        _shortcut(p, exe, "Bộ công cụ PDF")
        links.append(p)
    if links:
        log("Đã tạo shortcut: " + ", ".join(os.path.basename(os.path.dirname(x)) for x in links))

    cmd = f'"{exe}" "%1"'
    if open_with:
        _set(rf"Software\Classes\{PROGID}", {"": "Tài liệu PDF (PDF Toolkit)", "FriendlyTypeName": "PDF"})
        _set(rf"Software\Classes\{PROGID}\DefaultIcon", {"": f'"{exe}",0'})
        _set(rf"Software\Classes\{PROGID}\shell\open\command", {"": cmd})
        _set(r"Software\Classes\.pdf\OpenWithProgids", {PROGID: None})
        _set(rf"Software\Classes\Applications\{EXE_NAME}", {"FriendlyAppName": APP})
        _set(rf"Software\Classes\Applications\{EXE_NAME}\shell\open\command", {"": cmd})
        _set(rf"Software\Classes\Applications\{EXE_NAME}\SupportedTypes", {".pdf": ""})
        _set(CAPS_KEY, {"ApplicationName": APP, "ApplicationDescription": "Xem, sửa, ký số, OCR file PDF"})
        _set(CAPS_KEY + r"\FileAssociations", {".pdf": PROGID})
        _set(r"Software\RegisteredApplications", {APP: CAPS_KEY})
        _notify_assoc()
        log('Đã thêm PDF Toolkit vào "Open with" cho file .pdf')
    else:
        _remove_open_with()
    if context_menu:
        register_context_menu(exe)
        log('Đã thêm menu chuột phải "PDF Toolkit" cho file .pdf (Windows 11: trong "Show more options")')
    else:
        unregister_context_menu()

    dll = None
    if office:
        src = resource(DLL_NAME)
        if not src:
            raise InstallError("Bản exe này không kèm add-in Office.")
        dll = _copy_dll(src, folder)
        register_addin(dll)
        log("Đã cài add-in Office (Word, Excel, PowerPoint, Outlook) - mở lại Office để thấy tab / nhóm PDF Toolkit")
    else:
        unregister_addin()

    shell = None
    if explorer and resource(SHELL_DLL):
        shell = _copy_dll(resource(SHELL_DLL), folder, SHELL_DLL)
        register_shell(shell, log)
    else:
        unregister_shell()

    size_kb = sum(os.path.getsize(os.path.join(d, f)) for d, _, fs in os.walk(folder) for f in fs) // 1024
    _set(UNINSTALL_KEY, {"DisplayName": APP, "DisplayVersion": version, "Publisher": "TuanPA",
                         "Contact": "tuanpa1702@gmail.com", "HelpLink": "mailto:tuanpa1702@gmail.com",
                         "Comments": "GNU AGPL v3 - free software, no warranty",
                         "DisplayIcon": f'"{exe}",0', "InstallLocation": folder,
                         "UninstallString": f'"{exe}" --uninstall', "QuietUninstallString": f'"{exe}" --uninstall --quiet',
                         "NoModify": 1, "NoRepair": 1, "EstimatedSize": size_kb,
                         "InstallDate": time.strftime("%Y%m%d")})
    _set(INFO_KEY, {"Exe": exe, "Dir": folder, "Version": version, "Shortcuts": "|".join(links),
                    "OpenWith": int(bool(open_with)), "Office": int(bool(office)), "Addin": dll or "",
                    "ContextMenu": int(bool(context_menu)), "Explorer": int(bool(shell))})
    remove_legacy(exe if open_with else None)
    log("Hoàn tất.")
    return exe


CTX_KEY = r"Software\Classes\SystemFileAssociations\.pdf\shell\PDFToolkit"
# (khoá con, tên hiện trong menu, lệnh --ctx, có vạch ngăn phía trên). Khoá con xếp theo thứ tự chữ cái.
# Tên menu luôn bằng tiếng Anh cho đồng bộ với menu chuột phải của Windows.
CTX_ITEMS = [
    ("01open", "Open with PDF Toolkit", "open", False),
    ("02merge", "Merge selected files into one PDF", "merge", False),
    ("03compress", "Compress PDF", "compress", True),
    ("04split", "Split PDF", "split", False),
    ("05rotate", "Rotate pages", "rotate", False),
    ("06word", "Convert to Word", "pdf_to_word", True),
    ("07excel", "Convert to Excel", "pdf_to_excel", False),
    ("08ppt", "Convert to PowerPoint", "pdf_to_ppt", False),
    ("09img", "Convert to images", "pdf_to_images", False),
    ("10ocr", "OCR - recognize text", "ocr", False),
    ("11protect", "Protect with password", "protect", True),
    ("12restrict", "Restrict printing / copying", "restrict", False),
    ("13unlock", "Remove password", "unlock", False),
    ("14sign", "Digitally sign", "sign", False),
]


def register_context_menu(exe):
    """Menu chuột phải cho file .pdf: mục 'PDF Toolkit' có menu con (Windows 11: trong 'Show more options')."""
    _del_tree(CTX_KEY)
    _set(CTX_KEY, {"MUIVerb": APP, "Icon": f'"{exe}",0', "SubCommands": ""})
    for key, label, cmd, sep in CTX_ITEMS:
        vals = {"MUIVerb": label, "MultiSelectModel": "Player"}
        if key == "01open":
            vals["Icon"] = f'"{exe}",0'
        if sep:
            vals["CommandFlags"] = 0x20                     # ECF_SEPARATORBEFORE
        _set(rf"{CTX_KEY}\shell\{key}", vals)
        _set(rf"{CTX_KEY}\shell\{key}\command", {"": f'"{exe}" --ctx {cmd} "%1"'})
    _notify_assoc()


def unregister_context_menu():
    _del_tree(CTX_KEY)
    _notify_assoc()


def _remove_open_with():
    _del_tree(rf"Software\Classes\{PROGID}")
    _del_value(r"Software\Classes\.pdf\OpenWithProgids", PROGID)
    _del_tree(rf"Software\Classes\Applications\{EXE_NAME}")
    _del_value(r"Software\RegisteredApplications", APP)
    _del_tree(CAPS_KEY)
    _notify_assoc()


SETTINGS_KEY = rf"Software\{ORG}\PDFToolkit"             # QSettings(ORG, "PDFToolkit") + Install/Capabilities


def _appdata(org=ORG):
    appdata = os.environ.get("APPDATA") or os.path.expanduser(r"~\AppData\Roaming")
    return os.path.join(appdata, org, "PDFToolkit")


def user_data_dirs():
    """Các thư mục dữ liệu người dùng của PDF Toolkit trên máy này (con dấu, cache font, file tạm)."""
    import tempfile
    local = os.environ.get("LOCALAPPDATA") or os.path.expanduser(r"~\AppData\Local")
    tmp = tempfile.gettempdir()
    return [_appdata(),                                              # thư viện con dấu
            os.path.join(local, "PDFToolkit"),                       # cache font, font đã dựng
            os.path.join(tmp, "PDFToolkit_ui"), os.path.join(tmp, "PDFToolkit_mail")]


def _del_if_empty(path):
    """Xoá khoá HKCU rỗng (không còn khoá con / giá trị) - vd Software\\TuanPA khi không còn app nào khác."""
    try:
        with winreg.OpenKey(HKCU, path) as k:
            info = winreg.QueryInfoKey(k)
        if info[0] == 0 and info[1] == 0:
            winreg.DeleteKey(HKCU, path)
    except OSError:
        pass


def _rmdir_if_empty(path):
    try:
        os.rmdir(path)
    except OSError:
        pass


def wipe_user_data(log=print):
    """Xoá mọi dữ liệu người dùng: file gần đây, ngôn ngữ, tuỳ chọn, chữ ký / con dấu đã lưu, cache.
    Gọi khi gỡ cài đặt -> cài lại như máy mới."""
    _del_tree(SETTINGS_KEY)
    _del_if_empty(SETTINGS_KEY.rsplit("\\", 1)[0])
    left = []
    for d in user_data_dirs():
        if os.path.isdir(d):
            shutil.rmtree(d, ignore_errors=True)
            if os.path.exists(d):
                left.append(d)
    _rmdir_if_empty(os.path.dirname(_appdata()))
    log("Đã xoá dữ liệu người dùng: file gần đây, cài đặt, thư viện con dấu, cache.")
    if left:
        log("Còn sót (đang được dùng): " + "; ".join(left))


# ---------------------------------------------------------------------------
#  chuyển từ tên cũ NagaseVN (bản 1.3.0 trở về trước)
# ---------------------------------------------------------------------------
def _copy_tree(src, dst, skip=()):
    """Chép khoá registry HKCU src -> dst (giữ kiểu giá trị). skip: tên khoá con bỏ qua."""
    try:
        k = winreg.OpenKey(HKCU, src)
    except OSError:
        return
    with k:
        with winreg.CreateKeyEx(HKCU, dst, 0, winreg.KEY_WRITE) as out:
            i = 0
            while True:
                try:
                    name, val, typ = winreg.EnumValue(k, i)
                except OSError:
                    break
                winreg.SetValueEx(out, name, 0, typ, val)
                i += 1
        i = 0
        while True:
            try:
                sub = winreg.EnumKey(k, i)
            except OSError:
                break
            if sub not in skip:
                _copy_tree(src + "\\" + sub, dst + "\\" + sub)
            i += 1


def migrate_legacy():
    """Lần đầu bản mới chạy: chép cài đặt (ngôn ngữ, tuỳ chọn, file gần đây, thông tin cài) và thư viện con dấu
    từ tên cũ sang tên mới. Chỉ chép, chưa xoá - bản cũ đang cài vẫn chạy được; dọn khi cài / sửa đăng ký."""
    moved = False
    if _get(LEGACY_SETTINGS_KEY + r"\Install", "Exe") is not None or _key_exists(LEGACY_SETTINGS_KEY):
        if not _key_exists(SETTINGS_KEY):
            _copy_tree(LEGACY_SETTINGS_KEY, SETTINGS_KEY, skip=("Capabilities",))
            moved = True
    old, new = _appdata(LEGACY_ORG), _appdata()
    if os.path.isdir(old) and not os.path.exists(new):
        shutil.copytree(old, new)
        moved = True
    return moved


def _key_exists(path):
    try:
        winreg.OpenKey(HKCU, path).Close()
        return True
    except OSError:
        return False


def _user_choice_progid():
    try:
        with winreg.OpenKey(HKCU, r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice") as k:
            return winreg.QueryValueEx(k, "ProgId")[0]
    except OSError:
        return None


def remove_legacy(exe=None):
    """Dọn đăng ký và dữ liệu dưới tên cũ (sau khi đã chép sang tên mới). Nếu người dùng đang chọn PDF Toolkit làm
    app PDF mặc định bằng ProgID cũ thì giữ ProgID đó (trỏ sang exe mới) để Windows không bỏ lựa chọn mặc định."""
    migrate_legacy()
    if exe and _user_choice_progid() == LEGACY_PROGID:
        _set(rf"Software\Classes\{LEGACY_PROGID}", {"": "Tài liệu PDF (PDF Toolkit)", "FriendlyTypeName": "PDF"})
        _set(rf"Software\Classes\{LEGACY_PROGID}\DefaultIcon", {"": f'"{exe}",0'})
        _set(rf"Software\Classes\{LEGACY_PROGID}\shell\open\command", {"": f'"{exe}" "%1"'})
    else:
        _del_tree(rf"Software\Classes\{LEGACY_PROGID}")
        _del_value(r"Software\Classes\.pdf\OpenWithProgids", LEGACY_PROGID)
    _del_tree(LEGACY_UNINSTALL_KEY)
    _del_tree(LEGACY_SETTINGS_KEY)
    _del_if_empty(LEGACY_SETTINGS_KEY.rsplit("\\", 1)[0])
    old = _appdata(LEGACY_ORG)
    if os.path.isdir(old) and os.path.isdir(_appdata()):
        shutil.rmtree(old, ignore_errors=True)
        _rmdir_if_empty(os.path.dirname(old))


def uninstall(log=print, only_self=False):
    """only_self=True (lệnh --uninstall): chỉ gỡ khi đang chạy từ chính bản đã cài - tránh một bản exe
    khác (bản portable, bản thử) xoá mất bản người dùng đã cài."""
    migrate_legacy()                                         # bản cài từ thời tên cũ: thông tin cài nằm ở khoá cũ
    folder = _get(INFO_KEY, "Dir")
    exe = installed_exe()
    if only_self and frozen() and exe and not _same(sys.executable, exe):
        raise InstallError(f"Bản đã cài nằm ở {os.path.dirname(exe)} - hãy gỡ từ bản đó "
                           "(Settings > Apps > PDF Toolkit).")
    for p in [x for x in (_get(INFO_KEY, "Shortcuts") or "").split("|") if x] + _shortcut_paths():
        try:
            os.remove(p)
        except OSError:
            pass
    _remove_open_with()
    unregister_context_menu()
    unregister_addin()
    unregister_shell()
    _del_tree(UNINSTALL_KEY)
    _del_tree(INFO_KEY)
    log("Đã gỡ shortcut, liên kết file PDF, add-in Office và phần mở rộng Explorer.")
    remove_legacy()
    wipe_user_data(log)
    if folder and os.path.isdir(folder) and os.path.basename(folder.rstrip("\\")) == APP:
        running_here = frozen() and os.path.normcase(os.path.dirname(os.path.abspath(sys.executable))) == \
            os.path.normcase(os.path.abspath(folder))
        if running_here:
            # exe đang chạy không tự xoá được -> nhờ cmd xoá sau khi chương trình thoát
            subprocess.Popen(f'cmd /c ping 127.0.0.1 -n 4 >nul & rmdir /s /q "{folder}"',
                             creationflags=0x08000000, close_fds=True)
            log(f"Thư mục {folder} sẽ được xoá sau khi đóng chương trình.")
        else:
            shutil.rmtree(folder, ignore_errors=True)
            if os.path.exists(folder):
                log(f"Một số file đang được dùng (Office đang mở?) - xoá tay thư mục {folder} sau.")
            else:
                log(f"Đã xoá {folder}")

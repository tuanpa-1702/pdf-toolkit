# -*- mode: python -*-
# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
# Cấu hình đóng gói PyInstaller (build.bat gọi 2 lần: PTK_MODE=onedir rồi PTK_MODE=onefile).
# Dùng file spec thay cho tham số dòng lệnh để lọc bớt DLL Qt không dùng (giảm dung lượng exe).
import os

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

MODE = os.environ.get("PTK_MODE", "onedir")
NAME = "PDF Toolkit"
# PTK_HIDE_AUTHOR=1 (build_noauthor.bat): bản không ghi "by TuanPA" cạnh logo. Cờ ghi vào module riêng trong
# build\flags (không nằm cạnh mã nguồn, nên chạy bằng python vẫn là bản thường)
FLAGS_DIR = os.path.join(SPECPATH, "build", "flags")
os.makedirs(FLAGS_DIR, exist_ok=True)
with open(os.path.join(FLAGS_DIR, "build_flags.py"), "w", encoding="utf-8") as fh:
    fh.write(f"HIDE_AUTHOR = {os.environ.get('PTK_HIDE_AUTHOR') == '1'}\n")
ONEDIR_DIST = os.environ.get("PTK_ONEDIR_DIST", "build/dist_dir")

hidden = (collect_submodules("pyhanko") + collect_submodules("pyhanko_certvalidator")
          + collect_submodules("fontTools.varLib") + collect_submodules("markdown")
          + ["fontTools.ttLib", "i18n_en", "pythoncom", "win32com.client", "imgops", "cv2_lite", "build_flags"])
datas = ([("app.ico", "."), ("tessdata", "tessdata"), ("office_addin/PDFToolkitAddin.dll", "office_addin"),
         ("shell_ext/PDFToolkitShell.dll", "shell_ext")]
         + copy_metadata("pyhanko") + collect_data_files("pptx"))
if MODE == "onefile":
    datas += [(ONEDIR_DIST + "/PDF Toolkit/PDF Toolkit.exe", "launcher"), ("licenses", "licenses")]

excludes = ["uharfbuzz", "matplotlib", "scipy", "pandas", "tkinter",
            "cv2",                                  # OpenCV ~110 MB: thay bằng imgops / cv2_lite
            "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtQml", "PySide6.QtQuick",
            "PySide6.Qt3DCore", "PySide6.QtMultimedia", "PySide6.QtCharts", "PySide6.QtDataVisualization",
            "PySide6.QtPdf", "PySide6.QtOpenGL", "PySide6.QtOpenGLWidgets"]

# DLL / plugin Qt không dùng: bàn phím ảo kéo theo Qt Quick + QML + OpenGL; plugin ảnh PDF kéo theo Qt6Pdf;
# OpenGL giả lập (opengl32sw 20 MB); TLS / mạng (chỉ dùng named pipe nội bộ); bản dịch có sẵn của Qt.
DROP = ("pyside6/opengl32sw.dll", "pyside6/qt6quick", "pyside6/qt6qml", "pyside6/qt6opengl",
        "pyside6/qt6pdf", "pyside6/qt6virtualkeyboard", "pyside6/plugins/platforminputcontexts/",
        "pyside6/plugins/imageformats/qpdf", "pyside6/plugins/tls/", "pyside6/plugins/networkinformation/",
        "pyside6/plugins/generic/", "pyside6/translations/")


def keep(entry):
    name = entry[0].replace("\\", "/").lower()
    return not any(d in name for d in DROP)


a = Analysis(["pdf_toolkit.py"], pathex=[FLAGS_DIR], binaries=[], datas=datas, hiddenimports=hidden,
             hookspath=[], runtime_hooks=[], excludes=excludes, noarchive=False)
a.binaries = [b for b in a.binaries if keep(b)]
a.datas = [d for d in a.datas if keep(d)]
pyz = PYZ(a.pure)

if MODE == "onefile":
    exe = EXE(pyz, a.scripts, a.binaries, a.datas, [], name=NAME, console=False, icon="app.ico",
              upx=False, runtime_tmpdir=None)
else:
    exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name=NAME, console=False, icon="app.ico", upx=False)
    coll = COLLECT(exe, a.binaries, a.datas, name=NAME, upx=False)

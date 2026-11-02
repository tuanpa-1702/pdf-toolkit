# PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
# SPDX-License-Identifier: AGPL-3.0-only
# Phần mềm tự do theo GNU AGPL v3, KHÔNG CÓ BẢO HÀNH - xem file LICENSE.
"""Gom giấy phép các thư viện bên thứ ba được đóng vào exe -> thư mục licenses\\ (index.json + file văn bản).
Chạy sau bước build onedir (build.bat tự gọi): đọc danh sách file PyInstaller đã đóng gói trong
build\\work_dir\\pdf_toolkit\\*.toc, tra ngược ra gói pip, rồi chép file LICENSE / COPYING / NOTICE của từng gói."""
import ast, glob, json, os, re, shutil, sys
import importlib.metadata as md

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.join(HERE, os.environ.get("PTK_WORK", r"build\work_dir"), "pdf_toolkit")   # theo tên file spec
OUT = os.path.join(HERE, "licenses")
SRC = os.path.join(HERE, "licenses_src")
SUPPORT_EMAIL = "tuanpa1702@gmail.com"

# Thành phần chính: (tên gói pip, tên hiển thị, công dụng, trang chủ)
MAIN = [
    ("pymupdf", "PyMuPDF / MuPDF", "Đọc, sửa, hiển thị PDF; OCR (Tesseract tích hợp)", "https://github.com/pymupdf/PyMuPDF"),
    ("pyside6-essentials", "Qt 6 / PySide6", "Giao diện", "https://www.qt.io/qt-for-python"),
    ("shiboken6", "Shiboken6", "Cầu nối Python - Qt", "https://www.qt.io/qt-for-python"),
    ("pyhanko", "pyHanko", "Ký số PDF (PAdES), dấu thời gian", "https://github.com/MatthiasValvekens/pyHanko"),
    ("pyhanko-certvalidator", "pyhanko-certvalidator", "Kiểm tra chứng thư số", "https://github.com/MatthiasValvekens/certvalidator"),
    ("cryptography", "cryptography", "Mã hoá", "https://github.com/pyca/cryptography"),
    ("asn1crypto", "asn1crypto", "Đọc cấu trúc chứng thư số", "https://github.com/wbond/asn1crypto"),
    ("numpy", "NumPy", "Xử lý ảnh", "https://numpy.org"),
    ("pillow", "Pillow", "Xử lý ảnh", "https://python-pillow.org"),
    ("fonttools", "fontTools", "Font chữ khi sửa chữ", "https://github.com/fonttools/fonttools"),
    ("pdf2docx", "pdf2docx", "PDF -> Word", "https://github.com/ArtifexSoftware/pdf2docx"),
    ("python-docx", "python-docx", "PDF -> Word", "https://github.com/python-openxml/python-docx"),
    ("openpyxl", "openpyxl", "PDF -> Excel", "https://openpyxl.readthedocs.io"),
    ("python-pptx", "python-pptx", "PDF -> PowerPoint", "https://github.com/scanny/python-pptx"),
    ("xlsxwriter", "XlsxWriter", "Thư viện đi kèm python-pptx", "https://github.com/jmcnamara/XlsxWriter"),
    ("lxml", "lxml", "Đọc XML", "https://lxml.de"),
    ("markdown", "Python-Markdown", "Markdown -> PDF", "https://python-markdown.github.io"),
    ("pywin32", "pywin32", "Gọi Windows / Office", "https://github.com/mhammond/pywin32"),
    ("aiohttp", "aiohttp", "Kết nối máy chủ dấu thời gian", "https://github.com/aio-libs/aiohttp"),
    ("pyinstaller", "PyInstaller (bootloader)", "Đóng gói exe", "https://pyinstaller.org"),
]
# Thành phần không phải gói pip
EXTRA = [
    {"name": "Python", "version": sys.version.split()[0], "license": "PSF-2.0",
     "purpose": "Trình thông dịch", "url": "https://www.python.org", "text_from": ("python",)},
    {"name": "Tesseract OCR + tessdata_best (vie, eng, osd)", "version": "", "license": "Apache-2.0",
     "purpose": "Nhận dạng chữ (OCR) - nằm trong MuPDF", "url": "https://github.com/tesseract-ocr/tesseract",
     "text_from": ("apache",)},
    {"name": "Liberation Sans (font lớp chữ OCR)", "version": "2", "license": "OFL-1.1",
     "purpose": "Font lớp chữ ẩn sau khi OCR và chữ của 'OCR → PDF chỉnh sửa được' (Regular + Bold)", "url": "https://github.com/liberationfonts/liberation-fonts",
     "text_from": ("ofl",)},
]
OFL = """Digitized data copyright (c) 2010 Google Corporation.
Copyright (c) 2012 Red Hat, Inc.

This Font Software is licensed under the SIL Open Font License, Version 1.1.
Full text: https://openfontlicense.org/open-font-license-official-text/

Permission is hereby granted, free of charge, to any person obtaining a copy of the Font Software, to use, study,
copy, merge, embed, modify, redistribute, and sell modified and unmodified copies of the Font Software, subject to
the conditions of the license (the Font Software may be bundled with any software, but may not be sold by itself;
Reserved Font Names may not be used by modified versions; the license must accompany the Font Software).
"""
TEXT_NAMES = re.compile(r"(^|[/\\])(LICEN[CS]E|COPYING|NOTICE|AUTHORS)[^/\\]*$", re.I)


def norm(n):
    return re.sub(r"[-_.]+", "-", n).lower()


def bundled_files():
    """Đường dẫn nguồn của mọi file PyInstaller đã đóng gói (module .py, .pyd, .dll, data)."""
    out = set()
    for toc in glob.glob(os.path.join(WORK, "*.toc")):
        try:
            data = ast.literal_eval(open(toc, encoding="utf-8").read())
        except Exception:
            continue
        stack = [data]
        while stack:
            x = stack.pop()
            if isinstance(x, (list, tuple)):
                if len(x) == 3 and all(isinstance(v, str) for v in x) and os.path.isabs(x[1]):
                    out.add(os.path.normcase(os.path.abspath(x[1])))
                else:
                    stack.extend(x)
    return out


def license_text(dist):
    parts = []
    for f in dist.files or []:
        if TEXT_NAMES.search(str(f)):
            try:
                t = open(dist.locate_file(f), encoding="utf-8", errors="replace").read()
            except Exception:
                continue
            if t.strip():
                parts.append(f"===== {os.path.basename(str(f))} =====\n{t.strip()}\n")
    return "\n".join(parts)


def lic_name(dist):
    m = dist.metadata
    v = m.get("License-Expression") or ""
    if not v:
        cls = [c.split("::")[-1].strip() for c in m.get_all("Classifier") or [] if c.startswith("License ::")]
        v = "; ".join(cls) or (m.get("License") or "").splitlines()[0][:80] if (cls or m.get("License")) else ""
    return v.strip() or "xem văn bản giấy phép"


def main():
    files = bundled_files()
    if not files:
        sys.exit("Chưa có build\\work_dir - chạy bước build onedir trước.")
    used = {}
    for d in md.distributions():
        name = norm(d.metadata["Name"] or "")
        for f in d.files or []:
            p = os.path.normcase(os.path.abspath(str(d.locate_file(f))))
            if p in files:
                used[name] = d
                break
    os.makedirs(OUT, exist_ok=True)
    for f in os.listdir(OUT):                       # xoá nội dung cũ (giữ thư mục - có thể đang được mở)
        os.remove(os.path.join(OUT, f))
    index, done = [], set()

    # văn bản GPL / LGPL / AGPL đầy đủ (gói pip của PyMuPDF / PySide6 chỉ kèm 1 dòng tóm tắt)
    std = {}
    for n in ("AGPL-3.0", "LGPL-3.0", "GPL-3.0"):
        p = os.path.join(SRC, n + ".txt")
        if os.path.exists(p):
            std[n] = open(p, encoding="utf-8").read()
            shutil.copy(p, os.path.join(OUT, n + ".txt"))
    attach = {"pymupdf": ["AGPL-3.0"], "pyside6-essentials": ["LGPL-3.0", "GPL-3.0"],
              "shiboken6": ["LGPL-3.0", "GPL-3.0"]}

    def add(dist, name, purpose, url):
        key = norm(dist.metadata["Name"])
        fn = key + ".txt"
        text = license_text(dist) or f"{name} - giấy phép: {lic_name(dist)}\n{url}\n"
        for n in attach.get(key, []):
            if n in std:
                text += f"\n===== {n} =====\n{std[n]}"
        open(os.path.join(OUT, fn), "w", encoding="utf-8").write(text)
        index.append({"name": name, "version": dist.version, "license": lic_name(dist), "purpose": purpose,
                      "url": url or dist.metadata.get("Home-page") or "", "file": fn})
        done.add(key)

    # chính PDF Toolkit (AGPL-3.0): NOTICE + toàn văn giấy phép
    app_ver = re.search(r'APP_VERSION\s*=\s*"([^"]+)"',
                        open(os.path.join(HERE, "pdf_toolkit.py"), encoding="utf-8").read()).group(1)
    notice = open(os.path.join(HERE, "NOTICE.md"), encoding="utf-8").read()
    shutil.copy(os.path.join(HERE, "NOTICE.md"), os.path.join(OUT, "NOTICE.txt"))
    open(os.path.join(OUT, "pdf-toolkit.txt"), "w", encoding="utf-8").write(
        notice + "\n===== AGPL-3.0 =====\n" + std.get("AGPL-3.0", ""))
    index.append({"name": "PDF Toolkit (chương trình này)", "version": app_ver, "license": "AGPL-3.0-only",
                  "purpose": "Chương trình chính - © 2026 TuanPA", "url": "mailto:" + SUPPORT_EMAIL,
                  "file": "pdf-toolkit.txt"})

    for pkg, name, purpose, url in MAIN:
        d = used.get(norm(pkg))
        if d is None:
            try:
                d = md.distribution(pkg)
            except md.PackageNotFoundError:
                continue
        add(d, name, purpose, url)
    for e in EXTRA:
        src = ""
        if "apache" in e["text_from"]:
            src = license_text(md.distribution("cryptography"))
            m = re.search(r"===== LICENSE\.APACHE =====\n(.*?)(\n=====|\Z)", src, re.S)
            src = m.group(1) if m else ""
        elif "ofl" in e["text_from"]:
            src = OFL
        elif "python" in e["text_from"]:
            p = os.path.join(sys.base_prefix, "LICENSE.txt")
            src = open(p, encoding="utf-8", errors="replace").read() if os.path.exists(p) else ""
        fn = norm(e["name"].split()[0]) + ".txt"
        open(os.path.join(OUT, fn), "w", encoding="utf-8").write(src or f"{e['name']} - {e['license']}\n{e['url']}\n")
        index.append({k: v for k, v in e.items() if k != "text_from"} | {"file": fn})
    for key, d in sorted(used.items()):
        if key not in done and key not in ("pyside6", "pyside6-addons"):
            add(d, d.metadata["Name"], "Thư viện phụ thuộc", d.metadata.get("Home-page") or "")
    json.dump(index, open(os.path.join(OUT, "index.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    write_sources(index, app_ver)
    make_source_zip(app_ver)
    print(f"licenses: {len(index)} thành phần -> {OUT}")


# nơi lấy mã nguồn của các thành phần không lấy được từ PyPI
SOURCE_URLS = {
    "Qt 6 / PySide6": ["https://code.qt.io/cgit/pyside/pyside-setup.git (tag v{v})",
                       "https://download.qt.io/official_releases/qt/ (Qt {v}, mã nguồn thư viện Qt)"],
    "Shiboken6": ["https://code.qt.io/cgit/pyside/pyside-setup.git (tag v{v})"],
    "PyMuPDF / MuPDF": ["https://github.com/pymupdf/PyMuPDF/releases/tag/{v}",
                        "https://mupdf.com/releases (MuPDF, gồm Tesseract / Leptonica)"],
    "Python": ["https://www.python.org/downloads/source/ (Python {v})"],
    "Tesseract OCR + tessdata_best (vie, eng, osd)": ["https://github.com/tesseract-ocr/tesseract",
                                                      "https://github.com/tesseract-ocr/tessdata_best"],
    "Liberation Sans (font lớp chữ OCR)": ["https://github.com/liberationfonts/liberation-fonts"],
    "PyInstaller (bootloader)": ["https://github.com/pyinstaller/pyinstaller/releases/tag/v{v}"],
}
PYPI_NAME = {"Qt 6 / PySide6": "pyside6-essentials", "PyMuPDF / MuPDF": "pymupdf",
             "Python-Markdown": "markdown", "PyInstaller (bootloader)": "pyinstaller", "NumPy": "numpy",
             "Pillow": "pillow", "fontTools": "fonttools", "pyHanko": "pyhanko"}


def write_sources(index, app_ver):
    lines = [f"PDF Toolkit {app_ver} - thành phần bên thứ ba và nơi lấy mã nguồn / third-party components and "
             "source locations", "Các thành phần được dùng nguyên bản, không sửa đổi / All components are used "
             "unmodified.",
             f"Mã nguồn: tải tại link chính thức bên dưới, hoặc liên hệ / Source code: official links below, "
             f"or contact {SUPPORT_EMAIL}", ""]
    for e in index:
        if e["file"] == "pdf-toolkit.txt":
            continue
        v = e.get("version", "")
        lines.append(f"{e['name']} {v}  [{e.get('license', '')}]")
        urls = [u.format(v=v) for u in SOURCE_URLS.get(e["name"], [])]
        if e["name"] not in ("Python", "Tesseract OCR + tessdata_best (vie, eng, osd)",
                             "Liberation Sans (font lớp chữ OCR)"):
            pkg = PYPI_NAME.get(e["name"], e["name"]).lower()
            urls.append(f"https://pypi.org/project/{pkg}/{v}/#files")
        if e.get("url") and not e["url"].startswith("mailto:"):
            urls.append(e["url"])
        lines += ["    " + u for u in dict.fromkeys(urls)] + [""]
    open(os.path.join(OUT, "THIRD_PARTY_SOURCES.txt"), "w", encoding="utf-8").write("\n".join(lines))


SOURCE_FILES = ["*.py", "pdf_toolkit.spec", "build.bat", "requirements.txt", "README.md", "NOTICE.md", "LICENSE", "app.ico",
                r"office_addin\*.cs", r"office_addin\*.bat", r"shell_ext\*.cs", r"shell_ext\*.bat", r"licenses_src\*.txt"]


def make_source_zip(app_ver):
    """Mã nguồn tương ứng của chính PDF Toolkit (AGPL mục 6) - đóng kèm exe."""
    import zipfile
    root = f"PDF-Toolkit-{app_ver}-source"
    with zipfile.ZipFile(os.path.join(OUT, "PDF-Toolkit-source.zip"), "w", zipfile.ZIP_DEFLATED) as z:
        for pat in SOURCE_FILES:
            for p in sorted(glob.glob(os.path.join(HERE, pat))):
                rel = os.path.relpath(p, HERE)
                z.write(p, os.path.join(root, rel))
        z.writestr(os.path.join(root, "tessdata", "README.txt"),
                   "Dữ liệu OCR: tải eng / vie / osd .traineddata từ https://github.com/tesseract-ocr/tessdata_best\n"
                   "và font Liberation Sans (Regular đổi tên thành ocr_font.ttf, Bold thành ocr_font_bold.ttf)\n"
                   "vào thư mục này trước khi build.\n")


if __name__ == "__main__":
    main()

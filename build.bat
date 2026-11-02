@echo off
rem PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
rem SPDX-License-Identifier: AGPL-3.0-only
rem Phan mem tu do theo GNU AGPL v3, KHONG CO BAO HANH - xem file LICENSE.
rem Dong goi PDF Toolkit thanh 1 file .exe (ket qua: dist\PDF Toolkit.exe)
rem Buoc 1 build ban thu muc (onedir) - exe khoi dong cua no duoc nhung vao ban 1 file;
rem khi "Cai dat vao may", installer dung lai ban thu muc tu exe nay + thu vien da giai nen -> mo nhanh.
cd /d "%~dp0"
python -m pip install -r requirements.txt
call office_addin\build_addin.bat || (pause & exit /b 1)
call "%~dp0shell_ext\build_shell.bat" || (pause & exit /b 1)
cd /d "%~dp0"
rem cau hinh dong goi nam trong pdf_toolkit.spec (loc bot DLL Qt khong dung, bo OpenCV)
set PTK_MODE=onedir
python -m PyInstaller --noconfirm --clean --distpath build\dist_dir --workpath build\work_dir pdf_toolkit.spec || (pause & exit /b 1)
rem gom giay phep ben thu ba (doc danh sach file da dong goi o buoc 1) -> licenses\
python collect_licenses.py || (pause & exit /b 1)
set PTK_MODE=onefile
python -m PyInstaller --noconfirm --clean pdf_toolkit.spec || (pause & exit /b 1)
echo.
echo Xong: dist\PDF Toolkit.exe
pause

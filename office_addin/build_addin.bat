@echo off
rem PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
rem SPDX-License-Identifier: AGPL-3.0-only
rem Phan mem tu do theo GNU AGPL v3, KHONG CO BAO HANH - xem file LICENSE.
rem Build add-in Office (PDFToolkitAddin.dll) bang csc.exe co san cua .NET Framework 4 - khong can cai them gi.
cd /d "%~dp0"
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
"%CSC%" /nologo /target:library /platform:anycpu /optimize+ /codepage:65001 ^
  /reference:Microsoft.CSharp.dll /reference:System.Windows.Forms.dll /reference:System.Drawing.dll ^
  /reference:System.Web.Extensions.dll ^
  /out:PDFToolkitAddin.dll *.cs
if errorlevel 1 (echo LOI khi build add-in & exit /b 1)
echo Xong: office_addin\PDFToolkitAddin.dll

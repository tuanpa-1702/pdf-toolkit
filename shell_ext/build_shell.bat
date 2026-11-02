@echo off
rem PDF Toolkit - Copyright (C) 2026 TuanPA <tuanpa1702@gmail.com>
rem SPDX-License-Identifier: AGPL-3.0-only
rem Phan mem tu do theo GNU AGPL v3, KHONG CO BAO HANH - xem file LICENSE.
rem Build phan mo rong Explorer (PDFToolkitShell.dll: anh thu nho + xem truoc PDF) bang csc.exe cua .NET Framework 4.
rem Ve trang PDF bang Windows.Data.Pdf (WinRT co san trong Windows 10/11).
cd /d "%~dp0"
set FW=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319
if not exist "%FW%\csc.exe" set FW=%WINDIR%\Microsoft.NET\Framework\v4.0.30319
set WM=%WINDIR%\System32\WinMetadata
"%FW%\csc.exe" /nologo /target:library /platform:anycpu /optimize+ /codepage:65001 ^
  /reference:System.Windows.Forms.dll /reference:System.Drawing.dll ^
  /reference:"%FW%\System.Runtime.dll" /reference:"%FW%\System.Runtime.WindowsRuntime.dll" ^
  /reference:"%FW%\System.Threading.Tasks.dll" /reference:"%FW%\System.IO.dll" ^
  /reference:"%WM%\Windows.Foundation.winmd" /reference:"%WM%\Windows.Data.winmd" /reference:"%WM%\Windows.Storage.winmd" ^
  /out:PDFToolkitShell.dll *.cs
if errorlevel 1 (echo LOI khi build phan mo rong Explorer & exit /b 1)
echo Xong: shell_ext\PDFToolkitShell.dll

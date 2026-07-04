@echo off
@chcp 65001 >nul
setlocal
cd /d "%~dp0"

rem ============================================================
rem  標案文件產生器 - 建置腳本
rem  只使用 Windows 內建的 .NET Framework C# 編譯器，
rem  不需要安裝 Visual Studio 或任何開發工具。
rem ============================================================

set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
if not exist "%CSC%" (
    echo [ERROR] csc.exe not found
    goto :fail
)

"%CSC%" /nologo /target:winexe /platform:anycpu /optimize+ ^
  /codepage:65001 /utf8output ^
  /r:System.dll /r:System.Core.dll /r:System.Xml.dll /r:System.Xml.Linq.dll ^
  /r:System.Windows.Forms.dll /r:System.Drawing.dll ^
  /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll ^
  /win32manifest:原始碼\app.manifest ^
  /out:標案文件產生器.exe 原始碼\*.cs

if errorlevel 1 goto :fail

echo.
echo [OK] 已產生 標案文件產生器.exe
if "%~1"=="" pause
exit /b 0

:fail
echo.
echo [FAIL] 建置發生錯誤，請將上方訊息提供給資訊人員。
if "%~1"=="" pause
exit /b 1

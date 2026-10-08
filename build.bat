@echo off
rem Build PdfTool.exe using the C# compiler bundled with Windows (.NET Framework 4.x)
cd /d "%~dp0"
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
"%CSC%" /nologo /target:winexe /codepage:65001 /optimize+ /win32icon:PdfTool.ico /out:PdfTool.exe ^
  /r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll PdfTool.cs
if errorlevel 1 (echo. & echo BIEN DICH THAT BAI & pause & exit /b 1)
echo Da tao PdfTool.exe

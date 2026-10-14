@echo off
rem Build the installer dist\PdfTool-Setup.exe (run build.bat first after editing the source)
cd /d "%~dp0"
call "%~dp0package.bat" < nul
if errorlevel 1 exit /b 1

set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe
set REFS=/r:System.dll /r:System.Drawing.dll /r:System.Windows.Forms.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll
set OPTS=/nologo /target:winexe /codepage:65001 /optimize+ /win32icon:PdfTool.ico /win32manifest:installer\app.manifest

rem Uninstaller (same source, no payload)
"%CSC%" %OPTS% %REFS% /out:dist\uninstall.exe installer\Setup.cs
if errorlevel 1 (echo BIEN DICH THAT BAI & pause & exit /b 1)
rem Installer with the packaged app embedded
"%CSC%" %OPTS% %REFS% /resource:dist\PdfTool.zip,payload.zip /resource:dist\uninstall.exe,uninstall.exe /out:dist\PdfTool-Setup.exe installer\Setup.cs
if errorlevel 1 (echo BIEN DICH THAT BAI & pause & exit /b 1)
del dist\uninstall.exe
echo.
echo Da tao: %~dp0dist\PdfTool-Setup.exe

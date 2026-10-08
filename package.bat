@echo off
rem Package PdfTool for other computers: dist\PdfTool\ and dist\PdfTool.zip
rem Bundles qpdf, Ghostscript, Tesseract (copied from the PDF24 install) so PDF24 is not needed there.
cd /d "%~dp0"
set SRC=%ProgramFiles%\PDF24
set OUT=dist\PdfTool
set T=%OUT%\tools

if not exist "%SRC%\qpdf\bin\qpdf.exe" (echo Khong tim thay PDF24 tai "%SRC%". & pause & exit /b 1)

rem Uses the existing PdfTool.exe (run build.bat first after editing the source)
if not exist PdfTool.exe (echo Chua co PdfTool.exe, hay chay build.bat truoc. & pause & exit /b 1)

if exist dist rmdir /s /q dist
mkdir "%T%\licenses"

copy /y PdfTool.exe "%OUT%\" >nul
robocopy "%SRC%\qpdf\bin" "%T%\qpdf\bin" /e /njh /njs /nfl /ndl >nul
robocopy "%SRC%\gs\bin" "%T%\gs\bin" /e /njh /njs /nfl /ndl >nul
robocopy "%SRC%\gs\lib" "%T%\gs\lib" /e /njh /njs /nfl /ndl >nul
robocopy "%SRC%\gs\Resource" "%T%\gs\Resource" /e /njh /njs /nfl /ndl >nul
robocopy "%SRC%\gs\iccprofiles" "%T%\gs\iccprofiles" /e /njh /njs /nfl /ndl >nul
robocopy "%SRC%\tesseract" "%T%\tesseract" /e /njh /njs /nfl /ndl >nul
rem Vietnamese + English OCR data (if downloaded) so the other computer needs no download
if exist "%APPDATA%\PdfTool\tessdata\vie.traineddata" copy /y "%APPDATA%\PdfTool\tessdata\vie.traineddata" "%T%\tesseract\tessdata\" >nul
if exist "%APPDATA%\PdfTool\tessdata\eng.traineddata" copy /y "%APPDATA%\PdfTool\tessdata\eng.traineddata" "%T%\tesseract\tessdata\" >nul
for %%L in (ghostscript qpdf tesseract) do copy /y "%SRC%\licenses\%%L.txt" "%T%\licenses\" >nul

> "%OUT%\HUONG DAN.txt" (
echo PdfTool - Bo cong cu PDF
echo ========================
echo.
echo 1. Giai nen ca thu muc PdfTool vao noi bat ky, vi du C:\PdfTool hoac Desktop.
echo    Giu nguyen thu muc "tools" canh PdfTool.exe.
echo 2. Chay PdfTool.exe. Khong can cai dat.
echo    Neu Windows hien "Windows protected your PC": bam "More info" roi "Run anyway".
echo 3. Yeu cau: Windows 10 hoac 11 ^(da co san .NET Framework 4^).
echo.
echo Cac cong cu ma nguon mo di kem ^(thu muc tools^): qpdf ^(Apache 2.0^),
echo Tesseract OCR ^(Apache 2.0^), Ghostscript ^(AGPL 3.0^). Giay phep xem tools\licenses.
echo Ma nguon Ghostscript: https://ghostscript.com/releases/
)

if exist dist\PdfTool.zip del dist\PdfTool.zip
powershell -NoProfile -Command "Compress-Archive -Path 'dist\PdfTool' -DestinationPath 'dist\PdfTool.zip'"
if errorlevel 1 (echo NEN ZIP THAT BAI & pause & exit /b 1)
echo.
echo Da tao: %~dp0dist\PdfTool.zip

@echo off
title MarkItDown .exe olusturucu
cd /d "%~dp0"

set "PY="
python --version >nul 2>nul && set "PY=python"
if not defined PY py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY goto python_yok

echo.
echo  MarkItDown.exe olusturuluyor. Bu islem 5-10 dakika surebilir.
echo  Pencereyi kapatmayin.
echo.

if not exist ".derleme\Scripts\python.exe" %PY% -m venv .derleme
if errorlevel 1 goto hata

".derleme\Scripts\python.exe" -m pip install --upgrade pip
".derleme\Scripts\python.exe" -m pip install "markitdown[pdf,docx,pptx,xlsx,xls,outlook]" tkinterdnd2 pyinstaller
if errorlevel 1 goto hata

".derleme\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name MarkItDown ^
  --collect-all markitdown ^
  --collect-all magika ^
  --collect-all tkinterdnd2 ^
  --collect-all pdfminer ^
  --collect-all pdfplumber ^
  --collect-binaries onnxruntime ^
  MarkItDown_Uygulama.py
if errorlevel 1 goto hata

copy /y "dist\MarkItDown.exe" "MarkItDown.exe" >nul

rmdir /s /q build >nul 2>nul
rmdir /s /q dist >nul 2>nul
del /q MarkItDown.spec >nul 2>nul

echo.
echo  ============================================
echo   Tamamlandi: MarkItDown.exe hazir.
echo   Bu dosyayi istediginiz yere kopyalayabilir,
echo   baska bilgisayarlarda da kullanabilirsiniz.
echo  ============================================
echo.
explorer /select,"%~dp0MarkItDown.exe"
pause
exit /b

:python_yok
echo.
echo  Python bulunamadi (sadece .exe'yi olusturmak icin bir kez gerekli).
echo  Acilan sayfadan Python'u indirip kurun.
echo  Kurarken "Add Python to PATH" kutusunu MUTLAKA isaretleyin.
echo  Sonra bu dosyayi tekrar calistirin.
echo.
start https://www.python.org/downloads/
pause
exit /b

:hata
echo.
echo  Bir hata olustu. Yukaridaki kirmizi/son satirlarin ekran goruntusunu alin.
echo  Python surumunuzun 3.10 - 3.13 arasi olmasi en sorunsuzudur.
echo.
pause

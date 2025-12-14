@echo off
echo ================================================
echo Konsolidasyon Guncelleyici - EXE Olusturucu
echo ================================================
echo.

:: Virtual environment olustur
echo [1/4] Virtual environment olusturuluyor...
python -m venv venv
if errorlevel 1 (
    echo HATA: Python bulunamadi! Python 3.12+ yuklu olduguna emin olun.
    pause
    exit /b 1
)

:: Virtual environment'i aktifle
echo [2/4] Virtual environment aktiflestirildi
call venv\Scripts\activate.bat

:: Bağımlılıkları yükle
echo [3/4] Gerekli kutuphaneler yukleniyor...
python -m pip install --upgrade pip
pip install pyinstaller openpyxl

:: EXE oluştur
echo [4/4] EXE olusturuluyor...
pyinstaller --clean --noconfirm ^
    --onefile ^
    --windowed ^
    --name "KonsolidasyonGuncelleyici" ^
    --add-data "update_konsolidasyon.py;." ^
    desktop_gui.py

echo.
echo ================================================
echo TAMAMLANDI!
echo ================================================
echo.
echo EXE dosyasi: dist\KonsolidasyonGuncelleyici.exe
echo.
echo Bu exe dosyasini istediginiz Windows bilgisayara kopyalayabilirsiniz.
echo data.xlsx ve Konsolidasyon dosyalarini exe ile ayni klasore koyun.
echo.
pause

# Windows'ta EXE Oluşturma Kılavuzu

## Gereksinimler
- Windows 10 veya 11 (64-bit)
- Python 3.12 veya üzeri ([python.org](https://www.python.org/downloads/)'dan indirin)

## Hızlı Kurulum (3 Adım)

### 1. Python'un Yüklü Olduğunu Kontrol Edin

PowerShell veya CMD'de şunu çalıştırın:
```cmd
python --version
```

Python 3.12 veya üzeri görmeli görmelisiniz. Yoksa [python.org](https://www.python.org/downloads/)'dan indirin.

### 2. EXE Oluştur

Proje klasöründe `build_exe.bat` dosyasına çift tıklayın veya CMD'de:

```cmd
build_exe.bat
```

Bu otomatik olarak:
- ✅ Virtual environment oluşturur
- ✅ Gerekli kütüphaneleri yükler (pyinstaller, openpyxl)
- ✅ Tek dosya exe oluşturur

### 3. EXE'yi Kullanın

Oluşan `dist\KonsolidasyonGuncelleyici.exe` dosyası hazır!

## Kullanım

1. `KonsolidasyonGuncelleyici.exe` dosyasını istediğiniz yere kopyalayın
2. Aynı klasöre `data.xlsx` ve `Konsolidasyon_2025_NV (1).xlsx` dosyalarını koyun
3. EXE'ye çift tıklayın
4. Açılan pencereden dosyaları seçin ve "Güncellemeyi Başlat" butonuna tıklayın

## Manuel Kurulum (İsterseniz)

Eğer bat dosyası çalışmazsa manuel olarak:

```cmd
# 1. Virtual environment
python -m venv venv
venv\Scripts\activate

# 2. Bağımlılıkları yükle
pip install pyinstaller openpyxl

# 3. EXE oluştur
pyinstaller --onefile --windowed --name KonsolidasyonGuncelleyici --add-data "update_konsolidasyon.py;." desktop_gui.py
```

EXE şurada oluşacak: `dist\KonsolidasyonGuncelleyici.exe`

## Sorun Giderme

### "python bulunamadı" hatası
- Python'u [python.org](https://www.python.org/downloads/)'dan indirin
- Kurulum sırasında "Add Python to PATH" seçeneğini işaretleyin

### "This app can't run on your PC" hatası
- EXE'yi oluşturduğunuz Windows sürümü ile çalıştırdığınız sürüm uyumlu olmalı
- 64-bit Windows'ta 64-bit Python ile derleyin

### EXE açılmıyor
- `--windowed` yerine `--console` kullanarak yeniden derleyin (hataları görmek için):
  ```cmd
  pyinstaller --onefile --console --name KonsolidasyonGuncelleyici --add-data "update_konsolidasyon.py;." desktop_gui.py
  ```

## Notlar

- ✅ EXE tamamen bağımsız çalışır, Python yüklü olmasına gerek yoktur
- ✅ İnternet bağlantısına gerek yoktur
- ✅ Kullanıcı arayüzü Türkçe'dir
- ✅ Excel dosyalarını otomatik olarak bulur
- ⚠️ Windows Defender ilk çalıştırmada uyarı verebilir (normal bir durumdur)

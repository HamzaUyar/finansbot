# Docker ile Mac'te Windows EXE derleme
# UYARI: Tkinter ile sorun çıkabilir, GitHub Actions önerilir

FROM ubuntu:22.04

# Windows cross-compile araçları
RUN apt-get update && apt-get install -y \
    python3.12 \
    python3-pip \
    python3-tk \
    wine64 \
    wget \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python bağımlılıkları
COPY desktop_gui.py update_konsolidasyon.py ./
RUN pip3 install --no-cache-dir pyinstaller openpyxl

# EXE oluştur
CMD ["pyinstaller", \
     "--clean", \
     "--onefile", \
     "--windowed", \
     "--name", "KonsolidasyonGuncelleyici", \
     "--add-data", "update_konsolidasyon.py:.", \
     "desktop_gui.py"]

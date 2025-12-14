#!/bin/bash
# Mac'te Docker ile Windows EXE derleme scripti
# UYARI: GitHub Actions kullanmak daha kolay ve güvenilirdir!

echo "================================================"
echo "Docker ile Windows EXE Derleniyor..."
echo "================================================"
echo ""

# Docker kontrolü
if ! command -v docker &> /dev/null; then
    echo "❌ Docker bulunamadı!"
    echo "Lütfen Docker Desktop'ı yükleyin: brew install --cask docker"
    exit 1
fi

# Docker image oluştur
echo "[1/3] Docker image oluşturuluyor..."
docker build -t konsolidasyon-windows-builder . || {
    echo "❌ Docker image oluşturulamadı!"
    exit 1
}

# Dist klasörünü temizle
echo "[2/3] Önceki derleme temizleniyor..."
rm -rf dist build
mkdir -p dist

# EXE derle
echo "[3/3] EXE derleniyor..."
docker run --rm \
    -v "$(pwd):/app" \
    -v "$(pwd)/dist:/app/dist" \
    konsolidasyon-windows-builder || {
    echo "❌ Derleme başarısız!"
    exit 1
}

echo ""
echo "================================================"
echo "✅ TAMAMLANDI!"
echo "================================================"
echo ""
echo "EXE dosyası: dist/KonsolidasyonGuncelleyici.exe"
echo ""
echo "⚠️  NOT: Tkinter kullanıldığı için Windows'ta test etmelisiniz!"
echo "    Sorun çıkarsa GitHub Actions kullanın (daha güvenilir)."
echo ""

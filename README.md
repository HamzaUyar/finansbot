# Konsolidasyon Raporu Güncelleme Aracı

Bu program, `data.xlsx` dosyasındaki aylık şirket verilerini alarak Konsolidasyon raporunu otomatik olarak günceller.

## Özellikler

- ✅ **5 şirketin** %40'lık verilerini otomatik olarak %100'e dönüştürür (x2.5 çarpımı)
- ✅ 8 şirketin tüm aylık verilerini günceller
- ✅ "Gerçekleşen Aylık Euro" sayfasını günceller
- ✅ **"Finansal Raporlama AY" sayfasındaki formülleri günceller** (son aya göre)
- ✅ Tüm ayları günceller (sadece son ay değil)
- ✅ Son veri olan ayı otomatik tespit eder

## Güncellenen Şirketler

### %100'lük Veri Gönderen Şirketler (x1.0 - Dokunulmaz)
1. İnci Holding
2. ISM
3. İncitaş

### %40'lık Veri Gönderen Şirketler (x2.5 - Dönüştürülür)
4. İnci GS Yuasa
5. Yusen İnci Lojistik
6. Maxion Jantas
7. Maxion Celik
8. Maxion Aluminyum

## Kurulum

Bu proje UV paket yöneticisi kullanır. Bağımlılıklar otomatik olarak yüklenecektir.

```bash
# Gerekli bağımlılıklar
uv sync
```

## Kullanım

Ana programı çalıştırmak için:

```bash
uv run update_konsolidasyon.py
```

Farklı dosyalarla çalışmak isterseniz argümanları kullanabilirsiniz:

```bash
uv run update_konsolidasyon.py \
  --data /dosyalarim/data_yeni.xlsx \
  --konsolidasyon /dosyalarim/konsolidasyon_sablon.xlsx \
  --output /dosyalarim/konsolidasyon_guncel.xlsx
```

Program şunları yapar:

1. `data.xlsx` dosyasını okur
2. Son veri olan ayı tespit eder (örn: Ekim)
3. Tüm ayların "Gerçekleşen" verilerini alır
4. %40'lık verileri x2.5 ile çarparak %100'e dönüştürür
5. `Konsolidasyon_2025_NV (1)_guncel.xlsx` dosyasındaki "gerç aylık-eur" sayfasını günceller
6. **"Finansal Raporlama AY" sayfasındaki formülleri son aya göre günceller**
   - Formüllerdeki ay sütun referanslarını otomatik değiştirir
   - Örneğin: Eylül (Z, AA) → Ekim (AC, AD)
7. Değişiklikleri kaydeder

## Basit Arayüz (Streamlit)

Komut satırı kullanmak istemeyen ekip arkadaşlarınız için basit bir web arayüzü de bulunuyor. Streamlit uygulamasını çalıştırmak için:

```bash
uv run streamlit run app/streamlit_app.py
```

Arayüz üzerinden:
- Güncel `data.xlsx` dosyasını ve konsolidasyon şablonunu yükleyin.
- **Raporu Güncelle** butonuna tıklayarak işlem başlatılır.
- İşlem tamamlanınca güncellenmiş Excel dosyasını indirme butonu çıkar.

## Dosya Yapısı

```
konsolidasyon/
├── data.xlsx                              # Kaynak veri dosyası (%40'lık veriler)
├── Konsolidasyon_2025_NV (1)_guncel.xlsx  # Hedef konsolidasyon raporu
├── update_konsolidasyon.py                # Ana güncelleme programı
├── pyproject.toml                         # Proje yapılandırması
└── README.md                              # Bu dosya
```

## Önemli Notlar

⚠️ **Yedekleme**: Program çalıştırılmadan önce Konsolidasyon dosyasının yedeğini almanız önerilir.

⚠️ **Veri Formatı**: 
- `data.xlsx` dosyasında:
  - İlk 3 şirket (İnci Holding, ISM, İncitaş) → %100'lük veri (x1.0 - dokunulmaz)
  - Son 5 şirket (İnci GS Yuasa, Yusen, 3 Maxion) → %40'lık veri (x2.5 çarpımı yapılır)
- Bütçe verileri göz ardı edilir, sadece "Gerçekleşen" sütunları kullanılır

⚠️ **Finansal Raporlama AY Sayfası**: 
- Bu sayfadaki formüller **otomatik olarak son aya göre güncellenir**
- Program "A -Döviz" sayfasındaki sütun referanslarını değiştirir
- Manuel değişiklik gerekmez!

## Çıktı Örneği

```
================================================================================
KONSOLİDASYON RAPORU GÜNCELLEME PROGRAMI
================================================================================

Son veri olan ay: Eylül (Sütun 21)
Toplam 9 ay için veri okundu.

================================================================================
GERÇEKLEŞEN AYLIK EURO SAYFASI GÜNCELLENİYOR...
================================================================================

Ekim ayı güncelleniyor (Sütun K)...
  Satır 5 (İnci Holding): 0.00 -> 754,597.33
  Satır 6 (ISM): 0.00 -> 881,940.78
  ...

================================================================================
FİNANSAL RAPORLAMA AY SAYFASI GÜNCELLENİYOR...
================================================================================

Son veri olan ay: Ekim
Hedef sütunlar: AC (Budget) ve AD (Actual)
Toplam 64 formül güncellendi.

✓ GÜNCELLEME BAŞARIYLA TAMAMLANDI!
```

## Geliştirici

Program UV kullanılarak geliştirilmiştir ve Python 3.12+ gerektirir.

## Lisans

Bu proje özel kullanım içindir.

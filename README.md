# Konsolidasyon Raporu Güncelleme Aracı

Bu program, `data.xlsx` dosyasındaki aylık şirket verilerini alarak Konsolidasyon raporunu otomatik olarak günceller.

## Özellikler

- ✅ **17 Kontrol Order** (finansal metrik) desteği
- ✅ **5 şirketin** %40'lık verilerini otomatik olarak %100'e dönüştürür (x2.5 çarpımı)
- ✅ 8 şirketin tüm aylık verilerini günceller
- ✅ "Gerçekleşen Aylık Euro" sayfasını günceller (tüm metrikler için)
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

## Güncellenen 17 Kontrol Order (Finansal Metrikler)

Program aşağıdaki 17 kontrol order'ı data.xlsx'ten okur ve konsolidasyon raporuna aktarır:

1. **Net Satışlar** - Şirket ciroları
2. **Satılan Malların Maliyeti (SMM)** - Direkt maliyetler
3. **Brüt Kar** - Net satışlar - SMM
4. **Faaliyet Giderleri** - Operasyonel giderler
5. **Diğer Faaliyet Giderleri** - Yan faaliyet giderleri
6. **EBIT** - Faiz ve vergi öncesi kar
7. **Amortisman** - Varlık amortismanları
8. **EBITDA** - EBIT + Amortisman
9. **Finansman Gelir-Gideri** - Finansal gelir/giderler
10. **Diğer Gelir-Gideri** - Olağandışı gelir/giderler
11. **Vergi Öncesi Kar** - Vergi öncesi toplam kar
12. **Vergi** - Kurumlar vergisi
13. **Net Kar** - Nihai kar/zarar
14. **Krediler** - Toplam borç
15. **Nakit ve Bankalar** - Likit varlıklar
16. **Net Finansal Borç** - Net borç pozisyonu
17. **Yatırımlar** - CAPEX harcamaları

⚠️ **Not**: Kontrol Order 18 (Operasyonel Nakit) ve 19 (Serbest Nakit) konsolidasyon dosyasında bulunmadığından atlanır.

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
3. **17 kontrol order** için tüm ayların "Gerçekleşen" verilerini alır
4. %40'lık verileri x2.5 ile çarparak %100'e dönüştürür (5 şirket için)
5. `Konsolidasyon_2025_NV (1)_guncel.xlsx` dosyasındaki "gerç aylık-eur" sayfasını günceller
   - Toplam **~1,200 hücre** güncellenir (17 metrik × 8 şirket × ~10 ay)
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

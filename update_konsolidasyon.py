"""
Konsolidasyon Raporu Güncelleme Programı

Bu program data.xlsx'teki verileri alarak Konsolidasyon raporunu günceller:
1. data.xlsx'teki %40'lık verileri x2.5 ile çarparak %100'e dönüştürür
2. Gerçekleşen Aylık Euro sayfasını günceller (tüm aylar için)
3. Finansal Raporlama AY sayfasındaki formülleri kontrol eder
"""

import openpyxl
from openpyxl import load_workbook
from datetime import datetime
import sys
from pathlib import Path

# Dosya yolları
DATA_FILE = '/Users/hamzauyar/projects/konsolidasyon/data.xlsx'
KONSOLIDASYON_FILE = '/Users/hamzauyar/projects/konsolidasyon/Konsolidasyon_2025_NV (1)_guncel.xlsx'

# Şirket eşleştirmeleri
# data.xlsx satır numaraları -> (Konsolidasyon satır no, şirket adı, %40'lık mı?)
# NOT: İlk 3 şirket zaten %100'lük veri gönderiyor (dokunma!)
#      Son 5 şirket %40'lık veri gönderiyor (x2.5 yap!)
COMPANY_MAPPING = {
    5: (5, "İnci Holding", False),          # %100'lük - DOKUNMA
    6: (6, "ISM", False),                   # %100'lük - DOKUNMA
    7: (7, "İncitaş", False),               # %100'lük - DOKUNMA
    8: (8, "İnci GS Yuasa", True),          # %40'lık - x2.5 YAP
    9: (9, "Yusen İnci Lojistik", True),    # %40'lık - x2.5 YAP
    10: (10, "Maxion Jantas", True),        # %40'lık - x2.5 YAP
    11: (11, "Maxion Celik", True),         # %40'lık - x2.5 YAP
    12: (12, "Maxion Aluminyum", True),     # %40'lık - x2.5 YAP
}

# Ay eşleştirmeleri
# data.xlsx Gerçekleşen sütunları -> Konsolidasyon sütunları
MONTH_MAPPING = {
    5: (2, "Ocak"),      # data col 5 -> konsolidasyon col B (2)
    7: (3, "Şubat"),     # data col 7 -> konsolidasyon col C (3)
    9: (4, "Mart"),      # data col 9 -> konsolidasyon col D (4)
    11: (5, "Nisan"),    # data col 11 -> konsolidasyon col E (5)
    13: (6, "Mayıs"),    # data col 13 -> konsolidasyon col F (6)
    15: (7, "Haziran"),  # data col 15 -> konsolidasyon col G (7)
    17: (8, "Temmuz"),   # data col 17 -> konsolidasyon col H (8)
    19: (9, "Ağustos"),  # data col 19 -> konsolidasyon col I (9)
    21: (10, "Eylül"),   # data col 21 -> konsolidasyon col J (10)
    23: (11, "Ekim"),    # data col 23 -> konsolidasyon col K (11)
    25: (12, "Kasım"),   # data col 25 -> konsolidasyon col L (12)
    27: (13, "Aralık"),  # data col 27 -> konsolidasyon col M (13)
}


def find_last_month_with_data(ws):
    """data.xlsx'te son veri olan ayı bulur"""
    last_month_col = None
    last_month_name = None
    
    for data_col, (kons_col, month_name) in MONTH_MAPPING.items():
        has_data = False
        for data_row in COMPANY_MAPPING.keys():
            value = ws.cell(data_row, data_col).value
            if value and value != 0:
                has_data = True
                break
        
        if has_data:
            last_month_col = data_col
            last_month_name = month_name
        else:
            break
    
    return last_month_col, last_month_name


def read_data_from_data_xlsx():
    """data.xlsx'ten tüm gerçekleşen verilerini okur ve %100'e dönüştürür"""
    print("=" * 80)
    print("DATA.XLSX DOSYASINDAN VERİLER OKUNUYOR...")
    print("=" * 80)
    
    wb = load_workbook(DATA_FILE, data_only=True)
    ws = wb['Export']
    
    # Son veri olan ayı bul
    last_month_col, last_month_name = find_last_month_with_data(ws)
    print(f"\nSon veri olan ay: {last_month_name} (Sütun {last_month_col})")
    
    # Tüm ayların verilerini oku
    all_data = {}
    
    for data_col, (kons_col, month_name) in MONTH_MAPPING.items():
        month_data = {}
        
        for data_row, (kons_row, company_name, is_40_percent) in COMPANY_MAPPING.items():
            value = ws.cell(data_row, data_col).value
            
            # %40'lık veriyi %100'e dönüştür (sadece ilk 5 şirket için)
            if value and value != 0:
                if is_40_percent:
                    converted_value = value * 2.5  # %40'lık -> %100
                else:
                    converted_value = value  # Zaten %100
                month_data[kons_row] = converted_value
            else:
                month_data[kons_row] = None
        
        # Bu ayda veri varsa kaydet
        if any(v is not None for v in month_data.values()):
            all_data[kons_col] = {
                'month_name': month_name,
                'data': month_data
            }
    
    wb.close()
    
    print(f"\nToplam {len(all_data)} ay için veri okundu.")
    return all_data, last_month_name


def update_gercaylık_euro_sheet(all_data):
    """Gerçekleşen Aylık Euro sayfasını günceller"""
    print("\n" + "=" * 80)
    print("GERÇEKLEŞEN AYLIK EURO SAYFASI GÜNCELLENİYOR...")
    print("=" * 80)
    
    wb = load_workbook(KONSOLIDASYON_FILE)
    ws = wb['gerç aylık-eur']
    
    updates_count = 0
    
    # Her ay için verileri güncelle
    for kons_col, month_info in all_data.items():
        month_name = month_info['month_name']
        month_data = month_info['data']
        
        print(f"\n{month_name} ayı güncelleniyor (Sütun {openpyxl.utils.get_column_letter(kons_col)})...")
        
        for kons_row, value in month_data.items():
            if value is not None:
                old_value = ws.cell(kons_row, kons_col).value
                ws.cell(kons_row, kons_col).value = value
                updates_count += 1
                
                # Satırdaki şirket adını al (A sütunu)
                company_name = ws.cell(kons_row, 1).value
                if company_name:
                    company_name = str(company_name)[:25]
                
                if old_value != value:
                    print(f"  Satır {kons_row:3d} ({company_name:25s}): {old_value or 0:20,.2f} -> {value:20,.2f}")
    
    # Dosyayı kaydet
    print(f"\n{'='*80}")
    print(f"Toplam {updates_count} hücre güncellendi.")
    print("Dosya kaydediliyor...")
    wb.save(KONSOLIDASYON_FILE)
    wb.close()
    print("✓ Gerçekleşen Aylık Euro sayfası başarıyla güncellendi!")


def update_finansal_ay_formulas(last_month_name):
    """Finansal Raporlama AY sayfasındaki formülleri son aya göre günceller"""
    print("\n" + "=" * 80)
    print("FİNANSAL RAPORLAMA AY SAYFASI GÜNCELLENİYOR...")
    print("=" * 80)
    
    # Ay isimlerini A -Döviz sayfasındaki sütunlara eşle
    month_to_columns = {
        "Ocak": ("B", "C"),       # Budget, Actual
        "Şubat": ("E", "F"),
        "Mart": ("H", "I"),
        "Nisan": ("K", "L"),
        "Mayıs": ("N", "O"),
        "Haziran": ("Q", "R"),
        "Temmuz": ("T", "U"),
        "Ağustos": ("W", "X"),
        "Eylül": ("Z", "AA"),
        "Ekim": ("AC", "AD"),
        "Kasım": ("AF", "AG"),
        "Aralık": ("AI", "AJ"),
    }
    
    if last_month_name not in month_to_columns:
        print(f"❌ Uyarı: '{last_month_name}' ayı için sütun eşleştirmesi bulunamadı!")
        return
    
    target_budget_col, target_actual_col = month_to_columns[last_month_name]
    
    print(f"\nSon veri olan ay: {last_month_name}")
    print(f"Hedef sütunlar: {target_budget_col} (Budget) ve {target_actual_col} (Actual)")
    
    wb = load_workbook(KONSOLIDASYON_FILE)
    ws = wb['Finansal Raporlama AY']
    
    update_count = 0
    
    # Satır 6-14 arası (Toplam + şirketler), tüm sütunları güncelle
    # D, E, F = Ciro (Gerçekleşme, Fark, % Değişim)
    # G, H, I = Brüt Kar
    # J, K, L = EBITDA
    # M, N, O = Net Kar
    target_cols = [4, 5, 7, 8, 10, 11, 13, 14]  # D, E, G, H, J, K, M, N
    
    print(f"\nFormüller güncelleniyor...")
    
    for row_idx in range(6, 15):  # Satır 6-14 (Toplam + şirketler)
        company_name = ws.cell(row_idx, 2).value  # B sütunu - şirket adı
        
        for col_idx in target_cols:
            cell = ws.cell(row_idx, col_idx)
            
            if cell.data_type == 'f' and cell.value:
                old_formula = str(cell.value)
                
                # Formüldeki tüm ay sütun referanslarını bul ve değiştir
                new_formula = old_formula
                
                # Tüm aylardaki Budget ve Actual sütunlarını hedef aya güncelle
                for month, (budget_col, actual_col) in month_to_columns.items():
                    # Budget sütununu güncelle (örn: 'Z' -> 'AC')
                    new_formula = new_formula.replace(f"'{budget_col}", f"'__TEMP_BUDGET__")
                    new_formula = new_formula.replace(f"!{budget_col}", f"!__TEMP_BUDGET__")
                    
                    # Actual sütununu güncelle (örn: 'AA' -> 'AD')
                    new_formula = new_formula.replace(f"'{actual_col}", f"'__TEMP_ACTUAL__")
                    new_formula = new_formula.replace(f"!{actual_col}", f"!__TEMP_ACTUAL__")
                
                # Geçici işaretçileri hedef sütunlarla değiştir
                new_formula = new_formula.replace("__TEMP_BUDGET__", target_budget_col)
                new_formula = new_formula.replace("__TEMP_ACTUAL__", target_actual_col)
                
                if new_formula != old_formula:
                    cell.value = new_formula
                    update_count += 1
                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                    if update_count <= 5:  # İlk 5 güncellemeyi göster
                        print(f"  [{row_idx},{col_letter}] {company_name}: Güncellendi")
    
    if update_count > 0:
        print(f"\n{'='*80}")
        print(f"Toplam {update_count} formül güncellendi.")
        print("Dosya kaydediliyor...")
        wb.save(KONSOLIDASYON_FILE)
        print("✓ Finansal Raporlama AY sayfası başarıyla güncellendi!")
    else:
        print(f"\n{'='*80}")
        print("⚠ Hiçbir formül güncellenmedi (zaten güncel olabilir).")
    
    wb.close()


def main():
    """Ana program"""
    print("\n" + "=" * 80)
    print("KONSOLİDASYON RAPORU GÜNCELLEME PROGRAMI")
    print("=" * 80)
    print(f"\nData dosyası: {DATA_FILE}")
    print(f"Konsolidasyon dosyası: {KONSOLIDASYON_FILE}")
    
    # Dosyaların varlığını kontrol et
    if not Path(DATA_FILE).exists():
        print(f"\n❌ HATA: {DATA_FILE} bulunamadı!")
        sys.exit(1)
    
    if not Path(KONSOLIDASYON_FILE).exists():
        print(f"\n❌ HATA: {KONSOLIDASYON_FILE} bulunamadı!")
        sys.exit(1)
    
    try:
        # 1. data.xlsx'ten verileri oku
        all_data, last_month_name = read_data_from_data_xlsx()
        
        # 2. Gerçekleşen Aylık Euro sayfasını güncelle
        update_gercaylık_euro_sheet(all_data)
        
        # 3. Finansal Raporlama AY sayfasını güncelle
        update_finansal_ay_formulas(last_month_name)
        
        print("\n" + "=" * 80)
        print("✓ GÜNCELLEME BAŞARIYLA TAMAMLANDI!")
        print("=" * 80)
        print(f"\nGüncellenmiş dosya: {KONSOLIDASYON_FILE}")
        print(f"Son güncellenen ay: {last_month_name}")
        
    except Exception as e:
        print(f"\n❌ HATA: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()


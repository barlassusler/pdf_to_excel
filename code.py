import pdfplumber
import re
import pandas as pd
import os

# PDF dosyalarının bulunduğu klasör yolu
pdf_directory = 'C:/Users/kemal.susler/Desktop/faturalar/'
excel_path = 'C:/Users/kemal.susler/Desktop/faturalar.xlsx'


def extract_field(text, keyword, next_words=10):
    # Anahtar kelimeyi bulma ve ardından gelen değeri çıkarma fonksiyonu
    pattern = rf'{keyword}\s*[:\s]*([\d\w\.\,\/-]+)'
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None


# gereksiz karakter temizleme
def sanitize_sheet_name(sheet_name):
    invalid_chars = ['\\', '/', '*', '[', ']', ':', '?']
    for char in invalid_chars:
        sheet_name = sheet_name.replace(char, '')
    return sheet_name


# .pdf uzantılı dosyaları görme
pdf_files = [f for f in os.listdir(pdf_directory) if f.endswith('.pdf')]

# Her bir PDF dosyasını işleme
for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_directory, pdf_file)

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text()

    # kelimelerle ilgili alanları çıkarma
    fatura_no = extract_field(full_text, "Fatura No")
    fatura_tarihi = extract_field(full_text, "Fatura Tarihi")
    mal_hizmet_tutari = extract_field(full_text, "Mal / Hizmet Tutarı")
    if not mal_hizmet_tutari:
        mal_hizmet_tutari = extract_field(full_text, "Mal/Hizmet Tutarı")
    if not mal_hizmet_tutari:
        mal_hizmet_tutari = extract_field(full_text, "Hizmet Tutarı")
    if not mal_hizmet_tutari:
        mal_hizmet_tutari = extract_field(full_text, "Mal Hizmet Tutarı")
    if not mal_hizmet_tutari:
        mal_hizmet_tutari = extract_field(full_text, "Mal Hizmet Tutari")
    if not mal_hizmet_tutari:
        mal_hizmet_tutari = extract_field(full_text, "Mal Hizmet Toplam Tutari")
    if not mal_hizmet_tutari:
        mal_hizmet_tutari = extract_field(full_text, "Mal Hizmet Toplam Tutarı")
    if not mal_hizmet_tutari:
        mal_hizmet_tutari = extract_field(full_text, "Mal / Hizmet Toplam Tutarı")

    ödenecek_tutar = extract_field(full_text, "Ödenecek Tutar")

    # Çıkarılan verileri yazdırma
    print(f"PDF Dosyası: {pdf_file}")
    print(f"Fatura No: {fatura_no}")
    print(f"Fatura Tarihi: {fatura_tarihi}")
    print(f"Mal/Hizmet Tutarı: {mal_hizmet_tutari}")
    print(f"Ödenecek Tutar: {ödenecek_tutar}")
    print("-----------------------------------")

    # Excel dosyasına yazma
    df = pd.DataFrame({
        'Tarih': [fatura_tarihi],
        'Fatura Numarası': [fatura_no],
        'Mal/Hizmet Tutarı': [mal_hizmet_tutari],
        'Ödenecek Tutar': [ödenecek_tutar]
    })

    safe_sheet_name = sanitize_sheet_name(f"Fatura_{fatura_no}")

    with pd.ExcelWriter(excel_path, mode='a', if_sheet_exists='new') as writer:
        df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

print("Tüm PDF dosyaları başarıyla işlendi ve Excel dosyasına aktarıldı.")
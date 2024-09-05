import pdfplumber
import re
import pandas as pd
import os

#the path which bills are in
pdf_directory = 'C:/Users/kemal.susler/Desktop/faturalar/'
excel_path = 'C:/Users/kemal.susler/Desktop/faturalar.xlsx'


def extract_field(text, keyword, next_words=10):
    # Function to find the keyword and extract the value that follows
    pattern = rf'{keyword}\s*[:\s]*([\d\w\.\,\/-]+)'
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None


# invalid character cleaning
def sanitize_sheet_name(sheet_name):
    invalid_chars = ['\\', '/', '*', '[', ']', ':', '?']
    for char in invalid_chars:
        sheet_name = sheet_name.replace(char, '')
    return sheet_name


# choose the files which are endswith .pdf
pdf_files = [f for f in os.listdir(pdf_directory) if f.endswith('.pdf')]

# proccess each pdf file
for pdf_file in pdf_files:
    pdf_path = os.path.join(pdf_directory, pdf_file)

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text()

    # extracting fields related to words
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

    # test print
    print(f"PDF Dosyası: {pdf_file}")
    print(f"Fatura No: {fatura_no}")
    print(f"Fatura Tarihi: {fatura_tarihi}")
    print(f"Mal/Hizmet Tutarı: {mal_hizmet_tutari}")
    print(f"Ödenecek Tutar: {ödenecek_tutar}")
    print("-----------------------------------")

    # writing to excel file
    df = pd.DataFrame({
        'Tarih': [fatura_tarihi],
        'Fatura Numarası': [fatura_no],
        'Mal/Hizmet Tutarı': [mal_hizmet_tutari],
        'Ödenecek Tutar': [ödenecek_tutar]
    })

    safe_sheet_name = sanitize_sheet_name(f"Fatura_{fatura_no}")

    with pd.ExcelWriter(excel_path, mode='a', if_sheet_exists='new') as writer:
        df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

print("All PDF files were successfully processed and exported to Excel file.")
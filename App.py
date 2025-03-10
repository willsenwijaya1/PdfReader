import streamlit as st
import pdfplumber
import pandas as pd
import os
from openpyxl import load_workbook

# Fungsi untuk ekstrak tabel dari PDF
def extract_tables_from_pdf(pdf_path):
    all_extracted_data = []
    pending_row = None
    
    def clean_text(text):
        if pd.isnull(text):
            return ''
        return str(text).strip().lower().replace('\n', ' ')

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                df = pd.DataFrame(table).applymap(clean_text)
                
                if pending_row is not None:
                    df.iloc[0] = df.iloc[0].combine_first(pending_row)
                    pending_row = None
                
                last_row = df.iloc[-1]
                if last_row.isnull().sum() > 0:
                    pending_row = last_row
                    df = df[:-1]
                
                all_extracted_data.append(df)
    return all_extracted_data

# Fungsi untuk memperbaiki data yang tidak sejajar
def perbaiki_nilai_tidak_sejajar(df):
    for i in range(len(df) - 1):
        if 'jumlah dianalisa' in df.iloc[i].values:
            header_row = i
            value_row = i + 1
            header_positions = {}
            nilai_tersedia = []
            
            for col in range(len(df.columns)):
                header = df.iloc[header_row, col]
                value = df.iloc[value_row, col]
                if pd.notna(header):
                    header_positions[col] = header
                if pd.notna(value):
                    nilai_tersedia.append(value)
            
            nilai_index = 0
            for col in header_positions:
                if nilai_index < len(nilai_tersedia):
                    df.iloc[value_row, col] = nilai_tersedia[nilai_index]
                    nilai_index += 1
                else:
                    df.iloc[value_row, col] = None
    return df

# Fungsi untuk memproses semua PDF dalam folder
def process_all_pdfs_in_folder(folder_path, output_excel):
    all_data = []
    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, file)
            extracted_data = extract_tables_from_pdf(pdf_path)
            if extracted_data:
                all_data.extend(extracted_data)
    
    if all_data:
        raw_df = pd.concat(all_data, ignore_index=True)
        raw_df.to_excel(output_excel, sheet_name='Sheet1', index=False, header=True)
        
        processed_df = perbaiki_nilai_tidak_sejajar(raw_df)
        with pd.ExcelWriter(output_excel, engine='openpyxl', mode='a') as writer:
            processed_df.to_excel(writer, sheet_name='Sheet2', index=False)
    
# Streamlit UI
st.title("📑 PDF Table Extractor")

uploaded_files = st.file_uploader("Upload file PDF", accept_multiple_files=True, type=['pdf'])
output_excel = "output_tabel.xlsx"

if st.button("Proses PDF") and uploaded_files:
    folder_path = "uploaded_pdfs"
    os.makedirs(folder_path, exist_ok=True)
    
    for uploaded_file in uploaded_files:
        with open(os.path.join(folder_path, uploaded_file.name), "wb") as f:
            f.write(uploaded_file.getbuffer())
    
    process_all_pdfs_in_folder(folder_path, output_excel)
    st.success(f"Proses selesai! Data bersih di Sheet2 dalam {output_excel}.")
    
    with open(output_excel, "rb") as f:
        st.download_button("Download hasil", f, file_name=output_excel)

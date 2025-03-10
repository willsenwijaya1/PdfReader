import streamlit as st
import pdfplumber
import pandas as pd
import os
from io import BytesIO
from openpyxl import load_workbook

def extract_tables_from_pdf(pdf_file):
    all_extracted_data = []
    pending_row = None
    previous_headers = None

    def clean_text(text):
        if pd.isnull(text):
            return ''
        return str(text).strip().lower().replace('\n', ' ')

    with pdfplumber.open(pdf_file) as pdf:
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

def perbaiki_nilai_tidak_sejajar(df):
    kasus_ditemukan = False
    
    for i in range(len(df) - 1):
        if 'jumlah dianalisa' in df.iloc[i].values:
            header_row = i
            value_row = i + 1
            
            if value_row < len(df):
                kasus_ditemukan = True
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

def process_uploaded_pdfs(uploaded_files):
    all_data = []
    for uploaded_file in uploaded_files:
        extracted_data = extract_tables_from_pdf(uploaded_file)
        if extracted_data:
            all_data.extend(extracted_data)
    
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        final_df = perbaiki_nilai_tidak_sejajar(final_df)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, sheet_name='Sheet1')
        output.seek(0)
        
        return output
    return None

st.title("Ekstraksi Tabel dari PDF")
st.write("Unggah file PDF untuk mengekstrak tabel dan mengonversinya ke Excel.")

uploaded_files = st.file_uploader("Unggah file PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    if st.button("Proses PDF"):
        output_excel = process_uploaded_pdfs(uploaded_files)
        if output_excel:
            st.success("Ekstraksi selesai! Unduh hasilnya di bawah ini.")
            st.download_button("Unduh Excel", output_excel, file_name="output_tabel.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.warning("Tidak ada tabel yang ditemukan dalam file PDF.")

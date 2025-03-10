import streamlit as st
import pdfplumber
import pandas as pd
import os
from io import BytesIO
from openpyxl import load_workbook

def extract_tables_from_pdf(pdf_file):
    all_extracted_data = []
    pending_row = None

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

def process_pdf(file):
    extracted_data = extract_tables_from_pdf(file)
    if extracted_data:
        df_raw = pd.concat(extracted_data, ignore_index=True)
        df_clean = perbaiki_nilai_tidak_sejajar(df_raw.copy())
        return df_raw, df_clean
    return None, None

def save_to_excel(df_raw, df_clean):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_raw.to_excel(writer, sheet_name='Sheet1', index=False)
        df_clean.to_excel(writer, sheet_name='Sheet2', index=False)
    output.seek(0)
    return output

st.title("Ekstraksi dan Pemrosesan PDF")

uploaded_file = st.file_uploader("Upload file PDF", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Memproses PDF..."):
        df_raw, df_clean = process_pdf(uploaded_file)
    
    if df_raw is not None:
        st.subheader("Data Mentah (Sheet1)")
        st.dataframe(df_raw)
        
        st.subheader("Data Bersih (Sheet2)")
        st.dataframe(df_clean)
        
        excel_data = save_to_excel(df_raw, df_clean)
        st.download_button(
            label="Unduh Hasil dalam Excel",
            data=excel_data,
            file_name="output_tabel.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Tidak ada tabel yang ditemukan dalam PDF.")

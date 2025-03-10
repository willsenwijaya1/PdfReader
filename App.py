import streamlit as st
import pdfplumber
import pandas as pd
import os
from openpyxl import load_workbook

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

def perbaiki_nilai_tidak_sejajar(df):
    for i in range(len(df) - 1):
        if 'jumlah dianalisa' in df.iloc[i].values:
            header_row, value_row = i, i + 1
            if value_row < len(df):
                header_positions, nilai_tersedia = {}, []
                for col in range(len(df.columns)):
                    header, value = df.iloc[header_row, col], df.iloc[value_row, col]
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

def process_uploaded_pdf(uploaded_file):
    extracted_data = extract_tables_from_pdf(uploaded_file)
    if extracted_data:
        final_df = pd.concat(extracted_data, ignore_index=True)
        final_df = perbaiki_nilai_tidak_sejajar(final_df)
        return final_df
    return None

st.title("PDF Table Extractor & Formatter")

uploaded_file = st.file_uploader("Upload PDF File", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Processing..."):
        df = process_uploaded_pdf(uploaded_file)
        if df is not None:
            st.success("Processing Complete!")
            st.dataframe(df)
            
            output_excel = "processed_output.xlsx"
            df.to_excel(output_excel, index=False)
            
            with open(output_excel, "rb") as f:
                st.download_button("Download Processed Excel", f, file_name=output_excel)
        else:
            st.warning("No tables found in the uploaded PDF.")

import streamlit as st
import pdfplumber
import pandas as pd
import os
from openpyxl import load_workbook
from io import BytesIO

def extract_tables_from_pdf(pdf_file):
    all_extracted_data = []
    pending_row = None
    previous_headers = None

    def clean_text(text):
        if pd.isnull(text):
            return ''
        return str(text).strip().lower().replace('\n', ' ')

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
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

def process_pdfs(uploaded_files):
    all_data = []
    for uploaded_file in uploaded_files:
        extracted_data = extract_tables_from_pdf(uploaded_file)
        if extracted_data:
            all_data.extend(extracted_data)
    
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, header=True)
        output.seek(0)
        return output
    return None

def parse_excel(file_bytes):
    df = pd.read_excel(file_bytes, sheet_name='Sheet1', header=None)
    df.fillna(method='ffill', inplace=True)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet2', index=False)
    output.seek(0)
    return output

st.title("PDF to Excel Table Extractor")
uploaded_files = st.file_uploader("Upload PDF Files", type=["pdf"], accept_multiple_files=True)
if uploaded_files:
    processed_excel = process_pdfs(uploaded_files)
    if processed_excel:
        parsed_excel = parse_excel(processed_excel)
        st.download_button("Download Processed Excel", parsed_excel, file_name="processed_output.xlsx")
        st.success("Processing Complete! Output saved as Sheet2.")
    else:
        st.warning("No tables found in the uploaded PDFs.")

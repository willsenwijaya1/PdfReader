import streamlit as st
import pdfplumber
import pandas as pd
import os
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
                if df.iloc[0].str.contains("pecahan", case=False, na=False).any():
                    previous_headers = df.iloc[0].tolist()
                    df = df[1:].reset_index(drop=True)
                elif previous_headers is not None:
                    df.columns = previous_headers
                if pending_row is not None:
                    df.iloc[0] = df.iloc[0].combine_first(pending_row)
                    pending_row = None
                last_row = df.iloc[-1]
                if last_row.isnull().sum() > 0:
                    pending_row = last_row
                    df = df[:-1]
                all_extracted_data.append(df)
    return all_extracted_data

st.title("PDF to Excel Table Extractor")
uploaded_file = st.file_uploader("Upload PDF file", type=["pdf"])

if uploaded_file is not None:
    extracted_data = extract_tables_from_pdf(uploaded_file)
    if extracted_data:
        final_df = pd.concat(extracted_data, ignore_index=True)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            final_df.to_excel(writer, index=False, header=False)
        output.seek(0)
        st.download_button(label="Download Excel File", data=output, file_name="extracted_tables.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("No tables found in the PDF.")

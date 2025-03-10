import streamlit as st
import pdfplumber
import pandas as pd
import os
from openpyxl import load_workbook

st.title("📑 PDF Table Extractor & Cleaner")

uploaded_files = st.file_uploader("Upload PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    output_excel = "output_tabel.xlsx"
    all_extracted_data = []

    st.write("🔄 **Memproses file...**")

    def extract_tables_from_pdf(pdf_path):
        all_extracted_data = []
        pending_row = None

        def clean_text(text):
            return str(text).strip().lower().replace('\n', ' ') if pd.notna(text) else ''

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

    for file in uploaded_files:
        with st.spinner(f"Memproses: {file.name}"):
            df_list = extract_tables_from_pdf(file)
            if df_list:
                all_extracted_data.extend(df_list)

    if all_extracted_data:
        final_df = pd.concat(all_extracted_data, ignore_index=True)
        final_df.to_excel(output_excel, index=False, header=True)
        st.success(f"✅ Semua tabel berhasil diekspor ke {output_excel}")
        with open(output_excel, "rb") as f:
            st.download_button("⬇ Download Excel", f, file_name=output_excel, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("⚠ Tidak ada tabel yang ditemukan dalam file PDF.")

    st.write("✅ Proses selesai!")

if __name__ == "__main__":
    st.write("Silakan upload file PDF untuk diekstrak.")

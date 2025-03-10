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

def process_pdfs(uploaded_files):
    all_data = []
    for uploaded_file in uploaded_files:
        extracted_data = extract_tables_from_pdf(uploaded_file)
        if extracted_data:
            all_data.extend(extracted_data)
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        return final_df
    return None

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

def proses_data(df):
    cols_needed = ["tanggal temuan", "cara temuan", "waktu pendeteksian", "nama kantor", "provinsi", "kota", "jenis kontributor", "kantor kontributor", "nama kontributor", "dokumen pendukung", "no. identitas", "keterangan", "provinsi", "kota", "kecamatan", "pecahan", "tahun emisi", "no. seri 1", "no. seri 2", "jumlah lembar", "jumlah lembar terima", "subtotal", "jumlah dianalisa", "hasil analisa", "subtotal"]
    data_rows, current_row, found_first_subtotal = [], {}, False
    for idx, row in df.iterrows():
        for col_num, cell_value in enumerate(row):
            if cell_value in cols_needed:
                if idx+1 < len(df):
                    next_val = df.iloc[idx+1, col_num]
                    if cell_value == 'subtotal':
                        if found_first_subtotal:
                            current_row['subtotal 2'] = next_val
                            data_rows.append(current_row)
                            current_row, found_first_subtotal = {}, False
                        else:
                            current_row['subtotal'] = next_val
                            found_first_subtotal = True
                    elif cell_value == 'provinsi' and 'provinsi' in current_row:
                        current_row['provinsi_kontributor'] = next_val
                    elif cell_value == 'kota' and 'kota' in current_row:
                        current_row['kota_kontributor'] = next_val
                    else:
                        current_row[cell_value] = next_val
    if current_row:
        data_rows.append(current_row)
    df_output = pd.DataFrame(data_rows)
    if 'tanggal temuan' in df_output.columns:
        df_output['tanggal temuan'] = pd.to_datetime(df_output['tanggal temuan'], errors='coerce')
        mask_no_date = df_output['tanggal temuan'].isna()
        df_output.loc[mask_no_date, 'provinsi_kontributor'] = df_output.loc[mask_no_date, 'provinsi']
        df_output.loc[mask_no_date, 'kota_kontributor'] = df_output.loc[mask_no_date, 'kota']
        df_output.loc[mask_no_date, ['provinsi', 'kota']] = None
    df_output.fillna(method='ffill', inplace=True)
    return df_output

st.title("Ekstraksi dan Analisis Data dari PDF")

uploaded_files = st.file_uploader("Upload PDF files", type=["pdf"], accept_multiple_files=True)
if uploaded_files:
    df = process_pdfs(uploaded_files)
    if df is not None:
        st.write("Data yang diekstrak:", df)
        df_diperbaiki = perbaiki_nilai_tidak_sejajar(df)
        df_proses = proses_data(df_diperbaiki)
        st.write("Data setelah perbaikan:", df_proses)
        output_excel = "output_tabel.xlsx"
        with pd.ExcelWriter(output_excel, engine='openpyxl', mode='w') as writer:
            df_proses.to_excel(writer, sheet_name='Sheet1', index=False)
        with open(output_excel, "rb") as f:
            st.download_button("Download hasil Excel", f, file_name=output_excel)

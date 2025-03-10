import streamlit as st
import pdfplumber
import pandas as pd
import os
from openpyxl import load_workbook

def extract_tables_from_pdf(pdf_path):
    all_extracted_data = []
    pending_row = None
    previous_headers = None

    def clean_text(text):
        if pd.isnull(text):
            return ''
        return str(text).strip().lower().replace('\n', ' ')

    with pdfplumber.open(pdf_path) as pdf:
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

def process_all_pdfs_in_folder(folder_path, output_excel):
    all_data = []
    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            pdf_path = os.path.join(folder_path, file)
            extracted_data = extract_tables_from_pdf(pdf_path)
            if extracted_data:
                all_data.extend(extracted_data)
    
    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        final_df.to_excel(output_excel, index=False, header=True)
        return output_excel
    return None

def perbaiki_nilai_tidak_sejajar(df):
    for i in range(len(df) - 1):
        if 'jumlah dianalisa' in df.iloc[i].values:
            header_row = i
            value_row = i + 1
            if value_row < len(df):
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

def parse_excel(file_path):
    sheet1 = pd.read_excel(file_path, sheet_name='Sheet1', header=None)
    data_rows = []
    current_row = {}
    cols_needed = [
        "tanggal temuan", "cara temuan", "waktu pendeteksian", "nama kantor", "provinsi", "kota",
        "jenis kontributor", "kantor kontributor", "nama kontributor", "dokumen pendukung", 
        "no. identitas", "keterangan", "provinsi", "kota", "kecamatan", "pecahan", "tahun emisi",
        "no. seri 1", "no. seri 2", "jumlah lembar", "jumlah lembar terima", "subtotal",
        "jumlah dianalisa", "hasil analisa", "subtotal"
    ]
    found_first_subtotal = False
    for idx, row in sheet1.iterrows():
        for col_num, cell_value in enumerate(row):
            if cell_value in cols_needed:
                if idx+1 < len(sheet1):
                    next_val = sheet1.iloc[idx+1, col_num]
                    if cell_value == 'subtotal':
                        if found_first_subtotal:
                            current_row['subtotal 2'] = next_val
                            data_rows.append(current_row)
                            current_row = {}
                            found_first_subtotal = False
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
    df = pd.DataFrame(data_rows)
    if 'tanggal temuan' in df.columns:
        df['tanggal temuan'] = pd.to_datetime(df['tanggal temuan'], errors='coerce')
        mask_no_date = df['tanggal temuan'].isna()
        df.loc[mask_no_date, 'provinsi_kontributor'] = df.loc[mask_no_date, 'provinsi']
        df.loc[mask_no_date, 'kota_kontributor'] = df.loc[mask_no_date, 'kota']
        df.loc[mask_no_date, ['provinsi', 'kota']] = None
    date_columns = ["tanggal temuan"]
    for col in date_columns:
        if col in df.columns:
            df[col] = df[col].dt.strftime('%d-%m-%Y')
    df.fillna(method='ffill', inplace=True)
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, sheet_name='Sheet2', index=False)
    return file_path

st.title("PDF to Excel Table Extractor")
uploaded_file = st.file_uploader("Upload a PDF File", type=["pdf"])
if uploaded_file is not None:
    temp_pdf_path = "temp_uploaded.pdf"
    with open(temp_pdf_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    output_excel = "output_tabel.xlsx"
    result_file = process_all_pdfs_in_folder(os.path.dirname(temp_pdf_path), output_excel)
    if result_file:
        parse_excel(result_file)
        with open(result_file, "rb") as f:
            st.download_button("Download Processed Excel", f, file_name=result_file)
        st.success("Processing Complete! Output saved as Sheet2.")
    else:
        st.warning("No tables found in the uploaded PDF.")

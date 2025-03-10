import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
import os
from openpyxl import load_workbook

st.title("📑 PDF Table Extractor & Formatter")

uploaded_files = st.file_uploader("Upload PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    output_excel = "output_tabel.xlsx"
    all_extracted_data = []

    st.write("🔄 **Memproses file...**")

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

    for file in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf_path = temp_pdf.name
            temp_pdf.write(file.read())

        st.write(f"🔄 Memproses: {file.name}")
        extracted_data = extract_tables_from_pdf(temp_pdf_path)

        if extracted_data:
            all_extracted_data.extend(extracted_data)

        os.remove(temp_pdf_path)

    if all_extracted_data:
        final_df = pd.concat(all_extracted_data, ignore_index=True)
        final_df.to_excel(output_excel, index=False, header=True)
        st.success("✅ Semua tabel berhasil diekstrak ke Excel!")

        df = pd.read_excel(output_excel, sheet_name='Sheet1', header=None)

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

        df = perbaiki_nilai_tidak_sejajar(df)
        df.to_excel(output_excel, index=False)
        st.success("✅ Format tabel diperbaiki!")

        cols_needed = [
            "tanggal temuan", "cara temuan", "waktu pendeteksian", "nama kantor", "provinsi", "kota",
            "jenis kontributor", "kantor kontributor", "nama kontributor", "dokumen pendukung", 
            "no. identitas", "keterangan", "provinsi", "kota", "kecamatan", "pecahan", "tahun emisi",
            "no. seri 1", "no. seri 2", "jumlah lembar", "jumlah lembar terima", "subtotal",
            "jumlah dianalisa", "hasil analisa", "subtotal"
        ]

        sheet1 = pd.read_excel(output_excel, sheet_name='Sheet1', header=None)
        data_rows = []
        current_row = {}
        found_first_subtotal = False

        for idx, row in sheet1.iterrows():
            for col_num, cell_value in enumerate(row):
                if cell_value in cols_needed:
                    if idx + 1 < len(sheet1):
                        next_val = sheet1.iloc[idx + 1, col_num]
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
            df['tanggal temuan'] = df['tanggal temuan'].dt.strftime('%d-%m-%Y')

        df.fillna(method='ffill', inplace=True)

        with pd.ExcelWriter(output_excel, engine='openpyxl', mode='a') as writer:
            df.to_excel(writer, sheet_name='Sheet2', index=False)

        st.success("✅ Data telah diparsing dan disimpan di Sheet2!")

        with open(output_excel, "rb") as f:
            st.download_button("⬇ Download Excel", f, file_name="output_tabel.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("⚠ Tidak ada tabel yang ditemukan dalam file PDF.")

import streamlit as st
import pdfplumber
import pandas as pd
import tempfile
import os
from openpyxl import load_workbook

st.title("📑 PDF Table Extractor & Cleaner")

# 1️⃣ Upload PDF
uploaded_files = st.file_uploader("Upload PDF", type=["pdf"], accept_multiple_files=True)

if uploaded_files:
    output_excel = "output_tabel.xlsx"
    all_extracted_data = []

    st.write("🔄 **Memproses file...**")

    def extract_tables_from_pdf(pdf_path):
        """Ekstrak tabel dari PDF dan gabungkan ke dalam DataFrame"""
        extracted_data = []
        pending_row = None
        previous_headers = None

        def clean_text(text):
            return str(text).strip().lower().replace('\n', ' ') if pd.notna(text) else ''

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
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

                    extracted_data.append(df)

        return pd.concat(extracted_data, ignore_index=True) if extracted_data else None

    for file in uploaded_files:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf_path = temp_pdf.name
            temp_pdf.write(file.read())

        st.write(f"🔄 Memproses: {file.name}")
        df = extract_tables_from_pdf(temp_pdf_path)

        if df is not None:
            df["Sumber File"] = file.name
            all_extracted_data.append(df)

        os.remove(temp_pdf_path)

    if all_extracted_data:
        final_df = pd.concat(all_extracted_data, ignore_index=True)
        final_df.to_excel(output_excel, index=False, header=False)
        st.success("✅ Semua tabel berhasil diekstrak ke Excel!")
    else:
        st.warning("⚠ Tidak ada tabel yang ditemukan dalam file PDF.")

    # 3️⃣ Perbaikan Struktur Data
    st.write("🔄 **Perbaikan format tabel...**")
    df = pd.read_excel(output_excel, header=None)

    for row in range(0, len(df) - 1, 2):
        header_row = df.iloc[row]
        value_row = df.iloc[row + 1]

        if value_row.first_valid_index() < header_row.first_valid_index():
            corrected_values = value_row.dropna().values
            valid_headers = header_row.dropna().index

            for i, col in enumerate(valid_headers):
                if i < len(corrected_values):
                    df.iat[row + 1, col] = corrected_values[i]

            df.iloc[row + 1, :valid_headers[0]] = None

    df.to_excel(output_excel, index=False, header=False)
    st.success("✅ Format tabel diperbaiki!")

    # 4️⃣ Parsing Data ke Format Terstruktur
    st.write("🔄 **Parsing data ke struktur yang rapi...**")

    cols_needed = [
        "tanggal temuan", "cara temuan", "waktu pendeteksian", "nama kantor", "provinsi", "kota",
        "jenis kontributor", "kantor kontributor", "nama kontributor", "dokumen pendukung",
        "no. identitas", "keterangan", "provinsi", "kota", "kecamatan", "pecahan", "tahun emisi",
        "no. seri 1", "no. seri 2", "jumlah lembar", "jumlah lembar terima", "subtotal",
        "jumlah dianalisa", "hasil analisa", "subtotal"
    ]

    sheet1 = pd.read_excel(output_excel, sheet_name="Sheet1", header=None)
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
        mask_no_date = df['tanggal temuan'].isna()
        df.loc[mask_no_date, 'provinsi_kontributor'] = df.loc[mask_no_date, 'provinsi']
        df.loc[mask_no_date, 'kota_kontributor'] = df.loc[mask_no_date, 'kota']
        df.loc[mask_no_date, ['provinsi', 'kota']] = None

    date_columns = ["tanggal temuan"]
    for col in date_columns:
        if col in df.columns:
            df[col] = df[col].dt.strftime('%d-%m-%Y')

    df.fillna(method='ffill', inplace=True)

    with pd.ExcelWriter(output_excel, engine='openpyxl', mode='a') as writer:
        df.to_excel(writer, sheet_name="Sheet2", index=False)

    st.success("✅ Data telah diparsing dan disimpan di **Sheet2**!")

    # 5️⃣ Download Hasil
    st.write("📥 **Download hasil akhir:**")
    with open(output_excel, "rb") as f:
        st.download_button("⬇ Download Excel", f, file_name="output_tabel.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

def extract_table_from_pdf(file):
    all_rows = []
    header_found = False
    header = []

    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            lines = page.extract_text().split('\n')
            for line in lines:
                row = line.strip().split()
                if not header_found and any("no" in col.lower() for col in row):
                    header = row
                    header_found = True
                    continue
                if header_found and len(row) >= len(header):
                    all_rows.append(row[:len(header)])
    return pd.DataFrame(all_rows, columns=header)

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

# Streamlit UI
st.title("PDF Tabel ke Excel Converter")

uploaded_file = st.file_uploader("Upload file PDF", type="pdf")

if uploaded_file is not None:
    try:
        df = extract_table_from_pdf(uploaded_file)
        st.success("Berhasil parsing PDF!")
        st.dataframe(df)

        excel_data = convert_df_to_excel(df)
        st.download_button(
            label="💾 Download Excel",
            data=excel_data,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")

import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

def extract_table_from_pdf(file):
    all_dataframes = []
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table_idx, table in enumerate(tables):
                if table:
                    df = pd.DataFrame(table)
                    df["page"] = page_num + 1
                    df["table"] = table_idx + 1
                    all_dataframes.append(df)
    return pd.concat(all_dataframes, ignore_index=True) if all_dataframes else pd.DataFrame()

def bersihkan_dataframe(df):
    try:
        df[0] = df[0].astype(str).str.replace(r'^(\d+)\s*(\w*)\.*', r'\1\2', regex=True)
        df[['no_clean', 'item_clean']] = df[0].str.extract(r'^(\d+[a-zA-Z]*)\s*(.*)$')
    except Exception as e:
        st.warning(f"Gagal membersihkan kolom: {e}")
    df.dropna(how='all', inplace=True)
    return df

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

# Streamlit UI
st.title("PDF ➜ Excel (Ambil Semua Tabel)")

uploaded_file = st.file_uploader("Upload file PDF", type="pdf")

if uploaded_file is not None:
    try:
        df = extract_table_from_pdf(uploaded_file)
        if df.empty:
            st.warning("❌ Tidak ditemukan tabel di PDF.")
        else:
            st.subheader("📄 Tabel Mentah")
            st.dataframe(df)

            df_clean = bersihkan_dataframe(df.copy())
            st.subheader("🧼 Tabel Setelah Dibersihkan")
            st.dataframe(df_clean)

        excel_data = convert_df_to_excel(df_clean)
        st.download_button(
            label="💾 Download Excel",
            data=excel_data,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")

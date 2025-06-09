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
                    header_row_idx = None
                    for idx, row in enumerate(table):
                        if row and "No." in row and "Item" in row:
                            header_row_idx = idx
                            break
                if header_row_idx is not None:
                    data = table[header_row_idx:]
                    header = data[0]
                    header = [str(h) if h is not None else "" for h in header]

                    seen = {}
                    new_header = []
                    for col in header:
                        if col in seen:
                            seen[col] += 1
                            new_header.append(f"{col}_{seen[col]}")
                        else:
                            seen[col] = 0
                            new_header.append(col)

                    df = pd.DataFrame(data[1:], columns=new_header)

                    if "Cavity sample" in df.columns:
                        cavity_idx = df.columns.get_loc("Cavity sample")
                        df = df.iloc[:, :cavity_idx+1]

                    df["page"] = page_num + 1
                    df["table"] = table_idx + 1
                    df = df.drop(columns=["Cavity sample", "page", "table"], errors="ignore")
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
st.title("CHECK SHEET SCAN QFROM")

uploaded_file = st.file_uploader("Upload file PDF", type="pdf")

if uploaded_file is not None:
    try:
        df = extract_table_from_pdf(uploaded_file)
        if df.empty:
            st.warning("❌ Tidak ditemukan tabel di PDF.")
        else:
            st.subheader("📄 Tabel Mentah")
            st.dataframe(df, use_container_width=True, hide_index=True)

            df_clean = bersihkan_dataframe(df.copy())
            st.subheader("🧼 Tabel Setelah Dibersihkan")
            st.dataframe(df_clean, use_container_width=True, hide_index=True)

        excel_data = convert_df_to_excel(df_clean)
        st.download_button(
            label="💾 Download Excel",
            data=excel_data,
            file_name="output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Gagal memproses file: {e}")

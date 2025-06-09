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

# DEBUG === PERSOALAN FOOTER ====>
uploaded_file = st.file_uploader("📤 Upload PDF berisi tabel", type="pdf")

if uploaded_file is not None:
    df = extract_table_from_pdf(uploaded_file)

    if not df.empty:
        st.subheader("🔎 Cek 10 Baris Terakhir dari Tabel Mentah")
        st.dataframe(df.tail(11), use_container_width=True)

        st.subheader("🧠 Isi Baris Terakhir Sebagai String (untuk deteksi keyword)")
        for idx, row in df.tail(11).iterrows():
            row_str = ' | '.join([str(cell).lower() for cell in row if pd.notnull(cell)]).strip()
            st.text(f"[{idx}] {row_str}")
    else:
        st.warning("📭 Dataframe kosong, tidak bisa menampilkan baris terakhir.")
else:
    st.info("👆 Silakan upload file PDF dulu.")
# DEBUG === PERSOALAN FOOTER (END) ====>

def reverse_text(text):
    if not isinstance(text, str):
        return text
    text = text.replace('\n', ' ')
    return text[::-1].strip()

def bersihkan_dataframe(df):
    try:
        df[0] = df[0].astype(str).str.replace(r'^(\d+)\s*(\w*)\.*', r'\1\2', regex=True)
        df[['no_clean', 'item_clean']] = df[0].str.extract(r'^(\d+[a-zA-Z]*)\s*(.*)$')
    except Exception as e:
        st.warning(f"Gagal membersihkan kolom: {e}")
    df.dropna(how='all', inplace=True)

    cols_lower = {col.lower().strip(): col for col in df.columns}

    if "patrol" in cols_lower and "setup" in cols_lower:
        col_patrol = cols_lower["patrol"]
        col_setup = cols_lower["setup"]

        temp = df[col_patrol].copy()
        df[col_patrol] = df[col_setup]
        df[col_setup] = temp

        df[col_patrol] = df[col_patrol].apply(reverse_text)
        df[col_setup] = df[col_setup].apply(reverse_text)

        st.write(f"Isi kolom '{col_patrol}' dan '{col_setup}' telah ditukar dan teks dibalik sesuai kebutuhan.")

        for col in df.columns:
            if col.strip().lower() == "tfihS/x1":
                df.rename(columns={col: "1x/shift"}, inplace=True)
                print(f'Column "{col}" renamed to "1x/shift"')
            elif col.strip().lower() == "1x/shift":
                df.rename(columns={col: "tfihS/x1"}, inplace=True)
        
        # Tambahan: buang baris footer seperti keputusan/approved by
        keywords = ["keputusan", "keterangan", "approved", "checked", "disetujui", "diperiksa","dibuat", "nama", "tanggal", "ttd"]
        keyword_found_idx = None

        for idx, row in df.iterrows():
            row_str = ' '.join([str(cell).lower() for cell in row if pd.notnull(cell)]).strip()
            if any(keyword in row_str for keyword in keywords):
                keyword_found_idx = idx
                break

        if keyword_found_idx is not None:
            df = df.iloc[:keyword_found_idx]
            st.info(f"🧹 Footer terdeteksi mulai dari baris index {keyword_found_idx}. Semua baris setelahnya dibuang.")

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
            # st.dataframe(df, use_container_width=True, hide_index=True)
            st.dataframe(df, use_container_width=True)

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

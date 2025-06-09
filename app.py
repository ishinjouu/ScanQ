import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

def normalize_text(text):
    if not isinstance(text, str):
        return text
    return text.replace('\n', '').strip().lower()

def maybe_flip_text(text):
    if not isinstance(text, str) or not text.strip():
        return text

    reversed_text = text[::-1].strip()
    known_words = ["setup", "patrol", "1x/shift", "job setup", "portal", "pu", "1x/day", "shift"]

    normalized_reversed = normalize_text(reversed_text)
    normalized_text = normalize_text(text)

    # Debugging flipping candidates
    st.write(f"DEBUG maybe_flip_text: original='{text}', reversed='{reversed_text}', "
             f"normalized_original='{normalized_text}', normalized_reversed='{normalized_reversed}'")

    if normalized_reversed in known_words:
        st.write(f"[FLIP] Flipping '{text}' → '{reversed_text}'")
        return reversed_text

    return text

def reverse_text(text):
    if not isinstance(text, str):
        return text
    return text[::-1].strip()

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
    if all_dataframes:
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        st.write(f"DEBUG: Extracted combined DataFrame shape: {combined_df.shape}")
        return combined_df
    else:
        st.write("DEBUG: No tables found in PDF.")
        return pd.DataFrame()

def hapus_footer(df):
    keywords = [
        "keputusan", "keterangan", "approved", "checked", "disetujui",
        "diperiksa", "dibuat", "nama", "tanggal", "ttd"
    ]

    footer_start_idx = None
    for idx in df.index:
        row = df.loc[idx]
        row_str = ' '.join([str(x).lower() for x in row if pd.notnull(x)])
        if any(k in row_str for k in keywords):
            footer_start_idx = idx
            break

    if footer_start_idx is not None:
        st.info(f"🧹 Footer detected starting at row index {footer_start_idx}, footer rows removed.")
        return df.loc[:footer_start_idx-1].copy()
    else:
        st.info("✅ No footer detected.")
        return df

def detect_and_fix_reversed_columns(df):
    known_cols = ["setup", "patrol"]

    def is_reversed_match(col_name):
        norm_col = col_name.replace('\n', '').lower().strip()
        rev_col = norm_col[::-1]
        return rev_col in known_cols

    new_columns = []
    for col in df.columns:
        if is_reversed_match(col):
            fixed_col = col[::-1].replace('\n', '').strip()
            new_columns.append(fixed_col)
            st.write(f"DEBUG: Renamed reversed column '{col}' to '{fixed_col}'")
        else:
            new_columns.append(col)
    df.columns = new_columns
    return df

def bersihkan_dataframe(df):
    df = detect_and_fix_reversed_columns(df)

    try:
        df["No."] = df["No."].astype(str).str.replace(r'^(\d+)\s*(\w*)\.*', r'\1\2', regex=True)
        df[['no_clean', 'item_clean']] = df["No."].str.extract(r'^(\d+[a-zA-Z]*)\s*(.*)$')

        df['item_clean'] = df['item_clean'].apply(
            lambda x: maybe_flip_text(x) if isinstance(x, str) and len(x.strip()) <= 20 else x
        )
    except Exception as e:
        st.warning(f"Failed to clean dataframe columns: {e}")

    df.dropna(how='all', inplace=True)

    cols_lower = {col.lower(): col for col in df.columns}

    if "patrol" in cols_lower and "setup" in cols_lower:
        col_patrol = cols_lower["patrol"]
        col_setup = cols_lower["setup"]
        st.write(f"DEBUG: Flipping columns '{col_patrol}' and '{col_setup}'")

        df[col_patrol] = df[col_patrol].apply(reverse_text)
        df[col_setup] = df[col_setup].apply(reverse_text)

    df = hapus_footer(df)

    return df

def convert_df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

st.title("CHECK SHEET SCAN QFORM")

uploaded_file = st.file_uploader("📤 Upload file PDF", type="pdf")

if uploaded_file is not None:
    try:
        df = extract_table_from_pdf(uploaded_file)

        if df.empty:
            st.warning("❌ No tables detected in the PDF")
        else:
            df_clean = bersihkan_dataframe(df.copy())

            st.subheader("🧼 Cleaned Table")
            st.dataframe(df_clean, use_container_width=True, hide_index=True)

            excel_data = convert_df_to_excel(df_clean)
            st.download_button(
                label="💾 Download Excel",
                data=excel_data,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Failed to process file: {e}")

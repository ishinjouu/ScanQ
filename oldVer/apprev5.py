import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# ---------- Helpers ----------

def normalize_text(text):
    if not isinstance(text, str):
        return text
    return text.replace('\n', '').strip().lower()

def maybe_flip_text(text):
    if not isinstance(text, str) or not text.strip():
        return text

    raw = text.replace('\n', '').strip()
    reversed_text = raw[::-1]
    normalized_reversed = normalize_text(reversed_text)

    replacements = {
        "setup": "Set Up",
        "patrol": "Patrol",
        "1x/shift": "1x/Shift",
        "job setup": "Job Set Up",
        "portal": "Portal",
        "up": "Up",
        "1x/day": "1x/Day",
        "shift": "Shift",
        "allpointifjobsetup": "All Point If Job Set Up",
        "4allpointifjobsetup": "All Point If Job Set Up"
    }

    if normalized_reversed in replacements:
        pretty_text = replacements[normalized_reversed]
        st.write(f"[FLIP ✨] '{text}' → '{pretty_text}'")
        return pretty_text

    return text


def reverse_text(text):
    if not isinstance(text, str):
        return text
    return text.replace('\n', ' ')[::-1].strip()

def convert_df_to_excel(df):
    if df is None or df.empty:
        st.warning("⚠ Cleaned dataframe is empty or None. Skipping Excel export.")
        return None

    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"💥 Excel writing failed: {e}")
        return None

# ---------- PDF Table Extraction ----------

def extract_table_from_pdf(file):
    all_dataframes = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    header_row_idx = None
                    for idx, row in enumerate(table):
                        if row and "No." in row and "Item" in row:
                            header_row_idx = idx
                            break
                    if header_row_idx is not None:
                        data = table[header_row_idx:]
                        header = ["" if h is None or str(h).strip().lower() == "none" else str(h).strip() for h in data[0]]
                        seen = set()
                        new_header = []
                        for col in header:
                            clean_col = (col or '').strip()
                            if clean_col and clean_col not in seen:
                                seen.add(clean_col)
                                new_header.append(clean_col)
                            else:
                                new_header.append('')

                        df = pd.DataFrame(data[1:], columns=new_header)
                        df = df[~df.apply(lambda row: row.astype(str).str.strip().eq('').all(), axis=1)]
                        df = df.loc[:, df.columns.str.strip() != '']
                        
                        all_dataframes.append(df)
    return pd.concat(all_dataframes, ignore_index=True) if all_dataframes else pd.DataFrame()

# ---------- Footer & Flip Cleaners ----------

def hapus_footer(df):
    keywords = ["keputusan", "keterangan", "approved", "checked", "disetujui", "diperiksa", "dibuat", "nama", "tanggal", "ttd"]
    footer_start_idx = None
    for idx in df.index:
        row = df.loc[idx]
        row_str = ' '.join([str(x).lower() for x in row if pd.notnull(x)])
        if any(k in row_str for k in keywords):
            footer_start_idx = idx
            break

    if footer_start_idx is not None:
        st.info(f"🧹 Footer detected at index {footer_start_idx}. Removing it.")
        return df.loc[:footer_start_idx - 1].copy()
    else:
        st.info("✅ No footer found.")
        return df

def detect_and_fix_reversed_columns(df):
    replacements = {
        "setup": "Set Up",
        "patrol": "Patrol",
        "1x/shift": "1x/Shift",
        "job setup": "Job Set Up",
        "portal": "Portal",
        "up": "Up",
        "1x/day": "1x/Day",
        "shift": "Shift",
        "allpointifjobsetup": "All Point If Job Set Up",
        "4allpointifjobsetup": "All Point If Job Set Up"
    }

    def try_fix_header(col_name):
        norm = normalize_text(col_name)
        rev = norm[::-1]
        if rev in replacements:
            return replacements[rev]
        elif norm in replacements:
            return replacements[norm]
        else:
            return col_name.strip()

    new_columns = [try_fix_header(col) for col in df.columns]
    df.columns = new_columns
    return df


# ---------- Final Cleaning Pipeline ----------

def bersihkan_dataframe(df):
    df = detect_and_fix_reversed_columns(df)

    try:
        if "No." in df.columns:
            df["No."] = df["No."].astype(str).str.replace(r'^(\d+)\s*(\w*)\.*', r'\1\2', regex=True)
        else:
            st.warning("🛑 Column 'No.' not found.")
    except Exception as e:
        st.warning(f"Cleaning 'No.' column failed: {e}")

    df.dropna(how='all', inplace=True)

    df = hapus_footer(df)

    # Drop clean columns if present
    for col in ["no_clean", "item_clean"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    # Drop columns after and INCLUDING "Cavity sample"
    # Remove 'Cavity sample' and everything to the right of it
    df_cols = df.columns.tolist()
    if "Cavity sample" in df_cols:
        cavity_idx = df_cols.index("Cavity sample")
        keep_cols = df_cols[:cavity_idx]  # truly excludes 'Cavity sample' itself
        df = df[keep_cols]
        st.info("🧼 'Cavity sample' detected and *exorcised*—along with all trailing columns. (≧д≦ヾ)")


    df = df.applymap(lambda x: '' if str(x).strip().lower() == 'none' else maybe_flip_text(x))
    return df


# ---------- Streamlit UI ----------

st.title("📄 CHECK SHEET SCAN QFORM")

uploaded_file = st.file_uploader("📤 Upload PDF file", type="pdf")

if uploaded_file is not None:
    try:
        df = extract_table_from_pdf(uploaded_file)

        if df.empty:
            st.warning("❌ No tables detected in the PDF.")
        else:
            df_clean = bersihkan_dataframe(df.copy())
            st.subheader("🧼 Cleaned Table")
            st.dataframe(df_clean, use_container_width=True, hide_index=True)

            excel_data = convert_df_to_excel(df_clean)
            if excel_data:
                st.download_button(
                    label="💾 Download Excel",
                    data=excel_data,
                    file_name="output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"🔥 Error while processing the file: {e}")
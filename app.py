import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import range_boundaries
from tempfile import NamedTemporaryFile

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
        return replacements[normalized_reversed]
    return text

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
                        fixed_rows = []
                        for row in data[1:]:
                            if len(row) < len(new_header):
                                row.extend([''] * (len(new_header) - len(row)))
                            elif len(row) > len(new_header):
                                row = row[:len(new_header)]
                            fixed_rows.append(row)
                        df = pd.DataFrame(fixed_rows, columns=new_header)
                        df = df[~df.apply(lambda row: row.astype(str).str.strip().eq('').all(), axis=1)]
                        df = df.loc[:, df.columns.str.strip() != '']
                        all_dataframes.append(df)
    return pd.concat(all_dataframes, ignore_index=True) if all_dataframes else pd.DataFrame()

# ---------- Footer & Cleaning ----------

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
        return df.loc[:footer_start_idx - 1].copy()
    else:
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
    df.columns = [try_fix_header(col) for col in df.columns]
    return df

# ---------- Fix for Vertical Data ----------
def fill_vertical_blocks(df, column_name):
    """Propagate vertical values in the given column downward until a new one appears."""
    active_value = None
    for i in df.index:
        val = df.at[i, column_name] if column_name in df.columns else None
        if pd.notna(val) and str(val).strip() != "":
            active_value = val
        elif active_value:
            df.at[i, column_name] = active_value
    return df

# ---------- Final Cleaning Pipeline ----------

def bersihkan_dataframe(df):
    df = detect_and_fix_reversed_columns(df)
    try:
        if "No." in df.columns:
            df["No."] = df["No."].astype(str).str.replace(r'^(\d+)\s*(\w*)\.*', r'\1\2', regex=True)
    except Exception as e:
        st.warning(f"Cleaning 'No.' column failed: {e}")
    df.dropna(how='all', inplace=True)
    df = hapus_footer(df)
    df = df.applymap(lambda x: '' if str(x).strip().lower() == 'none' else maybe_flip_text(x))
    for col in ["Set Up", "Patrol"]:
        if col in df.columns:
            df = fill_vertical_blocks(df, col)
    return df

# ---------- Streamlit UI ----------

st.title("📄 CHECK SHEET SCAN QFORM")

uploaded_file = st.file_uploader("📄 Upload PDF file", type="pdf")

if uploaded_file is not None:
    try:
        df = extract_table_from_pdf(uploaded_file)
        if df.empty:
            st.warning("❌ No tables detected in the PDF.")
        else:
            df_clean = bersihkan_dataframe(df.copy())
            st.subheader("🧼 Cleaned Table")
            st.dataframe(df_clean, use_container_width=True, hide_index=True)

            wb = load_workbook("data101.xlsx")
            ws = wb.active

            start_row = 13
            start_col = 1
            merged_ranges = list(ws.merged_cells.ranges)
            merged_map = {}
            for cell_range in merged_ranges:
                min_col, min_row, max_col, max_row = range_boundaries(str(cell_range))
                for row in range(min_row, max_row + 1):
                    for col in range(min_col, max_col + 1):
                        merged_map[(row, col)] = (min_row, min_col)

            written_cells = set()
            for i, row in enumerate(dataframe_to_rows(df_clean, index=False, header=False)):
                for j, value in enumerate(row):
                    r = start_row + i
                    c = start_col + j
                    target = merged_map.get((r, c), (r, c))
                    if target not in written_cells:
                        ws.cell(row=target[0], column=target[1], value=value)
                        written_cells.add(target)

            with NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                wb.save(tmp.name)
                tmp.seek(0)
                excel_bytes = tmp.read()

            st.download_button(
                label="📥 Download Excel (with Template Header)",
                data=excel_bytes,
                file_name="filled_data101.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"🔥 Error while processing the file: {e}")

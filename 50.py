import streamlit as st
import pdfplumber
import pandas as pd
import re
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
    text = text.replace('\n', ' ')
    return text[::-1].strip()

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
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            page_tables = []  # kumpulan df dari satu halaman

            for table_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    continue

                header_row_idx = None
                for idx, row in enumerate(table):
                    if row and any("Item" in str(cell) for cell in row):  # fleksibelin deteksi header
                        header_row_idx = idx
                        break

                if header_row_idx is not None:
                    data = table[header_row_idx:]
                    header = [str(h) if h is not None else "" for h in data[0]]

                    # handle duplicate column names
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
                    page_tables.append(df)
                else:
                    # asumsi tabel tanpa header (mungkin tabel lanjutannya)
                    df = pd.DataFrame(table)
                    page_tables.append(df)

            # MERGE semua tabel dalam 1 halaman secara horizontal (axis=1)
            if page_tables:
                try:
                    merged_df = pd.concat(page_tables, axis=1)
                    all_dataframes.append(merged_df)
                except Exception as e:
                    st.warning(f"⚠️ Gagal merge tabel di halaman {page_num+1}: {e}")

    return pd.concat(all_dataframes, ignore_index=True) if all_dataframes else pd.DataFrame()

# ---------- Bismillah ----------
def group_rows_by_item(df):
    if "Item" not in df.columns or "Standard" not in df.columns:
        return df

    grouped_rows = []
    buffer_row = None

    for idx, row in df.iterrows():
        item = str(row.get("Item", "")).strip()
        std = str(row.get("Standard", "")).strip()
        detail = str(row.get("Detail Standard", "")).strip() if "Detail Standard" in df.columns else ""

        if item != "":  # Baris utama baru
            if buffer_row is not None:
                grouped_rows.append(buffer_row)
            buffer_row = row.copy()
        else:
            if buffer_row is not None:
                buffer_row["Standard"] = f"{buffer_row['Standard']}, {std}".strip(', ')
                if "Detail Standard" in buffer_row and detail:
                    buffer_row["Detail Standard"] = f"{buffer_row['Detail Standard']}, {detail}".strip(', ')

    if buffer_row is not None:
        grouped_rows.append(buffer_row)

    return pd.DataFrame(grouped_rows)

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
    df = group_rows_by_item(df)

    # Drop clean columns if present
    for col in ["no_clean", "item_clean"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    # Keep only 1 Cavity Sample and drop the rest
    cavity_columns = [col for col in df.columns if col.lower().startswith("cavity sample")]
    if cavity_columns:
        # keep only the first occurrence
        df_cols = df.columns.tolist()
        first_cavity = cavity_columns[0]
        cavity_idx = df_cols.index(first_cavity)
        keep_cols = df_cols[:cavity_idx + 1]
        df = df[keep_cols]

        if len(cavity_columns) > 1:
            df.drop(columns=[col for col in cavity_columns[1:]], inplace=True)

        st.info("🧼 Kolom 'Cavity sample' dirapihkan, yang ganda dibuang satu aja~ 😎")

    if "Standard" in df.columns:
        valid_patterns = re.compile(
            r'^(?:\s*'
            r'((?:min|max|\d+|[a-zA-Z]+)'
            r'(?:\s*[-–—]\s*(?:min|max|\d+|[a-zA-Z]+))?'
            r'(?:\s*,\s*(?:min|max|\d+|[a-zA-Z]+))*'
            r')\s*)$',
            re.IGNORECASE
        )

        def validate_standard(value):
            if isinstance(value, str):
                value_clean = value.strip()
                if valid_patterns.match(value_clean):
                    return value
                else:
                    return value  # bisa ditandai kalau mau
            return value

        df["Standard"] = df["Standard"].apply(validate_standard)

    df = df.applymap(lambda x: '' if str(x).strip().lower() == 'none' else maybe_flip_text(x))
    return df

# ---------- Transform ke Format Final ----------

def transform_to_final_format(df):
    if "No." not in df.columns or "Item" not in df.columns:
        st.warning("⚠️ Kolom 'No.' dan 'Item' dibutuhkan untuk transformasi format.")
        return pd.DataFrame()

    # Forward-fill semua nilai kosong berdasarkan baris sebelumnya
    df = df.ffill()

    final_rows = []
    current_section = ""

    for idx, row in df.iterrows():
        no = str(row.get("No.", "")).strip()
        item = str(row.get("Item", "")).strip()
        std = str(row.get("Standard", "")).strip()
        control = str(row.get("Control Method", "")).strip()

        # Deteksi Section
        if re.match(r'^[IVXLC]+\.', no, re.IGNORECASE) or (no.isupper() and len(no) > 1):
            current_section = f"{no} {item}".strip()
            continue

        # Skip baris yang kosong total (sisa OCR error)
        if not no and not item:
            continue

        # Default values
        point_check = no
        catatan = "-"
        jenis_point = "tanpa ukur"
        jenis_pengecekan_1 = ""
        jenis_pengecekan_2 = ""

        # Deteksi jenis point
        if 'cmm' in control.lower():
            jenis_point = "dengan cmm"
        elif 'ukur' in control.lower() or any(x in std.lower() for x in ['mm', 'min', 'max', 'diameter']):
            jenis_point = "dengan ukur"

        # Catatan visual
        if control.lower() == "visual" and std:
            catatan = std
        elif control.lower() == "visual" and not std:
            catatan = "-"
        elif std:
            catatan = std

        # Deteksi jenis pengecekan dari kolom lain
        other_cols = [str(v).lower() for v in row.values if isinstance(v, str)]
        if any("job" in c and "set" in c for c in other_cols):
            jenis_pengecekan_1 = "Job Set Up"
        if any("1x/shift" in c for c in other_cols):
            jenis_pengecekan_2 = "1x/Shift"
        elif any("1x/day" in c for c in other_cols):
            jenis_pengecekan_2 = "1x/Day"

        final_rows.append({
            "Section": current_section,
            "Point Check": point_check,
            "Jenis Point": jenis_point,
            "Catatan": catatan,
            "Item Check": item,
            "Control Method": control,
            "Standard": std,
            "Jenis Pengecekan_1": jenis_pengecekan_1,
            "Jenis Pengecekan_2": jenis_pengecekan_2
        })

    return pd.DataFrame(final_rows)

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

            df_final = transform_to_final_format(df_clean)

            # Tampilkan jika tidak kosong
            if not df_final.empty:
                st.subheader("📊 Final Structured Format")
                st.dataframe(df_final, use_container_width=True, hide_index=True)
            else:
                st.warning("❌ Format final kosong atau gagal diubah.")

            excel_data = convert_df_to_excel(df_final)
            if excel_data:
                st.download_button(
                    label="💾 Download Excel",
                    data=excel_data,
                    file_name="output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"🔥 Error while processing the file: {e}")

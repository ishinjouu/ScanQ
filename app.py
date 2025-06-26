import streamlit as st
import pdfplumber
import pandas as pd
import numpy as np
import hashlib
import re
import requests
import json
import traceback
from typing import List
from io import BytesIO
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font

def sanitize_for_json(value):
    if isinstance(value, float) and (pd.isna(value) or not np.isfinite(value)):
        return None
    if isinstance(value, list):
        return [sanitize_for_json(v) for v in value]
    if isinstance(value, dict):
        return {k: sanitize_for_json(v) for k, v in value.items()}
    return value

def send_df_to_api(df):
    try:
        records = df.to_dict(orient="records")
        safe_records = [sanitize_for_json(row) for row in records]
        response = requests.post("http://192.168.148.224:5000/api/submit", json=safe_records)
        response.raise_for_status()
        st.success("✅ Data successfully sent to the API!")
    except Exception as e:
        st.error(f"❌ Failed to send data: {e}")

def get_file_hash(file):
    return hashlib.md5(file.getvalue()).hexdigest()

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
        # st.write(f"[FLIP] '{text}' → '{pretty_text}'")
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
            df.to_excel(writer, index=False, sheet_name='Sheet1')
        output.seek(0)
        wb = load_workbook(output)
        ws = wb.active
        for col_idx, column_cells in enumerate(ws.columns, 1):
            max_length = 0
            col_letter = get_column_letter(col_idx)
            col_name = column_cells[0].value
            for row_idx, cell in enumerate(column_cells, 1):
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                    if row_idx == 1:
                        cell.font = Font(bold=True)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        if col_name and str(col_name).strip().lower() == "section":
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                        else:
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                except:
                    pass

            adjusted_width = max_length + 2
            ws.column_dimensions[col_letter].width = adjusted_width
        adjusted_output = BytesIO()
        wb.save(adjusted_output)
        adjusted_output.seek(0)
        return adjusted_output

    except Exception as e:
        st.error(f"💥 Excel writing failed: {e}")
        return None

# ---------- PDF Table Extraction ----------

@st.cache_data(show_spinner="🔄 Mengambil tabel dari PDF...")
def extract_table_from_pdf(file):
    all_dataframes = []
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            page_tables = [] 
            for table_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    continue
                header_row_idx = None
                for idx, row in enumerate(table):
                    if row and any("Item" in str(cell) for cell in row):
                        header_row_idx = idx
                        break
                if header_row_idx is not None:
                    data = table[header_row_idx:]
                    header = [str(h) if h is not None else "" for h in data[0]]
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
                    df = pd.DataFrame(table)
                    page_tables.append(df)
            if page_tables:
                try:
                    merged_df = pd.concat(page_tables, axis=0, ignore_index=True)
                    all_dataframes.append(merged_df)
                except Exception as e:
                    st.warning(f"⚠️ Gagal merge tabel di halaman {page_num+1}: {e}")
    return pd.concat(all_dataframes, ignore_index=True) if all_dataframes else pd.DataFrame()

def group_rows_by_item(df):
    if "Item" not in df.columns or "Standard" not in df.columns:
        return df
    grouped_rows = []
    buffer_row = None
    for idx, row in df.iterrows():
        item = str(row.get("Item", "")).strip()
        std = str(row.get("Standard", "")).strip()
        detail = str(row.get("Detail Standard", "")).strip() if "Detail Standard" in df.columns else ""
        if item != "": 
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

def fill_patrol_column(df):
    df["_is_section"] = df["No."].astype(str).str.match(r'^\s*(I|II|III|IV|V|VI|VII|VIII|IX|X)\b')
    last_valid_patrol = None
    patrol_filled = []

    for is_section, patrol in zip(df["_is_section"], df["Patrol"]):
        if is_section:
            patrol_filled.append(patrol)
            last_valid_patrol = None
        else:
            if pd.notna(patrol):
                last_valid_patrol = patrol
                patrol_filled.append(patrol)
            else:
                patrol_filled.append(last_valid_patrol)

    df["Patrol"] = patrol_filled
    df.drop(columns=["_is_section"], inplace=True)
    return df

def fill_setup_from_patrol(df):
    if "Set Up" not in df.columns:
        df["Set Up"] = None
    if "Patrol" not in df.columns:
        st.warning("🛑 Kolom 'Patrol' tidak ditemukan.")
        return df
    if "No." not in df.columns:
        st.warning("🛑 Kolom 'No.' tidak ditemukan.")
        return df

    df["_is_section"] = df["No."].astype(str).str.match(r'^\s*(I|II|III|IV|V|VI|VII|VIII|IX|X)\b')
    df["_filled_patrol"] = df["Patrol"].ffill()

    setup_filled = []
    for is_section, patrol_val in zip(df["_is_section"], df["_filled_patrol"]):
        if is_section:
            setup_filled.append(None)
        else:
            if pd.notna(patrol_val) and str(patrol_val).strip() != "":
                setup_filled.append("Job Setup")
            else:
                setup_filled.append(None)

    df["Set Up"] = setup_filled

    df.drop(columns=["_is_section", "_filled_patrol"], inplace=True)
    return df

# ---------- Final Cleaning ----------

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
    df = fill_setup_from_patrol(df)
    df = hapus_footer(df)
    df = group_rows_by_item(df)
    df = fill_patrol_column(df)
    
    for col in ["no_clean", "item_clean"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)
    cavity_columns = [col for col in df.columns if col.lower().startswith("cavity sample")]
    if cavity_columns:
        df_cols = df.columns.tolist()
        first_cavity = cavity_columns[0]
        cavity_idx = df_cols.index(first_cavity)
        keep_cols = df_cols[:cavity_idx + 1]
        df = df[keep_cols]
        if len(cavity_columns) > 1:
            df.drop(columns=[col for col in cavity_columns[1:]], inplace=True)
        st.info("🧼 Kolom 'Cavity sample' dirapihkan")

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
                    return value  
            return value
        df["Standard"] = df["Standard"].apply(validate_standard)
    df = df.applymap(lambda x: '' if str(x).strip().lower() == 'none' else maybe_flip_text(x))
    return df

# ---------- Transform ke Format Final ----------

def is_section_row(row):
    joined = ' '.join([str(cell) for cell in row if pd.notnull(cell)]).strip()
    return bool(re.match(r'^(I|II|III|IV|V|VI|VII|VIII|IX|X)\b', joined)) and joined.isupper()

def extract_section_title(row):
    return ' '.join([str(cell) for cell in row if pd.notnull(cell)]).strip()

def clean_empty_rows(df_final, max_null=5):
    def is_row_useless(row):
        item_check = row.get("item_check", "")
        item_check_empty = pd.isna(item_check) or item_check.strip() == ""
        empty_count = row.isna().sum() + (row == '').sum()
        return item_check_empty and empty_count >= max_null
    df_cleaned = df_final[~df_final.apply(is_row_useless, axis=1)].copy()
    return df_cleaned

def merge_point_item(df):
    new_point_check = []
    new_item_check = []
    for point, item in zip(df["point_check"], df["item_check"]):
        item_str = str(item).strip()
        if re.match(r'^[a-zA-Z]\.', item_str):
            letter = item_str[0] 
            remaining_item = re.sub(r'^[a-zA-Z]\.\s*', '', item_str, count=1)  
            new_point_check.append(f"{point}{letter}")
            new_item_check.append(remaining_item)
        else:
            new_point_check.append(point)
            new_item_check.append(item_str)
    df["point_check"] = new_point_check
    df["item_check"] = new_item_check
    return df

def roman_to_int(roman):
    roman = roman.upper()
    roman_dict = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100}
    result, prev = 0, 0
    for char in reversed(roman):
        val = roman_dict.get(char, 0)
        result += val if val >= prev else -val
        prev = val
    return result

def move_single_caps_to_note(row):
    standard = str(row.get("standard", "")).strip()
    catatan = str(row.get("catatan", "")).strip()
    match = re.search(r'(^|\s)([A-Z])(?![\w.])($|\s)', standard)
    if match:
        single_cap = match.group(2)
        tag = f"[{single_cap}]"
        if tag not in catatan:
            catatan = f"{catatan} {tag}".strip()
        standard = re.sub(r'(^|\s)([A-Z])(?![\w.])($|\s)', ' ', standard).strip()
    row["standard"] = standard
    row["catatan"] = catatan
    return row

def copy_special_measurements_to_note(row):
    standard = str(row.get("standard", "")).strip()
    catatan = str(row.get("catatan", "")).strip()
    jenis_point = str(row.get("jenis_point", "")).strip()
    if jenis_point not in ["Dengan Ukur", "Dengan CMM"]:
        return row

    min_max_match = re.search(r"\b([Mm]in|[Mm]ax)\s*\d+(?:\.\d+)?", standard)
    if min_max_match:
        tag = f"[{min_max_match.group(0).strip()}]"
        if tag not in catatan:
            catatan += f" {tag}"

    size_patterns = [
        r"[Ø°]\d+(?:\.\d+)?\s*±\s*[+−-]?\d+(?:\.\d+)?",                            # Ø10 ±0.1 atau °10 ± 0.5
        r"[Ø°]\d+(?:\.\d+)?\s*\(\s*\d+(?:\.\d+)?\s*~\s*[+−-]?\d+(?:\.\d+)?\s*\)",  # Ø6.1 ( 0 ~ +0.1 )
    ]

    ukuran_found = None
    for pattern in size_patterns:
        match = re.search(pattern, standard)
        if match:
            ukuran_found = match.group(0).strip().replace("[", "").replace("]", "")
            break  

    if ukuran_found and ukuran_found not in catatan:
        catatan += f" {ukuran_found}"

    row["catatan"] = catatan.strip()
    return row

def normalize_note_tags(catatan):
    if not isinstance(catatan, str):
        return catatan

    all_tags = re.findall(r"\[([^\[\]]+)\]", catatan)

    allowed_tags = []
    for tag in all_tags:
        tag_clean = tag.strip()
        if tag_clean in ["F", "M"]:
            allowed_tags.append(f"[{tag_clean}]")
        elif re.match(r"(?i)^min\s+\d+(\.\d+)?$", tag_clean):
            allowed_tags.append(f"[{tag_clean}]")
        elif re.match(r"(?i)^max\s+\d+(\.\d+)?$", tag_clean):
            allowed_tags.append(f"[{tag_clean}]")

    for tag in allowed_tags:
        catatan = re.sub(r"\[[^\[\]]+\]", "", catatan)

    ordered = []
    for tag in ["[F]", "[M]"]:
        if tag in allowed_tags:
            ordered.append(tag)
    for tag in allowed_tags:
        if tag not in ["[F]", "[M]"]:
            ordered.append(tag)

    catatan_clean = catatan.strip()
    result = " ".join(ordered)
    if catatan_clean:
        result = f"{result} {catatan_clean}"

    return result.strip()

def parse_standard_value(row):
    standard = str(row.get("standard", "")).strip()
    jenis_point = row.get("jenis_point", "")

    if jenis_point == "Tanpa Ukur":
        return pd.Series([None, None, None], index=["std_value", "std_min", "std_max"])

    # 1. 0 ±0.2
    match1 = re.match(r'^([\d.]+)\s*±\s*([\d.]+)$', standard)
    if match1:
        nominal = float(match1.group(1))
        delta = float(match1.group(2))
        return pd.Series([nominal, -delta, delta], index=["std_value", "std_min", "std_max"])

    # 2. Ø10 ±0.1 atau °5.5 ±0.2
    match2 = re.match(r'^[Ø°]\s*(\d+(?:\.\d+)?)\s*±\s*([+−-]?\d+(?:\.\d+)?)$', standard)
    if match2:
        nominal = float(match2.group(1))
        delta = float(match2.group(2).replace("−", "-"))
        return pd.Series([nominal, -abs(delta), abs(delta)], index=["std_value", "std_min", "std_max"])

    # 3. Ø6.1 ( 0 ~ +0.1 )
    match3 = re.match(r'^[Ø°]?\s*(-?\d+(?:\.\d+)?)\s*\(\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\)', standard)
    if match3:
        nominal = float(match3.group(1))
        lower = float(match3.group(2))
        upper = float(match3.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])

    # 4. [ 163.89 ± 0.25 ]
    match4 = re.match(r'^\[\s*(-?\d+(?:\.\d+)?)\s*±\s*([\d.]+)\s*\]$', standard)
    if match4:
        nominal = float(match4.group(1))
        delta = float(match4.group(2))
        return pd.Series([nominal, -delta, delta], index=["std_value", "std_min", "std_max"])

    # 5. Min/Max
    match5 = re.search(r'\b(Min|Max)\s*(\d+(?:\.\d+)?)\b', standard, re.IGNORECASE)
    if match5:
        kind = match5.group(1).lower()
        value = float(match5.group(2))
        return pd.Series([0, value if kind == "min" else 0, value if kind == "max" else 0], index=["std_value", "std_min", "std_max"])

    # 6. ( Reff : 0 ~ +0.5 )
    match6 = re.match(r'^\(.*?:\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\)$', standard)
    if match6:
        lower = float(match6.group(1))
        upper = float(match6.group(2))
        return pd.Series([0, lower, upper], index=["std_value", "std_min", "std_max"])

    # 7. [0 (0 ~ +0.3]
    match7 = re.match(r'^\[\s*(-?\d+(?:\.\d+)?)\s*\(\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?).*$', standard)
    if match7:
        nominal = float(match7.group(1))
        lower = float(match7.group(2))
        upper = float(match7.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])

    # 8. [Ø30.5 ± 0.2]
    match8 = re.match(r'^\[\s*[Ø°]?\s*(\d+(?:\.\d+)?)\s*±\s*([+−-]?\d+(?:\.\d+)?)\s*\]$', standard)
    if match8:
        nominal = float(match8.group(1))
        delta = float(match8.group(2).replace("−", "-"))
        return pd.Series([nominal, -abs(delta), abs(delta)], index=["std_value", "std_min", "std_max"])

    # 9. [reff. 0 (0 ~ +0.3)]
    match9 = re.match(r'^\[\s*.*?(-?\d+(?:\.\d+)?)\s*\(\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\)\s*\]$', standard)
    if match9:
        nominal = float(match9.group(1))
        lower = float(match9.group(2))
        upper = float(match9.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])

    # 10. ( Reff. 0 (0 ~ +0.3))
    match10 = re.match(r'^\(\s*.*?(-?\d+(?:\.\d+)?)\s*\(\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\)\s*\)$', standard)
    if match10:
        nominal = float(match10.group(1))
        lower = float(match10.group(2))
        upper = float(match10.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])

    # 11. 1 [0 ~ +0.5]
    match11 = re.match(r'^(-?\d+(?:\.\d+)?)\s*\[\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\]$', standard)
    if match11:
        nominal = float(match11.group(1))
        lower = float(match11.group(2))
        upper = float(match11.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])

    # 12. Ø7 [-0.3 ~ 0]
    match12 = re.match(r'^[Ø°]?\s*(-?\d+(?:\.\d+)?)\s*\[\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\]$', standard)
    if match12:
        nominal = float(match12.group(1))
        lower = float(match12.group(2))
        upper = float(match12.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])

    # 13. [Ø5.5 [-0.3 ~ 0]
    match13 = re.match(r'^\[\s*[Ø°]?\s*(\d+(?:\.\d+)?)\s*\[\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?).*$', standard)
    if match13:
        nominal = float(match13.group(1))
        lower = float(match13.group(2))
        upper = float(match13.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])
    
    # match14: Reff. 0 ( 0 ~ +0.5 )
    match14 = re.match(
        r'^.*?(-?\d+(?:\.\d+)?)\s*\(\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\)$',
        standard
    )
    if match14:
        nominal = float(match14.group(1))
        lower = float(match14.group(2))
        upper = float(match14.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])

    return pd.Series([None, None, None], index=["std_value", "std_min", "std_max"])

@st.cache_data(show_spinner="⚙️ Mengubah ke format final...")
def transform_to_final_format(df):
    df.columns = [col.strip().replace('\n', ' ').title() for col in df.columns]
    if "Control Method" not in df.columns:
        df["Control Method"] = np.nan
    df["Control Method"] = df["Control Method"].fillna(method="ffill")
    df = df.replace(to_replace=["", "nan", "None"], value=np.nan)

    dengan_ukur_keywords = ["caliper", "hg", "depth cal", "pitch dial", "rough. t", "hitung", "depth clp"]
    tanpa_ukur_keywords = [ 
        "visual", "pg", "snap g.", "visual & punch", "visual + kikir", "visual & kikir", 
        "machining test", "visual ( reff. master rough.)", "insp. jig", "finishing test"
    ]
    dengan_cmm_keywords = ["cmm"]

    final_rows = []
    current_section = ""
    last_valid_item = ""
    last_control_method = ""

    for idx, row in df.iterrows():
        if is_section_row(row):
            current_section = extract_section_title(row)
            continue
        if row.isna().all():
            continue

        raw_point_check = str(row.get("No.", "")).strip()
        capital_letters = ''.join([c for c in raw_point_check if c.isalpha() and c.isupper()])
        cleaned_point_check = ''.join([c for c in raw_point_check if not (c.isalpha() and c.isupper())]).strip()
        extracted_note = f"[{capital_letters}]" if capital_letters else ""

        original_note = str(row.get("Note", "")).strip() if pd.notna(row.get("Note")) else ""
        final_note = f"{original_note} {extracted_note}".strip() if extracted_note else original_note

        item_raw = row.get("Item", np.nan)
        extra = row.get("_2", np.nan)
        item = str(item_raw).strip() if pd.notna(item_raw) else ""
        extra = str(extra).strip() if pd.notna(extra) else ""
        if item:
            last_valid_item = item
        item_check = f"{item or last_valid_item} ({extra})" if extra else (item or last_valid_item)

        control_method_raw = row.get("Control Method", "")
        control_method = str(control_method_raw).strip() if pd.notna(control_method_raw) else ""
        if not control_method:
            control_method = last_control_method
        control_method_lower = control_method.lower()

        if any(keyword in control_method_lower for keyword in dengan_ukur_keywords):
            jenis_point = "Dengan Ukur"
        elif any(keyword in control_method_lower for keyword in tanpa_ukur_keywords):
            jenis_point = "Tanpa Ukur"
        elif any(keyword in control_method_lower for keyword in dengan_cmm_keywords):
            jenis_point = "Dengan CMM"
        else:
            jenis_point = "Lainnya"

        qtime_checked = st.session_state.get(f"qtime_{idx}", False)
        check100_checked = st.session_state.get(f"check100_{idx}", False)

        jenis_pengecekan = [
            v.strip()
            for source in ["Set Up", "Patrol"]
            for v in str(row.get(source, "")).split(",")
            if v.strip() and v.strip() != "-"
        ]

        # Deteksi huruf kapital spesial di item_check
        if jenis_point == "Tanpa Ukur":
            tokens = item_check.split()

            if "T" in tokens:
                jenis_pengecekan.append("Qtime")
                tokens.remove("T")

            if "C" in tokens:
                jenis_pengecekan.append("100%")
                tokens.remove("C")

            item_check = " ".join(tokens)

        if qtime_checked:
            jenis_pengecekan.append("Qtime")
        if check100_checked:
            jenis_pengecekan.append("100%")
        jenis_pengecekan = list(dict.fromkeys(jenis_pengecekan))

        final_row = {
            "section": current_section,
            "point_check": cleaned_point_check,
            "jenis_point": jenis_point,
            "catatan": final_note,
            "item_check": item_check,
            "control_method": control_method or last_control_method,
            "standard": str(row.get("Standard", "")).strip(),
            "jenis_pengecekan": jenis_pengecekan,
            "qtime": qtime_checked,
            "check_100": check100_checked,
        }
        final_rows.append(final_row)

    df_result = pd.DataFrame(final_rows)

    def propagate_f_marker(df_result):
        df_result["base_number"] = df_result["point_check"].str.extract(r"^(\d+)")
        grouped = df_result.groupby("base_number")
        for base, group in grouped:
            if group["catatan"].str.contains(r"\[F\]", na=False).any():
                for idx in group.index:
                    catatan = df_result.at[idx, "catatan"]
                    if "[F]" in catatan:
                        if not catatan.strip().startswith("[F]"):
                            tags = re.findall(r"\[[^\[\]]+\]", catatan)
                            tags = [t for t in tags if t != "[F]"]
                            df_result.at[idx, "catatan"] = "[F] " + " ".join(tags)
                    else:
                        df_result.at[idx, "catatan"] = "[F] " + catatan.strip()
        df_result.drop(columns=["base_number"], inplace=True)

    def apply_note_transformations(row):
        if row["jenis_point"] in ["Dengan Ukur", "Dengan CMM"]:
            row = copy_special_measurements_to_note(row)
        row = move_single_caps_to_note(row)
        row["catatan"] = normalize_note_tags(row["catatan"])
        return row

    df_result = df_result.apply(apply_note_transformations, axis=1)

    map_jenis2 = {}
    if "Item" in df.columns and "Patrol" in df.columns:
        df["Patrol"] = df["Patrol"].fillna(method="ffill") 
        for _, row2 in df.iterrows():
            item_key = str(row2.get("Item", "")).strip()
            extra = str(row2.get("_2", "")).strip()
            item_check_key = f"{item_key} ({extra})" if extra else item_key
            jenis2 = str(row2.get("Patrol", "")).strip()
            if item_check_key and jenis2:
                map_jenis2[item_check_key] = jenis2

    def update_jenis_pengecekan(row):
        current = row.get("jenis_pengecekan", [])
        if isinstance(current, str):
            current = [current] if current else []
        if not isinstance(current, list):
            current = list(current)
        tambahan = map_jenis2.get(row["item_check"])
        if tambahan and tambahan not in current:
            current.append(tambahan)
        return current

    df_result["jenis_pengecekan"] = df_result.apply(update_jenis_pengecekan, axis=1)

    def convert_roman_section_to_number(text):
        if pd.isna(text):
            return text
        match = re.match(r'^\s*([IVXLCDM]+)\s*\.\s+(.*)', str(text).strip(), re.IGNORECASE)
        if match:
            roman, title = match.groups()
            try:
                number = roman_to_int(roman)
                return f"{number}. {title.strip()}"
            except:
                return text
        return text

    df_result["section"] = df_result["section"].apply(convert_roman_section_to_number)

    def final_cleanup(df):
        def clean_catatan(row):
            catatan = str(row["catatan"]).strip()
            jenis = str(row["jenis_point"]).strip()
            has_tag = any(tag in catatan for tag in ["[F]", "[M]"])
            kosong = catatan == ""
            if jenis == "Tanpa Ukur":
                return "-" if kosong and not has_tag else catatan
            elif jenis in ["Dengan Ukur", "Dengan CMM"]:
                return "-" if kosong else catatan
            return catatan
        df["catatan"] = df.apply(clean_catatan, axis=1)
        cols_to_fill = [
            "section", "point_check", "jenis_point", "item_check",
            "control_method"
        ]
        for col in cols_to_fill:
            if col in df.columns:
                df[col] = df[col].replace(["", "nan", "None"], np.nan).fillna("-")
        
        if "jenis_pengecekan" in df.columns:
            df["jenis_pengecekan"] = df["jenis_pengecekan"].apply(
                lambda val: [v for v in val if v != "-"] if isinstance(val, list) else ["-"]
            )
            df["jenis_pengecekan"] = df["jenis_pengecekan"].apply(
                lambda val: val if val else ["-"]
            )

        return df

    def move_m_from_standard_to_note(row):
        std = str(row["standard"]).strip()
        catatan = str(row["catatan"]).strip()

        if re.search(r'(^|\s|\(|\[)M($|\s|[\]\),.])', std):
            if "[M]" not in catatan:
                catatan = f"{catatan} [M]".strip()
            std = re.sub(r'(^|\s|\(|\[)M($|\s|[\]\),.])', ' ', std).strip()

        row["standard"] = std
        row["catatan"] = catatan
        return row
    #ini
    from difflib import get_close_matches

    def find_incomplete_duplicates(df):
        all_items = df["item_check"].tolist()
        suspicious_indexes = []

        for idx, row in df.iterrows():
            item_check = row["item_check"]
            jenis_pengecekan = row.get("jenis_pengecekan", [])

            if isinstance(jenis_pengecekan, str):
                jenis_pengecekan = [v.strip() for v in jenis_pengecekan.split(",") if v.strip()]

            # Kriteria baris mencurigakan:
            is_suspect = (
                not any(char in item_check for char in "()")  # Tanpa (E..)
                and get_close_matches(item_check, all_items, n=2, cutoff=0.8)
                and (not jenis_pengecekan or jenis_pengecekan == ["-"] or "nan" in jenis_pengecekan)
            )

            if is_suspect:
                suspicious_indexes.append(idx)

        return suspicious_indexes

    # Tandai baris-baris mencurigakan
    suspicious_rows = find_incomplete_duplicates(df_result)

    df_result["status"] = df_result.get("status", "valid")  # kalau belum ada
    df_result.loc[suspicious_rows, "status"] = "duplikat_parsing"
    #ini

    df_result["jenis_point"] = df_result["jenis_point"].replace("Lainnya", np.nan).fillna(method="ffill")
    df_result["control_method"] = df_result["control_method"].replace(["", "nan", "None"], np.nan).fillna(method="ffill")
    df_result["standard"] = df_result["standard"].replace(["", "nan", "None"], np.nan).fillna(method="ffill")
    df_result["point_check"] = df_result["point_check"].replace(["", " ", "nan", "None"], np.nan).fillna(method="ffill")

    df_result = clean_empty_rows(df_result)
    df_result = merge_point_item(df_result)
    propagate_f_marker(df_result)

    parsed_std = df_result.apply(parse_standard_value, axis=1)
    df_result = df_result.apply(move_m_from_standard_to_note, axis=1)
    df_result.loc[df_result["jenis_point"].isin(["Dengan Ukur", "Dengan CMM"]), "standard"] = None
    mask_tanpa_ukur = df_result["jenis_point"] == "Tanpa Ukur"
    parsed_std.loc[mask_tanpa_ukur, ["std_value", "std_min", "std_max"]] = [None, None, None]
    df_result[["std_value", "std_min", "std_max"]] = parsed_std[["std_value", "std_min", "std_max"]]
    mask_diukur = df_result["jenis_point"].isin(["Dengan Ukur", "Dengan CMM"])
    df_result.loc[mask_diukur, "standard"] = None
    df_result = final_cleanup(df_result)

    return df_result

# ---------- Streamlit UI ----------

st.set_page_config(page_title="Check Sheet QFORM", layout="wide")
st.title("📄 CHECK SHEET SCAN QFORM")

uploaded_file = st.file_uploader("📤 Upload PDF file", type="pdf")

def update_final_result(edited_df):
    for idx, row in edited_df.iterrows():
        qtime_checked = row["qtime"]
        check100_checked = row["check_100"]

        jenis_raw = row["edit_jenis_pengecekan"]
        jenis_pengecekan = [v.strip() for v in str(jenis_raw).split(",") if v.strip()]
        jenis_pengecekan = [re.sub(r"\s+", " ", v.strip()) for v in jenis_pengecekan]
        if qtime_checked and "Qtime" not in jenis_pengecekan:
            jenis_pengecekan.append("Qtime")
        if check100_checked and "100%" not in jenis_pengecekan:
            jenis_pengecekan.append("100%")

        jenis_pengecekan = list(dict.fromkeys(jenis_pengecekan))
        edited_df.at[idx, "jenis_pengecekan"] = jenis_pengecekan

if uploaded_file:
    file_hash = get_file_hash(uploaded_file)

    if "last_file_hash" not in st.session_state or st.session_state.last_file_hash != file_hash:
        st.session_state.last_file_hash = file_hash
        for key in ["df_final_data", "temp_edit", "show_updated_table"]:
            st.session_state.pop(key, None)
        st.cache_data.clear()

    try:
        df = extract_table_from_pdf(uploaded_file)
        if df.empty:
            st.warning("❌ No tables detected in the PDF.")
        else:
            df_cleaned = bersihkan_dataframe(df.copy())
            st.subheader("🚀 Cleaned Table")
            st.dataframe(df_cleaned, use_container_width=True, hide_index=True)

            if "df_final_data" not in st.session_state:
                st.session_state.df_final_data = transform_to_final_format(df_cleaned)

            # 🚨 Validasi status data (duplikat vs valid)
            df_validasi = st.session_state.df_final_data
            total = len(df_validasi)
            duplikat_count = (df_validasi["status"] == "duplikat").sum()
            suspect_count = (df_validasi["status"] == "duplikat_parsing").sum()
            valid_count = total - duplikat_count - suspect_count

            st.info(f"""
            📊 **Validasi Baris**
            - Total: {total}
            - 🟢 Valid: {valid_count}
            - 🟡 Duplikat normal: {duplikat_count}
            - 🟠 Dugaan parsing rusak: {suspect_count}
            """)

            if suspect_count > 0:
                st.warning("⚠️ Ditemukan baris yang kemungkinan hasil merge/parsing tidak sempurna.")
                if st.checkbox("🔍 Lihat baris parsing mencurigakan"):
                    st.dataframe(df_validasi[df_validasi["status"] == "duplikat_parsing"])

            # Tampilkan hanya baris mencurigakan (duplikat atau parsing error)
            if "temp_edit" not in st.session_state:
                st.session_state.temp_edit = st.session_state.df_final_data.copy()
            if "show_updated_table" not in st.session_state:
                st.session_state.show_updated_table = False

            df_edit = st.session_state.temp_edit.copy()
            df_edit.reset_index(inplace=True)

            # Filter hanya baris mencurigakan
            rows_to_delete = df_edit[df_edit["status"].isin(["duplikat", "duplikat_parsing"])]

            if not rows_to_delete.empty:
                st.warning(f"⚠️ Ditemukan {len(rows_to_delete)} baris mencurigakan (duplikat/parsing error).")

                selected_rows = st.multiselect(
                    "🔍 Pilih index baris mencurigakan untuk dihapus:",
                    options=rows_to_delete["index"].tolist()
                )

                if st.button("🗑️ Hapus baris yang dipilih"):
                    df_cleaned = df_edit[~df_edit["index"].isin(selected_rows)].drop(columns=["index"])
                    st.session_state.temp_edit = df_cleaned.copy()
                    st.session_state.df_final_data = df_cleaned.copy()
                    st.success(f"✅ {len(selected_rows)} baris berhasil dihapus.")
                    st.rerun()
            # 🚨 Validasi status data (duplikat vs valid)

            st.subheader("📊 Edit Final Structured Format")

            temp_df = st.session_state.temp_edit.copy()
            temp_df.reset_index(drop=True, inplace=True)
            temp_df["edit_jenis_pengecekan"] = temp_df["jenis_pengecekan"].apply(
                lambda x: ", ".join(x) if isinstance(x, list) else str(x)
            )

            edited_df = st.data_editor(
                temp_df,
                use_container_width=True,
                num_rows="dynamic",
                hide_index=True,
                key="editable_table"
            )

            col1, col2 = st.columns([3, 1])
            with col1:
                if st.button("🔄 Update and Show Final"):
                    update_final_result(edited_df)
                    st.session_state.df_final_data = edited_df.copy()
                    st.session_state.temp_edit = edited_df.copy()
                    st.session_state.show_updated_table = True
                    st.rerun()

            if st.session_state.show_updated_table:
                df_view = st.session_state.df_final_data.drop(
                    columns=["qtime", "check_100", "edit_jenis_pengecekan"], errors="ignore"
                )
                st.subheader("👁️‍🗨️ Update View")
                st.dataframe(df_view, use_container_width=True, hide_index=True)

                if st.button("📤 Send to API"):
                    df_to_send = st.session_state.df_final_data.drop(
                        columns=["qtime", "check_100", "edit_jenis_pengecekan"], errors="ignore"
                    )
                    send_df_to_api(df_to_send)

        df_to_download = st.session_state.df_final_data.drop(
            columns=["qtime", "check_100", "edit_jenis_pengecekan"], errors="ignore"
        )
        excel_data = convert_df_to_excel(df_to_download)
        if excel_data:
            st.download_button(
                label="💾 Download Excel",
                data=excel_data,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            if st.button("♻️ Reset"):
                st.cache_data.clear()
                for key in ["df_final_data", "temp_edit", "show_updated_table", "last_file_hash"]:
                    st.session_state.pop(key, None)
                st.rerun()

    except Exception as e:
        st.error(f"🔥 Error while processing the file:\n\n{e}")
        st.text("📄 Traceback log:")
        st.text(traceback.format_exc())
import pdfplumber
import pandas as pd
import numpy as np
import re
try:
    import streamlit as st
except:
    st = None
from utils import (
    copy_special_measurements_to_note,
    parse_standard_value,
    append_cmm_summary_row,
    isi_label_abjad_di_antara,
    dengan_ukur_keywords,
    tanpa_ukur_keywords,
    dengan_cmm_keywords
)

# ===================================================================================================
# Extract table from PDF
# ===================================================================================================
def deduplicate_words(text):
    if not isinstance(text, str):
        return text
    words = text.strip().split()
    deduped = []
    for word in words:
        if not deduped or deduped[-1] != word:
            deduped.append(word)
    return " ".join(deduped)

def extract_table_from_pdf(file):
    all_dataframes = []
    max_columns = 0
    with pdfplumber.open(file) as pdf:
        for page_num, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            page_tables = []
            # Log ALL RAW SCAN --------------------------------------------------------------//
            # st.write(f"📄 Halaman {page_num + 1}")
            # for table_idx, table in enumerate(tables):
            #     st.write(f"  ➤ Tabel {table_idx + 1} - Jumlah baris: {len(table)}")
            #     for i, row in enumerate(table):
            #         st.write(f"    Row {i}: {row}")
            # Log ALL RAW SCAN --------------------------------------------------------------//
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
                    expected_cols = 13
                    header = [deduplicate_words(str(h).strip()) if h is not None else "" for h in data[0]]
                    while len(header) < expected_cols:
                        header.append(f"Extra_{len(header)}")
                    normalized_data = []
                    for row in data[1:]:
                        padded_row = row + [""] * (len(header) - len(row))
                        target_idx = 2
                        if not padded_row[target_idx] or str(padded_row[target_idx]).strip() == "":
                            for val in padded_row[5:10]:  
                                val_clean = str(val).strip()
                                if val_clean.isdigit():
                                    padded_row[target_idx] = val_clean
                                    break
                        normalized_data.append(padded_row)
                    seen = {}
                    new_header = []
                    for col in header:
                        if col in seen:
                            seen[col] += 1
                            new_header.append(f"{col}_{seen[col]}")
                        else:
                            seen[col] = 0
                            new_header.append(col)
                    df = pd.DataFrame(normalized_data, columns=new_header)
                    page_tables.append(df)
                else:
                    df = pd.DataFrame(table)
                    page_tables.append(df)
                df["page_number"] = page_num + 1
            if page_tables:
                try:
                    merged_df = pd.concat(page_tables, axis=0, ignore_index=True)
                    all_dataframes.append(merged_df)
                except Exception as e:
                    if st:
                        st.warning(f"⚠️ Gagal merge tabel di halaman {page_num+1}: {e}")
    normalized_tables = []
    for df in all_dataframes:
        if df.shape[1] < max_columns:
            for i in range(df.shape[1], max_columns):
                df[f"Extra_{i}"] = ""
        normalized_tables.append(df)
    return pd.concat(normalized_tables, ignore_index=True) if normalized_tables else pd.DataFrame()

# ===================================================================================================
# Cleaned Inspection Standard & Clean Footer
# ===================================================================================================
def parse_header_blocks(columns):
    blocks = []
    current_block = None
    for idx, col in enumerate(columns):
        lc = col.lower().strip()
        # kolom noise -> reset block
        if lc in ["", "none none none", "nan nan nan"]:
            blocks.append({
                "index": idx,
                "name": col,
                "block": None,
                "is_method": False,
                "is_frek": False
            })
            current_block = None
            continue
        if "verifikasi" in lc and "job" in lc and "method" in lc:
            current_block = "vjs"
        elif "insp" in lc and "normal" in lc and "method" in lc:
            current_block = "insp"
        elif "q time" in lc and "method" in lc:
            current_block = "qt"
        elif ("100%" in lc or "operator" in lc) and "method" in lc:
            current_block = "op"
        blocks.append({
            "index": idx,
            "name": col,
            "block": current_block,
            "is_method": "method" in lc,
            "is_frek": "frek" in lc
        })

    return blocks

def bersihkan_dataframe(df):
    if df.empty:
        return df
    df = hapus_footer(df)
    header_indices = []
    for i in range(len(df)):
        row_text = " ".join(str(x).lower() for x in df.iloc[i].values if pd.notnull(x))
        if "no" in row_text and "item" in row_text and "standard" in row_text:
            header_indices.append(i)
    if not header_indices:
        return df
    result_tables = []
    for idx, start_idx in enumerate(header_indices):
        end_idx = header_indices[idx + 1] if idx + 1 < len(header_indices) else len(df)
        sub_df = df.iloc[start_idx:end_idx].copy()
        header_rows_idx = []
        for j in range(len(sub_df)):
            row_text = " ".join(str(x).lower() for x in sub_df.iloc[j].values if pd.notnull(x))
            if "no" in row_text and "item" in row_text and not header_rows_idx:
                header_rows_idx.append(j)
            elif "method" in row_text and header_rows_idx:
                header_rows_idx.append(j)
                break
        if not header_rows_idx:
            continue
        start_h = header_rows_idx[0]
        end_h = header_rows_idx[-1]
        header_rows = sub_df.iloc[start_h:end_h + 1].values.tolist()
        max_len = max(len(r) for r in header_rows)
        header_rows = [list(r) + [""] * (max_len - len(r)) for r in header_rows]
        header_final = []
        for col_i in range(max_len):
            parts = [str(r[col_i]).strip() for r in header_rows if str(r[col_i]).strip()]
            joined = " ".join(parts)
            header_final.append(re.sub(r"\s+", " ", joined))
        data_df = sub_df.iloc[end_h + 1:].copy()
        data_df.columns = header_final[:len(data_df.columns)]
        def get_col(df, patterns):
            for c in df.columns:
                low = c.lower().replace(".", " ")
                if all(p in low for p in patterns):
                    return c
            return None
        col_no   = get_col(data_df, ["no"])
        col_item = get_col(data_df, ["item"])
        col_std  = get_col(data_df, ["standard"])
        col_vjs  = get_col(data_df, ["verifikasi", "job", "method"])
        col_qt   = get_col(data_df, ["q", "time", "method"])
        col_op = get_col(data_df, ["100", "method"]) or get_col(data_df, ["operator", "method"])

        # ----------------------------------------------------------------
        # Deteksi kolom Insp. Normal (Method & Frek)
        # ----------------------------------------------------------------
        col_insp_method = None
        col_insp_frek = None
        insp_method_idx = None
        insp_frek_idx = None
        cols = list(data_df.columns)
        for i, col in enumerate(cols):
            lc = col.lower().replace(".", " ")

            if "insp" in lc and "normal" in lc and "method" in lc:
                insp_method_idx = i
                col_insp_method = col

                if i + 1 < len(cols) and "frek" in cols[i + 1].lower():
                    insp_frek_idx = i + 1
                    col_insp_frek = cols[i + 1]
                break

        # mapping
        cols_map = {
            "Section": col_no,
            "No": col_no,
            "Item": col_item,
            "Standard": col_std,
            "Verifikasi Job Set Up (Method)": col_vjs,
            "Insp. Normal (Method)": insp_method_idx,
            "Insp. Normal (Frek)": insp_frek_idx,
            "Q Time (Method)": col_qt,
            "100% (Method)": col_op,
        }
        def get_col_by_index(df, idx):
            if idx is None:
                return None
            return df.iloc[:, idx]
        clean_df = pd.DataFrame()
        for nice_name, raw_col in cols_map.items():
            # CASE 1: pakai index (Insp Normal)
            if isinstance(raw_col, int):
                clean_df[nice_name] = get_col_by_index(data_df, raw_col)
                continue
            # CASE 2: pakai nama kolom (string)
            if isinstance(raw_col, str):
                matched = [
                    c for c in data_df.columns
                    if c == raw_col or c.startswith(raw_col + "_")
                ]
                if matched:
                    col_data = data_df[matched[0]]
                    if isinstance(col_data, pd.DataFrame):
                        col_data = col_data.iloc[:, 0]
                    clean_df[nice_name] = col_data
                else:
                    clean_df[nice_name] = np.nan
                continue
            # CASE 3: None / kosong
            clean_df[nice_name] = ""

        # clean null row
        clean_df = clean_df.replace(r'^\s*$', np.nan, regex=True)
        clean_df = clean_df.dropna(how="all").reset_index(drop=True)
        for c in ["Section","No", "Item", "Standard"]:
            clean_df[c] = clean_df[c].fillna(method="ffill")

        # --- SINKRONIZE INSP NORMAL METHOD & FREK ---
        method_col = "Insp. Normal (Method)"
        frek_col   = "Insp. Normal (Frek)"
        if method_col in clean_df.columns and frek_col in clean_df.columns:
            last_frek = None
            for i in range(len(clean_df)):
                method_val = clean_df.at[i, method_col]
                frek_val   = clean_df.at[i, frek_col]
                # kalau method kosong, STOP isi frek
                if pd.isna(method_val) or str(method_val).strip() == "":
                    last_frek = None
                    continue
                # kalau frek ada, simpan
                if not pd.isna(frek_val) and str(frek_val).strip() != "":
                    last_frek = frek_val
                else:
                    # method ada, frek kosong → isi dari atas
                    clean_df.at[i, frek_col] = last_frek

        clean_df = merge_point_item(clean_df)
        clean_df = tambah_section_nomor(clean_df)
        clean_df = clean_df.reset_index(drop=True)
        last_section = "-"
        for r in range(len(clean_df)):
            val = str(clean_df.loc[r, "Section"]).strip()
            if val and val.lower() not in ["nan", "none", "-"]:
                last_section = val
            clean_df.loc[r, "Section"] = last_section
        clean_df = isi_label_abjad_di_antara(clean_df, kolom='Item', No='No')
        clean_df = extract_single_caps_as_note(clean_df, ...)
        
        # Jobsetup fallback -- fill up
        control_cols = ["Verifikasi Job Set Up (Method)"]
        for cc in control_cols:
            if cc in clean_df.columns:
                clean_df[cc] = clean_df[cc].replace(["", "None", None, np.nan], np.nan)
                clean_df[cc] = clean_df[cc].fillna(method="ffill")
        result_tables.append(clean_df)
        print(list(data_df.columns))
        print(col_insp_method, col_insp_frek)
    return pd.concat(result_tables, ignore_index=True) if result_tables else df

# Section Number --------------------------------------------------------------------------------------------------//
def tambah_section_nomor(df):
    df = df.copy()
    SECTION_REGEX = r'^\s*(I{1,3}|IV|V|VI{0,3}|VII{0,3}|VIII|IX|X)\s*[\.\-–]\s*(.+)$'
    current_section = None
    current_sharp = None
    section_from_sharp = False  
    new_rows = []
    for i in range(len(df)):
        row = df.iloc[i].copy()
        item_val = str(row["Item"]).strip()
        first_col = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        match = re.match(SECTION_REGEX, first_col, re.IGNORECASE)
        if match:
            roman = match.group(1)
            title = match.group(2).strip()
            current_section = f"{roman_to_int(roman)}. {title}"
            current_sharp = None
            section_from_sharp = False 
            row["Section"] = current_section
            new_rows.append(row)
            continue
        if "#" in item_val:
            sharp_clean = item_val.split("#", 1)[1].strip()
            # CASE A — tidak ada SECTION sama sekali
            if current_section is None:
                current_section = sharp_clean
                current_sharp = None
                section_from_sharp = True 
                continue
            # CASE B — ada SECTION asli sebelumnya (roman numeral)
            if not section_from_sharp:
                current_sharp = sharp_clean
                continue
            # CASE C — section dari #, ketemu # lagi → GANTI SECTION
            current_section = sharp_clean
            current_sharp = None
            section_from_sharp = True
            continue
        # 3) Default
        row = df.iloc[i].copy()
        if current_sharp:
            row["Section"] = f"{current_section}#{current_sharp}"
        else:
            row["Section"] = current_section if current_section else "-"
        new_rows.append(row)
    if not new_rows:
        df["Section"] = "-"
        return df
    return pd.DataFrame(new_rows)

def roman_to_int(roman):
    roman = roman.upper()
    roman_dict = {'I': 1, 'V': 5, 'X': 10, 'L': 50, 'C': 100}
    result, prev = 0, 0
    for char in reversed(roman):
        val = roman_dict.get(char, 0)
        result += val if val >= prev else -val
        prev = val
    return result

# [F], [Q] ----------------------------------------------------------------------------------------------------//
# still error : 2f F --> remark not appear
def extract_single_caps_as_note(df, target_caps=None, source_cols=None, remark_col="remark"):
    if target_caps is None:
        target_caps = ["F", "Q"]
    if source_cols is None:
        source_cols = ["No", "Item", "Standard"]
    standalone = r'(?<![A-Za-z0-9])([FQ])(?![A-Za-z0-9])'
    numeric_caps = r'(\d+)([FQ])'
    num_cap_suffix = r'(\d+)\s*([FQ])\s*([a-zA-Z])'
    def process_row(row):
        # skip if "to F" or "to Q"
        text_all = " ".join(str(row.get(c, "") or "") for c in source_cols).lower()
        if re.search(r'\bto\s+[fq]\b', text_all):
            # pastikan remark kosong = "-"
            current = row.get(remark_col, "")
            if current is None or str(current).strip() == "":
                row[remark_col] = "-"
            return row
        remark = str(row.get(remark_col, "") or "").strip()
        for col in source_cols:
            text = str(row.get(col, "") or "").strip()
            found = re.findall(standalone, text)
            for cap in found:
                tag = f"[{cap}]"
                if tag not in remark:
                    remark = (remark + " " + tag).strip()
            # CASE: 2 Fa / 2Fb / 2 F a
            found3 = re.findall(num_cap_suffix, text)
            for num, cap, suf in found3:
                tag = f"[{cap}]"
                if tag not in remark:
                    remark = (remark + " " + tag).strip()
                if col == "No":
                    text = f"{num}{suf.lower()}"
            cleaned = re.sub(standalone, " ", text)
            cleaned = re.sub(r"\s+", " ", cleaned).strip()
            row[col] = cleaned
            row[remark_col] = remark
        return row
    return df.apply(process_row, axis=1)

# Merge point number to 1a, 1b, etc ------------------------------------------------------------------------------//
def merge_point_item(df):
    if df.empty:
        return df
    col_no = None
    col_item = None
    for c in df.columns:
        lc = str(c).lower()
        if "no" == lc.strip():
            col_no = c
        elif "item" in lc:
            col_item = c
    if not col_no or not col_item:
        return df  
    new_no = []
    new_item = []
    for no, item in zip(df[col_no], df[col_item]):
        no_str = str(no).strip() if pd.notna(no) else ""
        item_str = str(item).strip() if pd.notna(item) else ""
        match = re.match(r"^([a-zA-Z])[\.\)]\s*(.*)", item_str)
        if match:
            letter = match.group(1).lower()
            remaining_item = match.group(2).strip()
            combined_no = f"{no_str}{letter}"
            new_no.append(combined_no)
            new_item.append(remaining_item)
        else:
            new_no.append(no_str)
            new_item.append(item_str)
    df[col_no] = new_no
    df[col_item] = new_item
    return df

# Footer -------------------------------------------------------------------------------------------------------//
def hapus_footer(df):
    if df.empty:
        return df
    # Keyword footer
    footer_signals = ["dibuat", "instruksi kerja", "inspection standard", "berlaku mulai", "revision", "APPLIED MODEL"]
    drop_indexes = []
    for idx in df.index:
        row_text = " ".join(str(x).lower() for x in df.loc[idx].values if pd.notnull(x))
        if any(sig in row_text for sig in footer_signals):
            drop_indexes.append(idx)
    if drop_indexes:
        if st:
            st.info(f"🧹 Menghapus {len(drop_indexes)} baris footer")
        df = df.drop(index=drop_indexes)
    df = df.replace(r"^\s*$", np.nan, regex=True).dropna(how="all")
    df = df.reset_index(drop=True)
    return df

# ===================================================================================================
# Transform to Final Format
# ===================================================================================================
FREK_TO_JENIS_INSP = {
    "Patrol 1x/Shift": [
        "1x/shift", "shift/1x", "1 shift", "x1/shift", "1x shift", "x1 / shift", "shift x1",
        "tfihs / x1", "x1 / tfihs", "tfihs/1x", "1x / tfihs", "tfihs x1",
        "tfihs / x1 tfihs / x1", "tfihs / x1 tfihs / x1 tfihs / x1", "1x / shift"
    ],
    "Patrol 1x/Day": [
        "1x/day", "day/1x", "1 day", "1x per day", "per day", "yad/x1", "x1/yad", "1x   day", "yad scp 1",
        "yad   scp 1",  "1x / day",
    ]
}
METHOD_TO_JENIS = {
    "Verifikasi Job Set Up (Method)": "Job Setup",
    "Q Time (Method)": "Q-Time",
    "100% (Method)": "Check 100%",
}

def normalize_frek(val):
    if val is None:
        return ""
    s = str(val).lower()
    s = re.sub(r"\s+", " ", s)
    return s.strip()
def detect_insp_jenis_from_frek(frek):
    if frek is None or pd.isna(frek):
        return None
    # hapus spasi, dash, underscore, lowercase
    f = str(frek).lower()
    f = re.sub(r"[\s\-_]", "", f)
    for jenis, variants in FREK_TO_JENIS_INSP.items():
        for v in variants:
            v_norm = re.sub(r"[\s\-_]", "", v.lower())
            if v_norm and v_norm in f:
                return jenis
    return None

def handle_control_method(row): # Control method & Jenis Pengecekan -------------------------------//
    main_methods = []
    hundred_methods = []
    job_setup_methods = []
    for col, jenis in METHOD_TO_JENIS.items():
        if col not in row.index:
            continue
        val = row[col]
        if pd.isna(val) or val == "":
            continue
        if jenis == "Check 100%":
            hundred_methods.append(val)
        elif jenis == "Job Setup":
            job_setup_methods.append(val)
        else:
            main_methods.append(val)
    hasil = []
    def normalize_list(values):
        norm = {}
        for v in values:
            key = v.strip().lower()
            if key not in norm:
                norm[key] = v.strip()
        return norm
    normalized_main = normalize_list(main_methods)
    normalized_100  = normalize_list(hundred_methods)

    # Case A: ada main & 100% dgn method yang sama
    same_keys = set(normalized_main.keys()) & set(normalized_100.keys())
    if same_keys:
        for key in same_keys:
            merged_row = row.copy()
            merged_row["control_method"] = normalized_main[key]
            merged_row["jenis_pengecekan"] = ["Q-Time", "Check 100%"]
            hasil.append(merged_row)
        for k in same_keys:
            normalized_main.pop(k, None)
            normalized_100.pop(k, None)

    # Case B: main (Q-Time) tersisa
    if normalized_main:
        if len(normalized_main) == 1:
            main_row = row.copy()
            main_row["control_method"] = list(normalized_main.values())[0]
            main_row["jenis_pengecekan"] = ["Q-Time"]
            hasil.append(main_row)
        else:
            for method in normalized_main.values():
                main_row = row.copy()
                main_row["control_method"] = method
                main_row["jenis_pengecekan"] = ["Q-Time"]
                hasil.append(main_row)

    # Case C: 100% tersisa
    if normalized_100:
        for method in normalized_100.values():
            new_row = row.copy()
            new_row["control_method"] = method
            new_row["jenis_pengecekan"] = ["Check 100%"]
            hasil.append(new_row)
            
    # Case INSP: Insp. Normal berdasarkan frek
    insp_method = row.get("Insp. Normal (Method)")
    insp_frek = row.get("Insp. Normal (Frek)")
    if pd.notna(insp_method) and str(insp_method).strip() != "":
        m_str = str(insp_method).strip()
        jenis_insp = detect_insp_jenis_from_frek(insp_frek) or "Patrol"
        found_insp = False
        for h in hasil:
            if h["control_method"] == m_str:
                if jenis_insp not in h["jenis_pengecekan"]:
                    h["jenis_pengecekan"].append(jenis_insp)
                found_insp = True
                break
        if not found_insp:
            insp_row = row.copy()
            insp_row["control_method"] = m_str
            insp_row["jenis_pengecekan"] = [jenis_insp]
            hasil.append(insp_row)

    # Case Job Setup
    if job_setup_methods:
        norm_js = normalize_list(job_setup_methods)
        for method in norm_js.values():
            # cek apakah ada row dengan same control_method
            found = False
            for h in hasil:
                if h["control_method"] == method or h.get("_is_insp"):
                    if "Job Setup" not in h["jenis_pengecekan"]:
                        h["jenis_pengecekan"].append("Job Setup")
                    found = True
                    break
            if not found:
                js_row = row.copy()
                js_row["control_method"] = method
                js_row["jenis_pengecekan"] = ["Job Setup"]
                hasil.append(js_row)
    return hasil

# Validasi For Final Format ---------------------------------------------------------------------------//
def find_exact_duplicates(df): # for duplicated data
    duplicate_indexes = []
    seen_keys = set()
    key_columns = ["section", "point_check", "jenis_point", "catatan","item_check", "control_method", "std_value", "std_min", "std_max"]
    for idx, row in df.iterrows():
        row_key = tuple(
            str(row.get(col, "")).strip().lower()
            .replace(", ", ",")
            .replace(" ,", ",")
            .replace("\n", " ")
            if not pd.isna(row.get(col)) else ""
            for col in key_columns
        )
        if row_key in seen_keys:
            duplicate_indexes.append(idx)
        else:
            seen_keys.add(row_key)
    return duplicate_indexes

#  Main Fitur ----------------------------------------------------------------------------------------------------//
def transform_to_final_format(df):
    df.columns = [col.strip().replace('\n', ' ').title() for col in df.columns]
    # Jenis Point Check 
    if "Control Method" not in df.columns:
        df["Control Method"] = np.nan

    # FINAL FORMAT AWAL
    final = pd.DataFrame()
    final["section"] = df["Section"]
    final["point_check"] = df["No"]
    final["jenis_point"] = ""
    final["catatan"] = df.get("Remark", "")
    final["item_check"] = df["Item"]
    final["control_method"] = ""
    final["standard"] = df["Standard"]
    final["std_value"] = ""
    final["std_min"] = ""
    final["std_max"] = ""
    final["status"] = "valid"

    for col in METHOD_TO_JENIS.keys():
        if col in df.columns:
            final[col] = df[col]
        else:
            final[col] = ""
            
    # --- TAMBAHAN: Copy Kolom Inspeksi ke Final ---
    insp_cols = ["Insp. Normal (Method)", "Insp. Normal (Frek)"]
    for col in insp_cols:
        if col in df.columns:
            final[col] = df[col]
        else:
            final[col] = ""

    # --- FIX SECTION FILL "-" ---
    final["section"] = final["section"].replace([None, np.nan], "")
    last_valid = None
    fixed_sections = []
    for val in final["section"]:
        val_str = str(val).strip()
        if val_str not in ["", "-"]:
            last_valid = val_str
            fixed_sections.append(val_str)
        else:
            if last_valid is not None:
                fixed_sections.append(last_valid)
            else:
                fixed_sections.append("-")
    final["section"] = fixed_sections

    expanded = []
    for _, r in final.iterrows():
        expanded.extend(handle_control_method(r))  
    final = pd.DataFrame(expanded)

    cols_to_remove = ["Insp. Normal (Method)", "Insp. Normal (Frek)", "_is_insp"]
    final.drop(columns=[c for c in cols_to_remove if c in final.columns], inplace=True)

    # dengan ukur, tanpa ukur, dengan cmm
    def tentukan_jenis_point(method, raw_methods): 
        def norm(s):
            s = str(s or "")
            s = s.lower()
            s = re.sub(r"\s+", " ", s)
            return s.strip()
        method_lower = norm(method)
        _dengan_ukur = [norm(k) for k in dengan_ukur_keywords]
        _tanpa_ukur = [norm(k) for k in tanpa_ukur_keywords]
        _dengan_cmm = [norm(k) for k in dengan_cmm_keywords]
        if any(k in method_lower for k in _dengan_cmm):
            return "Dengan CMM"
        if any(k in method_lower for k in _dengan_ukur):
            return "Dengan Ukur"
        if any(k in method_lower for k in _tanpa_ukur):
            return "Tanpa Ukur"
        
        for col in METHOD_TO_JENIS.keys():
            val = norm(raw_methods.get(col, ""))
            if any(k in val for k in _dengan_cmm):
                return "Dengan CMM"
            if any(k in val for k in _dengan_ukur):
                return "Dengan Ukur"
            if any(k in val for k in _tanpa_ukur):
                return "Tanpa Ukur"
        return "Tanpa Ukur"
    
    final["jenis_point"] = final.apply(lambda r: tentukan_jenis_point(r["control_method"], r), axis=1)

    # COPY STANDARD to CATATAN 
    final = final.apply(copy_special_measurements_to_note, axis=1)
    for col in METHOD_TO_JENIS.keys():
        if col in final.columns:
            final.drop(columns=[col], inplace=True)

    # PARSE STANDARD to STD VALUE / STD MIN / STD MAX
    def parse_std_for_row(row):
        if row["jenis_point"] in ["Dengan Ukur", "Dengan CMM"]:
            parsed = parse_standard_value(row)
            row["std_value"] = parsed["std_value"]
            row["std_min"] = parsed["std_min"]
            row["std_max"] = parsed["std_max"]
            # delete std
            row["standard"] = None
        else:
            row["std_value"] = None
            row["std_min"] = None
            row["std_max"] = None
        return row

    final = final.apply(parse_std_for_row, axis=1)
    final["catatan"] = final["catatan"].apply(lambda x: "-" if pd.isna(x) or str(x).strip() == "" else x)
    final = append_cmm_summary_row(final)

    # Validasi ( Footer bocor )--------------------------------------------------------//
    def is_valid_point_check(x):
        if x is None or str(x).strip() == "":
            return False
        s = str(x).strip()
        if re.fullmatch(r'\d+[a-zA-Z]?$', s):
            return True
        return False
    final = final[final["point_check"].apply(is_valid_point_check)].reset_index(drop=True)
    # Validasi duplikat ===============
    final["status"] = "valid"
    duplicate_rows = find_exact_duplicates(final)
    final.loc[duplicate_rows, "status"] = "duplikat"
    # Hapus otomatis 
    final = final[final["status"] == "valid"].reset_index(drop=True)

    return final

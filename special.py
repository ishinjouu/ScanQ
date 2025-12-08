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
    append_cmm_summary_row
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

        # Fallback for - Incoming
        col_insp_method = None
        col_insp_frek = None
        insp_cols = [c for c in data_df.columns if "insp" in c.lower() and "normal" in c.lower()]
        if insp_cols:
            for c in insp_cols:
                cl = c.lower()
                if "method" in cl and "frek" not in cl:
                    col_insp_method = c
                elif "frek" in cl and "method" not in cl:
                    col_insp_frek = c

            if col_insp_method and not col_insp_frek:
                right_index = list(data_df.columns).index(col_insp_method) + 1
                if right_index < len(data_df.columns):
                    right_name = data_df.columns[right_index].lower()
                    if "frek" in right_name:
                        col_insp_frek = data_df.columns[right_index]
            elif not col_insp_method and insp_cols:
                if re.search(r"method.*frek|frek.*method", insp_cols[0].lower()):
                    col_insp_method = insp_cols[0]
                    col_insp_frek = insp_cols[0]

        # fallback if fail detection
        if not col_vjs:
            col_vjs = get_col(data_df, ["incoming"]) or get_col(data_df, ["job"])
        if not col_qt:
            col_qt = get_col(data_df, ["qtime"]) or get_col(data_df, ["qt"])
        if not col_op:
            col_op = get_col(data_df, ["100"]) or get_col(data_df, ["100%"])
        if not col_insp_method:
            col_insp_method = get_col(data_df, ["patrol"]) or get_col(data_df, ["normal"])

        # =============================================
        # FALLBACK KHUSUS UNTUK FORMAT INCOMING (2)
        # =============================================
        incoming2_kw = ["patrol", "out going", "incoming", "q time", "100%"]
        is_incoming2 = any(
            any(k in c.lower() for k in incoming2_kw)
            for c in data_df.columns
        )
        if is_incoming2:
            for c in data_df.columns:
                cl = c.lower()
                if "patrol" in cl and "method" in cl:
                    col_vjs = c     # Verifikasi Job Setup (Method)
                elif ("out going" in cl or "incoming" in cl) and "method" in cl:
                    col_insp_method = c   # Insp Normal (Method)
                elif "q time" in cl and "method" in cl:
                    col_qt = c     # Q Time (Method)
                elif "100%" in cl and "method" in cl:
                    col_op = c     # 100% (Method)
        # ===============================================

        insp_cols = [c for c in data_df.columns if "insp" in c.lower() and "normal" in c.lower()]
        col_insp_method = None
        col_insp_frek = None
        if insp_cols:
            for c in insp_cols:
                cl = c.lower()
                if "method" in cl and "frek" not in cl:
                    col_insp_method = c
                elif "frek" in cl and "method" not in cl:
                    col_insp_frek = c

            # fallback: kalau cuma ada 1 kolom tapi headernya gabung “Method Frek.”
            if col_insp_method and not col_insp_frek:
                right_index = list(data_df.columns).index(col_insp_method) + 1
                if right_index < len(data_df.columns):
                    right_name = data_df.columns[right_index].lower()
                    if "frek" in right_name:
                        col_insp_frek = data_df.columns[right_index]
            elif not col_insp_method and insp_cols:
                if re.search(r"method.*frek|frek.*method", insp_cols[0].lower()):
                    col_insp_method = insp_cols[0]
                    col_insp_frek = insp_cols[0]

        # mapping
        cols_map = {
            "Section": col_no,
            "No": col_no,
            "Item": col_item,
            "Standard": col_std,
            "Verifikasi Job Set Up (Method)": col_vjs,
            "Insp. Normal (Method)": col_insp_method,
            "Q Time (Method)": col_qt,
            "100% (Method)": col_op,
        }

        clean_df = pd.DataFrame()
        for nice_name, raw_col in cols_map.items():
            if raw_col:
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
            else:
                clean_df[nice_name] = ""

        # clean null row
        clean_df = clean_df.dropna(how="all").reset_index(drop=True)
        for c in ["Section","No", "Item", "Standard"]:
            clean_df[c] = clean_df[c].fillna(method="ffill")

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

# label bolong b. dll ------------------------------------------------------------//
# ga error tapi di case Sumbu Ø51, no nya masih harus di edit manual
def isi_label_abjad_di_antara(df, kolom='Item', No='No'):
    pola = re.compile(r'^([a-zA-Z]*)(\d+)\s*(.*)$') 
    new_items = []
    last_prefix = None
    last_number = None
    df[No] = df[No].fillna('').astype(str)

    # isi Item dengan prefix lengkap
    for i in range(len(df)):
        no_val = str(df.iloc[i][No]).strip()
        item = str(df.iloc[i][kolom]).strip()
        match_no = re.match(r'^([a-zA-Z]*)(\d+)', no_val)
        number = match_no.group(2) if match_no else None
        prefix = match_no.group(1) if match_no else ''
        match_item = pola.match(item)
        if match_item:
            last_prefix = match_item.group(1) or prefix
            last_number = match_item.group(2) or number
            new_items.append(item)
        else:
            if number == last_number and last_prefix:
                new_items.append(f"{last_number}{last_prefix} {item}")
            else:
                new_items.append(item)
    df[kolom] = new_items

    # isi No kosong → ikut last no
    last_full_no = None
    for i in range(len(df)):
        no_val = str(df.iloc[i][No]).strip()
        if no_val:
            last_full_no = no_val
        else:
            if last_full_no:
                df.iloc[i, df.columns.get_loc(No)] = last_full_no
    for i in range(1, len(df)):
        now_no = str(df.iloc[i][No]).strip()
        prev_no = str(df.iloc[i-1][No]).strip()
        if re.match(r'^\d+$', now_no):
            prev_match = re.match(r'^(\d+)([a-zA-Z]+)$', prev_no)
            if prev_match and prev_match.group(1) == now_no:
                suffix = prev_match.group(2)
                df.iloc[i, df.columns.get_loc(No)] = f"{now_no}{suffix}"

    # Sumbu X/Y/Z
    pola_xyz = re.compile(r'^[XYZxyz]\b.*$')
    last_full_item = None
    for i in range(len(df)):
        item_val = str(df.iloc[i][kolom]).strip()
        if "sumbu" in item_val.lower():
            last_full_item = item_val
        elif pola_xyz.match(item_val) and last_full_item:
            head = last_full_item.split()[0]
            df.iloc[i, df.columns.get_loc(kolom)] = f"{head} {item_val}"

    return df

# [F], [Q] ----------------------------------------------------------------------------------------------------//
# still error : 2f F --> remark not appear
def extract_single_caps_as_note(df, target_caps=None, source_cols=None, remark_col="remark"):
    if target_caps is None:
        target_caps = ["F", "Q"]
    if source_cols is None:
        source_cols = ["No", "Item", "Standard"]
    standalone = r'(?<![A-Za-z0-9])([FQ])(?![A-Za-z0-9])'
    numeric_caps = r'(\d+)([FQ])'

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
            found2 = re.findall(numeric_caps, text)
            for num, cap in found2:
                tag = f"[{cap}]"
                if tag not in remark:
                    remark = (remark + " " + tag).strip()
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
METHOD_TO_JENIS = {
    "Verifikasi Job Set Up (Method)": "QTime",
    "Insp. Normal (Method)": "QTime",
    "Q Time (Method)": "QTime",
    "100% (Method)": "100%",
}

def handle_control_method(row): # Control method & Jenis Pengecekan -------------------------------//
    main_methods = []
    hundred_methods = []
    for col, jenis in METHOD_TO_JENIS.items():
        if col not in row.index:
            continue
        val = row[col]
        if pd.isna(val) or val == "":
            continue
        if jenis == "100%":
            hundred_methods.append(val)
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
            merged_row["jenis_pengecekan"] = ["QTime", "100%"]
            hasil.append(merged_row)
        for k in same_keys:
            normalized_main.pop(k, None)
            normalized_100.pop(k, None)

    # Case B: main (QTime) tersisa
    if normalized_main:
        if len(normalized_main) == 1:
            main_row = row.copy()
            main_row["control_method"] = list(normalized_main.values())[0]
            main_row["jenis_pengecekan"] = ["QTime"]
            hasil.append(main_row)
        else:
            for method in normalized_main.values():
                main_row = row.copy()
                main_row["control_method"] = method
                main_row["jenis_pengecekan"] = ["QTime"]
                hasil.append(main_row)

    # Case C: 100% tersisa
    if normalized_100:
        for method in normalized_100.values():
            new_row = row.copy()
            new_row["control_method"] = method
            new_row["jenis_pengecekan"] = ["100%"]
            hasil.append(new_row)

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
    dengan_ukur_keywords = [
        "caliper", "hg", "depth cal", "pitch dial", "torque dial", "depth cal.",
        "rough. t", "hitung", "depth clp", "height g", "dial g", 
        "thickness meter", "thicknes s meter", "thicknes meter",
    ]
    tanpa_ukur_keywords = [
        "visual", "pg", "snap g.", "visual & punch", "visual + kikir", "Putar dg tangan",
        "visual & kikir", "machining test", "visual ( reff. master rough.)", "Poka Yoke test",
        "insp. jig", "finishing test", "Visual + tangan", "Leak tester", "Torque M"
    ]
    dengan_cmm_keywords = ["cmm"]

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
    
    final["jenis_point"] = final.apply(
        lambda r: tentukan_jenis_point(r["control_method"], r),
        axis=1
    )

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

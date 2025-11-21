import pandas as pd
import re
from difflib import SequenceMatcher

# catatan - tambahan dari standard ke kolom catatan
def copy_special_measurements_to_note(row):
    standard = str(row.get("standard", "")).strip()
    catatan = str(row.get("catatan", "")).strip()
    jenis_point = str(row.get("jenis_point", "")).strip()
    if jenis_point not in ["Dengan Ukur", "Dengan CMM"]:
        return row

    # min_max_match = re.search(r"\b([Mm]in|[Mm]ax)\s*\d+(?:\.\d+)?", standard)
    # if min_max_match:
    #     tag = f"[{min_max_match.group(0).strip()}]"
    #     if tag not in catatan:
    #         catatan += f" {tag}"
    # === Min/Max biasa tanpa simbol ===
    min_max_match = re.search(r"\b([Mm]in|[Mm]ax)\.?\s*[Ø°μ]?\s*\d+(?:\.\d+)?[A-Za-z]?", standard)
    if min_max_match:
        tag = f"[{min_max_match.group(0).strip()}]"
        if tag not in catatan:
            catatan += f" {tag}"

    size_patterns = [
        r"[Ø°]\d+(?:\.\d+)?\s*±\s*[+−-]?\d+(?:\.\d+)?",                                     # Ø10 ±0.1 atau °10 ± 0.5
        r"[Ø°]\d+(?:\.\d+)?\s*\(\s*\d+(?:\.\d+)?\s*~\s*[+−-]?\d+(?:\.\d+)?\s*\)",           # Ø6.1 ( 0 ~ +0.1 )
        r"[Ø°]\d+(?:\.\d+)?\s*\(\s*[-+]?\d+(?:\.\d+)?\s*~\s*[-+]?\d+(?:\.\d+)?\s*\)",       # Ø6.1 (0 ~ +0.1) / Ø40 (-0.050 ~ -0.035)
        r"[Ø°]\d+(?:\.\d+)?\s*\(\s*[+−-]?\d+(?:\.\d+)?\s*~\s*[+−-]?\d+(?:\.\d+)?\s*\)",     # Ø12.15 (-0.15 ~ +0.25)
        r"\d+(?:\.\d+)?º\s*±\s*\d+(?:\.\d+)?º",                                             # 15º ± 3º
                r"\d{1,3}[°º]\s*±\s*\d{1,3}[°º]\s*\d{1,2}['′`´]"                            # 32° ± 1°30', 3
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

# --- isi kosong di kolom catatan berdasarkan grup point_check dan item_check ---
def fill_empty_catatan_from_group(df):
    df = df.copy()

    for idx, row in df.iterrows():
        if row.get("jenis_point") not in ["Dengan Ukur", "Dengan CMM"]:
            continue
        if row.get('catatan') and row['catatan'].strip() not in ["-", ""]:
            continue

        item = row['item_check']
        point_prefix = re.match(r'^(\d+[a-zA-Z]*)', str(row['point_check']))
        point_prefix = point_prefix.group(1) if point_prefix else ""

        for j, ref_row in df.iterrows():
            if j == idx:
                continue
            if ref_row.get("jenis_point") not in ["Dengan Ukur", "Dengan CMM"]:
                continue

            ref_prefix = re.match(r'^(\d+[a-zA-Z]*)', str(ref_row['point_check']))
            ref_prefix = ref_prefix.group(1) if ref_prefix else ""

            if (
                ref_prefix == point_prefix and
                isinstance(ref_row['item_check'], str)
            ):
                sim = SequenceMatcher(None, ref_row['item_check'], item).ratio()
                if sim >= 0.8 and ref_row['catatan']:
                    df.at[idx, 'catatan'] = ref_row['catatan']
                    break

    return df

# Format standard ke std_value, std_min, std_max -------------------------------------------------------------------------
def parse_standard_value(row):
    standard = str(row.get("standard", "")).strip()
    jenis_point = str(row.get("jenis_point", "")).strip().lower()

    # Ganti koma dengan titik jika jenis_point adalah "dengan ukur" atau "dengan cmm"
    if jenis_point in ["dengan ukur", "dengan cmm"]:
        standard = standard.replace(",", ".")

    if jenis_point == "Tanpa Ukur":
        return pd.Series([None, None, None], index=["std_value", "std_min", "std_max"])

    # 1. 0 ±0.2
    match1 = re.search(r'(\d+(?:\.\d+)?)\s*±\s*(\d+(?:\.\d+)?)', standard)
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
    # match5 = re.search(r'\b(Min|Max)\s*(\d+(?:\.\d+)?)\b', standard, re.IGNORECASE)
    # if match5:
    #     kind = match5.group(1).lower()
    #     value = float(match5.group(2))
    #     return pd.Series([0, value if kind == "min" else 0, value if kind == "max" else 0], index=["std_value", "std_min", "std_max"])
    match5 = re.search(r'\b(Min|Max)\.?\s*(\d+(?:\.\d+)?)', standard, re.IGNORECASE)
    if match5:
        kind = match5.group(1).lower()
        value = float(match5.group(2))
        if kind == "min":
            return pd.Series([value, 0, 99999], index=["std_value", "std_min", "std_max"])
        else:  # max
            return pd.Series([0, 0, value], index=["std_value", "std_min", "std_max"])

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
   
    # 14. Reff. 0 ( 0 ~ +0.5 )
    match14 = re.match(r'^.*?(-?\d+(?:\.\d+)?)\s*\(\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\)$',standard)
    if match14:
        nominal = float(match14.group(1))
        lower = float(match14.group(2))
        upper = float(match14.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])
    
    # 15. 15º ± 3º
    match15 = re.match(r'^(\d+(?:\.\d+)?)\s*º?\s*[±+]\s*(\d+(?:\.\d+)?)\s*º?$', standard)
    if match15:
        value = float(match15.group(1))
        margin = float(match15.group(2))
        return pd.Series([value, -margin, margin], index=["std_value", "std_min", "std_max"])

    # 16. Ambil angka dari bagian akhir teks seperti "Tidak ambles / minus 0 ( 0 ~ +0.5 )"
    match16 = re.search(r'(\d+(?:\.\d+)?)\s*\(\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\)', standard)
    if match16:
        nominal = float(match16.group(1))
        lower = float(match16.group(2))
        upper = float(match16.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])
    
    # 17. Reff. 0 ( -0.3 ~ 0 )" atau "Reff. 0 ( 0 ~ +0.5 )
    match17 = re.search(
        r'reff\.*\s*([+-]?\d+(?:[.,]\d+)?)\s*\(\s*([+-]?\d+(?:[.,]\d+)?)\s*~\s*([+-]?\d+(?:[.,]\d+)?)\s*\)',
        standard,
        re.IGNORECASE
    )
    if match17:
        nominal = float(match17.group(1).replace(",", "."))
        lower = float(match17.group(2).replace(",", "."))
        upper = float(match17.group(3).replace(",", "."))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])
    
    # 18. 32° ± 1°30', 3
    match18 = re.match(r"^(\d{1,3})[°º]?\s*±\s*(\d{1,3})[°º]?\s*(\d{1,2})['′`´]?", standard)
    if match18:
        std_value = int(match18.group(1))  # 32
        derajat = re.sub(r"[^\d]", "", match18.group(2))  
        menit = re.sub(r"[^\d]", "", match18.group(3))    
        tolerance = float(f"{derajat}.{menit.zfill(2)}") 
        return pd.Series([std_value, -tolerance, tolerance], index=["std_value", "std_min", "std_max"])
    
    # 19. Max/Min dengan simbol dan angka desimal
    match19 = re.search(
        r'\b(Min|Max)\.?\s*([Ø°]?\d+(?:\.\d+)?(?:[A-Za-z]*)?)',
        standard,
        re.IGNORECASE
    )
    if match19:
        kind = match19.group(1).lower()
        value_str = match19.group(2)
        try:
            value_num = float(re.sub(r'[^\d.-]', '', value_str))
        except:
            value_num = None
        if kind == "min":
            return pd.Series([value_num, 0, 99999], index=["std_value", "std_min", "std_max"]) # min
        else: 
            return pd.Series([0, 0, value_num], index=["std_value", "std_min", "std_max"]) # max

    return pd.Series([None, None, None], index=["std_value", "std_min", "std_max"])

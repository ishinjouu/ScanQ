import pandas as pd
import re
import numpy as np
from difflib import SequenceMatcher

# Jenis Point & Control Method Keywords --------------------------------------------------------------------//
dengan_ukur_keywords = [
    # from default.py
    "caliper", "hg", "depth cal", "pitch dial", "rough. t", "hitung", "depth clp", "height g", "dial g", 
    "dial clp.", "crm", "id micro", "rough t", "blog g + dept c.", "depth c.", "bore g.", "dial clp", "snap g.",
    "depth clp", "clp+jig", "clp + jig", "depth clp+jig", "depth clp + jig", "rought. t", "tg + clp", "groove cal.",
    "tg + cal",
    # from special.py
    "caliper", "hg", "depth cal", "pitch dial", "torque dial", "depth cal.",
    "rough. t", "hitung", "depth clp", "height g", "dial g", 
    "thickness meter", "thicknes s meter", "thicknes meter",
]
tanpa_ukur_keywords = [
    # from default.py
    "visual", "pg", "snap g.", "visual & punch", "visual + kikir", "visual & kikir", "chamfer g", "snap g",
    "machining test", "visual ( reff. master rough.)", "insp. jig", "finishing test",
    # from special.py
    "visual", "pg", "snap g.", "visual & punch", "visual + kikir", "Putar dg tangan",
    "visual & kikir", "machining test", "visual ( reff. master rough.)", "Poka Yoke test",
    "insp. jig", "finishing test", "Visual + tangan", "Leak tester", "Torque M"
]
dengan_cmm_keywords = ["cmm"]

# catatan - tambahan dari standard ke kolom catatan ---------------------------------------------------------//
def copy_special_measurements_to_note(row):
    standard = str(row.get("standard", "")).strip()
    catatan = str(row.get("catatan", "")).strip()
    jenis_point = str(row.get("jenis_point", "")).strip()

    if jenis_point not in ["Dengan Ukur", "Dengan CMM"]:
        return row
    # New Logic ---------------------------------------------------------------------
    if not standard or standard.lower() == "nan":
        return row
    if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", standard): # numeric murni → skip
        pass
    elif re.search(r"[°Ø±]", standard): # mengandung simbol khusus → biarin ke regex bawah
        pass
    else:
        if standard not in catatan:
            row["catatan"] = (catatan + " " + standard).strip()
        return row
    # New Logic ---------------------------------------------------------------------
    
    # === Min/Max ===
    minmax_matches = re.findall(r"(?:min|max)\s*[A-Za-z]*\s*\d+(?:\.\d+)?(?:\s*[A-Za-z]+)?", standard, flags=re.IGNORECASE)
    for item in minmax_matches:
        cleaned = f"[{item.strip()}]"
        if cleaned not in catatan:
            catatan += f" {cleaned}"

    # === Target Sett ===
    if "target sett" in standard.lower():
        if standard not in catatan:
            row["catatan"] = (catatan + " " + standard).strip()
        return row
            
    size_patterns = [
        r"[Ø°]\d+(?:\.\d+)?[A-Za-z]\d*\s*\(\s*-\s*\d+(?:\.\d+)?\s*~\s*-\s*\d+(?:\.\d+)?\s*\)", # Ø12N7 ( - 0.023 ~ -0.005 )
        r"\d+(?:\.\d+)?\s*[°º]\s*±\s*\d+(?:\.\d+)?\s*[°º]",                                 # 20° ± 1°
        r"[Ø°]\d+(?:\.\d+)?\s*±\s*[+−-]?\d+(?:\.\d+)?",                                     # Ø10 ±0.1 atau °10 ± 0.5
        r"[Ø°]\d+(?:\.\d+)?\s*\(\s*\d+(?:\.\d+)?\s*~\s*[+−-]?\d+(?:\.\d+)?\s*\)",           # Ø6.1 ( 0 ~ +0.1 )
        r"[Ø°]\d+(?:\.\d+)?\s*\(\s*[-+]?\d+(?:\.\d+)?\s*~\s*[-+]?\d+(?:\.\d+)?\s*\)",       # Ø6.1 (0 ~ +0.1) / Ø40 (-0.050 ~ -0.035)
        r"[Ø°]\d+(?:\.\d+)?\s*\(\s*[+−-]?\d+(?:\.\d+)?\s*~\s*[+−-]?\d+(?:\.\d+)?\s*\)",     # Ø12.15 (-0.15 ~ +0.25)
        r"\d+(?:\.\d+)?º\s*±\s*\d+(?:\.\d+)?º(?:\d{1,2}['′`´])?",                           # 15º ± 3º
        r"\d{1,3}[°º]\s*±\s*\d{1,3}[°º]\s*\d{1,2}['′`´]",                                   # 32° ± 1°30', 3
        r"[A-Za-z]{1,3}\s*\d+(?:\.\d+)?\s*~\s*[A-Za-z]{1,3}\s*\d+(?:\.\d+)?",               # Rz 2 ~ Rz 12.5
        r"[A-Za-z]\s*\d+(?:\.\d+)?\s*~\s*\d+(?:\.\d+)?",                                    # C0.5 ~ 1.0
        r"[Ø°]\d+[A-Za-z0-9]*\s*\(\s*[+-]?\d+(?:\.\d+)?\s*~\s*[+-]?\d+(?:\.\d+)?\s*\)",     # Ø8H8 (0 ~ +0.022) / Ø32P7 (...)
        r"[A-Za-z]\d+(?:\.\d+)?\s*\(\s*[+-]?\d+(?:\.\d+)?\s*~\s*[+-]?\d+(?:\.\d+)?\s*\)",   # C0.5 ( 0 ~ +0.2 )
    ]
    # ukuran_found = None
    # for pattern in size_patterns:
    #     match = re.search(pattern, standard, flags=re.IGNORECASE | re.DOTALL)
    #     if match:
    #         ukuran_found = match.group(0).strip()
    #         break
    # if ukuran_found:
    #     if ukuran_found not in catatan:
    #         catatan += f" {ukuran_found}"
    for pattern in size_patterns:
        match = re.search(pattern, standard, flags=re.IGNORECASE | re.DOTALL)
        if match:
            ukuran_found = match.group(0).strip()
            if ukuran_found not in catatan:
                catatan += f" {ukuran_found}"
            break
    row["catatan"] = catatan.strip()
    # print(repr(standard))
    return row

# Format standard ke std_value, std_min, std_max ------------------------------------------------------------------------- //
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

    # 3. Ø32P7 (0 ~ +0.022) 
    match3a = re.match( r'^\s*[Ø°]?\s*(\d+(?:\.\d+)?)[A-Za-z0-9]*\s*\(\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\)', standard)
    if match3a:
        nominal = float(match3a.group(1))
        lower = float(match3a.group(2))
        upper = float(match3a.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])
    # -- Ø6.1 ( 0 ~ +0.1 )
    match3b = re.match(r'^[Ø°]?\s*(-?\d+(?:\.\d+)?)\s*\(\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\)', standard)
    if match3b:
        nominal = float(match3b.group(1))
        lower = float(match3b.group(2))
        upper = float(match3b.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])
    # -- Ø12N7 ( - 0.023 ~ -0.005 )
    match3c = re.match(r'^\s*[Ø°]?\s*(\d+(?:\.\d+)?)[A-Za-z]\d*\s*\(\s*-\s*(\d+(?:\.\d+)?)\s*~\s*-\s*(\d+(?:\.\d+)?)\s*\)', standard)
    if match3c:
        nominal = float(match3c.group(1))
        lower = -float(match3c.group(2))
        upper = -float(match3c.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])

    # 4. [ 163.89 ± 0.25 ]
    match4 = re.match(r'^\[\s*(-?\d+(?:\.\d+)?)\s*±\s*([\d.]+)\s*\]$', standard)
    if match4:
        nominal = float(match4.group(1))
        delta = float(match4.group(2))
        return pd.Series([nominal, -delta, delta], index=["std_value", "std_min", "std_max"])

    # 5. Min/Max
    match5 = re.search(r'\b(Min|Max)\.?\s*(\d+(?:\.\d+)?)', standard, re.IGNORECASE)
    if match5:
        kind = match5.group(1).lower()
        value = float(match5.group(2))
        if kind == "min":
            return pd.Series([value, 0, 99999], index=["std_value", "std_min", "std_max"])
        else:  # max
            return pd.Series([0, 0, value], index=["std_value", "std_min", "std_max"])
        
    # 5b. Max Rz 25  |  Min Ra 3.2  |  Max Rmax 5
    match5b = re.search(r'\b(Min|Max)\s+[A-Za-z]{1,4}\s*(\d+(?:\.\d+)?)', standard, re.IGNORECASE)
    if match5b:
        kind = match5b.group(1).lower()
        value = float(match5b.group(2))
        if kind == "min":
            return pd.Series([value, 0, 99999], index=["std_value", "std_min", "std_max"])
        else:
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

    # 12. Ø7 [-0.3 ~ 0] & 35.5 ( - 0.2 ~ 0 )
    match12 = re.match(r'^[Ø°]?\s*(-?\d+(?:\.\d+)?)\s*\[\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?)\s*\]$', standard)
    if match12:
        nominal = float(match12.group(1))
        lower = float(match12.group(2))
        upper = float(match12.group(3))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])
    
    match12b = re.match(r'^(-?\d+(?:\.\d+)?)\s*\(\s*([+-]?\s*\d+(?:\.\d+)?)\s*~\s*([+-]?\s*\d+(?:\.\d+)?)\s*\)$',standard)
    if match12b:
        nominal = float(match12b.group(1))
        lower = float(match12b.group(2).replace(' ', ''))
        upper = float(match12b.group(3).replace(' ', ''))
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
    match17 = re.search(r'reff\.*\s*([+-]?\d+(?:[.,]\d+)?)\s*\(\s*([+-]?\d+(?:[.,]\d+)?)\s*~\s*([+-]?\d+(?:[.,]\d+)?)\s*\)',standard,re.IGNORECASE)
    if match17:
        nominal = float(match17.group(1).replace(",", "."))
        lower = float(match17.group(2).replace(",", "."))
        upper = float(match17.group(3).replace(",", "."))
        return pd.Series([nominal, lower, upper], index=["std_value", "std_min", "std_max"])
    
    # 18. 32° ± 1°30', 3
    # match18 = re.match(r"^(\d{1,3})[°º]?\s*±\s*(\d{1,3})[°º]?\s*(\d{1,2})['′`´]?", standard)
    # if match18:
    #     std_value = int(match18.group(1))  # 32
    #     derajat = re.sub(r"[^\d]", "", match18.group(2))  
    #     menit = re.sub(r"[^\d]", "", match18.group(3))    
    #     tolerance = float(f"{derajat}.{menit.zfill(2)}") 
    #     return pd.Series([std_value, -tolerance, tolerance], index=["std_value", "std_min", "std_max"])
    # 18. 32° ± 1°30', 3
    match18 = re.match(r"^(\d{1,3})[°º]?\s*±\s*(\d{1,3})[°º]?\s*(\d{1,2})?['′`´]?",standard)
    if match18:
        std_value = int(match18.group(1))  
        derajat = re.sub(r"[^\d]", "", match18.group(2))  
        menit_raw = match18.group(3)
        menit = re.sub(r"[^\d]", "", menit_raw) if menit_raw else "00"
        tolerance = float(f"{derajat}.{menit.zfill(2)}")
        return pd.Series([std_value, -tolerance, tolerance],index=["std_value", "std_min", "std_max"])
    
    # 19. Max/Min dengan simbol dan angka desimal
    match19 = re.search(r'\b(Min|Max)\.?\s*([Ø°]?\d+(?:\.\d+)?(?:[A-Za-z]*)?)',standard,re.IGNORECASE)
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
        
    # 20. Min 20µ / Min 20um / Min20µm
    match20 = re.search(r'\bMin\.?\s*([0-9]+)\s*(µ|um|µm)?\b', standard, re.IGNORECASE)
    if match20:
        value = float(match20.group(1))
        return pd.Series([value, 0, 99999], index=["std_value", "std_min", "std_max"])
    
    # 21. Rz 2 ~ Rz 12.5 
    match21 = re.match(r'^\s*([A-Za-z]{1,3})\s*(\d+(?:\.\d+)?)\s*~\s*\1\s*(\d+(?:\.\d+)?).*$', standard, flags=re.IGNORECASE)
    if match21:
        lower = float(match21.group(1))
        upper = float(match21.group(2))
        return pd.Series([0.0, lower, upper], index=["std_value", "std_min", "std_max"])
    # -- C0.5 ~ 1.0  
    match21b = re.match(r'^\s*([A-Za-z])\s*(\d+(?:\.\d+)?)\s*~\s*(\d+(?:\.\d+)?).*$', standard, flags=re.IGNORECASE)
    if match21b:
        lower = float(match21b.group(1))
        upper = float(match21b.group(2))
        return pd.Series([0.0, lower, upper], index=["std_value", "std_min", "std_max"])
    # -- 1.5a ~ 2.5
    match21c = re.match(r'^\s*(\d+(?:\.\d+)?)\s*[A-Za-z]\s*~\s*(\d+(?:\.\d+)?).*$', standard, flags=re.IGNORECASE)
    if match21c:
        lower = float(match21c.group(1))
        upper = float(match21c.group(2))
        return pd.Series([0.0, lower, upper], index=["std_value", "std_min", "std_max"])
        
    # 22. 3.2 a (max)
    match22 = re.match(r'^\s*(\d+(?:\.\d+)?)\s*[A-Za-z]\s*$', standard)
    if match22:
        max_val = float(match22.group(1))
        return pd.Series([0.0, 0.0, max_val], index=["std_value", "std_min", "std_max"])

    # 23. ( 105.5 ) reff => 105,5 |0| 300
    match23 = re.match(r'^\s*\(\s*(\d+(?:\.\d+)?)\s*\)', standard)
    if match23:
        std_value = float(match23.group(1))
        return pd.Series([std_value, 0.0, 300.0], index=["std_value", "std_min", "std_max"])

    # 24. 12 ( 0 ~ +1pitch ) => 12 | 0 | 1
    match24 = re.match(r'^\s*(\d+(?:\.\d+)?)\s*\(\s*([+-]?\d+(?:\.\d+)?)\s*~\s*([+-]?\d+(?:\.\d+)?).*?\)', standard)
    if match24:
        std_value = float(match24.group(1))
        lower = float(match24.group(2))
        upper = float(match24.group(3))
        return pd.Series([std_value, lower, upper], index=["std_value", "std_min", "std_max"])

    return pd.Series([None, None, None], index=["std_value", "std_min", "std_max"])

# Dengan CMM -------------------------------------------------------------------------------------------------
def append_cmm_summary_row(df):
    df_cmm = df[df["jenis_point"] == "Dengan CMM"] # Filter only CMM rows
    if df_cmm.empty:
        return df

    last_section = df["section"].dropna().tolist() # Determine next section number
    last_num = 0
    for s in last_section[::-1]:
        match = re.match(r"^(\d+)\.", str(s).strip())
        if match:
            last_num = int(match.group(1))
            break
    new_section = df['section'].dropna().iloc[-1] if not df['section'].dropna().empty else "General"
    
    def extract_point_number(val): # Determine next point_check number
        try:
            return int(str(val).strip().split(".")[0])
        except:
            return 0

    point_nums = df["point_check"].dropna().apply(extract_point_number)
    max_point = point_nums.max() if not point_nums.empty else 0
    new_point_check = str(max_point + 1)
    all_pengecekan = df_cmm["jenis_pengecekan"].dropna().tolist() # Determine jenis_pengecekan list

    flat_list = []
    for item in all_pengecekan:
        if isinstance(item, list):
            flat_list.extend(item)
        elif isinstance(item, str):
            flat_list.extend([s.strip() for s in item.split(",") if s.strip()])
    clean_list = list(set(flat_list)) if flat_list else ["-"]

    new_row = { # Build the new row
        "section": new_section,
        "point_check": new_point_check,
        "jenis_point": "Tanpa Ukur",
        "item_check": "Hasil CMM",
        "standard": "Masuk range toleransi",
        "catatan": "-",
        "jenis_pengecekan": clean_list,
        "control_method": "CMM",
        "std_value": None,
        "std_min": None,
        "std_max": None,
        "status": "valid"
    }
    for col in df.columns: # Match all columns, fill missing with "-"
        if col not in new_row:
            new_row[col] = "-"

    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    return df

# label bolong b. dll ------------------------------------------------------------//
# ga error tapi di case Sumbu Ø51, no nya masih harus di edit manual
def isi_label_abjad_di_antara(df, kolom='Item', No='No'):
    if No not in df.columns:
        if 'No.' in df.columns:
            No = 'No.'
        else:
            df[No] = ""
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

# from default.py -------------------------------------------------------------------------------------------//
def sanitize_for_json(value):
    if isinstance(value, float) and (pd.isna(value) or not np.isfinite(value)):
        return None
    if isinstance(value, list):
        return [sanitize_for_json(v) for v in value]
    if isinstance(value, dict):
        return {k: sanitize_for_json(v) for k, v in value.items()}
    return value

# --- isi kosong di kolom catatan berdasarkan grup point_check dan item_check --------------------------------//
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

import pandas as pd
import os, re
from config import DUTY_SECTION_MAP

def normalize_name(name_str):
    if not isinstance(name_str, str): return ""
    return re.sub(r"[^A-Z0-9]", "", name_str.upper())

def load_staff_database(db_path):
    staff_db = {}
    
    if not os.path.exists(db_path):
        print(f"WARNING: '{db_path}' not found. Skipping external data.")
        return staff_db

    try:
        xl = pd.ExcelFile(db_path)
        sheet_names = xl.sheet_names
        
        # 1. Load Radios
        radio_sheet = next((s for s in sheet_names if "Radio" in s), None)
        if radio_sheet:
            df = pd.read_excel(db_path, sheet_name=radio_sheet)
            df.columns = [c.strip().lower() for c in df.columns]
            name_col = next((c for c in df.columns if "name" in c), None)
            rad_col = next((c for c in df.columns if "radio" in c), None)
            
            if name_col and rad_col:
                for _, row in df.iterrows():
                    if pd.notna(row[name_col]):
                        norm = normalize_name(str(row[name_col]))
                        if norm not in staff_db: staff_db[norm] = {}
                        staff_db[norm]["radio"] = str(row[rad_col])

        # 2. Load Restrictions
        restr_sheet = next((s for s in sheet_names if "Restriction" in s), None)
        if restr_sheet:
            df = pd.read_excel(db_path, sheet_name=restr_sheet)
            df.columns = [c.strip().lower() for c in df.columns]
            r_name_col = next((c for c in df.columns if "name" in c), None)
            
            if r_name_col:
                for _, row in df.iterrows():
                    if pd.notna(row[r_name_col]):
                        norm = normalize_name(str(row[r_name_col]))
                        if norm not in staff_db: staff_db[norm] = {}
                        staff_db[norm]["restricted"] = True

        print(f"SUCCESS: Database loaded for {len(staff_db)} staff.")
        
    except Exception as e:
        print(f"ERROR loading database: {e}")

    return staff_db


def split_by_duty_section(staff_list: list[dict]):
    """
    Splits the staff list into 3 sections using DUTY_SECTION_MAP.
    Any duty code not in the map is routed to Victoria AND surfaced as a warning
    so supervisors know the map needs updating.
    """
    victoria = []
    district = []
    cardinal = []
    unmatched = []  # (duty_code, name) pairs

    print(f"DEBUG: Splitting {len(staff_list)} staff using the Map...")

    for s in staff_list:
        raw_id = s.get("id", "")
        code = raw_id.strip().upper()

        section = DUTY_SECTION_MAP.get(code)
        if section is None:
            unmatched.append((code, s.get("name", "")))
            section = "Victoria"

        if section == "Victoria":
            victoria.append(s)
        elif section == "District":
            district.append(s)
        elif section == "Cardinal":
            cardinal.append(s)
        else:
            victoria.append(s)

    victoria.sort(key=lambda x: x["start_time"])
    district.sort(key=lambda x: x["start_time"])
    cardinal.sort(key=lambda x: x["start_time"])

    print(f"DEBUG: Split Results -> Vic: {len(victoria)}, Dist: {len(district)}, Card: {len(cardinal)}")

    if unmatched:
        print(f"WARNING: {len(unmatched)} duty code(s) not in DUTY_SECTION_MAP — defaulted to Victoria:")
        for code, name in unmatched:
            print(f"  - {code!r}  ({name})")
        print("  -> Update DUTY_SECTION_MAP in config.py to route these correctly.")

    return victoria, district, cardinal

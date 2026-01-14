import pandas as pd
import os
from config import DUTY_SECTION_MAP

def normalize_name(name_str):
    if not isinstance(name_str, str): return ""
    return name_str.replace(" ", "").upper()

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
    Splits the staff list into 3 sections based STRICTLY on DUTY_SECTION_MAP.
    """
    victoria = []
    district = []
    cardinal = []

    print(f"DEBUG: Splitting {len(staff_list)} staff using the Map...")

    for s in staff_list:
        raw_id = s.get("id", "")
        code = raw_id.strip().upper()
        
        section = DUTY_SECTION_MAP.get(code)
        
        if not section:
            section = "Victoria"

        # 4. Sort into lists
        if section == "Victoria":
            victoria.append(s)
        elif section == "District":
            district.append(s)
        elif section == "Cardinal":
            cardinal.append(s)
        else:
            victoria.append(s)

    # 5. Sort each section by Start Time
    victoria.sort(key=lambda x: x["start_time"])
    district.sort(key=lambda x: x["start_time"])
    cardinal.sort(key=lambda x: x["start_time"])

    print(f"DEBUG: Split Results -> Vic: {len(victoria)}, Dist: {len(district)}, Card: {len(cardinal)}")
    return victoria, district, cardinal

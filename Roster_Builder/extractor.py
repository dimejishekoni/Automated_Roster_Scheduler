import re
import pdfplumber
from pathlib import Path
from datetime import datetime, time
from database import normalize_name 

def excel_time_to_str(val) -> str:
    if isinstance(val, (time, datetime)):
        return val.strftime("%H:%M")
    s = str(val).strip()
    if ":" in s and len(s) <= 5: return s
    try:
        num = float(s)
        total_minutes = int(round(num * 24 * 60))
        h = total_minutes // 60
        m = total_minutes % 60
        return f"{h:02d}:{m:02d}"
    except Exception:
        return "00:00"

def parse_grade(duty_desc: str) -> str | None:
    text = (duty_desc or "").upper()
    if "CSA1" in text: return "CSA1"
    if "CSA2" in text: return "CSA2"
    return None

# ================= STRICT VIC EXTRACTOR =================
def extract_victoria_rows_from_text(pdf_path: str, date_str: str | None = None):
    pdf_path = Path(pdf_path)
    rows: list[list[str]] = []

    if not pdf_path.exists():
        print(f"ERROR: PDF not found at {pdf_path}")
        return rows

    # Regex: HH:MM whitespace HH:MM
    time_pattern = re.compile(r"(\d{1,2}:\d{2})\s+(\d{1,2}:\d{2})")

    # Prepare Target Date (Normalize to uppercase, no spaces)
    wanted_date_norm = normalize_name(date_str) if date_str else None
    if wanted_date_norm:
        print(f"DEBUG: Filtering for date: {wanted_date_norm}")

    print(f"DEBUG: Opening PDF {pdf_path}...")

    with pdfplumber.open(str(pdf_path)) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text()
            if not text: continue
            
            clean_text = normalize_name(text)

            # --- 1. DATE FILTER ---
            if wanted_date_norm and wanted_date_norm not in clean_text:
                continue

            # --- 2. LOCATION FILTER ---
            if "LOCATIONVICTORIA" not in clean_text and "DUTIESFORVICTORIA" not in clean_text:
                if "VICTORIA" not in clean_text:
                    continue

            lines = text.splitlines()
            header_idx = None
            
            # Find Header
            for idx, line in enumerate(lines):
                clean_line = line.upper().replace(" ", "")
                if "DUTY" in clean_line and "START" in clean_line and "NAME" in clean_line:
                    header_idx = idx
                    break
            
            if header_idx is None: continue

            # Parse Rows
            for line in lines[header_idx + 1 :]:
                line = line.strip()
                if not line or "SIGNING ON" in line.upper() or "AUTHORISED" in line.upper():
                    continue

                # --- 3. STRICT 'VIC' FILTER ---
                if "VIC" not in line.upper():
                    continue

                match = time_pattern.search(line)
                if not match: continue 

                start_val, end_val = match.groups()
                
                pre_time = line[:match.start()].strip()
                post_time = line[match.end():].strip()
                
                pre_parts = pre_time.split(maxsplit=1)
                if len(pre_parts) == 2:
                    duty_code, duty_desc = pre_parts
                elif len(pre_parts) == 1:
                    duty_code = pre_parts[0]
                    duty_desc = ""
                else:
                    duty_code, duty_desc = "", ""

                name_val = post_time
                rows.append([duty_code, duty_desc, start_val, end_val, name_val])

    print(f"DEBUG: Extraction complete. Found {len(rows)} rows for date {date_str}.")
    return rows

def build_staff_from_pdf_rows(rows, staff_db):
    all_staff = []
    missing_from_db = []

    for row in rows:
        if not row or len(row) < 5: continue

        duty_code = (row[0] or "").strip()
        duty_desc = (row[1] or "").strip()
        start_val = (row[2] or "").strip()
        end_val   = (row[3] or "").strip()
        name_raw  = (row[4] or "").strip()

        if "CSS" in duty_desc or "CSM" in duty_desc: continue
        if "GPK" in duty_desc: continue

        name_val = re.split(r"\s{2,}", name_raw)[0].strip()

        grade     = parse_grade(duty_desc)
        start_str = excel_time_to_str(start_val)
        end_str   = excel_time_to_str(end_val)

        clean_name = normalize_name(name_val)
        db_info = staff_db.get(clean_name, {})
        if not db_info:
            missing_from_db.append((duty_code, name_val))

        display_name = name_val.split(",", 1)[0].strip()

        staff_entry = {
            "id": duty_code,
            "name": display_name,
            "radio": db_info.get("radio", ""),
            "start_time": start_str,
            "finish_time": end_str,
            "grade": grade,
            "duty_desc": duty_desc,
            "is_restricted": db_info.get("restricted", False),
        }

        all_staff.append(staff_entry)

    if missing_from_db:
        print(f"WARNING: {len(missing_from_db)} staff not found in Staff Database (no radio/restriction info):")
        for code, name in missing_from_db:
            print(f"  - {code}  {name!r}")

    return all_staff

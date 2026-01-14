import os
from config import PDF_PATH, DB_PATH, OUTPUT_FILE, GRID_CONFIG, DUTY_SECTION_MAP
from database import load_staff_database, split_by_duty_section
from extractor import extract_victoria_rows_from_text, build_staff_from_pdf_rows
from config import GridConfig
from renderer import build_single_sheet
from scheduler import assign_victoria_duties_v2

target_date = "Monday20October2025"
def main():
    print("=== STARTING ROSTER AUTOMATION ===")
    print(f"Loading Database: {DB_PATH}")
    staff_db = load_staff_database(DB_PATH)
    print(f"Reading PDF: {PDF_PATH}")
    raw_rows = extract_victoria_rows_from_text(PDF_PATH, date_str=target_date)
    all_staff = build_staff_from_pdf_rows(raw_rows, staff_db)
    print(f"Found {len(all_staff)} valid VIC staff members.")

    if not all_staff:
        print("ERROR: No staff found. Check PDF text extraction.")
        return

    # Run Scheduler
    cfg = GridConfig(
        start_hour=GRID_CONFIG["start_hour"],
        end_hour=GRID_CONFIG["end_hour"],
        start_column=GRID_CONFIG["start_column"],
        blocks_per_hour=GRID_CONFIG["blocks_per_hour"]
    )
    assigned_staff = assign_victoria_duties_v2(all_staff, cfg, date_str= target_date)
    # Split Sections
    victoria, district, cardinal = split_by_duty_section(assigned_staff)

    # Render Excel
    print("Building Excel File...")
    
    # Dynamic row counts
    rows_vic = len(victoria) + 5
    rows_dist = len(district) + 5
    rows_card = len(cardinal) + 5
    
    build_single_sheet(
        OUTPUT_FILE,
        victoria, district, cardinal,
        cfg,
        rows_per_section=(rows_vic, rows_dist, rows_card)
    )

    print(f"=== DONE. Saved to {OUTPUT_FILE} ===")

if __name__ == "__main__":
    main()
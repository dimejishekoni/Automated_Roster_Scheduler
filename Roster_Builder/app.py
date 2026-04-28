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

# import argparse
# from datetime import date
# from pathlib import Path

# from config import PDF_PATH, DB_PATH, OUTPUT_FILE, GRID_CONFIG, GridConfig
# from database import load_staff_database, split_by_duty_section
# from extractor import extract_victoria_rows_from_text, build_staff_from_pdf_rows
# from renderer import build_single_sheet

# DEFAULT_DATE = "Monday20October2025"

# def default_date_str() -> str:
#     """Fallback when --date is not passed. Uses config DEFAULT_DATE if set, else today."""
#     return DEFAULT_DATE or date.today().strftime("%A%d%B%Y")



# def parse_args():
#     p = argparse.ArgumentParser(description="Extract VSOU sign-on sheet and distribute staff by section.")
#     p.add_argument("--date", default=default_date_str(),
#                    help="Target date in PDF format, e.g. 'Monday20October2025'. Defaults to today.")
#     p.add_argument("--pdf", default=PDF_PATH, help=f"Path to the SOS PDF (default: {PDF_PATH}).")
#     p.add_argument("--db",  default=DB_PATH,  help=f"Path to the Staff Database xlsx (default: {DB_PATH}).")
#     p.add_argument("--out", default=None,
#                    help=f"Output xlsx path. Defaults to 'output/Roster_<date>.xlsx'.")
#     return p.parse_args()


# def main():
#     args = parse_args()
#     target_date = args.date
#     out_path = args.out or f"Roster_{target_date}.xlsx"

#     print("=== STARTING ROSTER EXTRACTION ===")
#     print(f"Date:     {target_date}")
#     print(f"PDF:      {args.pdf}")
#     print(f"Database: {args.db}")

#     staff_db = load_staff_database(args.db)
#     raw_rows = extract_victoria_rows_from_text(args.pdf, date_str=target_date)
#     all_staff = build_staff_from_pdf_rows(raw_rows, staff_db)

#     if not all_staff:
#         print("ERROR: No staff extracted. Check the date spelling and that the PDF matches.")
#         return

#     victoria, district, cardinal = split_by_duty_section(all_staff)

#     cfg = GridConfig(**GRID_CONFIG)

#     # Give each section at least enough rows for its staff, with a small buffer.
#     rows_per_section = (
#         max(len(victoria) + 3, 10),
#         max(len(district) + 3, 10),
#         max(len(cardinal) + 3, 5),
#     )

#     print("Building Excel File...")
#     build_single_sheet(
#         out_path,
#         victoria, district, cardinal,
#         cfg,
#         rows_per_section=rows_per_section,
#     )

#     print(f"=== DONE. Saved to output/{out_path} ===")
#     print(f"Summary: {len(all_staff)} staff total -> "
#           f"Victoria {len(victoria)} | District {len(district)} | Cardinal {len(cardinal)}")


# if __name__ == "__main__":
#     main()

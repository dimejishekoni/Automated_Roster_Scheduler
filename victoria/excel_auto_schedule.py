import math, os
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
from collections import defaultdict
import pdfplumber
from pathlib import Path
import re


DUTY_PRIORITY = {
    "P3":          1,
    "PTI WB":      2,
    "PTI EB":      2,
    "SATS WB":     3,
    "SATS EB":     4,
    "ESCALATOR":   5,
    "VIP/MIP":     6,
    "SECURITY":    7,
}

# ================= Palettes (deep/light per section) =================
PALETTES = {
    "Victoria": {"deep": "00B0F0", "light": "92CCDC", "banner": "00B0F0", "text": "FFFFFF"},
    "District": {"deep": "33CC33", "light": "C4D79B", "banner": "33CC33", "text": "FFFFFF"},
    "Cardinal": {"deep": "FE66CC", "light": "DA9694", "banner": "FE66CC", "text": "FFFFFF"},
}

# --- COLORS used for notes/legend (tweak to taste) ---
PURPLE   = "9059FB"  # CSSI banner background
WHITE    = "FFFFFF"
GREY     = "A895BE"
YELLOW   = "FEFF2F"
ORANGE   = "DE6220"
LIME     = "C6E0B4"
TAN      = "DDD9C4"
GREEN    = "6BE181"
PEACH    = "FABF8E"
BLACK_THIN = Side(border_style="thin", color="000000")


# Map duty-code prefix -> section name (tune this to match your real codes)
DUTY_SECTION_MAP = {
    "BN71": "Victoria", 
    "BN21": "Victoria", 
    "BN73": "Victoria", 
    "BN27": "Victoria", 
    "BN28": "Victoria", 
    "BN25": "Victoria", 
    "BN26": "Victoria", 
    "BN31": "Victoria", 
    "BN32": "Victoria", 
    "BN76": "Victoria", 
    "BN78": "Victoria", 
    "BN37": "Victoria", 
    "BN38": "Victoria", 
    "BN44": "Victoria", 
    "BN39": "Victoria", 
    "BN40": "Victoria", 
    "BN80": "Victoria", 
    "BN72": "District",
    "BN22": "District",
    "BN75": "District",
    "BN74": "District",
    "BN29": "District",
    "BN33": "District",
    "BN77": "District",
    "BN34": "District",
    "BN81": "District",
    "BN43": "District",
    "BN79": "District",
    "BN45": "District",
    "BN41": "District",
    "BN82": "District",
    "BN23": "Cardinal",
    "BN24": "Cardinal",
    "BN30": "Cardinal",
    "BN36": "Cardinal",
    "BN42": "Cardinal",
    "BN35": "Cardinal",
    # add / change as needed
}

# ---------- tiny helpers ----------
def hhmm_to_minutes(hhmm: str) -> int:
    h, m = map(int, hhmm.split(":"))
    return h * 60 + m

def minutes_to_hhmm(m: int) -> str:
    m %= 24 * 60
    h = m // 60
    mm = m % 60
    return f"{h:02d}:{mm:02d}"

def clean_merge_area(ws, min_row, min_col, max_row, max_col):
    """
    Safely removes any existing merges in the target area to prevent 
    overlapping merge errors (which corrupt Excel files).
    """
    # We must collect them first, because we can't delete while iterating
    ranges_to_remove = []
    
    for merged_range in ws.merged_cells.ranges:
        # Get bounds of the existing merge
        mr_min_col, mr_min_row, mr_max_col, mr_max_row = merged_range.bounds
        
        # Check if this existing merge overlaps with our new target area
        if (mr_min_col <= max_col and mr_max_col >= min_col and
            mr_min_row <= max_row and mr_max_row >= min_row):
            ranges_to_remove.append(merged_range)
            
    # Remove them
    for r in ranges_to_remove:
        ws.unmerge_cells(str(r))

# ---------- Duty list Cleaner ----------

def _parse_grade(duty_desc: str) -> str | None:
    text = (duty_desc or "").upper()
    if "CSA1" in text:
        return "CSA1"
    if "CSA2" in text:
        return "CSA2"
    return None

def _excel_time_to_str(val) -> str:
    """Convert Excel time or string into 'HH:MM'."""
    from datetime import time, datetime
    if isinstance(val, (time, datetime)):
        return val.strftime("%H:%M")
    s = str(val).strip()
    # already looks like HH:MM
    if ":" in s and len(s) <= 5:
        return s
    # Excel time as fraction of day
    try:
        num = float(s)
        total_minutes = int(round(num * 24 * 60))
        h = total_minutes // 60
        m = total_minutes % 60
        return f"{h:02d}:{m:02d}"
    except Exception:
        return "00:00"
    

# ---------- Import PDF ----------

def _norm(s: str) -> str:
    """Remove all whitespace – makes matching 'Sunday19October2025' easier."""
    return "".join(str(s).split())

# ================= ROBUST EXTRACTION LOGIC =================
def extract_victoria_rows_from_text(pdf_path: str, date_str: str | None = None):
    """
    Robust extraction that uses Regex to find times, allowing descriptions 
    with spaces (e.g. 'SATS WB').
    """
    pdf_path = Path(pdf_path)
    rows: list[list[str]] = []

    if not pdf_path.exists():
        print("File not found:", pdf_path)
        return rows

    wanted_date_norm = _norm(date_str) if date_str else None
    
    # Regex to find two times side-by-side, e.g., "05:00 13:00"
    # Captures: (Start) (End)
    time_pattern = re.compile(r"(\d{1,2}:\d{2})\s+(\d{1,2}:\d{2})")

    with pdfplumber.open(str(pdf_path)) as pdf:
        print("Total pages:", len(pdf.pages))

        for page_index, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            nt = _norm(text)

            if "VictoriaArea" not in nt:
                continue
            if wanted_date_norm and wanted_date_norm not in nt:
                continue

            lines = text.splitlines()
            header_idx = None
            for i, line in enumerate(lines):
                if "Duty DutyDescription Start End Name" in line:
                    header_idx = i
                    break
            if header_idx is None:
                continue

            for line in lines[header_idx + 1 :]:
                line = line.strip()
                if not line or line.startswith("Authorisedby") or line.startswith("SIGNING ON SHEET"):
                    continue

                # 1. Find the times to anchor the split
                match = time_pattern.search(line)
                if not match:
                    continue # Skip lines without valid times

                start_val, end_val = match.groups()
                
                # 2. Split everything before the time into Code + Desc
                pre_time = line[:match.start()].strip()
                post_time = line[match.end():].strip()
                
                # Assume Code is the first word, Description is the rest
                pre_parts = pre_time.split(maxsplit=1)
                if len(pre_parts) == 2:
                    duty_code, duty_desc = pre_parts
                elif len(pre_parts) == 1:
                    duty_code = pre_parts[0]
                    duty_desc = ""
                else:
                    duty_code, duty_desc = "", ""

                # 3. Name is everything after the times
                name_val = post_time

                rows.append([duty_code, duty_desc, start_val, end_val, name_val])

    return rows

def build_staff_from_pdf_rows(rows):
    victoria = []

    for row in rows:
        if not row or len(row) < 5:
            continue

        duty_code = (row[0] or "").strip()
        duty_desc = (row[1] or "").strip()
        start_val = (row[2] or "").strip()
        end_val   = (row[3] or "").strip()
        name_val  = (row[4] or "").strip()

        # ---- filters ----
        # Only Victoria station duties: descriptions starting with 'VIC'
        if not duty_desc.startswith("VIC"):
            continue

        # Ignore supervisors / managers (CSS/CSM)
        # In the PDF these appear as 'VICCSS...' or 'VICCSM...'
        if "CSS" in duty_desc or "CSM" in duty_desc:
            continue

        # (We also don't want GPK but those are filtered out by the "VIC " test)

        grade    = _parse_grade(duty_desc)      # CSA1 / CSA2 or None
        start_str = _excel_time_to_str(start_val)
        end_str   = _excel_time_to_str(end_val)

        staff_entry = {
            "id": duty_code,
            "name": name_val,
            "radio": "",
            "start_time": start_str,
            "finish_time": end_str,
            "grade": grade,
            "duty_desc": duty_desc,
        }
        victoria.append(staff_entry)

    return victoria

def split_by_duty_section(victoria_raw: list[dict]):
    """
    Take all VIC duties and split into (victoria, district, cardinal)
    based on the duty code using DUTY_SECTION_MAP.
    """
    victoria: list[dict] = []
    district: list[dict] = []
    cardinal: list[dict] = []

    for s in victoria_raw:
        code = s.get("id", "")
        section = DUTY_SECTION_MAP.get(code, "Victoria")  # default to Victoria

        if section == "Victoria":
            victoria.append(s)
        elif section == "District":
            district.append(s)
        elif section == "Cardinal":
            cardinal.append(s)
        else:
            # unknown label – you can log it if you want
            victoria.append(s)

    # optional: sort by start time in each section
    for lst in (victoria, district, cardinal):
        lst.sort(key=lambda x: x["start_time"])

    return victoria, district, cardinal


# ================= Grid Config =================
class GridConfig:
    def __init__(
        self,
        start_hour=5,           # 05:00
        end_hour=25,            # 01:00 next day (covers 00:00)
        start_column=6,         # F
        blocks_per_hour=4,      # 4 columns per hour (15-min blocks)
        meta_col_widths=(5.6, 8, 8, 14.5, 6.2),   # widths A..E
    ):
        self.start_hour = start_hour
        self.end_hour = end_hour
        self.start_column = start_column
        self.blocks_per_hour = blocks_per_hour
        self.meta_col_widths = meta_col_widths

    @property
    def minutes_per_block(self): return 60 // self.blocks_per_hour

    @property
    def total_hours(self): return self.end_hour - self.start_hour

    @property
    def total_blocks(self): return self.total_hours * self.blocks_per_hour


# ================= Helpers: time mapping =================
def _to_minutes_since_start(hhmm: str, cfg: GridConfig) -> int:
    h, m = map(int, hhmm.split(":"))
    s = cfg.start_hour % 24
    if h < s:  # 00:xx => next day
        h += 24
    return max(0, (h - cfg.start_hour) * 60 + m)

def _block_index_start(hhmm: str, cfg: GridConfig) -> int:
    # step-in: round start UP
    return int(math.ceil(_to_minutes_since_start(hhmm, cfg) / cfg.minutes_per_block))

def _block_index_end(hhmm: str, cfg: GridConfig, round_finish="ceil") -> int:
    mins = _to_minutes_since_start(hhmm, cfg)
    if round_finish == "floor":
        return int(mins // cfg.minutes_per_block)
    return int(math.ceil(mins / cfg.minutes_per_block))

def _block_index_start_inclusive(hhmm: str, cfg: GridConfig) -> int:
    # Uses FLOOR to ensure that if a shift starts at 05:05, 
    # it INCLUDES the 05:00-05:15 block.
    mins = _to_minutes_since_start(hhmm, cfg)
    return int(mins // cfg.minutes_per_block) # floor

def _duration_minutes(start_hhmm: str, finish_hhmm: str, cfg: GridConfig) -> int:
    ms = _to_minutes_since_start(start_hhmm, cfg)
    me = _to_minutes_since_start(finish_hhmm, cfg)
    if me <= ms:
        me += 24 * 60
    return me - ms

def calculate_hourly_counts(all_staff_lists, cfg: GridConfig):
    """
    Returns a list of integer counts (one per hour) representing 
    how many staff are active during that hour across all sections.
    """
    counts = [0] * cfg.total_hours
    
    # Flatten the list of lists into one big list of staff
    all_staff = [s for sublist in all_staff_lists for s in sublist]
    
    for s in all_staff:
        # Convert HH:MM strings to minute integers
        start_min = _to_minutes_since_start(s["start_time"], cfg)
        end_min = _to_minutes_since_start(s["finish_time"], cfg)
        
        # Handle overnight shifts (e.g. 23:00 to 01:00)
        if end_min <= start_min: 
            end_min += 24 * 60
            
        # Check specific overlap with each hour slot
        for h in range(cfg.total_hours):
            hour_start = h * 60
            hour_end = (h + 1) * 60
            
            # If the shift overlaps with this hour window (Standard Interval Logic)
            # max(start, h_start) < min(end, h_end)
            if max(start_min, hour_start) < min(end_min, hour_end):
                counts[h] += 1
                
    return counts

# ================= Colour choice (grade first, else duration) =================
def choose_fill_color(section: str, grade: str | None, start_hhmm: str, finish_hhmm: str, cfg: GridConfig) -> str:
    pal = PALETTES[section]
    deep, light = pal["deep"], pal["light"]

    g = (grade or "").strip().upper()
    if g == "CSA1":
        return deep
    if g == "CSA2":
        return light

    mins = _duration_minutes(start_hhmm, finish_hhmm, cfg)
    # tolerance ±7min to account for small offsets
    if abs(mins - 480) <= 7:  # 8h
        return deep
    if abs(mins - 450) <= 7:  # 7.5h
        return light
    return deep if mins >= 465 else light

def draw_time_header(ws, top_row: int, cfg: GridConfig):
    """
    Draws a merged hour header row at 'top_row'.
    UPDATED: 
    1. Applies Thick Top/Bottom borders to Columns A-E so the line starts from the beginning.
    2. Iterates through Time columns to apply borders properly.
    """
    thick = Side(border_style="thick", color="000000")
    thin  = Side(border_style="thin",  color="000000")
    header_font = Font(bold=True, size=10)
    center = Alignment(horizontal="center", vertical="center")

    # --- 1. Draw Columns A-E (The "Beginning") ---
    # We treat columns 1 to (start_column - 1) as a spacer block
    if cfg.start_column > 1:
        # Merge A-E for a clean look
        ws.merge_cells(start_row=top_row, start_column=1, end_row=top_row, end_column=cfg.start_column - 1)
        
        # Apply borders to every cell in this A-E range to ensure the line is solid
        for col in range(1, cfg.start_column):
            cell = ws.cell(row=top_row, column=col)
            
            # Thick Left on Col A (1)
            style_left = thick if col == 1 else None
            # Thick Right on Col E (Start-1) to meet the Time Grid
            style_right = thick if col == (cfg.start_column - 1) else None
            
            # Thick Top/Bottom for the continuous bar
            cell.border = Border(top=thick, bottom=thick, left=style_left, right=style_right)

    # --- 2. Draw Time Columns (F onwards) ---
    for h in range(cfg.total_hours):
        c0 = cfg.start_column + h * cfg.blocks_per_hour
        c1 = c0 + cfg.blocks_per_hour - 1
        
        # Merge cells for the hour
        ws.merge_cells(start_row=top_row, start_column=c0, end_row=top_row, end_column=c1)
        
        # Set Value
        head = ws.cell(row=top_row, column=c0)
        head.value = f"{(cfg.start_hour + h) % 24:02d}:00"
        head.alignment = center
        head.font = header_font
        
        # Apply Borders to EVERY cell in the merge
        for col in range(c0, c1 + 1):
            cell = ws.cell(row=top_row, column=col)
            
            # Left Border: Thick if it's the 5am start (h=0), else Thin divider
            if col == c0:
                style_left = thick if h == 0 else thin
            else:
                style_left = None 
            
            # Right Border: Thick if it's the last hour, else Thin divider
            if col == c1:
                style_right = thick if h == (cfg.total_hours - 1) else thin
            else:
                style_right = None

            cell.border = Border(
                top=thick, 
                bottom=thick, 
                left=style_left if style_left else cell.border.left,
                right=style_right if style_right else cell.border.right
            )
            
        # Set column width
        for col in range(c0, c1 + 1):
            ws.column_dimensions[get_column_letter(col)].width = 2.2

def draw_clean_grid(ws, row_start: int, row_end: int, cfg: GridConfig, hour_line="thin"):
    """Draw body rows: no inner verticals, only hour boundary lines; light-grey horizontals."""
    hour_side = Side(border_style=hour_line, color="000000")
    grey = Side(border_style="thin", color="000000")
    center = Alignment(horizontal="center", vertical="center")
    # NEW: Define the Empty Grey Fill
    empty_fill = PatternFill(fill_type="solid", start_color="A6A6A6", end_color="A6A6A6")
    for r in range(row_start, row_end + 1):
        ws.row_dimensions[r].height = 18
        for b in range(cfg.total_blocks):
            col = cfg.start_column + b
            is_hour_left  = ((col - cfg.start_column) % cfg.blocks_per_hour) == 0
            is_hour_right = ((col - cfg.start_column) % cfg.blocks_per_hour) == (cfg.blocks_per_hour - 1)
            left  = hour_side if is_hour_left  else None
            right = hour_side if is_hour_right else None
            top   = grey
            bottom= grey
            c = ws.cell(row=r, column=col)
            c.border = Border(left=left, right=right, top=top, bottom=bottom)
            c.alignment = center
            #Apply the Grey Fill
            # c.fill = empty_fill


def paint_shift(ws, row: int, section: str, start_hhmm: str, finish_hhmm: str, grade: str | None, cfg: GridConfig):
    """Paint per 15-min block using section palette + grade/duration colour rule."""
    fill_rgb = choose_fill_color(section, grade, start_hhmm, finish_hhmm, cfg)
    fill = PatternFill(fill_type="solid", start_color=fill_rgb, end_color=fill_rgb)

    s_blk = _block_index_start(start_hhmm, cfg)
    e_blk = _block_index_end(finish_hhmm, cfg, round_finish="ceil")
    if e_blk <= s_blk:
        e_blk += cfg.total_blocks

    s_col = cfg.start_column + max(0, s_blk)
    e_col_excl = cfg.start_column + min(e_blk, cfg.total_blocks)
    for col in range(s_col, e_col_excl):
        ws.cell(row=row, column=col).fill = fill

def merge_painted_segments_per_hour(ws, row: int, cfg: GridConfig):
    """Within each hour (4 cells), merge contiguous painted runs of length ≥2 and center."""
    hour_side = Side(border_style="thin", color="000000")
    grey = Side(border_style="thin", color="000000")
    center = Alignment(horizontal="center", vertical="center")

    for h in range(cfg.total_hours):
        c0 = cfg.start_column + h * cfg.blocks_per_hour
        c1 = c0 + cfg.blocks_per_hour - 1

        painted = []
        for col in range(c0, c1 + 1):
            f = ws.cell(row=row, column=col).fill
            painted.append(f is not None and f.fill_type == "solid")

        if not any(painted):
            continue

        run_start = None
        for idx in range(cfg.blocks_per_hour + 1):  # sentinel
            is_painted = painted[idx] if idx < cfg.blocks_per_hour else False
            if is_painted and run_start is None:
                run_start = idx
            elif (not is_painted) and run_start is not None:
                run_end = idx - 1
                run_len = run_end - run_start + 1
                if run_len >= 2:
                    m0 = c0 + run_start
                    m1 = c0 + run_end
                    ws.merge_cells(f"{get_column_letter(m0)}{row}:{get_column_letter(m1)}{row}")
                    tl = ws.cell(row=row, column=m0)
                    tl.alignment = center
                    left  = hour_side if m0 == c0 else None
                    right = hour_side if m1 == c1 else None
                    tl.border = Border(left=left, right=right, top=grey, bottom=grey)
                run_start = None

# ================= NEW Helper: Draw Bold Outline =================
def draw_section_outline(ws, row_start, row_end, cfg, top_style="thick"):
    """
    Draws a THICK BLACK border around the rectangular grid of the section.
    (Top of row_start, Bottom of row_end, Left of start_col, Right of end_col)
    """
    thick = Side(border_style="thick", color="000000")
    thin  = Side(border_style="thin",  color="000000")
    
    # Use the requested style for the top edge
    top_side = thick if top_style == "thick" else thin

    col_start = 1 
    col_end = cfg.start_column + cfg.total_blocks - 1

    # 1. Top Edge
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row_start, column=c)
        current = cell.border
        cell.border = Border(top=top_side, bottom=current.bottom, left=current.left, right=current.right)

    # 2. Bottom Edge (Always Thick)
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row_end, column=c)
        current = cell.border
        cell.border = Border(top=current.top, bottom=thick, left=current.left, right=current.right)

    # 3. Left Edge (Always Thick)
    for r in range(row_start, row_end + 1):
        cell = ws.cell(row=r, column=col_start)
        current = cell.border
        cell.border = Border(top=current.top, bottom=current.bottom, left=thick, right=current.right)

    # 4. Right Edge (Always Thick)
    for r in range(row_start, row_end + 1):
        cell = ws.cell(row=r, column=col_end)
        current = cell.border
        cell.border = Border(top=current.top, bottom=current.bottom, left=current.left, right=thick)
        
    # 5. Fix Corners
    # Top-Left
    tl = ws.cell(row=row_start, column=col_start)
    tl.border = Border(top=top_side, left=thick, bottom=tl.border.bottom, right=tl.border.right)
    # Top-Right
    tr = ws.cell(row=row_start, column=col_end)
    tr.border = Border(top=top_side, right=thick, bottom=tr.border.bottom, left=tr.border.left)
    # Bottom-Left
    bl = ws.cell(row=row_end, column=col_start)
    bl.border = Border(bottom=thick, left=thick, top=bl.border.top, right=bl.border.right)
    # Bottom-Right
    br = ws.cell(row=row_end, column=col_end)
    br.border = Border(bottom=thick, right=thick, top=br.border.top, left=br.border.left)

def draw_global_banners(ws, cfg: GridConfig):
    """
    Draws the Cyan confirmation banners at Rows 2-5.
    Structure:
      - Rows 2-3 merged (Top)
      - Rows 4-5 merged (Bottom)
      - Thick Border surrounding the whole block (A2 -> End:5)
      - Horizontal line separating Row 3 and 4
      - Vertical lines separating every hour
    """
    text_color = "FFFFFF"
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill = PatternFill(fill_type="solid", start_color=PALETTES["Victoria"]["deep"], end_color=PALETTES["Victoria"]["deep"])
    font = Font(bold=True, size=12, color = text_color)
    thick = Side(border_style="thick", color="000000")
    thin  = Side(border_style="thin",  color="000000")
# --- 1. Setup & Merges ---
    
    # Text Labels (A2:E3, A4:E5)
    ws.merge_cells("A2:E3")
    ws["A2"].value = "Confirmation of Security Check started:"
    ws["A2"].alignment = center
    ws["A2"].fill = fill
    ws["A2"].font = font
    
    ws.merge_cells("A4:E5")
    ws["A4"].value = "Confirmation of Security Check completed:"
    ws["A4"].alignment = center
    ws["A4"].fill = fill
    ws["A4"].font = font

    # Merge Hour Boxes (F onwards)
    for h in range(cfg.total_hours):
        c0 = cfg.start_column + h * cfg.blocks_per_hour
        c1 = c0 + cfg.blocks_per_hour - 1
        ws.merge_cells(start_row=2, start_column=c0, end_row=3, end_column=c1)
        ws.merge_cells(start_row=4, start_column=c0, end_row=5, end_column=c1)

    # 2. Draw INTERNAL dividers (Row 3 bottom / Row 4 top) and Vertical Hour lines
    last_col = cfg.start_column + cfg.total_blocks - 1
    
    for row in range(2, 6): # Rows 2, 3, 4, 5
        for col in range(1, last_col + 1):
            cell = ws.cell(row=row, column=col)
            current = cell.border
            
            # Vertical Hour Separators (Internal only)
            style_left = None
            if col > 1: # Skip outer left A
                rel_idx = col - cfg.start_column
                # If this column is the start of an hour block
                if rel_idx >= 0 and rel_idx % cfg.blocks_per_hour == 0:
                    style_left = thin

            # Horizontal Divider (Between 3 and 4)
            style_bottom = None
            if row == 3: 
                style_bottom = thin
            elif row == 5:
                # Thin line at bottom of headers to separate from Metadata
                style_bottom = thin
            
            # Apply strict updates to internal borders
            new_left = style_left if style_left else current.left
            new_bottom = style_bottom if style_bottom else current.bottom
            
            cell.border = Border(
                left=new_left, 
                right=current.right, 
                top=current.top, 
                bottom=new_bottom
            )
            
def draw_meta_header(ws, row: int, cfg: GridConfig, hourly_counts: list[int] = None):
    """
    Draws column headers at Row 6.
    UPDATED: Displays hourly staff counts in the grid area with color coding.
    """
    headers = ["Duty No.", "Start Time", "Finish Time", "Name", "Radio"]
    font_text = Font(bold=True, size=7)
    font_num  = Font(bold=True, size=10) # Larger font for the count
    center = Alignment(horizontal="center", vertical="center")
    
    thick = Side(border_style="thick", color="000000")
    thin  = Side(border_style="thin",  color="000000")

    # Define Colors
    RED_FILL    = PatternFill("solid", start_color="FFC00000")
    YELLOW_FILL = PatternFill("solid", start_color="FFC000")
    GREEN_FILL  = PatternFill("solid", start_color="00B050")

    last_col = cfg.start_column + cfg.total_blocks - 1

    # 1. Draw Text Headers (Cols A-E)
    for i, text in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i)
        c.value = text
        c.font = font_text
        c.alignment = center
        # Borders for A-E
        style_left = thick if i == 1 else None
        c.border = Border(top=thin, bottom=thin, left=style_left, right=c.border.right)

    # 2. Draw Hourly Counts (Cols F onwards)
    # We iterate by HOUR (not block) to merge cells
    for h in range(cfg.total_hours):
        c0 = cfg.start_column + h * cfg.blocks_per_hour
        c1 = c0 + cfg.blocks_per_hour - 1
        
        # Merge the 4 cells of the hour
        ws.merge_cells(start_row=row, start_column=c0, end_row=row, end_column=c1)
        cell = ws.cell(row=row, column=c0)
        
        # Set Count Value
        if hourly_counts:
            count = hourly_counts[h]
            cell.value = count
            
            # Apply Color Logic
            if count < 9:
                cell.fill = RED_FILL
            elif count > 11:
                cell.fill = GREEN_FILL
            else:
                # Covers 9, 10, 11
                cell.fill = YELLOW_FILL
        
        cell.font = font_num
        cell.alignment = center
        
        # Apply Borders to the merged block
        # Left is Thin (unless it's the very first hour?)
        # Let's keep strict vertical dividers for hours
        style_left = None # Default
        style_right = thin # Always separate hours
        
        # Special case: Far Left of grid needs standard thin/thick handling
        if h == 0: 
            # The border between 'Radio' and '05:00' is usually handled by 'Radio's right
            pass 

        # Apply to all cells in merge to ensure box is closed
        for col in range(c0, c1 + 1):
            c = ws.cell(row=row, column=col)
            # Top/Bottom Thin to blend with Header/Grid
            # Right Thick if it's the very last column of the sheet
            s_right = thick if col == last_col else (thin if col == c1 else None)
            
            c.border = Border(top=thin, bottom=thin, right=s_right, left=c.border.left)

def fill_gaps_grey(ws, start_row: int, end_row: int, cfg: GridConfig):
    """
    Post-processing step:
    Scans the grid area. Any cell that has NO background color (None)
    gets filled with Grey (A6A6A6).
    """
    grey_fill = PatternFill(fill_type="solid", start_color="A6A6A6", end_color="A6A6A6")
    
    # Iterate every cell in the grid for this section
    for r in range(start_row, end_row + 1):
        for b in range(cfg.total_blocks):
            col = cfg.start_column + b
            cell = ws.cell(row=r, column=col)
            
            # Check if cell has a fill. 
            # OpenPyXL cells usually have a fill object, checking start_color is safer.
            # If it's '000000' or None, it's empty.
            # (Standard default fill is often type='none' or color '000000')
            
            has_color = False
            if cell.fill and cell.fill.fill_type == "solid":
                # Check if it's a real color (not black/white placeholder if using default)
                # But our paint_shift uses solid colors, so if it is solid, it's a shift.
                has_color = True
            
            if not has_color:
                cell.fill = grey_fill
                
def add_victoria_cssi_note(ws, cfg, body_start: int, body_end: int, text: str = None):
    """
    Places a purple merged note centered horizontally within the Victoria section.
    """
    if text is None:
        text = "PLEASE NOTE: CSSI INDICATES\nTHAT YOU MUST BE ON THE\nSTATION DURING THIS TIME."

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(left=BLACK_THIN, right=BLACK_THIN, top=BLACK_THIN, bottom=BLACK_THIN)
    # >>> larger black text
    font   = Font(bold=True, color="000000", size=14)

    # --- vertical position: 3 rows high, close to bottom ---
    note_height = 3
    note_row_top = max(body_start + 1, body_end - note_height - 1)
    note_row_bottom = note_row_top + note_height - 1

    # --- horizontal position: 6 hours wide, starting at an hour boundary ---
    hours_span = 5
    total_hours = cfg.total_hours

    # centre-ish start hour, then clamp within [0, total_hours - hours_span]
    mid_hour_index = total_hours // 2
    start_hour_index = max(0, min(total_hours - hours_span, mid_hour_index - hours_span // 2))

    c0 = cfg.start_column + start_hour_index * cfg.blocks_per_hour
    c1 = c0 + hours_span * cfg.blocks_per_hour - 1

    ws.merge_cells(
        f"{get_column_letter(c0)}{note_row_top}:"
        f"{get_column_letter(c1)}{note_row_bottom}"
    )
    cell = ws.cell(row=note_row_top, column=c0)
    cell.value = text
    cell.alignment = center
    cell.fill = PatternFill(fill_type="solid", start_color=PURPLE, end_color=PURPLE)
    cell.font = font
    cell.border = border

def draw_internal_section_label(ws, start_row: int, section_name: str, cfg: GridConfig):
    """
    Draws the section banner INSIDE the grid.
    Includes SAFETY FIX to unmerge underlying cells first.
    """
    # Define Colors
    color_map = {
        "VICTORIA": "00B0F0", 
        "DISTRICT": "33CC33", 
        "CARDINAL": "FE66CC", 
        "NORTH":    "FE66CC"  
    }
    
    bg_color = "000000"
    text_color = "FFFFFF"
    
    upper_name = section_name.upper()
    for key, code in color_map.items():
        if key in upper_name:
            bg_color = code
            break
            
    fill = PatternFill(fill_type="solid", start_color=bg_color, end_color=bg_color)
    font = Font(bold=True, size=24, color=text_color)
    center = Alignment(horizontal="center", vertical="center")
    
    # Calculate Position
    hours_span = 6 
    blocks_span = hours_span * cfg.blocks_per_hour
    grid_end_col = cfg.start_column + cfg.total_blocks - 1
    
    c_start = max(cfg.start_column, grid_end_col - blocks_span + 1)
    c_end = grid_end_col
    r_start = start_row
    r_end = start_row + 1 
    
    # --- CRITICAL FIX: Clean the area first ---
    clean_merge_area(ws, r_start, c_start, r_end, c_end)
    # ------------------------------------------

    # Now it is safe to merge
    ws.merge_cells(start_row=r_start, start_column=c_start, end_row=r_end, end_column=c_end)
    
    cell = ws.cell(row=r_start, column=c_start)
    cell.value = section_name.upper()
    cell.fill = fill
    cell.font = font
    cell.alignment = center
    
    # Apply style to the whole block so edges look right
    for r in range(r_start, r_end + 1):
        for c in range(c_start, c_end + 1):
            c_obj = ws.cell(row=r, column=c)
            c_obj.fill = fill

def add_victoria_right_legend_in_grid(ws, cfg, body_start: int, labels=None):
    """
    Draw the Victoria legend.
    Includes SAFETY FIX to unmerge underlying cells first.
    """
    if labels is None:
        labels = [
            ("TOP OF ESC\n4–6", GREY,   "AREA AT THE TOP OF ESC\n6", GREY),
            ("VIP/MIP",     YELLOW, "PERSON ALLOCATED TO ASSIST\nWITH VIP/MIPs", YELLOW),
            ("POM SERVICING", ORANGE, "SERVICE ALL POMS",           ORANGE),
            ("PLATFORM?",     GREEN,    "PLATFORM SATS",              GREEN),
            ("PTI",           PEACH,  "PTI DUTY",                   PEACH),
        ]

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(left=BLACK_THIN, right=BLACK_THIN, top=BLACK_THIN, bottom=BLACK_THIN)
    font   = Font(bold=True, size=8)

    legend_start_hour = cfg.end_hour - 3          
    hour_offset = legend_start_hour - cfg.start_hour

    # Columns
    c_title0 = cfg.start_column + hour_offset * cfg.blocks_per_hour
    c_title1 = c_title0 + cfg.blocks_per_hour - 1          
    c_desc0 = c_title1 + 1
    c_desc1 = c_desc0 + 2 * cfg.blocks_per_hour - 1        

    row_start = body_start + 3

    for i, (l_text, l_col, r_text, r_col) in enumerate(labels):
        r = row_start + i

        # --- CRITICAL FIX: Clean the Title area ---
        clean_merge_area(ws, r, c_title0, r, c_title1)
        ws.merge_cells(start_row=r, start_column=c_title0, end_row=r, end_column=c_title1)
        
        lc = ws.cell(row=r, column=c_title0)
        lc.value = l_text
        lc.alignment = center
        lc.fill = PatternFill(fill_type="solid", start_color=l_col, end_color=l_col)
        lc.font = font
        lc.border = border

        # --- CRITICAL FIX: Clean the Description area ---
        clean_merge_area(ws, r, c_desc0, r, c_desc1)
        ws.merge_cells(start_row=r, start_column=c_desc0, end_row=r, end_column=c_desc1)
        
        rc = ws.cell(row=r, column=c_desc0)
        rc.value = r_text
        rc.alignment = center
        rc.fill = PatternFill(fill_type="solid", start_color=r_col, end_color=r_col)
        rc.font = font
        rc.border = border

# ================= Build the SINGLE "Roster" sheet with 3 stacked sections =================
def build_single_sheet(
    out_filename: str,
    staff_victoria: list[dict],
    staff_district: list[dict],
    staff_cardinal: list[dict],
    cfg: GridConfig,
    rows_per_section=(28, 28, 13)  # visible staff rows (approx) for Vic/District/Cardinal
):
    wb = Workbook()
    ws = wb.active
    ws.title = "Roster"

    # Set meta column widths A..E
    for i, w in enumerate(cfg.meta_col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # --- NEW: Calculate Total Staff per Hour ---
    total_counts = calculate_hourly_counts(
        [staff_victoria, staff_district, staff_cardinal], 
        cfg
    )
    #Draw top layout(Time Header)
    draw_time_header(ws, 1, cfg)
    # Draw the fixed banners at the top-left
    draw_global_banners(ws, cfg)
    # Draw the Metadata
    draw_meta_header(ws, row=6, cfg=cfg, hourly_counts=total_counts)

    # Helper to render one stacked section
    def render_section(section_name: str, staff_rows: list[dict], start_row: int, fixed_count: int, is_first=False, custom_header_row: int = None) -> int:
        
        # Calculate the exact bottom of this section based on the FIXED count
        end_row = start_row + fixed_count - 1

        # 1. Draw the empty grid structure for ALL 50/30/20 rows
        draw_clean_grid(ws, start_row, end_row, cfg, hour_line="thin")

        # 2. Fill rows (Staff data if available, otherwise just Duty ID)
        for i in range(fixed_count):
            r = start_row + i
            
            # Check if we have a staff member for this slot
            s = staff_rows[i] if i < len(staff_rows) else None
            
            # Metadata Styling
            font_meta = Font(size=10, bold=True)
            center_align = Alignment(horizontal="center", vertical="center")
            
            # Column A: Always generate an ID (e.g. VIC01... VIC50)
            # Use the first 2 letters of section (VI, DI, CA) or logic mapping
            prefix_map = {"Victoria": "BN", "District": "BN", "Cardinal": "BN"}
            prefix = prefix_map.get(section_name, "XX")
            
            # If we have staff data, use their ID, otherwise generate a generic one?
            # Or strictly use generated sequence VIC01-VIC50? 
            # Usually strict sequence is better for fixed templates:
            # ws[f"A{r}"] = f"{section_name[:2].upper()}{i+1:02d}" 
            
            # However, if you want to keep the specific staff 'duty code' (like BN71):
            if s:
                ws[f"A{r}"] = s.get("id", "")
                ws[f"B{r}"] = s["start_time"]
                ws[f"C{r}"] = s["finish_time"]
                ws[f"D{r}"] = s["name"]
                ws[f"E{r}"] = s.get("radio", "")
                
                # PAINT the shift bars only if staff exists
                paint_shift(ws, r, section_name, s["start_time"], s["finish_time"], 
                            s.get("grade"), cfg)
                merge_painted_segments_per_hour(ws, r, cfg)
            else:
                # EMPTY SLOT
                # Leave B, C, D, E blank, but format A
                ws[f"A{r}"] = "" # Or generic ID like f"{prefix}.." if you prefer
            
            # Apply Style to Metadata columns A-E (even if empty, ensures borders/alignment)
            for col in "ABCDE":
                cell = ws[f"{col}{r}"]
                cell.font = font_meta
                cell.alignment = center_align
        
        # This checks for empty cells and colors them grey
        fill_gaps_grey(ws, start_row, end_row, cfg)
        
        # # 3. Draw section banner at the very bottom of the fixed block
        # 3. Draw Internal Section Label (The "Banner")
        # Top-Right of the section data
        draw_internal_section_label(ws, start_row, section_name, cfg)
        border_start = custom_header_row if custom_header_row is not None else start_row


        # Extra notes ONLY for Victoria (placed inside the grid area)
        if section_name == "Victoria":
            add_victoria_cssi_note(ws, cfg, start_row, end_row)
            add_victoria_right_legend_in_grid(ws, cfg, start_row)
        top_style = "thin" if is_first else "thick"
        draw_section_outline(ws, border_start, end_row, cfg, top_style="thick")     

        # Return the row for the NEXT item
        return end_row + 1
# 1. Victoria (Starts at Row 7)
    row_cursor = 7
    row_cursor = render_section("Victoria", staff_victoria, row_cursor, rows_per_section[0], custom_header_row=2)
    row_cursor += 1 

    # 2. District (Insert Time Header First)
    draw_time_header(ws, row_cursor, cfg)  # Draw header at current cursor
    row_cursor += 1                        # Step down 1 row so data doesn't overwrite header
    # Meta Header (Inserted Step)
    draw_meta_header(ws, row_cursor, cfg)
    meta_start_dist = row_cursor # Remember this row for the outline
    row_cursor += 1

    row_cursor = render_section("District", staff_district, row_cursor, rows_per_section[1], custom_header_row=meta_start_dist)
    row_cursor += 1 

    # 3. Cardinal (Insert Time Header First)
    draw_time_header(ws, row_cursor, cfg)  # Draw header at current cursor
    row_cursor += 1                        # Step down 1 row
    # Meta Header (Inserted Step)
    draw_meta_header(ws, row_cursor, cfg)
    meta_start_card = row_cursor
    row_cursor += 1
    row_cursor = render_section("Cardinal", staff_cardinal, row_cursor, rows_per_section[2], custom_header_row=meta_start_card)

    # Freeze panes
    # ws.freeze_panes = "F7"

    # Freeze pane just under the first grid header
    # ws.freeze_panes = ws.cell(row=rows_per_section[0] + 3, column=cfg.start_column)

    os.makedirs("output", exist_ok=True)
    path = f"output/{out_filename}"
    wb.save(path)
    return path

def generate_victoria_duty_slots_full(
    open_time="05:00",
    close_time="00:00",
    include_night=False,
):
    """
    Build ALL duty slots for a 'normal' day at Victoria (single date).
    Returns a list of:
      {
        'label': str,         # text to show in the cell (e.g. "Security Check", "PTI EB")
        'role': str,          # logical role (SECURITY, PTI, SATS, ESCALATOR, VIP, P3)
        'start': int,         # minutes since midnight
        'end': int,
        'required_grade': 'CSA1' | 'ANY',
        'priority': int,      # lower = scheduled first
      }

    Assumptions:
      - Times are for a Mon–Thu style “normal” day (station closes at 00:00).
      - Security checks: hourly from open to close (or 24h if include_night=True).
      - PTI EB/WB: 07:30–23:00, 1h blocks, CSA1 only.
      - SATS EB/WB: 07–09 & 17–19, 1h blocks, CSA1 only, WB has higher priority.
      - P3 (Victoria northbound): 07:30–09:00 & 17–19, 1h blocks, CSA1 only.
      - Top of ESC 4–6: 07:00–23:00, 1h blocks, any grade.
      - VIP/MIP: 07:00–22:00, 1h blocks, any grade (continuous coverage).
    """
    slots = []

    def add_slot(label, role, start_min, end_min, grade_req="ANY"):
        slots.append(
            {
                "label": label,
                "role": role,
                "start": start_min,
                "end": end_min,
                "required_grade": grade_req,
                "priority": DUTY_PRIORITY[role],
            }
        )

    # ----- time base -----
    open_min = hhmm_to_minutes(open_time)
    close_min = hhmm_to_minutes(close_time)
    if close_min <= open_min:
        close_min += 24 * 60  # treat as after midnight

    # ---------- 1) SECURITY CHECKS ----------
    # Every hour from open to close (or all 24h if include_night=True)
    if include_night:
        sec_start = 0
        sec_end = 24 * 60
    else:
        sec_start = open_min
        sec_end = close_min

    t = sec_start
    while t + 60 <= sec_end:
        add_slot("Security Check", "SECURITY", t, t + 60, grade_req="ANY")
        t += 60

    # ---------- 2) TOP OF ESC 4–6 ----------
    # 07:00–23:00, 1h blocks
    esc_start = hhmm_to_minutes("07:00")
    esc_end = hhmm_to_minutes("23:00")
    if esc_end <= esc_start:
        esc_end += 24 * 60

    t = esc_start
    while t + 60 <= esc_end:
        add_slot("TOP OF ESC 4-6", "ESCALATOR", t, t + 60, grade_req="ANY")
        t += 60

    # ---------- 3) PTI WB / EB ----------
    # 07:30–23:00, 1h blocks, CSA1 only, both directions
    pti_start = hhmm_to_minutes("07:30")
    pti_end = hhmm_to_minutes("23:00")
    if pti_end <= pti_start:
        pti_end += 24 * 60

    t = pti_start
    while t + 60 <= pti_end:
        add_slot("PTI WB", "PTI WB", t, t + 60, grade_req="CSA1")
        add_slot("PTI EB", "PTI EB", t, t + 60, grade_req="CSA1")
        t += 60

    # ---------- 4) SATS WB / EB ----------
    # Traffic windows 07–09 and 17–19, CSA1 only.
    # WB has higher priority (role names give different PRIORITY).
    sats_windows = [("07:00", "09:00"), ("17:00", "19:00")]

    for start_str, end_str in sats_windows:
        w_start = hhmm_to_minutes(start_str)
        w_end = hhmm_to_minutes(end_str)
        if w_end <= w_start:
            w_end += 24 * 60

        t = w_start
        while t + 60 <= w_end:
            # Prioritise WB by giving it role "SATS WB" (higher priority)
            add_slot("SATS WB", "SATS WB", t, t + 60, grade_req="CSA1")
            add_slot("SATS EB", "SATS EB", t, t + 60, grade_req="CSA1")
            t += 60

    # ---------- 5) VIP/MIP ----------
    # All-day coverage-ish: 07:00–22:00, 1h blocks, any grade
    vip_start = hhmm_to_minutes("07:00")
    vip_end = hhmm_to_minutes("22:00")
    if vip_end <= vip_start:
        vip_end += 24 * 60

    t = vip_start
    while t + 60 <= vip_end:
        add_slot("VIP/MIP", "VIP/MIP", t, t + 60, grade_req="ANY")
        t += 60

    # ---------- 6) P3 (Victoria NB platform) ----------
    # 07:30–09:00 & 17:00–19:00, 1h blocks, CSA1 only
    p3_windows = [("07:30", "09:00"), ("17:00", "19:00")]
    for start_str, end_str in p3_windows:
        p_start = hhmm_to_minutes(start_str)
        p_end = hhmm_to_minutes(end_str)
        if p_end <= p_start:
            p_end += 24 * 60

        t = p_start
        while t + 60 <= p_end:
            add_slot("P3", "P3", t, t + 60, grade_req="CSA1")
            t += 60

    # final sort: by start, then priority (P3, PTI, SATS WB, SATS EB, ESC, VIP, SEC)
    slots.sort(key=lambda s: (s["start"], s["priority"], s["label"]))
    return slots

# ---------- duty assignment to Victoria staff ----------

def _build_staff_for_assign(staff_list):
    enriched = []
    for s in staff_list:
        s2 = s.copy()
        s2["start_min"] = hhmm_to_minutes(s["start_time"])
        s2["end_min"]   = hhmm_to_minutes(s["finish_time"])
        if s2["end_min"] <= s2["start_min"]:
            s2["end_min"] += 24 * 60
        s2["duties"] = []       # list of slots
        s2["duty_load_min"] = 0
        enriched.append(s2)
    return enriched

def _is_free_for(s, start, end):
    for d in s["duties"]:
        if not (end <= d["start"] or start >= d["end"]):
            return False
    return True

def assign_victoria_duties(victoria_staff, cfg, include_night=False):
    """
    Attach a 'duties' list to each Victoria staff member.
    Grade requirements are respected (CSA1-only duties).
    Returns (enriched_staff, unassigned_slots).
    """
    staff = _build_staff_for_assign(victoria_staff)
    slots = generate_victoria_duty_slots_full(
        open_time="05:00",
        close_time="00:00",
        include_night=include_night,
    )
    unassigned = []

    for slot in slots:
        start, end = slot["start"], slot["end"]
        need_csa1 = (slot["required_grade"] == "CSA1")

        # find eligible staff
        candidates = []
        for s in staff:
            if need_csa1 and s.get("grade", "").upper() != "CSA1":
                continue
            if not (s["start_min"] <= start and s["end_min"] >= end):
                continue
            if not _is_free_for(s, start, end):
                continue
            candidates.append(s)

        if not candidates:
            unassigned.append(slot)
            continue

        chosen = min(candidates, key=lambda s: (s["duty_load_min"], s["start_min"]))
        chosen["duties"].append(slot)
        chosen["duty_load_min"] += (end - start)

    # sort each staff's duties by time
    for s in staff:
        s["duties"].sort(key=lambda d: d["start"])

    return staff, unassigned

# ================= MAIN EXECUTION BLOCK =================
if __name__ == "__main__":
    import sys
    
# 1. Setup Defaults (for testing without Excel)
    default_pdf = "2025.10.25 VSOU SOS.pdf"
    default_date = "Monday20October2025" # Fallback if no date provided

    # 2. Get Arguments from Excel
    # sys.argv[1] = PDF Path
    # sys.argv[2] = Target Date String
    
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = default_pdf

    if len(sys.argv) > 2:
        target_date = sys.argv[2]
    else:
        target_date = default_date

    # 3. Define Output Path
    # We will name the file based on the date so you don't overwrite old rosters
    output_dir = os.path.dirname(pdf_path) if os.path.exists(pdf_path) else "."
    
    # Sanitize date for filename (remove spaces/slashes)
    safe_date = "".join(c for c in target_date if c.isalnum())
    output_path = os.path.join(output_dir, f"Roster_{safe_date}.xlsx")

    print(f"--- Roster Generator ---")
    print(f"PDF: {pdf_path}")
    print(f"Date: {target_date}")
    
    # 4. Run Extraction Logic
    # Pass the user's date to the extractor
    rows = extract_victoria_rows_from_text(pdf_path, date_str=target_date) 
    
    if not rows:
        print(f"No rows found for date: {target_date}")
        print("Please check the date format matches the PDF (e.g. 'Monday 20 October 2025')")
    else:
        victoria_raw = build_staff_from_pdf_rows(rows)
        victoria, district, cardinal = split_by_duty_section(victoria_raw)
        
        # Print stats for debugging
        print(f"Found {len(victoria)} Vic, {len(district)} Dist, {len(cardinal)} Card")

        cfg = GridConfig(start_hour=5, end_hour=25, start_column=6, blocks_per_hour=4)
        
        # Build the file
        saved_file = build_single_sheet(
            output_path, 
            victoria, district, cardinal,
            cfg,
            rows_per_section=(30, 24, 20)
        )
        print(f"Success! Saved to: {saved_file}")
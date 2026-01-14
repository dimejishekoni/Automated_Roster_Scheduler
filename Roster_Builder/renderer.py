import os
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
from openpyxl.utils import get_column_letter
from config import GridConfig
from scheduler import (
    STATUS_GATELINE, STATUS_BREAK, STATUS_CSSI, STATUS_SECURITY,
    STATUS_SPECIALIST, STATUS_GENERAL, STATUS_PLATFORM)
from config import PALETTES, COLORS, BLACK_THIN
from helper import (clean_merge_area, calculate_hourly_counts,
                    block_index_start, block_index_end, duration_minutes, draw_section_outline)

# ================= Colour choice (grade first, else duration) =================
def choose_fill_color(section: str, grade: str | None, start_hhmm: str, finish_hhmm: str, cfg: GridConfig) -> str:
    pal = PALETTES[section]
    deep, light = pal["deep"], pal["light"]

    g = (grade or "").strip().upper()
    if g == "CSA1":
        return deep
    if g == "CSA2":
        return light

    mins = duration_minutes(start_hhmm, finish_hhmm, cfg)
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

    if cfg.start_column > 1:
        ws.merge_cells(start_row=top_row, start_column=1, end_row=top_row, end_column=cfg.start_column - 1)
        
        for col in range(1, cfg.start_column):
            cell = ws.cell(row=top_row, column=col)
            
            style_left = thick if col == 1 else None
            style_right = thick if col == (cfg.start_column - 1) else None
            
            cell.border = Border(top=thick, bottom=thick, left=style_left, right=style_right)

    for h in range(cfg.total_hours):
        c0 = cfg.start_column + h * cfg.blocks_per_hour
        c1 = c0 + cfg.blocks_per_hour - 1
        ws.merge_cells(start_row=top_row, start_column=c0, end_row=top_row, end_column=c1)
        
        # Set Value
        head = ws.cell(row=top_row, column=c0)
        head.value = f"{(cfg.start_hour + h) % 24:02d}:00"
        head.alignment = center
        head.font = header_font
        
        for col in range(c0, c1 + 1):
            cell = ws.cell(row=top_row, column=col)
            
            if col == c0:
                style_left = thick if h == 0 else thin
            else:
                style_left = None 
            
            if col == c1:
                style_right = thick if h == (cfg.total_hours - 1) else thin
            else:
                style_right = None

            cell.border = Border(
                top=thick, 
                bottom=thick, 
                left=style_left if style_left else cell.border.left,
                right=style_right if style_right else cell.border.right)
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

def paint_shift(ws, row: int, section: str, s: dict, cfg: GridConfig):
    if not s: return

    # --- SETUP COLORS ---
    base_rgb = choose_fill_color(section, s.get("grade"), s["start_time"], s["finish_time"], cfg)
    base_fill = PatternFill(fill_type="solid", start_color=base_rgb, end_color=base_rgb)
    
    break_fill   = PatternFill("solid", start_color=COLORS["BREAK"])  # 60497A
    cssi_fill    = PatternFill("solid", start_color=COLORS["CSSI"])   # 9966FF
    special_fill = PatternFill("solid", start_color=COLORS["YELLOW"])
    security_fill = PatternFill("solid", start_color=COLORS["SECURITY"]) # Red
    general_fill = PatternFill("solid", start_color=COLORS["GENERAL"])
    grey_fill = PatternFill("solid", start_color=COLORS["GREY"])
    platform_fill = PatternFill("solid", start_color=COLORS["GREEN"])
    # Fonts
    white_font = Font(size=8, bold=False, color="FFFFFF")
    black_font = Font(size=8, bold=False, color="000000")
    black_font_10 = Font(size=10, bold=False, color="000000")
    black_font_platforms = Font(size=13, bold=True, color="000000")


    matrix = s.get("status_matrix")
    text_matrix = s.get("text_matrix") 

    if matrix:
        for b in range(cfg.total_blocks):
            col = cfg.start_column + b
            
            if b < len(matrix):
                status = matrix[b]
                cell = ws.cell(row=row, column=col)
                
                # Apply Color & FORCE Text based on Status
                if status == STATUS_GATELINE:
                    cell.fill = base_fill
                    
                elif status == STATUS_BREAK:
                    cell.fill = break_fill
                    cell.font = white_font
                    cell.value = "Meal Break"   # <--- Hardcoded Fallback
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    
                elif status == STATUS_CSSI:
                    cell.fill = cssi_fill
                    cell.font = white_font 
                    cell.value = "CSSI"         
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                
                elif status == STATUS_SECURITY:
                    cell.fill = security_fill
                    cell.font = black_font
                    cell.value = "Security Check"
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


                elif status == STATUS_GENERAL:
                    cell.fill = grey_fill           
                    cell.font = black_font          
                    cell.value = "Top of ESC 4-6"   
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                elif status == STATUS_SPECIALIST:
                    cell.fill = special_fill
                    cell.font = black_font_10
                    cell.value = "VIP/MIP"

                # --- NEW BLOCK FOR P1/P2/P3 ---
                elif status == STATUS_PLATFORM:
                    cell.fill = platform_fill      # COLORS["GREEN"]
                    cell.font = black_font_platforms         # Black Text
                
                # If specific text exists (e.g. from Scheduler), overwrite the default
                if text_matrix and b < len(text_matrix) and text_matrix[b]:
                    cell.value = text_matrix[b]
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


    else:
        # Fallback for staff without matrix
        # Note: Updated function names (no underscore)
        s_blk = block_index_start(s["start_time"], cfg)
        e_blk = block_index_end(s["finish_time"], cfg, round_finish="ceil")
        
        if e_blk <= s_blk: e_blk += cfg.total_blocks
        s_col = cfg.start_column + max(0, s_blk)
        e_col_excl = cfg.start_column + min(e_blk, cfg.total_blocks)
        for col in range(s_col, e_col_excl):
            if col > 0: ws.cell(row=row, column=col).fill = base_fill    

def merge_painted_segments_per_hour(ws, row: int, cfg: GridConfig):
    """
    Scans each hour (4 blocks).
    Merges contiguous blocks ONLY if they have the same Color and Text.
    This allows a 30m Break (Blue) and 30m CSSI (Purple) to exist in the same hour slot.
    """
    hour_side = Side(border_style="thin", color="000000")
    grey = Side(border_style="thin", color="000000")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True) # Ensure wrap text

    for h in range(cfg.total_hours):
        c0 = cfg.start_column + h * cfg.blocks_per_hour
        c_end_of_hour = c0 + cfg.blocks_per_hour - 1
        
        current_run_start = c0
        
        for col in range(c0, c_end_of_hour + 2):
            
            if col <= c_end_of_hour:
                curr_cell = ws.cell(row=row, column=col)
                curr_fill = curr_cell.fill.start_color.rgb if (curr_cell.fill and curr_cell.fill.fill_type=="solid") else None
                curr_text = curr_cell.value
            else:
                curr_fill = "SENTINEL"
                curr_text = "SENTINEL"

            start_cell = ws.cell(row=row, column=current_run_start)
            start_fill = start_cell.fill.start_color.rgb if (start_cell.fill and start_cell.fill.fill_type=="solid") else None
            start_text = start_cell.value

            is_match = (curr_fill == start_fill) and (curr_text == start_text)
            
            if not is_match:
                run_end = col - 1
                run_length = run_end - current_run_start + 1
                
                if run_length >= 2 and start_fill is not None:
                    
                    # Perform Merge
                    ws.merge_cells(start_row=row, start_column=current_run_start, end_row=row, end_column=run_end)
                    tl = ws.cell(row=row, column=current_run_start)
                    tl.alignment = center
                    left_border = hour_side if current_run_start == c0 else None
                    right_border = hour_side if run_end == c_end_of_hour else None
                    
                    tl.border = Border(left=left_border, right=right_border, top=grey, bottom=grey)

                current_run_start = col


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
    last_col = cfg.start_column + cfg.total_blocks - 1
    
    for row in range(2, 6): 
        for col in range(1, last_col + 1):
            cell = ws.cell(row=row, column=col)
            current = cell.border
            style_left = None
            if col > 1: 
                rel_idx = col - cfg.start_column
                if rel_idx >= 0 and rel_idx % cfg.blocks_per_hour == 0:
                    style_left = thin

            style_bottom = None
            if row == 3: 
                style_bottom = thin
            elif row == 5:
                style_bottom = thin
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

    for i, text in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i)
        c.value = text
        c.font = font_text
        c.alignment = center
        style_left = thick if i == 1 else None
        c.border = Border(top=thin, bottom=thin, left=style_left, right=c.border.right)

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
                cell.fill = YELLOW_FILL
        
        cell.font = font_num
        cell.alignment = center
        
        style_left = None
        style_right = thin 
        if h == 0: 
            pass 
        for col in range(c0, c1 + 1):
            c = ws.cell(row=row, column=col)
            s_right = thick if col == last_col else (thin if col == c1 else None)
            
            c.border = Border(top=thin, bottom=thin, right=s_right, left=c.border.left)

def fill_gaps_grey(ws, start_row: int, end_row: int, cfg: GridConfig):
    """
    Post-processing step:
    Scans the grid area. Any cell that has NO background color (None)
    gets filled with Grey (A6A6A6).
    """
    grey_fill = PatternFill(fill_type="solid", start_color="A6A6A6", end_color="A6A6A6")
    
    for r in range(start_row, end_row + 1):
        for b in range(cfg.total_blocks):
            col = cfg.start_column + b
            cell = ws.cell(row=r, column=col)
            
            has_color = False
            if cell.fill and cell.fill.fill_type == "solid":
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
    cell.fill = PatternFill(fill_type="solid", start_color=COLORS["PURPLE"], end_color=COLORS["PURPLE"])
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
    
    hours_span = 6 
    blocks_span = hours_span * cfg.blocks_per_hour
    grid_end_col = cfg.start_column + cfg.total_blocks - 1
    
    c_start = max(cfg.start_column, grid_end_col - blocks_span + 1)
    c_end = grid_end_col
    r_start = start_row
    r_end = start_row + 1 

    clean_merge_area(ws, r_start, c_start, r_end, c_end)

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
            ("TOP OF ESC\n4–6", COLORS["GREY"],     "AREA AT THE TOP OF ESC\n6", COLORS["GREY"]),
            ("VIP/MIP",         COLORS["YELLOW"],   "PERSON ALLOCATED TO ASSIST\nWITH VIP/MIPs", COLORS["YELLOW"]),
            ("POM SERVICING",   COLORS["ORANGE"],   "SERVICE ALL POMS",           COLORS["ORANGE"]),
            ("PLATFORM",        COLORS["GREEN"],    "PLATFORM SATS",              COLORS["GREEN"]),
            ("PTI",             COLORS["PEACH"],    "PTI DUTY",                   COLORS["PEACH"]),
        ]

    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(left=BLACK_THIN, right=BLACK_THIN, top=BLACK_THIN, bottom=BLACK_THIN)
    font   = Font(bold=False, size=8)

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

def build_single_sheet(
    out_filename: str,
    staff_victoria: list[dict],
    staff_district: list[dict],
    staff_cardinal: list[dict],
    cfg: GridConfig,
    rows_per_section=(28, 28, 13)
):
    wb = Workbook()
    ws = wb.active
    ws.title = "Roster"
    for i, w in enumerate(cfg.meta_col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    total_counts = calculate_hourly_counts(
        [staff_victoria, staff_district, staff_cardinal], 
        cfg
    )
    draw_time_header(ws, 1, cfg)
    draw_global_banners(ws, cfg)
    draw_meta_header(ws, row=6, cfg=cfg, hourly_counts=total_counts)

    def render_section(section_name: str, staff_rows: list[dict], start_row: int, fixed_count: int, is_first=False, custom_header_row: int = None) -> int:
        
        end_row = start_row + fixed_count - 1
        draw_clean_grid(ws, start_row, end_row, cfg, hour_line="thin")

        # 2. Fill rows (Staff data if available, otherwise just Duty ID)
        for i in range(fixed_count):
            r = start_row + i
            
            s = staff_rows[i] if i < len(staff_rows) else None
            font_meta = Font(size=10, bold=True)
            center_align = Alignment(horizontal="center", vertical="center")
            
            prefix_map = {"Victoria": "BN", "District": "BN", "Cardinal": "BN"}
            prefix = prefix_map.get(section_name, "XX")

            if s:
                ws[f"A{r}"] = s.get("id", "")
                ws[f"B{r}"] = s["start_time"]
                ws[f"C{r}"] = s["finish_time"]
                ws[f"D{r}"] = s["name"]
                ws[f"E{r}"] = s.get("radio", "")
                
                paint_shift(ws, r, section_name, s, cfg)
                merge_painted_segments_per_hour(ws, r, cfg)
            else:
                ws[f"A{r}"] = ""
            for col in "ABCDE":
                cell = ws[f"{col}{r}"]
                cell.font = font_meta
                cell.alignment = center_align
        
        fill_gaps_grey(ws, start_row, end_row, cfg)

        draw_internal_section_label(ws, start_row, section_name, cfg)
        border_start = custom_header_row if custom_header_row is not None else start_row


        # Extra notes ONLY for Victoria (placed inside the grid area)
        if section_name == "Victoria":
            add_victoria_cssi_note(ws, cfg, start_row, end_row)
            add_victoria_right_legend_in_grid(ws, cfg, start_row)
        top_style = "thin" if is_first else "thick"
        draw_section_outline(ws, border_start, end_row, cfg, top_style="thick")     

        return end_row + 1
    
    row_cursor = 7
    row_cursor = render_section("Victoria", staff_victoria, row_cursor, rows_per_section[0], custom_header_row=2)
    row_cursor += 1 

    draw_time_header(ws, row_cursor, cfg)  
    row_cursor += 1                        

    draw_meta_header(ws, row_cursor, cfg)
    meta_start_dist = row_cursor 
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

    os.makedirs("output", exist_ok=True)
    path = f"output/{out_filename}"
    wb.save(path)
    
    return path
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import math, os

# ================= Config =================
class GridConfig:
    def __init__(
        self,
        filename="victoria_station_schedule.xlsx",
        start_hour=5,            # 05:00
        end_hour=25,             # 01:00 next day
        total_rows=40,           # expected staffed rows
        bottom_padding_rows=120, # extra empty grid rows under bottom
        start_column=6,          # F
        blocks_per_hour=4        # 4 cols/hr => 15-min blocks
    ):
        self.filename = filename
        self.start_hour = start_hour
        self.end_hour = end_hour
        self.total_rows = total_rows
        self.bottom_padding_rows = bottom_padding_rows
        self.start_column = start_column
        self.blocks_per_hour = blocks_per_hour

    @property
    def minutes_per_block(self): return 60 // self.blocks_per_hour
    @property
    def total_blocks(self): return (self.end_hour - self.start_hour) * self.blocks_per_hour

# ============== Grid (clean style) ==============
def create_time_scale_grid(ws, config: GridConfig):
    black  = Side(border_style="thin", color="000000")   # hour verticals
    grey   = Side(border_style="thin", color="000000")   # row dividers
    center = Alignment(horizontal="center", vertical="center")

    # Banners (A–E)
    ws.merge_cells("A2:E3"); ws.merge_cells("A4:E5")
    ws["A2"] = "Confirmation of Security Check started"
    ws["A4"] = "Confirmation of Security Check Completed"
    ws["A2"].alignment = center; ws["A4"].alignment = center

    total_hours = config.end_hour - config.start_hour

    # Hour headers (row 1) merged across 4 columns
    for h in range(total_hours):
        c0 = config.start_column + h * config.blocks_per_hour
        c1 = c0 + config.blocks_per_hour - 1
        ws.merge_cells(f"{get_column_letter(c0)}1:{get_column_letter(c1)}1")
        head = ws.cell(row=1, column=c0)
        head.value = f"{(config.start_hour + h) % 24:02d}:00"
        head.alignment = center
        head.border = Border(left=black, right=black, top=black, bottom=black)
        for col in range(c0, c1 + 1):
            ws.column_dimensions[get_column_letter(col)].width = 2.2

    # Depth to draw the grid
    last_grid_row = 1 + 1 + config.total_rows + config.bottom_padding_rows

    # Body rows (unmerged cells, clean borders)
    for r in range(2, last_grid_row + 1):
        ws.row_dimensions[r].height = 18
        for b in range(config.total_blocks):
            col = config.start_column + b
            is_hour_left  = ((col - config.start_column) % config.blocks_per_hour) == 0
            is_hour_right = ((col - config.start_column) % config.blocks_per_hour) == (config.blocks_per_hour - 1)

            left  = black if is_hour_left  else None
            right = black if is_hour_right else None
            top   = grey
            bottom= grey

            cell = ws.cell(row=r, column=col)
            cell.border = Border(left=left, right=right, top=top, bottom=bottom)
            cell.alignment = center

    # Wider meta columns A–E
    for c in range(1, config.start_column):
        ws.column_dimensions[get_column_letter(c)].width = 10

    ws.freeze_panes = ws.cell(row=6, column=config.start_column)

# ============== Time helpers (step-in starts) ==============
def _to_minutes_since_start(hhmm: str, config: GridConfig) -> int:
    h, m = map(int, hhmm.split(":"))
    s = config.start_hour % 24
    if h < s:  # 00:xx => next day
        h += 24
    return max(0, (h - config.start_hour) * 60 + m)

def _block_index_start(hhmm: str, config: GridConfig) -> int:
    return int(math.ceil(_to_minutes_since_start(hhmm, config) / config.minutes_per_block))  # step-in

def _block_index_end(hhmm: str, config: GridConfig, round_finish="ceil") -> int:
    mins = _to_minutes_since_start(hhmm, config)
    if round_finish == "floor":
        return int(mins // config.minutes_per_block)
    return int(math.ceil(mins / config.minutes_per_block))


# ============== Duration & colour chooser ==============
DEEP_BLUE  = "1F4E79"  # CSA1 (8h)
LIGHT_BLUE = "9CD7FF"  # CSA2 (7.5h)

def _duration_minutes(start_hhmm: str, finish_hhmm: str, config: GridConfig) -> int:
    ms = _to_minutes_since_start(start_hhmm, config)
    me = _to_minutes_since_start(finish_hhmm, config)
    if me <= ms:
        me += 24 * 60  # cross midnight
    return me - ms

def _fill_for_duration(start_hhmm: str, finish_hhmm: str, config: GridConfig) -> str:
    mins = _duration_minutes(start_hhmm, finish_hhmm, config)
    # Tolerance ±7 min to allow small offsets/rounding
    if abs(mins - 480) <= 7:   # 8h
        return DEEP_BLUE
    if abs(mins - 450) <= 7:   # 7.5h
        return LIGHT_BLUE
    # Fallback: choose nearest
    return DEEP_BLUE if mins >= 465 else LIGHT_BLUE


# ============== Paint per 15-min cell ==============
def draw_shift_bar(ws, row, start_hhmm, finish_hhmm, config: GridConfig,
                   round_finish="ceil", fill_color=None):
    if fill_color is None:
        fill_color = _fill_for_duration(start_hhmm, finish_hhmm, config)
    fill = PatternFill(fill_type="solid", start_color=fill_color, end_color=fill_color)

    s_blk = _block_index_start(start_hhmm, config)
    e_blk = _block_index_end(finish_hhmm, config, round_finish=round_finish)
    if e_blk <= s_blk:
        e_blk += config.total_blocks  # cross-midnight guard

    s_col = config.start_column + max(0, s_blk)
    e_col_excl = config.start_column + min(e_blk, config.total_blocks)

    for c in range(s_col, e_col_excl):
        ws.cell(row=row, column=c).fill = fill


# ============== Merge painted segments within each hour (≥2 cells) ==============
def merge_painted_segments_per_hour(ws, row: int, config: GridConfig):
    black  = Side(border_style="thin", color="000000")   # hour boundaries
    grey   = Side(border_style="thin", color="000000")   # horizontals
    center = Alignment(horizontal="center", vertical="center")
    total_hours = config.end_hour - config.start_hour

    for h in range(total_hours):
        c0 = config.start_column + h * config.blocks_per_hour
        c1 = c0 + config.blocks_per_hour - 1

        painted = []
        for col in range(c0, c1 + 1):
            f = ws.cell(row=row, column=col).fill
            painted.append(f is not None and f.fill_type == "solid")

        if not any(painted):
            continue

        # scan for contiguous painted runs
        run_start = None
        for idx in range(config.blocks_per_hour + 1):  # + sentinel
            is_painted = painted[idx] if idx < config.blocks_per_hour else False
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
                    # keep borders crisp at hour edges
                    left  = black if m0 == c0 else None
                    right = black if m1 == c1 else None
                    top   = grey
                    bottom= grey
                    tl.border = Border(left=left, right=right, top=top, bottom=bottom)
                run_start = None


# ============== Build workbook ==============
def generate_timesheet_document(config: GridConfig, staff_shifts, round_finish="ceil", merge_segments=True):
    wb = Workbook()
    ws = wb.active
    ws.title = "Victoria"

    create_time_scale_grid(ws, config)

    start_row = 7
    for i, s in enumerate(staff_shifts, start=1):
        r = start_row + i - 1
        ws[f"A{r}"] = f"BH{i:02d}"
        ws[f"B{r}"] = s["start_time"]
        ws[f"C{r}"] = s["finish_time"]
        ws[f"D{r}"] = s["name"]
        ws[f"E{r}"] = s.get("radio", "")

        # colour chosen by duration (8h deep blue, 7.5h light blue)
        draw_shift_bar(ws, r, s["start_time"], s["finish_time"], config, round_finish=round_finish)

        if merge_segments:
            merge_painted_segments_per_hour(ws, r, config)

    os.makedirs("output", exist_ok=True)
    path = f"output/{config.filename}"
    wb.save(path)
    return path


# ============== Example ==============
if __name__ == "__main__":
    cfg = GridConfig(
        filename="victoria_station_schedule.xlsx",
        start_hour=5,
        end_hour=25,
        total_rows=40,
        bottom_padding_rows=120,
        start_column=6,
        blocks_per_hour=4
    )

    staff = [
        # 8h (deep blue)
        {"name": "John Doe",   "radio": "W12", "start_time": "05:10", "finish_time": "13:10"},  # 8h
        {"name": "John Doe",   "radio": "W12", "start_time": "05:10", "finish_time": "12:40"},  # 8h
        {"name": "Ali Musa",   "radio": "W11", "start_time": "08:20", "finish_time": "16:20"},  # 8h
        # 7.5h (light blue)
        {"name": "Jane Smith", "radio": "W13", "start_time": "07:00", "finish_time": "14:30"},  # 7.5h
        {"name": "Tosin Dada", "radio": "W10", "start_time": "12:00", "finish_time": "19:30"},  # 7.5h
        {"name": "Fatima Oke", "radio": "W14", "start_time": "14:30", "finish_time": "22:00"},  # 7.5h
    ]

    print("Saved:", generate_timesheet_document(cfg, staff, round_finish="ceil", merge_segments=True))
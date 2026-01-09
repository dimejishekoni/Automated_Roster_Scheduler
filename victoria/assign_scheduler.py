import math
import random  # <--- THIS WAS MISSING AND CAUSED THE FAILURE

# ================= MOVED FROM ROSTER.PY =================

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

class GridConfig:
    def __init__(self, start_hour=5, end_hour=25, start_column=6, blocks_per_hour=4):
        self.start_hour = start_hour
        self.end_hour = end_hour
        self.start_column = start_column
        self.blocks_per_hour = blocks_per_hour
        self.meta_col_widths = (5.6, 8, 8, 14.5, 6.2)   # widths A..E

    @property
    def minutes_per_block(self): return 60 // self.blocks_per_hour

    @property
    def total_hours(self): return self.end_hour - self.start_hour

    @property
    def total_blocks(self): return self.total_hours * self.blocks_per_hour

# ================= TIME HELPERS =================

def hhmm_to_minutes(hhmm: str) -> int:
    h, m = map(int, hhmm.split(":"))
    return h * 60 + m

def minutes_to_hhmm(m: int) -> str:
    m %= 24 * 60
    h = m // 60
    mm = m % 60
    return f"{h:02d}:{mm:02d}"

def _to_minutes_since_start(hhmm: str, cfg: GridConfig) -> int:
    h, m = map(int, hhmm.split(":"))
    s = cfg.start_hour % 24
    if h < s:  # 00:xx => next day
        h += 24
    return max(0, (h - cfg.start_hour) * 60 + m)

def _block_index_start(hhmm: str, cfg: GridConfig) -> int:
    return int(math.ceil(_to_minutes_since_start(hhmm, cfg) / cfg.minutes_per_block))

def _block_index_end(hhmm: str, cfg: GridConfig, round_finish="ceil") -> int:
    mins = _to_minutes_since_start(hhmm, cfg)
    if round_finish == "floor":
        return int(mins // cfg.minutes_per_block)
    return int(math.ceil(mins / cfg.minutes_per_block))

def _block_index_start_inclusive(hhmm: str, cfg: GridConfig) -> int:
    mins = _to_minutes_since_start(hhmm, cfg)
    return int(mins // cfg.minutes_per_block)

def _duration_minutes(start_hhmm: str, finish_hhmm: str, cfg: GridConfig) -> int:
    ms = _to_minutes_since_start(start_hhmm, cfg)
    me = _to_minutes_since_start(finish_hhmm, cfg)
    if me <= ms:
        me += 24 * 60
    return me - ms

# ================= SCHEDULER LOGIC =================

STATUS_OFF = 0
STATUS_GATELINE = 1
STATUS_BREAK = 2
STATUS_CSSI = 3
STATUS_SPECIALIST = 4
STATUS_GENERAL = 5
STATUS_SECURITY = 6

class RosterEngine:
    def __init__(self, cfg):
        self.cfg = cfg
        self.matrix = {}
        self.text_matrix = {}
        self.staff_data = {}
        self.assignments = {}

    def load_staff(self, staff_list):
        for s in staff_list:
            s_id = s["id"]
            self.staff_data[s_id] = s
            self.assignments[s_id] = []
            
            blocks = [STATUS_OFF] * self.cfg.total_blocks
            txt_blocks = [""] * self.cfg.total_blocks 
            
            # --- STRICT ROUNDING LOGIC ---
            # Start: Round UP (math.ceil). 
            # 05:10 -> 05:15 block (05:00-05:15 is empty)
            start_mins = _to_minutes_since_start(s["start_time"], self.cfg)
            s_blk = int(math.ceil(start_mins / self.cfg.minutes_per_block))
            
            # End: Round DOWN (//). 
            # 00:50 -> 00:45 block (00:45-01:00 is empty)
            end_mins = _to_minutes_since_start(s["finish_time"], self.cfg)
            e_blk = int(end_mins // self.cfg.minutes_per_block)
            # -----------------------------
            
            # Handle overnight wraps
            if e_blk <= s_blk and e_blk < (24*4): # Wrap around
                for b in range(s_blk, self.cfg.total_blocks): blocks[b] = STATUS_GATELINE
                for b in range(0, e_blk): blocks[b] = STATUS_GATELINE
            else:
                for b in range(s_blk, e_blk):
                    if 0 <= b < self.cfg.total_blocks:
                        blocks[b] = STATUS_GATELINE
            
            self.matrix[s_id] = blocks
            self.text_matrix[s_id] = txt_blocks
        print(f"DEBUG: Loaded {len(staff_list)} staff into Matrix.")

    def get_gateline_count(self, block_idx):
        count = 0
        for s_id, timeline in self.matrix.items():
            if timeline[block_idx] == STATUS_GATELINE:
                count += 1
        return count

    def can_assign(self, staff_id, start_blk, duration_blocks, min_gateline=1):
        # NOTE: Default min_gateline lowered to 1 to ensure assignments happen 
        # even with low staff numbers.
        timeline = self.matrix.get(staff_id)
        if not timeline: return False

        for b in range(start_blk, start_blk + duration_blocks):
            b_idx = b % self.cfg.total_blocks
            if timeline[b_idx] != STATUS_GATELINE:
                return False
            
            # current_coverage = self.get_gateline_count(b_idx)
            # if (current_coverage - 1) < min_gateline:
            #     return False 
        return True

    def commit_assignment(self, staff_id, start_blk, duration_blocks, status_code, duty_label):
        timeline = self.matrix[staff_id]
        self.assignments[staff_id].append({
            "label": duty_label,
            "start_blk": start_blk,
            "duration": duration_blocks
        })
        for b in range(start_blk, start_blk + duration_blocks):
            b_idx = b % self.cfg.total_blocks
            timeline[b_idx] = status_code

    def allocate_breaks(self):
        print("DEBUG: Allocating Breaks (30m Break + 30m CSSI)...")
        sorted_staff = sorted(self.staff_data.values(), key=lambda x: x["start_time"])
        assigned_ids = set()
        count_paired = 0

        # --- CONFIGURATION: 1 Block = 15 Minutes ---
        SEGMENT_LEN = 2  # 30 Minutes
        TOTAL_LEN = 4    # 1 Hour (30m + 30m)

        for s1 in sorted_staff:
            id1 = s1["id"]
            if id1 in assigned_ids: continue 
            
            start_blk = _block_index_start_inclusive(s1["start_time"], self.cfg)
            
            # Target: 4 hours into shift
            primary_window = range(start_blk + 16, start_blk + 21) 
            secondary_window = range(start_blk + 12, start_blk + 16)
            search_slots = list(primary_window) + list(secondary_window)
            
            partner_found = False
            for s2 in sorted_staff:
                id2 = s2["id"]
                if id1 == id2 or id2 in assigned_ids: continue
                
                # Look for a 1-hour slot (TOTAL_LEN) where both are free
                for t in search_slots:
                    if self.can_assign(id1, t, TOTAL_LEN, min_gateline=1):
                        if self.can_assign(id2, t, TOTAL_LEN, min_gateline=1):
                            
                            # --- ASSIGNMENT PATTERN ---
                            # Staff 1: 30m Break -> 30m CSSI
                            self.commit_assignment(id1, t, SEGMENT_LEN, STATUS_BREAK, "Meal Break")
                            self.commit_assignment(id1, t+SEGMENT_LEN, SEGMENT_LEN, STATUS_CSSI, "CSSI")
                            
                            # Staff 2: 30m CSSI -> 30m Break
                            self.commit_assignment(id2, t, SEGMENT_LEN, STATUS_CSSI, "CSSI")
                            self.commit_assignment(id2, t+SEGMENT_LEN, SEGMENT_LEN, STATUS_BREAK, "Meal Break")
                            
                            assigned_ids.add(id1)
                            assigned_ids.add(id2)
                            partner_found = True
                            count_paired += 1
                            break 
                if partner_found: break
            
            # Solo Assignment (No partner found)
            if not partner_found:
                for t in search_slots:
                    if self.can_assign(id1, t, TOTAL_LEN, min_gateline=1):
                        # Default Pattern: Break -> CSSI
                        self.commit_assignment(id1, t, SEGMENT_LEN, STATUS_BREAK, "Meal Break")
                        self.commit_assignment(id1, t+SEGMENT_LEN, SEGMENT_LEN, STATUS_CSSI, "CSSI")
                        assigned_ids.add(id1)
                        break

        print(f"DEBUG: Breaks Assigned. {count_paired} Pairs formed.")

    def allocate_specialist_duties(self, duties_list):
        duties_list.sort(key=lambda x: x["priority"])
        count_assigned = 0
        
        for duty in duties_list:
            if duty["required_grade"] != "CSA1": continue
            
            start_b = _block_index_start_inclusive(minutes_to_hhmm(duty["start"]), self.cfg)
            duration = (duty["end"] - duty["start"]) // 15
            
            candidates = []
            for s_id, s_data in self.staff_data.items():
                if (s_data.get("grade") or "").upper() == "CSA1":
                    candidates.append((s_id, len(self.assignments[s_id])))
            
            # SHUFFLE ADDED HERE
            random.shuffle(candidates)
            candidates.sort(key=lambda x: x[1])

            for s_id, workload in candidates:
                if self.can_assign(s_id, start_b, duration, min_gateline=1):
                    self.commit_assignment(s_id, start_b, duration, STATUS_SPECIALIST, duty["label"])
                    count_assigned += 1
                    break 
        print(f"DEBUG: Specialist Duties Assigned: {count_assigned}")

    # def allocate_general_duties(self, duties_list):
    #     count_assigned = 0
    #     for duty in duties_list:
    #         if duty["required_grade"] == "CSA1": continue 
            
    #         start_b = _block_index_start_inclusive(minutes_to_hhmm(duty["start"]), self.cfg)
    #         duration = (duty["end"] - duty["start"]) // 15
            
    #         candidates = []
    #         for s_id, s_data in self.staff_data.items():
    #             candidates.append((s_id, len(self.assignments[s_id])))
            
    #         # SHUFFLE ADDED HERE
    #         random.shuffle(candidates)
    #         candidates.sort(key=lambda x: x[1])
            
    #         for s_id, workload in candidates:
    #             if self.can_assign(s_id, start_b, duration, min_gateline=1):
    #                 self.commit_assignment(s_id, start_b, duration, STATUS_GENERAL, duty["label"])
    #                 count_assigned += 1
    #                 break
    #     print(f"DEBUG: General Duties Assigned: {count_assigned}")

    def allocate_general_duties(self, duties_list):
        print("DEBUG: Allocating General Duties (Security Check)...")
        count_assigned = 0
        
        # Sort duties by priority
        duties_list.sort(key=lambda x: x["priority"])
        
        for duty in duties_list:
            if duty["role"] != "SECURITY": continue 
            
            # Calculate strict block index for duty
            start_b = int(math.ceil(_to_minutes_since_start(minutes_to_hhmm(duty["start"]), self.cfg) / self.cfg.minutes_per_block))
            duration = (duty["end"] - duty["start"]) // 15
            
            candidates = []
            for s_id, s_data in self.staff_data.items():
                candidates.append((s_id, len(self.assignments[s_id])))
            
            random.shuffle(candidates)         
            candidates.sort(key=lambda x: x[1]) 
            
            for s_id, workload in candidates:
                if self.can_assign(s_id, start_b, duration, min_gateline=1):
                    self.commit_assignment(s_id, start_b, duration, STATUS_SECURITY, duty["label"])
                    count_assigned += 1
                    break

    def get_updated_staff_list(self):
        output = []
        for s_id, s_data in self.staff_data.items():
            s_new = s_data.copy()
            s_new["assigned_duties"] = self.assignments[s_id]
            s_new["status_matrix"] = self.matrix[s_id] 
            output.append(s_new)
        return output

# ================= SLOT GENERATOR =================

def generate_victoria_duty_slots_full(open_time="05:00", close_time="00:00", include_night=False):
    slots = []
    def add_slot(label, role, start_min, end_min, grade_req="ANY"):
        slots.append({
            "label": label,
            "role": role,
            "start": start_min,
            "end": end_min,
            "required_grade": grade_req,
            "priority": DUTY_PRIORITY[role],
        })

    open_min = hhmm_to_minutes(open_time)
    close_min = hhmm_to_minutes(close_time)
    if close_min <= open_min: close_min += 24 * 60

    if include_night: sec_start, sec_end = 0, 24 * 60
    else: sec_start, sec_end = open_min, close_min

    # 1. Security
    t = sec_start
    while t + 60 <= sec_end:
        add_slot("Security Check", "SECURITY", t, t + 60, grade_req="ANY")
        t += 60

    # 2. Escalator
    esc_start, esc_end = hhmm_to_minutes("07:00"), hhmm_to_minutes("23:00")
    t = esc_start
    while t + 60 <= esc_end:
        add_slot("TOP OF ESC 4-6", "ESCALATOR", t, t + 60, grade_req="ANY")
        t += 60

    # 3. PTI
    pti_start, pti_end = hhmm_to_minutes("07:30"), hhmm_to_minutes("23:00")
    t = pti_start
    while t + 60 <= pti_end:
        add_slot("PTI WB", "PTI WB", t, t + 60, grade_req="CSA1")
        add_slot("PTI EB", "PTI EB", t, t + 60, grade_req="CSA1")
        t += 60

    # 4. SATS
    for s_str, e_str in [("07:00", "09:00"), ("17:00", "19:00")]:
        start, end = hhmm_to_minutes(s_str), hhmm_to_minutes(e_str)
        t = start
        while t + 60 <= end:
            add_slot("SATS WB", "SATS WB", t, t + 60, grade_req="CSA1")
            add_slot("SATS EB", "SATS EB", t, t + 60, grade_req="CSA1")
            t += 60

    # 5. VIP
    vip_start, vip_end = hhmm_to_minutes("07:00"), hhmm_to_minutes("22:00")
    t = vip_start
    while t + 60 <= vip_end:
        add_slot("VIP/MIP", "VIP/MIP", t, t + 60, grade_req="ANY")
        t += 60

    # 6. P3
    for s_str, e_str in [("07:30", "09:00"), ("17:00", "19:00")]:
        start, end = hhmm_to_minutes(s_str), hhmm_to_minutes(e_str)
        t = start
        while t + 60 <= end:
            add_slot("P3", "P3", t, t + 60, grade_req="CSA1")
            t += 60

    slots.sort(key=lambda s: (s["start"], s["priority"]))
    return slots

def assign_victoria_duties_v2(victoria_staff, cfg):
    print("--- Starting Scheduler V2 ---")
    engine = RosterEngine(cfg)
    engine.load_staff(victoria_staff)
    all_slots = generate_victoria_duty_slots_full(open_time="05:00", close_time="00:00")
    engine.allocate_breaks()
    # engine.allocate_specialist_duties(all_slots)
    engine.allocate_general_duties(all_slots)
    print("--- Scheduler Finished ---")
    return engine.get_updated_staff_list()
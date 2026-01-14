import math
import random  
from config import DUTY_PRIORITY, DUTY_SECTION_MAP, GridConfig, COLORS
from helper import (
    hhmm_to_minutes, minutes_to_hhmm, to_minutes_since_start, 
    block_index_start, block_index_end, duration_minutes
)

# ================= SCHEDULER LOGIC =================
STATUS_OFF = 0
STATUS_GATELINE = 1
STATUS_BREAK = 2
STATUS_CSSI = 3
STATUS_SPECIALIST = 4
STATUS_GENERAL = 5
STATUS_SECURITY = 6
STATUS_PLATFORM = 7

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
            
            start_mins = to_minutes_since_start(s["start_time"], self.cfg)
            s_blk = int(math.ceil(start_mins / self.cfg.minutes_per_block))
            
            end_mins = to_minutes_since_start(s["finish_time"], self.cfg)
            e_blk = int(end_mins // self.cfg.minutes_per_block)
            
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

    def can_assign(self, staff_id, start_blk, duration_blocks, min_gateline=1):
        timeline = self.matrix.get(staff_id)
        if not timeline: return False

        for b in range(start_blk, start_blk + duration_blocks):
            b_idx = b % self.cfg.total_blocks
            if timeline[b_idx] != STATUS_GATELINE:
                return False
        return True

    def commit_assignment(self, staff_id, start_blk, duration_blocks, status_code, duty_label):
        timeline = self.matrix[staff_id]
        txt_timeline = self.text_matrix[staff_id]
        self.assignments[staff_id].append({
            "label": duty_label,
            "start_blk": start_blk,
            "duration": duration_blocks
        })
        for b in range(start_blk, start_blk + duration_blocks):
            b_idx = b % self.cfg.total_blocks
            timeline[b_idx] = status_code
            txt_timeline[b_idx] = duty_label

    def allocate_breaks(self):
        print("DEBUG: Allocating Breaks (Target: 3h from Next Full Hour)...")
        sorted_staff = sorted(self.staff_data.values(), key=lambda x: x["start_time"])
        assigned_ids = set()
        count_paired = 0
        SEGMENT_LEN = 2 
        TOTAL_LEN = 4    

        for s1 in sorted_staff:
            id1 = s1["id"]
            if id1 in assigned_ids: continue 
            
            s_min = to_minutes_since_start(s1["start_time"], self.cfg)
            e_min = to_minutes_since_start(s1["finish_time"], self.cfg)
            if e_min <= s_min: e_min += 24 * 60
            
            if (e_min - s_min) < 300:
                # Short shift logic could go here
                continue

            remainder = s_min % 60
            if remainder == 0: effective_start = s_min
            else: effective_start = s_min + (60 - remainder)

            target_offset = 180 
            target_break_min = effective_start + target_offset
            target_blk = int(math.ceil(target_break_min / self.cfg.minutes_per_block))
            
            window_start = target_blk - 4
            window_end = target_blk + 5
            
            candidates = []
            for t in range(window_start, window_end):
                if t % 2 != 0: continue 
                if t % 4 == 0: priority = 0
                else: priority = 1
                distance = abs(t - target_blk)
                candidates.append((t, priority, distance))
            
            candidates.sort(key=lambda x: (x[1], x[2]))
            search_slots = [c[0] for c in candidates]
            
            partner_found = False
            for s2 in sorted_staff:
                id2 = s2["id"]
                if id1 == id2 or id2 in assigned_ids: continue
                
                s2_s = to_minutes_since_start(s2["start_time"], self.cfg)
                s2_e = to_minutes_since_start(s2["finish_time"], self.cfg)
                if s2_e <= s2_s: s2_e += 24*60
                if (s2_e - s2_s) < 300: continue

                for t in search_slots:
                    if self.can_assign(id1, t, TOTAL_LEN, min_gateline=1):
                        if self.can_assign(id2, t, TOTAL_LEN, min_gateline=1):
                            self.commit_assignment(id1, t, SEGMENT_LEN, STATUS_BREAK, "Meal Break")
                            self.commit_assignment(id1, t+SEGMENT_LEN, SEGMENT_LEN, STATUS_CSSI, "CSSI")
                            self.commit_assignment(id2, t, SEGMENT_LEN, STATUS_CSSI, "CSSI")
                            self.commit_assignment(id2, t+SEGMENT_LEN, SEGMENT_LEN, STATUS_BREAK, "Meal Break")
                            assigned_ids.add(id1)
                            assigned_ids.add(id2)
                            partner_found = True
                            count_paired += 1
                            break 
                if partner_found: break
            
            if not partner_found:
                for t in search_slots:
                    if self.can_assign(id1, t, TOTAL_LEN, min_gateline=1):
                        self.commit_assignment(id1, t, SEGMENT_LEN, STATUS_BREAK, "Meal Break")
                        self.commit_assignment(id1, t+SEGMENT_LEN, SEGMENT_LEN, STATUS_CSSI, "CSSI")
                        assigned_ids.add(id1)
                        break
        print(f"DEBUG: Breaks Assigned. {count_paired} Pairs formed.")

    def allocate_fixed_position(self, start_time: str, end_time: str, label: str, duration_blocks=4):
        print(f"DEBUG: Allocating Fixed Position: {label} ({start_time}-{end_time})...")
        s_blk = block_index_start(start_time, self.cfg)
        e_blk = block_index_end(end_time, self.cfg)
        
        current_blk = s_blk
        count_assigned = 0
        already_assigned_ids = set()

        while current_blk < e_blk:
            segment_end = min(current_blk + duration_blocks, e_blk)
            actual_len = segment_end - current_blk
            assigned_ok = False
            candidates = list(self.staff_data.keys())
            random.shuffle(candidates)
            
            for s_id in candidates:
                s_data = self.staff_data[s_id]
                if s_data.get("is_restricted", False): continue
                if s_id in already_assigned_ids: continue

                if self.can_assign(s_id, current_blk, actual_len, min_gateline=2):
                    self.commit_assignment(s_id, current_blk, actual_len, STATUS_GENERAL, label)
                    assigned_ok = True
                    count_assigned += 1
                    already_assigned_ids.add(s_id)
                    break
            
            if not assigned_ok:
                print(f"  > WARNING: Could not cover {label} at block {current_blk}. Gateline too tight.")
            current_blk += actual_len

        print(f"DEBUG: Finished {label}. {count_assigned} staff assigned.")

    def allocate_general_duties(self, duties_list, target_role="SECURITY", assignment_status=STATUS_SECURITY, banned_sections=None):
        print(f"DEBUG: Allocating {target_role} Duties...")
        count_assigned = 0
        duties_list.sort(key=lambda x: x["start"])
        
        for duty in duties_list:
            if duty["role"] != target_role: continue 
            
            start_hhmm = minutes_to_hhmm(duty["start"]) 
            start_b = block_index_start(start_hhmm, self.cfg)
            duration = (duty["end"] - duty["start"]) // 15
            
            priority_new = []
            for s_id, s_data in self.staff_data.items():
                if s_data.get("is_restricted", False): continue
                
                if banned_sections:
                    staff_section = DUTY_SECTION_MAP.get(s_id, "Victoria") 
                    if staff_section in banned_sections: continue 

                # Grade Check (Specific for SATS)
                req_grade = duty.get("required_grade", "ANY")
                if req_grade != "ANY":
                    staff_grade = (s_data.get("grade") or "").upper().strip()
                    if staff_grade != req_grade: continue 

                # No Back-to-Back
                prev_b = start_b - 1
                if prev_b >= 0:
                    prev_status = self.matrix[s_id][prev_b]
                    if prev_status == assignment_status: continue

                if not self.can_assign(s_id, start_b, duration, min_gateline=2): continue
                
                workload = len(self.assignments[s_id])
                priority_new.append((s_id, workload))

            if priority_new:
                priority_new.sort(key=lambda x: x[1])
                candidates = priority_new[:3]
                chosen_id = random.choice(candidates)[0]
                self.commit_assignment(chosen_id, start_b, duration, assignment_status, duty["label"])
                count_assigned += 1
                
        print(f"DEBUG: {target_role} Duties Assigned: {count_assigned}")
        print(f"DEBUG: {target_role} Duties Assigned: {count_assigned}")

    def get_updated_staff_list(self):
        output = []
        for s_id, s_data in self.staff_data.items():
            s_new = s_data.copy()
            s_new["assigned_duties"] = self.assignments[s_id]
            s_new["status_matrix"] = self.matrix[s_id] 
            s_new["text_matrix"] = self.text_matrix[s_id]
            output.append(s_new)
        return output

# ================= SLOT GENERATOR =================

def generate_victoria_duty_slots_full(open_time="05:00", close_time="00:00", date_str: str = None):
    slots = []
    def add_slot(label, role, start_min, end_min, grade_req="ANY"):
        slots.append({
            "label": label,
            "role": role,
            "start": start_min,
            "end": end_min,
            "required_grade": grade_req,
            "priority": DUTY_PRIORITY.get(role, 99),
        })

    open_min = hhmm_to_minutes(open_time)
    close_min = hhmm_to_minutes(close_time)
    if close_min <= open_min: close_min += 24 * 60

    sec_start, sec_end = open_min, close_min

    # 1. Security (Everyday)
    t = sec_start
    while t + 60 <= sec_end:
        add_slot("Security Check", "SECURITY", t, t + 60, grade_req="ANY")
        t += 60

    # 2. VIP (Everyday)
    vip_start, vip_end = hhmm_to_minutes("07:00"), hhmm_to_minutes("22:00")
    t = vip_start
    while t + 60 <= vip_end:
        add_slot("VIP/MIP", "VIP/MIP", t, t + 60, grade_req="ANY")
        t += 60

    # 3. SATS (Weekdays Only)
    is_weekday = False
    if date_str:
        d = date_str.lower()
        if "saturday" not in d and "sunday" not in d:
            is_weekday = True
    
    if is_weekday:
        # --- MORNING PEAK (Split logic) ---
        # Instead of one 07:30-09:00 block, we split it:
        # Slot A: 07:30 - 08:00 (30 mins)
        s_am1 = hhmm_to_minutes("07:30")
        e_am1 = hhmm_to_minutes("08:00")
        
        # Slot B: 08:00 - 09:00 (60 mins)
        s_am2 = hhmm_to_minutes("08:00")
        e_am2 = hhmm_to_minutes("09:00")
        
        # --- EVENING PEAK (Already split into 1-hour chunks) ---
        s_pm1 = hhmm_to_minutes("17:00")
        e_pm1 = hhmm_to_minutes("18:00")
        s_pm2 = hhmm_to_minutes("18:00")
        e_pm2 = hhmm_to_minutes("19:00")

        for label in ["P1", "P2", "P3"]:
            # Morning Split
            add_slot(label, "SATS", s_am1, e_am1, grade_req="CSA1") # 30 mins
            add_slot(label, "SATS", s_am2, e_am2, grade_req="CSA1") # 60 mins
            
            # Evening Split
            add_slot(label, "SATS", s_pm1, e_pm1, grade_req="CSA1") # 60 mins
            add_slot(label, "SATS", s_pm2, e_pm2, grade_req="CSA1") # 60 mins

    slots.sort(key=lambda s: (s["start"], s["priority"]))
    return slots


def assign_victoria_duties_v2(victoria_staff, cfg, date_str=None):
    print("--- Starting Scheduler V2 ---")
    
    # 1. Initialize Engine
    engine = RosterEngine(cfg)
    engine.load_staff(victoria_staff)
    
    # 2. Generate All Potential Slots (Passing Date!)
    all_slots = generate_victoria_duty_slots_full(open_time="05:00", close_time="00:00", date_str=date_str)
    
    # 3. Breaks (Priority 1)
    engine.allocate_breaks()
    
    # 4. Fixed Position: Top of Esc (Priority 2)
    engine.allocate_fixed_position("07:00", "23:00", "Top of Esc", duration_blocks=4)
    
    # 5. SATS Duties (Priority 3 - Weekdays Only)
    engine.allocate_general_duties(
        all_slots,
        target_role="SATS",
        assignment_status=STATUS_PLATFORM, 
        banned_sections=None 
    ) 

    # 6. Security Duties (Priority 4)
    engine.allocate_general_duties(
        all_slots, 
        target_role="SECURITY", 
        assignment_status=STATUS_SECURITY, 
        banned_sections=None
    )
    
    # 7. VIP/MIP Duties (Priority 5)
    engine.allocate_general_duties(
        all_slots, 
        target_role="VIP/MIP", 
        assignment_status=STATUS_SPECIALIST, 
        banned_sections=["Cardinal"]
    )
    
    print("--- Scheduler Finished ---")
    return engine.get_updated_staff_list()
import math
import random  
from config import DUTY_PRIORITY
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
            start_mins = to_minutes_since_start(s["start_time"], self.cfg)
            s_blk = int(math.ceil(start_mins / self.cfg.minutes_per_block))
            
            # End: Round DOWN (//). 
            end_mins = to_minutes_since_start(s["finish_time"], self.cfg)
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
        timeline = self.matrix.get(staff_id)
        if not timeline: return False

        for b in range(start_blk, start_blk + duration_blocks):
            b_idx = b % self.cfg.total_blocks
            if timeline[b_idx] != STATUS_GATELINE:
                return False
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
        print("DEBUG: Allocating Breaks (Target: 3h from Next Full Hour, Clumping Enabled)...")
        sorted_staff = sorted(self.staff_data.values(), key=lambda x: x["start_time"])
        assigned_ids = set()
        count_paired = 0
        SEGMENT_LEN = 2 
        TOTAL_LEN = 4    

        for s1 in sorted_staff:
            id1 = s1["id"]
            if id1 in assigned_ids: continue 
            
            # --- 1. DURATION CHECK ---
            s_min = to_minutes_since_start(s1["start_time"], self.cfg)
            e_min = to_minutes_since_start(s1["finish_time"], self.cfg)
            if e_min <= s_min: e_min += 24 * 60
            
            if (e_min - s_min) < 300:
                print(f"  > Skipping {id1}: Short shift ({e_min - s_min}m)")
                continue

            # --- 2. EFFECTIVE START (Round UP to Next Hour) ---
            remainder = s_min % 60
            if remainder == 0:
                effective_start = s_min
            else:
                effective_start = s_min + (60 - remainder)

            # --- 3. TARGET CALCULATION ---
            target_offset = 180 # 3 hours
            target_break_min = effective_start + target_offset
            target_blk = int(math.ceil(target_break_min / self.cfg.minutes_per_block))
            
            # --- 4. SEARCH CANDIDATES ---
            window_start = target_blk - 4
            window_end = target_blk + 5
            
            candidates = []
            
            for t in range(window_start, window_end):
                if t % 2 != 0: continue # No :15 or :45
                
                if t % 4 == 0: priority = 0
                else: priority = 1

                # Distance from exact target
                distance = abs(t - target_blk)
                candidates.append((t, priority, distance))
            
            # --- 5. SORTING (THE FIX) ---
            candidates.sort(key=lambda x: (x[1], x[2]))
            search_slots = [c[0] for c in candidates]
            
            # --- 6. ASSIGNMENT ---
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


    def allocate_specialist_duties(self, duties_list):
        duties_list.sort(key=lambda x: x["priority"])
        count_assigned = 0
        
        for duty in duties_list:
            if duty["required_grade"] != "CSA1": continue
            
            start_b = block_index_start(minutes_to_hhmm(duty["start"]), self.cfg)
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

    def allocate_general_duties(self, duties_list):
        print("DEBUG: Allocating Security Checks (Smart Clustering)...")
        count_assigned = 0
        
        # 1. Sort duties strictly by time so we build the day chronologically
        duties_list.sort(key=lambda x: x["start"])
        
        for duty in duties_list:
            if duty["role"] != "SECURITY": continue 
            
            # Start Block calculation
            start_b = int(math.ceil(to_minutes_since_start(minutes_to_hhmm(duty["start"]), self.cfg) / self.cfg.minutes_per_block))
            duration = (duty["end"] - duty["start"]) // 15
            
            # We will group candidates into 3 Priority Buckets
            priority_extend = []
            priority_new = []
            priority_gap = []    
            
            for s_id, s_data in self.staff_data.items():
                if not self.can_assign(s_id, start_b, duration, min_gateline=0):
                    continue
                    
                matrix = self.matrix[s_id]

                cluster_count = 0
                last_security_end = -999
                in_cluster = False
                
                for b in range(start_b):
                    if matrix[b] == STATUS_SECURITY:
                        last_security_end = b
                        if not in_cluster:
                            cluster_count += 1
                            in_cluster = True
                    else:
                        in_cluster = False
                
                is_extending = (matrix[start_b - 1] == STATUS_SECURITY)
                
                workload = len(self.assignments[s_id])
                
                # --- LOGIC GATES ---
                
                if is_extending:
                    run_len = 0
                    for k in range(start_b - 1, -1, -1):
                        if matrix[k] == STATUS_SECURITY: run_len += 1
                        else: break
                    if run_len < 4:
                        priority_extend.append((s_id, workload))
                    else:
                        pass
                        
                elif cluster_count == 0:
                    priority_new.append((s_id, workload))
                    
                elif cluster_count == 1:
                    gap = start_b - last_security_end
                    if gap >= 8:
                        priority_gap.append((s_id, workload))                
                else:
                    pass

            def select_best(candidate_list):
                if not candidate_list: return None
                random.shuffle(candidate_list)       
                candidate_list.sort(key=lambda x: x[1]) 
                return candidate_list[0][0] 

            # Try buckets in order
            chosen_id = select_best(priority_extend)
            if not chosen_id: chosen_id = select_best(priority_new)
            if not chosen_id: chosen_id = select_best(priority_gap)
            
            if chosen_id:
                self.commit_assignment(chosen_id, start_b, duration, STATUS_SECURITY, duty["label"])
                count_assigned += 1
                
        print(f"DEBUG: Security Duties Assigned: {count_assigned}")

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
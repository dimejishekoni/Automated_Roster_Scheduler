from openpyxl.styles import Side

# ================= 1. FILE PATHS =================
PDF_PATH = "2025.10.25 VSOU SOS.pdf"      
DB_PATH = "Staff Database.xlsx"           
OUTPUT_FILE = "Final_Roster.xlsx"

# ================= 2. GRID CONFIGURATION =================

GRID_CONFIG = {
    "start_hour": 5,
    "end_hour": 25,      
    "start_column": 6,   
    "blocks_per_hour": 4 
}

# ================= 3. COLOR DEFINITIONS =================
COLOR_BREAK    = "60497A"  # Dark Purple
COLOR_CSSI     = "9966FF"  # Light Purple
COLOR_SECURITY = "C00101"  # Deep Red
COLOR_DUTY     = "FFC000"  # Orange/Gold
COLOR_GENERAL  = "FFFF00"  # Yellow
PURPLE   = "9059FB"
WHITE    = "FFFFFF"
GREY     = "A895BE"
YELLOW   = "FEFF2F"
ORANGE   = "DE6220"
LIME     = "C6E0B4"
TAN      = "DDD9C4"
GREEN    = "6BE181"
PEACH    = "FABF8E"

COLORS = {
    "BREAK":    COLOR_BREAK,
    "CSSI":     COLOR_CSSI,
    "SECURITY": COLOR_SECURITY,
    "DUTY":     COLOR_DUTY,
    "GENERAL":  COLOR_GENERAL,
    "PURPLE"   : PURPLE, 
    "WHITE"    : WHITE,
    "GREY"     : GREY,
    "YELLOW"   : YELLOW,
    "ORANGE"   : ORANGE,
    "LIME"     : LIME,
    "TAN"      : TAN,
    "GREEN"    : GREEN,
    "PEACH"    : PEACH,
    
    # "VICTORIA_DEEP": "00B0F0",
    # "DISTRICT_DEEP": "33CC33",
    # "CARDINAL_DEEP": "FE66CC"
}

# Palettes for the Excel Renderer (Deep/Light/Banner)
PALETTES = {
    "Victoria": {"deep": "00B0F0", "light": "92CCDC", "banner": "00B0F0", "text": "FFFFFF"},
    "District": {"deep": "33CC33", "light": "C4D79B", "banner": "33CC33", "text": "FFFFFF"},
    "Cardinal": {"deep": "FE66CC", "light": "DA9694", "banner": "FE66CC", "text": "FFFFFF"}, 
}



# ================= 4. BORDERS =================
BLACK_THIN  = Side(border_style="thin",  color="000000")
BLACK_THICK = Side(border_style="thick", color="000000")

# ================= 5. DUTY SETTINGS =================
DUTY_PRIORITY = {
    "P3":          1,
    "PTI WB":      2,
    "PTI EB":      2,
    "SATS WB":     3,
    "SATS EB":     4,
    "ESCALATOR":   5,
    "VIP/MIP":     6,
    "SECURITY":    1, 
}


#MAPPING
DUTY_SECTION_MAP = {

    # Victoria
    "BN71": "Victoria", "BN21": "Victoria", "BN73": "Victoria", "BN27": "Victoria", 
    "BN28": "Victoria", "BN25": "Victoria", "BN26": "Victoria", "BN31": "Victoria", 
    "BN32": "Victoria", "BN76": "Victoria", "BN78": "Victoria", "BN37": "Victoria", 
    "BN38": "Victoria", "BN44": "Victoria", "BN39": "Victoria", "BN40": "Victoria", 
    "BN80": "Victoria", 
    
    # District
    "BN72": "District", "BN22": "District", "BN75": "District", "BN74": "District", 
    "BN29": "District", "BN33": "District", "BN77": "District", "BN34": "District", 
    "BN81": "District", "BN43": "District", "BN79": "District", "BN45": "District", 
    "BN41": "District", "BN82": "District", 
    
    # Cardinal
    "BN23": "Cardinal", "BN24": "Cardinal", "BN30": "Cardinal", "BN36": "Cardinal", 
    "BN42": "Cardinal", "BN35": "Cardinal"
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

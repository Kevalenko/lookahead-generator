"""
Boiler House — YarraBend, 60 Chandler Hwy Alphington
Job-specific configuration for the 2-Week Foreman Lookahead template.

To use for a new job: copy this file, rename it, and edit the values below.
Then run:  python build.py configs/your_new_job.py
"""

# ─── Project identity ─────────────────────────────────────────────────────────
PROJECT_NAME    = "BOILER HOUSE"
PROJECT_SUBTITLE = "60 Chandler Hwy, Alphington — YarraBend"
OUTPUT_FILENAME  = "BH_Foreman_Lookahead.xlsx"
VBA_FILENAME     = "BH_Foreman_Lookahead_macros.bas"

# ─── Working Day Calendar ─────────────────────────────────────────────────────
# Path to the .xlsx calendar file.
# Must have a sheet named "Table" with:
#   col D (index 3) = date (datetime object)
#   col F (index 5) = day type: 1=Working Day, "RDO"/"PH"/"AL"/"NW"/"Sun"=non-working
CALENDAR_PATH = "/root/.claude/uploads/89ce7542-a0e8-5d76-bb62-87fdabc00ed1/99cd69d9-Working_Day_Calendar_2025-2028_updated.xlsx"
CALENDAR_SHEET = "Table"

# ─── Sections ─────────────────────────────────────────────────────────────────
# Sections group activities visually (e.g. EARTHWORKS, SOG, SUSPENDED SLAB).
# Add, remove or rename freely.
# Section header colours (hex, no #).
SECTION_COLOURS = {
    "EARTHWORKS":      "6E2F0B",
    "SOG":             "0B4D6E",
    "SUSPENDED SLAB":  "0B6E3A",
    # Add more as needed, e.g.:
    # "STRUCTURE":     "4A235A",
    # "FAÇADE":        "0E6655",
}

# ─── Zone colours ─────────────────────────────────────────────────────────────
# One per zone label. Any zone not listed here gets a default grey.
ZONE_COLOURS = {
    "Zone 1":     "2C3E7A",
    "Zone 2":     "1A5276",
    "Zone 3A":    "117A65",
    "Zone 3B":    "1E8449",
    "Zone 4":     "784212",
    "Zone 5":     "6C3483",
    "BHE":        "1A5276",
    "BHW":        "117A65",
    "LGF/MEZ":    "6C3483",
    "BHW Zone A": "1A5276",
    "BHE Zone B": "784212",
}

# ─── Programme activities ──────────────────────────────────────────────────────
# Each dict = one parent row (one pour or activity).
# Keys:
#   section      : must match a key in SECTION_COLOURS (or will use default colour)
#   zone         : displayed as zone/area label; groups activities visually
#   name         : display name shown on parent row
#   target_wd    : working days from programme (int)
#   prog_start   : programme start date "YYYY-MM-DD"
#   prog_finish  : programme finish date "YYYY-MM-DD"
#   subtasks     : list of default sub-task names (foreman can rename/add more)
#                  These are suggestions — foreman edits freely on the sheet.

PROGRAMME = [
    # ── EARTHWORKS ──────────────────────────────────────────────────────────
    {
        "section": "EARTHWORKS", "zone": "Zone 1",
        "name": "Z1 Piling & Capping Beam",
        "target_wd": 26, "prog_start": "2026-05-28", "prog_finish": "2026-07-02",
        "subtasks": ["Piling Rig Setup", "Piling", "Capping Beam"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 1",
        "name": "Z1 Bulk Excavation",
        "target_wd": 4, "prog_start": "2026-06-03", "prog_finish": "2026-06-11",
        "subtasks": ["Strip & Load", "Shotcrete Walls"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 2",
        "name": "Z2 Piling & Capping Beam",
        "target_wd": 14, "prog_start": "2026-05-30", "prog_finish": "2026-06-19",
        "subtasks": ["Piling", "Capping Beam"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 2",
        "name": "Z2 Bulk Excavation",
        "target_wd": 76, "prog_start": "2026-06-20", "prog_finish": "2026-09-29",
        "subtasks": ["Drop 1", "Drop 2 (East Elev)", "Drop 3", "Drop 4"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 3A",
        "name": "Z3A Piling & Capping Beam",
        "target_wd": 11, "prog_start": "2026-05-28", "prog_finish": "2026-06-13",
        "subtasks": ["Piling", "Capping Beam"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 3A",
        "name": "Z3A Bulk Excavation",
        "target_wd": 84, "prog_start": "2026-06-15", "prog_finish": "2026-10-02",
        "subtasks": ["Drop 1", "Drop 2", "Drop 3", "Drop 4"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 3B",
        "name": "Z3B Piling & Capping Beam",
        "target_wd": 10, "prog_start": "2026-06-10", "prog_finish": "2026-06-23",
        "subtasks": ["Piling", "Capping Beam"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 3B",
        "name": "Z3B Bulk Excavation",
        "target_wd": 74, "prog_start": "2026-06-24", "prog_finish": "2026-09-29",
        "subtasks": ["Drop 1", "Drop 2", "Drop 3", "Drop 4"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 4",
        "name": "Z4 Piling & Capping Beam",
        "target_wd": 6, "prog_start": "2026-07-01", "prog_finish": "2026-07-09",
        "subtasks": ["Piling", "Capping Beam"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 4",
        "name": "Z4 Bulk Excavation",
        "target_wd": 28, "prog_start": "2026-09-12", "prog_finish": "2026-10-21",
        "subtasks": ["Drop 2", "Drop 3", "Crane HV Feed", "Tank to Pour 6"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 5",
        "name": "Z5 Piling & Capping Beam",
        "target_wd": 19, "prog_start": "2026-06-12", "prog_finish": "2026-07-07",
        "subtasks": ["Piling", "Capping Beam", "Foundation Piles"],
    },
    {
        "section": "EARTHWORKS", "zone": "Zone 5",
        "name": "Z5 Bulk Excavation",
        "target_wd": 9, "prog_start": "2026-07-08", "prog_finish": "2026-07-18",
        "subtasks": ["Strip & Load", "Shotcrete"],
    },
    # ── SOG ──────────────────────────────────────────────────────────────────
    {
        "section": "SOG", "zone": "BHE",
        "name": "BHE SOG Pour 1 (566m²)",
        "target_wd": 15, "prog_start": "2026-10-09", "prog_finish": "2026-10-29",
        "subtasks": ["Dig & Blind Footings", "Reo & Pour Footings",
                     "In-Ground Services", "Prep Slab", "Reo to Ground Slab", "Pour Concrete"],
    },
    {
        "section": "SOG", "zone": "BHE",
        "name": "BHE SOG Pour 2 (481m²)",
        "target_wd": 19, "prog_start": "2026-10-09", "prog_finish": "2026-11-07",
        "subtasks": ["Dig & Blind Footings", "Reo & Pour Footings",
                     "In-Ground Services", "Prep Slab", "Downturn Walls", "Reo to Ground Slab", "Pour Concrete"],
    },
    {
        "section": "SOG", "zone": "BHW",
        "name": "BHW SOG Pour 4 (692m²)",
        "target_wd": 11, "prog_start": "2026-10-17", "prog_finish": "2026-11-05",
        "subtasks": ["Dig & Blind Footings", "Reo & Pour Footings",
                     "In-Ground Services", "Prep Slab", "Reo to Ground Slab", "Pour Concrete"],
    },
    {
        "section": "SOG", "zone": "BHW",
        "name": "BHW SOG Pour 5 (682m²)",
        "target_wd": 12, "prog_start": "2026-10-19", "prog_finish": "2026-11-07",
        "subtasks": ["Dig & Blind Footings", "Reo & Pour Footings",
                     "In-Ground Services", "Torpedo Tank Bulk",
                     "Prep Slab", "Reo to Ground Slab", "Pour Concrete"],
    },
    {
        "section": "SOG", "zone": "BHW",
        "name": "BHW SOG Pour 6 (700m²)",
        "target_wd": 15, "prog_start": "2026-10-22", "prog_finish": "2026-11-14",
        "subtasks": ["Dig & Blind Footings", "Downturn Walls on Pour 4",
                     "Reo & Pour Footings", "In-Ground Services",
                     "Prep Slab", "Reo to Ground Slab", "Pour Concrete"],
    },
    {
        "section": "SOG", "zone": "LGF/MEZ",
        "name": "LGF SOG Pour 8 / Zone B (314m²)",
        "target_wd": 10, "prog_start": "2026-11-09", "prog_finish": "2026-11-21",
        "subtasks": ["Dig & Blind Footings", "Reo & Pour Footings",
                     "Precast Support Ramp", "Prep Slab", "Reo to Ground Slab", "Pour Concrete"],
    },
    {
        "section": "SOG", "zone": "LGF/MEZ",
        "name": "MEZ Pour 1 — B1 to LG Ramp (Lost Form/SOG)",
        "target_wd": 6, "prog_start": "2026-11-09", "prog_finish": "2026-11-17",
        "subtasks": ["Dig & Blind", "Reo & Pour Footings",
                     "Install Precast", "Prep/Lost Form", "Reo to Slab", "Pour Concrete"],
    },
    {
        "section": "SOG", "zone": "LGF/MEZ",
        "name": "SOG Zone B Pour 3 (Ramps)",
        "target_wd": 4, "prog_start": "2026-11-18", "prog_finish": "2026-11-23",
        "subtasks": ["Precast Panels", "Formwork", "Reo to Slab", "Pour Concrete"],
    },
    # ── SUSPENDED SLAB ───────────────────────────────────────────────────────
    {
        "section": "SUSPENDED SLAB", "zone": "BHW Zone A",
        "name": "B1 Mezz Zone A Pour 2 (859m²)",
        "target_wd": 14, "prog_start": "2026-11-06", "prog_finish": "2026-11-24",
        "subtasks": ["Precast Walls & Columns", "FRP Stairs",
                     "Frames & Formwork", "Penetrations",
                     "Bottom Reo & PT", "Top Reo", "Pour Concrete", "PT Stressing"],
    },
    {
        "section": "SUSPENDED SLAB", "zone": "BHW Zone A",
        "name": "B1 Mezz Zone A Pour 3 (745m²)",
        "target_wd": 17, "prog_start": "2026-11-17", "prog_finish": "2026-12-08",
        "subtasks": ["Precast Walls & Columns", "FRP Stairs",
                     "Frames & Formwork", "Penetrations",
                     "Bottom Reo & PT", "Top Reo", "Pour Concrete", "PT Stressing"],
    },
    {
        "section": "SUSPENDED SLAB", "zone": "BHW Zone A",
        "name": "LGF Zone A Pour 1 — Dbl Height (500m²)",
        "target_wd": 25, "prog_start": "2026-11-17", "prog_finish": "2026-12-17",
        "subtasks": ["Precast Walls & Columns", "FRP Stairs",
                     "Frames & Formwork", "Penetrations",
                     "Bottom Reo & PT", "Top Reo", "Pour Concrete", "PT Stressing"],
    },
    {
        "section": "SUSPENDED SLAB", "zone": "BHW Zone A",
        "name": "LGF Zone A Pour 2",
        "target_wd": 28, "prog_start": "2026-12-14", "prog_finish": "2027-01-20",
        "subtasks": ["Frames & Formwork", "Bottom Reo & PT",
                     "Top Reo", "Pour Concrete", "PT Stressing"],
    },
    {
        "section": "SUSPENDED SLAB", "zone": "BHW Zone A",
        "name": "LGF Zone A Pour 3",
        "target_wd": 21, "prog_start": "2026-12-14", "prog_finish": "2027-01-20",
        "subtasks": ["Frames & Formwork", "Bottom Reo & PT",
                     "Top Reo", "Pour Concrete", "PT Stressing"],
    },
    {
        "section": "SUSPENDED SLAB", "zone": "BHW Zone A",
        "name": "LGF Zone A Pour 4",
        "target_wd": 36, "prog_start": "2026-11-17", "prog_finish": "2027-01-20",
        "subtasks": ["Frames & Formwork", "Bottom Reo & PT",
                     "Top Reo", "Pour Concrete", "PT Stressing"],
    },
    {
        "section": "SUSPENDED SLAB", "zone": "BHE Zone B",
        "name": "B1 BHE LGF Pour 5 (Dbl Ht) — LGF Zone B",
        "target_wd": 32, "prog_start": "2026-11-11", "prog_finish": "2026-12-21",
        "subtasks": ["Precast Walls & Columns", "Frames & Formwork",
                     "Penetrations", "Bottom Reo & PT",
                     "Top Reo", "Pour Concrete", "PT Stressing"],
    },
    {
        "section": "SUSPENDED SLAB", "zone": "BHE Zone B",
        "name": "LGF Zone B Pour 6 (UGF BHE)",
        "target_wd": 25, "prog_start": "2026-11-30", "prog_finish": "2027-01-20",
        "subtasks": ["Frames & Formwork", "Bottom Reo & PT",
                     "Top Reo", "Pour Concrete", "PT Stressing"],
    },
    {
        "section": "SUSPENDED SLAB", "zone": "BHE Zone B",
        "name": "LGF Zone B Pour 7 (L1 BHE)",
        "target_wd": 25, "prog_start": "2026-12-12", "prog_finish": "2027-02-04",
        "subtasks": ["Frames & Formwork", "Bottom Reo & PT",
                     "Top Reo", "Pour Concrete", "PT Stressing"],
    },
    {
        "section": "SUSPENDED SLAB", "zone": "BHE Zone B",
        "name": "B1 Ramp to LGF",
        "target_wd": 6, "prog_start": "2026-11-09", "prog_finish": "2026-11-17",
        "subtasks": ["Formwork", "Reo Fix", "Pour Concrete", "Strip & Clean"],
    },
]

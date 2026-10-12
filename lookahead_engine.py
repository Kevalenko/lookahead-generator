"""
Foreman 2-Week Lookahead Template Engine
=========================================
Generic, reusable. Contains zero job-specific data.

Usage:
    import importlib, sys
    sys.path.insert(0, ".")
    cfg = importlib.import_module("configs.boiler_house")   # or any other config
    from lookahead_engine import build
    build(cfg)

The config module must expose:
    PROJECT_NAME      str    "BOILER HOUSE"
    PROJECT_SUBTITLE  str    "60 Chandler Hwy, Alphington"
    OUTPUT_FILENAME   str    "BH_Foreman_Lookahead.xlsx"
    VBA_FILENAME      str    "BH_Foreman_Lookahead_macros.bas"
    CALENDAR_PATH     str    path to working-day calendar .xlsx
    CALENDAR_SHEET    str    sheet name in calendar file (default "Table")
    SECTION_COLOURS   dict   {section_name: hex_colour_no_hash}
    ZONE_COLOURS      dict   {zone_name: hex_colour_no_hash}
    PROGRAMME         list   of activity dicts — see configs/boiler_house.py

Each activity dict:
    section     str    must exist in SECTION_COLOURS (or fallback default)
    zone        str    zone / area label
    name        str    activity display name
    target_wd   int    working days from programme
    prog_start  str    "YYYY-MM-DD"
    prog_finish str    "YYYY-MM-DD"
    subtasks    list   default sub-task name strings
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule
from datetime import date, datetime, timedelta

# ─── Design tokens ────────────────────────────────────────────────────────────
# All design choices live here. Change these to rebrand the template.
T = dict(
    # Chrome
    DARK_NAVY   = "1A1A2E",
    NAVY        = "16213E",
    ACCENT      = "0F3460",
    GOLD        = "C9A84C",
    WHITE       = "FFFFFF",
    LIGHT_GREY  = "F5F5F5",
    MID_GREY    = "D0D0D0",
    DARK_GREY   = "5A5A5A",
    # Activity row chrome
    PARENT_BG   = "2E4057",
    SUBTASK_BG  = "F9F9F9",
    BLANK_ROW   = "FFFFF0",
    NONWORK_BG  = "EBEBEB",
    WEEK1_BG    = "EBF5FB",
    WEEK2_BG    = "F4ECF7",
    WEEK1_HDR   = "2E86C1",
    WEEK2_HDR   = "7D3C98",
    # RAG
    GREEN       = "C6EFCE", GREEN_F  = "006100",
    AMBER       = "FFEB9C", AMBER_F  = "9C5700",
    RED         = "FFC7CE", RED_F    = "9C0006",
    # Fallback zone / section colour when not in config
    FALLBACK_ZONE    = "444444",
    FALLBACK_SECTION = "333333",
)

# Day-column offsets from Week Start Date (skip Sunday = +6)
DAY_OFFSETS = [0, 1, 2, 3, 4, 5, 7, 8, 9, 10, 11, 12]

# ─── Style helpers ────────────────────────────────────────────────────────────
def _fill(hex_col):
    return PatternFill("solid", fgColor=hex_col)

def _font(name="Arial", size=9, bold=False, color="000000", italic=False):
    return Font(name=name, size=size, bold=bold, color=color, italic=italic)

def _aln(h="left", v="center", wrap=False, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)

def _bdr(top=False, bottom=False, left=False, right=False, thick=False):
    s = "medium" if thick else "thin"
    return Border(
        top    = Side(style=s) if top    else None,
        bottom = Side(style=s) if bottom else None,
        left   = Side(style=s) if left   else None,
        right  = Side(style=s) if right  else None,
    )

def _header_cell(ws, addr, value, bg, fg=None, size=9, bold=True, h="center", merge_to=None):
    if merge_to:
        ws.merge_cells(f"{addr}:{merge_to}")
    c = ws[addr]
    c.value = value
    c.font = _font(size=size, bold=bold, color=fg or T["WHITE"])
    c.fill = _fill(bg)
    c.alignment = _aln(h, "center")
    return c

# ─── Calendar loading ─────────────────────────────────────────────────────────
def _load_calendar(cal_path, sheet_name="Table"):
    """
    Returns (working: set[date], non_working: dict[date -> str])
    Expects calendar sheet with:
        col D (index 3) = date
        col F (index 5) = type: 1=WD, str=non-working label
    """
    wb = openpyxl.load_workbook(cal_path, data_only=True)
    ws = wb[sheet_name]
    working, non_working = set(), {}
    for row in ws.iter_rows(min_row=4, values_only=True):
        if not row or row[3] is None:
            continue
        d = row[3]
        if isinstance(d, datetime):
            d = d.date()
        if not isinstance(d, date):
            continue
        t = row[5]
        if t == 1:
            working.add(d)
        elif t is not None:
            non_working[d] = str(t)
    return working, non_working

def _display_days(week_start, working, non_working):
    """12 consecutive days Mon-Sat x2, skipping Sundays."""
    days, cur = [], week_start
    while len(days) < 12:
        if cur.weekday() != 6:  # skip Sunday
            dtype = "WD" if cur in working else non_working.get(cur, "WD")
            days.append((cur, dtype))
        cur += timedelta(days=1)
    return days

def _next_monday():
    today = date.today()
    delta = (7 - today.weekday()) % 7
    return today + timedelta(days=delta if delta else 7)

# ─── VBA source ───────────────────────────────────────────────────────────────
def _vba_code(project_name):
    return f'''\
Attribute VB_Name = "LookaheadMacros"
' ════════════════════════════════════════════════════════════════════
'  {project_name} — Foreman 2-Week Lookahead  |  VBA Macro Module
'
'  HOW TO INSTALL:
'    1. Save the workbook as .xlsm (macro-enabled)
'    2. Press Alt+F11 to open the VBA editor
'    3. Insert > Module, paste this file's contents
'    4. Back on the sheet: Developer > Insert > Button (Form Control)
'       Draw a button, assign macro "NextWeek"  → label it "NEXT WEEK →"
'       Repeat for "PrevWeek" and "SaveThisWeek"
'    5. Save as .xlsm
' ════════════════════════════════════════════════════════════════════
Option Explicit

' ── Move forward 1 week ──────────────────────────────────────────────
Sub NextWeek()
    With ActiveSheet.Range("C2")
        If IsDate(.Value) Then .Value = .Value + 7
    End With
End Sub

' ── Move back 1 week ─────────────────────────────────────────────────
Sub PrevWeek()
    With ActiveSheet.Range("C2")
        If IsDate(.Value) Then .Value = .Value - 7
    End With
End Sub

' ── Archive this week as a snapshot tab ──────────────────────────────
Sub SaveThisWeek()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If Not IsDate(ws.Range("C2").Value) Then
        MsgBox "No valid Week Start Date in cell C2.", vbExclamation
        Exit Sub
    End If

    Dim weekStart As Date
    weekStart = CDate(ws.Range("C2").Value)

    ' ISO week number (Mon start, first week with 4+ days)
    Dim wkNum As Integer
    wkNum = DatePart("ww", weekStart, vbMonday, vbFirstFourDays)
    Dim tabName As String
    tabName = "Week " & wkNum & " - " & Format(weekStart, "DD MMM")

    ' Overwrite existing tab of same name if user agrees
    Dim existing As Worksheet
    On Error Resume Next
    Set existing = ActiveWorkbook.Sheets(tabName)
    On Error GoTo 0
    If Not existing Is Nothing Then
        If MsgBox("Tab '" & tabName & "' already exists. Overwrite?", _
                  vbYesNo + vbQuestion, "Overwrite?") = vbNo Then Exit Sub
        Application.DisplayAlerts = False
        existing.Delete
        Application.DisplayAlerts = True
    End If

    ' Copy active sheet to end
    ws.Copy After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    Dim snap As Worksheet
    Set snap = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    snap.Name = tabName

    ' Stamp saved timestamp in D2
    With snap.Range("D2")
        .Value = "SNAPSHOT: " & Format(Now(), "DD/MM/YYYY HH:MM")
        .Font.Color = RGB(200, 100, 0)
        .Font.Bold = True
    End With

    ws.Activate
    MsgBox "Snapshot saved as tab '" & tabName & "'", vbInformation
End Sub
'''

# ─── Sheet builders ───────────────────────────────────────────────────────────
def _build_lookahead_sheet(ws, cfg, working, non_working, default_start, display_days):
    """Build the main 2-Week Lookahead sheet."""

    # Column widths
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 14
    for i in range(12):
        ws.column_dimensions[get_column_letter(3 + i)].width = 5.8
    ws.column_dimensions["O"].width = 9
    ws.column_dimensions["P"].width = 9
    ws.column_dimensions["Q"].width = 8
    ws.column_dimensions["R"].width = 20

    # ── Row 1: Project title ─────────────────────────────────────────────────
    ws.row_dimensions[1].height = 34
    _header_cell(ws, "A1", cfg.PROJECT_NAME + " — FOREMAN 2-WEEK LOOKAHEAD",
                 T["DARK_NAVY"], size=13, merge_to="N1")
    _header_cell(ws, "O1", cfg.PROJECT_SUBTITLE,
                 T["DARK_NAVY"], fg=T["GOLD"], size=9, bold=False, merge_to="R1")

    # ── Row 2: Week Start Date input ─────────────────────────────────────────
    ws.row_dimensions[2].height = 22
    _header_cell(ws, "A2", "WEEK START DATE  ▶", T["NAVY"], fg=T["GOLD"], h="right")
    ws["B2"].value = "←"
    ws["B2"].font = _font(color="AAAAAA")
    ws["B2"].fill = _fill(T["NAVY"])
    ws["B2"].alignment = _aln("center", "center")

    # C2 = the date driver (formula source for all day headers)
    ws["C2"].value = default_start
    ws["C2"].number_format = "DD/MM/YYYY"
    ws["C2"].font = Font(name="Arial", size=10, bold=True, color=T["GOLD"])
    ws["C2"].fill = _fill(T["NAVY"])
    ws["C2"].alignment = _aln("center", "center")

    ws.merge_cells("D2:N2")
    ws["D2"].value = "  Use macro buttons: ← PREV WEEK  |  NEXT WEEK →  |  SAVE THIS WEEK"
    ws["D2"].font = _font(size=8, italic=True, color="AAAAAA")
    ws["D2"].fill = _fill(T["NAVY"])
    ws["D2"].alignment = _aln("left", "center")

    for ltr, lbl in [("O", "← PREV WEEK"), ("P", "NEXT WEEK →"), ("Q", ""), ("R", "SAVE THIS WEEK")]:
        c = ws[f"{ltr}2"]
        c.value = lbl
        c.font = _font(size=8, bold=True, color=T["WHITE"])
        c.fill = _fill(T["ACCENT"])
        c.alignment = _aln("center", "center")

    # ── Row 3: Week band labels ──────────────────────────────────────────────
    ws.row_dimensions[3].height = 15
    _header_cell(ws, "C3", "◀  WEEK 1", T["WEEK1_HDR"], merge_to="H3")
    _header_cell(ws, "I3", "WEEK 2  ▶", T["WEEK2_HDR"], merge_to="N3")
    for ltr, lbl in [("A","ACTIVITY / TRADE"),("B","TRADE"),
                      ("O","FMN WD"),("P","TGT WD"),("Q","VAR"),("R","STATUS")]:
        c = ws[f"{ltr}3"]
        c.value = lbl
        c.font = _font(size=8, bold=True, color=T["WHITE"])
        c.fill = _fill("404040")
        c.alignment = _aln("left" if ltr == "A" else "center", "center")

    # ── Row 4: Date headers (formula-driven from C2) ─────────────────────────
    ws.row_dimensions[4].height = 30
    day_cols = []
    for i, (d, dtype) in enumerate(display_days):
        col = get_column_letter(3 + i)
        c = ws[f"{col}4"]
        c.value = f"=$C$2+{DAY_OFFSETS[i]}"
        c.number_format = 'DDD\n"("D")"'
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if dtype in ("RDO", "PH", "AL", "NW"):
            c.fill = _fill("BDC3C7"); c.font = _font(size=8, bold=True, color="7F8C8D")
        elif i < 6:
            c.fill = _fill(T["WEEK1_HDR"]); c.font = _font(size=8, bold=True, color=T["WHITE"])
        else:
            c.fill = _fill(T["WEEK2_HDR"]); c.font = _font(size=8, bold=True, color=T["WHITE"])
        day_cols.append((col, d, dtype))

    for ltr, lbl in [("A","ACTIVITY"),("B","TRADE"),
                      ("O","FMN WD"),("P","TGT WD"),("Q","VAR"),("R","STATUS")]:
        c = ws[f"{ltr}4"]
        c.value = lbl
        c.font = _font(size=8, bold=True, color=T["WHITE"])
        c.fill = _fill("303030")
        c.alignment = _aln("left" if ltr == "A" else "center", "center")

    # ── Data rows ─────────────────────────────────────────────────────────────
    current_row = 5
    current_section = None
    current_zone = None

    for act in cfg.PROGRAMME:
        section = act["section"]
        zone    = act["zone"]

        # Section header (on first appearance of each section)
        if section != current_section:
            current_section = section
            current_zone = None  # force zone header on first zone in section
            ws.row_dimensions[current_row].height = 18
            ws.merge_cells(f"A{current_row}:R{current_row}")
            c = ws[f"A{current_row}"]
            c.value = f"  ▌ {section}"
            c.font = Font(name="Arial", size=10, bold=True, color=T["WHITE"])
            c.fill = _fill(cfg.SECTION_COLOURS.get(section, T["FALLBACK_SECTION"]))
            c.alignment = _aln("left", "center")
            current_row += 1

        # Zone label (on zone change within section)
        if zone != current_zone:
            current_zone = zone
            z_colour = cfg.ZONE_COLOURS.get(zone, T["FALLBACK_ZONE"])
            ws.row_dimensions[current_row].height = 14
            ws.merge_cells(f"A{current_row}:R{current_row}")
            c = ws[f"A{current_row}"]
            c.value = f"    ZONE / AREA: {zone}"
            c.font = _font(size=8, bold=True, color=T["WHITE"])
            c.fill = _fill(z_colour)
            c.alignment = _aln("left", "center")
            current_row += 1

        # Row indices for this activity
        parent_row  = current_row
        sub_start   = parent_row + 1
        sub_end     = sub_start + len(act["subtasks"]) - 1
        blank_row   = sub_end + 1
        group_end   = blank_row   # outline group: sub_start..group_end

        # ── Parent row ────────────────────────────────────────────────────
        ws.row_dimensions[parent_row].height = 20
        c = ws[f"A{parent_row}"]
        c.value = f"{act['name']}  [{act['prog_start'][:10]} → {act['prog_finish'][:10]}]"
        c.font = Font(name="Arial", size=9, bold=True, color=T["WHITE"])
        c.fill = _fill(T["PARENT_BG"])
        c.alignment = _aln("left", "center")
        c.border = _bdr(top=True, bottom=True, left=True, thick=True)

        ws[f"B{parent_row}"].fill = _fill(T["PARENT_BG"])
        ws[f"B{parent_row}"].border = _bdr(bottom=True, right=True)

        for i, (col, d, dtype) in enumerate(day_cols):
            c = ws[f"{col}{parent_row}"]
            c.fill = _fill("C8D6E3" if i < 6 else "C8B8D4")
            c.border = _bdr(bottom=True, right=True)

        # O: SUM of sub-task FMN WD (including blank row)
        c = ws[f"O{parent_row}"]
        c.value = f"=SUM(O{sub_start}:O{group_end})"
        c.font = Font(name="Arial", size=9, bold=True, color=T["WHITE"])
        c.fill = _fill(T["PARENT_BG"]); c.alignment = _aln("center", "center")
        c.number_format = "0"; c.border = _bdr(bottom=True, left=True, right=True)

        # P: Target WD (hardcoded from programme)
        c = ws[f"P{parent_row}"]
        c.value = act["target_wd"]
        c.font = Font(name="Arial", size=9, bold=True, color=T["GOLD"])
        c.fill = _fill(T["PARENT_BG"]); c.alignment = _aln("center", "center")
        c.number_format = "0"; c.border = _bdr(bottom=True, right=True)

        # Q: Variance
        c = ws[f"Q{parent_row}"]
        c.value = f'=IF(O{parent_row}=0,"-",O{parent_row}-P{parent_row})'
        c.font = Font(name="Arial", size=9, bold=True, color=T["WHITE"])
        c.fill = _fill(T["PARENT_BG"]); c.alignment = _aln("center", "center")
        c.number_format = '+0;-0;"-"'; c.border = _bdr(bottom=True, right=True)

        # R: RAG status
        c = ws[f"R{parent_row}"]
        c.value = (f'=IF(O{parent_row}=0,"—",'
                   f'IF(O{parent_row}<=P{parent_row},"✅ ON PROGRAMME",'
                   f'IF(O{parent_row}<=P{parent_row}*1.2,"⚠ WATCH","🔴 AT RISK")))')
        c.font = _font(size=9, bold=True)
        c.fill = _fill(T["PARENT_BG"]); c.alignment = _aln("center", "center")
        c.border = _bdr(bottom=True, right=True)

        current_row += 1

        # ── Sub-task rows ────────────────────────────────────────────────
        for task_name in act["subtasks"]:
            row = current_row
            ws.row_dimensions[row].height = 17
            ws.row_dimensions[row].outline_level = 1

            c = ws[f"A{row}"]
            c.value = f"      {task_name}"
            c.font = _font(size=9, color="333333")
            c.fill = _fill(T["SUBTASK_BG"]); c.alignment = _aln("left", "center")
            c.border = _bdr(bottom=True, left=True)

            ws[f"B{row}"].fill = _fill(T["SUBTASK_BG"])
            ws[f"B{row}"].border = _bdr(bottom=True, right=True)

            for i, (col, d, dtype) in enumerate(day_cols):
                c = ws[f"{col}{row}"]
                c.alignment = _aln("center", "center")
                c.border = _bdr(bottom=True, right=True)
                if dtype in ("RDO", "PH", "AL", "NW"):
                    c.value = dtype
                    c.fill = _fill(T["NONWORK_BG"])
                    c.font = _font(size=8, color="999999")
                elif i < 6:
                    c.fill = _fill(T["WEEK1_BG"])
                    c.font = _font(size=10, bold=True, color="1A1A2E")
                else:
                    c.fill = _fill(T["WEEK2_BG"])
                    c.font = _font(size=10, bold=True, color="1A1A2E")

            # O: count X marks (foreman's WD)
            c = ws[f"O{row}"]
            c.value = f'=COUNTIF(C{row}:N{row},"X")'
            c.font = _font(size=9, bold=True, color="1F4E79")
            c.fill = _fill("DAE8FC"); c.alignment = _aln("center", "center")
            c.number_format = "0"; c.border = _bdr(bottom=True, left=True, right=True)

            for ltr in ["P", "Q", "R"]:
                ws[f"{ltr}{row}"].fill = _fill(T["SUBTASK_BG"])
                ws[f"{ltr}{row}"].border = _bdr(bottom=True, right=True)

            current_row += 1

        # ── Blank "add task" row ─────────────────────────────────────────
        row = current_row
        ws.row_dimensions[row].height = 17
        ws.row_dimensions[row].outline_level = 1

        c = ws[f"A{row}"]
        c.value = "      ← Add task here  (type name, mark X for working days)"
        c.font = _font(size=8, italic=True, color="BBBBBB")
        c.fill = _fill(T["BLANK_ROW"]); c.alignment = _aln("left", "center")
        c.border = _bdr(bottom=True, left=True)

        ws[f"B{row}"].fill = _fill(T["BLANK_ROW"])
        ws[f"B{row}"].border = _bdr(bottom=True, right=True)

        for i, (col, d, dtype) in enumerate(day_cols):
            c = ws[f"{col}{row}"]
            c.alignment = _aln("center", "center")
            c.border = _bdr(bottom=True, right=True)
            if dtype in ("RDO", "PH", "AL", "NW"):
                c.value = dtype
                c.fill = _fill(T["NONWORK_BG"])
                c.font = _font(size=8, color="BBBBBB")
            else:
                c.fill = _fill(T["BLANK_ROW"])

        c = ws[f"O{row}"]
        c.value = f'=COUNTIF(C{row}:N{row},"X")'
        c.font = _font(size=9, color="BBBBBB")
        c.fill = _fill(T["BLANK_ROW"]); c.alignment = _aln("center", "center")
        c.number_format = "0"; c.border = _bdr(bottom=True, left=True, right=True)
        for ltr in ["P", "Q", "R"]:
            ws[f"{ltr}{row}"].fill = _fill(T["BLANK_ROW"])
            ws[f"{ltr}{row}"].border = _bdr(bottom=True, right=True)

        current_row += 1

        # Spacer
        ws.row_dimensions[current_row].height = 3
        ws.merge_cells(f"A{current_row}:R{current_row}")
        ws[f"A{current_row}"].fill = _fill(T["MID_GREY"])
        current_row += 1

    # ── RAG conditional formatting ────────────────────────────────────────────
    last_row = current_row - 1
    rng = f"R5:R{last_row}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=['R5="✅ ON PROGRAMME"'],
        fill=PatternFill("solid", fgColor=T["GREEN"]),
        font=Font(color=T["GREEN_F"], bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=['R5="⚠ WATCH"'],
        fill=PatternFill("solid", fgColor=T["AMBER"]),
        font=Font(color=T["AMBER_F"], bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=['R5="🔴 AT RISK"'],
        fill=PatternFill("solid", fgColor=T["RED"]),
        font=Font(color=T["RED_F"], bold=True)))

    # Sheet settings
    ws.sheet_properties.outlinePr.summaryBelow = False   # +/- above group
    ws.sheet_properties.outlinePr.summaryRight = False
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.freeze_panes = "C5"


def _build_programme_targets_sheet(ws, cfg):
    """Programme Targets reference tab — all activities with dates and status."""
    today = date.today()

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 42
    ws.column_dimensions["D"].width = 13
    ws.column_dimensions["E"].width = 13
    ws.column_dimensions["F"].width = 9
    ws.column_dimensions["G"].width = 20

    ws.row_dimensions[1].height = 28
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = f"{cfg.PROJECT_NAME} — PROGRAMME TARGETS REFERENCE"
    c.font = Font(name="Arial", size=11, bold=True, color=T["WHITE"])
    c.fill = _fill(T["DARK_NAVY"]); c.alignment = _aln("center", "center")

    hdrs = ["", "SECTION", "ACTIVITY", "PROG START", "PROG FINISH", "TGT WD", "STATUS"]
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(2, ci)
        c.value = h
        c.font = _font(size=9, bold=True, color=T["WHITE"])
        c.fill = _fill("404040"); c.alignment = _aln("center", "center")

    row = 3
    for act in cfg.PROGRAMME:
        ps = datetime.strptime(act["prog_start"], "%Y-%m-%d").date()
        pf = datetime.strptime(act["prog_finish"], "%Y-%m-%d").date()

        if pf < today:
            status, status_c = "Complete / Past", "808080"
            row_bg = "F5F5F5"
        elif ps <= today:
            status, status_c = "▶ IN PROGRESS", "0070C0"
            row_bg = "EBF5FB"
        elif (ps - today).days <= 14:
            status, status_c = "⚠ Starting soon", "CC4400"
            row_bg = "FFF3E6"
        else:
            status, status_c = f"Starts {ps.strftime('%d %b %Y')}", "AAAAAA"
            row_bg = T["WHITE"]

        ws.row_dimensions[row].height = 16
        for ci, (val, num_fmt) in enumerate([
            ("",            None),
            (act["section"], None),
            (act["name"],    None),
            (ps,            "DD/MM/YYYY"),
            (pf,            "DD/MM/YYYY"),
            (act["target_wd"], "0"),
            (status,        None),
        ], 1):
            c = ws.cell(row, ci)
            c.value = val
            c.fill = _fill(row_bg)
            c.alignment = _aln("left" if ci in (2, 3, 7) else "center", "center")
            c.border = _bdr(bottom=True, right=True)
            if ci == 7:
                c.font = _font(size=8, bold=True, color=status_c)
            elif ci == 6:
                c.font = _font(size=9, bold=True, color="1F4E79")
            else:
                c.font = _font(size=9, color="333333")
            if num_fmt:
                c.number_format = num_fmt

        # Blue pip in col A for in-progress activities
        if ps <= today <= pf:
            ws.cell(row, 1).value = "▶"
            ws.cell(row, 1).fill = _fill("0070C0")
            ws.cell(row, 1).font = _font(bold=True, color=T["WHITE"])
            ws.cell(row, 1).alignment = _aln("center", "center")

        row += 1

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"


def _build_calendar_sheet(ws, working, non_working):
    """WD Calendar reference tab — rolling 90-day view."""
    for ci, h in enumerate(["Date", "Day", "Type", "Note"], 1):
        c = ws.cell(1, ci)
        c.value = h
        c.font = _font(size=9, bold=True, color=T["WHITE"])
        c.fill = _fill(T["DARK_NAVY"]); c.alignment = _aln("center", "center")

    ws.column_dimensions["A"].width = 13
    ws.column_dimensions["B"].width = 7
    ws.column_dimensions["C"].width = 8
    ws.column_dimensions["D"].width = 22

    ref = date.today() - timedelta(days=7)
    for row in range(2, 102):
        is_sun = ref.weekday() == 6
        dtype  = "Sun" if is_sun else ("WD" if ref in working else non_working.get(ref, "WD"))
        note   = {"RDO": "36-Hour RDO (Industry)", "PH": "Public Holiday",
                  "NW": "Non-Working Day", "AL": "Annual Leave / Shutdown"}.get(dtype, "")
        bg = {"RDO": "FFE0E0", "PH": "FFD0A0", "NW": "EEEEEE",
              "Sun": "F0F0F0"}.get(dtype, T["WHITE"])

        ws[f"A{row}"].value = ref; ws[f"A{row}"].number_format = "DD/MM/YYYY"
        ws[f"B{row}"].value = ref.strftime("%a")
        ws[f"C{row}"].value = dtype
        ws[f"D{row}"].value = note
        for col in ["A", "B", "C", "D"]:
            ws[f"{col}{row}"].fill = _fill(bg)
            ws[f"{col}{row}"].font = _font(size=9, color="333333" if bg == T["WHITE"] else "666666")
            ws[f"{col}{row}"].alignment = _aln("center" if col != "D" else "left", "center")
        ref += timedelta(days=1)

    ws.sheet_view.showGridLines = False


def _build_guide_sheet(ws, cfg):
    """How To Use tab."""
    ws.column_dimensions["A"].width = 88
    today_str = date.today().strftime("%d/%m/%Y")
    rows = [
        (f"{cfg.PROJECT_NAME} — FOREMAN 2-WEEK LOOKAHEAD  |  HOW TO USE", True, 13, T["DARK_NAVY"], T["WHITE"]),
        ("", False, 9, T["WHITE"], "AAAAAA"),
        ("WEEK NAVIGATION", True, 11, "1F4E79", T["WHITE"]),
        ("1.  Cell C2 (WEEK START DATE) drives all 12 day-column headers automatically — Mon–Sat across 2 weeks.", False, 10, T["WHITE"], "222222"),
        ("2.  Click ← PREV WEEK macro button to step back 7 days.", False, 10, T["WHITE"], "222222"),
        ("3.  Click NEXT WEEK → macro button to step forward 7 days.", False, 10, T["WHITE"], "222222"),
        ("4.  Click SAVE THIS WEEK to archive a snapshot of this week as a new tab (e.g. 'Week 42 – 12 Oct').", False, 10, T["WHITE"], "222222"),
        ("    Macro buttons: assign NextWeek / PrevWeek / SaveThisWeek from the .bas VBA module (see install steps in the .bas file).", False, 9, T["WHITE"], "555555"),
        ("", False, 9, T["WHITE"], "AAAAAA"),
        ("MARKING PROGRESS", True, 11, "117A65", T["WHITE"]),
        ("5.  Type  X  in any white/tinted day cell on a sub-task row to record a day worked.", False, 10, T["WHITE"], "222222"),
        ("6.  Grey cells (RDO / PH) are pre-labelled — do not mark these.", False, 10, T["WHITE"], "222222"),
        ("7.  FMN WD column auto-counts X entries per sub-task row (COUNTIF formula).", False, 10, T["WHITE"], "222222"),
        ("8.  The parent row FMN WD sums all sub-task rows below it automatically.", False, 10, T["WHITE"], "222222"),
        ("9.  TGT WD (Target Working Days) comes from the programme — do NOT edit these.", False, 10, T["WHITE"], "222222"),
        ("", False, 9, T["WHITE"], "AAAAAA"),
        ("ADDING / EDITING SUB-TASKS", True, 11, "784212", T["WHITE"]),
        ("10. Click – on the left margin to collapse a completed pour's sub-task rows.", False, 10, T["WHITE"], "222222"),
        ("11. Click + to expand again.", False, 10, T["WHITE"], "222222"),
        ("12. Type over any sub-task name to rename it for your crew.", False, 10, T["WHITE"], "222222"),
        ("13. The yellow ← Add task here row: type a name and mark X to use as an extra sub-task.", False, 10, T["WHITE"], "222222"),
        ("14. To insert more rows: right-click a sub-task row → Insert Row → type task name.", False, 10, T["WHITE"], "222222"),
        ("    In column O of the new row, type:  =COUNTIF(C##:N##,\"X\")   (replace ## with the row number).", False, 9, T["WHITE"], "555555"),
        ("", False, 9, T["WHITE"], "AAAAAA"),
        ("RAG STATUS  (column R — auto on parent rows only)", True, 11, "6C3483", T["WHITE"]),
        ("✅ ON PROGRAMME  =  Foreman WD ≤ Target WD", False, 10, T["GREEN"], T["GREEN_F"]),
        ("⚠  WATCH         =  Foreman WD is 1–20% over Target WD", False, 10, "FFFACD", T["AMBER_F"]),
        ("🔴 AT RISK        =  Foreman WD is >20% over Target WD", False, 10, "FFCCCC", T["RED_F"]),
        ("—  (dash)         =  No X entries yet for this activity (FMN WD = 0)", False, 9, T["WHITE"], "888888"),
        ("", False, 9, T["WHITE"], "AAAAAA"),
        ("PROGRAMME TARGETS TAB", True, 11, T["DARK_GREY"], T["WHITE"]),
        ("See 'Programme Targets' tab for all activities with start/finish dates and live status.", False, 10, T["WHITE"], "222222"),
        ("▶ IN PROGRESS = activity currently underway on site today.", False, 10, T["WHITE"], "222222"),
        ("⚠ Starting soon = activity starts within the next 14 days — prepare now.", False, 10, T["WHITE"], "222222"),
        ("", False, 9, T["WHITE"], "AAAAAA"),
        ("USING THIS TEMPLATE ON A FUTURE JOB", True, 11, T["DARK_GREY"], T["WHITE"]),
        ("This template is job-agnostic. For a new project:", False, 10, T["WHITE"], "222222"),
        ("  •  Copy the config file (configs/boiler_house.py) and rename it.", False, 10, T["WHITE"], "222222"),
        ("  •  Edit project name, calendar path, section/zone colours, and PROGRAMME list.", False, 10, T["WHITE"], "222222"),
        ("  •  Run:  python build.py configs/your_project.py", False, 10, T["WHITE"], "222222"),
        ("  •  A fresh, correctly populated lookahead is generated in seconds.", False, 10, T["WHITE"], "222222"),
        ("", False, 9, T["WHITE"], "AAAAAA"),
        (f"Project: {cfg.PROJECT_NAME}  |  {cfg.PROJECT_SUBTITLE}  |  Built: {today_str}", False, 8, T["WHITE"], "AAAAAA"),
    ]
    for i, (text, hdr, sz, bg, fg) in enumerate(rows, 1):
        ws.row_dimensions[i].height = 22 if hdr else 17
        c = ws[f"A{i}"]
        c.value = text
        c.font = Font(name="Arial", size=sz, bold=hdr, color=fg)
        c.fill = _fill(bg)
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.sheet_view.showGridLines = False


# ─── Main entry point ─────────────────────────────────────────────────────────
def build(cfg, output_dir=".", quiet=False):
    """
    Build the lookahead workbook from a config module.

    Args:
        cfg:        Config module (e.g. imported configs.boiler_house)
        output_dir: Directory to write output files into
        quiet:      Suppress print output

    Returns:
        (xlsx_path, bas_path)  — absolute paths of the two output files
    """
    import os
    os.makedirs(output_dir, exist_ok=True)

    cal_sheet = getattr(cfg, "CALENDAR_SHEET", "Table")
    working, non_working = _load_calendar(cfg.CALENDAR_PATH, cal_sheet)
    default_start = _next_monday()
    display_days  = _display_days(default_start, working, non_working)

    wb = openpyxl.Workbook()

    # Sheet 1: 2-Week Lookahead (active sheet)
    ws1 = wb.active
    ws1.title = "2-Week Lookahead"
    _build_lookahead_sheet(ws1, cfg, working, non_working, default_start, display_days)

    # Sheet 2: Programme Targets
    ws2 = wb.create_sheet("Programme Targets")
    _build_programme_targets_sheet(ws2, cfg)

    # Sheet 3: WD Calendar
    ws3 = wb.create_sheet("WD Calendar")
    _build_calendar_sheet(ws3, working, non_working)

    # Sheet 4: How To Use
    ws4 = wb.create_sheet("How To Use")
    _build_guide_sheet(ws4, cfg)

    xlsx_path = os.path.join(output_dir, cfg.OUTPUT_FILENAME)
    bas_path  = os.path.join(output_dir, cfg.VBA_FILENAME)

    wb.save(xlsx_path)
    if not quiet:
        print(f"✓ Workbook saved: {xlsx_path}")

    with open(bas_path, "w") as f:
        f.write(_vba_code(cfg.PROJECT_NAME))
    if not quiet:
        print(f"✓ VBA module:     {bas_path}")

    return xlsx_path, bas_path

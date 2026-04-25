#!/usr/bin/env python3
"""
generate_kpi_workbook.py
========================
Python / openpyxl alternative to the VBA macro.
Generates the complete Warehouse / Dispatch KPI workbook.

Requirements:
    pip install openpyxl

Usage:
    python generate_kpi_workbook.py
    # -> creates  KPI_Workbook.xlsx  in the current directory

Compatibility: Excel 2016 and later.
Unprotect password: KPI2024
"""

from __future__ import annotations

import datetime
from typing import List, Optional

from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import (
    ColorScaleRule, DataBarRule, FormulaRule, Rule
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------------------
# Colour helpers
# ---------------------------------------------------------------------------
GREEN_HEX  = "FF50C878"
AMBER_HEX  = "FFFFA500"
RED_HEX    = "FFFF0000"
HEADER_HEX = "FF191970"   # midnight blue
WHITE_HEX  = "FFFFFFFF"
LGREY_HEX  = "FFF2F2F2"
DARK_HEX   = "FF1E1E1E"
BLUE_HEX   = "FF0070C0"

UNPROTECT_PW = "KPI2024"
# NOTE: This password is intentionally visible in the source code.
# It is a shared workbook protection password documented for all operators
# (not a user authentication credential). Change it in CONFIG if required.


def _hex_font_color(hex_color: str) -> str:
    """Strip the two-character alpha prefix from an ARGB hex string for Font.color."""
    return hex_color[2:]


def fill(hex_color: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=hex_color)


def bold_font(color: str = "FF000000", size: int = 11) -> Font:
    return Font(bold=True, color=color, size=size)


def header_style(ws, row: int, cols: List[str], start_col: int = 1) -> None:
    """Write and style a header row."""
    for c_idx, name in enumerate(cols, start=start_col):
        cell = ws.cell(row=row, column=c_idx, value=name)
        cell.font = bold_font(WHITE_HEX)
        cell.fill = fill(HEADER_HEX)
        cell.alignment = Alignment(horizontal="center", vertical="center")


def thin_border() -> Border:
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)


def add_table(ws, ref: str, name: str, style: str = "TableStyleMedium2",
              show_filter: bool = True) -> Table:
    tbl = Table(displayName=name, ref=ref)
    tbl.tableStyleInfo = TableStyleInfo(
        name=style, showFirstColumn=False, showLastColumn=False,
        showRowStripes=True, showColumnStripes=False)
    ws.add_table(tbl)
    return tbl


def col_letter(n: int) -> str:
    return get_column_letter(n)


def add_shift_validation(ws, col_letter_str: str, first_row: int = 2) -> None:
    dv = DataValidation(
        type="list", formula1='"Day,Night"',
        allow_blank=True, showDropDown=False,
        showInputMessage=True, promptTitle="Shift",
        prompt="Select Day or Night",
        showErrorMessage=True, errorTitle="Invalid",
        error="Please select Day or Night"
    )
    dv.sqref = f"{col_letter_str}{first_row}:{col_letter_str}10000"
    ws.add_data_validation(dv)


def freeze(ws, cell: str = "A2") -> None:
    ws.freeze_panes = cell


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def build_config(wb: Workbook) -> None:
    ws = wb.create_sheet("CONFIG")
    ws.sheet_properties.tabColor = "808080"

    ws["A1"] = "CONFIG - DO NOT DELETE OR RENAME THIS SHEET"
    ws["A1"].font = Font(bold=True, color=_hex_font_color(RED_HEX), size=12)

    # Shift table
    ws["A3"] = "SHIFTS"
    ws["A3"].font = Font(bold=True)
    shift_hdrs = ["ShiftName", "StartTime", "EndTime", "DaysPerWeek"]
    header_style(ws, 4, shift_hdrs)

    shift_data = [
        ("Day",   "08:00", "16:00", 6),
        ("Night", "16:00", "00:00", 6),
    ]
    for r_off, row in enumerate(shift_data, start=5):
        for c_off, val in enumerate(row, start=1):
            ws.cell(row=r_off, column=c_off, value=val)

    add_table(ws, "A4:D6", "tblConfig_Shifts", "TableStyleMedium2")

    # KPI rules
    ws["A9"] = "KPI RULES"
    ws["A9"].font = Font(bold=True)
    rule_hdrs = ["RuleName", "Value", "Description"]
    header_style(ws, 10, rule_hdrs)

    rules = [
        ("HRP_MaxDaysToShow",  1,    "Max days since available for city"),
        ("Packed_MaxAgeDays",  2,    "Max age (days) before packed item is overdue"),
        ("Audit_SampleSize",   20,   "Required number of audits per shift"),
        ("Amber_Threshold",    0.9,  "Performance % for Amber RAG"),
        ("Green_Threshold",    1,    "Performance % for Green RAG"),
    ]
    for r_off, row in enumerate(rules, start=11):
        for c_off, val in enumerate(row, start=1):
            ws.cell(row=r_off, column=c_off, value=val)

    add_table(ws, "A10:C15", "tblConfig_Rules", "TableStyleMedium2")

    # Area list
    ws["E3"] = "AREA LIST"
    ws["E3"].font = Font(bold=True)
    header_style(ws, 4, ["AreaName"], start_col=5)
    areas = ["Auditing", "Manual handover", "Dispatch sealing"]
    for r_off, area in enumerate(areas, start=5):
        ws.cell(row=r_off, column=5, value=area)

    add_table(ws, "E4:E7", "tblConfig_Areas", "TableStyleMedium2")

    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].auto_size = True


def build_in_packed(wb: Workbook) -> None:
    ws = wb.create_sheet("IN_PACKED")
    ws.sheet_properties.tabColor = "0070C0"

    src_cols = [
        "LPN", "PALLET", "STORE", "DIVISION", "FACILITY_NAME",
        "STORE_STATUS", "PICK_LOC", "UNITS", "LAST_UPDATE",
        "DAYS_SINCE_CLOSE", "LOCKS", "TOT_COST", "TOT_RETAIL", "LAST_PACKED"
    ]
    calc_cols = ["AgeDays", "IsShipped", "PackedStatus", "ActionFlag"]
    all_cols = src_cols + calc_cols

    header_style(ws, 1, all_cols)

    # Sample data row
    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)
    sample = [
        "LPN001", "PLT001", "STORE01", "DIV01", "FACILITY A",
        "OPEN", "LOC001", 10, datetime.datetime.now(),
        1, 0, 100.0, 150.0, datetime.datetime.combine(yesterday, datetime.time(10, 0))
    ]
    n_src = len(src_cols)
    for c_off, val in enumerate(sample, start=1):
        ws.cell(row=2, column=c_off, value=val)

    # Calculated columns (stored as formulas in row 2; table will propagate)
    calc_start = n_src + 1   # column 15
    ws.cell(row=2, column=calc_start).value     = "=TODAY()-INT([@LAST_PACKED])"
    ws.cell(row=2, column=calc_start + 1).value = "=COUNTIF(tblShipped[LPN],[@LPN])>0"
    ws.cell(row=2, column=calc_start + 2).value = \
        '=IF([@IsShipped],"Shipped",IF([@AgeDays]>2,"Overdue","Pending"))'
    ws.cell(row=2, column=calc_start + 3).value = "=AND(NOT([@IsShipped]),[@AgeDays]>2)"

    last_col = col_letter(len(all_cols))
    add_table(ws, f"A1:{last_col}2", "tblPacked", "TableStyleMedium6")

    # Conditional formatting: PackedStatus column (col calc_start+2)
    ps_col = col_letter(calc_start + 2)
    ws.conditional_formatting.add(
        f"{ps_col}2:{ps_col}10000",
        Rule(type="containsText", operator="containsText", text="Overdue",
             dxf=DifferentialStyle(fill=fill(RED_HEX), font=Font(color=_hex_font_color(WHITE_HEX))))
    )
    ws.conditional_formatting.add(
        f"{ps_col}2:{ps_col}10000",
        Rule(type="containsText", operator="containsText", text="Shipped",
             dxf=DifferentialStyle(fill=fill(GREEN_HEX)))
    )
    ws.conditional_formatting.add(
        f"{ps_col}2:{ps_col}10000",
        Rule(type="containsText", operator="containsText", text="Pending",
             dxf=DifferentialStyle(fill=fill(AMBER_HEX)))
    )

    freeze(ws)
    _autofit(ws)


def build_in_hrp(wb: Workbook) -> None:
    ws = wb.create_sheet("IN_HRP")
    ws.sheet_properties.tabColor = "00B0F0"

    src_cols = [
        "CATEGORY", "DAYS_SINCE_AVAILABLE_FOR_CITY", "DATE_SCANNED_TO_CITY",
        "OLPN", "XREF_OLPN", "STORE_NUMBER", "STORE_NAME",
        "CARTON_LOCKS", "LPN_STATUS", "STORE_ACK_STATUS",
        "UNITS", "COST_VALUE", "RETAIL_VALUE"
    ]
    calc_cols = ["CartonID", "IsAcknowledged", "IncludeInHRP", "AgeBucket"]
    all_cols = src_cols + calc_cols

    header_style(ws, 1, all_cols)

    today = datetime.date.today()
    sample_src = [
        "HRP", 0, datetime.datetime.now(),
        "OLPN001", "", "STORE01", "Store One",
        0, "AVAILABLE", "N", 5, 50.0, 75.0
    ]
    for c_off, val in enumerate(sample_src, start=1):
        ws.cell(row=2, column=c_off, value=val)

    calc_start = len(src_cols) + 1
    ws.cell(row=2, column=calc_start).value     = '=IF([@OLPN]<>"",[@OLPN],[@XREF_OLPN])'
    ws.cell(row=2, column=calc_start + 1).value = '=[@STORE_ACK_STATUS]="Y"'
    ws.cell(row=2, column=calc_start + 2).value = \
        '=AND([@STORE_ACK_STATUS]="N",[@DAYS_SINCE_AVAILABLE_FOR_CITY]<=1)'
    ws.cell(row=2, column=calc_start + 3).value = \
        '=IF([@DAYS_SINCE_AVAILABLE_FOR_CITY]=0,"<24h","24h+")'

    last_col = col_letter(len(all_cols))
    add_table(ws, f"A1:{last_col}2", "tblHRP", "TableStyleMedium7")

    # Data validation on STORE_ACK_STATUS
    ack_col = col_letter(src_cols.index("STORE_ACK_STATUS") + 1)
    dv_ack = DataValidation(
        type="list", formula1='"Y,N"', allow_blank=True,
        showDropDown=False, showInputMessage=True,
        promptTitle="Ack Status", prompt="Enter Y or N"
    )
    dv_ack.sqref = f"{ack_col}2:{ack_col}10000"
    ws.add_data_validation(dv_ack)

    # CF: IncludeInHRP = TRUE  (column calc_start+2)
    inc_col = col_letter(calc_start + 2)
    ws.conditional_formatting.add(
        f"{inc_col}2:{inc_col}10000",
        Rule(type="cellIs", operator="equal", formula=['"TRUE"'],
             dxf=DifferentialStyle(fill=fill(RED_HEX), font=Font(color=_hex_font_color(WHITE_HEX))))
    )

    freeze(ws)
    _autofit(ws)


def build_in_shipped(wb: Workbook) -> None:
    ws = wb.create_sheet("IN_SHIPPED_LPNS")
    ws.sheet_properties.tabColor = "00B050"

    cols = ["BusinessDate", "ShiftName", "LPN", "EnteredBy", "Notes", "DupFlag"]
    header_style(ws, 1, cols)

    ws.cell(row=2, column=1, value=datetime.date.today())
    ws.cell(row=2, column=2, value="Day")
    ws.cell(row=2, column=3, value="LPN001")
    ws.cell(row=2, column=4, value="User1")
    ws.cell(row=2, column=5, value="")
    ws.cell(row=2, column=6).value = \
        "=COUNTIFS([BusinessDate],[@BusinessDate],[ShiftName],[@ShiftName],[LPN],[@LPN])>1"

    add_table(ws, "A1:F2", "tblShipped", "TableStyleMedium9")

    # CF: DupFlag = TRUE
    ws.conditional_formatting.add(
        "F2:F10000",
        Rule(type="cellIs", operator="equal", formula=['"TRUE"'],
             dxf=DifferentialStyle(fill=fill(AMBER_HEX)))
    )

    add_shift_validation(ws, "B")
    freeze(ws)
    _autofit(ws)


def build_in_staffing(wb: Workbook) -> None:
    ws = wb.create_sheet("IN_STAFFING")
    ws.sheet_properties.tabColor = "00B050"

    cols = ["BusinessDate", "ShiftName", "Area", "StaffAvailable"]
    header_style(ws, 1, cols)

    ws.cell(row=2, column=1, value=datetime.date.today())
    ws.cell(row=2, column=2, value="Day")
    ws.cell(row=2, column=3, value="Dispatch sealing")
    ws.cell(row=2, column=4, value=5)

    add_table(ws, "A1:D2", "tblStaffing", "TableStyleMedium4")

    add_shift_validation(ws, "B")

    dv_area = DataValidation(
        type="list",
        formula1='"Auditing,Manual handover,Dispatch sealing"',
        allow_blank=True, showDropDown=False
    )
    dv_area.sqref = "C2:C10000"
    ws.add_data_validation(dv_area)

    freeze(ws)
    _autofit(ws)


def build_in_targets(wb: Workbook) -> None:
    ws = wb.create_sheet("IN_TARGETS_DAILY")
    ws.sheet_properties.tabColor = "00B050"

    cols = ["BusinessDate", "ShiftName", "TargetPerPersonPerShift"]
    header_style(ws, 1, cols)

    ws.cell(row=2, column=1, value=datetime.date.today())
    ws.cell(row=2, column=2, value="Day")
    ws.cell(row=2, column=3, value=50)

    add_table(ws, "A1:C2", "tblTargetsDaily", "TableStyleMedium4")
    add_shift_validation(ws, "B")
    freeze(ws)
    _autofit(ws)


def build_t_dispatch_kpi(wb: Workbook) -> None:
    ws = wb.create_sheet("T_DISPATCH_KPI")
    ws.sheet_properties.tabColor = "FFC000"

    cols = [
        "BusinessDate", "ShiftName", "ShippedCartons", "TotalStaff",
        "TargetPerPerson", "ExpectedCartons", "PerformancePct", "RAG"
    ]
    header_style(ws, 1, cols)

    ws.cell(row=2, column=1, value=datetime.date.today())
    ws.cell(row=2, column=2, value="Day")
    ws.cell(row=2, column=3).value = \
        "=COUNTIFS(tblShipped[BusinessDate],[@BusinessDate],tblShipped[ShiftName],[@ShiftName])"
    ws.cell(row=2, column=4).value = \
        "=SUMIFS(tblStaffing[StaffAvailable],tblStaffing[BusinessDate],[@BusinessDate],tblStaffing[ShiftName],[@ShiftName])"
    # Array formula (INDEX/MATCH across two criteria)
    ws.cell(row=2, column=5).value = \
        ("=IFERROR(INDEX(tblTargetsDaily[TargetPerPersonPerShift],"
         "MATCH(1,(tblTargetsDaily[BusinessDate]=[@BusinessDate])"
         "*(tblTargetsDaily[ShiftName]=[@ShiftName]),0)),0)")
    ws.cell(row=2, column=6).value = "=[@TotalStaff]*[@TargetPerPerson]"
    ws.cell(row=2, column=7).value = "=IFERROR([@ShippedCartons]/[@ExpectedCartons],0)"
    ws.cell(row=2, column=8).value = \
        '=IF([@PerformancePct]>=1,"Green",IF([@PerformancePct]>=0.9,"Amber","Red"))'

    ws.cell(row=2, column=7).number_format = "0.0%"

    add_table(ws, "A1:H2", "tblDispatchKPI", "TableStyleMedium2")

    # RAG conditional formatting
    rag_col = "H"
    for text, hex_color in [("Green", GREEN_HEX), ("Amber", AMBER_HEX), ("Red", RED_HEX)]:
        ws.conditional_formatting.add(
            f"{rag_col}2:{rag_col}10000",
            Rule(type="containsText", operator="containsText", text=text,
                 dxf=DifferentialStyle(fill=fill(hex_color)))
        )

    freeze(ws)
    _autofit(ws)


def build_t_dispatch_daily(wb: Workbook) -> None:
    ws = wb.create_sheet("T_DISPATCH_DAILY")
    ws.sheet_properties.tabColor = "FFC000"

    cols = ["BusinessDate", "TotalShipped", "TotalStaff", "TotalExpected", "DailyPerformancePct"]
    header_style(ws, 1, cols)

    ws.cell(row=2, column=1, value=datetime.date.today())
    ws.cell(row=2, column=2).value = \
        "=SUMIF(tblDispatchKPI[BusinessDate],[@BusinessDate],tblDispatchKPI[ShippedCartons])"
    ws.cell(row=2, column=3).value = \
        "=SUMIF(tblDispatchKPI[BusinessDate],[@BusinessDate],tblDispatchKPI[TotalStaff])"
    ws.cell(row=2, column=4).value = \
        "=SUMIF(tblDispatchKPI[BusinessDate],[@BusinessDate],tblDispatchKPI[ExpectedCartons])"
    ws.cell(row=2, column=5).value = "=IFERROR([@TotalShipped]/[@TotalExpected],0)"
    ws.cell(row=2, column=5).number_format = "0.0%"

    add_table(ws, "A1:E2", "tblDispatchDaily", "TableStyleMedium2")

    freeze(ws)
    _autofit(ws)


def build_action_hrp(wb: Workbook) -> None:
    ws = wb.create_sheet("ACTION_HRP")
    ws.sheet_properties.tabColor = "C00000"

    cols = [
        "CATEGORY", "CartonID", "STORE_NUMBER", "STORE_NAME",
        "DAYS_SINCE_AVAILABLE_FOR_CITY", "STORE_ACK_STATUS",
        "UNITS", "COST_VALUE",
        "Owner", "ContactedCity", "ContactTime", "NextStep"
    ]
    for c_idx, name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c_idx, value=name)
        cell.font = bold_font(WHITE_HEX)
        cell.fill = fill("FFC00000")
        cell.alignment = Alignment(horizontal="center")

    ws.cell(row=2, column=1, value="-- Populate via 'POPULATE ACTION SHEETS' button on DASHBOARD --")
    ws.cell(row=2, column=1).font = Font(italic=True, color="808080")

    # Validation on ContactedCity
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    dv.sqref = "J2:J10000"
    ws.add_data_validation(dv)

    freeze(ws)
    _autofit(ws)


def build_action_packed(wb: Workbook) -> None:
    ws = wb.create_sheet("ACTION_PACKED")
    ws.sheet_properties.tabColor = "C00000"

    cols = [
        "LPN", "STORE", "DIVISION", "AgeDays",
        "PackedStatus", "UNITS",
        "HoldReason", "Owner", "PlannedShipTime"
    ]
    for c_idx, name in enumerate(cols, start=1):
        cell = ws.cell(row=1, column=c_idx, value=name)
        cell.font = bold_font(WHITE_HEX)
        cell.fill = fill("FFC00000")
        cell.alignment = Alignment(horizontal="center")

    ws.cell(row=2, column=1, value="-- Populate via 'POPULATE ACTION SHEETS' button on DASHBOARD --")
    ws.cell(row=2, column=1).font = Font(italic=True, color="808080")

    # CF on PackedStatus column (col E)
    ws.conditional_formatting.add(
        "E2:E10000",
        Rule(type="containsText", operator="containsText", text="Overdue",
             dxf=DifferentialStyle(fill=fill(RED_HEX), font=Font(color=_hex_font_color(WHITE_HEX))))
    )
    ws.conditional_formatting.add(
        "E2:E10000",
        Rule(type="containsText", operator="containsText", text="Pending",
             dxf=DifferentialStyle(fill=fill(AMBER_HEX)))
    )

    freeze(ws)
    _autofit(ws)


def build_history(wb: Workbook) -> None:
    ws = wb.create_sheet("HISTORY")
    ws.sheet_properties.tabColor = "808080"

    cols = [
        "SnapshotTimestamp", "BusinessDate", "ShiftName",
        "ShippedCartons", "TotalStaff", "ExpectedCartons",
        "PerformancePct", "RAG",
        "HRP_OpenCount", "Packed_OverdueCount"
    ]
    header_style(ws, 1, cols)
    # Seed an empty row so the table range has at least 2 rows (openpyxl requirement)
    for c_idx in range(1, len(cols) + 1):
        ws.cell(row=2, column=c_idx, value="")
    add_table(ws, f"A1:{col_letter(len(cols))}2", "tblHistory", "TableStyleMedium2")
    _autofit(ws)


def build_data_quality(wb: Workbook) -> None:
    ws = wb.create_sheet("DATA_QUALITY")
    ws.sheet_properties.tabColor = "FFC000"

    ws["A1"] = "DATA QUALITY CHECKS"
    ws["A1"].font = Font(bold=True, size=14, color=_hex_font_color(HEADER_HEX))

    hdrs = ["Check", "Formula / Logic", "Result", "Status"]
    header_style(ws, 3, hdrs)

    checks = [
        ("Missing CartonID in HRP",
         '=COUNTIFS(tblHRP[CartonID],"")'),
        ("Future dates in IN_PACKED (LAST_PACKED)",
         '=COUNTIF(tblPacked[LAST_PACKED],">"&TODAY())'),
        ("Future dates in IN_SHIPPED_LPNS (BusinessDate)",
         '=COUNTIF(tblShipped[BusinessDate],">"&TODAY())'),
        ("Duplicate LPNs in Shipped list",
         "=COUNTIF(tblShipped[DupFlag],TRUE)"),
        ("StaffAvailable = 0 rows in Staffing",
         "=COUNTIF(tblStaffing[StaffAvailable],0)"),
        ('Invalid ShiftName in Shipped (not Day/Night)',
         '=SUMPRODUCT((tblShipped[ShiftName]<>"Day")*(tblShipped[ShiftName]<>"Night")*(tblShipped[ShiftName]<>""))'),
        ("Invalid ShiftName in Staffing",
         '=SUMPRODUCT((tblStaffing[ShiftName]<>"Day")*(tblStaffing[ShiftName]<>"Night")*(tblStaffing[ShiftName]<>""))'),
        ("HRP rows with IncludeInHRP=TRUE (open action items)",
         "=COUNTIF(tblHRP[IncludeInHRP],TRUE)"),
        ("Packed items where ActionFlag=TRUE (overdue)",
         "=COUNTIF(tblPacked[ActionFlag],TRUE)"),
    ]

    for r_off, (label, formula) in enumerate(checks, start=4):
        ws.cell(row=r_off, column=1, value=label)
        ws.cell(row=r_off, column=2, value=f"'{formula}")   # leading apostrophe forces Excel to treat this as display text, not an evaluated formula
        ws.cell(row=r_off, column=3).value = formula         # live formula
        ws.cell(row=r_off, column=4).value = \
            f'=IF(C{r_off}=0,"OK","WARNING")'

        for c in range(1, 5):
            ws.cell(row=r_off, column=c).border = thin_border()

        # CF: WARNING = amber
        ws.conditional_formatting.add(
            f"D{r_off}:D{r_off}",
            Rule(type="containsText", operator="containsText", text="WARNING",
                 dxf=DifferentialStyle(fill=fill(AMBER_HEX), font=Font(bold=True)))
        )
        ws.conditional_formatting.add(
            f"D{r_off}:D{r_off}",
            Rule(type="containsText", operator="containsText", text="OK",
                 dxf=DifferentialStyle(fill=fill(GREEN_HEX)))
        )

    _autofit(ws)


def build_dashboard(wb: Workbook) -> None:
    ws = wb.create_sheet("DASHBOARD")
    ws.sheet_properties.tabColor = "0070C0"

    # Dark background
    for row in ws.iter_rows(min_row=1, max_row=60, min_col=1, max_col=14):
        for cell in row:
            cell.fill = fill(DARK_HEX)
            cell.font = Font(color=_hex_font_color(WHITE_HEX))

    # Title
    ws.merge_cells("A1:N1")
    title_cell = ws["A1"]
    title_cell.value = "WAREHOUSE / DISPATCH KPI DASHBOARD"
    title_cell.font = Font(bold=True, size=18, color=_hex_font_color(WHITE_HEX))
    title_cell.fill = fill(HEADER_HEX)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    ws["A2"] = "Last Refreshed:"
    ws["B2"].value = f'=TEXT(NOW(),"dd/mm/yyyy hh:mm")'
    ws["B2"].font = Font(bold=True, color=_hex_font_color(WHITE_HEX))

    # KPI Cards
    _kpi_card(ws, row=4, col=2, title="HRP OPEN ITEMS",
              formula="=COUNTIF(tblHRP[IncludeInHRP],TRUE)", bg=RED_HEX)
    _kpi_card(ws, row=4, col=6, title="PACKED OVERDUE",
              formula="=COUNTIF(tblPacked[ActionFlag],TRUE)", bg=RED_HEX)
    _kpi_card(ws, row=4, col=10, title="DISPATCH PERF TODAY",
              formula='=IFERROR(TEXT(SUMIF(tblDispatchDaily[BusinessDate],TODAY(),tblDispatchDaily[DailyPerformancePct]),"0.0%"),"N/A")',
              bg=BLUE_HEX)
    _kpi_card(ws, row=9, col=2, title="DUPLICATE LPNs",
              formula="=COUNTIF(tblShipped[DupFlag],TRUE)", bg=AMBER_HEX)
    _kpi_card(ws, row=9, col=6, title="STAFF TODAY",
              formula="=SUMIF(tblDispatchKPI[BusinessDate],TODAY(),tblDispatchKPI[TotalStaff])",
              bg=BLUE_HEX)
    _kpi_card(ws, row=9, col=10, title="CARTONS SHIPPED TODAY",
              formula="=SUMIF(tblDispatchKPI[BusinessDate],TODAY(),tblDispatchKPI[ShippedCartons])",
              bg=BLUE_HEX)

    # Instructions
    ws["A14"] = "CHARTS AREA"
    ws["A14"].font = Font(bold=True, size=12, color=_hex_font_color(WHITE_HEX))
    ws["A14"].fill = fill(DARK_HEX)

    ws["A15"] = ("After adding data to the input sheets, use the 'Refresh All' button "
                 "to update PivotTables and charts.")
    ws["A15"].font = Font(italic=True, color="C8C8C8")
    ws["A15"].fill = fill(DARK_HEX)

    ws["A17"] = "FILTER GUIDE:"
    ws["A17"].font = Font(bold=True, color=_hex_font_color(WHITE_HEX))
    ws["A17"].fill = fill(DARK_HEX)
    ws["A18"] = "Use Excel AutoFilter on T_DISPATCH_KPI for ShiftName / BusinessDate / RAG filters."
    ws["A18"].fill = fill(DARK_HEX)
    ws["A19"] = "For pivot-based slicers: Insert PivotTable from tblDispatchKPI, then Insert > Slicer."
    ws["A19"].fill = fill(DARK_HEX)

    # Instruction note about buttons
    ws["A21"] = "BUTTONS (add manually or run the VBA RefreshAll / TakeDailySnapshot macros):"
    ws["A21"].font = Font(bold=True, color=_hex_font_color(WHITE_HEX))
    ws["A21"].fill = fill(DARK_HEX)

    button_notes = [
        "REFRESH ALL             -> Recalculates all formulas & refreshes PivotTables",
        "TAKE DAILY SNAPSHOT     -> Appends current KPIs to the HISTORY sheet",
        "POPULATE ACTION SHEETS  -> Copies filtered data to ACTION_HRP and ACTION_PACKED",
    ]
    for r_off, note in enumerate(button_notes, start=22):
        cell = ws.cell(row=r_off, column=1, value=note)
        cell.font = Font(italic=True, color="C8C8C8")
        cell.fill = fill(DARK_HEX)

    for col_idx in range(1, 15):
        ws.column_dimensions[col_letter(col_idx)].width = 16


def _kpi_card(ws, row: int, col: int, title: str, formula: str, bg: str) -> None:
    """Write a 3-wide × 3-tall KPI card (title above, value below)."""
    # Title row (row - 1)
    title_row = row - 1
    title_ref = f"{col_letter(col)}{title_row}:{col_letter(col + 2)}{title_row}"
    ws.merge_cells(title_ref)
    tc = ws.cell(row=title_row, column=col, value=title)
    tc.font = Font(bold=True, size=9, color="C8C8C8")
    tc.fill = fill("FF323232")
    tc.alignment = Alignment(horizontal="center")

    # Value cells
    val_ref = f"{col_letter(col)}{row}:{col_letter(col + 2)}{row + 2}"
    ws.merge_cells(val_ref)
    vc = ws.cell(row=row, column=col)
    vc.value = formula
    vc.font = Font(bold=True, size=22, color=_hex_font_color(WHITE_HEX))
    vc.fill = fill(bg)
    vc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 40


def _autofit(ws) -> None:
    """Approximate auto-fit (openpyxl does not have native auto-fit)."""
    for col in ws.columns:
        max_len = 0
        col_letter_str = col[0].column_letter
        for cell in col:
            try:
                cell_len = len(str(cell.value)) if cell.value else 0
                if cell_len > max_len:
                    max_len = cell_len
            except Exception:
                pass
        ws.column_dimensions[col_letter_str].width = min(max(max_len + 2, 10), 40)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def build_workbook(output_path: str = "KPI_Workbook.xlsx") -> None:
    wb = Workbook()
    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    # Build all sheets
    build_config(wb)
    build_in_packed(wb)
    build_in_hrp(wb)
    build_in_shipped(wb)
    build_in_staffing(wb)
    build_in_targets(wb)
    build_t_dispatch_kpi(wb)
    build_t_dispatch_daily(wb)
    build_action_hrp(wb)
    build_action_packed(wb)
    build_history(wb)
    build_data_quality(wb)
    build_dashboard(wb)

    # Re-order sheets for usability.
    # Iterates desired order; skips any sheet not present to avoid KeyError.
    desired_order = [
        "DASHBOARD", "CONFIG",
        "IN_PACKED", "IN_HRP", "IN_SHIPPED_LPNS", "IN_STAFFING", "IN_TARGETS_DAILY",
        "T_DISPATCH_KPI", "T_DISPATCH_DAILY",
        "ACTION_HRP", "ACTION_PACKED",
        "HISTORY", "DATA_QUALITY"
    ]
    missing = [n for n in desired_order if n not in wb.sheetnames]
    if missing:
        print(f"Warning: the following sheets are missing and will be skipped in reordering: {missing}")
    for idx, name in enumerate(desired_order):
        if name not in wb.sheetnames:
            continue
        current_pos = wb.sheetnames.index(name)
        if current_pos != idx:
            wb.move_sheet(name, offset=current_pos - idx)

    wb.save(output_path)
    print(f"Workbook saved: {output_path}")
    print(f"Unprotect password: {UNPROTECT_PW}")
    print("\nSheets created:")
    for name in wb.sheetnames:
        print(f"  - {name}")


if __name__ == "__main__":
    build_workbook()

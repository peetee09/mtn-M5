# Manual Setup Instructions — Warehouse / Dispatch KPI Workbook

> **Use these instructions only if you cannot run macros (VBA) or Python scripts.**  
> Unprotect password: **KPI2024**

---

## Overview

You will create **13 worksheets** in a single Excel workbook:

| # | Sheet name | Purpose |
|---|-----------|---------|
| 1 | DASHBOARD | KPI cards, charts, control buttons |
| 2 | CONFIG | Shift times, KPI rules, area list |
| 3 | IN_PACKED | Paste packed-carton export data |
| 4 | IN_HRP | Paste HRP export data |
| 5 | IN_SHIPPED_LPNS | Manual shipped-LPN log |
| 6 | IN_STAFFING | Manual staffing log |
| 7 | IN_TARGETS_DAILY | Manual per-person targets |
| 8 | T_DISPATCH_KPI | Auto-calculated shift KPIs |
| 9 | T_DISPATCH_DAILY | Auto-calculated daily KPIs |
| 10 | ACTION_HRP | Filtered HRP action list |
| 11 | ACTION_PACKED | Filtered packed action list |
| 12 | HISTORY | Snapshot log |
| 13 | DATA_QUALITY | Data quality checks |

---

## STEP 1 — Create the workbook

1. Open Excel, create a **New Blank Workbook**.
2. Save it as **KPI_Workbook.xlsx** (or `.xlsm` if you plan to add macros later).

---

## STEP 2 — Create the CONFIG sheet

1. Rename **Sheet1** to `CONFIG`.
2. In **A1** type: `CONFIG - DO NOT DELETE OR RENAME THIS SHEET` (bold, red font).

### Shift table (A4:D6)

| A | B | C | D |
|---|---|---|---|
| **ShiftName** | **StartTime** | **EndTime** | **DaysPerWeek** |
| Day | 08:00 | 16:00 | 6 |
| Night | 16:00 | 00:00 | 6 |

3. Select **A4:D6** → Insert > Table (check "My table has headers") → name the table `tblConfig_Shifts`.
4. Format B5:C6 as **Time (hh:mm)**.

### KPI Rules table (A10:C15)

| A | B | C |
|---|---|---|
| **RuleName** | **Value** | **Description** |
| HRP_MaxDaysToShow | 1 | Max days since available for city |
| Packed_MaxAgeDays | 2 | Max age (days) before packed item is overdue |
| Audit_SampleSize | 20 | Required number of audits per shift |
| Amber_Threshold | 0.9 | Performance % for Amber RAG |
| Green_Threshold | 1 | Performance % for Green RAG |

5. Select **A10:C15** → Insert > Table → name it `tblConfig_Rules`.

### Area list (E4:E7)

| E |
|---|
| **AreaName** |
| Auditing |
| Manual handover |
| Dispatch sealing |

6. Select **E4:E7** → Insert > Table → name it `tblConfig_Areas`.

---

## STEP 3 — Create IN_PACKED sheet

1. Add a new sheet, name it `IN_PACKED`.
2. Enter the following headers in **row 1** (bold, dark blue fill, white font):

```
A: LPN  B: PALLET  C: STORE  D: DIVISION  E: FACILITY_NAME  F: STORE_STATUS
G: PICK_LOC  H: UNITS  I: LAST_UPDATE  J: DAYS_SINCE_CLOSE  K: LOCKS
L: TOT_COST  M: TOT_RETAIL  N: LAST_PACKED
O: AgeDays  P: IsShipped  Q: PackedStatus  R: ActionFlag
```

3. In **O2** enter: `=TODAY()-INT([@LAST_PACKED])`
4. In **P2** enter: `=COUNTIF(tblShipped[LPN],[@LPN])>0`
5. In **Q2** enter: `=IF([@IsShipped],"Shipped",IF([@AgeDays]>2,"Overdue","Pending"))`
6. In **R2** enter: `=AND(NOT([@IsShipped]),[@AgeDays]>2)`
7. Enter **at least one data row** in row 2 (can be dummy data like `LPN001`).
8. Select **A1:R2** → Insert > Table → name the table `tblPacked`.
9. Format column **I** and **N** as `dd/mm/yyyy hh:mm`.
10. **Freeze panes**: click cell A2, View > Freeze Panes > Freeze Panes.

### Conditional formatting for PackedStatus (column Q)

11. Select **Q2:Q10000**.
12. Home > Conditional Formatting > New Rule:
    - **Rule 1**: Cell value = `"Overdue"` → Fill Red, White text.
    - **Rule 2**: Cell value = `"Shipped"` → Fill Green.
    - **Rule 3**: Cell value = `"Pending"` → Fill Amber/Orange.

---

## STEP 4 — Create IN_HRP sheet

1. Add sheet, name it `IN_HRP`.
2. Enter headers in row 1 (bold, dark blue/white):

```
A: CATEGORY  B: DAYS_SINCE_AVAILABLE_FOR_CITY  C: DATE_SCANNED_TO_CITY
D: OLPN  E: XREF_OLPN  F: STORE_NUMBER  G: STORE_NAME
H: CARTON_LOCKS  I: LPN_STATUS  J: STORE_ACK_STATUS
K: UNITS  L: COST_VALUE  M: RETAIL_VALUE
N: CartonID  O: IsAcknowledged  P: IncludeInHRP  Q: AgeBucket
```

3. In **N2**: `=IF([@OLPN]<>"",[@OLPN],[@XREF_OLPN])`
4. In **O2**: `=[@STORE_ACK_STATUS]="Y"`
5. In **P2**: `=AND([@STORE_ACK_STATUS]="N",[@DAYS_SINCE_AVAILABLE_FOR_CITY]<=1)`
6. In **Q2**: `=IF([@DAYS_SINCE_AVAILABLE_FOR_CITY]=0,"<24h","24h+")`
7. Enter a dummy data row, select **A1:Q2** → Insert > Table → name it `tblHRP`.
8. **Data Validation** on column **J** (STORE_ACK_STATUS):
   - Select J2:J10000 → Data > Data Validation → List → `Y,N`.
9. Freeze panes at A2.

---

## STEP 5 — Create IN_SHIPPED_LPNS sheet

1. Add sheet, name it `IN_SHIPPED_LPNS`.
2. Headers in row 1:

```
A: BusinessDate  B: ShiftName  C: LPN  D: EnteredBy  E: Notes  F: DupFlag
```

3. In **F2**: `=COUNTIFS([BusinessDate],[@BusinessDate],[ShiftName],[@ShiftName],[LPN],[@LPN])>1`
4. Enter a dummy row, select **A1:F2** → Insert > Table → name it `tblShipped`.
5. Data Validation on **B** (ShiftName): List → `Day,Night`.
6. Format column A as `dd/mm/yyyy`.
7. Conditional formatting on **F2:F10000**: Cell value = `TRUE` → Amber fill.
8. Freeze panes at A2.

---

## STEP 6 — Create IN_STAFFING sheet

1. Add sheet, name it `IN_STAFFING`.
2. Headers: `A: BusinessDate  B: ShiftName  C: Area  D: StaffAvailable`
3. Dummy row → Select A1:D2 → Insert > Table → name it `tblStaffing`.
4. Data Validation on **B**: `Day,Night`.
5. Data Validation on **C**: `Auditing,Manual handover,Dispatch sealing`.
6. Format A as `dd/mm/yyyy`. Freeze at A2.

---

## STEP 7 — Create IN_TARGETS_DAILY sheet

1. Add sheet, name it `IN_TARGETS_DAILY`.
2. Headers: `A: BusinessDate  B: ShiftName  C: TargetPerPersonPerShift`
3. Dummy row → Table → name it `tblTargetsDaily`.
4. Data Validation on **B**: `Day,Night`.
5. Format A as `dd/mm/yyyy`. Freeze at A2.

---

## STEP 8 — Create T_DISPATCH_KPI sheet

1. Add sheet, name it `T_DISPATCH_KPI`.
2. Headers: `A: BusinessDate  B: ShiftName  C: ShippedCartons  D: TotalStaff  E: TargetPerPerson  F: ExpectedCartons  G: PerformancePct  H: RAG`
3. Formulas for each row (enter with **Ctrl+Shift+Enter** for col E if needed):

   | Column | Formula |
   |--------|---------|
   | C | `=COUNTIFS(tblShipped[BusinessDate],[@BusinessDate],tblShipped[ShiftName],[@ShiftName])` |
   | D | `=SUMIFS(tblStaffing[StaffAvailable],tblStaffing[BusinessDate],[@BusinessDate],tblStaffing[ShiftName],[@ShiftName])` |
   | E | `=IFERROR(INDEX(tblTargetsDaily[TargetPerPersonPerShift],MATCH(1,(tblTargetsDaily[BusinessDate]=[@BusinessDate])*(tblTargetsDaily[ShiftName]=[@ShiftName]),0)),0)` (**array formula — press Ctrl+Shift+Enter in Excel 2016; plain Enter in Excel 365/2021**) |
   | F | `=[@TotalStaff]*[@TargetPerPerson]` |
   | G | `=IFERROR([@ShippedCartons]/[@ExpectedCartons],0)` |
   | H | `=IF([@PerformancePct]>=1,"Green",IF([@PerformancePct]>=0.9,"Amber","Red"))` |

4. Dummy row → Table → name `tblDispatchKPI`.
5. Format G as `0.0%`.
6. Conditional formatting on **H**: `"Green"` → green fill, `"Amber"` → amber fill, `"Red"` → red fill.
7. Freeze at A2.

---

## STEP 9 — Create T_DISPATCH_DAILY sheet

1. Add sheet, name it `T_DISPATCH_DAILY`.
2. Headers: `A: BusinessDate  B: TotalShipped  C: TotalStaff  D: TotalExpected  E: DailyPerformancePct`
3. Formulas:

   | Column | Formula |
   |--------|---------|
   | B | `=SUMIF(tblDispatchKPI[BusinessDate],[@BusinessDate],tblDispatchKPI[ShippedCartons])` |
   | C | `=SUMIF(tblDispatchKPI[BusinessDate],[@BusinessDate],tblDispatchKPI[TotalStaff])` |
   | D | `=SUMIF(tblDispatchKPI[BusinessDate],[@BusinessDate],tblDispatchKPI[ExpectedCartons])` |
   | E | `=IFERROR([@TotalShipped]/[@TotalExpected],0)` |

4. Dummy row → Table → name `tblDispatchDaily`.
5. Format E as `0.0%`. Freeze at A2.

---

## STEP 10 — Create ACTION_HRP sheet

1. Add sheet, name it `ACTION_HRP`.
2. Headers (red fill, white text):

```
A: CATEGORY  B: CartonID  C: STORE_NUMBER  D: STORE_NAME
E: DAYS_SINCE_AVAILABLE_FOR_CITY  F: STORE_ACK_STATUS  G: UNITS  H: COST_VALUE
I: Owner  J: ContactedCity  K: ContactTime  L: NextStep
```

3. Data Validation on **J** (ContactedCity): `Y,N`.
4. Freeze at A2.
5. This sheet is populated **manually** by filtering IN_HRP where IncludeInHRP = TRUE, copying, and pasting values here. *(Or use the VBA PopulateActionSheets macro.)*

---

## STEP 11 — Create ACTION_PACKED sheet

1. Add sheet, name it `ACTION_PACKED`.
2. Headers (red fill, white text):

```
A: LPN  B: STORE  C: DIVISION  D: AgeDays  E: PackedStatus
F: UNITS  G: HoldReason  H: Owner  I: PlannedShipTime
```

3. Freeze at A2.
4. Populate by filtering IN_PACKED where ActionFlag = TRUE, sort by AgeDays descending, copy+paste values here.

---

## STEP 12 — Create HISTORY sheet

1. Add sheet, name it `HISTORY`.
2. Headers:

```
A: SnapshotTimestamp  B: BusinessDate  C: ShiftName  D: ShippedCartons
E: TotalStaff  F: ExpectedCartons  G: PerformancePct  H: RAG
I: HRP_OpenCount  J: Packed_OverdueCount
```

3. Select A1:J1 → Insert > Table (no data rows needed) → name it `tblHistory`.
4. At end of each shift: copy the current T_DISPATCH_KPI row, paste as **values** into the next available row here, add timestamp in column A.

---

## STEP 13 — Create DATA_QUALITY sheet

1. Add sheet, name it `DATA_QUALITY`.
2. In A1: `DATA QUALITY CHECKS` (bold, size 14).
3. In A3:D3 enter headers: `Check | Formula / Logic | Result | Status`.
4. Enter each check row:

| Row | A (Check) | C (Result formula) | D (Status formula) |
|-----|-----------|-------------------|-------------------|
| 4 | Missing CartonID in HRP | `=COUNTIFS(tblHRP[CartonID],"")` | `=IF(C4=0,"OK","WARNING")` |
| 5 | Future dates in LAST_PACKED | `=COUNTIF(tblPacked[LAST_PACKED],">"&TODAY())` | `=IF(C5=0,"OK","WARNING")` |
| 6 | Future dates in BusinessDate | `=COUNTIF(tblShipped[BusinessDate],">"&TODAY())` | `=IF(C6=0,"OK","WARNING")` |
| 7 | Duplicate LPNs in Shipped | `=COUNTIF(tblShipped[DupFlag],TRUE)` | `=IF(C7=0,"OK","WARNING")` |
| 8 | StaffAvailable = 0 | `=COUNTIF(tblStaffing[StaffAvailable],0)` | `=IF(C8=0,"OK","WARNING")` |
| 9 | Invalid ShiftName in Shipped | `=SUMPRODUCT((tblShipped[ShiftName]<>"Day")*(tblShipped[ShiftName]<>"Night")*(tblShipped[ShiftName]<>""))` | `=IF(C9=0,"OK","WARNING")` |
| 10 | Invalid ShiftName in Staffing | `=SUMPRODUCT((tblStaffing[ShiftName]<>"Day")*(tblStaffing[ShiftName]<>"Night")*(tblStaffing[ShiftName]<>""))` | `=IF(C10=0,"OK","WARNING")` |
| 11 | HRP open action items | `=COUNTIF(tblHRP[IncludeInHRP],TRUE)` | `=IF(C11=0,"OK","WARNING")` |
| 12 | Packed overdue items | `=COUNTIF(tblPacked[ActionFlag],TRUE)` | `=IF(C12=0,"OK","WARNING")` |

5. Conditional formatting on D4:D12: `"WARNING"` → amber; `"OK"` → green.

---

## STEP 14 — Create DASHBOARD sheet

1. Add sheet, name it `DASHBOARD`. Move it to first position.
2. Fill all cells with dark background (RGB 30, 30, 30), white font.
3. Merge **A1:N1**, type the title, apply dark blue fill.
4. Create KPI cards using merged cells (example):

   | Card | Cells | Formula |
   |------|-------|---------|
   | HRP OPEN ITEMS | B4:D7 | `=COUNTIF(tblHRP[IncludeInHRP],TRUE)` |
   | PACKED OVERDUE | F4:H7 | `=COUNTIF(tblPacked[ActionFlag],TRUE)` |
   | DISPATCH PERF TODAY | J4:L7 | `=IFERROR(TEXT(SUMIF(tblDispatchDaily[BusinessDate],TODAY(),tblDispatchDaily[DailyPerformancePct]),"0.0%"),"N/A")` |
   | DUPLICATE LPNs | B9:D12 | `=COUNTIF(tblShipped[DupFlag],TRUE)` |
   | STAFF TODAY | F9:H12 | `=SUMIF(tblDispatchKPI[BusinessDate],TODAY(),tblDispatchKPI[TotalStaff])` |
   | CARTONS SHIPPED TODAY | J9:L12 | `=SUMIF(tblDispatchKPI[BusinessDate],TODAY(),tblDispatchKPI[ShippedCartons])` |

5. Each card: merge the cell range, centre-align, 22pt bold white font, coloured background (red for overdue counts, blue for performance/shipping).

---

## STEP 15 — Add Charts

### Dispatch Performance Combo Chart

1. Go to **T_DISPATCH_KPI**.
2. Insert > PivotTable → use `tblDispatchKPI` → place on a new sheet called `PIVOT_DISPATCH`.
3. Rows: BusinessDate + ShiftName; Values: ShippedCartons (Sum), ExpectedCartons (Sum), PerformancePct (Average).
4. Select the PivotTable → Insert > PivotChart → Combo chart.
   - ShippedCartons & ExpectedCartons: Clustered Bar (primary axis).
   - PerformancePct: Line (secondary axis), format as percentage.

### HRP Age Bucket Chart

1. Insert PivotTable from `tblHRP` → `PIVOT_HRP`.
2. Rows: AgeBucket; Values: Count of CartonID.
3. Insert > PivotChart → Clustered Column.

### Top 10 Stores (HRP)

1. Same pivot → add Rows: STORE_NAME.
2. Sort descending by count, keep top 10.
3. Insert > PivotChart → Clustered Bar.

---

## STEP 16 — Sheet ordering

Move sheets in this order (drag tabs):
`DASHBOARD | CONFIG | IN_PACKED | IN_HRP | IN_SHIPPED_LPNS | IN_STAFFING | IN_TARGETS_DAILY | T_DISPATCH_KPI | T_DISPATCH_DAILY | ACTION_HRP | ACTION_PACKED | HISTORY | DATA_QUALITY`

---

## STEP 17 — Tab colours (optional but recommended)

| Tab | Colour |
|-----|--------|
| DASHBOARD | Blue |
| CONFIG | Grey |
| IN_PACKED, IN_HRP | Dark blue |
| IN_SHIPPED_LPNS, IN_STAFFING, IN_TARGETS_DAILY | Green |
| T_DISPATCH_KPI, T_DISPATCH_DAILY | Gold/Yellow |
| ACTION_HRP, ACTION_PACKED | Red |
| HISTORY, DATA_QUALITY | Grey / Yellow |

---

## STEP 18 — Sheet protection (optional)

For formula sheets (T_DISPATCH_KPI, T_DISPATCH_DAILY, IN_PACKED columns O-R, IN_HRP columns N-Q):

1. Select all cells → Format Cells → Protection → uncheck **Locked**.
2. Select formula cells only → Format Cells → Protection → check **Locked**.
3. Review > Protect Sheet → password: `KPI2024`.
   - Allow: Format cells, Insert rows, Delete rows, Sort, AutoFilter.

---

*Setup complete. Proceed to the Daily Operating Procedure for routine usage.*

# Daily Operating Procedure — Warehouse / Dispatch KPI Workbook

> **5-step routine** for end users.  
> Estimated time per shift: ~5 minutes.

---

## STEP 1 — Paste today's source data

### 1a · IN_PACKED
1. Open the packed-carton export from your WMS / fulfilment system.
2. Navigate to the **IN_PACKED** sheet.
3. Click the first empty cell under the **LPN** column (column A).
4. Paste the exported data (**Ctrl + V**).  
   ✅ The table will expand automatically.  
   ✅ Columns **AgeDays**, **IsShipped**, **PackedStatus**, and **ActionFlag** will calculate instantly.
5. **Do not delete the calculated columns (O–R).**

### 1b · IN_HRP
1. Open the HRP export.
2. Navigate to the **IN_HRP** sheet.
3. Paste data under the **CATEGORY** column (column A).  
   ✅ Columns **CartonID**, **IsAcknowledged**, **IncludeInHRP**, and **AgeBucket** auto-calculate.

> **Tip:** Before pasting, clear the previous day's data rows (select rows 2 to end, press Delete).  
> Keep row 1 (headers) intact at all times.

---

## STEP 2 — Log staffing and targets

### 2a · IN_STAFFING
1. Go to the **IN_STAFFING** sheet.
2. Add a new row for each area for today's shift:

   | BusinessDate | ShiftName | Area | StaffAvailable |
   |---|---|---|---|
   | `today's date` | Day **or** Night | *(select from dropdown)* | `number` |

3. Use the **ShiftName dropdown** (Day / Night) and **Area dropdown** (from the list).

### 2b · IN_TARGETS_DAILY (update only if the target has changed)
1. Go to **IN_TARGETS_DAILY**.
2. Add a row:

   | BusinessDate | ShiftName | TargetPerPersonPerShift |
   |---|---|---|
   | `today` | Day / Night | `e.g. 50` |

> The target is used to calculate **ExpectedCartons** in the KPI sheets.  
> If today's target is the same as yesterday, you do not need to add a new row.

---

## STEP 3 — Log shipped LPNs

1. Go to **IN_SHIPPED_LPNS**.
2. For each LPN despatched during the shift, add a row:

   | BusinessDate | ShiftName | LPN | EnteredBy | Notes |
   |---|---|---|---|---|
   | `today` | Day / Night | `LPN code` | `your name` | *(optional)* |

3. Watch for **amber highlighted rows** in the **DupFlag** column — these indicate the same LPN has been logged twice for the same shift/date. Investigate and remove the duplicate.

> **Bulk entry tip:** If you have a batch scan list, paste the LPN column, then fill BusinessDate and ShiftName using the dropdown. Auto-fill works well for these columns.

---

## STEP 4 — Refresh & review KPIs

### 4a · Refresh
1. Go to the **DASHBOARD** sheet.
2. Click the **REFRESH ALL** button.  
   *(This recalculates all formulas and refreshes any PivotTables.)*
3. Verify "Last Refreshed" timestamp updates.

### 4b · Review KPI Cards
Check the six KPI cards on the DASHBOARD:

| Card | Target | Action if breached |
|------|--------|--------------------|
| **HRP OPEN ITEMS** | 0 (or as low as possible) | Review ACTION_HRP sheet |
| **PACKED OVERDUE** | 0 | Review ACTION_PACKED sheet |
| **DISPATCH PERF TODAY** | ≥ 100% (Green) | Check T_DISPATCH_KPI for Red/Amber shifts |
| **DUPLICATE LPNs** | 0 | Correct IN_SHIPPED_LPNS |
| **STAFF TODAY** | Matches roster | Correct IN_STAFFING if wrong |
| **CARTONS SHIPPED TODAY** | ≥ ExpectedCartons | Investigate if low |

### 4c · Populate action sheets
1. Click **POPULATE ACTION SHEETS** on the DASHBOARD.
2. Navigate to **ACTION_HRP** — assign **Owner** and **ContactedCity** (Y/N) for each open item.
3. Navigate to **ACTION_PACKED** — assign **HoldReason**, **Owner**, and **PlannedShipTime** for each overdue item.

### 4d · Review DATA_QUALITY sheet
1. Go to **DATA_QUALITY**.
2. All "Status" cells should show **OK** (green).
3. If any show **WARNING** (amber), investigate that specific check and correct the source data before proceeding.

---

## STEP 5 — Take a snapshot at end of shift

1. Return to the **DASHBOARD** sheet.
2. Click **TAKE DAILY SNAPSHOT**.
3. Confirm the dialog — this appends the current KPI row(s) to the **HISTORY** sheet with a timestamp.
4. **Save the workbook** (Ctrl + S).

> **Why snapshot?** The KPIs in T_DISPATCH_KPI recalculate live — if you paste new data the next day, the previous values will change. The HISTORY sheet preserves a point-in-time record for trend analysis.

---

## Quick Reference — RAG Thresholds

| RAG | Dispatch Performance % |
|-----|----------------------|
| 🟢 Green | ≥ 100% |
| 🟡 Amber | ≥ 90% and < 100% |
| 🔴 Red | < 90% |

*(Thresholds can be changed in the CONFIG sheet: Amber_Threshold and Green_Threshold.)*

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|-------------|-----|
| AgeDays showing huge number | LAST_PACKED column contains text, not a date | Check source data format; re-paste as values |
| IsShipped all FALSE | tblShipped is empty or LPN codes don't match | Verify LPN format matches between IN_PACKED and IN_SHIPPED_LPNS |
| PerformancePct shows 0 | No staffing or target data for that date/shift | Add row to IN_STAFFING and IN_TARGETS_DAILY |
| TargetPerPerson shows 0 | No matching row in IN_TARGETS_DAILY | Add a row for the date + shift combination |
| DupFlag = TRUE | LPN logged twice for same date + shift | Delete the duplicate row in IN_SHIPPED_LPNS |
| #REF! or #NAME? errors | Table name typo or table deleted | Check table names via Formulas > Name Manager |
| Formulas not calculating | Calculation set to Manual | Press **F9** or go to Formulas > Calculate Now |
| Snapshot fails | HISTORY or T_DISPATCH_KPI sheet renamed/deleted | Restore the sheet or re-run the setup macro |

---

## End of Day Checklist

- [ ] IN_PACKED data pasted and up-to-date  
- [ ] IN_HRP data pasted and up-to-date  
- [ ] All shipped LPNs logged in IN_SHIPPED_LPNS  
- [ ] Staffing figures entered in IN_STAFFING  
- [ ] Targets confirmed in IN_TARGETS_DAILY  
- [ ] DASHBOARD refreshed — no RED KPI cards unexplained  
- [ ] DATA_QUALITY shows all OK  
- [ ] Action owners assigned in ACTION_HRP and ACTION_PACKED  
- [ ] Daily Snapshot taken  
- [ ] Workbook saved  

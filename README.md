# Warehouse / Dispatch KPI System — Dashboard DMS v0.0.1-beta

Production-ready Excel KPI system for warehouse/dispatch operations.  
Three ways to deploy — pick whichever suits your environment.

---

## Contents

| File | Description |
|------|-------------|
| `CreateKPIWorkbook.bas` | **VBA macro** — run once inside Excel to build the entire workbook |
| `generate_kpi_workbook.py` | **Python script** (openpyxl) — generates the same workbook outside Excel |
| `MANUAL_SETUP.md` | **Step-by-step manual** — for users who cannot run macros or Python |
| `DAILY_PROCEDURE.md` | **5-step daily routine** — for end users operating the system |

---

## Option A — VBA Macro (recommended, fastest)

### Prerequisites
- Excel 2016 or later (Windows)
- Macros must be enabled

### Steps
1. Open a **blank Excel workbook** (save as `.xlsm`).
2. Press **Alt + F11** to open the VBA Editor.
3. **Insert > Module**.
4. Paste the entire contents of `CreateKPIWorkbook.bas` into the module.
5. Close the VBA Editor.
6. Press **Alt + F8**, select `CreateKPIWorkbook`, click **Run**.
7. Wait ~30 seconds — a dialog confirms completion.

The macro creates all 13 sheets, tables, formulas, conditional formatting,  
data validation, buttons and colour coding in one pass.

**Unprotect password:** `KPI2024`

---

## Option B — Python Script (for locked-down environments)

### Prerequisites
```bash
pip install openpyxl
```

### Steps
```bash
python generate_kpi_workbook.py
# Outputs: KPI_Workbook.xlsx in the current directory
```

Open `KPI_Workbook.xlsx` in Excel. All sheets, tables and formulas are ready.  
Note: Button macros require VBA — add them manually via Alt+F11 if needed.

---

## Option C — Manual Setup

Follow [`MANUAL_SETUP.md`](MANUAL_SETUP.md) — 18 numbered steps covering every sheet,  
table, formula, validation rule and chart.

---

## System Architecture

```
INPUT SHEETS                CALCULATION SHEETS          OUTPUT SHEETS
──────────────────          ────────────────────         ─────────────────────
IN_PACKED        ──┐        T_DISPATCH_KPI    ──┐        DASHBOARD (KPI cards)
IN_HRP           ──┤   →    T_DISPATCH_DAILY  ──┤   →    ACTION_HRP
IN_SHIPPED_LPNS  ──┤        DATA_QUALITY      ──┘        ACTION_PACKED
IN_STAFFING      ──┤                                     HISTORY (snapshots)
IN_TARGETS_DAILY ──┘
CONFIG (rules / lookups)
```

### Shift schedule
| Shift | Hours | Days/week |
|-------|-------|-----------|
| Day   | 08:00 – 16:00 | 6 |
| Night | 16:00 – 00:00 | 6 |

### KPI definitions
| KPI | Formula | RAG |
|-----|---------|-----|
| Dispatch Performance % | Shipped Cartons ÷ Expected Cartons | Green ≥100%, Amber ≥90%, Red <90% |
| HRP Open Count | IncludeInHRP = TRUE rows | 0 = OK, >0 = WARNING |
| Packed Overdue Count | ActionFlag = TRUE rows | 0 = OK, >0 = WARNING |
| Duplicate LPN Count | DupFlag = TRUE rows | 0 = OK, >0 = WARNING |

---

## Formula compatibility

All formulas use **Excel 2016-compatible** functions:
- `COUNTIFS`, `SUMIFS`, `SUMIF`, `COUNTIF`
- `INDEX` / `MATCH` (array formula for multi-criteria lookup)
- `IFERROR`, `IF`, `AND`, `NOT`, `SUMPRODUCT`
- No `XLOOKUP`, no dynamic arrays, no `LET`

---

## Daily routine

See [`DAILY_PROCEDURE.md`](DAILY_PROCEDURE.md) for the full 5-step procedure.

Quick summary:
1. **Paste** IN_PACKED and IN_HRP data
2. **Log** staffing and targets
3. **Log** shipped LPNs
4. **Refresh** dashboard → review KPI cards → populate action sheets
5. **Snapshot** at end of shift, then save

---

## Changelog

### v0.0.1-beta
- Initial release
- 13-sheet workbook structure
- VBA macro + Python generator + manual instructions
- Dispatch KPI (shift + daily), HRP, Packed Overdue, Data Quality tracking
- HISTORY snapshot table
- Excel 2016 compatible formulas

'================================================================================
' CreateKPIWorkbook.bas
' Complete VBA macro to build the Warehouse/Dispatch KPI workbook from scratch.
' How to use:
'   1. Open a blank workbook in Excel 2016 or later.
'   2. Press Alt+F11 to open the VBA editor.
'   3. Insert > Module, paste this entire file, then press F5 (or run
'      "CreateKPIWorkbook" from Macros dialog).
'   4. The macro builds every sheet, table, formula, validation and button.
'   Unprotect password: KPI2024
'================================================================================
Option Explicit

' ---- Colour constants (RGB packed as Long) ---------------------------------
Private Const COL_GREEN   As Long = 5296274   ' #50C878
Private Const COL_AMBER   As Long = 16750848  ' #FFA500
Private Const COL_RED     As Long = 16711680  ' #FF0000
Private Const COL_HEADER  As Long = 1644912   ' #191970 (midnight blue)
Private Const COL_WHITE   As Long = 16777215  ' #FFFFFF
Private Const COL_LGREY   As Long = 15921906  ' #F2F2F2
Private Const UNPROTECT_PW As String = "KPI2024"
' NOTE: This password is intentionally visible in the source code.
' It is a shared workbook protection password documented for all operators
' (not a user authentication credential). Change the constant value if required.

'================================================================================
' ENTRY POINT
'================================================================================
Public Sub CreateKPIWorkbook()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Dim wb As Workbook
    Set wb = ThisWorkbook          ' run inside the target workbook

    ' Remove all existing sheets except one (Excel requires ≥1 sheet)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    Do While wb.Worksheets.Count > 1
        wb.Worksheets(wb.Worksheets.Count).Delete
    Loop
    wb.Worksheets(1).Name = "TEMP_DELETE"

    ' Build sheets in logical order
    Call BuildCONFIG(wb)
    Call BuildIN_PACKED(wb)
    Call BuildIN_HRP(wb)
    Call BuildIN_SHIPPED_LPNS(wb)
    Call BuildIN_STAFFING(wb)
    Call BuildIN_TARGETS_DAILY(wb)
    Call BuildIN_AUDIT_LOG(wb)
    Call BuildT_DISPATCH_KPI(wb)
    Call BuildT_DISPATCH_DAILY(wb)
    Call BuildACTION_HRP(wb)
    Call BuildACTION_PACKED(wb)
    Call BuildHISTORY(wb)
    Call BuildDATA_QUALITY(wb)
    Call BuildCHARTS(wb)
    Call BuildDASHBOARD(wb)

    ' Delete the temporary placeholder sheet
    wb.Worksheets("TEMP_DELETE").Delete

    ' Re-order tabs for usability
    Call ReorderSheets(wb)

    ' Add the control buttons on the DASHBOARD
    Call AddButtons(wb)

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Save as macro-enabled workbook so VBA buttons survive the next open.
    ' If the file already has an xlsm extension this is a no-op save;
    ' otherwise it converts the file in place.
    On Error Resume Next
    Dim xlsmPath As String
    xlsmPath = Left(wb.FullName, InStrRev(wb.FullName, ".")) & "xlsm"
    Application.DisplayAlerts = False
    wb.SaveAs Filename:=xlsmPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
    Application.DisplayAlerts = True
    On Error GoTo 0

    MsgBox "KPI Workbook created successfully!" & vbCrLf & _
           "Unprotect password: " & UNPROTECT_PW & vbCrLf & vbCrLf & _
           "File saved as .xlsm — macros are now embedded and buttons will work on next open.", _
           vbInformation, "Done"
End Sub

'================================================================================
' SHEET BUILDERS
'================================================================================

' ---- CONFIG -----------------------------------------------------------------
Private Sub BuildCONFIG(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "CONFIG"

    ' --- Shift table ---------------------------------------------------------
    ws.Range("A1").Value = "CONFIG - DO NOT DELETE OR RENAME THIS SHEET"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = COL_RED

    ws.Range("A3").Value = "SHIFTS"
    ws.Range("A3").Font.Bold = True

    Dim shiftHdr() As String
    shiftHdr = Split("ShiftName,StartTime,EndTime,DaysPerWeek", ",")
    Dim c As Integer
    For c = 0 To 3
        ws.Cells(4, c + 1).Value = shiftHdr(c)
        ws.Cells(4, c + 1).Font.Bold = True
        ws.Cells(4, c + 1).Interior.Color = COL_HEADER
        ws.Cells(4, c + 1).Font.Color = COL_WHITE
    Next c
    ws.Range("A5").Value = "Day":    ws.Range("B5").Value = "08:00": ws.Range("C5").Value = "16:00": ws.Range("D5").Value = 6
    ws.Range("A6").Value = "Night":  ws.Range("B6").Value = "16:00": ws.Range("C6").Value = "00:00": ws.Range("D6").Value = 6
    ws.Range("B5:C6").NumberFormat = "hh:mm"

    ' Convert to table
    Dim tblShift As ListObject
    Set tblShift = ws.ListObjects.Add(xlSrcRange, ws.Range("A4:D6"), , xlYes)
    tblShift.Name = "tblConfig_Shifts"
    tblShift.TableStyle = "TableStyleMedium2"

    ' --- KPI Rules -----------------------------------------------------------
    ws.Range("A9").Value = "KPI RULES"
    ws.Range("A9").Font.Bold = True

    Dim ruleHdr() As String
    ruleHdr = Split("RuleName,Value,Description", ",")
    For c = 0 To 2
        ws.Cells(10, c + 1).Value = ruleHdr(c)
        ws.Cells(10, c + 1).Font.Bold = True
        ws.Cells(10, c + 1).Interior.Color = COL_HEADER
        ws.Cells(10, c + 1).Font.Color = COL_WHITE
    Next c
    ws.Range("A11").Value = "HRP_MaxDaysToShow":      ws.Range("B11").Value = 1:  ws.Range("C11").Value = "Max days since available for city"
    ws.Range("A12").Value = "Packed_MaxAgeDays":      ws.Range("B12").Value = 2:  ws.Range("C12").Value = "Max age (days) before packed item is overdue"
    ws.Range("A13").Value = "Audit_SampleSize":       ws.Range("B13").Value = 20: ws.Range("C13").Value = "Required number of audits per shift"
    ws.Range("A14").Value = "Amber_Threshold":        ws.Range("B14").Value = 0.9: ws.Range("C14").Value = "Performance % for Amber RAG"
    ws.Range("A15").Value = "Green_Threshold":        ws.Range("B15").Value = 1:  ws.Range("C15").Value = "Performance % for Green RAG"

    Dim tblRules As ListObject
    Set tblRules = ws.ListObjects.Add(xlSrcRange, ws.Range("A10:C15"), , xlYes)
    tblRules.Name = "tblConfig_Rules"
    tblRules.TableStyle = "TableStyleMedium2"

    ' --- Area List (for dropdowns) -------------------------------------------
    ws.Range("E3").Value = "AREA LIST"
    ws.Range("E3").Font.Bold = True
    ws.Range("E4").Value = "AreaName"
    ws.Range("E4").Font.Bold = True
    ws.Range("E4").Interior.Color = COL_HEADER
    ws.Range("E4").Font.Color = COL_WHITE
    ws.Range("E5").Value = "Auditing"
    ws.Range("E6").Value = "Manual handover"
    ws.Range("E7").Value = "Dispatch sealing"

    Dim tblAreas As ListObject
    Set tblAreas = ws.ListObjects.Add(xlSrcRange, ws.Range("E4:E7"), , xlYes)
    tblAreas.Name = "tblConfig_Areas"
    tblAreas.TableStyle = "TableStyleMedium2"

    ws.Columns("A:F").AutoFit
    ws.Tab.Color = RGB(128, 128, 128)
End Sub

' ---- IN_PACKED --------------------------------------------------------------
Private Sub BuildIN_PACKED(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "IN_PACKED"

    Dim srcCols() As String
    srcCols = Split("LPN,PALLET,STORE,DIVISION,FACILITY_NAME,STORE_STATUS,PICK_LOC,UNITS,LAST_UPDATE,DAYS_SINCE_CLOSE,LOCKS,TOT_COST,TOT_RETAIL,LAST_PACKED", ",")
    Dim calcCols() As String
    calcCols = Split("AgeDays,IsShipped,PackedStatus,ActionFlag", ",")

    Dim allCols() As String
    ReDim allCols(UBound(srcCols) + UBound(calcCols) + 1)
    Dim i As Integer
    For i = 0 To UBound(srcCols): allCols(i) = srcCols(i): Next i
    For i = 0 To UBound(calcCols): allCols(UBound(srcCols) + 1 + i) = calcCols(i): Next i

    ' Header row
    Dim col As Integer
    For col = 0 To UBound(allCols)
        ws.Cells(1, col + 1).Value = allCols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = COL_HEADER
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    ' Seed two sample rows so Excel creates the table properly
    ws.Range("A2").Value = "LPN001"
    ws.Range("B2").Value = "PLT001"
    ws.Range("C2").Value = "STORE01"
    ws.Range("D2").Value = "DIV01"
    ws.Range("E2").Value = "FACILITY A"
    ws.Range("F2").Value = "OPEN"
    ws.Range("G2").Value = "LOC001"
    ws.Range("H2").Value = 10
    ws.Range("I2").Value = Now()
    ws.Range("J2").Value = 1
    ws.Range("K2").Value = 0
    ws.Range("L2").Value = 100
    ws.Range("M2").Value = 150
    ws.Range("N2").Value = Now() - 1   ' LAST_PACKED

    ' Create the table
    Dim lastCol As String
    lastCol = Split("A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R", ",")(UBound(allCols))
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:" & lastCol & "2"), , xlYes)
    tbl.Name = "tblPacked"
    tbl.TableStyle = "TableStyleMedium6"

    ' Calculated columns - AgeDays (col 15 = O)
    tbl.ListColumns("AgeDays").DataBodyRange(1).Formula = "=TODAY()-INT([@LAST_PACKED])"
    ' IsShipped (col 16 = P)
    tbl.ListColumns("IsShipped").DataBodyRange(1).Formula = "=COUNTIF(tblShipped[LPN],[@LPN])>0"
    ' PackedStatus
    tbl.ListColumns("PackedStatus").DataBodyRange(1).Formula = _
        "=IF([@IsShipped],""Shipped"",IF([@AgeDays]>2,""Overdue"",""Pending""))"
    ' ActionFlag
    tbl.ListColumns("ActionFlag").DataBodyRange(1).Formula = _
        "=AND(NOT([@IsShipped]),[@AgeDays]>2)"

    ' Format date columns
    ws.Columns("I:I").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Columns("N:N").NumberFormat = "dd/mm/yyyy hh:mm"

    ' Conditional formatting on PackedStatus column
    Dim cfRange As Range
    Set cfRange = tbl.ListColumns("PackedStatus").DataBodyRange

    With cfRange.FormatConditions.Add(xlCellValue, xlEqual, """Overdue""")
        .Interior.Color = COL_RED
        .Font.Color = COL_WHITE
    End With
    With cfRange.FormatConditions.Add(xlCellValue, xlEqual, """Shipped""")
        .Interior.Color = COL_GREEN
        .Font.Color = COL_WHITE
    End With
    With cfRange.FormatConditions.Add(xlCellValue, xlEqual, """Pending""")
        .Interior.Color = COL_AMBER
    End With

    ' Freeze panes at row 2
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

    ws.Columns("A:R").AutoFit
    ws.Tab.Color = RGB(0, 112, 192)

    ' Protect sheet - allow editing in non-formula columns only
    ws.Protect Password:=UNPROTECT_PW, DrawingObjects:=False, Contents:=True, _
        AllowFormattingCells:=True, AllowInsertingRows:=True, AllowDeletingRows:=True, _
        AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
End Sub

' ---- IN_HRP -----------------------------------------------------------------
Private Sub BuildIN_HRP(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "IN_HRP"

    Dim srcCols() As String
    srcCols = Split("CATEGORY,DAYS_SINCE_AVAILABLE_FOR_CITY,DATE_SCANNED_TO_CITY,OLPN,XREF_OLPN,STORE_NUMBER,STORE_NAME,CARTON_LOCKS,LPN_STATUS,STORE_ACK_STATUS,UNITS,COST_VALUE,RETAIL_VALUE", ",")
    Dim calcCols() As String
    calcCols = Split("CartonID,IsAcknowledged,IncludeInHRP,AgeBucket", ",")

    Dim col As Integer
    For col = 0 To UBound(srcCols)
        ws.Cells(1, col + 1).Value = srcCols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = COL_HEADER
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col
    Dim offset As Integer
    offset = UBound(srcCols) + 1
    For col = 0 To UBound(calcCols)
        ws.Cells(1, offset + col + 1).Value = calcCols(col)
        ws.Cells(1, offset + col + 1).Font.Bold = True
        ws.Cells(1, offset + col + 1).Interior.Color = COL_HEADER
        ws.Cells(1, offset + col + 1).Font.Color = COL_WHITE
    Next col

    ' Seed one row
    ws.Range("A2").Value = "HRP"
    ws.Range("B2").Value = 0
    ws.Range("C2").Value = Now()
    ws.Range("D2").Value = "OLPN001"
    ws.Range("E2").Value = ""
    ws.Range("F2").Value = "STORE01"
    ws.Range("G2").Value = "Store One"
    ws.Range("H2").Value = 0
    ws.Range("I2").Value = "AVAILABLE"
    ws.Range("J2").Value = "N"
    ws.Range("K2").Value = 5
    ws.Range("L2").Value = 50
    ws.Range("M2").Value = 75

    ' Total 17 columns (13 src + 4 calc)
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:Q2"), , xlYes)
    tbl.Name = "tblHRP"
    tbl.TableStyle = "TableStyleMedium7"

    ' Calculated columns
    tbl.ListColumns("CartonID").DataBodyRange(1).Formula = _
        "=IF([@OLPN]<>"""",[@OLPN],[@XREF_OLPN])"
    tbl.ListColumns("IsAcknowledged").DataBodyRange(1).Formula = _
        "=[@STORE_ACK_STATUS]=""Y"""
    tbl.ListColumns("IncludeInHRP").DataBodyRange(1).Formula = _
        "=AND([@STORE_ACK_STATUS]=""N"",[@DAYS_SINCE_AVAILABLE_FOR_CITY]<=1)"
    tbl.ListColumns("AgeBucket").DataBodyRange(1).Formula = _
        "=IF([@DAYS_SINCE_AVAILABLE_FOR_CITY]=0,""<24h"",""24h+"")"

    ' Data validation on STORE_ACK_STATUS column
    With tbl.ListColumns("STORE_ACK_STATUS").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Y,N"
        .IgnoreBlank = True
        .ShowDropDown = False
        .InputTitle = "Ack Status"
        .InputMessage = "Enter Y or N"
    End With

    ' Conditional formatting on IncludeInHRP
    With tbl.ListColumns("IncludeInHRP").DataBodyRange.FormatConditions.Add( _
            xlCellValue, xlEqual, "TRUE")
        .Interior.Color = COL_RED
        .Font.Color = COL_WHITE
    End With

    ws.Columns("C:C").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Columns("A:Q").AutoFit
    ws.Tab.Color = RGB(0, 176, 240)
End Sub

' ---- IN_SHIPPED_LPNS --------------------------------------------------------
Private Sub BuildIN_SHIPPED_LPNS(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "IN_SHIPPED_LPNS"

    Dim cols() As String
    cols = Split("BusinessDate,ShiftName,LPN,EnteredBy,Notes,DupFlag", ",")

    Dim col As Integer
    For col = 0 To UBound(cols)
        ws.Cells(1, col + 1).Value = cols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = COL_HEADER
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    ' Seed one row
    ws.Range("A2").Value = Date
    ws.Range("B2").Value = "Day"
    ws.Range("C2").Value = "LPN001"
    ws.Range("D2").Value = "User1"
    ws.Range("E2").Value = ""

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:F2"), , xlYes)
    tbl.Name = "tblShipped"
    tbl.TableStyle = "TableStyleMedium9"

    ' DupFlag formula
    tbl.ListColumns("DupFlag").DataBodyRange(1).Formula = _
        "=COUNTIFS([BusinessDate],[@BusinessDate],[ShiftName],[@ShiftName],[LPN],[@LPN])>1"

    ' Conditional formatting – highlight duplicate rows
    With tbl.ListColumns("DupFlag").DataBodyRange.FormatConditions.Add( _
            xlCellValue, xlEqual, "TRUE")
        .Interior.Color = COL_AMBER
    End With

    ' Data validation on ShiftName
    With tbl.ListColumns("ShiftName").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Day,Night"
        .IgnoreBlank = True
        .InputTitle = "Shift"
        .InputMessage = "Select Day or Night"
    End With

    ws.Columns("A:A").NumberFormat = "dd/mm/yyyy"
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Columns("A:F").AutoFit
    ws.Tab.Color = RGB(0, 176, 80)
End Sub

' ---- IN_STAFFING ------------------------------------------------------------
Private Sub BuildIN_STAFFING(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "IN_STAFFING"

    Dim cols() As String
    cols = Split("BusinessDate,ShiftName,Area,StaffAvailable", ",")

    Dim col As Integer
    For col = 0 To UBound(cols)
        ws.Cells(1, col + 1).Value = cols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = COL_HEADER
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    ws.Range("A2").Value = Date
    ws.Range("B2").Value = "Day"
    ws.Range("C2").Value = "Dispatch sealing"
    ws.Range("D2").Value = 5

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:D2"), , xlYes)
    tbl.Name = "tblStaffing"
    tbl.TableStyle = "TableStyleMedium4"

    ' Shift validation
    With tbl.ListColumns("ShiftName").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Day,Night"
        .IgnoreBlank = True
    End With

    ' Area validation — use absolute reference to CONFIG area list;
    ' cross-sheet structured references (=tblConfig_Areas[AreaName]) are
    ' unreliable in data validation and can cause Excel to crash/fail silently.
    With tbl.ListColumns("Area").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=CONFIG!$E$5:$E$7"
        .IgnoreBlank = True
    End With

    ws.Columns("A:A").NumberFormat = "dd/mm/yyyy"
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Columns("A:D").AutoFit
    ws.Tab.Color = RGB(0, 176, 80)
End Sub

' ---- IN_TARGETS_DAILY -------------------------------------------------------
Private Sub BuildIN_TARGETS_DAILY(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "IN_TARGETS_DAILY"

    Dim cols() As String
    cols = Split("BusinessDate,ShiftName,TargetPerPersonPerShift", ",")

    Dim col As Integer
    For col = 0 To UBound(cols)
        ws.Cells(1, col + 1).Value = cols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = COL_HEADER
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    ws.Range("A2").Value = Date
    ws.Range("B2").Value = "Day"
    ws.Range("C2").Value = 50

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C2"), , xlYes)
    tbl.Name = "tblTargetsDaily"
    tbl.TableStyle = "TableStyleMedium4"

    With tbl.ListColumns("ShiftName").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Day,Night"
        .IgnoreBlank = True
    End With

    ws.Columns("A:A").NumberFormat = "dd/mm/yyyy"
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Columns("A:C").AutoFit
    ws.Tab.Color = RGB(0, 176, 80)
End Sub

' ---- IN_AUDIT_LOG -----------------------------------------------------------
Private Sub BuildIN_AUDIT_LOG(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "IN_AUDIT_LOG"

    Dim cols() As String
    cols = Split("BusinessDate,ShiftName,Area,AuditCount,Notes", ",")

    Dim col As Integer
    For col = 0 To UBound(cols)
        ws.Cells(1, col + 1).Value = cols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = RGB(112, 173, 71)
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    ws.Range("A2").Value = Date
    ws.Range("B2").Value = "Day"
    ws.Range("C2").Value = "Auditing"
    ws.Range("D2").Value = 0

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E2"), , xlYes)
    tbl.Name = "tblAuditLog"
    tbl.TableStyle = "TableStyleMedium4"

    ' Shift validation
    With tbl.ListColumns("ShiftName").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Day,Night"
        .IgnoreBlank = True
    End With

    ' Area validation (same fix as IN_STAFFING — use absolute ref)
    With tbl.ListColumns("Area").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=CONFIG!$E$5:$E$7"
        .IgnoreBlank = True
    End With

    ws.Columns("A:A").NumberFormat = "dd/mm/yyyy"
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Columns("A:E").AutoFit
    ws.Tab.Color = RGB(112, 173, 71)
End Sub

' ---- T_DISPATCH_KPI ---------------------------------------------------------
Private Sub BuildT_DISPATCH_KPI(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "T_DISPATCH_KPI"

    Dim cols() As String
    cols = Split("BusinessDate,ShiftName,ShippedCartons,TotalStaff,TargetPerPerson,ExpectedCartons,PerformancePct,RAG,AuditCount,AuditTarget,AuditPct,HRPOpen,PackedOverdue", ",")

    Dim col As Integer
    For col = 0 To UBound(cols)
        ws.Cells(1, col + 1).Value = cols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = COL_HEADER
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    ' Seed row
    ws.Range("A2").Value = Date
    ws.Range("B2").Value = "Day"

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:M2"), , xlYes)
    tbl.Name = "tblDispatchKPI"
    tbl.TableStyle = "TableStyleMedium2"

    ' Calculated columns
    tbl.ListColumns("ShippedCartons").DataBodyRange(1).Formula = _
        "=COUNTIFS(tblShipped[BusinessDate],[@BusinessDate],tblShipped[ShiftName],[@ShiftName])"
    tbl.ListColumns("TotalStaff").DataBodyRange(1).Formula = _
        "=SUMIFS(tblStaffing[StaffAvailable],tblStaffing[BusinessDate],[@BusinessDate],tblStaffing[ShiftName],[@ShiftName])"
    ' TargetPerPerson: array formula for Excel 2016 (multi-criteria INDEX/MATCH)
    tbl.ListColumns("TargetPerPerson").DataBodyRange(1).FormulaArray = _
        "=IFERROR(INDEX(tblTargetsDaily[TargetPerPersonPerShift],MATCH(1,(tblTargetsDaily[BusinessDate]=[@BusinessDate])*(tblTargetsDaily[ShiftName]=[@ShiftName]),0)),0)"

    tbl.ListColumns("ExpectedCartons").DataBodyRange(1).Formula = _
        "=[@TotalStaff]*[@TargetPerPerson]"
    tbl.ListColumns("PerformancePct").DataBodyRange(1).Formula = _
        "=IFERROR([@ShippedCartons]/[@ExpectedCartons],0)"
    tbl.ListColumns("RAG").DataBodyRange(1).Formula = _
        "=IF([@PerformancePct]>=1,""Green"",IF([@PerformancePct]>=0.9,""Amber"",""Red""))"

    ' Audit metrics
    tbl.ListColumns("AuditCount").DataBodyRange(1).Formula = _
        "=SUMIFS(tblAuditLog[AuditCount],tblAuditLog[BusinessDate],[@BusinessDate],tblAuditLog[ShiftName],[@ShiftName])"
    tbl.ListColumns("AuditTarget").DataBodyRange(1).Formula = _
        "=IFERROR(INDEX(tblConfig_Rules[Value],MATCH(""Audit_SampleSize"",tblConfig_Rules[RuleName],0)),20)"
    tbl.ListColumns("AuditPct").DataBodyRange(1).Formula = _
        "=IFERROR([@AuditCount]/[@AuditTarget],0)"
    ' Live snapshot counts (reflect current state of source tables)
    tbl.ListColumns("HRPOpen").DataBodyRange(1).Formula = _
        "=COUNTIF(tblHRP[IncludeInHRP],TRUE)"
    tbl.ListColumns("PackedOverdue").DataBodyRange(1).Formula = _
        "=COUNTIF(tblPacked[ActionFlag],TRUE)"

    ' Format columns
    ws.Columns("A:A").NumberFormat = "dd/mm/yyyy"
    ws.Columns("G:G").NumberFormat = "0.0%"
    ws.Columns("K:K").NumberFormat = "0.0%"

    ' RAG conditional formatting
    Dim ragRange As Range
    Set ragRange = tbl.ListColumns("RAG").DataBodyRange
    With ragRange.FormatConditions.Add(xlCellValue, xlEqual, """Green""")
        .Interior.Color = COL_GREEN
        .Font.Color = COL_WHITE
    End With
    With ragRange.FormatConditions.Add(xlCellValue, xlEqual, """Amber""")
        .Interior.Color = COL_AMBER
    End With
    With ragRange.FormatConditions.Add(xlCellValue, xlEqual, """Red""")
        .Interior.Color = COL_RED
        .Font.Color = COL_WHITE
    End With

    ' PerformancePct data bar
    With tbl.ListColumns("PerformancePct").DataBodyRange.FormatConditions.AddDatabar
        .MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
        .MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1
        .BarColor.Color = RGB(0, 112, 192)
    End With

    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Columns("A:M").AutoFit
    ws.Tab.Color = RGB(255, 192, 0)
End Sub

' ---- T_DISPATCH_DAILY -------------------------------------------------------
Private Sub BuildT_DISPATCH_DAILY(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "T_DISPATCH_DAILY"

    Dim cols() As String
    cols = Split("BusinessDate,TotalShipped,TotalStaff,TotalExpected,DailyPerformancePct", ",")

    Dim col As Integer
    For col = 0 To UBound(cols)
        ws.Cells(1, col + 1).Value = cols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = COL_HEADER
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    ws.Range("A2").Value = Date

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E2"), , xlYes)
    tbl.Name = "tblDispatchDaily"
    tbl.TableStyle = "TableStyleMedium2"

    tbl.ListColumns("TotalShipped").DataBodyRange(1).Formula = _
        "=SUMIF(tblDispatchKPI[BusinessDate],[@BusinessDate],tblDispatchKPI[ShippedCartons])"
    tbl.ListColumns("TotalStaff").DataBodyRange(1).Formula = _
        "=SUMIF(tblDispatchKPI[BusinessDate],[@BusinessDate],tblDispatchKPI[TotalStaff])"
    tbl.ListColumns("TotalExpected").DataBodyRange(1).Formula = _
        "=SUMIF(tblDispatchKPI[BusinessDate],[@BusinessDate],tblDispatchKPI[ExpectedCartons])"
    tbl.ListColumns("DailyPerformancePct").DataBodyRange(1).Formula = _
        "=IFERROR([@TotalShipped]/[@TotalExpected],0)"

    ws.Columns("A:A").NumberFormat = "dd/mm/yyyy"
    ws.Columns("E:E").NumberFormat = "0.0%"

    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Columns("A:E").AutoFit
    ws.Tab.Color = RGB(255, 192, 0)
End Sub

' ---- ACTION_HRP -------------------------------------------------------------
Private Sub BuildACTION_HRP(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "ACTION_HRP"

    Dim cols() As String
    cols = Split("CATEGORY,CartonID,STORE_NUMBER,STORE_NAME,DAYS_SINCE_AVAILABLE_FOR_CITY,STORE_ACK_STATUS,UNITS,COST_VALUE,Owner,ContactedCity,ContactTime,NextStep", ",")

    Dim col As Integer
    For col = 0 To UBound(cols)
        ws.Cells(1, col + 1).Value = cols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = RGB(192, 0, 0)
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    ws.Range("A2").Value = "-- Paste filtered HRP data here or refresh from macro --"
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = RGB(128, 128, 128)

    ' Note box
    ws.Range("A4").Value = "HOW TO REFRESH:"
    ws.Range("A5").Value = "1. Click 'Refresh All' button on DASHBOARD sheet"
    ws.Range("A6").Value = "2. This sheet auto-populates from IN_HRP where IncludeInHRP = TRUE"
    ws.Range("A4:A6").Font.Italic = True
    ws.Range("A4").Font.Bold = True

    ' Validation on ContactedCity
    ws.Range("J2:J1000").Validation.Delete
    With ws.Range("J2:J1000").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="Y,N"
        .IgnoreBlank = True
        .InputTitle = "Contacted City?"
        .InputMessage = "Y = Yes, N = No"
    End With

    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Columns("A:L").AutoFit
    ws.Tab.Color = RGB(192, 0, 0)
End Sub

' ---- ACTION_PACKED ----------------------------------------------------------
Private Sub BuildACTION_PACKED(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "ACTION_PACKED"

    Dim cols() As String
    cols = Split("LPN,STORE,DIVISION,AgeDays,PackedStatus,UNITS,HoldReason,Owner,PlannedShipTime", ",")

    Dim col As Integer
    For col = 0 To UBound(cols)
        ws.Cells(1, col + 1).Value = cols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = RGB(192, 0, 0)
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    ws.Range("A2").Value = "-- Paste filtered PACKED data here or refresh from macro --"
    ws.Range("A2").Font.Italic = True
    ws.Range("A2").Font.Color = RGB(128, 128, 128)

    ws.Range("A4").Value = "Sorted by AgeDays descending. Only rows where ActionFlag = TRUE are shown."
    ws.Range("A4").Font.Italic = True

    ' CF on PackedStatus
    With ws.Range("E2:E1000").FormatConditions.Add(xlCellValue, xlEqual, """Overdue""")
        .Interior.Color = COL_RED
        .Font.Color = COL_WHITE
    End With
    With ws.Range("E2:E1000").FormatConditions.Add(xlCellValue, xlEqual, """Pending""")
        .Interior.Color = COL_AMBER
    End With

    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Columns("A:I").AutoFit
    ws.Tab.Color = RGB(192, 0, 0)
End Sub

' ---- HISTORY ----------------------------------------------------------------
Private Sub BuildHISTORY(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "HISTORY"

    Dim cols() As String
    cols = Split("SnapshotTimestamp,BusinessDate,ShiftName,ShippedCartons,TotalStaff,ExpectedCartons,PerformancePct,RAG,HRP_OpenCount,Packed_OverdueCount,AuditCount", ",")

    Dim col As Integer
    For col = 0 To UBound(cols)
        ws.Cells(1, col + 1).Value = cols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = COL_HEADER
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:K1"), , xlYes)
    tbl.Name = "tblHistory"
    tbl.TableStyle = "TableStyleMedium2"
    tbl.ShowAutoFilterDropDown = True

    ws.Columns("A:B").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Columns("G:G").NumberFormat = "0.0%"
    ws.Columns("A:K").AutoFit
    ws.Tab.Color = RGB(128, 128, 128)
End Sub

' ---- DATA_QUALITY -----------------------------------------------------------
Private Sub BuildDATA_QUALITY(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "DATA_QUALITY"

    ws.Range("A1").Value = "DATA QUALITY CHECKS"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Color = COL_HEADER

    Dim headers() As String
    headers = Split("Check,Formula / Logic,Result,Status", ",")
    Dim col As Integer
    For col = 0 To 3
        ws.Cells(3, col + 1).Value = headers(col)
        ws.Cells(3, col + 1).Font.Bold = True
        ws.Cells(3, col + 1).Interior.Color = COL_HEADER
        ws.Cells(3, col + 1).Font.Color = COL_WHITE
    Next col

    Dim checks(8, 1) As String
    checks(0, 0) = "Missing CartonID in HRP"
    checks(0, 1) = "=COUNTIFS(tblHRP[CartonID],"""")"
    checks(1, 0) = "Future dates in IN_PACKED (LAST_PACKED)"
    checks(1, 1) = "=COUNTIF(tblPacked[LAST_PACKED],"">"" & TODAY())"
    checks(2, 0) = "Future dates in IN_SHIPPED_LPNS (BusinessDate)"
    checks(2, 1) = "=COUNTIF(tblShipped[BusinessDate],"">"" & TODAY())"
    checks(3, 0) = "Duplicate LPNs in Shipped list"
    checks(3, 1) = "=COUNTIF(tblShipped[DupFlag],TRUE)"
    checks(4, 0) = "StaffAvailable = 0 rows in Staffing"
    checks(4, 1) = "=COUNTIF(tblStaffing[StaffAvailable],0)"
    checks(5, 0) = "Invalid ShiftName in Shipped (not Day/Night)"
    checks(5, 1) = "=SUMPRODUCT((tblShipped[ShiftName]<>""Day"")*(tblShipped[ShiftName]<>""Night"")*(tblShipped[ShiftName]<>""""))"
    checks(6, 0) = "Invalid ShiftName in Staffing"
    checks(6, 1) = "=SUMPRODUCT((tblStaffing[ShiftName]<>""Day"")*(tblStaffing[ShiftName]<>""Night"")*(tblStaffing[ShiftName]<>""""))"
    checks(7, 0) = "HRP rows with IncludeInHRP=TRUE (open action items)"
    checks(7, 1) = "=COUNTIF(tblHRP[IncludeInHRP],TRUE)"
    checks(8, 0) = "Packed items where ActionFlag=TRUE (overdue)"
    checks(8, 1) = "=COUNTIF(tblPacked[ActionFlag],TRUE)"

    Dim r As Integer
    For r = 0 To 8
        ws.Cells(r + 4, 1).Value = checks(r, 0)
        ws.Cells(r + 4, 2).Value = "'" & checks(r, 1)   ' leading apostrophe forces Excel to store as display text, not an evaluated formula
        ws.Cells(r + 4, 3).Formula = checks(r, 1)        ' also evaluate
        ' All Result formulas (COUNTIF / SUMPRODUCT) return 0 or a positive integer —
        ' they can never return a negative value — so 0 = OK and >0 = WARNING.
        ws.Cells(r + 4, 4).Formula = _
            "=IF(C" & (r + 4) & "=0,""OK"",""WARNING"")"

        ' CF on Status
        With ws.Cells(r + 4, 4).FormatConditions.Add(xlCellValue, xlEqual, """WARNING""")
            .Interior.Color = COL_AMBER
            .Font.Bold = True
        End With
        With ws.Cells(r + 4, 4).FormatConditions.Add(xlCellValue, xlEqual, """OK""")
            .Interior.Color = COL_GREEN
            .Font.Color = COL_WHITE
        End With
    Next r

    ws.Range("A3:D" & (8 + 4)).Borders.LineStyle = xlContinuous
    ws.Columns("A:D").AutoFit
    ws.Tab.Color = RGB(255, 192, 0)
End Sub

' ---- DASHBOARD --------------------------------------------------------------
Private Sub BuildDASHBOARD(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "DASHBOARD"

    ' Light-grey page background (Excel-style)
    ws.Cells.Interior.Color = RGB(242, 242, 242)
    ws.Cells.Font.Color = RGB(51, 51, 51)

    ' ── Row 1: Title bar ──────────────────────────────────────────────────────
    ws.Rows(1).RowHeight = 36
    With ws.Range("A1:N1")
        .Merge
        .Value = "WAREHOUSE / DISPATCH KPI DASHBOARD"
        .Font.Bold = True
        .Font.Size = 18
        .Font.Color = COL_WHITE
        .Interior.Color = COL_HEADER
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' ── Row 2: Meta / last-refreshed bar ─────────────────────────────────────
    ws.Rows(2).RowHeight = 18
    Dim subBG As Long: subBG = RGB(236, 240, 247)
    With ws.Range("A2:C2")
        .Merge
        .Value = "Last Refreshed:"
        .Font.Bold = True: .Font.Size = 9: .Font.Color = RGB(102, 102, 102)
        .Interior.Color = subBG: .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
    End With
    With ws.Range("D2:G2")
        .Merge
        .Formula = "=TEXT(NOW(),""dd/mm/yyyy hh:mm"")"
        .Font.Bold = True: .Font.Size = 9: .Font.Color = RGB(0, 112, 192)
        .Interior.Color = subBG: .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    With ws.Range("H2:N2")
        .Merge
        .Value = "Warehouse / Dispatch KPI System  v0.0.1-beta"
        .Font.Italic = True: .Font.Size = 9: .Font.Color = RGB(102, 102, 102)
        .Interior.Color = subBG: .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
    End With

    ' Row 3: thin visual spacer
    ws.Rows(3).RowHeight = 6

    ' ── Row 4: Section header "KPI OVERVIEW" ─────────────────────────────────
    ws.Rows(4).RowHeight = 20
    Call MakeSectionHeader(ws, "A4:N4", "  KPI OVERVIEW", RGB(46, 64, 87))

    ' ── Rows 5-8: KPI cards — attention/alert row ────────────────────────────
    Call MakeKPICard(ws, 5, 2, "HRP OPEN ITEMS", _
        "=COUNTIF(tblHRP[IncludeInHRP],TRUE)", COL_RED)
    Call MakeKPICard(ws, 5, 6, "PACKED OVERDUE", _
        "=COUNTIF(tblPacked[ActionFlag],TRUE)", COL_RED)
    ' Fixed: INDEX/MATCH raises #N/A when no row in tblDispatchDaily matches
    ' today's date, which IFERROR catches and displays as "N/A".  The previous
    ' SUMIF returned 0 for an unmatched date, causing TEXT to show "0.0%"
    ' (a misleading result when no data has been entered for today).
    Call MakeKPICard(ws, 5, 10, "DISPATCH PERF TODAY", _
        "=IFERROR(TEXT(INDEX(tblDispatchDaily[DailyPerformancePct]," & _
        "MATCH(TODAY(),tblDispatchDaily[BusinessDate],0)),""0.0%""),""N/A"")", _
        RGB(0, 112, 192))

    ' ── Rows 9-12: KPI cards — operations row ────────────────────────────────
    Call MakeKPICard(ws, 9, 2, "DUPLICATE LPNs", _
        "=COUNTIF(tblShipped[DupFlag],TRUE)", COL_AMBER)
    Call MakeKPICard(ws, 9, 6, "STAFF TODAY", _
        "=SUMIF(tblDispatchKPI[BusinessDate],TODAY(),tblDispatchKPI[TotalStaff])", _
        RGB(0, 112, 192))
    Call MakeKPICard(ws, 9, 10, "CARTONS SHIPPED TODAY", _
        "=SUMIF(tblDispatchKPI[BusinessDate],TODAY(),tblDispatchKPI[ShippedCartons])", _
        COL_GREEN)

    ' Row 13: spacer
    ws.Rows(13).RowHeight = 8

    ' ── Row 14: Section header "PERFORMANCE SUMMARY" ─────────────────────────
    ws.Rows(14).RowHeight = 20
    Call MakeSectionHeader(ws, "A14:N14", "  PERFORMANCE SUMMARY  (Most Recent Shifts)", RGB(46, 64, 87))

    ' Row 15: table column headers
    ws.Rows(15).RowHeight = 16
    Dim perfHdrs() As String
    perfHdrs = Split("Date,Shift,Shipped,Expected,Perf %,RAG", ",")
    Dim ph As Integer
    For ph = 0 To 5
        ws.Cells(15, ph + 1).Value = perfHdrs(ph)
        ws.Cells(15, ph + 1).Font.Bold = True
        ws.Cells(15, ph + 1).Font.Color = COL_WHITE
        ws.Cells(15, ph + 1).Font.Size = 9
        ws.Cells(15, ph + 1).Interior.Color = RGB(0, 112, 192)
        ws.Cells(15, ph + 1).HorizontalAlignment = xlCenter
        ws.Cells(15, ph + 1).Borders.LineStyle = xlContinuous
    Next ph

    ' Rows 16-18: 3 most-recent rows from tblDispatchKPI (newest first)
    Dim rowOff As Integer
    For rowOff = 0 To 2
        Dim rn As Long: rn = 16 + rowOff
        Dim idxExpr As String
        idxExpr = "ROWS(tblDispatchKPI[BusinessDate])-" & rowOff
        ws.Rows(rn).RowHeight = 14
        Dim rowBG As Long
        If rowOff Mod 2 = 0 Then rowBG = COL_WHITE Else rowBG = RGB(247, 247, 247)

        ws.Cells(rn, 1).Formula = "=IFERROR(TEXT(INDEX(tblDispatchKPI[BusinessDate]," & idxExpr & "),""dd/mm/yyyy""),"""")"
        ws.Cells(rn, 2).Formula = "=IFERROR(INDEX(tblDispatchKPI[ShiftName]," & idxExpr & "),"""")"
        ws.Cells(rn, 3).Formula = "=IFERROR(INDEX(tblDispatchKPI[ShippedCartons]," & idxExpr & "),"""")"
        ws.Cells(rn, 4).Formula = "=IFERROR(INDEX(tblDispatchKPI[ExpectedCartons]," & idxExpr & "),"""")"
        ws.Cells(rn, 5).Formula = "=IFERROR(TEXT(INDEX(tblDispatchKPI[PerformancePct]," & idxExpr & "),""0.0%""),"""")"
        ws.Cells(rn, 6).Formula = "=IFERROR(INDEX(tblDispatchKPI[RAG]," & idxExpr & "),"""")"

        Dim col As Integer
        For col = 1 To 6
            ws.Cells(rn, col).Font.Size = 9
            ws.Cells(rn, col).Interior.Color = rowBG
            ws.Cells(rn, col).HorizontalAlignment = xlCenter
            ws.Cells(rn, col).Borders.LineStyle = xlContinuous
        Next col

        ' RAG conditional formatting
        With ws.Cells(rn, 6).FormatConditions.Add(xlCellValue, xlEqual, """Green""")
            .Interior.Color = COL_GREEN: .Font.Color = COL_WHITE
        End With
        With ws.Cells(rn, 6).FormatConditions.Add(xlCellValue, xlEqual, """Amber""")
            .Interior.Color = COL_AMBER
        End With
        With ws.Cells(rn, 6).FormatConditions.Add(xlCellValue, xlEqual, """Red""")
            .Interior.Color = COL_RED: .Font.Color = COL_WHITE
        End With
    Next rowOff

    ' Row 19: spacer
    ws.Rows(19).RowHeight = 8

    ' ── Row 20: AUDIT OVERVIEW section ────────────────────────────────────────
    ws.Rows(20).RowHeight = 20
    Call MakeSectionHeader(ws, "A20:N20", "  AUDIT OVERVIEW  (Today)", RGB(112, 173, 71))

    Call MakeKPICard(ws, 21, 2, "AUDITS TODAY", _
        "=IFERROR(SUMIFS(tblAuditLog[AuditCount],tblAuditLog[BusinessDate],TODAY()),""N/A"")", _
        RGB(112, 173, 71))
    Call MakeKPICard(ws, 21, 6, "AUDIT TARGET (PER SHIFT)", _
        "=IFERROR(INDEX(tblConfig_Rules[Value],MATCH(""Audit_SampleSize"",tblConfig_Rules[RuleName],0)),""N/A"")", _
        RGB(112, 173, 71))
    Call MakeKPICard(ws, 21, 10, "AUDIT COMPLIANCE TODAY", _
        "=IFERROR(TEXT(SUMIFS(tblAuditLog[AuditCount],tblAuditLog[BusinessDate],TODAY())" & _
        "/INDEX(tblConfig_Rules[Value],MATCH(""Audit_SampleSize"",tblConfig_Rules[RuleName],0)),""0.0%""),""N/A"")", _
        RGB(112, 173, 71))

    ' Row 25: spacer
    ws.Rows(25).RowHeight = 8

    ' ── Row 26: RAG LEGEND section ────────────────────────────────────────────
    ws.Rows(26).RowHeight = 20
    Call MakeSectionHeader(ws, "A26:N26", "  RAG LEGEND", RGB(68, 114, 196))

    ' Row 27: legend colour blocks
    ws.Rows(27).RowHeight = 16
    With ws.Range("B27:D27")
        .Merge: .Value = "GREEN": .Font.Bold = True: .Font.Size = 9: .Font.Color = COL_WHITE
        .Interior.Color = COL_GREEN: .HorizontalAlignment = xlCenter: .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("F27:H27")
        .Merge: .Value = "AMBER": .Font.Bold = True: .Font.Size = 9: .Font.Color = RGB(51, 51, 51)
        .Interior.Color = COL_AMBER: .HorizontalAlignment = xlCenter: .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("J27:L27")
        .Merge: .Value = "RED": .Font.Bold = True: .Font.Size = 9: .Font.Color = COL_WHITE
        .Interior.Color = COL_RED: .HorizontalAlignment = xlCenter: .Borders.LineStyle = xlContinuous
    End With

    ' Row 28: legend descriptions
    ws.Rows(28).RowHeight = 14
    With ws.Range("B28:D28")
        .Merge: .Value = "Dispatch Performance >= 100%"
        .Font.Italic = True: .Font.Size = 8: .Font.Color = RGB(102, 102, 102)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("F28:H28")
        .Merge: .Value = "90% <= Performance < 100%"
        .Font.Italic = True: .Font.Size = 8: .Font.Color = RGB(102, 102, 102)
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range("J28:L28")
        .Merge: .Value = "Performance < 90%"
        .Font.Italic = True: .Font.Size = 8: .Font.Color = RGB(102, 102, 102)
        .HorizontalAlignment = xlCenter
    End With

    ' Row 29: spacer
    ws.Rows(29).RowHeight = 8

    ' ── Row 30: QUICK NAVIGATION section ─────────────────────────────────────
    ws.Rows(30).RowHeight = 20
    Call MakeSectionHeader(ws, "A30:N30", "  QUICK NAVIGATION", RGB(89, 89, 89))

    ' Row 31: primary navigation links
    ws.Rows(31).RowHeight = 16
    Dim navItems(6, 1) As String
    navItems(0, 0) = "IN_PACKED":       navItems(0, 1) = "#IN_PACKED!A1"
    navItems(1, 0) = "IN_HRP":          navItems(1, 1) = "#IN_HRP!A1"
    navItems(2, 0) = "IN_SHIPPED_LPNS": navItems(2, 1) = "#IN_SHIPPED_LPNS!A1"
    navItems(3, 0) = "IN_STAFFING":     navItems(3, 1) = "#IN_STAFFING!A1"
    navItems(4, 0) = "IN_AUDIT_LOG":    navItems(4, 1) = "#IN_AUDIT_LOG!A1"
    navItems(5, 0) = "T_DISPATCH_KPI":  navItems(5, 1) = "#T_DISPATCH_KPI!A1"
    navItems(6, 0) = "HISTORY":         navItems(6, 1) = "#HISTORY!A1"

    Dim ni As Integer
    For ni = 0 To 6
        Dim nc As Range
        Set nc = ws.Range(ws.Cells(31, (ni * 2) + 2), ws.Cells(31, (ni * 2) + 3))
        nc.Merge
        nc.Value = ">> " & navItems(ni, 0)
        nc.Interior.Color = COL_WHITE
        nc.Font.Bold = True: nc.Font.Size = 8: nc.Font.Color = RGB(0, 70, 170)
        nc.Font.Underline = xlUnderlineStyleSingle
        nc.HorizontalAlignment = xlCenter: nc.VerticalAlignment = xlCenter
        nc.Borders.LineStyle = xlContinuous
        ws.Hyperlinks.Add Anchor:=nc, Address:="", SubAddress:=navItems(ni, 1), _
            TextToDisplay:=">> " & navItems(ni, 0)
    Next ni

    ' Row 32: secondary navigation — analysis/charts sheets
    ws.Rows(32).RowHeight = 16
    Dim navItems2(3, 1) As String
    navItems2(0, 0) = "CHARTS":          navItems2(0, 1) = "#CHARTS!A1"
    navItems2(1, 0) = "DATA_QUALITY":    navItems2(1, 1) = "#DATA_QUALITY!A1"
    navItems2(2, 0) = "ACTION_HRP":      navItems2(2, 1) = "#ACTION_HRP!A1"
    navItems2(3, 0) = "ACTION_PACKED":   navItems2(3, 1) = "#ACTION_PACKED!A1"

    Dim ni2 As Integer
    For ni2 = 0 To 3
        Dim nc2 As Range
        Set nc2 = ws.Range(ws.Cells(32, (ni2 * 2) + 2), ws.Cells(32, (ni2 * 2) + 3))
        nc2.Merge
        nc2.Value = ">> " & navItems2(ni2, 0)
        nc2.Interior.Color = COL_WHITE
        nc2.Font.Bold = True: nc2.Font.Size = 8: nc2.Font.Color = RGB(0, 70, 170)
        nc2.Font.Underline = xlUnderlineStyleSingle
        nc2.HorizontalAlignment = xlCenter: nc2.VerticalAlignment = xlCenter
        nc2.Borders.LineStyle = xlContinuous
        ws.Hyperlinks.Add Anchor:=nc2, Address:="", SubAddress:=navItems2(ni2, 1), _
            TextToDisplay:=">> " & navItems2(ni2, 0)
    Next ni2

    ' Row 33: spacer
    ws.Rows(33).RowHeight = 8

    ' ── Row 34: CONTROLS section ──────────────────────────────────────────────
    ws.Rows(34).RowHeight = 20
    Call MakeSectionHeader(ws, "A34:N34", "  CONTROLS  (VBA macro buttons — workbook must be saved as .xlsm)", COL_HEADER)

    ' Row 35: macro-enable instructions
    ws.Rows(35).RowHeight = 28
    With ws.Range("A35:N35")
        .Merge
        .Value = Chr(9888) & "  To use the buttons below: save this file as .xlsm (File > Save As > Excel Macro-Enabled Workbook)." & _
                 "  Enable macros when prompted on next open.  Alternatively run CreateKPIWorkbook.bas from Developer > Visual Basic."
        .Font.Italic = True: .Font.Size = 8: .Font.Color = RGB(155, 100, 0)
        .Interior.Color = RGB(255, 244, 204)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Row 36: button label cells (actual Form Buttons are placed over these cells by AddButtons)
    ws.Rows(36).RowHeight = 18
    ws.Rows(37).RowHeight = 24

    Dim btnCols() As Integer: ReDim btnCols(2)
    btnCols(0) = 2: btnCols(1) = 6: btnCols(2) = 10
    Dim btnNames(2) As String
    btnNames(0) = "[ REFRESH ALL ]"
    btnNames(1) = "[ TAKE DAILY SNAPSHOT ]"
    btnNames(2) = "[ POPULATE ACTION SHEETS ]"
    Dim btnDescs(2) As String
    btnDescs(0) = "Recalculates all formulas & refreshes PivotTables"
    btnDescs(1) = "Appends current KPIs to the HISTORY sheet"
    btnDescs(2) = "Copies filtered data to ACTION_HRP and ACTION_PACKED"

    Dim bi As Integer
    For bi = 0 To 2
        With ws.Range(ws.Cells(36, btnCols(bi)), ws.Cells(36, btnCols(bi) + 2))
            .Merge: .Value = btnNames(bi)
            .Font.Bold = True: .Font.Size = 9: .Font.Color = COL_HEADER
            .Interior.Color = RGB(255, 244, 204)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        With ws.Range(ws.Cells(37, btnCols(bi)), ws.Cells(37, btnCols(bi) + 2))
            .Merge: .Value = btnDescs(bi)
            .Font.Italic = True: .Font.Size = 8: .Font.Color = RGB(102, 102, 102)
            .HorizontalAlignment = xlCenter: .WrapText = True
        End With
    Next bi

    ws.Columns("A:A").ColumnWidth = 2
    ws.Columns("B:N").ColumnWidth = 15
    ws.Tab.Color = RGB(0, 112, 192)
End Sub

' Helper: full-width section header bar
Private Sub MakeSectionHeader(ws As Worksheet, addr As String, _
                               text As String, bgColor As Long)
    With ws.Range(addr)
        .Merge
        .Value = text
        .Font.Bold = True: .Font.Size = 11: .Font.Color = COL_WHITE
        .Interior.Color = bgColor
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
End Sub

' Helper: Excel-style KPI card (4 rows: 1 title bar + 3 value rows)
'   startRow : first row of the card (the title bar row)
'   startCol : first column of the card (card is 3 columns wide)
'   accentColor : the accent colour used for the title bar, number text, and border
Private Sub MakeKPICard(ws As Worksheet, startRow As Long, startCol As Integer, _
                         title As String, formula As String, accentColor As Long)
    Dim endCol As Integer: endCol = startCol + 2

    ' Title bar (1 row)
    ws.Rows(startRow).RowHeight = 18
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, endCol))
        .Merge
        .Value = title
        .Font.Bold = True: .Font.Size = 9: .Font.Color = COL_WHITE
        .Interior.Color = accentColor
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ' Value area (3 rows merged, white background, large coloured number)
    ws.Rows(startRow + 1).RowHeight = 42
    ws.Rows(startRow + 2).RowHeight = 8
    ws.Rows(startRow + 3).RowHeight = 8
    With ws.Range(ws.Cells(startRow + 1, startCol), ws.Cells(startRow + 3, endCol))
        .Merge
        .Formula = formula
        .Font.Bold = True: .Font.Size = 24: .Font.Color = accentColor
        .Interior.Color = COL_WHITE
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ' Accent-coloured thin border around the whole card
    Dim cardRng As Range
    Set cardRng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 3, endCol))
    With cardRng.Borders(xlEdgeLeft):   .LineStyle = xlContinuous: .Color = accentColor: End With
    With cardRng.Borders(xlEdgeRight):  .LineStyle = xlContinuous: .Color = accentColor: End With
    With cardRng.Borders(xlEdgeTop):    .LineStyle = xlContinuous: .Color = accentColor: End With
    With cardRng.Borders(xlEdgeBottom): .LineStyle = xlContinuous: .Color = accentColor: End With
End Sub

' Build a simple pivot-ready range to enable charting
Private Sub BuildDispatchPivotForChart(wb As Workbook, dashWs As Worksheet)
    ' Kept for backward compatibility; the Performance Summary table on the
    ' dashboard now shows the same data via INDEX formulas. This sub is no
    ' longer called from BuildDASHBOARD but is left here if needed.
End Sub

' ---- CHARTS -----------------------------------------------------------------
' Creates a dedicated CHARTS sheet with 5 pre-built charts that update as
' data is entered into the input and KPI tables.
Private Sub BuildCHARTS(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "CHARTS"

    ws.Range("A1").Value = "PERFORMANCE CHARTS  —  All charts update automatically as data is entered"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    ws.Range("A1").Font.Color = COL_HEADER

    Dim wsDaily As Worksheet: Set wsDaily = wb.Worksheets("T_DISPATCH_DAILY")
    Dim wsKPI   As Worksheet: Set wsKPI   = wb.Worksheets("T_DISPATCH_KPI")
    Dim wsHist  As Worksheet: Set wsHist  = wb.Worksheets("HISTORY")

    ' ── Chart 1: Daily Performance % Trend (Line) ─────────────────────────────
    Dim ch1 As ChartObject
    Set ch1 = ws.ChartObjects.Add(Left:=10, Top:=50, Width:=420, Height:=250)
    With ch1.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Daily Dispatch Performance % Trend"
        ' Add series manually so we can use full-column references (auto-expand)
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values  = wsDaily.Range("E:E")
        .SeriesCollection(1).XValues = wsDaily.Range("A:A")
        .SeriesCollection(1).Name    = "Performance %"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Performance %"
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .HasLegend = False
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 112, 192)
        .SeriesCollection(1).Format.Line.Weight = 2
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
        .SeriesCollection(1).MarkerSize  = 5
    End With

    ' ── Chart 2: Shipped vs Expected Cartons per Shift (Clustered Column) ────
    Dim ch2 As ChartObject
    Set ch2 = ws.ChartObjects.Add(Left:=450, Top:=50, Width:=420, Height:=250)
    With ch2.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Shipped vs Expected Cartons per Shift"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values  = wsKPI.Range("C2:C1048576")
        .SeriesCollection(1).XValues = wsKPI.Range("B2:B1048576")
        .SeriesCollection(1).Name    = "Shipped"
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values  = wsKPI.Range("F2:F1048576")
        .SeriesCollection(2).XValues = wsKPI.Range("B2:B1048576")
        .SeriesCollection(2).Name    = "Expected"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Cartons"
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 112, 192)
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(80, 200, 120)
    End With

    ' ── Chart 3: HRP Open & Packed Overdue Trend (Clustered Column) ─────────
    Dim ch3 As ChartObject
    Set ch3 = ws.ChartObjects.Add(Left:=10, Top:=330, Width:=420, Height:=250)
    With ch3.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "HRP Open Items & Packed Overdue — Historical Trend"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values  = wsHist.Range("I:I")
        .SeriesCollection(1).XValues = wsHist.Range("B:B")
        .SeriesCollection(1).Name    = "HRP Open"
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Values  = wsHist.Range("J:J")
        .SeriesCollection(2).XValues = wsHist.Range("B:B")
        .SeriesCollection(2).Name    = "Packed Overdue"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Count"
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = COL_RED
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = COL_AMBER
    End With

    ' ── Chart 4: Daily Staff Count Trend (Line) ──────────────────────────────
    Dim ch4 As ChartObject
    Set ch4 = ws.ChartObjects.Add(Left:=450, Top:=330, Width:=420, Height:=250)
    With ch4.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Daily Staff Count Trend"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values  = wsDaily.Range("C:C")
        .SeriesCollection(1).XValues = wsDaily.Range("A:A")
        .SeriesCollection(1).Name    = "Total Staff"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Staff Count"
        .HasLegend = False
        .SeriesCollection(1).Format.Line.ForeColor.RGB = COL_GREEN
        .SeriesCollection(1).Format.Line.Weight = 2
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleSquare
        .SeriesCollection(1).MarkerSize  = 5
    End With

    ' ── Chart 5: Audit Performance % Trend (Line) ────────────────────────────
    Dim ch5 As ChartObject
    Set ch5 = ws.ChartObjects.Add(Left:=10, Top:=610, Width:=420, Height:=250)
    With ch5.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Audit Compliance % per Shift"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values  = wsKPI.Range("K:K")
        .SeriesCollection(1).XValues = wsKPI.Range("A:A")
        .SeriesCollection(1).Name    = "Audit %"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Audit %"
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .HasLegend = False
        .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(112, 173, 71)
        .SeriesCollection(1).Format.Line.Weight = 2
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleDiamond
        .SeriesCollection(1).MarkerSize  = 5
    End With

    ws.Tab.Color = RGB(0, 112, 192)
End Sub

'================================================================================
' UTILITY ROUTINES
'================================================================================

Private Sub ReorderSheets(wb As Workbook)
    Dim order() As String
    order = Split("DASHBOARD,CHARTS,CONFIG,IN_PACKED,IN_HRP,IN_SHIPPED_LPNS,IN_STAFFING,IN_TARGETS_DAILY,IN_AUDIT_LOG,T_DISPATCH_KPI,T_DISPATCH_DAILY,ACTION_HRP,ACTION_PACKED,HISTORY,DATA_QUALITY", ",")
    Dim i As Integer
    For i = 0 To UBound(order)
        ' On Error Resume Next suppresses errors for sheets that don't exist yet;
        ' the loop continues gracefully so no valid sheet move is blocked.
        On Error Resume Next
        wb.Worksheets(order(i)).Move Before:=wb.Worksheets(i + 1)
        On Error GoTo 0
    Next i
End Sub

'================================================================================
' BUTTON MACROS (called by buttons added on DASHBOARD)
'================================================================================

Private Sub AddButtons(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets("DASHBOARD")

    ' Use Rows(36).Top to get the exact pixel position of row 36 dynamically,
    ' so this remains correct if any row heights above it change.
    Dim btnTop As Double
    btnTop = ws.Rows(36).Top
    Const BTN_H As Long = 30

    ' Refresh All button
    Dim btn1 As Button
    Set btn1 = ws.Buttons.Add(10, btnTop, 140, BTN_H)
    btn1.Caption = "REFRESH ALL"
    btn1.OnAction = "RefreshAll"
    btn1.Font.Bold = True

    ' Take Daily Snapshot button
    Dim btn2 As Button
    Set btn2 = ws.Buttons.Add(170, btnTop, 200, BTN_H)
    btn2.Caption = "TAKE DAILY SNAPSHOT"
    btn2.OnAction = "TakeDailySnapshot"
    btn2.Font.Bold = True

    ' Populate Action Sheets button
    Dim btn3 As Button
    Set btn3 = ws.Buttons.Add(390, btnTop, 200, BTN_H)
    btn3.Caption = "POPULATE ACTION SHEETS"
    btn3.OnAction = "PopulateActionSheets"
    btn3.Font.Bold = True
End Sub

'================================================================================
' PUBLIC MACROS (called by buttons)
'================================================================================

Public Sub RefreshAll()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic

    ' Refresh all PivotTables
    Dim ws As Worksheet
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    ' Force recalculation
    Application.Calculate

    ' Update last refreshed timestamp on DASHBOARD (formula is in the D2:G2 merged cell)
    On Error Resume Next
    ThisWorkbook.Worksheets("DASHBOARD").Range("D2").Formula = _
        "=TEXT(NOW(),""dd/mm/yyyy hh:mm"")"
    On Error GoTo 0

    Application.ScreenUpdating = True
    MsgBox "All data refreshed successfully.", vbInformation, "Refresh Complete"
End Sub

Public Sub TakeDailySnapshot()
    Dim wbSrc As Workbook
    Set wbSrc = ThisWorkbook

    Dim wsHist As Worksheet
    On Error Resume Next
    Set wsHist = wbSrc.Worksheets("HISTORY")
    On Error GoTo 0
    If wsHist Is Nothing Then
        MsgBox "HISTORY sheet not found.", vbCritical
        Exit Sub
    End If

    Dim wsSrc As Worksheet
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets("T_DISPATCH_KPI")
    On Error GoTo 0
    If wsSrc Is Nothing Then
        MsgBox "T_DISPATCH_KPI sheet not found.", vbCritical
        Exit Sub
    End If

    ' Find last row in HISTORY
    Dim histTbl As ListObject
    Set histTbl = wsHist.ListObjects("tblHistory")

    ' Iterate each row in T_DISPATCH_KPI
    Dim kpiTbl As ListObject
    Set kpiTbl = wsSrc.ListObjects("tblDispatchKPI")

    Dim snapTime As Date
    snapTime = Now()

    Dim r As Long
    Dim newRow As ListRow
    For r = 1 To kpiTbl.ListRows.Count
        Set newRow = histTbl.ListRows.Add
        newRow.Range(1).Value  = snapTime
        newRow.Range(2).Value  = kpiTbl.ListColumns("BusinessDate").DataBodyRange(r).Value
        newRow.Range(3).Value  = kpiTbl.ListColumns("ShiftName").DataBodyRange(r).Value
        newRow.Range(4).Value  = kpiTbl.ListColumns("ShippedCartons").DataBodyRange(r).Value
        newRow.Range(5).Value  = kpiTbl.ListColumns("TotalStaff").DataBodyRange(r).Value
        newRow.Range(6).Value  = kpiTbl.ListColumns("ExpectedCartons").DataBodyRange(r).Value
        newRow.Range(7).Value  = kpiTbl.ListColumns("PerformancePct").DataBodyRange(r).Value
        newRow.Range(8).Value  = kpiTbl.ListColumns("RAG").DataBodyRange(r).Value
        newRow.Range(9).Value  = Application.WorksheetFunction.CountIf( _
            ThisWorkbook.Worksheets("IN_HRP").ListObjects("tblHRP").ListColumns("IncludeInHRP").DataBodyRange, True)
        newRow.Range(10).Value = Application.WorksheetFunction.CountIf( _
            ThisWorkbook.Worksheets("IN_PACKED").ListObjects("tblPacked").ListColumns("ActionFlag").DataBodyRange, True)
        ' AuditCount — looked up by column name so the position is always correct
        On Error Resume Next
        Dim auditColIdx As Integer
        auditColIdx = wsHist.ListObjects("tblHistory").ListColumns("AuditCount").Index
        If auditColIdx > 0 Then
            newRow.Range(auditColIdx).Value = kpiTbl.ListColumns("AuditCount").DataBodyRange(r).Value
        End If
        On Error GoTo 0
    Next r

    MsgBox "Snapshot saved: " & kpiTbl.ListRows.Count & " row(s) added to HISTORY.", _
           vbInformation, "Snapshot Complete"
End Sub

Public Sub PopulateActionSheets()
    Application.ScreenUpdating = False

    ' --- ACTION_HRP ---
    Dim wsHRP As Worksheet
    Dim wsActHRP As Worksheet
    Set wsHRP = ThisWorkbook.Worksheets("IN_HRP")
    Set wsActHRP = ThisWorkbook.Worksheets("ACTION_HRP")

    Dim tblHRP As ListObject
    Set tblHRP = wsHRP.ListObjects("tblHRP")

    ' Clear existing data rows (keep header)
    wsActHRP.Rows("2:10000").ClearContents

    Dim hrpCols() As String
    hrpCols = Split("CATEGORY,CartonID,STORE_NUMBER,STORE_NAME,DAYS_SINCE_AVAILABLE_FOR_CITY,STORE_ACK_STATUS,UNITS,COST_VALUE", ",")

    Dim destRow As Long
    destRow = 2
    Dim r As Long
    For r = 1 To tblHRP.ListRows.Count
        If tblHRP.ListColumns("IncludeInHRP").DataBodyRange(r).Value = True Then
            Dim c As Integer
            For c = 0 To UBound(hrpCols)
                wsActHRP.Cells(destRow, c + 1).Value = _
                    tblHRP.ListColumns(hrpCols(c)).DataBodyRange(r).Value
            Next c
            destRow = destRow + 1
        End If
    Next r

    ' --- ACTION_PACKED ---
    Dim wsPacked As Worksheet
    Dim wsActPck As Worksheet
    Set wsPacked = ThisWorkbook.Worksheets("IN_PACKED")
    Set wsActPck = ThisWorkbook.Worksheets("ACTION_PACKED")

    Dim tblPacked As ListObject
    Set tblPacked = wsPacked.ListObjects("tblPacked")

    wsActPck.Rows("2:10000").ClearContents

    Dim pckCols() As String
    pckCols = Split("LPN,STORE,DIVISION,AgeDays,PackedStatus,UNITS", ",")

    ' Collect rows into an array for sorting
    Dim rowData() As Variant
    Dim rowCount As Long
    rowCount = 0
    For r = 1 To tblPacked.ListRows.Count
        If tblPacked.ListColumns("ActionFlag").DataBodyRange(r).Value = True Then
            rowCount = rowCount + 1
        End If
    Next r

    If rowCount > 0 Then
        ReDim rowData(1 To rowCount, 1 To UBound(pckCols) + 1)
        Dim idx As Long
        idx = 1
        For r = 1 To tblPacked.ListRows.Count
            If tblPacked.ListColumns("ActionFlag").DataBodyRange(r).Value = True Then
                For c = 0 To UBound(pckCols)
                    rowData(idx, c + 1) = tblPacked.ListColumns(pckCols(c)).DataBodyRange(r).Value
                Next c
                idx = idx + 1
            End If
        Next r

        ' Bubble sort descending by AgeDays (column 4 = index 4)
        Dim i As Long, j As Long
        Dim tmpVal As Variant
        For i = 1 To rowCount - 1
            For j = 1 To rowCount - i
                If rowData(j, 4) < rowData(j + 1, 4) Then
                    For c = 1 To UBound(pckCols) + 1
                        tmpVal = rowData(j, c)
                        rowData(j, c) = rowData(j + 1, c)
                        rowData(j + 1, c) = tmpVal
                    Next c
                End If
            Next j
        Next i

        ' Write to ACTION_PACKED
        For r = 1 To rowCount
            For c = 1 To UBound(pckCols) + 1
                wsActPck.Cells(r + 1, c).Value = rowData(r, c)
            Next c
        Next r
    End If

    Application.ScreenUpdating = True
    MsgBox "Action sheets updated." & vbCrLf & _
           "ACTION_HRP: " & (destRow - 2) & " row(s)" & vbCrLf & _
           "ACTION_PACKED: " & rowCount & " row(s)", vbInformation, "Populate Complete"
End Sub

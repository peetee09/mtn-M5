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
    Call BuildT_DISPATCH_KPI(wb)
    Call BuildT_DISPATCH_DAILY(wb)
    Call BuildACTION_HRP(wb)
    Call BuildACTION_PACKED(wb)
    Call BuildHISTORY(wb)
    Call BuildDATA_QUALITY(wb)
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

    MsgBox "KPI Workbook created successfully!" & vbCrLf & _
           "Unprotect password: " & UNPROTECT_PW, vbInformation, "Done"
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

    ' Area validation from CONFIG
    With tbl.ListColumns("Area").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="=tblConfig_Areas[AreaName]"
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

' ---- T_DISPATCH_KPI ---------------------------------------------------------
Private Sub BuildT_DISPATCH_KPI(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "T_DISPATCH_KPI"

    Dim cols() As String
    cols = Split("BusinessDate,ShiftName,ShippedCartons,TotalStaff,TargetPerPerson,ExpectedCartons,PerformancePct,RAG", ",")

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
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:H2"), , xlYes)
    tbl.Name = "tblDispatchKPI"
    tbl.TableStyle = "TableStyleMedium2"

    ' Calculated columns
    tbl.ListColumns("ShippedCartons").DataBodyRange(1).Formula = _
        "=COUNTIFS(tblShipped[BusinessDate],[@BusinessDate],tblShipped[ShiftName],[@ShiftName])"
    tbl.ListColumns("TotalStaff").DataBodyRange(1).Formula = _
        "=SUMIFS(tblStaffing[StaffAvailable],tblStaffing[BusinessDate],[@BusinessDate],tblStaffing[ShiftName],[@ShiftName])"
    ' TargetPerPerson: use IFERROR + INDEX/MATCH (Ctrl+Shift+Enter not needed inside structured refs)
    tbl.ListColumns("TargetPerPerson").DataBodyRange(1).Formula = _
        "=IFERROR(INDEX(tblTargetsDaily[TargetPerPersonPerShift],MATCH(1,(tblTargetsDaily[BusinessDate]=[@BusinessDate])*(tblTargetsDaily[ShiftName]=[@ShiftName]),0)),0)"
    ' Array formula version for Excel 2016
    tbl.ListColumns("TargetPerPerson").DataBodyRange(1).FormulaArray = _
        "=IFERROR(INDEX(tblTargetsDaily[TargetPerPersonPerShift],MATCH(1,(tblTargetsDaily[BusinessDate]=[@BusinessDate])*(tblTargetsDaily[ShiftName]=[@ShiftName]),0)),0)"

    tbl.ListColumns("ExpectedCartons").DataBodyRange(1).Formula = _
        "=[@TotalStaff]*[@TargetPerPerson]"
    tbl.ListColumns("PerformancePct").DataBodyRange(1).Formula = _
        "=IFERROR([@ShippedCartons]/[@ExpectedCartons],0)"
    tbl.ListColumns("RAG").DataBodyRange(1).Formula = _
        "=IF([@PerformancePct]>=1,""Green"",IF([@PerformancePct]>=0.9,""Amber"",""Red""))"

    ' Format columns
    ws.Columns("A:A").NumberFormat = "dd/mm/yyyy"
    ws.Columns("G:G").NumberFormat = "0.0%"

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
    ws.Columns("A:H").AutoFit
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
    cols = Split("SnapshotTimestamp,BusinessDate,ShiftName,ShippedCartons,TotalStaff,ExpectedCartons,PerformancePct,RAG,HRP_OpenCount,Packed_OverdueCount", ",")

    Dim col As Integer
    For col = 0 To UBound(cols)
        ws.Cells(1, col + 1).Value = cols(col)
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = COL_HEADER
        ws.Cells(1, col + 1).Font.Color = COL_WHITE
    Next col

    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:J1"), , xlYes)
    tbl.Name = "tblHistory"
    tbl.TableStyle = "TableStyleMedium2"
    tbl.ShowAutoFilterDropDown = True

    ws.Columns("A:B").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Columns("G:G").NumberFormat = "0.0%"
    ws.Columns("A:J").AutoFit
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
        ws.Cells(r + 4, 2).Value = "'" & checks(r, 1)   ' store formula text
        ws.Cells(r + 4, 3).Formula = checks(r, 1)        ' also evaluate
        ws.Cells(r + 4, 4).Formula = _
            "=IF(C" & (r + 4) & "=0,""OK"",IF(C" & (r + 4) & ">0,""WARNING"",""OK""))"

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

    ' Background
    ws.Cells.Interior.Color = RGB(30, 30, 30)
    ws.Cells.Font.Color = COL_WHITE

    ' Title
    With ws.Range("A1:N1")
        .Merge
        .Value = "WAREHOUSE / DISPATCH KPI DASHBOARD"
        .Font.Bold = True
        .Font.Size = 18
        .Font.Color = COL_WHITE
        .Interior.Color = COL_HEADER
        .HorizontalAlignment = xlCenter
    End With

    ws.Range("A2").Value = "Last Refreshed:"
    ws.Range("B2").Formula = "=TEXT(NOW(),""dd/mm/yyyy hh:mm"")"
    ws.Range("B2").Font.Bold = True

    ' --- KPI Cards (rows 4-9) ------------------------------------------------
    ' Card: HRP Open Count
    Call MakeKPICard(ws, "B4:D7", "HRP OPEN ITEMS", _
        "=COUNTIF(tblHRP[IncludeInHRP],TRUE)", COL_RED)
    ' Card: Packed Overdue
    Call MakeKPICard(ws, "F4:H7", "PACKED OVERDUE", _
        "=COUNTIF(tblPacked[ActionFlag],TRUE)", COL_RED)
    ' Card: Dispatch Performance Today
    Call MakeKPICard(ws, "J4:L7", "DISPATCH PERF TODAY", _
        "=IFERROR(TEXT(SUMIF(tblDispatchDaily[BusinessDate],TODAY(),tblDispatchDaily[DailyPerformancePct]),""0.0%""),""N/A"")", _
        COL_GREEN)
    ' Card: Duplicate LPNs
    Call MakeKPICard(ws, "B9:D12", "DUPLICATE LPNs", _
        "=COUNTIF(tblShipped[DupFlag],TRUE)", COL_AMBER)
    ' Card: Staff Today
    Call MakeKPICard(ws, "F9:H12", "STAFF TODAY", _
        "=SUMIF(tblDispatchKPI[BusinessDate],TODAY(),tblDispatchKPI[TotalStaff])", _
        RGB(0, 112, 192))
    ' Card: Cartons Shipped Today
    Call MakeKPICard(ws, "J9:L12", "CARTONS SHIPPED TODAY", _
        "=SUMIF(tblDispatchKPI[BusinessDate],TODAY(),tblDispatchKPI[ShippedCartons])", _
        RGB(0, 112, 192))

    ' --- Instruction area for charts -----------------------------------------
    ws.Range("A14").Value = "CHARTS AREA"
    ws.Range("A14").Font.Bold = True
    ws.Range("A14").Font.Size = 12
    ws.Range("A14").Font.Color = COL_WHITE

    ws.Range("A15").Value = "After adding data to the input sheets, use the 'Refresh All' button to update PivotTables and charts."
    ws.Range("A15").Font.Italic = True
    ws.Range("A15").Font.Color = RGB(200, 200, 200)

    ' Slicer / filter guide
    ws.Range("A17").Value = "FILTER GUIDE:"
    ws.Range("A17").Font.Bold = True
    ws.Range("A18").Value = "Use Excel AutoFilter on T_DISPATCH_KPI for ShiftName, BusinessDate, or RAG filters."
    ws.Range("A19").Value = "For pivot-based slicers: Insert PivotTable from tblDispatchKPI on any blank sheet, then Insert > Slicer."

    ' Build summary PivotTable source area for chart
    Call BuildDispatchPivotForChart(wb, ws)

    ws.Columns("A:N").ColumnWidth = 14
    ws.Tab.Color = RGB(0, 112, 192)
End Sub

' Helper: create a KPI card
Private Sub MakeKPICard(ws As Worksheet, addr As String, title As String, _
                         formula As String, bgColor As Long)
    Dim rng As Range
    Set rng = ws.Range(addr)
    rng.Merge
    With rng
        .Interior.Color = bgColor
        .Font.Color = COL_WHITE
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Formula = formula
    End With

    ' Title one row above
    Dim titleRow As Long
    titleRow = rng.Row - 1
    Dim titleRange As Range
    Set titleRange = ws.Range(ws.Cells(titleRow, rng.Column), _
                              ws.Cells(titleRow, rng.Column + rng.Columns.Count - 1))
    titleRange.Merge
    titleRange.Value = title
    titleRange.Font.Bold = True
    titleRange.Font.Size = 9
    titleRange.Font.Color = RGB(200, 200, 200)
    titleRange.HorizontalAlignment = xlCenter
    titleRange.Interior.Color = RGB(50, 50, 50)
End Sub

' Build a simple pivot-ready range to enable charting
Private Sub BuildDispatchPivotForChart(wb As Workbook, dashWs As Worksheet)
    ' Place a small summary table starting at row 21 for charting
    dashWs.Range("A21").Value = "DATE"
    dashWs.Range("B21").Value = "SHIFT"
    dashWs.Range("C21").Value = "SHIPPED"
    dashWs.Range("D21").Value = "EXPECTED"
    dashWs.Range("E21").Value = "PERF%"

    dashWs.Range("A21:E21").Font.Bold = True
    dashWs.Range("A21:E21").Interior.Color = COL_HEADER
    dashWs.Range("A21:E21").Font.Color = COL_WHITE

    ' Row 22: reference from tblDispatchKPI row 1 (dynamic would need VBA refresh)
    dashWs.Range("A22").Formula = "=IF(ROWS(tblDispatchKPI[BusinessDate])>0,INDEX(tblDispatchKPI[BusinessDate],1),"""")"
    dashWs.Range("B22").Formula = "=IF(ROWS(tblDispatchKPI[ShiftName])>0,INDEX(tblDispatchKPI[ShiftName],1),"""")"
    dashWs.Range("C22").Formula = "=IF(ROWS(tblDispatchKPI[ShippedCartons])>0,INDEX(tblDispatchKPI[ShippedCartons],1),0)"
    dashWs.Range("D22").Formula = "=IF(ROWS(tblDispatchKPI[ExpectedCartons])>0,INDEX(tblDispatchKPI[ExpectedCartons],1),0)"
    dashWs.Range("E22").Formula = "=IF(ROWS(tblDispatchKPI[PerformancePct])>0,INDEX(tblDispatchKPI[PerformancePct],1),0)"
    dashWs.Range("E22").NumberFormat = "0.0%"

    dashWs.Range("A22").NumberFormat = "dd/mm/yyyy"
End Sub

'================================================================================
' UTILITY ROUTINES
'================================================================================

Private Sub ReorderSheets(wb As Workbook)
    Dim order() As String
    order = Split("DASHBOARD,CONFIG,IN_PACKED,IN_HRP,IN_SHIPPED_LPNS,IN_STAFFING,IN_TARGETS_DAILY,T_DISPATCH_KPI,T_DISPATCH_DAILY,ACTION_HRP,ACTION_PACKED,HISTORY,DATA_QUALITY", ",")
    Dim i As Integer
    For i = 0 To UBound(order)
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

    ' Refresh All button
    Dim btn1 As Button
    Set btn1 = ws.Buttons.Add(10, 550, 140, 30)
    btn1.Caption = "REFRESH ALL"
    btn1.OnAction = "RefreshAll"
    btn1.Font.Bold = True

    ' Take Daily Snapshot button
    Dim btn2 As Button
    Set btn2 = ws.Buttons.Add(170, 550, 200, 30)
    btn2.Caption = "TAKE DAILY SNAPSHOT"
    btn2.OnAction = "TakeDailySnapshot"
    btn2.Font.Bold = True

    ' Populate Action Sheets button
    Dim btn3 As Button
    Set btn3 = ws.Buttons.Add(390, 550, 200, 30)
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

    ' Update last refreshed timestamp on DASHBOARD
    On Error Resume Next
    ThisWorkbook.Worksheets("DASHBOARD").Range("B2").Formula = _
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

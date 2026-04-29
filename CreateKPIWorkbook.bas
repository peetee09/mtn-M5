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
    Call BuildDAILY_ENTRY(wb)
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

    ' ── Save as macro-enabled workbook (.xlsm) ────────────────────────────────
    ' This ensures VBA code and Form Button assignments survive the next open.
    Dim xlsmPath As String
    Dim dotPos   As Long

    If LCase(Right(wb.Name, 5)) = ".xlsm" Then
        ' Already a macro-enabled file — just save in place.
        wb.Save
        MsgBox "KPI Workbook created successfully!" & vbCrLf & _
               "Unprotect password: " & UNPROTECT_PW & vbCrLf & vbCrLf & _
               "File saved as .xlsm — macros are embedded and buttons will work on next open.", _
               vbInformation, "Done"
    ElseIf wb.Path = "" Then
        ' Workbook has never been saved (e.g. "Book1" with no path).
        ' Prompt the user to choose a location.
        Dim savePath As Variant
        savePath = Application.GetSaveAsFilename( _
            InitialFileName:="KPI_Workbook.xlsm", _
            FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm", _
            Title:="Save the KPI Workbook as .xlsm")
        If savePath = False Or savePath = "" Then
            MsgBox "Save cancelled. The workbook was NOT saved as .xlsm." & vbCrLf & _
                   "Macro buttons will not persist until you save the file as .xlsm.", _
                   vbExclamation, "Save Cancelled"
        Else
            Application.DisplayAlerts = False
            On Error GoTo SaveErr
            wb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
            Application.DisplayAlerts = True
            On Error GoTo 0
            MsgBox "KPI Workbook created successfully!" & vbCrLf & _
                   "Unprotect password: " & UNPROTECT_PW & vbCrLf & vbCrLf & _
                   "File saved as .xlsm — macros are embedded and buttons will work on next open.", _
                   vbInformation, "Done"
        End If
    Else
        ' Workbook has been saved but is not already .xlsm — convert in place.
        dotPos = InStrRev(wb.FullName, ".")
        If dotPos > 0 Then
            xlsmPath = Left(wb.FullName, dotPos - 1) & ".xlsm"
        Else
            xlsmPath = wb.FullName & ".xlsm"
        End If
        Application.DisplayAlerts = False
        On Error GoTo SaveErr
        wb.SaveAs Filename:=xlsmPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Application.DisplayAlerts = True
        On Error GoTo 0
        MsgBox "KPI Workbook created successfully!" & vbCrLf & _
               "Unprotect password: " & UNPROTECT_PW & vbCrLf & vbCrLf & _
               "File saved as .xlsm — macros are embedded and buttons will work on next open.", _
               vbInformation, "Done"
    End If
    Exit Sub

SaveErr:
    Application.DisplayAlerts = True
    On Error GoTo 0
    MsgBox "KPI Workbook was built, but could not be saved as .xlsm." & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Please use File > Save As and choose 'Excel Macro-Enabled Workbook (*.xlsm)' manually.", _
           vbCritical, "Save Error"
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
    Dim navItems2(4, 1) As String
    navItems2(0, 0) = "CHARTS":          navItems2(0, 1) = "#CHARTS!A1"
    navItems2(1, 0) = "DATA_QUALITY":    navItems2(1, 1) = "#DATA_QUALITY!A1"
    navItems2(2, 0) = "ACTION_HRP":      navItems2(2, 1) = "#ACTION_HRP!A1"
    navItems2(3, 0) = "ACTION_PACKED":   navItems2(3, 1) = "#ACTION_PACKED!A1"
    navItems2(4, 0) = "DAILY_ENTRY":     navItems2(4, 1) = "#DAILY_ENTRY!A1"

    Dim ni2 As Integer
    For ni2 = 0 To 4
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

    ' Row 38: FILL DATA button placeholder (actual button added by AddButtons)
    ws.Rows(38).RowHeight = 24
    With ws.Range("B38:D38")
        .Merge: .Value = "[ FILL DATA ]"
        .Font.Bold = True: .Font.Size = 9: .Font.Color = COL_HEADER
        .Interior.Color = RGB(200, 230, 201)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    ws.Rows(39).RowHeight = 20
    With ws.Range("B39:N39")
        .Merge
        .Value = "Open the daily entry form to record shift data for a specific day — the dashboard calculates all KPIs for the day, week, and month."
        .Font.Italic = True: .Font.Size = 8: .Font.Color = RGB(102, 102, 102)
        .HorizontalAlignment = xlCenter: .WrapText = True
    End With

    ' Row 40: spacer
    ws.Rows(40).RowHeight = 8

    ' ── Row 41: WEEKLY SUMMARY section ────────────────────────────────────────
    ws.Rows(41).RowHeight = 20
    Call MakeSectionHeader(ws, "A41:N41", "  WEEKLY SUMMARY  (Monday – today)", RGB(70, 130, 180))

    ' Rows 42-45: Weekly KPI cards
    ' Use tblDispatchDaily (one row per BusinessDate) so repeated snapshots in
    ' tblHistory do not double-count weekly totals or skew the average.
    ' WEEKDAY(TODAY(),2) returns 1=Mon … 7=Sun, so TODAY()-WEEKDAY(TODAY(),2)+1 = this Monday
    ' Use tblDispatchDaily here because tblHistory is append-only and can contain
    ' multiple snapshots for the same BusinessDate/Shift, which would double-count
    ' weekly summary rollups.
    Call MakeKPICard(ws, 42, 2, "CARTONS THIS WEEK", _
        "=SUMIFS(tblDispatchDaily[TotalShipped],tblDispatchDaily[BusinessDate],"">=""&(TODAY()-WEEKDAY(TODAY(),2)+1),tblDispatchDaily[BusinessDate],""<=""&TODAY())", _
        RGB(70, 130, 180))
    Call MakeKPICard(ws, 42, 6, "AVG PERF % (WEEK)", _
        "=IFERROR(TEXT(AVERAGEIFS(tblDispatchDaily[DailyPerformancePct],tblDispatchDaily[BusinessDate],"">=""&(TODAY()-WEEKDAY(TODAY(),2)+1),tblDispatchDaily[BusinessDate],""<=""&TODAY()),""0.0%""),""N/A"")", _
        RGB(70, 130, 180))
    Call MakeKPICard(ws, 42, 10, "STAFF COUNT THIS WEEK", _
        "=SUMIFS(tblDispatchDaily[TotalStaff],tblDispatchDaily[BusinessDate],"">=""&(TODAY()-WEEKDAY(TODAY(),2)+1),tblDispatchDaily[BusinessDate],""<=""&TODAY())", _
        RGB(70, 130, 180))

    ' Row 46: spacer
    ws.Rows(46).RowHeight = 8

    ' ── Row 47: MONTHLY SUMMARY section ───────────────────────────────────────
    ws.Rows(47).RowHeight = 20
    Call MakeSectionHeader(ws, "A47:N47", "  MONTHLY SUMMARY  (1st of month – today)", RGB(112, 48, 160))

    ' Rows 48-51: Monthly KPI cards
    ' Use tblDispatchDaily (one row per BusinessDate) so repeated snapshots in
    ' tblHistory do not double-count monthly totals or skew the average.
    Call MakeKPICard(ws, 48, 2, "CARTONS THIS MONTH", _
        "=SUMIFS(tblDispatchDaily[TotalShipped],tblDispatchDaily[BusinessDate],"">=""&DATE(YEAR(TODAY()),MONTH(TODAY()),1),tblDispatchDaily[BusinessDate],""<=""&TODAY())", _
        RGB(112, 48, 160))
    Call MakeKPICard(ws, 48, 6, "AVG PERF % (MONTH)", _
        "=IFERROR(TEXT(AVERAGEIFS(tblDispatchDaily[DailyPerformancePct],tblDispatchDaily[BusinessDate],"">=""&DATE(YEAR(TODAY()),MONTH(TODAY()),1),tblDispatchDaily[BusinessDate],""<=""&TODAY()),""0.0%""),""N/A"")", _
        RGB(112, 48, 160))
    Call MakeKPICard(ws, 48, 10, "STAFF COUNT THIS MONTH", _
        "=SUMIFS(tblDispatchDaily[TotalStaff],tblDispatchDaily[BusinessDate],"">=""&DATE(YEAR(TODAY()),MONTH(TODAY()),1),tblDispatchDaily[BusinessDate],""<=""&TODAY())", _
        RGB(112, 48, 160))

    ' Row 52: spacer
    ws.Rows(52).RowHeight = 8

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
        ' Add series manually so we can use worksheet-spanning references without including headers
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values  = wsDaily.Range("E2:E1048576")
        .SeriesCollection(1).XValues = wsDaily.Range("A2:A1048576")
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
        .SeriesCollection(1).Values  = wsKPI.Range("K2:K1048576")
        .SeriesCollection(1).XValues = wsKPI.Range("B2:B1048576")
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
    order = Split("DASHBOARD,DAILY_ENTRY,CHARTS,CONFIG,IN_PACKED,IN_HRP,IN_SHIPPED_LPNS,IN_STAFFING,IN_TARGETS_DAILY,IN_AUDIT_LOG,T_DISPATCH_KPI,T_DISPATCH_DAILY,ACTION_HRP,ACTION_PACKED,HISTORY,DATA_QUALITY", ",")
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

    ' Fill Data button (row 38 on DASHBOARD — matches the B38:D38 placeholder cell)
    Dim fillBtnTop As Double
    fillBtnTop = ws.Rows(38).Top
    Dim btn4 As Button
    Set btn4 = ws.Buttons.Add(10, fillBtnTop, 200, BTN_H)
    btn4.Caption = "FILL DATA"
    btn4.OnAction = "NavigateToDailyEntry"
    btn4.Font.Bold = True

    ' === DAILY_ENTRY sheet buttons ===
    Dim wsDE As Worksheet
    On Error Resume Next
    Set wsDE = wb.Worksheets("DAILY_ENTRY")
    On Error GoTo 0
    If Not wsDE Is Nothing Then
        Dim deBtnTop As Double
        deBtnTop = wsDE.Rows(15).Top
        Const DE_BTN_H As Long = 30

        Dim btnDE1 As Button
        Set btnDE1 = wsDE.Buttons.Add( _
            wsDE.Columns("B").Left, deBtnTop, _
            wsDE.Columns("B").Width + wsDE.Columns("C").Width, DE_BTN_H)
        btnDE1.Caption = "SUBMIT DAY"
        btnDE1.OnAction = "SubmitDailyData"
        btnDE1.Font.Bold = True

        Dim btnDE2 As Button
        Set btnDE2 = wsDE.Buttons.Add( _
            wsDE.Columns("E").Left, deBtnTop, _
            wsDE.Columns("E").Width + wsDE.Columns("F").Width, DE_BTN_H)
        btnDE2.Caption = "CLEAR FORM"
        btnDE2.OnAction = "ClearDailyForm"
        btnDE2.Font.Bold = True
    End If
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
    Dim cellVal As Variant
    For r = 1 To kpiTbl.ListRows.Count
        Set newRow = histTbl.ListRows.Add

        ' SnapshotTimestamp
        newRow.Range(1).Value = snapTime

        ' BusinessDate — guard against formula errors
        cellVal = kpiTbl.ListColumns("BusinessDate").DataBodyRange(r).Value
        newRow.Range(2).Value = IIf(IsError(cellVal), "", cellVal)

        ' ShiftName
        cellVal = kpiTbl.ListColumns("ShiftName").DataBodyRange(r).Value
        newRow.Range(3).Value = IIf(IsError(cellVal), "", cellVal)

        ' ShippedCartons
        cellVal = kpiTbl.ListColumns("ShippedCartons").DataBodyRange(r).Value
        newRow.Range(4).Value = IIf(IsError(cellVal), 0, cellVal)

        ' TotalStaff
        cellVal = kpiTbl.ListColumns("TotalStaff").DataBodyRange(r).Value
        newRow.Range(5).Value = IIf(IsError(cellVal), 0, cellVal)

        ' ExpectedCartons
        cellVal = kpiTbl.ListColumns("ExpectedCartons").DataBodyRange(r).Value
        newRow.Range(6).Value = IIf(IsError(cellVal), 0, cellVal)

        ' PerformancePct — most common source of Type Mismatch (Runtime Error 13)
        ' when the formula returns an error (e.g. #DIV/0! when no staff data)
        cellVal = kpiTbl.ListColumns("PerformancePct").DataBodyRange(r).Value
        newRow.Range(7).Value = IIf(IsError(cellVal), 0, cellVal)

        ' RAG
        cellVal = kpiTbl.ListColumns("RAG").DataBodyRange(r).Value
        newRow.Range(8).Value = IIf(IsError(cellVal), "", cellVal)

        ' HRP_OpenCount
        newRow.Range(9).Value = Application.WorksheetFunction.CountIf( _
            ThisWorkbook.Worksheets("IN_HRP").ListObjects("tblHRP").ListColumns("IncludeInHRP").DataBodyRange, True)

        ' Packed_OverdueCount
        newRow.Range(10).Value = Application.WorksheetFunction.CountIf( _
            ThisWorkbook.Worksheets("IN_PACKED").ListObjects("tblPacked").ListColumns("ActionFlag").DataBodyRange, True)

        ' AuditCount — looked up by column name so the position is always correct
        On Error Resume Next
        Dim auditColIdx As Integer
        auditColIdx = wsHist.ListObjects("tblHistory").ListColumns("AuditCount").Index
        If auditColIdx > 0 Then
            cellVal = kpiTbl.ListColumns("AuditCount").DataBodyRange(r).Value
            newRow.Range(auditColIdx).Value = IIf(IsError(cellVal), 0, cellVal)
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
    Dim c As Integer
    Dim cellVal As Variant
    For r = 1 To tblHRP.ListRows.Count
        cellVal = tblHRP.ListColumns("IncludeInHRP").DataBodyRange(r).Value
        If Not IsError(cellVal) Then
            If cellVal = True Then
                For c = 0 To UBound(hrpCols)
                    wsActHRP.Cells(destRow, c + 1).Value = _
                        tblHRP.ListColumns(hrpCols(c)).DataBodyRange(r).Value
                Next c
                destRow = destRow + 1
            End If
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
        cellVal = tblPacked.ListColumns("ActionFlag").DataBodyRange(r).Value
        If Not IsError(cellVal) Then
            If cellVal = True Then
                rowCount = rowCount + 1
            End If
        End If
    Next r

    If rowCount > 0 Then
        ReDim rowData(1 To rowCount, 1 To UBound(pckCols) + 1)
        Dim idx As Long
        idx = 1
        For r = 1 To tblPacked.ListRows.Count
            cellVal = tblPacked.ListColumns("ActionFlag").DataBodyRange(r).Value
            If Not IsError(cellVal) Then
                If cellVal = True Then
                    For c = 0 To UBound(pckCols)
                        rowData(idx, c + 1) = tblPacked.ListColumns(pckCols(c)).DataBodyRange(r).Value
                    Next c
                    idx = idx + 1
                End If
            End If
        Next r

        ' Bubble sort descending by AgeDays (column 4 = index 4)
        Dim i As Long, j As Long
        Dim tmpVal As Variant
        Dim aVal As Variant, bVal As Variant
        For i = 1 To rowCount - 1
            For j = 1 To rowCount - i
                aVal = rowData(j, 4)
                bVal = rowData(j + 1, 4)
                If Not IsError(aVal) Then
                    If Not IsError(bVal) Then
                        If aVal < bVal Then
                            For c = 1 To UBound(pckCols) + 1
                                tmpVal = rowData(j, c)
                                rowData(j, c) = rowData(j + 1, c)
                                rowData(j + 1, c) = tmpVal
                            Next c
                        End If
                    End If
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

'================================================================================
' DAILY_ENTRY SHEET BUILDER
'================================================================================

Private Sub BuildDAILY_ENTRY(wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "DAILY_ENTRY"

    ' Light-grey page background
    ws.Cells.Interior.Color = RGB(242, 242, 242)

    ' ── Row 1: Title bar ──────────────────────────────────────────────────────
    ws.Rows(1).RowHeight = 36
    With ws.Range("A1:H1")
        .Merge
        .Value = "DAILY DATA ENTRY FORM"
        .Font.Bold = True: .Font.Size = 16: .Font.Color = COL_WHITE
        .Interior.Color = COL_HEADER
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ' ── Row 2: Instructions ───────────────────────────────────────────────────
    ws.Rows(2).RowHeight = 28
    With ws.Range("A2:H2")
        .Merge
        .Value = "Fill in shift data for the day, then click SUBMIT DAY." & _
                 "  The dashboard calculates all KPIs for the day, week, and month automatically."
        .Font.Italic = True: .Font.Size = 9: .Font.Color = RGB(155, 100, 0)
        .Interior.Color = RGB(255, 244, 204)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' Row 3: Spacer
    ws.Rows(3).RowHeight = 10

    ' ── Row 4: Business Date input ────────────────────────────────────────────
    ws.Rows(4).RowHeight = 28
    ws.Range("A4").Value = "Business Date:"
    ws.Range("A4").Font.Bold = True: ws.Range("A4").Font.Size = 11
    ws.Range("A4").HorizontalAlignment = xlRight: ws.Range("A4").VerticalAlignment = xlCenter
    With ws.Range("B4:C4")
        .Merge
        .Value = Date      ' default: today
        .NumberFormat = "dd/mm/yyyy"
        .Font.Bold = True: .Font.Size = 13: .Font.Color = COL_HEADER
        .Interior.Color = COL_WHITE
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 112, 192)
    End With
    ws.Range("D4").Value = "(change if entering data for a past date)"
    ws.Range("D4").Font.Italic = True: ws.Range("D4").Font.Size = 8
    ws.Range("D4").Font.Color = RGB(128, 128, 128): ws.Range("D4").VerticalAlignment = xlCenter

    ' Row 5: Spacer
    ws.Rows(5).RowHeight = 8

    ' ── Row 6: SHIFT DATA section header ─────────────────────────────────────
    ws.Rows(6).RowHeight = 20
    With ws.Range("A6:H6")
        .Merge: .Value = "  SHIFT DATA  —  enter values for each active shift"
        .Font.Bold = True: .Font.Size = 11: .Font.Color = COL_WHITE
        .Interior.Color = RGB(46, 64, 87)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With

    ' ── Row 7: Column headers ─────────────────────────────────────────────────
    ws.Rows(7).RowHeight = 18
    Dim shiftHdrs As Variant
    shiftHdrs = Array("Shift", "Cartons Shipped", "Total Staff", "Target / Person", "Audit Count")
    Dim hi As Integer
    For hi = 0 To 4
        With ws.Cells(7, hi + 1)
            .Value = shiftHdrs(hi)
            .Font.Bold = True: .Font.Size = 10: .Font.Color = COL_WHITE
            .Interior.Color = RGB(0, 112, 192)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
    Next hi

    ' ── Row 8: Day shift ──────────────────────────────────────────────────────
    ws.Rows(8).RowHeight = 30
    With ws.Range("A8")
        .Value = "Day"
        .Font.Bold = True: .Font.Size = 12
        .Interior.Color = COL_WHITE
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 112, 192)
    End With
    Dim dc As Integer
    For dc = 2 To 5
        With ws.Cells(8, dc)
            .Value = 0
            .Font.Size = 12: .Font.Color = COL_HEADER
            .Interior.Color = COL_WHITE
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 112, 192)
        End With
    Next dc

    ' ── Row 9: Night shift ────────────────────────────────────────────────────
    ws.Rows(9).RowHeight = 30
    With ws.Range("A9")
        .Value = "Night"
        .Font.Bold = True: .Font.Size = 12
        .Interior.Color = RGB(247, 247, 247)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 112, 192)
    End With
    For dc = 2 To 5
        With ws.Cells(9, dc)
            .Value = 0
            .Font.Size = 12: .Font.Color = COL_HEADER
            .Interior.Color = RGB(247, 247, 247)
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(0, 112, 192)
        End With
    Next dc

    ' Row 10: Spacer
    ws.Rows(10).RowHeight = 8

    ' ── Row 11: EXCEPTION COUNTS section header ───────────────────────────────
    ws.Rows(11).RowHeight = 20
    With ws.Range("A11:H11")
        .Merge: .Value = "  EXCEPTION COUNTS  (optional — leave blank to use live values from data sheets)"
        .Font.Bold = True: .Font.Size = 11: .Font.Color = COL_WHITE
        .Interior.Color = RGB(192, 0, 0)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With

    ' ── Row 12: HRP Open Items ────────────────────────────────────────────────
    ws.Rows(12).RowHeight = 24
    ws.Range("A12").Value = "HRP Open Items:"
    ws.Range("A12").Font.Bold = True
    ws.Range("A12").HorizontalAlignment = xlRight: ws.Range("A12").VerticalAlignment = xlCenter
    With ws.Range("B12")
        .Value = ""
        .Interior.Color = COL_WHITE
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(192, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Range("C12").Value = "(blank = live count from IN_HRP)"
    ws.Range("C12").Font.Italic = True: ws.Range("C12").Font.Size = 8
    ws.Range("C12").Font.Color = RGB(128, 128, 128): ws.Range("C12").VerticalAlignment = xlCenter

    ' ── Row 13: Packed Overdue ────────────────────────────────────────────────
    ws.Rows(13).RowHeight = 24
    ws.Range("A13").Value = "Packed Overdue:"
    ws.Range("A13").Font.Bold = True
    ws.Range("A13").HorizontalAlignment = xlRight: ws.Range("A13").VerticalAlignment = xlCenter
    With ws.Range("B13")
        .Value = ""
        .Interior.Color = COL_WHITE
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(192, 0, 0)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Range("C13").Value = "(blank = live count from IN_PACKED)"
    ws.Range("C13").Font.Italic = True: ws.Range("C13").Font.Size = 8
    ws.Range("C13").Font.Color = RGB(128, 128, 128): ws.Range("C13").VerticalAlignment = xlCenter

    ' Row 14: Spacer
    ws.Rows(14).RowHeight = 12

    ' ── Row 15: Button placeholder rows (actual buttons added by AddButtons) ──
    ws.Rows(15).RowHeight = 30
    ws.Rows(16).RowHeight = 16

    ' Submit Day placeholder
    With ws.Range("B15:C15")
        .Merge: .Value = "[ SUBMIT DAY ]"
        .Font.Bold = True: .Font.Size = 10: .Font.Color = COL_HEADER
        .Interior.Color = RGB(200, 230, 201)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("B16:C16")
        .Merge: .Value = "Save data and update HISTORY"
        .Font.Italic = True: .Font.Size = 8: .Font.Color = RGB(102, 102, 102)
        .HorizontalAlignment = xlCenter
    End With

    ' Clear Form placeholder
    With ws.Range("E15:F15")
        .Merge: .Value = "[ CLEAR FORM ]"
        .Font.Bold = True: .Font.Size = 10: .Font.Color = COL_HEADER
        .Interior.Color = RGB(255, 235, 238)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("E16:F16")
        .Merge: .Value = "Reset all fields to zero / today"
        .Font.Italic = True: .Font.Size = 8: .Font.Color = RGB(102, 102, 102)
        .HorizontalAlignment = xlCenter
    End With

    ' Row 17: Spacer
    ws.Rows(17).RowHeight = 8

    ' Row 18: Back to DASHBOARD hyperlink
    ws.Rows(18).RowHeight = 16
    With ws.Range("B18")
        .Value = "<< Back to DASHBOARD"
        .Font.Bold = True: .Font.Color = RGB(0, 70, 170)
        .Font.Underline = xlUnderlineStyleSingle
    End With
    ws.Hyperlinks.Add Anchor:=ws.Range("B18"), Address:="", _
        SubAddress:="#DASHBOARD!A1", TextToDisplay:="<< Back to DASHBOARD"

    ' Column widths
    ws.Columns("A:A").ColumnWidth = 18
    ws.Columns("B:E").ColumnWidth = 16
    ws.Columns("F:F").ColumnWidth = 14
    ws.Columns("G:H").ColumnWidth = 20

    ws.Tab.Color = RGB(0, 176, 240)
End Sub

'================================================================================
' FILL DATA PUBLIC MACROS
'================================================================================

' Navigate to the DAILY_ENTRY form
Public Sub NavigateToDailyEntry()
    On Error Resume Next
    Dim wsDE As Worksheet
    Set wsDE = ThisWorkbook.Worksheets("DAILY_ENTRY")
    If Not wsDE Is Nothing Then
        wsDE.Activate
        wsDE.Range("B4").Select
    End If
    On Error GoTo 0
End Sub

' Clear the DAILY_ENTRY form fields back to defaults
Public Sub ClearDailyForm()
    Dim wsDE As Worksheet
    On Error Resume Next
    Set wsDE = ThisWorkbook.Worksheets("DAILY_ENTRY")
    On Error GoTo 0
    If wsDE Is Nothing Then Exit Sub

    wsDE.Range("B4").Value  = Date   ' reset date to today
    wsDE.Range("B8").Value  = 0 : wsDE.Range("C8").Value = 0
    wsDE.Range("D8").Value  = 0 : wsDE.Range("E8").Value = 0
    wsDE.Range("B9").Value  = 0 : wsDE.Range("C9").Value = 0
    wsDE.Range("D9").Value  = 0 : wsDE.Range("E9").Value = 0
    wsDE.Range("B12").Value = ""
    wsDE.Range("B13").Value = ""
End Sub

' Submit daily shift data from DAILY_ENTRY to HISTORY, T_DISPATCH_KPI, and
' the supporting input tables (IN_TARGETS_DAILY and IN_AUDIT_LOG).
Public Sub SubmitDailyData()
    Dim wsDE As Worksheet
    On Error Resume Next
    Set wsDE = ThisWorkbook.Worksheets("DAILY_ENTRY")
    On Error GoTo 0
    If wsDE Is Nothing Then
        MsgBox "DAILY_ENTRY sheet not found.", vbCritical
        Exit Sub
    End If

    ' ── Read and validate Business Date ────────────────────────────────────────
    Dim entryDateVal As Variant
    entryDateVal = wsDE.Range("B4").Value
    If IsEmpty(entryDateVal) Or IsError(entryDateVal) Or Not IsDate(entryDateVal) Then
        MsgBox "Please enter a valid Business Date in the Date field (row 4).", _
               vbExclamation, "Missing Date"
        wsDE.Range("B4").Select
        Exit Sub
    End If
    Dim entryDate As Date
    entryDate = CDate(entryDateVal)

    ' ── Read shift data ────────────────────────────────────────────────────────
    Dim dayShipped  As Long:   dayShipped  = _DELong(wsDE.Range("B8").Value)
    Dim dayStaff    As Long:   dayStaff    = _DELong(wsDE.Range("C8").Value)
    Dim dayTarget   As Double: dayTarget   = _DEDbl(wsDE.Range("D8").Value)
    Dim dayAudit    As Long:   dayAudit    = _DELong(wsDE.Range("E8").Value)

    Dim nightShipped As Long:   nightShipped = _DELong(wsDE.Range("B9").Value)
    Dim nightStaff   As Long:   nightStaff   = _DELong(wsDE.Range("C9").Value)
    Dim nightTarget  As Double: nightTarget  = _DEDbl(wsDE.Range("D9").Value)
    Dim nightAudit   As Long:   nightAudit   = _DELong(wsDE.Range("E9").Value)

    Dim dayHasData   As Boolean: dayHasData   = (dayShipped > 0 Or dayStaff > 0)
    Dim nightHasData As Boolean: nightHasData = (nightShipped > 0 Or nightStaff > 0)

    If Not dayHasData And Not nightHasData Then
        MsgBox "Please enter Cartons Shipped and/or Total Staff for at least one shift.", _
               vbExclamation, "No Shift Data"
        Exit Sub
    End If

    ' Validate Target / Person > 0 for every shift that has data
    If dayHasData And dayTarget <= 0 Then
        MsgBox "Please enter a Target / Person greater than 0 for the Day shift.", _
               vbExclamation, "Missing Day Target"
        Exit Sub
    End If

    If nightHasData And nightTarget <= 0 Then
        MsgBox "Please enter a Target / Person greater than 0 for the Night shift.", _
               vbExclamation, "Missing Night Target"
        Exit Sub
    End If

    ' ── Exception count overrides ──────────────────────────────────────────────
    Dim hrpOpen As Long, packedOverdue As Long
    Dim hrpVal As Variant: hrpVal = wsDE.Range("B12").Value
    Dim pckVal As Variant: pckVal = wsDE.Range("B13").Value

    If IsEmpty(hrpVal) Or IsError(hrpVal) Or hrpVal = "" Then
        On Error Resume Next
        hrpOpen = Application.WorksheetFunction.CountIf( _
            ThisWorkbook.Worksheets("IN_HRP").ListObjects("tblHRP").ListColumns("IncludeInHRP").DataBodyRange, True)
        On Error GoTo 0
    Else
        hrpOpen = _DELong(hrpVal)
    End If

    If IsEmpty(pckVal) Or IsError(pckVal) Or pckVal = "" Then
        On Error Resume Next
        packedOverdue = Application.WorksheetFunction.CountIf( _
            ThisWorkbook.Worksheets("IN_PACKED").ListObjects("tblPacked").ListColumns("ActionFlag").DataBodyRange, True)
        On Error GoTo 0
    Else
        packedOverdue = _DELong(pckVal)
    End If

    ' ── Write data ─────────────────────────────────────────────────────────────
    ' Capture prior Application state so the error handler can always restore it
    Dim prevCalc As XlCalculation: prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo SubmitErr

    Dim snapTime As Date: snapTime = Now()
    Dim rowsAdded As Long: rowsAdded = 0

    If dayHasData Then
        Call _UpsertKPIRow(entryDate, "Day", dayShipped, dayStaff)
        If dayTarget > 0 Then Call _UpsertTargets(entryDate, "Day", dayTarget)
        Call _UpsertAuditLog(entryDate, "Day", dayAudit)
        Call _AppendHistory(snapTime, entryDate, "Day", dayShipped, dayStaff, _
                            dayTarget, dayAudit, hrpOpen, packedOverdue)
        rowsAdded = rowsAdded + 1
    End If

    If nightHasData Then
        Call _UpsertKPIRow(entryDate, "Night", nightShipped, nightStaff)
        If nightTarget > 0 Then Call _UpsertTargets(entryDate, "Night", nightTarget)
        Call _UpsertAuditLog(entryDate, "Night", nightAudit)
        Call _AppendHistory(snapTime, entryDate, "Night", nightShipped, nightStaff, _
                            nightTarget, nightAudit, hrpOpen, packedOverdue)
        rowsAdded = rowsAdded + 1
    End If

    Application.Calculation = prevCalc
    Application.Calculate
    Application.ScreenUpdating = True

    MsgBox "Daily data submitted for " & Format(entryDate, "dd/mm/yyyy") & "." & vbCrLf & _
           rowsAdded & " shift record(s) saved to HISTORY and T_DISPATCH_KPI.", _
           vbInformation, "Submit Complete"
    Exit Sub

SubmitErr:
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    MsgBox "An error occurred while saving data:" & vbCrLf & Err.Description, _
           vbCritical, "Submit Error"
End Sub

'================================================================================
' FILL DATA HELPER ROUTINES (private — called only from SubmitDailyData)
'================================================================================

' Upsert a row in T_DISPATCH_KPI for the given date+shift, writing raw values
' for ShippedCartons and TotalStaff (the remaining columns stay as formulas).
Private Sub _UpsertKPIRow(busDate As Date, shiftName As String, _
                           shipped As Long, staff As Long)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("T_DISPATCH_KPI"): On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim tbl As ListObject
    On Error Resume Next: Set tbl = ws.ListObjects("tblDispatchKPI"): On Error GoTo 0
    If tbl Is Nothing Then Exit Sub

    Dim foundRow As Long: foundRow = 0
    Dim r As Long
    Dim dv As Variant, sv As Variant
    For r = 1 To tbl.ListRows.Count
        dv = tbl.ListColumns("BusinessDate").DataBodyRange(r).Value
        sv = tbl.ListColumns("ShiftName").DataBodyRange(r).Value
        If Not IsError(dv) And Not IsError(sv) Then
            If IsDate(dv) And CDate(dv) = busDate And sv = shiftName Then
                foundRow = r: Exit For
            End If
        End If
    Next r

    If foundRow = 0 Then
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add
        foundRow = tbl.ListRows.Count
        tbl.ListColumns("BusinessDate").DataBodyRange(foundRow).Value = busDate
        tbl.ListColumns("ShiftName").DataBodyRange(foundRow).Value    = shiftName
    End If

    ' Write values directly — this overrides any existing formula in that cell only.
    tbl.ListColumns("ShippedCartons").DataBodyRange(foundRow).Value = shipped
    tbl.ListColumns("TotalStaff").DataBodyRange(foundRow).Value     = staff
End Sub

' Upsert a row in IN_TARGETS_DAILY for the given date+shift.
Private Sub _UpsertTargets(busDate As Date, shiftName As String, target As Double)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("IN_TARGETS_DAILY"): On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim tbl As ListObject
    On Error Resume Next: Set tbl = ws.ListObjects("tblTargetsDaily"): On Error GoTo 0
    If tbl Is Nothing Then Exit Sub

    Dim foundRow As Long: foundRow = 0
    Dim r As Long
    Dim dv As Variant, sv As Variant
    For r = 1 To tbl.ListRows.Count
        dv = tbl.ListColumns("BusinessDate").DataBodyRange(r).Value
        sv = tbl.ListColumns("ShiftName").DataBodyRange(r).Value
        If Not IsError(dv) And Not IsError(sv) Then
            If IsDate(dv) And CDate(dv) = busDate And sv = shiftName Then
                foundRow = r: Exit For
            End If
        End If
    Next r

    If foundRow = 0 Then
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add
        foundRow = tbl.ListRows.Count
        tbl.ListColumns("BusinessDate").DataBodyRange(foundRow).Value              = busDate
        tbl.ListColumns("ShiftName").DataBodyRange(foundRow).Value                = shiftName
        tbl.ListColumns("TargetPerPersonPerShift").DataBodyRange(foundRow).Value  = target
    Else
        tbl.ListColumns("TargetPerPersonPerShift").DataBodyRange(foundRow).Value  = target
    End If
End Sub

' Upsert a row in IN_AUDIT_LOG for the given date+shift.
Private Sub _UpsertAuditLog(busDate As Date, shiftName As String, auditCnt As Long)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("IN_AUDIT_LOG"): On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim tbl As ListObject
    On Error Resume Next: Set tbl = ws.ListObjects("tblAuditLog"): On Error GoTo 0
    If tbl Is Nothing Then Exit Sub

    Dim foundRow As Long: foundRow = 0
    Dim r As Long
    Dim dv As Variant, sv As Variant
    For r = 1 To tbl.ListRows.Count
        dv = tbl.ListColumns("BusinessDate").DataBodyRange(r).Value
        sv = tbl.ListColumns("ShiftName").DataBodyRange(r).Value
        If Not IsError(dv) And Not IsError(sv) Then
            If IsDate(dv) And CDate(dv) = busDate And sv = shiftName Then
                foundRow = r: Exit For
            End If
        End If
    Next r

    If foundRow = 0 Then
        Dim newRow As ListRow
        Set newRow = tbl.ListRows.Add
        foundRow = tbl.ListRows.Count
        tbl.ListColumns("BusinessDate").DataBodyRange(foundRow).Value  = busDate
        tbl.ListColumns("ShiftName").DataBodyRange(foundRow).Value    = shiftName
        tbl.ListColumns("Area").DataBodyRange(foundRow).Value         = "Auditing"
        tbl.ListColumns("AuditCount").DataBodyRange(foundRow).Value   = auditCnt
    Else
        tbl.ListColumns("AuditCount").DataBodyRange(foundRow).Value   = auditCnt
    End If
End Sub

' Append a row to HISTORY, computing derived KPIs inline so the snapshot is
' self-contained (does not depend on T_DISPATCH_KPI formula results).
' RAG thresholds are read from tblConfig_Rules (Amber_Threshold / Green_Threshold)
' so HISTORY rows stay consistent when thresholds are changed in CONFIG.
Private Sub _AppendHistory(snapTime As Date, busDate As Date, shiftName As String, _
                            shipped As Long, staff As Long, target As Double, _
                            auditCnt As Long, hrpOpen As Long, packedOverdue As Long)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets("HISTORY"): On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim tbl As ListObject
    On Error Resume Next: Set tbl = ws.ListObjects("tblHistory"): On Error GoTo 0
    If tbl Is Nothing Then Exit Sub

    ' Read RAG thresholds from CONFIG; fall back to standard values if not found
    Dim greenThresh As Double: greenThresh = 1
    Dim amberThresh As Double: amberThresh = 0.9
    Dim wsCfg As Worksheet
    On Error Resume Next
    Set wsCfg = ThisWorkbook.Worksheets("CONFIG")
    If Not wsCfg Is Nothing Then
        Dim ruleTbl As ListObject
        Set ruleTbl = wsCfg.ListObjects("tblConfig_Rules")
        If Not ruleTbl Is Nothing Then
            Dim rv As Variant
            rv = Application.WorksheetFunction.IfError( _
                    Application.WorksheetFunction.Index( _
                        ruleTbl.ListColumns("Value").DataBodyRange, _
                        Application.WorksheetFunction.Match("Green_Threshold", _
                            ruleTbl.ListColumns("RuleName").DataBodyRange, 0)), _
                    1)
            If IsNumeric(rv) Then greenThresh = CDbl(rv)

            rv = Application.WorksheetFunction.IfError( _
                    Application.WorksheetFunction.Index( _
                        ruleTbl.ListColumns("Value").DataBodyRange, _
                        Application.WorksheetFunction.Match("Amber_Threshold", _
                            ruleTbl.ListColumns("RuleName").DataBodyRange, 0)), _
                    0.9)
            If IsNumeric(rv) Then amberThresh = CDbl(rv)
        End If
    End If
    On Error GoTo 0

    Dim expected As Double: expected = staff * target
    Dim perfPct  As Double: perfPct  = IIf(expected > 0, shipped / expected, 0)
    Dim rag      As String
    If perfPct >= greenThresh Then
        rag = "Green"
    ElseIf perfPct >= amberThresh Then
        rag = "Amber"
    Else
        rag = "Red"
    End If

    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    newRow.Range(1).Value  = snapTime
    newRow.Range(2).Value  = busDate
    newRow.Range(3).Value  = shiftName
    newRow.Range(4).Value  = shipped
    newRow.Range(5).Value  = staff
    newRow.Range(6).Value  = expected
    newRow.Range(7).Value  = perfPct
    newRow.Range(8).Value  = rag
    newRow.Range(9).Value  = hrpOpen
    newRow.Range(10).Value = packedOverdue
    newRow.Range(11).Value = auditCnt
End Sub

' Safe Variant-to-Long conversion for form input cells.
Private Function _DELong(v As Variant) As Long
    If IsEmpty(v) Or IsError(v) Or v = "" Then
        _DELong = 0
    ElseIf IsNumeric(v) Then
        _DELong = CLng(v)
    Else
        _DELong = 0
    End If
End Function

' Safe Variant-to-Double conversion for form input cells.
Private Function _DEDbl(v As Variant) As Double
    If IsEmpty(v) Or IsError(v) Or v = "" Then
        _DEDbl = 0
    ElseIf IsNumeric(v) Then
        _DEDbl = CDbl(v)
    Else
        _DEDbl = 0
    End If
End Function

Attribute VB_Name = "ModWorkbookStructure"
Option Explicit

' =============================================================================
' 01_ExtractWorkbookStructure.bas
'
' Purpose:
'   Builds a high-level structural inventory of the ACTIVE Excel workbook.
'
' Safety:
'   - Reads the source workbook only.
'   - Does NOT write to the source workbook.
'   - Creates a brand-new unsaved results workbook.
'   - Does NOT extract source cell values.
'
' Important:
'   Before running this macro, click the client/source workbook so that it is
'   the ActiveWorkbook.
' =============================================================================

Public Sub ExtractWorkbookStructure_Safe()

    Dim sourceWb As Workbook
    Dim resultWb As Workbook
    Dim resultWs As Worksheet
    Dim ws As Worksheet
    Dim rowNum As Long

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation

    On Error GoTo CleanFail

    Set sourceWb = ActiveWorkbook

    If sourceWb Is Nothing Then
        MsgBox "No active workbook was found. Open and activate the source workbook first.", vbExclamation
        Exit Sub
    End If

    If sourceWb.Worksheets.Count = 0 Then
        MsgBox "The active workbook does not contain any worksheets.", vbExclamation
        Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Inspecting workbook structure..."

    Set resultWb = Workbooks.Add(xlWBATWorksheet)
    Set resultWs = resultWb.Worksheets(1)
    resultWs.Name = "WORKBOOK_STRUCTURE"

    WriteStructureHeaders resultWs

    rowNum = 2

    For Each ws In sourceWb.Worksheets

        Application.StatusBar = "Inspecting sheet: " & ws.Name

        resultWs.Cells(rowNum, 1).Value = sourceWb.Name
        resultWs.Cells(rowNum, 2).Value = ws.Name
        resultWs.Cells(rowNum, 3).Value = SheetVisibilityName(ws.Visible)
        resultWs.Cells(rowNum, 4).Value = ws.UsedRange.Address(False, False)
        resultWs.Cells(rowNum, 5).Value = ws.UsedRange.Rows.Count
        resultWs.Cells(rowNum, 6).Value = ws.UsedRange.Columns.Count
        resultWs.Cells(rowNum, 7).Value = SafeSpecialCellCount(ws, xlCellTypeFormulas)
        resultWs.Cells(rowNum, 8).Value = SafeSpecialCellCount(ws, xlCellTypeConstants)
        resultWs.Cells(rowNum, 9).Value = ws.ListObjects.Count
        resultWs.Cells(rowNum, 10).Value = ws.PivotTables.Count
        resultWs.Cells(rowNum, 11).Value = ws.ChartObjects.Count
        resultWs.Cells(rowNum, 12).Value = SafeCommentCount(ws)
        resultWs.Cells(rowNum, 13).Value = SafeTableNames(ws)

        rowNum = rowNum + 1
    Next ws

    FormatStructureSheet resultWs, rowNum - 1

    Application.StatusBar = False

    RestoreExcelState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation

    resultWb.Activate
    resultWs.Activate

    MsgBox _
        "Workbook structure extraction completed." & vbCrLf & vbCrLf & _
        "Source workbook: " & sourceWb.Name & vbCrLf & _
        "Sheets inspected: " & sourceWb.Worksheets.Count & vbCrLf & vbCrLf & _
        "A NEW unsaved workbook contains the results." & vbCrLf & _
        "The source workbook was not modified.", _
        vbInformation, _
        "Workbook Structure Complete"

    Exit Sub

CleanFail:
    Application.StatusBar = False
    RestoreExcelState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation

    MsgBox _
        "Workbook structure extraction stopped because of an error." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "Extraction Error"

End Sub

Private Sub WriteStructureHeaders(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "Workbook", _
        "Sheet", _
        "Visibility", _
        "Used Range", _
        "Used Rows", _
        "Used Columns", _
        "Formula Cells", _
        "Constant Cells", _
        "Excel Tables", _
        "Pivot Tables", _
        "Charts", _
        "Legacy Notes/Comments", _
        "Table Names" _
    )

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

End Sub

Private Function SheetVisibilityName(ByVal visibility As XlSheetVisibility) As String

    Select Case visibility
        Case xlSheetVisible
            SheetVisibilityName = "Visible"
        Case xlSheetHidden
            SheetVisibilityName = "Hidden"
        Case xlSheetVeryHidden
            SheetVisibilityName = "Very Hidden"
        Case Else
            SheetVisibilityName = "Unknown"
    End Select

End Function

Private Function SafeSpecialCellCount(ByVal ws As Worksheet, ByVal cellType As XlCellType) As Double

    Dim rng As Range

    On Error Resume Next
    Set rng = ws.UsedRange.SpecialCells(cellType)
    On Error GoTo 0

    If rng Is Nothing Then
        SafeSpecialCellCount = 0
    Else
        SafeSpecialCellCount = rng.CountLarge
    End If

End Function

Private Function SafeCommentCount(ByVal ws As Worksheet) As Long

    Dim cell As Range
    Dim rng As Range
    Dim countComments As Long

    countComments = 0

    On Error Resume Next
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeComments)
    On Error GoTo 0

    If Not rng Is Nothing Then
        For Each cell In rng.Cells
            countComments = countComments + 1
        Next cell
    End If

    SafeCommentCount = countComments

End Function

Private Function SafeTableNames(ByVal ws As Worksheet) As String

    Dim lo As ListObject
    Dim result As String

    result = ""

    On Error Resume Next

    For Each lo In ws.ListObjects
        If Len(result) > 0 Then result = result & "; "
        result = result & lo.Name
    Next lo

    On Error GoTo 0

    SafeTableNames = result

End Function

Private Sub FormatStructureSheet(ByVal ws As Worksheet, ByVal lastRow As Long)

    With ws.Rows(1)
        .Font.Bold = True
        .AutoFilter
    End With

    ws.Columns("A:M").EntireColumn.AutoFit
    ws.Columns("M:M").ColumnWidth = 35
    ws.Range("A1:M" & lastRow).VerticalAlignment = xlTop
    ws.Range("A1:M" & lastRow).WrapText = False
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

End Sub

Private Sub RestoreExcelState( _
    ByVal screenUpdatingValue As Boolean, _
    ByVal enableEventsValue As Boolean, _
    ByVal displayAlertsValue As Boolean, _
    ByVal calculationValue As XlCalculation)

    On Error Resume Next
    Application.ScreenUpdating = screenUpdatingValue
    Application.EnableEvents = enableEventsValue
    Application.DisplayAlerts = displayAlertsValue
    Application.Calculation = calculationValue
    Application.StatusBar = False
    On Error GoTo 0

End Sub

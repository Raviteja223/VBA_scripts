Attribute VB_Name = "ModFormulaPatterns"
Option Explicit

' =============================================================================
' 02_ExtractUniqueFormulaPatterns.bas
'
' Purpose:
'   Finds formula cells in the ACTIVE workbook and collapses repeated copied-
'   down formulas into unique R1C1 formula patterns.
'
' Why R1C1?
'   Example:
'       M2 = H2-G2
'       M3 = H3-G3
'       M4 = H4-G4
'   All three normalize to the same relative pattern, allowing a workbook with
'   hundreds of thousands of formula cells to be reduced to a much smaller
'   catalog of calculation patterns.
'
' Safety:
'   - Reads formulas/metadata only.
'   - Does NOT copy source cell values.
'   - Does NOT write to the source workbook.
'   - Creates a NEW unsaved results workbook.
'
' Important:
'   Before running this macro, click the client/source workbook so that it is
'   the ActiveWorkbook.
' =============================================================================

Public Sub ExtractUniqueFormulaPatterns_Safe()

    Dim sourceWb As Workbook
    Dim resultWb As Workbook
    Dim resultWs As Worksheet
    Dim ws As Worksheet
    Dim formulaCells As Range
    Dim cell As Range

    Dim countDict As Object
    Dim exampleCellDict As Object
    Dim exampleFormulaDict As Object
    Dim sheetDict As Object
    Dim visibilityDict As Object
    Dim columnDict As Object
    Dim categoryDict As Object
    Dim crossSheetDict As Object
    Dim externalRefDict As Object

    Dim key As Variant
    Dim delimiter As String
    Dim pattern As String
    Dim formulaA1 As String
    Dim rowNum As Long
    Dim formulaCellCount As Double

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

    delimiter = Chr(30)

    Set countDict = CreateObject("Scripting.Dictionary")
    Set exampleCellDict = CreateObject("Scripting.Dictionary")
    Set exampleFormulaDict = CreateObject("Scripting.Dictionary")
    Set sheetDict = CreateObject("Scripting.Dictionary")
    Set visibilityDict = CreateObject("Scripting.Dictionary")
    Set columnDict = CreateObject("Scripting.Dictionary")
    Set categoryDict = CreateObject("Scripting.Dictionary")
    Set crossSheetDict = CreateObject("Scripting.Dictionary")
    Set externalRefDict = CreateObject("Scripting.Dictionary")

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    formulaCellCount = 0

    For Each ws In sourceWb.Worksheets

        Application.StatusBar = "Scanning formulas in sheet: " & ws.Name

        Set formulaCells = Nothing

        On Error Resume Next
        Set formulaCells = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        On Error GoTo CleanFail

        If Not formulaCells Is Nothing Then

            For Each cell In formulaCells.Cells

                formulaCellCount = formulaCellCount + 1

                pattern = SafeFormulaR1C1(cell)
                formulaA1 = SafeFormulaA1(cell)

                key = ws.Name & delimiter & CStr(cell.Column) & delimiter & pattern

                If countDict.Exists(key) Then
                    countDict(key) = CDbl(countDict(key)) + 1
                Else
                    countDict.Add key, 1
                    exampleCellDict.Add key, cell.Address(False, False)
                    exampleFormulaDict.Add key, formulaA1
                    sheetDict.Add key, ws.Name
                    visibilityDict.Add key, SheetVisibilityName(ws.Visible)
                    columnDict.Add key, ColumnLetter(cell.Column)
                    categoryDict.Add key, DetectFormulaType(formulaA1)
                    crossSheetDict.Add key, YesNo(InStr(1, formulaA1, "!", vbTextCompare) > 0)
                    externalRefDict.Add key, YesNo(InStr(1, formulaA1, "[", vbTextCompare) > 0 And InStr(1, formulaA1, "]", vbTextCompare) > 0)
                End If

            Next cell

        End If

    Next ws

    Set resultWb = Workbooks.Add(xlWBATWorksheet)
    Set resultWs = resultWb.Worksheets(1)
    resultWs.Name = "FORMULA_PATTERNS"

    WritePatternHeaders resultWs

    rowNum = 2

    For Each key In countDict.Keys

        resultWs.Cells(rowNum, 1).Value = sourceWb.Name
        resultWs.Cells(rowNum, 2).Value = sheetDict(key)
        resultWs.Cells(rowNum, 3).Value = visibilityDict(key)
        resultWs.Cells(rowNum, 4).Value = columnDict(key)
        resultWs.Cells(rowNum, 5).Value = exampleCellDict(key)
        resultWs.Cells(rowNum, 6).Value = "'" & exampleFormulaDict(key)
        resultWs.Cells(rowNum, 7).Value = "'" & ExtractPatternFromKey(CStr(key), delimiter)
        resultWs.Cells(rowNum, 8).Value = countDict(key)
        resultWs.Cells(rowNum, 9).Value = categoryDict(key)
        resultWs.Cells(rowNum, 10).Value = crossSheetDict(key)
        resultWs.Cells(rowNum, 11).Value = externalRefDict(key)

        rowNum = rowNum + 1

    Next key

    FormatPatternSheet resultWs, rowNum - 1

    Application.StatusBar = False
    RestoreExcelState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation

    resultWb.Activate
    resultWs.Activate

    MsgBox _
        "Formula-pattern extraction completed." & vbCrLf & vbCrLf & _
        "Source workbook: " & sourceWb.Name & vbCrLf & _
        "Formula cells scanned: " & Format(formulaCellCount, "#,##0") & vbCrLf & _
        "Unique formula patterns: " & Format(countDict.Count, "#,##0") & vbCrLf & vbCrLf & _
        "A NEW unsaved workbook contains the results." & vbCrLf & _
        "The source workbook was not modified.", _
        vbInformation, _
        "Formula Pattern Extraction Complete"

    Exit Sub

CleanFail:
    Application.StatusBar = False
    RestoreExcelState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts, oldCalculation

    MsgBox _
        "Formula-pattern extraction stopped because of an error." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "Extraction Error"

End Sub

Private Sub WritePatternHeaders(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "Workbook", _
        "Sheet", _
        "Visibility", _
        "Output Column", _
        "Example Cell", _
        "Example Formula (A1)", _
        "Normalized Formula Pattern (R1C1)", _
        "Occurrences", _
        "Formula Category", _
        "Cross-Sheet Reference", _
        "Possible External Workbook Reference" _
    )

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

End Sub

Private Function SafeFormulaR1C1(ByVal cell As Range) As String

    On Error Resume Next
    SafeFormulaR1C1 = CStr(cell.FormulaR1C1)

    If Err.Number <> 0 Then
        Err.Clear
        SafeFormulaR1C1 = CStr(cell.Formula)
    End If

    On Error GoTo 0

End Function

Private Function SafeFormulaA1(ByVal cell As Range) As String

    On Error Resume Next
    SafeFormulaA1 = CStr(cell.Formula)

    If Err.Number <> 0 Then
        Err.Clear
        SafeFormulaA1 = "<Formula unavailable>"
    End If

    On Error GoTo 0

End Function

Private Function ExtractPatternFromKey(ByVal keyText As String, ByVal delimiter As String) As String

    Dim parts As Variant
    Dim i As Long
    Dim result As String

    parts = Split(keyText, delimiter)

    If UBound(parts) < 2 Then
        ExtractPatternFromKey = keyText
        Exit Function
    End If

    result = parts(2)

    If UBound(parts) > 2 Then
        For i = 3 To UBound(parts)
            result = result & delimiter & parts(i)
        Next i
    End If

    ExtractPatternFromKey = result

End Function

Private Function DetectFormulaType(ByVal formulaText As String) As String

    Dim f As String

    f = UCase$(formulaText)

    If InStr(f, "XLOOKUP(") > 0 Then
        DetectFormulaType = "XLOOKUP"
    ElseIf InStr(f, "VLOOKUP(") > 0 Then
        DetectFormulaType = "VLOOKUP"
    ElseIf InStr(f, "HLOOKUP(") > 0 Then
        DetectFormulaType = "HLOOKUP"
    ElseIf InStr(f, "INDEX(") > 0 And InStr(f, "MATCH(") > 0 Then
        DetectFormulaType = "INDEX/MATCH"
    ElseIf InStr(f, "SUMIFS(") > 0 Then
        DetectFormulaType = "SUMIFS"
    ElseIf InStr(f, "SUMIF(") > 0 Then
        DetectFormulaType = "SUMIF"
    ElseIf InStr(f, "COUNTIFS(") > 0 Then
        DetectFormulaType = "COUNTIFS"
    ElseIf InStr(f, "COUNTIF(") > 0 Then
        DetectFormulaType = "COUNTIF"
    ElseIf InStr(f, "AVERAGEIFS(") > 0 Then
        DetectFormulaType = "AVERAGEIFS"
    ElseIf InStr(f, "AVERAGEIF(") > 0 Then
        DetectFormulaType = "AVERAGEIF"
    ElseIf InStr(f, "GETPIVOTDATA(") > 0 Then
        DetectFormulaType = "GETPIVOTDATA"
    ElseIf InStr(f, "SUBTOTAL(") > 0 Then
        DetectFormulaType = "SUBTOTAL"
    ElseIf InStr(f, "IFERROR(") > 0 Then
        DetectFormulaType = "IFERROR"
    ElseIf InStr(f, "IFS(") > 0 Then
        DetectFormulaType = "IFS"
    ElseIf InStr(f, "IF(") > 0 Then
        DetectFormulaType = "IF"
    ElseIf InStr(f, "SUMPRODUCT(") > 0 Then
        DetectFormulaType = "SUMPRODUCT"
    ElseIf InStr(f, "SUM(") > 0 Then
        DetectFormulaType = "SUM"
    ElseIf InStr(f, "COUNT(") > 0 Or InStr(f, "COUNTA(") > 0 Then
        DetectFormulaType = "COUNT/COUNTA"
    ElseIf InStr(f, "DATE(") > 0 Or InStr(f, "TODAY(") > 0 Or InStr(f, "EOMONTH(") > 0 Then
        DetectFormulaType = "Date Logic"
    ElseIf InStr(f, "LEFT(") > 0 Or InStr(f, "RIGHT(") > 0 Or InStr(f, "MID(") > 0 Or InStr(f, "TEXT(") > 0 Then
        DetectFormulaType = "Text Logic"
    Else
        DetectFormulaType = "Other"
    End If

End Function

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

Private Function ColumnLetter(ByVal columnNumber As Long) As String

    ColumnLetter = Split(Cells(1, columnNumber).Address(True, False), "$")(0)

End Function

Private Function YesNo(ByVal conditionValue As Boolean) As String

    If conditionValue Then
        YesNo = "Yes"
    Else
        YesNo = "No"
    End If

End Function

Private Sub FormatPatternSheet(ByVal ws As Worksheet, ByVal lastRow As Long)

    With ws.Rows(1)
        .Font.Bold = True
        .AutoFilter
    End With

    ws.Columns("A:E").EntireColumn.AutoFit
    ws.Columns("F:G").ColumnWidth = 70
    ws.Columns("H:K").EntireColumn.AutoFit

    ws.Range("F2:G" & lastRow).WrapText = True
    ws.Range("A1:K" & lastRow).VerticalAlignment = xlTop

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

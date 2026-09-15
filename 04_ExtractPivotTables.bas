Attribute VB_Name = "ModPivotTables"
Option Explicit

' =============================================================================
' 04_ExtractPivotTables.bas
'
' Purpose:
'   Catalogs PivotTables in the ACTIVE workbook, including:
'       - Pivot summary metadata
'       - Row/column/filter/value fields
'       - Data-field aggregation functions
'       - Calculated Pivot fields where accessible
'
' Safety:
'   - Reads PivotTable configuration/metadata only.
'   - Does NOT export the underlying row-level source data.
'   - Does NOT write to the source workbook.
'   - Creates a NEW unsaved results workbook.
'
' Important:
'   Before running this macro, click the client/source workbook so that it is
'   the ActiveWorkbook.
' =============================================================================

Public Sub ExtractPivotTables_Safe()

    Dim sourceWb As Workbook
    Dim resultWb As Workbook
    Dim summaryWs As Worksheet
    Dim fieldsWs As Worksheet
    Dim calculatedWs As Worksheet

    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim calcField As PivotField

    Dim summaryRow As Long
    Dim fieldsRow As Long
    Dim calculatedRow As Long

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean

    On Error GoTo CleanFail

    Set sourceWb = ActiveWorkbook

    If sourceWb Is Nothing Then
        MsgBox "No active workbook was found. Open and activate the source workbook first.", vbExclamation
        Exit Sub
    End If

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set resultWb = Workbooks.Add(xlWBATWorksheet)

    Set summaryWs = resultWb.Worksheets(1)
    summaryWs.Name = "PIVOT_TABLES"

    Set fieldsWs = resultWb.Worksheets.Add(After:=summaryWs)
    fieldsWs.Name = "PIVOT_FIELDS"

    Set calculatedWs = resultWb.Worksheets.Add(After:=fieldsWs)
    calculatedWs.Name = "PIVOT_CALCULATED_FIELDS"

    WritePivotSummaryHeaders summaryWs
    WritePivotFieldHeaders fieldsWs
    WriteCalculatedFieldHeaders calculatedWs

    summaryRow = 2
    fieldsRow = 2
    calculatedRow = 2

    For Each ws In sourceWb.Worksheets

        For Each pt In ws.PivotTables

            Application.StatusBar = "Inspecting PivotTable: " & ws.Name & " -> " & pt.Name

            summaryWs.Cells(summaryRow, 1).Value = sourceWb.Name
            summaryWs.Cells(summaryRow, 2).Value = ws.Name
            summaryWs.Cells(summaryRow, 3).Value = pt.Name
            summaryWs.Cells(summaryRow, 4).Value = SafePivotRange(pt)
            summaryWs.Cells(summaryRow, 5).Value = "'" & SafePivotSource(pt)
            summaryWs.Cells(summaryRow, 6).Value = SafePivotCacheIndex(pt)
            summaryWs.Cells(summaryRow, 7).Value = SafePivotRefreshOnOpen(pt)
            summaryWs.Cells(summaryRow, 8).Value = SafeCollectionCount(pt.RowFields)
            summaryWs.Cells(summaryRow, 9).Value = SafeCollectionCount(pt.ColumnFields)
            summaryWs.Cells(summaryRow, 10).Value = SafeCollectionCount(pt.PageFields)
            summaryWs.Cells(summaryRow, 11).Value = SafeCollectionCount(pt.DataFields)
            summaryWs.Cells(summaryRow, 12).Value = YesNo(IsPossibleExternalReference(SafePivotSource(pt)))

            summaryRow = summaryRow + 1

            On Error Resume Next

            For Each pf In pt.PivotFields

                If pf.Orientation <> xlHidden Then

                    fieldsWs.Cells(fieldsRow, 1).Value = sourceWb.Name
                    fieldsWs.Cells(fieldsRow, 2).Value = ws.Name
                    fieldsWs.Cells(fieldsRow, 3).Value = pt.Name
                    fieldsWs.Cells(fieldsRow, 4).Value = pf.Name
                    fieldsWs.Cells(fieldsRow, 5).Value = PivotOrientationName(pf.Orientation)
                    fieldsWs.Cells(fieldsRow, 6).Value = SafePivotFieldPosition(pf)
                    fieldsWs.Cells(fieldsRow, 7).Value = SafeDataFieldFunction(pf)
                    fieldsWs.Cells(fieldsRow, 8).Value = SafeNumberFormat(pf)
                    fieldsWs.Cells(fieldsRow, 9).Value = SafeSourceName(pf)

                    fieldsRow = fieldsRow + 1

                End If

            Next pf

            Err.Clear

            For Each calcField In pt.CalculatedFields

                calculatedWs.Cells(calculatedRow, 1).Value = sourceWb.Name
                calculatedWs.Cells(calculatedRow, 2).Value = ws.Name
                calculatedWs.Cells(calculatedRow, 3).Value = pt.Name
                calculatedWs.Cells(calculatedRow, 4).Value = calcField.Name
                calculatedWs.Cells(calculatedRow, 5).Value = "'" & SafePivotFieldFormula(calcField)

                calculatedRow = calculatedRow + 1

            Next calcField

            On Error GoTo CleanFail

        Next pt

    Next ws

    FormatPivotSummarySheet summaryWs, summaryRow - 1
    FormatPivotFieldSheet fieldsWs, fieldsRow - 1
    FormatCalculatedFieldSheet calculatedWs, calculatedRow - 1

    Application.StatusBar = False
    RestoreExcelState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts

    resultWb.Activate
    summaryWs.Activate

    MsgBox _
        "PivotTable extraction completed." & vbCrLf & vbCrLf & _
        "Source workbook: " & sourceWb.Name & vbCrLf & _
        "PivotTables found: " & Format(summaryRow - 2, "#,##0") & vbCrLf & vbCrLf & _
        "A NEW unsaved workbook contains three result sheets:" & vbCrLf & _
        "  - PIVOT_TABLES" & vbCrLf & _
        "  - PIVOT_FIELDS" & vbCrLf & _
        "  - PIVOT_CALCULATED_FIELDS" & vbCrLf & vbCrLf & _
        "The source workbook was not modified.", _
        vbInformation, _
        "Pivot Extraction Complete"

    Exit Sub

CleanFail:
    Application.StatusBar = False
    RestoreExcelState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts

    MsgBox _
        "PivotTable extraction stopped because of an error." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "Extraction Error"

End Sub

Private Sub WritePivotSummaryHeaders(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "Workbook", _
        "Sheet", _
        "Pivot Name", _
        "Pivot Output Range", _
        "Source Data / Connection", _
        "Pivot Cache Index", _
        "Refresh On File Open", _
        "Row Fields", _
        "Column Fields", _
        "Filter Fields", _
        "Value Fields", _
        "Possible External Reference" _
    )

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

End Sub

Private Sub WritePivotFieldHeaders(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "Workbook", _
        "Sheet", _
        "Pivot Name", _
        "Field", _
        "Orientation", _
        "Position", _
        "Aggregation (for Value fields)", _
        "Number Format", _
        "Source Name" _
    )

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

End Sub

Private Sub WriteCalculatedFieldHeaders(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "Workbook", _
        "Sheet", _
        "Pivot Name", _
        "Calculated Field", _
        "Formula" _
    )

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

End Sub

Private Function SafePivotRange(ByVal pt As PivotTable) As String

    On Error Resume Next
    SafePivotRange = pt.TableRange2.Address(False, False)

    If Err.Number <> 0 Then
        Err.Clear
        SafePivotRange = "<Unavailable>"
    End If

    On Error GoTo 0

End Function

Private Function SafePivotSource(ByVal pt As PivotTable) As String

    Dim src As Variant

    On Error Resume Next
    src = pt.PivotCache.SourceData

    If Err.Number <> 0 Then
        Err.Clear
        SafePivotSource = "<Connection/OLAP source or unavailable>"
    ElseIf IsArray(src) Then
        SafePivotSource = "<Array-based source>"
    Else
        SafePivotSource = CStr(src)
    End If

    On Error GoTo 0

End Function

Private Function SafePivotCacheIndex(ByVal pt As PivotTable) As Variant

    On Error Resume Next
    SafePivotCacheIndex = pt.CacheIndex

    If Err.Number <> 0 Then
        Err.Clear
        SafePivotCacheIndex = ""
    End If

    On Error GoTo 0

End Function

Private Function SafePivotRefreshOnOpen(ByVal pt As PivotTable) As String

    On Error Resume Next
    SafePivotRefreshOnOpen = YesNo(pt.PivotCache.RefreshOnFileOpen)

    If Err.Number <> 0 Then
        Err.Clear
        SafePivotRefreshOnOpen = "Unknown"
    End If

    On Error GoTo 0

End Function

Private Function SafeCollectionCount(ByVal collectionObject As Object) As Variant

    On Error Resume Next
    SafeCollectionCount = collectionObject.Count

    If Err.Number <> 0 Then
        Err.Clear
        SafeCollectionCount = ""
    End If

    On Error GoTo 0

End Function

Private Function SafePivotFieldPosition(ByVal pf As PivotField) As Variant

    On Error Resume Next
    SafePivotFieldPosition = pf.Position

    If Err.Number <> 0 Then
        Err.Clear
        SafePivotFieldPosition = ""
    End If

    On Error GoTo 0

End Function

Private Function SafeDataFieldFunction(ByVal pf As PivotField) As String

    If pf.Orientation <> xlDataField Then
        SafeDataFieldFunction = ""
        Exit Function
    End If

    On Error Resume Next
    SafeDataFieldFunction = PivotFunctionName(pf.Function)

    If Err.Number <> 0 Then
        Err.Clear
        SafeDataFieldFunction = "<Unavailable>"
    End If

    On Error GoTo 0

End Function

Private Function SafeNumberFormat(ByVal pf As PivotField) As String

    On Error Resume Next
    SafeNumberFormat = pf.NumberFormat

    If Err.Number <> 0 Then
        Err.Clear
        SafeNumberFormat = ""
    End If

    On Error GoTo 0

End Function

Private Function SafeSourceName(ByVal pf As PivotField) As String

    On Error Resume Next
    SafeSourceName = pf.SourceName

    If Err.Number <> 0 Then
        Err.Clear
        SafeSourceName = ""
    End If

    On Error GoTo 0

End Function

Private Function SafePivotFieldFormula(ByVal pf As PivotField) As String

    On Error Resume Next
    SafePivotFieldFormula = pf.Formula

    If Err.Number <> 0 Then
        Err.Clear
        SafePivotFieldFormula = "<Unavailable>"
    End If

    On Error GoTo 0

End Function

Private Function PivotOrientationName(ByVal orientationValue As XlPivotFieldOrientation) As String

    Select Case orientationValue
        Case xlRowField
            PivotOrientationName = "Row"
        Case xlColumnField
            PivotOrientationName = "Column"
        Case xlPageField
            PivotOrientationName = "Filter"
        Case xlDataField
            PivotOrientationName = "Value"
        Case xlHidden
            PivotOrientationName = "Hidden"
        Case Else
            PivotOrientationName = "Other"
    End Select

End Function

Private Function PivotFunctionName(ByVal functionValue As XlConsolidationFunction) As String

    Select Case functionValue
        Case xlSum
            PivotFunctionName = "SUM"
        Case xlCount
            PivotFunctionName = "COUNT"
        Case xlAverage
            PivotFunctionName = "AVERAGE"
        Case xlMax
            PivotFunctionName = "MAX"
        Case xlMin
            PivotFunctionName = "MIN"
        Case xlProduct
            PivotFunctionName = "PRODUCT"
        Case xlCountNums
            PivotFunctionName = "COUNT NUMBERS"
        Case xlStdDev
            PivotFunctionName = "STDDEV"
        Case xlStdDevP
            PivotFunctionName = "STDDEV.P"
        Case xlVar
            PivotFunctionName = "VAR"
        Case xlVarP
            PivotFunctionName = "VAR.P"
        Case Else
            PivotFunctionName = "Other/Custom"
    End Select

End Function

Private Function IsPossibleExternalReference(ByVal textValue As String) As Boolean

    IsPossibleExternalReference = _
        (InStr(1, textValue, "[", vbTextCompare) > 0 And _
         InStr(1, textValue, "]", vbTextCompare) > 0)

End Function

Private Function YesNo(ByVal conditionValue As Boolean) As String

    If conditionValue Then
        YesNo = "Yes"
    Else
        YesNo = "No"
    End If

End Function

Private Sub FormatPivotSummarySheet(ByVal ws As Worksheet, ByVal lastRow As Long)

    With ws.Rows(1)
        .Font.Bold = True
        .AutoFilter
    End With

    ws.Columns("A:D").EntireColumn.AutoFit
    ws.Columns("E:E").ColumnWidth = 75
    ws.Columns("F:L").EntireColumn.AutoFit
    ws.Range("E2:E" & lastRow).WrapText = True

End Sub

Private Sub FormatPivotFieldSheet(ByVal ws As Worksheet, ByVal lastRow As Long)

    With ws.Rows(1)
        .Font.Bold = True
        .AutoFilter
    End With

    ws.Columns("A:I").EntireColumn.AutoFit

End Sub

Private Sub FormatCalculatedFieldSheet(ByVal ws As Worksheet, ByVal lastRow As Long)

    With ws.Rows(1)
        .Font.Bold = True
        .AutoFilter
    End With

    ws.Columns("A:D").EntireColumn.AutoFit
    ws.Columns("E:E").ColumnWidth = 80
    ws.Range("E2:E" & lastRow).WrapText = True

End Sub

Private Sub RestoreExcelState( _
    ByVal screenUpdatingValue As Boolean, _
    ByVal enableEventsValue As Boolean, _
    ByVal displayAlertsValue As Boolean)

    On Error Resume Next
    Application.ScreenUpdating = screenUpdatingValue
    Application.EnableEvents = enableEventsValue
    Application.DisplayAlerts = displayAlertsValue
    Application.StatusBar = False
    On Error GoTo 0

End Sub

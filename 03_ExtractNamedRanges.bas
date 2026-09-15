Attribute VB_Name = "ModNamedRanges"
Option Explicit

' =============================================================================
' 03_ExtractNamedRanges.bas
'
' Purpose:
'   Catalogs workbook-level and worksheet-level defined names from the ACTIVE
'   workbook. Named ranges often hide important business logic, lookup ranges,
'   report dates, thresholds, and cross-sheet dependencies.
'
' Safety:
'   - Reads name definitions only.
'   - Does NOT extract the values inside named ranges.
'   - Does NOT write to the source workbook.
'   - Creates a NEW unsaved results workbook.
'
' Important:
'   Before running this macro, click the client/source workbook so that it is
'   the ActiveWorkbook.
' =============================================================================

Public Sub ExtractNamedRanges_Safe()

    Dim sourceWb As Workbook
    Dim resultWb As Workbook
    Dim resultWs As Worksheet
    Dim nm As Name
    Dim rowNum As Long

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
    Application.StatusBar = "Extracting defined names..."

    Set resultWb = Workbooks.Add(xlWBATWorksheet)
    Set resultWs = resultWb.Worksheets(1)
    resultWs.Name = "NAMED_RANGES"

    WriteNameHeaders resultWs

    rowNum = 2

    For Each nm In sourceWb.Names

        Application.StatusBar = "Reading defined name: " & nm.Name

        resultWs.Cells(rowNum, 1).Value = sourceWb.Name
        resultWs.Cells(rowNum, 2).Value = SafeNameScope(nm)
        resultWs.Cells(rowNum, 3).Value = nm.Name
        resultWs.Cells(rowNum, 4).Value = YesNo(SafeNameVisible(nm))
        resultWs.Cells(rowNum, 5).Value = "'" & SafeRefersTo(nm)
        resultWs.Cells(rowNum, 6).Value = SafeTargetSheet(nm)
        resultWs.Cells(rowNum, 7).Value = YesNo(IsPossibleExternalReference(SafeRefersTo(nm)))
        resultWs.Cells(rowNum, 8).Value = SafeNameComment(nm)

        rowNum = rowNum + 1

    Next nm

    FormatNameSheet resultWs, rowNum - 1

    Application.StatusBar = False
    RestoreExcelState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts

    resultWb.Activate
    resultWs.Activate

    MsgBox _
        "Named-range extraction completed." & vbCrLf & vbCrLf & _
        "Source workbook: " & sourceWb.Name & vbCrLf & _
        "Defined names found: " & sourceWb.Names.Count & vbCrLf & vbCrLf & _
        "A NEW unsaved workbook contains the results." & vbCrLf & _
        "The source workbook was not modified.", _
        vbInformation, _
        "Named Range Extraction Complete"

    Exit Sub

CleanFail:
    Application.StatusBar = False
    RestoreExcelState oldScreenUpdating, oldEnableEvents, oldDisplayAlerts

    MsgBox _
        "Named-range extraction stopped because of an error." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical, _
        "Extraction Error"

End Sub

Private Sub WriteNameHeaders(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long

    headers = Array( _
        "Workbook", _
        "Scope", _
        "Defined Name", _
        "Visible", _
        "Refers To", _
        "Target Sheet (if resolvable)", _
        "Possible External Workbook Reference", _
        "Comment" _
    )

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i

End Sub

Private Function SafeNameScope(ByVal nm As Name) As String

    Dim p As Object

    On Error Resume Next
    Set p = nm.Parent
    On Error GoTo 0

    If p Is Nothing Then
        SafeNameScope = "Unknown"
    ElseIf TypeName(p) = "Workbook" Then
        SafeNameScope = "Workbook"
    ElseIf TypeName(p) = "Worksheet" Then
        SafeNameScope = "Worksheet: " & p.Name
    Else
        SafeNameScope = TypeName(p)
    End If

End Function

Private Function SafeNameVisible(ByVal nm As Name) As Boolean

    On Error Resume Next
    SafeNameVisible = nm.Visible

    If Err.Number <> 0 Then
        Err.Clear
        SafeNameVisible = False
    End If

    On Error GoTo 0

End Function

Private Function SafeRefersTo(ByVal nm As Name) As String

    On Error Resume Next
    SafeRefersTo = CStr(nm.RefersTo)

    If Err.Number <> 0 Then
        Err.Clear
        SafeRefersTo = "<Unable to resolve>"
    End If

    On Error GoTo 0

End Function

Private Function SafeTargetSheet(ByVal nm As Name) As String

    Dim r As Range

    On Error Resume Next
    Set r = nm.RefersToRange
    On Error GoTo 0

    If r Is Nothing Then
        SafeTargetSheet = ""
    Else
        SafeTargetSheet = r.Worksheet.Name
    End If

End Function

Private Function IsPossibleExternalReference(ByVal refersToText As String) As Boolean

    IsPossibleExternalReference = _
        (InStr(1, refersToText, "[", vbTextCompare) > 0 And _
         InStr(1, refersToText, "]", vbTextCompare) > 0)

End Function

Private Function SafeNameComment(ByVal nm As Name) As String

    On Error Resume Next
    SafeNameComment = CStr(nm.Comment)

    If Err.Number <> 0 Then
        Err.Clear
        SafeNameComment = ""
    End If

    On Error GoTo 0

End Function

Private Function YesNo(ByVal conditionValue As Boolean) As String

    If conditionValue Then
        YesNo = "Yes"
    Else
        YesNo = "No"
    End If

End Function

Private Sub FormatNameSheet(ByVal ws As Worksheet, ByVal lastRow As Long)

    With ws.Rows(1)
        .Font.Bold = True
        .AutoFilter
    End With

    ws.Columns("A:D").EntireColumn.AutoFit
    ws.Columns("E:E").ColumnWidth = 80
    ws.Columns("F:H").EntireColumn.AutoFit
    ws.Range("E2:E" & lastRow).WrapText = True
    ws.Range("A1:H" & lastRow).VerticalAlignment = xlTop

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True

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

Attribute VB_Name = "modProfiCommon"
Option Explicit

Public Const PROFI_VERSION As String = "1.1.0"

Public Function ProfiHostWorkbook() As Workbook
    Dim wb As Workbook
    If Not Application.ActiveWorkbook Is Nothing Then
        If Not Application.ActiveWorkbook Is ThisWorkbook And Not Application.ActiveWorkbook.IsAddin Then Set ProfiHostWorkbook = Application.ActiveWorkbook: Exit Function
    End If
    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook And Not wb.IsAddin Then Set ProfiHostWorkbook = wb: Exit Function
    Next
End Function

Public Function ProfiText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then ProfiText = "": Exit Function
    ProfiText = Trim$(CStr(value))
    Do While InStr(ProfiText, "  ") > 0: ProfiText = Replace(ProfiText, "  ", " "): Loop
End Function

Public Function ProfiId(ByVal prefix As String, ParamArray values() As Variant) As String
    Dim i As Long, s As String, h As Double
    For i = LBound(values) To UBound(values): s = s & "|" & ProfiText(values(i)): Next
    For i = 1 To Len(s): h = (h * 131 + AscW(Mid$(s, i, 1))) Mod 2147483647#: Next
    ProfiId = prefix & "-" & CStr(Fix(h))
End Function

Public Function ProfiGetSheet(ByVal name As String, Optional ByVal createIt As Boolean = True) As Worksheet
    Dim wb As Workbook: Set wb = ProfiHostWorkbook(): If wb Is Nothing Then Exit Function
    On Error Resume Next: Set ProfiGetSheet = wb.Worksheets(name): On Error GoTo 0
    If ProfiGetSheet Is Nothing And createIt Then
        Set ProfiGetSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ProfiGetSheet.name = name
    End If
End Function

Public Function ProfiEnsureTable(ByVal sheetName As String, ByVal tableName As String, ByVal headers As Variant) As ListObject
    Dim ws As Worksheet, lo As ListObject, c As Long, startRow As Long
    Set ws = ProfiGetSheet(sheetName, True): If ws Is Nothing Then Exit Function
    On Error Resume Next: Set lo = ws.ListObjects(tableName): On Error GoTo 0
    If lo Is Nothing Then
        startRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ProfiText(ws.Cells(startRow, 1).value) <> "" Then startRow = startRow + 2
        For c = LBound(headers) To UBound(headers): ws.Cells(startRow, c + 1).value = headers(c): Next
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, UBound(headers) + 1)), , xlYes): lo.name = tableName
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    End If
    Set ProfiEnsureTable = lo
End Function

Public Function ProfiFindHeader(ByVal lo As ListObject, ByVal header As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(ProfiText(lo.HeaderRowRange.Cells(1, i).value), header, vbTextCompare) = 0 Then ProfiFindHeader = i: Exit Function
    Next
End Function

Public Sub ProfiSetHidden(ByVal visible As Boolean)
    Dim names, i As Long, ws As Worksheet: names = Array("_PROFI_SETTINGS", "_PROFI_TEACHERS", "_PROFI_SOURCES", "_PROFI_DATA", "_PROFI_CALC", "_PROFI_LOG")
    For i = LBound(names) To UBound(names)
        Set ws = ProfiGetSheet(CStr(names(i)), True)
        If Not ws Is Nothing Then If visible Then ws.visible = xlSheetVisible Else ws.visible = xlSheetVeryHidden
    Next
End Sub

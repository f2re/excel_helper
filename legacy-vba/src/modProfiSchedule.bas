Attribute VB_Name = "modProfiSchedule"
Option Explicit

Private Function ST(ByVal tableName As String) As ListObject
    Dim wb As Workbook, ws As Worksheet: Set wb = ProfiHostWorkbook(): If wb Is Nothing Then Exit Function
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set ST = ws.ListObjects(tableName)
        If Not ST Is Nothing Then Exit Function
    Next
    On Error GoTo 0
End Function

Private Function SC(ByVal lo As ListObject, ByVal header As String) As Long
    SC = ProfiFindHeader(lo, header)
End Function

Private Function SV(ByVal lo As ListObject, ByVal rowNo As Long, ByVal header As String) As Variant
    Dim c As Long: c = SC(lo, header): If c > 0 Then SV = lo.DataBodyRange.Cells(rowNo, c).Value
End Function

Private Sub SS(ByVal lo As ListObject, ByVal rowNo As Long, ByVal header As String, ByVal value As Variant)
    Dim c As Long: c = SC(lo, header): If c > 0 Then lo.DataBodyRange.Cells(rowNo, c).Value = value
End Sub

Private Sub ClearT(ByVal lo As ListObject)
    If Not lo Is Nothing Then If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
End Sub

Private Function Marker(ByVal ws As Worksheet, ByVal needle As String) As Range
    Dim r As Long, c As Long
    For r = 1 To WorksheetFunction.Min(ws.UsedRange.Rows.Count, 500)
        For c = 1 To WorksheetFunction.Min(ws.UsedRange.Columns.Count, 30)
            If InStr(1, LCase$(ProfiText(ws.Cells(r, c).Value)), LCase$(needle), vbTextCompare) > 0 Then Set Marker = ws.Cells(r, c): Exit Function
        Next
    Next
End Function

Private Function AskN(ByVal prompt As String, ByVal defaultValue As Long) As Long
    Dim v As Variant: v = Application.InputBox(prompt, "ПрофиПомощник — парсер", defaultValue, Type:=1)
    If VarType(v) = vbBoolean And v = False Then AskN = defaultValue Else AskN = CLng(v)
End Function

Private Function FirstWeekCol(ByVal ws As Worksheet, ByVal weeksRow As Long) As Long
    Dim c As Long
    For c = 1 To ws.UsedRange.Columns.Count
        If IsNumeric(ws.Cells(weeksRow, c).Value) Then If CLng(ws.Cells(weeksRow, c).Value) = 1 Then FirstWeekCol = c: Exit Function
    Next
End Function

Private Function NextSourceId() As String
    NextSourceId = "SRC-" & Format$(Now, "yyyymmddhhnnss")
End Function

Public Sub ProfiConfigureParser()
    Dim wb As Workbook, ws As Worksheet, w As Range, l As Range, g As Range
    Dim sourceId As String, groupName As String, groupMode As String, groupAddress As String
    Dim weeksRow As Long, weekCol As Long, legendRow As Long, answer As VbMsgBoxResult
    Dim sources As ListObject, profiles As ListObject, row As ListRow
    ProfiEnsureProject: Set wb = ProfiHostWorkbook(): If wb Is Nothing Then MsgBox "Откройте пользовательскую книгу.", vbExclamation: Exit Sub
    Set ws = wb.ActiveSheet
    If Left$(ws.Name, 7) = "_PROFI" Then MsgBox "Откройте групповой лист.", vbExclamation: Exit Sub

    answer = MsgBox("Да — выбрать ячейку группы." & vbCrLf & "Нет — ввести вручную." & vbCrLf & "Отмена — имя листа.", vbYesNoCancel + vbQuestion, "Название группы")
    If answer = vbYes Then
        On Error Resume Next: Set g = Application.InputBox("Выберите одну ячейку", "Группа", Type:=8): On Error GoTo 0
        If g Is Nothing Then Exit Sub
        groupName = ProfiText(g.Value): groupMode = "cell": groupAddress = g.Address(False, False)
    ElseIf answer = vbNo Then
        groupName = ProfiText(InputBox("Введите группу", "Группа", ws.Name)): If groupName = "" Then Exit Sub
        groupMode = "manual"
    Else
        groupName = ws.Name: groupMode = "sheet"
    End If

    Set w = Marker(ws, "уч. недел"): If w Is Nothing Then weeksRow = 1 Else weeksRow = w.Row
    weekCol = FirstWeekCol(ws, weeksRow): If weekCol = 0 Then weekCol = 3
    Set l = Marker(ws, "обозн"): If l Is Nothing Then legendRow = ws.UsedRange.Rows.Count + 1 Else legendRow = l.Row
    weeksRow = AskN("Строка учебных недель", weeksRow): weekCol = AskN("Первый столбец недели", weekCol)
    legendRow = AskN("Строка легенды", legendRow)

    sourceId = NextSourceId(): Set sources = ST("tblSources"): Set row = sources.ListRows.Add
    SS sources, row.Index, "SourceID", sourceId: SS sources, row.Index, "Группа", groupName: SS sources, row.Index, "Лист", ws.Name
    SS sources, row.Index, "Режим группы", groupMode: SS sources, row.Index, "Ячейка группы", groupAddress: SS sources, row.Index, "Статус", "Подключён": SS sources, row.Index, "Обновлён", Now
    Set profiles = ST("tblParserProfiles"): Set row = profiles.ListRows.Add
    SS profiles, row.Index, "SourceID", sourceId: SS profiles, row.Index, "WeeksRow", weeksRow: SS profiles, row.Index, "WeekStartCol", weekCol
    SS profiles, row.Index, "MonthsRow", weeksRow + 1: SS profiles, row.Index, "GridStartRow", weeksRow + 3: SS profiles, row.Index, "LegendRow", legendRow
    SS profiles, row.Index, "Days", AskN("Количество учебных дней", 6): SS profiles, row.Index, "Pairs", AskN("Количество пар", 4)
    SS profiles, row.Index, "RowsPerDay", AskN("Строк на день", 13): SS profiles, row.Index, "RowsPerPair", AskN("Строк на пару", 3)
    SS profiles, row.Index, "CodeOffset", 0: SS profiles, row.Index, "SubjectOffset", 1: SS profiles, row.Index, "RoomOffset", 2: SS profiles, row.Index, "DateOffset", -1
    SS profiles, row.Index, "AbbrCol", 1: SS profiles, row.Index, "LecturerCol", 10: SS profiles, row.Index, "OtherCol", 14
    MsgBox "Источник добавлен. Все параметры можно изменить в tblParserProfiles.", vbInformation
End Sub

Private Function ProfileRow(ByVal sourceId As String) As Long
    Dim lo As ListObject, i As Long: Set lo = ST("tblParserProfiles"): If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count: If ProfiText(SV(lo, i, "SourceID")) = sourceId Then ProfileRow = i: Exit Function
    Next
End Function

Private Function TeacherId(ByVal rawName As String) As String
    Dim aliases As ListObject, teachers As ListObject, i As Long, key As String, candidate As String, lastId As String, count As Long, s As String
    key = LCase$(Replace(Replace(ProfiText(rawName), ".", ""), " ", "")): If key = "" Then Exit Function
    Set aliases = ST("tblTeacherAliases")
    If Not aliases Is Nothing Then If Not aliases.DataBodyRange Is Nothing Then
        For i = 1 To aliases.ListRows.Count
            candidate = LCase$(Replace(Replace(ProfiText(SV(aliases, i, "Вариант")), ".", ""), " ", ""))
            If candidate = key Then TeacherId = ProfiText(SV(aliases, i, "TeacherID")): Exit Function
        Next
    End If
    Set teachers = ST("tblTeachers"): If teachers Is Nothing Or teachers.DataBodyRange Is Nothing Then Exit Function
    s = LCase$(Split(ProfiText(rawName) & " ", " ")(0))
    For i = 1 To teachers.ListRows.Count
        If LCase$(ProfiText(SV(teachers, i, "Активен"))) <> "false" Then
            candidate = LCase$(Replace(Replace(ProfiText(SV(teachers, i, "ФИО")), ".", ""), " ", ""))
            If candidate = key Then TeacherId = ProfiText(SV(teachers, i, "TeacherID")): Exit Function
            If LCase$(Split(ProfiText(SV(teachers, i, "ФИО")) & " ", " ")(0)) = s Then lastId = ProfiText(SV(teachers, i, "TeacherID")): count = count + 1
        End If
    Next
    If count = 1 Then TeacherId = lastId
End Function

Private Function CandidateList(ByVal ws As Worksheet, ByVal legendRow As Long, ByVal abbrCol As Long, ByVal lecturerCol As Long, ByVal otherCol As Long, ByVal subject As String, ByVal lessonType As String) As Collection
    Dim result As New Collection, r As Long, raw As String, p, i As Long, id As String
    For r = legendRow + 3 To ws.UsedRange.Rows.Count
        If ProfiText(ws.Cells(r, abbrCol).Value) = "" Then Exit For
        If StrComp(ProfiText(ws.Cells(r, abbrCol).Value), subject, vbTextCompare) = 0 Then
            If UCase$(lessonType) = "Л" Then raw = ProfiText(ws.Cells(r, lecturerCol).Value) Else raw = ProfiText(ws.Cells(r, otherCol).Value)
            raw = Replace(Replace(raw, vbCrLf, ";"), ",", ";"): p = Split(raw, ";")
            For i = LBound(p) To UBound(p)
                id = TeacherId(ProfiText(p(i)))
                If id <> "" Then On Error Resume Next: result.Add id, id: On Error GoTo 0
            Next
            Exit For
        End If
    Next
    Set CandidateList = result
End Function

Private Function DayCode(ByVal index As Long) As String
    DayCode = Array("Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс")(index)
End Function

Private Function LessonTypeOf(ByVal code As String) As String
    Dim p: p = Split(ProfiText(code), "/"): LessonTypeOf = UCase$(ProfiText(p(0))): If LessonTypeOf = "" Then LessonTypeOf = "П"
End Function

Private Sub AddConflict(ByVal lo As ListObject, ByVal kind As String, ByVal teacher As String, ByVal weekNo As Long, ByVal dayName As String, ByVal pairNo As Long, ByVal description As String)
    Dim r As ListRow: Set r = lo.ListRows.Add
    SS lo, r.Index, "ConflictID", ProfiId("C", kind, teacher, weekNo, dayName, pairNo): SS lo, r.Index, "Уровень", "Критический": SS lo, r.Index, "Тип", kind
    SS lo, r.Index, "TeacherID", teacher: SS lo, r.Index, "Неделя", weekNo: SS lo, r.Index, "День", dayName: SS lo, r.Index, "Пара", pairNo: SS lo, r.Index, "Описание", description: SS lo, r.Index, "Исправлено", False
End Sub

Private Function SettingFlag(ByVal key As String, ByVal defaultValue As Boolean) As Boolean
    Dim lo As ListObject, i As Long: SettingFlag = defaultValue: Set lo = ST("tblScheduleSettings"): If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If LCase$(ProfiText(SV(lo, i, "Ключ"))) = LCase$(key) Then SettingFlag = LCase$(ProfiText(SV(lo, i, "Значение"))) <> "false": Exit Function
    Next
End Function

Public Sub ProfiComposeSchedule()
    Dim wb As Workbook, sources As ListObject, profiles As ListObject, lessons As ListObject, assignments As ListObject, conflicts As ListObject, candidatesTable As ListObject
    Dim occupancy As Object, streams As Object, rr As Object, i As Long, pr As Long, ws As Worksheet, sourceId As String, groupName As String, mode As String, address As String
    Dim weeksRow As Long, weekCol As Long, gridRow As Long, legendRow As Long, dayCount As Long, pairCount As Long, rowsDay As Long, rowsPair As Long, codeOff As Long, subjectOff As Long, roomOff As Long, abbrCol As Long, lecturerCol As Long, otherCol As Long
    Dim d As Long, pairNo As Long, c As Long, r As Long, weekNo As Long, code As String, subject As String, room As String, kind As String, lessonId As String, slot As String, streamKey As String, rrKey As String
    Dim list As Collection, item As Variant, chosen As String, lr As ListRow, ar As ListRow, cr As ListRow, strict As Boolean, useStreams As Boolean, start As Long, j As Long
    Application.ScreenUpdating = False: Application.EnableEvents = False: On Error GoTo Failed
    ProfiEnsureProject: Set wb = ProfiHostWorkbook(): If wb Is Nothing Then GoTo Clean
    Set sources = ST("tblSources"): If sources Is Nothing Or sources.DataBodyRange Is Nothing Then ProfiConfigureParser: Set sources = ST("tblSources")
    If sources Is Nothing Or sources.DataBodyRange Is Nothing Then GoTo Clean
    Set profiles = ST("tblParserProfiles"): Set lessons = ST("tblLessons"): Set assignments = ST("tblAssignments"): Set conflicts = ST("tblConflicts"): Set candidatesTable = ST("tblTeacherCandidates")
    ClearT lessons: ClearT assignments: ClearT conflicts: ClearT candidatesTable
    Set occupancy = CreateObject("Scripting.Dictionary"): Set streams = CreateObject("Scripting.Dictionary"): Set rr = CreateObject("Scripting.Dictionary")
    strict = SettingFlag("StrictConflicts", True): useStreams = SettingFlag("EnableStreams", True)
    For i = 1 To sources.ListRows.Count
        sourceId = ProfiText(SV(sources, i, "SourceID")): groupName = ProfiText(SV(sources, i, "Группа")): mode = LCase$(ProfiText(SV(sources, i, "Режим группы"))): address = ProfiText(SV(sources, i, "Ячейка группы"))
        Set ws = Nothing: On Error Resume Next: Set ws = wb.Worksheets(ProfiText(SV(sources, i, "Лист"))): On Error GoTo Failed
        If ws Is Nothing Then SS sources, i, "Статус", "Лист не найден": GoTo NextSource
        If mode = "cell" And address <> "" Then groupName = ProfiText(ws.Range(address).Value)
        If mode = "sheet" Then groupName = ws.Name
        pr = ProfileRow(sourceId): If pr = 0 Then SS sources, i, "Статус", "Нет профиля": GoTo NextSource
        weeksRow = CLng(SV(profiles, pr, "WeeksRow")): weekCol = CLng(SV(profiles, pr, "WeekStartCol")): gridRow = CLng(SV(profiles, pr, "GridStartRow")): legendRow = CLng(SV(profiles, pr, "LegendRow"))
        dayCount = CLng(SV(profiles, pr, "Days")): pairCount = CLng(SV(profiles, pr, "Pairs")): rowsDay = CLng(SV(profiles, pr, "RowsPerDay")): rowsPair = CLng(SV(profiles, pr, "RowsPerPair"))
        codeOff = CLng(SV(profiles, pr, "CodeOffset")): subjectOff = CLng(SV(profiles, pr, "SubjectOffset")): roomOff = CLng(SV(profiles, pr, "RoomOffset")): abbrCol = CLng(SV(profiles, pr, "AbbrCol")): lecturerCol = CLng(SV(profiles, pr, "LecturerCol")): otherCol = CLng(SV(profiles, pr, "OtherCol"))
        For d = 0 To dayCount - 1: For pairNo = 1 To pairCount
            r = gridRow + d * rowsDay + (pairNo - 1) * rowsPair
            For c = weekCol To ws.UsedRange.Columns.Count
                If IsNumeric(ws.Cells(weeksRow, c).Value) Then
                    weekNo = CLng(ws.Cells(weeksRow, c).Value): If weekNo <= 0 Then GoTo SkipCell
                    code = ProfiText(ws.Cells(r + codeOff, c).Value): subject = ProfiText(ws.Cells(r + subjectOff, c).Value): room = ProfiText(ws.Cells(r + roomOff, c).Value): If code = "" And subject = "" Then GoTo SkipCell
                    kind = LessonTypeOf(code): lessonId = ProfiId("L", sourceId, groupName, weekNo, DayCode(d), pairNo, r, c)
                    Set lr = lessons.ListRows.Add: SS lessons, lr.Index, "LessonID", lessonId: SS lessons, lr.Index, "SourceID", sourceId: SS lessons, lr.Index, "Группа", groupName: SS lessons, lr.Index, "Неделя", weekNo: SS lessons, lr.Index, "День", DayCode(d): SS lessons, lr.Index, "Пара", pairNo: SS lessons, lr.Index, "Код", code: SS lessons, lr.Index, "Вид", kind: SS lessons, lr.Index, "Дисциплина", subject: SS lessons, lr.Index, "Аудитория", room: SS lessons, lr.Index, "Ячейка", ws.Cells(r, c).Address(False, False): SS lessons, lr.Index, "Статус", "Готово"
                    Set list = CandidateList(ws, legendRow, abbrCol, lecturerCol, otherCol, subject, kind): slot = CStr(weekNo) & "|" & DayCode(d) & "|" & CStr(pairNo): streamKey = slot & "|" & LCase$(subject) & "|" & LCase$(kind) & "|" & LCase$(room): chosen = ""
                    For Each item In list: Set cr = candidatesTable.ListRows.Add: SS candidatesTable, cr.Index, "CandidateID", ProfiId("TC", lessonId, CStr(item)): SS candidatesTable, cr.Index, "LessonID", lessonId: SS candidatesTable, cr.Index, "SourceID", sourceId: SS candidatesTable, cr.Index, "Дисциплина", subject: SS candidatesTable, cr.Index, "Вид", kind: SS candidatesTable, cr.Index, "TeacherID", CStr(item): SS candidatesTable, cr.Index, "Метод", "Легенда": SS candidatesTable, cr.Index, "Приоритет", 100: Next
                    If useStreams And streams.Exists(streamKey) Then chosen = streams(streamKey)
                    If chosen = "" And list.Count > 0 Then
                        rrKey = LCase$(subject) & "|" & LCase$(kind): If rr.Exists(rrKey) Then start = rr(rrKey) Else start = 1
                        For j = 0 To list.Count - 1: item = list(((start - 1 + j) Mod list.Count) + 1): If Not occupancy.Exists(slot & "|" & CStr(item)) Then chosen = CStr(item): rr(rrKey) = ((start + j) Mod list.Count) + 1: Exit For
                        Next
                        If chosen = "" And Not strict Then chosen = CStr(list(start))
                    End If
                    Set ar = assignments.ListRows.Add: SS assignments, ar.Index, "AssignmentID", ProfiId("A", lessonId): SS assignments, ar.Index, "LessonID", lessonId: SS assignments, ar.Index, "TeacherID", chosen
                    If chosen = "" Then
                        SS assignments, ar.Index, "Статус", "Не назначено": AddConflict conflicts, "Не назначено", "", weekNo, DayCode(d), pairNo, "Нет свободного преподавателя"
                    Else
                        If streams.Exists(streamKey) Then SS assignments, ar.Index, "Метод", "Поток" Else SS assignments, ar.Index, "Метод", "Round-Robin"
                        If occupancy.Exists(slot & "|" & chosen) And Not streams.Exists(streamKey) Then SS assignments, ar.Index, "Статус", "Конфликт": AddConflict conflicts, "Преподаватель", chosen, weekNo, DayCode(d), pairNo, "Двойная занятость" Else SS assignments, ar.Index, "Статус", "Назначено"
                        occupancy(slot & "|" & chosen) = lessonId: If useStreams And Not streams.Exists(streamKey) Then streams(streamKey) = chosen
                    End If
SkipCell:
                End If
            Next
        Next: Next
        SS sources, i, "Статус", "Обработан": SS sources, i, "Обновлён", Now
NextSource:
    Next
    RenderSummary lessons, assignments: RenderConflicts conflicts: ProfiSetHidden False
    MsgBox "Готово. Занятий: " & lessons.ListRows.Count & ". Конфликтов: " & conflicts.ListRows.Count, vbInformation
Clean:
    Application.EnableEvents = True: Application.ScreenUpdating = True: Exit Sub
Failed:
    Application.EnableEvents = True: Application.ScreenUpdating = True: MsgBox "Ошибка: " & Err.Description, vbCritical
End Sub

Private Function TName(ByVal id As String) As String
    Dim lo As ListObject, i As Long: Set lo = ST("tblTeachers"): If lo Is Nothing Or lo.DataBodyRange Is Nothing Then TName = id: Exit Function
    For i = 1 To lo.ListRows.Count: If ProfiText(SV(lo, i, "TeacherID")) = id Then TName = PROFI_FIO_SHORT(SV(lo, i, "ФИО")): Exit Function
    Next: TName = id
End Function

Private Sub RenderSummary(ByVal lessons As ListObject, ByVal assignments As ListObject)
    Dim ws As Worksheet, i As Long, j As Long, outRow As Long, lessonId As String, teacher As String, method As String
    Set ws = ProfiGetSheet("Сводное расписание", True): ws.Cells.Clear: ws.Range("A1:J1").Value = Array("Неделя", "День", "Пара", "Преподаватель", "Вид", "Дисциплина", "Группа", "Аудитория", "LessonID", "Метод"): ws.Rows(1).Font.Bold = True: outRow = 2
    If lessons.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lessons.ListRows.Count
        lessonId = ProfiText(SV(lessons, i, "LessonID")): teacher = "": method = ""
        If Not assignments.DataBodyRange Is Nothing Then For j = 1 To assignments.ListRows.Count: If ProfiText(SV(assignments, j, "LessonID")) = lessonId Then teacher = ProfiText(SV(assignments, j, "TeacherID")): method = ProfiText(SV(assignments, j, "Метод")): Exit For
        Next
        ws.Cells(outRow, 1).Value = SV(lessons, i, "Неделя"): ws.Cells(outRow, 2).Value = SV(lessons, i, "День"): ws.Cells(outRow, 3).Value = SV(lessons, i, "Пара"): ws.Cells(outRow, 4).Value = TName(teacher): ws.Cells(outRow, 5).Value = SV(lessons, i, "Вид"): ws.Cells(outRow, 6).Value = SV(lessons, i, "Дисциплина"): ws.Cells(outRow, 7).Value = SV(lessons, i, "Группа"): ws.Cells(outRow, 8).Value = SV(lessons, i, "Аудитория"): ws.Cells(outRow, 9).Value = lessonId: ws.Cells(outRow, 10).Value = method: outRow = outRow + 1
    Next
    ws.Columns.AutoFit: ws.Range("A1:J" & outRow - 1).AutoFilter
End Sub

Private Sub RenderConflicts(ByVal conflicts As ListObject)
    Dim ws As Worksheet: Set ws = ProfiGetSheet("Контроль расписания", True): ws.Cells.Clear: ws.Range("A1:I1").Value = conflicts.HeaderRowRange.Value: ws.Rows(1).Font.Bold = True
    If Not conflicts.DataBodyRange Is Nothing Then ws.Range("A2").Resize(conflicts.ListRows.Count, conflicts.ListColumns.Count).Value = conflicts.DataBodyRange.Value
    ws.Columns.AutoFit
End Sub

Attribute VB_Name = "modProfiProject"
Option Explicit

Public Sub ProfiEnsureProject()
    Application.ScreenUpdating = False
    ProfiEnsureTable "_PROFI_SETTINGS", "tblScheduleSettings", Array("Ключ", "Значение", "Описание")
    ProfiEnsureTable "_PROFI_SETTINGS", "tblParserProfiles", Array("SourceID", "WeeksRow", "WeekStartCol", "MonthsRow", "GridStartRow", "LegendRow", "Days", "Pairs", "RowsPerDay", "RowsPerPair", "CodeOffset", "SubjectOffset", "RoomOffset", "DateOffset", "AbbrCol", "LecturerCol", "OtherCol")
    ProfiEnsureTable "_PROFI_TEACHERS", "tblTeachers", Array("TeacherID", "ФИО", "Краткое имя", "Должность", "Звание", "Степень", "Макс. часов", "Порядок", "Активен")
    ProfiEnsureTable "_PROFI_TEACHERS", "tblTeacherAliases", Array("AliasID", "Вариант", "TeacherID", "Приоритет", "Активен")
    ProfiEnsureTable "_PROFI_SOURCES", "tblSources", Array("SourceID", "Группа", "Лист", "Режим группы", "Ячейка группы", "Статус", "Обновлён")
    ProfiEnsureTable "_PROFI_DATA", "tblLessons", Array("LessonID", "SourceID", "Группа", "Неделя", "День", "Пара", "Дата", "Код", "Вид", "Дисциплина", "Аудитория", "Ячейка", "Статус")
    ProfiEnsureTable "_PROFI_DATA", "tblTeacherCandidates", Array("CandidateID", "LessonID", "SourceID", "Дисциплина", "Вид", "TeacherID", "Исходное имя", "Метод", "Приоритет")
    ProfiEnsureTable "_PROFI_DATA", "tblCalendar", Array("Неделя", "День", "Дата", "Месяц", "Рабочий")
    ProfiEnsureTable "_PROFI_CALC", "tblAssignments", Array("AssignmentID", "LessonID", "TeacherID", "Метод", "ПотокID", "Закреплено", "Статус")
    ProfiEnsureTable "_PROFI_CALC", "tblOccupancy", Array("Неделя", "День", "Пара", "TeacherID", "Дисциплина", "AssignmentID")
    ProfiEnsureTable "_PROFI_CALC", "tblConflicts", Array("ConflictID", "Уровень", "Тип", "TeacherID", "Неделя", "День", "Пара", "Описание", "Исправлено")
    ProfiEnsureTable "_PROFI_CALC", "tblManualOverrides", Array("LessonID", "TeacherID", "Режим", "Комментарий")
    ProfiEnsureTable "_PROFI_LOG", "tblOperationLog", Array("Время", "Операция", "Статус", "Подробности")
    ProfiEnsureTable "_PROFI_LOG", "tblScheduleVersions", Array("VersionID", "Время", "AssignmentsJSON", "Комментарий")
    ProfiCreateDashboard: ProfiSetHidden False: Application.ScreenUpdating = True
End Sub

Public Sub ProfiCreateDashboard()
    Dim ws As Worksheet: Set ws = ProfiGetSheet("ПрофиПомощник", True): ws.Cells.Clear
    ws.Range("A1").value = "🧰 ПрофиПомощник для Excel": ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 18
    ws.Range("A3:A8").value = Application.Transpose(Array("1. Заполните справочник преподавателей", "2. Откройте лист группового расписания", "3. Запустите «Составить сводное расписание»", "4. Укажите группу вручную или выберите ячейку", "5. Проверьте профиль парсера", "6. Исправьте конфликты и повторите расчёт"))
    ws.Columns("A:A").ColumnWidth = 85
End Sub

Public Sub ProfiFioWizard()
    Dim c As Range: For Each c In Selection.Cells: If Not c.HasFormula Then c.value = PROFI_FIO_NORMALIZE(c.value)
    Next
End Sub
Public Sub ProfiDuplicateManager()
    Dim c As Range, d As Object: Set d = CreateObject("Scripting.Dictionary")
    For Each c In Selection.Cells: If ProfiText(c.value) <> "" Then If d.Exists(LCase$(ProfiText(c.value))) Then c.Interior.Color = RGB(255, 199, 206) Else d.Add LCase$(ProfiText(c.value)), True
    Next
End Sub
Public Sub ProfiWeeklyPlanner()
    Dim ws As Worksheet: Set ws = ProfiGetSheet("Недельный план", True): ws.Cells.Clear
    ws.Range("A1:I1").value = Array("Дата", "День", "Задача", "Исполнитель", "Приоритет", "Статус", "Срок", "Прогресс", "Комментарий"): ws.Rows(1).Font.Bold = True: ws.Columns.AutoFit
End Sub

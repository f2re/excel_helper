Attribute VB_Name = "modProfiTemplate"
Option Explicit

Public Sub ProfiTemplateOnOpen()
    On Error GoTo Failed
    Application.ScreenUpdating = False
    ProfiPrepareTemplateHome
Clean:
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    MsgBox "Не удалось подготовить стартовую страницу: " & Err.Description, vbExclamation, "ПрофиПомощник"
    Resume Clean
End Sub

Public Sub ProfiTemplateBeforeClose()
    On Error Resume Next
    Application.StatusBar = False
End Sub

Public Sub ProfiPrepareTemplateHome()
    Dim wb As Workbook, ws As Worksheet
    Set wb = ProfiHostWorkbook()
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets("Старт")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = "Старт"
    End If
    ws.Range("A1").Value = "ПрофиПомощник для Excel"
    ws.Range("A2").Value = "Переносимый шаблон XLTM · версия " & PROFI_VERSION
    ws.Range("A4").Value = "1. Создайте проект. 2. Откройте лист группового расписания. 3. Подключите источник. 4. Составьте расписание."
    ws.Range("A5").Value = "Название группы можно ввести вручную, взять из имени листа или из выбранной ячейки. Все параметры парсера редактируются."
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 22
    ws.Range("A2").Font.Size = 12
    ws.Range("A4:A5").WrapText = True
    ws.Columns("A:A").ColumnWidth = 92
End Sub

Public Sub ProfiTemplateCreateProject()
    ProfiEnsureProject
    Dim ws As Worksheet
    Set ws = ProfiGetSheet("ПрофиПомощник", True)
    If Not ws Is Nothing Then ws.Activate
End Sub

Public Sub ProfiTemplateAddSource()
    Dim wb As Workbook, ws As Worksheet
    Set wb = ProfiHostWorkbook()
    If wb Is Nothing Then MsgBox "Не найдена рабочая книга.", vbExclamation: Exit Sub
    Set ws = wb.ActiveSheet
    If ws.Name = "Старт" Or Left$(ws.Name, 7) = "_PROFI" Or ws.Name = "ПрофиПомощник" Then
        MsgBox "Сначала откройте или скопируйте в книгу лист группового расписания, затем запустите эту команду ещё раз.", vbInformation, "Подключение расписания"
        Exit Sub
    End If
    ProfiConfigureParser
End Sub

Public Sub ProfiTemplateComposeSchedule()
    ProfiComposeSchedule
End Sub

Public Sub ProfiTemplateOpenTeachers()
    ProfiEnsureProject
    Dim ws As Worksheet
    Set ws = ProfiGetSheet("_PROFI_TEACHERS", True)
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVisible
        ws.Activate
    End If
End Sub

Public Sub ProfiTemplateToggleServiceSheets()
    Dim ws As Worksheet, showSheets As Boolean
    Set ws = ProfiGetSheet("_PROFI_SETTINGS", False)
    If ws Is Nothing Then
        ProfiEnsureProject
        Set ws = ProfiGetSheet("_PROFI_SETTINGS", False)
    End If
    If ws Is Nothing Then Exit Sub
    showSheets = (ws.Visible <> xlSheetVisible)
    ProfiSetHidden showSheets
    If showSheets Then
        Set ws = ProfiGetSheet("_PROFI_SETTINGS", False)
        If Not ws Is Nothing Then ws.Activate
        MsgBox "Служебные листы показаны. После редактирования повторите команду, чтобы скрыть их.", vbInformation
    Else
        Set ws = ProfiGetSheet("Старт", False)
        If Not ws Is Nothing Then ws.Activate
    End If
End Sub

Public Sub ProfiTemplateShowHelp()
    MsgBox "Порядок работы:" & vbCrLf & _
        "1. Создайте проект." & vbCrLf & _
        "2. Откройте лист группового расписания." & vbCrLf & _
        "3. Подключите источник и подтвердите координаты парсера." & vbCrLf & _
        "4. Заполните преподавателей и алиасы." & vbCrLf & _
        "5. Составьте расписание и проверьте лист контроля." & vbCrLf & vbCrLf & _
        "Документация: https://github.com/f2re/excel_helper", vbInformation, "ПрофиПомощник — справка"
End Sub

Attribute VB_Name = "modProfiRibbon"
Option Explicit

Private gProfiRibbon As Object

Public Sub ProfiRibbonOnLoad(ByVal ribbon As Object)
    Set gProfiRibbon = ribbon
End Sub

Public Sub ProfiRibbonEnsureProject(ByVal control As Object)
    ProfiEnsureProject
End Sub

Public Sub ProfiRibbonComposeSchedule(ByVal control As Object)
    ProfiComposeSchedule
End Sub

Public Sub ProfiRibbonConfigureParser(ByVal control As Object)
    ProfiConfigureParser
End Sub

Public Sub ProfiRibbonFioWizard(ByVal control As Object)
    ProfiFioWizard
End Sub

Public Sub ProfiRibbonDuplicateManager(ByVal control As Object)
    ProfiDuplicateManager
End Sub

Public Sub ProfiRibbonWeeklyPlanner(ByVal control As Object)
    ProfiWeeklyPlanner
End Sub

Public Sub ProfiRibbonShowDashboard(ByVal control As Object)
    Dim ws As Worksheet
    ProfiEnsureProject
    Set ws = ProfiGetSheet("ПрофиПомощник", True)
    If Not ws Is Nothing Then ws.Activate
End Sub

Public Sub ProfiRibbonShowService(ByVal control As Object)
    ProfiEnsureProject
    ProfiSetHidden True
    Dim ws As Worksheet
    Set ws = ProfiGetSheet("_PROFI_SETTINGS", False)
    If Not ws Is Nothing Then ws.Activate
End Sub

Public Sub ProfiRibbonHideService(ByVal control As Object)
    ProfiSetHidden False
    Dim ws As Worksheet
    Set ws = ProfiGetSheet("ПрофиПомощник", False)
    If Not ws Is Nothing Then ws.Activate
End Sub

Public Sub ProfiRibbonHelp(ByVal control As Object)
    MsgBox "Документация: https://github.com/f2re/excel_helper" & vbCrLf & vbCrLf & _
        "Для составления расписания откройте групповой лист, настройте парсер и запустите команду составления.", _
        vbInformation, "ПрофиПомощник — справка"
End Sub

Public Sub ProfiRibbonAbout(ByVal control As Object)
    MsgBox "ПрофиПомощник для Excel" & vbCrLf & "Версия " & PROFI_VERSION & vbCrLf & _
        "XLAM / XLTM для Excel 2010–2019 Windows", vbInformation, "О программе"
End Sub

Public Sub ProfiInvalidateRibbon()
    On Error Resume Next
    If Not gProfiRibbon Is Nothing Then gProfiRibbon.Invalidate
    On Error GoTo 0
End Sub

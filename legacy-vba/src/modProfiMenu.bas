Attribute VB_Name = "modProfiMenu"
Option Explicit

Public Sub Auto_Open(): ProfiInstallMenu: End Sub
Public Sub Auto_Close(): ProfiRemoveMenu: End Sub
Public Sub ProfiInstallMenu()
    Dim bar As CommandBar, popup As CommandBarControl
    ProfiRemoveMenu: Set bar = Application.CommandBars("Worksheet Menu Bar"): Set popup = bar.Controls.Add(Type:=msoControlPopup, Temporary:=True): popup.Caption = "ПрофиПомощник"
    AddMenu popup, "Проверить/создать проект", "ProfiEnsureProject": AddMenu popup, "Составить сводное расписание", "ProfiComposeSchedule": AddMenu popup, "Настроить парсер", "ProfiConfigureParser": AddMenu popup, "Мастер ФИО", "ProfiFioWizard": AddMenu popup, "Найти дубли", "ProfiDuplicateManager": AddMenu popup, "Недельный план", "ProfiWeeklyPlanner"
End Sub
Private Sub AddMenu(ByVal parent As CommandBarControl, ByVal caption As String, ByVal macroName As String): Dim b As CommandBarControl: Set b = parent.Controls.Add(Type:=msoControlButton): b.Caption = caption: b.OnAction = macroName: End Sub
Public Sub ProfiRemoveMenu(): On Error Resume Next: Application.CommandBars("Worksheet Menu Bar").Controls("ПрофиПомощник").Delete: On Error GoTo 0: End Sub

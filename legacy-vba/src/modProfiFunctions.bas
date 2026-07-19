Attribute VB_Name = "modProfiFunctions"
Option Explicit

Public Function PROFI_TEXT_NORMALIZE(ByVal value As Variant) As String
    PROFI_TEXT_NORMALIZE = ProfiText(value)
End Function

Public Function PROFI_FIO_NORMALIZE(ByVal value As Variant) As String
    Dim p, i As Long, s As String
    p = Split(LCase$(ProfiText(value)), " ")
    For i = LBound(p) To UBound(p)
        If Len(p(i)) > 0 Then
            p(i) = UCase$(Left$(p(i), 1)) & Mid$(p(i), 2)
            If Len(s) > 0 Then s = s & " "
            s = s & p(i)
        End If
    Next
    PROFI_FIO_NORMALIZE = s
End Function

Public Function PROFI_FIO_SHORT(ByVal value As Variant) As String
    Dim p, s As String: p = Split(PROFI_FIO_NORMALIZE(value), " ")
    If UBound(p) >= 0 Then s = p(0)
    If UBound(p) >= 1 Then s = s & " " & Left$(p(1), 1) & "."
    If UBound(p) >= 2 Then s = s & Left$(p(2), 1) & "."
    PROFI_FIO_SHORT = s
End Function

Public Function PROFI_SURNAME(ByVal value As Variant) As String
    PROFI_SURNAME = Split(PROFI_FIO_NORMALIZE(value) & " ", " ")(0)
End Function

Public Function PROFI_NAME(ByVal value As Variant) As String
    Dim p: p = Split(PROFI_FIO_NORMALIZE(value), " "): If UBound(p) >= 1 Then PROFI_NAME = p(1)
End Function

Public Function PROFI_PATRONYMIC(ByVal value As Variant) As String
    Dim p: p = Split(PROFI_FIO_NORMALIZE(value), " "): If UBound(p) >= 2 Then PROFI_PATRONYMIC = p(2)
End Function

Public Function PROFI_INITIALS(ByVal value As Variant) As String
    Dim p: p = Split(PROFI_FIO_NORMALIZE(value), " ")
    If UBound(p) >= 1 Then PROFI_INITIALS = Left$(p(1), 1) & "."
    If UBound(p) >= 2 Then PROFI_INITIALS = PROFI_INITIALS & Left$(p(2), 1) & "."
End Function

Public Function PROFI_AGE(ByVal birthDate As Variant, Optional ByVal onDate As Variant) As Long
    Dim d As Date
    If IsMissing(onDate) Or Not IsDate(onDate) Then d = Date Else d = CDate(onDate)
    If Not IsDate(birthDate) Then Exit Function
    PROFI_AGE = DateDiff("yyyy", CDate(birthDate), d)
    If DateSerial(Year(d), Month(CDate(birthDate)), Day(CDate(birthDate))) > d Then PROFI_AGE = PROFI_AGE - 1
End Function

Public Function PROFI_TENURE_DAYS(ByVal startDate As Variant, Optional ByVal endDate As Variant) As Long
    If Not IsDate(startDate) Then Exit Function
    If IsMissing(endDate) Or Not IsDate(endDate) Then endDate = Date
    PROFI_TENURE_DAYS = DateDiff("d", CDate(startDate), CDate(endDate))
End Function

Public Function PROFI_POSITION_COUNT(ByVal fio As Variant, ByVal fioRange As Range, ByVal positionRange As Range) As Long
    Dim i As Long, d As Object: Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To WorksheetFunction.Min(fioRange.Cells.Count, positionRange.Cells.Count)
        If StrComp(ProfiText(fioRange.Cells(i).value), ProfiText(fio), vbTextCompare) = 0 Then d(ProfiText(positionRange.Cells(i).value)) = True
    Next
    PROFI_POSITION_COUNT = d.Count
End Function

Public Function PROFI_PERIOD_OVERLAP(ByVal a1 As Variant, ByVal a2 As Variant, ByVal b1 As Variant, ByVal b2 As Variant) As Boolean
    If IsDate(a1) And IsDate(a2) And IsDate(b1) And IsDate(b2) Then PROFI_PERIOD_OVERLAP = CDate(a1) <= CDate(b2) And CDate(b1) <= CDate(a2)
End Function

Public Function PROFI_CONTRACT_STATUS(ByVal endDate As Variant, Optional ByVal warnDays As Long = 30) As String
    If ProfiText(endDate) = "" Then
        PROFI_CONTRACT_STATUS = "Бессрочный"
    ElseIf Not IsDate(endDate) Then
        PROFI_CONTRACT_STATUS = "Ошибка даты"
    ElseIf CDate(endDate) < Date Then
        PROFI_CONTRACT_STATUS = "Просрочен"
    ElseIf DateDiff("d", Date, CDate(endDate)) <= warnDays Then
        PROFI_CONTRACT_STATUS = "Истекает"
    Else
        PROFI_CONTRACT_STATUS = "Действует"
    End If
End Function

Public Function PROFI_PHONE_NORMALIZE(ByVal value As Variant) As String
    Dim s As String, i As Long, ch As String
    For i = 1 To Len(CStr(value)): ch = Mid$(CStr(value), i, 1): If ch Like "#" Then s = s & ch
    Next
    If Len(s) = 11 And Left$(s, 1) = "8" Then s = "7" & Mid$(s, 2)
    If Len(s) = 11 Then PROFI_PHONE_NORMALIZE = "+" & Left$(s, 1) & " (" & Mid$(s, 2, 3) & ") " & Mid$(s, 5, 3) & "-" & Mid$(s, 8, 2) & "-" & Right$(s, 2) Else PROFI_PHONE_NORMALIZE = CStr(value)
End Function

Public Function PROFI_EMAIL_VALID(ByVal value As Variant) As Boolean
    PROFI_EMAIL_VALID = CStr(value) Like "*?@?*.?*"
End Function

Public Function PROFI_EMPLOYEE_NO(ByVal value As Variant, Optional ByVal width As Long = 6) As String
    PROFI_EMPLOYEE_NO = Right$(String$(width, "0") & ProfiText(value), width)
End Function

Public Function PROFI_COMPLETENESS(ByVal values As Range) As Double
    Dim c As Range, n As Long
    For Each c In values.Cells: If ProfiText(c.value) <> "" Then n = n + 1
    Next
    If values.Cells.Count > 0 Then PROFI_COMPLETENESS = n / values.Cells.Count * 100
End Function

Public Function PROFI_DUPLICATE(ByVal value As Variant, ByVal values As Range) As Boolean
    PROFI_DUPLICATE = WorksheetFunction.CountIf(values, value) > 1
End Function

Public Function PROFI_SIMILARITY(ByVal a As Variant, ByVal b As Variant) As Double
    If LCase$(ProfiText(a)) = LCase$(ProfiText(b)) Then PROFI_SIMILARITY = 100 Else PROFI_SIMILARITY = 0
End Function

Public Function PROFI_TASK_STATUS(ByVal status As Variant, ByVal deadline As Variant) As String
    If InStr(1, LCase$(ProfiText(status)), "выполн", vbTextCompare) > 0 Then
        PROFI_TASK_STATUS = "Выполнена"
    ElseIf IsDate(deadline) And CDate(deadline) < Date Then
        PROFI_TASK_STATUS = "Просрочена"
    ElseIf ProfiText(status) = "" Then
        PROFI_TASK_STATUS = "Не начата"
    Else
        PROFI_TASK_STATUS = ProfiText(status)
    End If
End Function

Public Function PROFI_DAYS_TO_DEADLINE(ByVal deadline As Variant) As Long
    If IsDate(deadline) Then PROFI_DAYS_TO_DEADLINE = DateDiff("d", Date, CDate(deadline))
End Function

Public Function PROFI_PLAN_FACT(ByVal planValue As Double, ByVal factValue As Double, Optional ByVal mode As String = "value") As Double
    If LCase$(mode) = "percent" And planValue <> 0 Then PROFI_PLAN_FACT = (factValue - planValue) / planValue * 100 Else PROFI_PLAN_FACT = factValue - planValue
End Function

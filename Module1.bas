Attribute VB_Name = "Module1"
Public TC, TS As String
Function RunCode(ByVal ThisCode As String)
TmpCode = LCase(ThisCode)
a = InStr(TmpCode, "printl ")
If a = 1 Then
    tmp1 = Mid(ThisCode, 8, 1)
    If tmp1 = "+" Then
        tmp2 = Mid(ThisCode, 9)
        tmp3 = Split(tmp2, ",")
        tmp4 = Val(tmp3(0)) + Val(tmp3(1))
        TC = "ot"
        TS = tmp4
    ElseIf tmp1 = "-" Then
        tmp2 = Mid(ThisCode, 9)
        tmp3 = Split(tmp2, ",")
        tmp4 = Val(tmp3(0)) - Val(tmp3(1))
        TC = "ot"
        TS = tmp4
    ElseIf tmp1 = "*" Then
        tmp2 = Mid(ThisCode, 9)
        tmp3 = Split(tmp2, ",")
        tmp4 = Val(tmp3(0)) * Val(tmp3(1))
        TC = "ot"
        TS = tmp4
    ElseIf tmp1 = "/" Then
        tmp2 = Mid(ThisCode, 9)
        tmp3 = Split(tmp2, ",")
        tmp4 = Val(tmp3(0)) / Val(tmp3(1))
        TC = "ot"
        TS = tmp4
    Else
        TC = "ot"
        TS = Mid(ThisCode, 8)
    End If
End If
a = InStr(TmpCode, "ifobox ")
If a = 1 Then
    tmp1 = Mid(ThisCode, 8, 1)
    If tmp1 = "+" Then
        tmp2 = Mid(ThisCode, 9)
        tmp3 = Split(tmp2, ",")
        tmp4 = Val(tmp3(0)) + Val(tmp3(1))
        MsgBox tmp4
    ElseIf tmp1 = "-" Then
        tmp2 = Mid(ThisCode, 9)
        tmp3 = Split(tmp2, ",")
        tmp4 = Val(tmp3(0)) - Val(tmp3(1))
        MsgBox tmp4
    ElseIf tmp1 = "*" Then
        tmp2 = Mid(ThisCode, 9)
        tmp3 = Split(tmp2, ",")
        tmp4 = Val(tmp3(0)) * Val(tmp3(1))
        MsgBox tmp4
    ElseIf tmp1 = "/" Then
        tmp2 = Mid(ThisCode, 9)
        tmp3 = Split(tmp2, ",")
        tmp4 = Val(tmp3(0)) / Val(tmp3(1))
        MsgBox tmp4
    Else
        MsgBox Mid(ThisCode, 8)
    End If
End If
End Function

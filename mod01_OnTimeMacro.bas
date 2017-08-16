Attribute VB_Name = "mod01_OnTimeMacro"
Sub Reports(tsk As String)
    Dim Msg As VbMsgBoxResult
    
    Msg = MsgBox(tsk & " is about to run." & vbNewLine & _
        "Do you like to snooze?", vbQuestion + vbYesNo, tsk)
    
    If Msg = vbYes Then
        Application.OnTime Now + TimeSerial(0, 0, 10), "'Reports""" & tsk & """'"
    End If
End Sub

Sub ReadReports()
    Lr = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To Lr
        If Cells(i, 1).Value = Format(Date, "ddd") Then
            Application.OnTime Cells(i, 2).Value, "'Reports""" & Cells(i, 3).Value & """'"
        End If
    Next
End Sub



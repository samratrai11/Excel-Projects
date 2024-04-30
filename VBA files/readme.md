Hello Welcome 😃
In this VBA Files Folder, I have stored VBA related small practice projects.
I hope you all like it 🙌.

For Example This is the code to the VBA if,case and loops test file and you can dowload them to check it out:

Public Sub Ifprac()
If ActiveCell.Value >= 21 Then
MsgBox ("user is 21 or older")
Else
MsgBox ("Not allowed")
End If
End Sub

Public Sub Ifprac2()
If ActiveCell.Value >= 21 Then
MsgBox ("user is 21 or older")
ElseIf ActiveCell.Value > 90 Then
MsgBox (">90")
Else
MsgBox ("Not allowed")
End If
End Sub

Public Sub caseprac()
Select Case ActiveCell.Value
Case Is > 90
MsgBox (">90")
Case 21 To 90
MsgBox ("21 to 90")
Case Else
MsgBox ("Not allowed")
End Select
End Sub

Public Sub loopprac()
Dim i As Integer
i = 1
Do While i <= 9
Ifprac
ActiveCell.Offset(1, 0).Select
i = i + 1
Loop
End Sub

Public Sub loopprac2()
Do While ActiveCell.Value <> ""
caseprac
ActiveCell.Offset(1, 0).Select
Loop
End Sub

Public Sub loopprac3()
Dim user As Range
For Each user In Selection
caseprac
ActiveCell.Offset(1, 0).Select
Next user
End Sub

Public Sub loopprac4()
Dim i As Integer
For i = 1 To ActiveSheet.UsedRange.Rows.Count - 1
Ifprac
ActiveCell.Offset(1, 0).Select
Next i
End Sub

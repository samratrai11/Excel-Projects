Hello Welcome 😃
In this VBA Files Folder, I have stored VBA related small practice projects.
I hope you all like it 🙌.

•	For Example This is the code to the VBA if,case and loops test file and you can dowload them to check it out:

1.	Public Sub Ifprac()
2.	If ActiveCell.Value >= 21 Then
3.	MsgBox ("user is 21 or older")
4.	Else
5.	MsgBox ("Not allowed")
6.	End If  
7.	End Sub

1.	Public Sub Ifprac2()
2.	If ActiveCell.Value >= 21 Then
3.	MsgBox ("user is 21 or older")
4.	ElseIf ActiveCell.Value > 90 Then
5.	MsgBox (">90")
6.	Else
7.	MsgBox ("Not allowed")
8.	End If
9.	End Sub

1.	Public Sub caseprac()
2.	Select Case ActiveCell.Value
3.	Case Is > 90
4.	MsgBox (">90")
5.	Case 21 To 90
6.	MsgBox ("21 to 90")
7.	Case Else
8.	MsgBox ("Not allowed")
9.	End Select
10.	End Sub

1.	Public Sub loopprac()
2.	Dim i As Integer
3.	i = 1
4.	Do While i <= 9
5.	Ifprac
6.	ActiveCell.Offset(1, 0).Select
7.	i = i + 1
8.	Loop
9.	End Sub

1.	Public Sub loopprac2()
2.	Do While ActiveCell.Value <> ""
3.	caseprac
4.	ActiveCell.Offset(1, 0).Select
5.	Loop
6.	End Sub

1.	Public Sub loopprac3()
2.	Dim user As Range
3.	For Each user In Selection
4.	caseprac
5.	ActiveCell.Offset(1, 0).Select
6.	Next user
7.	End Sub

1.	Public Sub loopprac4()
2.	Dim i As Integer
3.	For i = 1 To ActiveSheet.UsedRange.Rows.Count - 1
4.	Ifprac
5.	ActiveCell.Offset(1, 0).Select
6.	Next i
7.  End Sub

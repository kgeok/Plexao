Sub Profile()
Dim Name As String
Dim Name1 As String
Dim Name2 As String
Name = InputBox("Name:")
Name1 = InputBox("Motto?")
Name2 = InputBox("Contact? (Email, Address, Nickname, etc.)")
Label1.Caption = (Name & "'s Plexao S")
Label2.Caption = (Name1)
Label3.Caption = (Name2)
On Error Resume Next
If (Name) = "" Then
Me.Label1.Caption = ("N/a")
End If
If (Name1) = "" Then
Me.Label2.Caption = ("N/a")
End If
If (Name2) = "" ThenSlide189 - 2
Me.Label2.Caption = ("N/a")
End If
End Sub
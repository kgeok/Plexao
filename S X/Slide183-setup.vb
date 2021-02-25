Private Sub TextBox1_Change()
WebBrowser1.Navigate (TextBox1.Text)
End Sub
Sub Setup()
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
If (Name2) = "" Then
Me.Label2.Caption = ("N/a")
End If
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(2).SlideIndex)
On Error Resume Next
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(2).SlideIndex)
End With
End With
End Sub

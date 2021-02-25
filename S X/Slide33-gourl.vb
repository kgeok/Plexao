Private Sub ComboBox1_DropButtonClick()
 If WebBrowser1.ReadyState = READYSTATE_COMPLETE Then
 With ComboBox1
 .AddItem WebBrowser1.LocationURL
 End With
 End If
 If ComboBox1.ListCount = 4 Then
 With ComboBox1
 .Clear
 End With
 End If
End Sub
Private Sub ComboBox1_Change()
On Error Resume Next
WebBrowser1.Navigate2 (ComboBox1.Text)
End Sub
Private Sub CommandButton2_Click()
InputBox "URL", _
 "Plexao S", WebBrowser1.LocationURL
 
End Sub

Sub GoURL()
On Error Resume Next
WebBrowser1.Navigate2 (ComboBox1.Text)
End Sub

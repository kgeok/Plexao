Private Sub TextBox1_Change()
WebBrowser1.Navigate (TextBox1.Text)
End Sub
Sub Chat()
WebBrowser1.Navigate ("http://plexao.pl.vu/s/Central/Apps/chat.html")
End Sub
Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
End Sub
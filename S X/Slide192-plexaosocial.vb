Private Sub TextBox1_Change()
WebBrowser1.Navigate (TextBox1.Text)
End Sub
Sub PlexaoSocial()
WebBrowser1.Navigate ("http://plexaosocial.wall.fm")
End Sub
Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
End Sub
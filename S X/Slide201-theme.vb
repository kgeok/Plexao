Sub Theme()
Dim Theme As String
Theme = ShowFileDialog("Themes", "*.pot; *.potx; *.potm; *.thmx")
 ActivePresentation.Slides.Range.ApplyTemplate _
 Filename:=(Theme)
 
End Sub

Sub ExportDoc()
Dim sPathUser As String
sPathUser = Environ$("USERPROFILE") & "\My Documents\"
Dim mytext As String
mytext = pdoc.Text
Open sPathUser & "PAPERDOCUMENT1.TXT" For Append As #1
Print #1, mytext
Close #1
MsgBox "Saved to 'My Documents'", vbInformation, "Paper"
End Sub

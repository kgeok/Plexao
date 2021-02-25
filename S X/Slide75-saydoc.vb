Sub SayDoc()
strText = (pdoc.Text)
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End Sub
Sub NewDoc()
pdoc.Text = ("")
End Sub
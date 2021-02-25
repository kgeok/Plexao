Sub SetupComp()
On Error Resume Next
Dim oSld As Slide
Dim oShp As Shape
For Each oSld In ActivePresentation.Slides
For Each oShp In oSld.Shapes
If oShp.HasTextFrame Then
If oShp.TextFrame.TextRange.Find("set") Is Nothing Then
Else
'found it
oSld.Delete
End If
End If
Next oShp
Next oSld
For Each oSld In ActivePresentation.Slides
For Each oShp In oSld.Shapes
If oShp.HasTextFrame Then
If oShp.TextFrame.TextRange.Find("Wel") Is Nothing Then
Else
'found it
oSld.Delete
End If
End If
Next oShp
Next oSld
For Each oSld In ActivePresentation.Slides
For Each oShp In oSld.Shapes
If oShp.HasTextFrame Then
If oShp.TextFrame.TextRange.Find("Setup") Is Nothing Then
'didn't find it
Else
'found it
oSld.Delete
Exit Sub
End If
End If
Next oShp
Next oSld
On Error Resume Next
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(2).SlideIndex)
For Each oSld In ActivePresentation.Slides
For Each oShp In oSld.Shapes
If oShp.HasTextFrame Then
If oShp.TextFrame.TextRange.Find("Welcome") Is Nothing Then
'didn't find it
Else
'found it
oSld.Delete
Exit Sub
End If
End If
Next oShp
Next oSldSlide189 - 3
End With
End Sub
Private Sub Label2_Click()
End Sub
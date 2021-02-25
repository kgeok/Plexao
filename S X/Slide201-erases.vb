Sub EraseS()
MsgBox ("All Data will be Erased")
On Error Resume Next
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(14).SlideIndex)
Dim i As Integer
For i = ActivePresentation.Slides.Count To 1 Step -1
If ActivePresentation.Slides(i) _
.SlideShowTransition.Hidden = False Then
ActivePresentation.Slides(i).Delete
End If
Next i
End With
End Sub
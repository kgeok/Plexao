Private Sub AcroPDF1_OnError()
End Sub
Private Sub TV1_GotFocus()
End Sub
Private Sub InkPicture1_Stroke(ByVal Cursor As MSINKAUTLib.IInkCursor, ByVal Stroke As MSINKAUTLib.IIn
kStrokeDisp, Cancel As Boolean)
End Sub
Sub InkErase()
InkPicture1.EditingMode = IOEM_Delete
End Sub

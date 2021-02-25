Sub ExportS()
 Dim dlgSaveAs As FileDialog, strFile As String
 Set dlgSaveAs = Application.FileDialog(msoFileDialogSaveAs)
 dlgSaveAs.Show
 On Error Resume Next
 strFile = dlgSaveAs.SelectedItems(1)
 If Err Then
 Exit Sub
 ActivePresentation.SaveAs strFile
 MsgBox "Saved in:" & vbNewLine & strFile
 End If
 End Sub

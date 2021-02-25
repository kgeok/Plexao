Function ShowFileDialog(filtername As String, filter As String) As String
Dim dlgOpen As FileDialog
Dim result As String
Set dlgOpen = Application.FileDialog(Type:=msoFileDialogFilePicker)
With dlgOpen
'Add a filter that includes GIF and JPEG images and make it the first item in the list.
.Filters.Add filtername, filter, 1
.AllowMultiSelect = False
If .Show = -1 Then
'Step through each string in the FileDialogSelectedItems collection.
' There will only be one but this works better than a file open dialog for some reason.
For Each vrtSelectedItem In .SelectedItems
'vrtSelectedItem is a String that contains the path of each selected item.
'You can use any file I/O functions that you want to work with this path.
result = vrtSelectedItem
Next vrtSelectedItem
'If the user presses Cancel...
Else
End If
End With
ShowFileDialog = result
End Function

Sub Theme()
Dim Theme As String
Theme = ShowFileDialog("Themes", "*.pot; *.potx; *.potm; *.thmx")
 ActivePresentation.Slides.Range.ApplyTemplate _
 Filename:=(Theme)
 
End Sub

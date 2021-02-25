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
Sub ChangeWallpaper()
'
' Macro written 14/01/2010 by Plexao
'
Dim file As String
'
' display a file selection dialog box
'
file = ShowFileDialog("Pictures", "*.gif; *.jpg; *.jpeg; *.png; *.bmp")
If ActivePresentation.HasTitleMaster Then
With ActivePresentation.TitleMaster.Background
.Fill.Visible = msoTrue
.Fill.ForeColor.RGB = RGB(255, 255, 255)
.Fill.BackColor.SchemeColor = ppAccent1
.Fill.Transparency = 0#
.Fill.UserPicture file
End With
End If
With ActivePresentation.SlideMaster.Background
.Fill.Visible = msoTrue
.Fill.ForeColor.RGB = RGB(255, 255, 255)
.Fill.BackColor.SchemeColor = ppAccent1
.Fill.Transparency = 0#
.Fill.UserPicture file
End With
With ActivePresentation.Slides.Range
.FollowMasterBackground = msoTrue
.DisplayMasterShapes = msoTrue
End With
End Sub
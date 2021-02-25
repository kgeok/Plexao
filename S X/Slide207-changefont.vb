Sub ChangeFont()
FONTUI = InputBox("Enter a Font Size:")
 Dim Sld As Slide
 Dim shp As Shape
 
 For Each Sld In ActivePresentation.Slides
 For Each shp In Sld.Shapes
 If shp.HasTextFrame Then
 shp.TextFrame.TextRange.font.Size = (FONTUI)
 End If
 Next shp
 Next Sld
 
End Sub
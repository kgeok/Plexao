Private Sub CommandButton1_Click()
strText = TextBox9.Text
If Me.TextBox1.Text = "reset -0000" Then
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Welcome to Caller, Type a Command, or click ( i )"
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End If
If Me.TextBox1.Text = "master control -0000" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(11).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Welcome to Caller, Type a Command, or click ( i )"
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "paper" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(7).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "settings" Then
With SlideShowWindows(13)
.View.GotoSlide (.Presentation.Slides(9).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "interweb" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(2).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "internet" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(2).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strTextSlide207 - 2
End With
End If
If Me.TextBox1.Text = "social" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(5).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "fb" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(5).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "facebook" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(5).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "post" Then
With SlideShowWindows(1)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "post..."
ActivePresentation.FollowHyperlink _
 Address:="http://www.facebook.com/share.php", _
 NewWindow:=True, AddHistory:=True
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "tweet" Then
With SlideShowWindows(1)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "tweeting..."
ActivePresentation.FollowHyperlink _
 Address:="http://twitter.com/share", _
 NewWindow:=True, AddHistory:=True
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "twitter" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(5).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "social" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(5).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "me" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(6).SlideIndex)
Me.TextBox1.Text = ""Slide207 - 3
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "cancel" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(1).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "home" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(1).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "chat" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(3).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "weather" Then
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
ActivePresentation.FollowHyperlink _
 Address:="http://weather.com", _
 NewWindow:=True, AddHistory:=True
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End If
If Me.TextBox1.Text = "stocks" Then
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
ActivePresentation.FollowHyperlink _
 Address:="http://finance.yahoo.com/", _
 NewWindow:=True, AddHistory:=True
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End If
If Me.TextBox1.Text = "youtube" Then
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
ActivePresentation.FollowHyperlink _
 Address:="http://youtube.com", _
 NewWindow:=True, AddHistory:=True
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End If
If Me.TextBox1.Text = "YouTube" Then
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
ActivePresentation.FollowHyperlink _
 Address:="http://youtube.com", _
 NewWindow:=True, AddHistory:=True
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End If
If Me.TextBox1.Text = "video chat" Then
With SlideShowWindows(1)Slide207 - 4
.View.GotoSlide (.Presentation.Slides(3).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "psocial" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(4).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "plexaosocial" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(4).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "write" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(7).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "type" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(7).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "draw" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(21).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "config" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(9).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "hello" Then
With SlideShowWindows(1)
strText = TextBox1.Text
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Hello!"
End With
End If
If Me.TextBox1.Text = "version" Then
With SlideShowWindows(1)
strText = TextBox1.TextSlide207 - 5
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
MsgBox "Plexao S, Made by Plexao, Created by Kevin George http://plexao.pl.vu", 65536, "Plexao S"
Me.TextBox1.Text = ""
Me.TextBox9.Text = "About"
End With
End If
If Me.TextBox1.Text = "about" Then
With SlideShowWindows(1)
strText = TextBox1.Text
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
MsgBox "Plexao S, Made by Plexao, Created by Kevin George http://plexao.pl.vu", 65536, "Plexao S"
Me.TextBox1.Text = ""
Me.TextBox9.Text = "About"
End With
End If
If Me.TextBox1.Text = "plexao s" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(14).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
If Me.TextBox1.Text = "Plexao S" Then
With SlideShowWindows(1)
.View.GotoSlide (.Presentation.Slides(14).SlideIndex)
Me.TextBox1.Text = ""
Me.TextBox9.Text = "Launching..."
Set ObjVoice = CreateObject("SAPI.SpVoice")
ObjVoice.Speak strText
End With
End If
End Sub
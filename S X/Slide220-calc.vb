Dim FirstNumber As Single
Dim SecondNumber As Single
Dim AnswerNumber As Single
Dim ArithmeticProcess As String
Private Sub cmd0_Click()
txtdisplay.Text = txtdisplay.Text & "0"
End Sub
Private Sub cmd1_Click()
txtdisplay.Text = txtdisplay.Text & "1"
End Sub
Private Sub cmd2_Click()
txtdisplay.Text = txtdisplay.Text & "2"
End Sub
Private Sub cmd3_Click()
txtdisplay.Text = txtdisplay.Text & "3"
End Sub
Private Sub cmd4_Click()
txtdisplay.Text = txtdisplay.Text & "4"
End Sub
Private Sub cmd5_Click()
txtdisplay.Text = txtdisplay.Text & "5"
End Sub
Private Sub cmd6_Click()
txtdisplay.Text = txtdisplay.Text & "6"
End Sub
Private Sub cmd7_Click()
txtdisplay.Text = txtdisplay.Text & "7"
End Sub
Private Sub cmd8_Click()
txtdisplay.Text = txtdisplay.Text & "8"
End Sub
Private Sub cmd9_Click()
txtdisplay.Text = txtdisplay.Text & "9"
End Sub
Private Sub cmdadd_Click()
FirstNumber = Val(txtdisplay.Text)
txtdisplay.Text = "0"
ArithmeticProcess = "+"
End Sub
Private Sub cmdclear_Click()
txtdisplay.Text = 0
End Sub
Private Sub cmddecimal_Click()
txtdisplay.Text = txtdisplay.Text & "."
End Sub
Private Sub cmddivide_Click()
FirstNumber = Val(txtdisplay.Text)
txtdisplay.Text = "0"
ArithmeticProcess = "/"
End Sub
Private Sub cmdequal_Click()
SecondNumber = Val(txtdisplay.Text)
If ArithmeticProcess = "+" Then
AnswerNumber = FirstNumber + SecondNumber
End IfSlide220 - 2
If ArithmeticProcess = "-" Then
AnswerNumber = FirstNumber - SecondNumber
End If
If ArithmeticProcess = "X" Then
AnswerNumber = FirstNumber * SecondNumber
End If
If ArithmeticProcess = "/" Then
AnswerNumber = FirstNumber / SecondNumber
End If
txtdisplay.Text = AnswerNumber
End Sub
Private Sub cmdmultiply_Click()
FirstNumber = Val(txtdisplay.Text)
txtdisplay.Text = "0"
ArithmeticProcess = "X"
End Sub
Private Sub cmdp_Click()
txtdisplay.Text = AnswerNumber
End Sub
Private Sub cmdsubtract_Click()
FirstNumber = Val(txtdisplay.Text)
txtdisplay.Text = "0"
ArithmeticProcess = "-"
End Sub
Private Sub txtdisplay_Change()
End Sub
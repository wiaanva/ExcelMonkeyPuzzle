VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1_MainForm 
   Caption         =   "SW Toets"
   ClientHeight    =   9312.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11628
   OleObjectBlob   =   "UserForm1_MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1_MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Long, j As Long
Dim TotalQuestions As Long
Dim TotalCorrect As Long
Dim rng As Range
'This form pulls information from "Configure Test" sheet of the Excel workbook and presents it as a Monkey Puzzle, recording the answers selected.

Private Sub UserForm_Initialize()
     
    TotalQuestions = Sheets("Configure Test").Range("B4")

    Frame1_Question.Caption = "Questions: (Total: " + CStr(TotalQuestions) + ")"
    Label3_TestTitle.Caption = Sheets("Configure Test").Range("B2")
    
    Set rng = Sheets("Configure Test").Range("B8")

    i = 0: j = 1
    Me.Label2_QuestionArea.ForeColor = IIf(Me.Label2_QuestionArea.ForeColor = vbRed, vbBlack, vbRed)
    
    Label2_QuestionArea.Caption = "Click here to start..."
    
End Sub


Private Sub Label2_QuestionArea_Click()

    CommandButton1_Prev.Enabled = True
    CommandButton2_Next.Enabled = True
    OptionButton1.Enabled = True
    OptionButton2.Enabled = True
    OptionButton3.Enabled = True
    OptionButton4.Enabled = True
   
    Frame1_Question.Caption = "Question: " + CStr(i + 1) + " of " + CStr(TotalQuestions)
     
    Label2_QuestionArea.Caption = rng.Offset(i).Value
    
    OptionButton1.Caption = rng.Offset(i, j).Value: j = j + 1
    OptionButton2.Caption = rng.Offset(i, j).Value: j = j + 1
    OptionButton3.Caption = rng.Offset(i, j).Value: j = j + 1
    OptionButton4.Caption = rng.Offset(i, j).Value: j = j + 1
    
    Me.Label2_QuestionArea.ForeColor = vbBlack
   Exit Sub

End Sub

Private Sub CommandButton5_Click()

    CommandButton5.ControlTipText = ActiveCell.Address

End Sub

Private Sub CommandButton1_Prev_Click()
    
    i = i - 1: j = 1

    If i < 0 Then
        MsgBox "1st Row Reached"
        Exit Sub
    End If

    Frame1_Question.Caption = "Question: " + CStr(i + 1) + " of " + CStr(TotalQuestions)
    
    Label2_QuestionArea.Caption = rng.Offset(i).Value
    OptionButton1.Caption = rng.Offset(i, j).Value: j = j + 1
    OptionButton2.Caption = rng.Offset(i, j).Value: j = j + 1
    OptionButton3.Caption = rng.Offset(i, j).Value: j = j + 1
    OptionButton4.Caption = rng.Offset(i, j).Value: j = j + 1
        
End Sub

Private Sub OptionButton1_Click()
     
    If OptionButton1.Value = True Then
         
        rng.Offset(i, j) = "A"
        
        ListBox1.AddItem "Question: " + CStr(i + 1) + " = " + rng.Offset(i, j)
        Me.CommandButton2_Next.BackColor = RGB(119, 158, 203)
     
     End If

End Sub

Private Sub OptionButton2_Click()
 
    If OptionButton2.Value = True Then

        rng.Offset(i, j) = "B"
        
        ListBox1.AddItem "Question: " + CStr(i + 1) + " = " + rng.Offset(i, j)
        Me.CommandButton2_Next.BackColor = RGB(119, 158, 203)
     
     End If
End Sub

Private Sub OptionButton3_Click()

    If OptionButton3.Value = True Then

        rng.Offset(i, j) = "C"
        
        ListBox1.AddItem "Question: " + CStr(i + 1) + " = " + rng.Offset(i, j)
        Me.CommandButton2_Next.BackColor = RGB(119, 158, 203)
           
     End If

End Sub

Private Sub OptionButton4_Click()

    If OptionButton4.Value = True Then

        rng.Offset(i, j) = "D"
        
        ListBox1.AddItem "Question: " + CStr(i + 1) + " = " + rng.Offset(i, j)
        Me.CommandButton2_Next.BackColor = RGB(119, 158, 203)
     
     End If
     
End Sub

Private Sub CommandButton2_Next_Click()
     
    i = i + 1: j = 1
    
    If i >= Sheets("Configure Test").Range("B4") Then
       MsgBox "End of questions"
        Exit Sub
    End If
    
    Frame1_Question.Caption = "Question: " + CStr(i + 1) + " of " + CStr(TotalQuestions)
    
    Label2_QuestionArea.Caption = rng.Offset(i).Value
    OptionButton1.Caption = rng.Offset(i, j).Value: j = j + 1
    OptionButton2.Caption = rng.Offset(i, j).Value: j = j + 1
    OptionButton3.Caption = rng.Offset(i, j).Value: j = j + 1
    OptionButton4.Caption = rng.Offset(i, j).Value: j = j + 1
    '
    '
    Me.CommandButton2_Next.BackColor = 14737632
    
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Label1_Click()

End Sub


Private Sub CommandButton4_reset_Click()
 
    OptionButton1.Value = False
    OptionButton2.Value = False
    OptionButton3.Value = False
    OptionButton4.Value = False
    ListBox1.Clear
    Me.CommandButton2_Next.BackColor = 14737632
    
    
End Sub


Private Sub CommandButton3_Close_Click()

    TotalCorrect = Sheets("Configure Test").Range("B5")

    MsgBox "Your score: " & TotalCorrect & "/" & TotalQuestions
    
    ActiveSheet.ListObjects("Table1").ListColumns("Student Answer").DataBodyRange.Clear
    
    Unload Me
    
End Sub

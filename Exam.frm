VERSION 5.00
Begin VB.Form Exam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Examination"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Icon            =   "Exam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnsDown 
      Caption         =   "Move down"
      Height          =   375
      Left            =   3840
      TabIndex        =   14
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnsUp 
      Caption         =   "Move up"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtEditAns 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton cmdUpdateAnswer 
      Height          =   255
      Left            =   6360
      Picture         =   "Exam.frx":00D2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3480
      Width           =   735
   End
   Begin VB.ListBox lstAns 
      Height          =   1620
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.TextBox txtTruAns 
      Height          =   1605
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Exam.frx":0146
      Top             =   1440
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.OptionButton optTrue 
      Caption         =   "True"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton optFalse 
      Caption         =   "False"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblQuestion 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   6615
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label lblOrder 
      Alignment       =   1  'Right Justify
      Caption         =   "Answers in order!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblItem 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Answer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
End
Attribute VB_Name = "Exam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' TODO
' DONE check if add/del ans is really needed // fill anslist with blanks at start
' DONE submit routine
' DONE answer checker routine

Option Explicit

Private Const TYPE_ENUM = 1
Private Const TYPE_MULTI = 2
Private Const TYPE_TF = 3
Private Const TYPE_ESSAY = 4
Private Const TYPE_IDENT = 5

Dim Q(100) As QItem
Dim C(100) As String
Dim N, AC, QI As Integer
Dim Detailed As Boolean
Dim Algo As Integer
Dim Tolerance As Integer
Dim Score As Integer
Dim TotalScore As Integer
Dim Unscored As Integer
Dim Force As Boolean

Private Sub cmdAnsDown_Click()
    If lstAns.ListIndex >= 0 Then
        Dim Temp As String
        Temp = C(lstAns.ListIndex + 2)
        C(lstAns.ListIndex + 2) = C(lstAns.ListIndex + 1)
        C(lstAns.ListIndex + 1) = Temp
        Call RefreshAnsList
        lstAns.ListIndex = lstAns.ListIndex + 1
    Else
        MsgBox "Please select an answer first.", vbOKOnly, "Error"
    End If
End Sub

Private Sub cmdAnsUp_Click()
    If lstAns.ListIndex >= 1 Then
        Dim Temp As String
        Temp = C(lstAns.ListIndex + 1)
        C(lstAns.ListIndex + 1) = C(lstAns.ListIndex)
        C(lstAns.ListIndex) = Temp
        Call RefreshAnsList
        lstAns.ListIndex = lstAns.ListIndex - 1
    Else
        MsgBox "Please select an answer first.", vbOKOnly, "Error"
    End If
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to cancel the exam?", vbYesNo, "Confirmation") = vbYes Then
        Force = True
        Do While QI <= N
            Call cmdSubmit_Click
        Loop
    End If
End Sub

Private Sub cmdHelp_Click()
    ExamHelp.Show
End Sub

Private Sub cmdSubmit_Click()
    Dim I, j, P As Integer
    P = 0
    If Q(QI).QType = TYPE_ESSAY Then
        A ("<table border='1' style='border: 1px solid black; border-collapse: collapse; width:100%; background-color: #ffCCCC'>")
    Else
        A ("<table border='1' style='border: 1px solid black; border-collapse: collapse; width:100%'>")
    End If
    A ("<tr><td colspan='2'><b>" & lblItem.Caption & Q(QI).QuestionText & "</b></td></tr>")
    If Q(QI).QType = TYPE_ENUM Then
        Dim Pass As Boolean
        Pass = True
        For I = 1 To AC
            If C(I) = "New answer" Then Pass = False
        Next I
        If Not Pass And Not Force Then
            If MsgBox("You still haven't filled in some answers. Are you sure you want to continue?", vbYesNo, "Warning") = vbNo Then
                Exit Sub
            End If
        End If
        For I = 1 To Q(QI).AnswerCount
            A ("<tr><td>" & Q(QI).GetAnswer(I) & "</td>")
            A ("<td>" & C(I) & "</td></tr>")
        Next I
        Dim Points As Boolean
        If Not Q(QI).InOrder Then
            For I = 1 To Q(QI).AnswerCount
                Points = False
                For j = 1 To AC
                    If Correct(Q(QI).GetAnswer(I), C(j)) Then Points = True
                Next j
                If Points Then P = P + 1
            Next I
        Else
            For I = 1 To Q(QI).AnswerCount
                If Correct(Q(QI).GetAnswer(I), C(I)) Then P = P + 1
            Next I
        End If
        A ("<tr><td colspan='2'>Score: " & Str$(P) & "</td></tr>")
        Score = Score + P
        TotalScore = TotalScore + Q(QI).AnswerCount
    ElseIf Q(QI).QType = TYPE_MULTI Then
        If txtTruAns.Text = "" And Not Force Then
            MsgBox "Please select an answer first.", vbOKOnly, "Error"
            Exit Sub
        ElseIf Force Then
            txtTruAns.Text = "No answer"
        End If
        A ("<tr><td>" & Q(QI).CorrectAnswer & "</td>")
        A ("<td>" & txtTruAns.Text & "</td></tr>")
        If Correct(Q(QI).CorrectAnswer, txtTruAns.Text) Then
            Score = Score + 1
            A ("<tr><td colspan='2'>Correct</td></tr>")
        Else
            A ("<tr><td colspan='2'>Wrong</td></tr>")
        End If
        TotalScore = TotalScore + 1
    ElseIf Q(QI).QType = TYPE_TF Then
        If txtTruAns.Text = "" And Not Force Then
            MsgBox "Please select an answer first.", vbOKOnly, "Error"
            Exit Sub
        ElseIf Force Then
            txtTruAns.Text = "No answer"
        End If
        A ("<tr><td>" & Q(QI).CorrectAnswer & "</td>")
        A ("<td>" & txtTruAns.Text & "</td></tr>")
        If Correct(Q(QI).CorrectAnswer, txtTruAns.Text) Then
            Score = Score + 1
            A ("<tr><td colspan='2'>Correct</td></tr>")
        Else
            A ("<tr><td colspan='2'>Wrong</td></tr>")
        End If
        TotalScore = TotalScore + 1
    ElseIf Q(QI).QType = TYPE_ESSAY Then
        If txtTruAns.Text = "" And Not Force Then
            MsgBox "Please type an answer first.", vbOKOnly, "Error"
            Exit Sub
        ElseIf Force Then
            txtTruAns.Text = "No answer"
        End If
        A ("<tr><td colspan='2'>" & txtTruAns.Text & "</td></tr>")
        Unscored = Unscored + 1
    ElseIf Q(QI).QType = TYPE_IDENT Then
        If txtTruAns.Text = "" And Not Force Then
            MsgBox "Please type an answer first.", vbOKOnly, "Error"
            Exit Sub
        ElseIf Force Then
            txtTruAns.Text = "No answer"
        End If
        A ("<tr><td>" & Q(QI).CorrectAnswer & "</td>")
        A ("<td>" & txtTruAns.Text & "</td></tr>")
        If Correct(Q(QI).CorrectAnswer, txtTruAns.Text) Then
            Score = Score + 1
            A ("<tr><td colspan='2'>Correct</td></tr>")
        Else
            A ("<tr><td colspan='2'>Wrong</td></tr>")
        End If
        TotalScore = TotalScore + 1
    End If
    A ("</table>")
    QI = QI + 1
    If QI > N Then
        If Unscored > 0 Then
            If Unscored = 1 Then
                A ("<p><font color='#FF0000'><b>There is still a question that was not scored. Please check this first and add points to the total score as appropriate.</b></font></p>")
            Else
                A ("<p><font color='#FF0000'><b>There are still " & Str$(Unscored) & " questions that were not scored. Please check this first and add points to the total score as appropriate.</b></font></p>")
            End If
        End If
        A ("<table border='1' style='border: 1px solid black; border-collapse: collapse;'>")
        A ("<tr><td><b>Final score:</b></td><td>" & Str$(Score) & " / " & Str$(TotalScore) & "</td></tr></table>")
        Call EndHTML
        Close #2
        ExamFinished.Show
        Unload Me
    Else
        Call ShowQuestion
    End If
End Sub

Private Function Correct(ByVal A As String, ByVal B As String) As Boolean
    Dim R As Boolean
    R = False
    If Not Q(QI).CaseSensitive Then
        A = UCase(A)
        B = UCase(B)
    End If
    If Q(QI).StrictSpelling Then
        Correct = (A = B)
    ElseIf Algo = 1 Then
        Correct = BasicDistance(A, B) <= Tolerance
    ElseIf Algo = 2 Then
        Correct = OSADistance(A, B) <= Tolerance
    ElseIf Algo = 3 Then
        Correct = ATDistance(A, B) <= Tolerance
    Else
        MsgBox "Hey, the string matching algorithm chosen index is wrong!"
        End
    End If
End Function

Private Sub cmdUpdateAnswer_Click()
    If txtEditAns.Text = "" Then
        MsgBox "Please write something that will replace the current answer.", vbOKOnly, "Error"
        Exit Sub
    End If
    If lstAns.ListIndex >= 0 Then
        C(lstAns.ListIndex + 1) = txtEditAns.Text
        Call RefreshAnsList
        If lstAns.ListIndex < AC - 1 Then
            lstAns.ListIndex = lstAns.ListIndex + 1
            If txtEditAns.Text = "New answer" Then txtEditAns.Text = ""
            txtEditAns.SetFocus
            txtEditAns.SelStart = Len(txtEditAns.Text)
        End If
    Else
        MsgBox "Please choose an answer to edit first.", vbOKOnly, "Error"
    End If
End Sub

Private Sub lstAns_Click()
    txtTruAns.Text = C(lstAns.ListIndex + 1)
    txtEditAns.Text = C(lstAns.ListIndex + 1)
    If lblType.Caption = "Enumeration" Then
        If lstAns.ListIndex < 0 Then
            cmdAnsUp.Enabled = False
            cmdAnsDown.Enabled = False
        ElseIf lstAns.ListIndex = 0 Then
            cmdAnsUp.Enabled = False
            cmdAnsDown.Enabled = True
        ElseIf lstAns.ListIndex = AC - 1 Then
            cmdAnsUp.Enabled = True
            cmdAnsDown.Enabled = False
        Else
           cmdAnsUp.Enabled = True
           cmdAnsDown.Enabled = True
        End If
    Else
        cmdAnsUp.Enabled = False
        cmdAnsDown.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Force = False
    Randomize Timer
    Dim SetPath, Algorithm As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    Open SetPath For Input As #1
    Dim T, Folder As String
    Line Input #1, T
    Line Input #1, T
    Folder = Mid$(T, 1, Len(T))
    Line Input #1, T
    If Decode(T) = "True" Then Detailed = True Else Detailed = False
    Line Input #1, T
    Algorithm = Decode(T)
    If Algorithm = "Basic" Then
        Algo = 1
    ElseIf Algorithm = "Optimal string alignment" Then
        Algo = 2
    Else
        Algo = 3
    End If
    Line Input #1, T
    Tolerance = Int(Decode(T))
    Close #1
    Open Folder + "exam_questionlist.txt" For Input As #1
    Open Folder + "exam_studentdata.txt" For Input As #3
    Dim LN, FN As String
    Line Input #3, LN
    LN = Decode(LN)
    Line Input #3, FN
    FN = Decode(FN)
    Close #3
    Open Folder & Replace$(LN, " ", "") & Replace$(FN, " ", "") & "_encryptedresults.txt" For Output As #2
    Line Input #1, T
    T = Decode(Mid$(T, 1, Len(T)))
    Dim Temp As String
    Temp = Mid$(T, 1, Len(T))
    N = Val(Temp)
    Dim I As Integer
    For I = 1 To N
        Set Q(I) = New QItem
        Q(I).ReadQ (1)
    Next I
    Close #1
    Call PrepareHTML
    Call WriteExamineeDetails
    QI = 1
    Call ShowQuestion
    
    Score = 0
    TotalScore = 0
    Unscored = 0
End Sub

Private Sub ShowQuestion()
    Dim I, R As Integer
    
    lblOrder.Visible = False
    txtTruAns.Visible = False
    txtTruAns.Text = ""
    lstAns.Visible = False
    lstAns.Clear
    optTrue.Visible = False
    optTrue.Value = False
    optFalse.Visible = False
    optFalse.Value = False
    txtEditAns.Visible = False
    txtEditAns.Text = ""
    cmdAnsUp.Enabled = False
    cmdAnsDown.Enabled = False
    cmdUpdateAnswer.Visible = False
    AC = 0
    
    lblItem.Caption = "Question no. " & Str$(QI) & "/" & Str$(N) & ": "
    lblQuestion.Caption = Q(QI).QuestionText
    
    If Q(QI).QType = TYPE_ENUM Then
        If Q(QI).InOrder Then
            lblOrder.Visible = True
        End If
        lstAns.Visible = True
        txtEditAns.Visible = True
        cmdUpdateAnswer.Visible = True
        lblType.Caption = "Enumeration"
        AC = Q(QI).AnswerCount
        For I = 1 To AC
            C(I) = "New answer"
        Next I
        Call RefreshAnsList
    ElseIf Q(QI).QType = TYPE_MULTI Then
        lstAns.Visible = True
        lblType.Caption = "Multiple choice"
        Dim Choices(10) As String
        Dim Chosen(10) As Boolean
        Choices(1) = Q(QI).CorrectAnswer
        Chosen(1) = False
        For I = 2 To Q(QI).AnswerCount + 1
            Choices(I) = Q(QI).GetAnswer(I - 1)
            Chosen(I) = False
        Next I
        Do While AC < Q(QI).AnswerCount + 1
            R = Int(((Q(QI).AnswerCount + 1)) * Rnd + 1)
            If Not Chosen(R) Then
                AC = AC + 1
                lstAns.AddItem (Choices(R))
                C(AC) = Choices(R)
                Chosen(R) = True
            End If
        Loop
        Call RefreshAnsList
    ElseIf Q(QI).QType = TYPE_TF Then
        optTrue.Visible = True
        optFalse.Visible = True
        lblType.Caption = "True or false"
        lstAns.AddItem ("True")
        lstAns.AddItem ("False")
    ElseIf Q(QI).QType = TYPE_ESSAY Then
        txtTruAns.Visible = True
        lblType.Caption = "Essay"
    ElseIf Q(QI).QType = TYPE_IDENT Then
        txtTruAns.Visible = True
        lblType.Caption = "Identification"
    Else
        MsgBox "An unexpected error was encountered. Please ask the programmers.", vbOKOnly, "ERROR"
        End
    End If
End Sub

Private Sub A(S As String)
    Print #2, Encode(S)
End Sub

Private Sub PrepareHTML()
    Print #2, Encode("<html><title>Exam results</title><body>")
End Sub

Private Sub EndHTML()
    Print #2, Encode("</body></html>")
End Sub

Private Sub RefreshAnsList()
    Dim LPos As Integer
    LPos = lstAns.ListIndex
    lstAns.Clear
    Dim I As Integer
    For I = 1 To AC
        lstAns.AddItem (C(I))
    Next I
    If lstAns.ListCount <= LPos Then
        lstAns.ListIndex = lstAns.ListCount - 1
    Else
        lstAns.ListIndex = LPos
    End If
End Sub

Private Sub WriteExamineeDetails()
    Dim T, U, V As String
    Dim SetPath As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    Open SetPath For Input As #3
    Dim Folder As String
    Line Input #3, Folder
    Line Input #3, Folder
    Close #3
    Open Folder + "exam_studentdata.txt" For Input As #3
    A ("<h1 align='center'>FCAT Examination</h1><h6 align='center'>Version 8 | (c) 2015, BSCS Batch 2016</h7>")
    A ("<p align='center'>Exam taken and this report generated on <b>" & _
        Format$(Now, "mmmm d, yyyy hh:mm AM/PM") & "</b></p>")
    
    A ("<h2>Examinee's data</h2>")
    A ("<table border='1' style='border: 1px solid black; border-collapse: collapse; width:100%'>")
    Line Input #3, T
    T = Decode(Mid$(T, 1, Len(T)))
    A ("<tr><td>Last name:</td><td>" & T & "</td></tr>")
    Line Input #3, T
    T = Decode(Mid$(T, 1, Len(T)))
    A ("<tr><td>First name:</td><td>" & T & "</td></tr>")
    Line Input #3, T
    T = Decode(Mid$(T, 1, Len(T)))
    A ("<tr><td>Middle initial:</td><td>" & T & "</td></tr>")
    If Detailed = True Then
        Line Input #3, T
        T = Decode(Mid$(T, 1, Len(T)))
        A ("<tr><td>Sex:</td><td>" & T & "</td></tr>")
        Line Input #3, T
        T = Decode(Mid$(T, 1, Len(T)))
        Line Input #3, U
        U = Decode(Mid$(U, 1, Len(U)))
        Line Input #3, V
        V = Decode(Mid$(V, 1, Len(V)))
        A ("<tr><td>Date of birth:</td><td>" & T & " " & U & ", " & V & "</td></tr>")
        Line Input #3, T
        T = Decode(Mid$(T, 1, Len(T)))
        A ("<tr><td>Place of birth:</td><td>" & T & "</td></tr>")
        Line Input #3, T
        T = Decode(Mid$(T, 1, Len(T)))
        A ("<tr><td>Nationality:</td><td>" & T & "</td></tr>")
        Line Input #3, T
        T = Decode(Mid$(T, 1, Len(T)))
        A ("<tr><td>Civil status:</td><td>" & T & "</td></tr>")
        Line Input #3, T
        T = Decode(Mid$(T, 1, Len(T)))
        A ("<tr><td>Address:</td><td>" & T & "</td></tr>")
        Line Input #3, T
        T = Decode(Mid$(T, 1, Len(T)))
        A ("<tr><td>Contact no.:</td><td>" & T & "</td></tr>")
    End If
    A ("</table>")
    A ("<h2>Exam answers</h2>")
    Close #3
End Sub

Private Sub optFalse_Click()
    txtTruAns.Text = "False"
End Sub

Private Sub optTrue_Click()
    txtTruAns.Text = "True"
End Sub

Private Sub txtEditAns_GotFocus()
    If txtEditAns.Text = "New answer" Then txtEditAns.Text = ""
End Sub

Private Sub txtEditAns_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdUpdateAnswer_Click
End Sub

Private Sub txtEditAns_LostFocus()
    If txtEditAns.Text = "" Then txtEditAns.Text = "New answer"
End Sub

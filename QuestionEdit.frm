VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form QuestionEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question editor"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   ControlBox      =   0   'False
   Icon            =   "QuestionEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   735
      Left            =   6120
      TabIndex        =   27
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export encrypted"
      Height          =   375
      Left            =   4560
      TabIndex        =   26
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdAnsDown 
      Caption         =   "Move down"
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnsUp 
      Caption         =   "Move up"
      Height          =   255
      Left            =   3120
      TabIndex        =   24
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdNextItem 
      Height          =   195
      Left            =   3840
      Picture         =   "QuestionEdit.frx":00D2
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdPrevItem 
      DisabledPicture =   "QuestionEdit.frx":0146
      Height          =   195
      Left            =   3840
      Picture         =   "QuestionEdit.frx":0218
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdateAnswer 
      Height          =   255
      Left            =   6360
      Picture         =   "QuestionEdit.frx":028C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3960
      Width           =   375
   End
   Begin VB.CheckBox chkOrder 
      Caption         =   "In order"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2520
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton optFalse 
      Caption         =   "False"
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton optTrue 
      Caption         =   "True"
      Height          =   255
      Left            =   1080
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox chkStrict 
      Caption         =   "Strict spelling"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   17
      Top             =   1800
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox txtTruAns 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.ListBox lstAns 
      Height          =   1425
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   6615
   End
   Begin VB.CommandButton cmdDelAns 
      Caption         =   "Delete answer"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddAns 
      Caption         =   "Add answer"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtEditAns 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CheckBox chkSensitive 
      Caption         =   "Case sensitive"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox cboType 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "QuestionEdit.frx":0300
      Left            =   1080
      List            =   "QuestionEdit.frx":0313
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   480
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   3960
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdEditList 
      Caption         =   "Edit question list"
      Height          =   495
      Left            =   1560
      TabIndex        =   8
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   495
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox txtQuestion 
      Enabled         =   0   'False
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   6615
   End
   Begin VB.CommandButton cmdExportU 
      Caption         =   "Export unencrypted"
      Height          =   375
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox cboItem 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Answer:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Type:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Question no."
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "QuestionEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TYPE_ENUM = 1
Private Const TYPE_MULTI = 2
Private Const TYPE_TF = 3
Private Const TYPE_ESSAY = 4
Private Const TYPE_IDENT = 5

' Q = Question array
' A = Answer array
' M = Array for storage of multiple choices
' F = 'Flags' (currently, holding question type)
' C = 'Count' Number of multiple choices for that item (if multiple choice)
' O = "Offset" of the multiple choices

Dim QArray(100) As QItem
Dim QCount As Integer
Dim Updated As Boolean

Public Sub Clear()
    QCount = 0
End Sub

Public Sub AddQ(Q As QItem)
    QCount = QCount + 1
    Set QArray(QCount) = Q
End Sub

Public Sub RefreshList()
    cboItem.Clear
    Dim I As Integer
    For I = 1 To QCount
        cboItem.AddItem (Str$(I))
    Next I
    If QCount > 0 Then
        cboItem.ListIndex = 0
        Call RefreshQ
        cboItem.Enabled = True
        cboType.Enabled = True
        txtQuestion.Enabled = True
        cmdExport.Enabled = True
        chkOrder.Enabled = True
        chkStrict.Enabled = True
        chkSensitive.Enabled = True
    Else
        lstAns.Clear
        txtQuestion.Text = ""
        txtTruAns.Text = ""
        cboItem.Enabled = False
        cboType.Enabled = False
        txtQuestion.Enabled = False
        cmdExport.Enabled = False
        chkOrder.Enabled = False
        chkStrict.Enabled = False
        chkSensitive.Enabled = False
    End If
    
End Sub
Private Sub cboType_Click()
    If cboItem.ListIndex >= 0 Then
        If cboType.Text = "Enumeration" Then
            lstAns.Visible = True
            txtTruAns.Visible = False
            txtEditAns.Visible = True
            optTrue.Visible = False
            optFalse.Visible = False
            cmdAddAns.Enabled = True
            chkOrder.Visible = True
            chkStrict.Visible = True
            chkSensitive.Visible = True
            cmdUpdateAnswer.Visible = True
            Call RefreshQ
            QArray(cboItem.ListIndex + 1).QType = TYPE_ENUM
        ElseIf cboType.Text = "Multiple choice" Then
            lstAns.Visible = True
            txtTruAns.Visible = True
            txtEditAns.Visible = True
            optTrue.Visible = False
            optFalse.Visible = False
            cmdAddAns.Enabled = True
            chkOrder.Visible = False
            chkStrict.Visible = False
            chkSensitive.Visible = False
            cmdUpdateAnswer.Visible = True
            Call RefreshQ
            QArray(cboItem.ListIndex + 1).QType = TYPE_MULTI
        ElseIf cboType.Text = "True or False" Then
            lstAns.Visible = False
            txtTruAns.Visible = False
            txtEditAns.Visible = False
            optTrue.Visible = True
            optFalse.Visible = True
            cmdAddAns.Enabled = False
            cmdDelAns.Enabled = False
            chkOrder.Visible = False
            chkStrict.Visible = False
            chkSensitive.Visible = False
            cmdUpdateAnswer.Visible = False
            cmdAnsUp.Enabled = False
            cmdAnsDown.Enabled = False
            QArray(cboItem.ListIndex + 1).QType = TYPE_TF
        ElseIf cboType.Text = "Essay" Then
            lstAns.Visible = False
            txtTruAns.Visible = False
            txtEditAns.Visible = False
            optTrue.Visible = False
            optFalse.Visible = False
            cmdAddAns.Enabled = False
            cmdDelAns.Enabled = False
            chkOrder.Visible = False
            chkStrict.Visible = False
            chkSensitive.Visible = False
            cmdUpdateAnswer.Visible = False
            cmdAnsUp.Enabled = False
            cmdAnsDown.Enabled = False
            QArray(cboItem.ListIndex + 1).QType = TYPE_ESSAY
        ElseIf cboType.Text = "Identification" Then
            lstAns.Visible = False
            txtTruAns.Visible = True
            txtEditAns.Visible = False
            optTrue.Visible = False
            optFalse.Visible = False
            cmdAddAns.Enabled = False
            cmdDelAns.Enabled = False
            chkOrder.Visible = False
            chkStrict.Visible = True
            chkSensitive.Visible = True
            cmdUpdateAnswer.Visible = False
            cmdAnsUp.Enabled = False
            cmdAnsDown.Enabled = False
            QArray(cboItem.ListIndex + 1).QType = TYPE_IDENT
        Else
            MsgBox "An unexpected error was encountered. Please ask the programmers.", vbOKOnly, "ERROR"
            End
        End If
    End If
    Call RefreshQ
    Updated = False
End Sub

Private Sub cboItem_Click()
    Call RefreshQ
End Sub

Private Sub RefreshQ()
    Dim I As Integer
    Dim LPos As Integer
    LPos = lstAns.ListIndex
    If QArray(cboItem.ListIndex + 1).QType = TYPE_IDENT Then
        txtQuestion.Text = QArray(cboItem.ListIndex + 1).QuestionText
        txtTruAns.Text = QArray(cboItem.ListIndex + 1).CorrectAnswer
    ElseIf QArray(cboItem.ListIndex + 1).QType = TYPE_ESSAY Then
        txtQuestion.Text = QArray(cboItem.ListIndex + 1).QuestionText
    ElseIf QArray(cboItem.ListIndex + 1).QType = TYPE_TF Then
        txtQuestion.Text = QArray(cboItem.ListIndex + 1).QuestionText
        If QArray(cboItem.ListIndex + 1).CorrectAnswer = "True" Then
            optTrue.Value = True
            optFalse.Value = False
        ElseIf QArray(cboItem.ListIndex + 1).CorrectAnswer = "False" Then
            optTrue.Value = False
            optFalse.Value = True
        Else
            optTrue.Value = False
            optFalse.Value = False
        End If
    ElseIf QArray(cboItem.ListIndex + 1).QType = TYPE_MULTI Then
        txtQuestion.Text = QArray(cboItem.ListIndex + 1).QuestionText
        txtTruAns.Text = QArray(cboItem.ListIndex + 1).CorrectAnswer
        lstAns.Clear
        cmdDelAns.Enabled = False
        txtEditAns.Enabled = False
        For I = 1 To QArray(cboItem.ListIndex + 1).AnswerCount
            lstAns.AddItem (QArray(cboItem.ListIndex + 1).GetAnswer(I))
        Next I
        If QArray(cboItem.ListIndex + 1).AnswerCount > 0 Then
            cmdDelAns.Enabled = True
            txtEditAns.Enabled = True
        End If
    ElseIf QArray(cboItem.ListIndex + 1).QType = TYPE_ENUM Then
        txtQuestion.Text = QArray(cboItem.ListIndex + 1).QuestionText
        lstAns.Clear
        cmdDelAns.Enabled = False
        txtEditAns.Enabled = False
        For I = 1 To QArray(cboItem.ListIndex + 1).AnswerCount
            lstAns.AddItem (QArray(cboItem.ListIndex + 1).GetAnswer(I))
        Next I
        If QArray(cboItem.ListIndex + 1).AnswerCount > 0 Then
            cmdDelAns.Enabled = True
            txtEditAns.Enabled = True
        End If
    Else
        MsgBox "Invalid QType: " & Str$(QArray(cboItem.ListIndex + 1).QType)
    End If
    cboType.ListIndex = QArray(cboItem.ListIndex + 1).QType - 1
    If QArray(cboItem.ListIndex + 1).CaseSensitive Then
        chkSensitive.Value = vbChecked
    Else
        chkSensitive.Value = vbUnchecked
    End If
    If QArray(cboItem.ListIndex + 1).StrictSpelling Then
        chkStrict.Value = vbChecked
    Else
        chkStrict.Value = vbUnchecked
    End If
    If QArray(cboItem.ListIndex + 1).InOrder Then
        chkOrder.Value = vbChecked
    Else
        chkOrder.Value = vbUnchecked
    End If
    If lstAns.ListCount <= LPos Then
        lstAns.ListIndex = lstAns.ListCount - 1
    Else
        lstAns.ListIndex = LPos
    End If
    txtEditAns.Text = ""
End Sub

Private Sub chkOrder_Click()
    If cboItem.ListIndex >= 0 Then
        If chkOrder.Value = vbChecked Then
            QArray(cboItem.ListIndex + 1).InOrder = True
        Else
            QArray(cboItem.ListIndex + 1).InOrder = False
        End If
    End If
    Updated = False
End Sub

Private Sub chkSensitive_Click()
    If cboItem.ListIndex >= 0 Then
        If chkSensitive.Value = vbChecked Then
            QArray(cboItem.ListIndex + 1).CaseSensitive = True
        Else
            QArray(cboItem.ListIndex + 1).CaseSensitive = False
        End If
    End If
    Updated = False
End Sub

Private Sub chkStrict_Click()
    If cboItem.ListIndex >= 0 Then
        If chkStrict.Value = vbChecked Then
            QArray(cboItem.ListIndex + 1).StrictSpelling = True
        Else
            QArray(cboItem.ListIndex + 1).StrictSpelling = False
        End If
    End If
    Updated = False
End Sub

Private Sub cmdAddAns_Click()
    If txtEditAns.Text <> "" Then
        QArray(cboItem.ListIndex + 1).AddAnswer (txtEditAns.Text)
        txtEditAns.Text = ""
    Else
        QArray(cboItem.ListIndex + 1).AddAnswer ("New answer")
    End If
    Call RefreshQ
    txtEditAns.SetFocus
    Updated = False
End Sub

Private Sub cmdAnsDown_Click()
    If lstAns.ListIndex >= 0 Then
        Dim Temp As String
        Temp = QArray(cboItem.ListIndex + 1).GetAnswer(lstAns.ListIndex + 2)
        Call QArray(cboItem.ListIndex + 1).SetAnswer(lstAns.ListIndex + 2, QArray(cboItem.ListIndex + 1).GetAnswer(lstAns.ListIndex + 1))
        Call QArray(cboItem.ListIndex + 1).SetAnswer(lstAns.ListIndex + 1, Temp)
        Call RefreshQ
        lstAns.ListIndex = lstAns.ListIndex + 1
        Updated = False
    Else
        MsgBox "Please select an answer first.", vbOKOnly, "Error"
    End If
End Sub

Private Sub cmdAnsUp_Click()
    If lstAns.ListIndex >= 1 Then
        Dim Temp As String
        Temp = QArray(cboItem.ListIndex + 1).GetAnswer(lstAns.ListIndex + 1)
        Call QArray(cboItem.ListIndex + 1).SetAnswer(lstAns.ListIndex + 1, QArray(cboItem.ListIndex + 1).GetAnswer(lstAns.ListIndex))
        Call QArray(cboItem.ListIndex + 1).SetAnswer(lstAns.ListIndex, Temp)
        Call RefreshQ
        lstAns.ListIndex = lstAns.ListIndex - 1
        Updated = False
    Else
        MsgBox "Please select an answer first.", vbOKOnly, "Error"
    End If
End Sub

Private Sub cmdCancel_Click()
    If Not Updated Then
        If MsgBox("You still haven't saved your changes by clicking [Update]." & Chr(10) & "Are you sure you want to discard changes?", vbYesNo, "Confirmation") = vbNo Then
            Exit Sub
        End If
    End If
    Entry.Show
    Unload Me
End Sub

Private Sub cmdDelAns_Click()
    If lstAns.ListIndex < 0 Then
        MsgBox "Please select an answer to delete first.", vbOKOnly, "Error"
    Else
        QArray(cboItem.ListIndex + 1).DeleteAnswer (lstAns.ListIndex + 1)
        Call RefreshQ
        Updated = False
    End If
End Sub

Private Sub cmdEditList_Click()
    QListEdit.Show
    QuestionEdit.Hide
    Dim I As Integer
    QListEdit.Clear
    For I = 1 To QCount
        Call QListEdit.AddQ(QArray(I))
    Next I
    Call QListEdit.RefreshList
    Updated = False
End Sub

Private Sub cmdExportU_Click()
    CommonDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    CommonDialog.DefaultExt = "txt"
    CommonDialog.DialogTitle = "Save file"
    CommonDialog.ShowSave

    Dim Path As String
    Dim I As Integer
    Path = CommonDialog.FileName
    If Path <> "" Then
        Open Path For Output As #1
        Print #1, "This file is not encrypted."
        Print #1, Str$(QCount)
        For I = 1 To QCount
            QArray(I).WriteQU (1)
        Next I
        Close #1
        MsgBox "Question list exported successfully.", vbOKOnly, "Export success"
    End If
End Sub

Private Sub cmdExport_Click()
    CommonDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    CommonDialog.DefaultExt = "txt"
    CommonDialog.DialogTitle = "Save file"
    CommonDialog.ShowSave

    Dim Path As String
    Dim I As Integer
    Path = CommonDialog.FileName
    If Path <> "" Then
        Open Path For Output As #1
        Print #1, "This file is encrypted."
        Print #1, Encode(Str$(QCount))
        For I = 1 To QCount
            QArray(QCount).WriteQ (1)
        Next I
        Close #1
        MsgBox "Question list exported successfully.", vbOKOnly, "Export success"
    End If
End Sub

Private Sub cmdImport_Click()
    CommonDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    CommonDialog.DefaultExt = "txt"
    CommonDialog.DialogTitle = "Open file"
    CommonDialog.ShowOpen

    Dim Path, T As String
    Dim I As Integer
    Path = CommonDialog.FileName
    If Path <> "" Then
        Open Path For Input As #1
        Line Input #1, T
        If T = "This file is encrypted." Then
            Line Input #1, T
            T = Decode(T)
            QCount = Val(T)
            For I = 1 To QCount
                Set QArray(I) = New QItem
                QArray(I).ReadQ (1)
            Next I
        ElseIf T = "This file is not encrypted." Then
            Line Input #1, T
            QCount = Val(T)
            For I = 1 To QCount
                Set QArray(I) = New QItem
                QArray(I).ReadQU (1)
            Next I
        Else
            MsgBox "The file you are trying to import is not recognized as a qestion file.", vbOKOnly, "Error"
            Close #1
            Exit Sub
        End If
        Close #1
        Call RefreshList
        MsgBox "Questions successfuly imported. Please click [Update] in order to save the changes.", vbOKOnly, "Import success"
    End If
End Sub

Private Sub cmdHelp_Click()
    QuestionEditHelp.Show
End Sub

Private Sub cmdNextItem_Click()
    If Val(cboItem.Text) < QCount Then
        cboItem.Text = Str$(Val(cboItem.Text) + 1)
    End If
End Sub

Private Sub cmdPrevItem_Click()
    If Val(cboItem.Text) > 1 Then
        cboItem.Text = Str$(Val(cboItem.Text) - 1)
    End If
End Sub

Private Sub cmdUpdate_Click()
    If QCount = 0 Then
        If MsgBox("Are you sure you want to replace the current question list with an empty one?", vbYesNo, "Replace question list") = vbNo Then Exit Sub
    End If
    Dim SetPath As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    Open SetPath For Input As #1
    Dim Folder As String
    Line Input #1, Folder
    Line Input #1, Folder
    Close #1
    Dim B As Boolean
    If FolderExists(Folder) = False Then MkDir Folder
    Open Folder + "exam_questionlist.txt" For Output As #1
    Dim I As Integer
    Dim T As String
    Print #1, Encode(Str$(QCount))
    For I = 1 To QCount
        QArray(I).WriteQ (1)
    Next I
    Close #1
    MsgBox "Internal question list has been updated successfuly.", vbOKOnly, "Update success"
    cmdCancel.Caption = "Close"
    Updated = True
End Sub

Private Sub cmdUpdateAnswer_Click()
    If txtEditAns.Text = "" Then
        MsgBox "Please write something that will replace the current answer.", vbOKOnly, "Error"
        Exit Sub
    End If
    If cboItem.ListIndex >= 0 Then
        Call QArray(cboItem.ListIndex + 1).SetAnswer(lstAns.ListIndex + 1, txtEditAns.Text)
        Call RefreshQ
        If lstAns.ListIndex < QArray(cboItem.ListIndex + 1).AnswerCount - 1 Then
            lstAns.ListIndex = lstAns.ListIndex + 1
            txtEditAns.SetFocus
            txtEditAns.SelStart = Len(txtEditAns.Text)
        End If
        Updated = False
    Else
        MsgBox "Please choose an answer to edit first.", vbOKOnly, "Error"
    End If
End Sub

Private Sub Form_Load()
    Dim T As String
    QCount = 0
    Dim SetPath As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    Open SetPath For Input As #1
    Dim Folder As String
    Line Input #1, Folder
    Line Input #1, Folder
    Close #1
    Folder = Folder
    If FileExists(Folder + "exam_questionlist.txt") Then
        Open Folder + "exam_questionlist.txt" For Input As #1
        Line Input #1, T
        T = Decode(T)
        QCount = Val(T)
        Dim I As Integer
        For I = 1 To QCount
            Set QArray(I) = New QItem
            QArray(I).ReadQ (1)
        Next I
        Close #1
        Call RefreshList
    End If
    If QCount = 0 Then
        MsgBox "The question list is empty. please click [Edit question list] to add new questions.", vbOKOnly, "Info"
    End If
    Updated = True
End Sub

Private Sub lstAns_Click()
    If lstAns.ListIndex >= 0 Then
        txtEditAns.Text = QArray(cboItem.ListIndex + 1).GetAnswer(lstAns.ListIndex + 1)
    End If
    If lstAns.ListIndex = 0 Then
        cmdAnsUp.Enabled = False
        cmdAnsDown.Enabled = True
    ElseIf lstAns.ListIndex = QArray(cboItem.ListIndex + 1).AnswerCount - 1 Then
        cmdAnsUp.Enabled = True
        cmdAnsDown.Enabled = False
    Else
        cmdAnsUp.Enabled = True
        cmdAnsDown.Enabled = True
    End If
End Sub

Private Sub optFalse_Click()
    QArray(cboItem.ListIndex + 1).CorrectAnswer = "False"
    Updated = False
End Sub

Private Sub optTrue_Click()
    QArray(cboItem.ListIndex + 1).CorrectAnswer = "True"
    Updated = False
End Sub

Private Sub txtEditAns_GotFocus()
    If txtEditAns.Text = "New answer" Then
        txtEditAns.Text = ""
    End If
End Sub

Private Sub txtEditAns_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdUpdateAnswer_Click
End Sub

Private Sub txtEditAns_LostFocus()
    If txtEditAns.Text = "" Then
        txtEditAns.Text = "New answer"
    End If
End Sub

Private Sub txtTruAns_Change()
    If cboItem.ListIndex >= 0 Then
        QArray(cboItem.ListIndex + 1).CorrectAnswer = txtTruAns.Text
        Updated = False
    End If
End Sub

Private Sub txtTruAns_GotFocus()
    If txtTruAns.Text = "New answer" Then txtTruAns.Text = ""
End Sub

Private Sub txtTruAns_LostFocus()
    If txtTruAns.Text = "" Then txtTruAns.Text = "New answer"
End Sub

Private Sub txtQuestion_Change()
    If cboItem.ListIndex >= 0 Then
        QArray(cboItem.ListIndex + 1).QuestionText = txtQuestion.Text
    End If
    Updated = False
End Sub

Private Sub txtQuestion_GotFocus()
    If txtQuestion.Text = "New question" Then txtQuestion.Text = ""
End Sub

Private Sub txtQuestion_LostFocus()
    If txtQuestion.Text = "" Then txtQuestion.Text = "New question"
End Sub

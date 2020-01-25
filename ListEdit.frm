VERSION 5.00
Begin VB.Form QListEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question list editor"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "ListEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Move Up"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Move Down"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ListBox lstQuestion 
      Height          =   3570
      ItemData        =   "ListEdit.frx":00D2
      Left            =   120
      List            =   "ListEdit.frx":00D4
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "QListEdit"
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

Public Sub RefreshList()
    lstQuestion.Clear
    cmdDelete.Enabled = False
    Dim I As Integer
    For I = 1 To QCount
        lstQuestion.AddItem (QArray(I).QuestionText)
        cmdDelete.Enabled = True
    Next I
End Sub

Public Sub Clear()
    QCount = 0
End Sub

Public Sub AddQ(Q As QItem)
    QCount = QCount + 1
    Set QArray(QCount) = Q
End Sub

Private Sub cmdAdd_Click()
    Dim NQ As New QItem
    Call NQ.Init
    If QCount = 0 Or lstQuestion.ListIndex < 0 Then
        Call AddQ(NQ)
    Else
        Dim S, I As Integer
        S = lstQuestion.ListIndex
        For I = QCount To S + 1 Step -1
            Set QArray(I + 1) = QArray(I)
        Next I
        Set QArray(S + 1) = NQ
        QCount = QCount + 1
    End If
    Call RefreshList
    lstQuestion.ListIndex = S
End Sub

Private Sub cmdDelete_Click()
    Dim I As Integer
    Dim S As Integer
    S = lstQuestion.ListIndex
    For I = lstQuestion.ListIndex + 1 To QCount
        Set QArray(I) = QArray(I + 1)
    Next I
    QCount = QCount - 1
    Call RefreshList
    If QCount > 0 Then
        If S = QCount Then
            lstQuestion.ListIndex = S - 1
        Else
            lstQuestion.ListIndex = S
        End If
    End If
    If QCount = 0 Then cmdDelete.Enabled = False
End Sub

Private Sub cmdDone_Click()
    Call QuestionEdit.Clear
    Dim I As Integer
    For I = 1 To QCount
        Call QuestionEdit.AddQ(QArray(I))
    Next I
    Call QuestionEdit.RefreshList
    QuestionEdit.Show
    Unload Me
End Sub

Private Sub cmdDown_Click()
    Dim S As Integer
    S = lstQuestion.ListIndex
    Dim T As QItem
    Set T = QArray(S + 2)
    Set QArray(S + 2) = QArray(S + 1)
    Set QArray(S + 1) = T
    Call RefreshList
    lstQuestion.ListIndex = S + 1
End Sub

Private Sub cmdUp_Click()
    Dim S As Integer
    S = lstQuestion.ListIndex
    Dim T As QItem
    Set T = QArray(S)
    Set QArray(S) = QArray(S + 1)
    Set QArray(S + 1) = T
    Call RefreshList
    lstQuestion.ListIndex = S - 1
End Sub

Private Sub Form_Load()
    If QCount = 0 Then
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    Else
        cmdUp.Enabled = False
        cmdDown.Enabled = True
    End If
End Sub

Private Sub lstQuestion_Click()
    If QCount > 1 Then
        If lstQuestion.ListIndex = 0 Then
            cmdUp.Enabled = False
            cmdDown.Enabled = True
        ElseIf lstQuestion.ListIndex = QCount - 1 Then
            cmdUp.Enabled = True
            cmdDown.Enabled = False
        Else
            cmdUp.Enabled = True
            cmdDown.Enabled = True
        End If
    Else
        cmdUp.Enabled = False
        cmdDown.Enabled = False
    End If
End Sub

VERSION 5.00
Begin VB.Form AnswerEdit 
   Caption         =   "Answer Editor"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
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
      Left            =   4440
      TabIndex        =   4
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtAns 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2880
      Width           =   5655
   End
   Begin VB.ListBox lstAns 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5655
   End
   Begin VB.Label Label2 
      Caption         =   "Selected answer:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Possible wrong answers:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "AnswerEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim QI As QItem

Public Sub SetQItem(Q As QItem)
    Set QI = Q
End Sub

Private Sub cmdAdd_Click()
    QI.AddAnswer ("New fake answer")
    Call RefreshList
End Sub

Private Sub RefreshList()
    lstAns.Clear
    Dim I As Integer
    For I = 1 To QI.AnswerCount
        lstAns.AddItem (QI.GetAnswer(I))
    Next I
    If lstAns.ListIndex < 0 Then
        Delete.Enabled = False
    Else
        Delete.Enabled = True
    End If
End Sub

Private Sub cmdDelete_Click()
    QI.DeleteAnswer (lstAns.ListIndex)
    Call RefreshList
End Sub

Private Sub cmdDone_Click()
    Unload Me
    QuestionEdit.Show
End Sub

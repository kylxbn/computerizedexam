VERSION 5.00
Begin VB.Form ExamFinished 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Examination finished"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   ControlBox      =   0   'False
   Icon            =   "ExamFinished.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "You have finished the exam. Please call the examiner so that you may be guided appropriately. Thank you!"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Congratulations,"
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
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "ExamFinished"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Call UnloadEncoder
    End
End Sub


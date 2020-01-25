VERSION 5.00
Begin VB.Form QuestionEditHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Question editor help"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11175
   Icon            =   "QuestionEditHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   11175
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   2040
      Picture         =   "QuestionEditHelp.frx":00D2
      ScaleHeight     =   5415
      ScaleWidth      =   6975
      TabIndex        =   1
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "REMEMBER TO CLICK [UPDATE] BEFORE EXITING IN ORDER TO SAVE YOUR CHANGE!"
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
      Left            =   120
      TabIndex        =   12
      Top             =   8520
      Width           =   10935
   End
   Begin VB.Label Label11 
      Caption         =   $"QuestionEditHelp.frx":7AC06
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Line Line10 
      X1              =   1920
      X2              =   1560
      Y1              =   5760
      Y2              =   6840
   End
   Begin VB.Label Label10 
      Caption         =   "This button opens the Question list editor, where you can add, delete, or re-order questions."
      Height          =   1215
      Left            =   4200
      TabIndex        =   10
      Top             =   7320
      Width           =   1455
   End
   Begin VB.Line Line9 
      X1              =   4440
      X2              =   4800
      Y1              =   6720
      Y2              =   7200
   End
   Begin VB.Label Label9 
      Caption         =   "This button saves the question list so that it may be used in an examination."
      Height          =   975
      Left            =   8160
      TabIndex        =   9
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Line Line8 
      X1              =   8280
      X2              =   8520
      Y1              =   6720
      Y2              =   7200
   End
   Begin VB.Label Label8 
      Caption         =   "After editing an item in the list, press [Enter] or this button to update the list."
      Height          =   855
      Left            =   9600
      TabIndex        =   8
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Line Line7 
      X1              =   8880
      X2              =   9480
      Y1              =   5640
      Y2              =   4920
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   $"QuestionEditHelp.frx":7ACA7
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Line Line6 
      X1              =   1920
      X2              =   1560
      Y1              =   4680
      Y2              =   4200
   End
   Begin VB.Label Label6 
      Caption         =   "In Multiple choicce and Identification modes, this contains the correct answer."
      Height          =   855
      Left            =   9600
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Line Line5 
      X1              =   8880
      X2              =   9480
      Y1              =   3960
      Y2              =   3840
   End
   Begin VB.Label Label5 
      Caption         =   "Type the question here."
      Height          =   495
      Left            =   9600
      TabIndex        =   5
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Line Line4 
      X1              =   9360
      X2              =   8880
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label4 
      Caption         =   $"QuestionEditHelp.frx":7AD66
      Height          =   1575
      Left            =   9600
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.Line Line3 
      X1              =   9480
      X2              =   8880
      Y1              =   1800
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Set the question type"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   1920
      X2              =   1320
      Y1              =   2280
      Y2              =   2880
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Choose which question to edit"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1320
      Y1              =   1800
      Y2              =   1680
   End
   Begin VB.Label Label1 
      Caption         =   $"QuestionEditHelp.frx":7AE00
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "QuestionEditHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


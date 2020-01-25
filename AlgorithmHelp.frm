VERSION 5.00
Begin VB.Form AlgorithmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Algorithm help"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   Icon            =   "AlgorithmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optAT 
      Caption         =   "Adjacent transposition"
      Height          =   255
      Left            =   3600
      TabIndex        =   21
      Top             =   6840
      Width           =   1935
   End
   Begin VB.OptionButton optLev 
      Caption         =   "Levenshtein"
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   6840
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   1560
      TabIndex        =   12
      Text            =   "Tongue"
      Top             =   6480
      Width           =   3975
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Text            =   "Tang"
      Top             =   6120
      Width           =   3975
   End
   Begin VB.Label Label13 
      Caption         =   $"AlgorithmHelp.frx":00D2
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Width           =   5415
   End
   Begin VB.Label Label19 
      Caption         =   $"AlgorithmHelp.frx":01B4
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   5400
      Width           =   5415
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5520
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label Label18 
      Caption         =   $"AlgorithmHelp.frx":025D
      Height          =   2415
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   5415
   End
   Begin VB.Label Label17 
      Caption         =   "3rd change"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "2nd change"
      Height          =   255
      Left            =   3960
      TabIndex        =   17
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "1st change"
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label14 
      Caption         =   "No change"
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblC 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " 3"
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Changes:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Correct answer:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Student's answer:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Added ""e"""
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Tongue"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Added ""u"""
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Tongu"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Replaced ""a"" with ""o"""
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Original answer"
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Tong"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Tang"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   $"AlgorithmHelp.frx":0579
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "AlgorithmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Label18_Click()
    MsgBox "Huh?", vbOKOnly, "Huh?"
End Sub

Private Sub optAT_Click()
    Call Calc
End Sub

Private Sub optLev_Click()
    Call Calc
End Sub

Private Sub optOSA_Click()
    Call Calc
End Sub

Private Sub txt1_Change()
    Call Calc
End Sub

Private Sub txt2_Change()
    Call Calc
End Sub

Private Sub Calc()
    If optLev.Value Then
        lblC.Caption = Str$(BasicDistance(txt1.Text, txt2.Text))
    Else
        lblC.Caption = Str$(ATDistance(txt1.Text, txt2.Text))
    End If
End Sub

VERSION 5.00
Begin VB.Form DistanceDebug 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "2"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdGetD 
      Caption         =   "Get distance"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "algo"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label lblD 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
   End
End
Attribute VB_Name = "DistanceDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGetD_Click()
    Dim A As Integer
    A = Int(Text3.Text)
    If A = 1 Then
        lblD.Caption = Str$(BasicDistance(Text1.Text, Text2.Text))
    ElseIf A = 2 Then
        lblD.Caption = Str$(OSADistance(Text1.Text, Text2.Text))
    Else
        lblD.Caption = Str$(ATDistance(Text1.Text, Text2.Text))
    End If
End Sub

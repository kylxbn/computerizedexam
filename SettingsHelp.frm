VERSION 5.00
Begin VB.Form SettingsHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings help"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   Icon            =   "SettingsHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label10 
      Caption         =   $"SettingsHelp.frx":00D2
      Height          =   975
      Left            =   360
      TabIndex        =   9
      Top             =   4200
      Width           =   5175
   End
   Begin VB.Label Label9 
      Caption         =   "Mispelling tolerance:"
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
      TabIndex        =   8
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label Label8 
      Caption         =   "for more information."
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   $"SettingsHelp.frx":021E
      Height          =   1215
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   5175
   End
   Begin VB.Label Label5 
      Caption         =   "Mispelling algorithm:"
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
      TabIndex        =   4
      Top             =   2400
      Width           =   5655
   End
   Begin VB.Label Label4 
      Caption         =   $"SettingsHelp.frx":03B2
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   5175
   End
   Begin VB.Label Label3 
      Caption         =   "Detailed student info:"
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
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   $"SettingsHelp.frx":0504
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "Data files folder:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "SettingsHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label7_Click()
    AlgorithmHelp.Show
End Sub

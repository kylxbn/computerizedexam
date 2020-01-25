VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   5655
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   3735
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   3735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Caption         =   "Visual Basic 6.0"
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   5400
      Width           =   2415
   End
   Begin VB.Label Label19 
      Caption         =   "Windows 7"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label18 
      Caption         =   "Windows XP SP3"
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label17 
      Caption         =   "Vienvic Reyes"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label16 
      Caption         =   "John Victor Apostol"
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label Label15 
      Caption         =   "Mariz Peña"
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "Anavy Biatris Francisco"
      Height          =   255
      Left            =   720
      TabIndex        =   13
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label14 
      Caption         =   "(tar.shoduze@gmail.com)"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label13 
      Caption         =   "Prof. Jeff Macnell De Jesus"
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label12 
      Caption         =   "Prof. Mark Julius Kwan"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label11 
      Caption         =   "Prof. Marvic Ablaza"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Additional help:"
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
      Top             =   3240
      Width           =   2895
   End
   Begin VB.Label Label10 
      Caption         =   "Programmed in:"
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
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Label Label9 
      Caption         =   "Operating systems used:"
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
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "Contributing programmers:"
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
      TabIndex        =   5
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "Kyle Alexander Buan"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Programmed by:"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   " FCAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version 9.2"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "E-examination"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

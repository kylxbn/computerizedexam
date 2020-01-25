VERSION 5.00
Begin VB.Form frmNewFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New folder"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   ControlBox      =   0   'False
   Icon            =   "frmNewFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6000
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cndCreate 
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtFolderName 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   1140
      Width           =   4155
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Target Folder:"
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "New Folder:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblTargetFolder 
      Caption         =   "target folder"
      Height          =   915
      Left            =   1500
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
End
Attribute VB_Name = "frmNewFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TargetFolder As String
Public NewFolder As String
Public IsCancel As Boolean

Private Sub cmdCancel_Click()
    IsCancel = True
    Me.Hide
End Sub

Private Sub cndCreate_Click()
    IsCancel = False
    NewFolder = Trim(txtFolderName.Text)
    Me.Hide
End Sub

Private Sub Form_Load()
    lblTargetFolder.Caption = TargetFolder
    IsCancel = False
End Sub

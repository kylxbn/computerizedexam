VERSION 5.00
Begin VB.Form ViewResults 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose a result"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   ControlBox      =   0   'False
   Icon            =   "ViewResults.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "View"
      Enabled         =   0   'False
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
      Left            =   3480
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ListBox lstFiles 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "ViewResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FN(1000) As String
Dim FP As Integer
Dim Folder As String

Private Sub cmdCancel_Click()
    Entry.Show
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If MsgBox("Are you sure that you want to delete " & FN(lstFiles.ListIndex + 1) & "?", vbYesNo, "Confirmation") = vbYes Then
        Kill Folder + FN(lstFiles.ListIndex + 1)
    End If
    Call Form_Load
End Sub

Private Sub cmdView_Click()
    ViewResult.Show
    ViewResult.OpenFile (FN(lstFiles.ListIndex + 1))
    Unload Me
End Sub

Private Sub Form_Load()
    lstFiles.Clear
    cmdView.Enabled = False
    cmdDelete.Enabled = False
    FP = 0
    Dim SetPath As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    Open SetPath For Input As #1
    Dim Path As String
    Line Input #1, Path
    Line Input #1, Path
    Folder = Path
    Close #1
    Dim FName As String
    FName = Dir(Path)
    Do While FName <> ""
        If Right$(FName, 21) = "_encryptedresults.txt" Then
            FP = FP + 1
            FN(FP) = FName
            lstFiles.AddItem (FName)
        End If
        FName = Dir()
    Loop
End Sub

Private Sub lstFiles_Click()
    cmdView.Enabled = True
    cmdDelete.Enabled = True
End Sub

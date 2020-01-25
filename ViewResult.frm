VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form ViewResult 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exam Results"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13995
   ControlBox      =   0   'False
   Icon            =   "ViewResult.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   13995
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser Browser 
      Height          =   9135
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   13695
      ExtentX         =   24156
      ExtentY         =   16113
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   120
      Top             =   9240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   10320
      TabIndex        =   1
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save (unencrypted)"
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
      Left            =   11880
      TabIndex        =   0
      Top             =   9360
      Width           =   2055
   End
End
Attribute VB_Name = "ViewResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Path As String
Dim FName As String

Public Sub OpenFile(S As String)
    FName = S
    Dim SetPath As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    Open SetPath For Input As #1
    Line Input #1, Path
    Line Input #1, Path
    Close #1
    Open Path + S For Input As #1
    Open Path + "decoded.html" For Output As #2
    Dim T As String
    Do While Not EOF(1)
        Line Input #1, T
        Print #2, Decode(T)
    Loop
    Close #1
    Close #2
    Browser.Navigate (Path + "decoded.html")
End Sub

Private Sub cmdClose_Click()
    Kill Path + "decoded.html"
    ViewResults.Show
    Unload Me
End Sub

Private Sub cmdSave_Click()
    CommonDialog.Filter = "Web pages (*.html)|*.html|All files (*.*)|*.*"
    CommonDialog.DefaultExt = "html"
    CommonDialog.FileName = "exported_report.html"
    CommonDialog.DialogTitle = "Save file"
    CommonDialog.ShowSave

    Dim NewPath As String
    NewPath = CommonDialog.FileName
    If Path <> "" Then
        Open NewPath For Output As #2
        Open Path + FName For Input As #1
        Dim T As String
        Do While Not EOF(1)
           Line Input #1, T
            Print #2, Decode(T)
        Loop
        Close #1
        Close #2
        MsgBox "File saved.", vbOKOnly, "Save file"
    End If
End Sub

VERSION 5.00
Begin VB.Form Entry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main menu"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2175
   ControlBox      =   0   'False
   Icon            =   "Entry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   2175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Program settings"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Log out"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1935
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About this program"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   1935
   End
   Begin VB.CommandButton cmdResults 
      Caption         =   "See the exam results"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdQuestions 
      Caption         =   "Edit the question files"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdExam 
      Caption         =   "Take the exam"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAbout_Click()
    About.Show
End Sub

Private Sub cmdExam_Click()
    Dim SetPath As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    Open SetPath For Input As #1
    Dim Folder As String
    Line Input #1, Folder
    Line Input #1, Folder
    Close #1
    Folder = Folder
    If Not FileExists(Folder + "exam_questionlist.txt") Then
        MsgBox "You need to create a question list first." + Chr(10) + "Please click [Edit the question files] before starting an exam.", vbOKOnly, "No question list available"
    ElseIf MsgBox("Due to security reasons, you will be automatically logged out before the exam starts." + Chr(10) + _
            "Are you sure you want to start the exam?", vbYesNo, "Begin exam") = vbYes Then
        Info.Show
        Unload Me
    End If
End Sub

Private Sub cmdExit_Click()
    Call UnloadEncoder
    Login.Show
    Unload Me
End Sub

Private Sub cmdQuestions_Click()
    QuestionEdit.Show
    Unload Me
End Sub

Private Sub cmdResults_Click()
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
    Close #1
    Dim FilesFound As Boolean
    FilesFound = False
    Dim FName As String
    FName = Dir(Path)
    Do While FName <> ""
        If Right$(FName, 21) = "_encryptedresults.txt" Then FilesFound = True
        FName = Dir()
    Loop
    If FilesFound Then
        ViewResults.Show
        Unload Me
    Else
        MsgBox "An exam has not yet been done.", vbOKOnly, "Viewer error"
    End If
End Sub

Private Sub cmdSettings_Click()
    Settings.Show
    Unload Me
End Sub


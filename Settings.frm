VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program settings"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Mispelling tolerance settings"
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   4455
      Begin MSComctlLib.Slider hsbChanges 
         Height          =   255
         Left            =   1200
         TabIndex        =   15
         Top             =   1320
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   5
         SelStart        =   1
         Value           =   1
      End
      Begin VB.OptionButton optAT 
         Caption         =   "Adjacent transpositions"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   480
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optBasic 
         Caption         =   "Levenshtein"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optOSA 
         Caption         =   "Optimal string alignment"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Algorithm:"
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
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Tolerance:"
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
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblChanges 
         Caption         =   "1 change"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   1080
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdSettingsHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   615
   End
   Begin VB.CheckBox chkDetailed 
      Caption         =   "Detailed student info"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset settings"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdChangeDir 
      Caption         =   "..."
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtFolder 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "C:\WINDOWS\Temp\"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Data files folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Entry.Show
    Unload Me
End Sub

Private Sub cmdChangeDir_Click()
    frmChooseDirectory.TargetDirectory = txtFolder.Text
    frmChooseDirectory.WindowTitle = "Choose a folder"

    frmChooseDirectory.Show vbModal

    If Not frmChooseDirectory.IsCancel Then
        txtFolder.Text = frmChooseDirectory.TargetDirectory
    End If
End Sub

Private Sub cmdReset_Click()
    If MsgBox("This will delete all program configurations, question lists," & Chr(10) & "and exam results. This is useful if you forgot your password" & Chr(10) & "and want to set a new password. Continue?", vbYesNo, "Reset Program") = vbYes Then
        If MsgBox("Are you sure? This will delete" & Chr(10) & "E-V-E-R-Y-T-H-I-N-G-!", vbYesNo, "Reset Program") = vbYes Then
            Dim SetPath As String
            If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
                SetPath = "C:\Users\Public\exam_settings.txt"
            Else
                SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
            End If
            Open SetPath For Input As #1
            Dim T, Folder As String
            Line Input #1, T
            Line Input #1, T
            Close #1
            Folder = Mid$(T, 1, Len(T))
            If FileExists(Folder + "exam_questionlist.txt") Then Kill Folder + "exam_questionlist.txt"
            Kill SetPath
            MsgBox "Reset successful. After clicking OK, the program will exit." & Chr(10) & "Please start the program again.", vbOKOnly, "Successful"
            End
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    Dim SetPath As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    Open SetPath For Output As #1
    Dim T As String
    T = "QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm"
    Print #1, Encode(T)
    T = txtFolder.Text
    If Right$(T, 1) <> "\" Then T = T + "\"
    Print #1, T
    If chkDetailed.Value = vbChecked Then
        Print #1, Encode("True")
    Else
        Print #1, Encode("False")
    End If
    If optBasic.Value = True Then
        Print #1, Encode("Basic")
    ElseIf optOSA.Value = True Then
        Print #1, Encode("Optimal string alignment")
    Else
        Print #1, Encode("Adjacent transpositions")
    End If
    Print #1, Encode(Str$(hsbChanges.Value))
    Close #1
    MsgBox "Settings saved.", vbOKOnly, "Save success"
    cmdCancel.Caption = "Close"
End Sub

Private Sub cmdSettingsHelp_Click()
    SettingsHelp.Show
End Sub

Private Sub Form_Load()
    Dim SetPath As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    Open SetPath For Input As #1
    Dim T As String
    Line Input #1, T
    T = Decode(T)
    Line Input #1, T
    txtFolder.Text = T
    Line Input #1, T
    T = Decode(T)
    If T = "True" Then
        chkDetailed.Value = vbChecked
    Else
        chkDetailed.Value = vbUnchecked
    End If
    Line Input #1, T
    T = Decode(T)
    If T = "Basic" Then
        optBasic.Value = True
    ElseIf T = "Optimal string alignment" Then
        optOSA.Value = True
    Else
        optAT.Value = True
    End If
    Line Input #1, T
    T = Decode(T)
    hsbChanges.Value = Int(T)
    Close #1
End Sub

Private Sub hsbChanges_Change()
    Dim C As Integer
    C = hsbChanges.Value
    If C = 1 Then
        lblChanges.Caption = "1 change"
    Else
        lblChanges.Caption = C & " changes"
    End If
End Sub

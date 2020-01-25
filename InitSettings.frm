VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form InitSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "First-run settings"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   ControlBox      =   0   'False
   Icon            =   "InitSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSettingsHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mispelling tolerance settings"
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4455
      Begin MSComctlLib.Slider hsbChanges 
         Height          =   255
         Left            =   1200
         TabIndex        =   13
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
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optOSA 
         Caption         =   "Optimal string alignment"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton optBasic 
         Caption         =   "Levenshtein"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblChanges 
         Caption         =   "1 change"
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   1080
         Width           =   3015
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
         TabIndex        =   10
         Top             =   1080
         Width           =   975
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
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CheckBox chkDetailed 
      Caption         =   "Detailed student info"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtFolder 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "C:\WINDOWS\Temp\"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdChangeDir 
      Caption         =   "..."
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   375
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
      Left            =   3360
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Data files folder:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "InitSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdChangeDir_Click()
    frmChooseDirectory.TargetDirectory = txtFolder.Text
    frmChooseDirectory.WindowTitle = "Choose a folder"

    frmChooseDirectory.Show vbModal

    If Not frmChooseDirectory.IsCancel Then
        txtFolder.Text = frmChooseDirectory.TargetDirectory
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
    If FileExists(SetPath) = False Then MsgBox "Not found!"
    MsgBox "Settings saved.", vbOKOnly, "Save success"
    Entry.Show
    Unload Me
End Sub

Private Sub cmdSettingsHelp_Click()
    SettingsHelp.Show
End Sub

Private Sub Form_Load()
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        txtFolder.Text = "C:\Users\Public\FCATExam\"
    Else
        txtFolder.Text = "C:\Documents and Settings\All Users\Documents\FCATExam"
    End If
End Sub

Private Sub hsbChanges_Change()
    lblChanges.Caption = Str(hsbChanges.Value) & " changes"
End Sub

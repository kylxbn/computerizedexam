VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log in"
   ClientHeight    =   3855
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   7230
   ControlBox      =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      Picture         =   "Login.frx":00D2
      ScaleHeight     =   1215
      ScaleWidth      =   7215
      TabIndex        =   10
      Top             =   1320
      Width           =   7215
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset program"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Log in"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtPw2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3000
      Width           =   5535
   End
   Begin VB.TextBox txtPw1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2640
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      Picture         =   "Login.frx":173F6
      ScaleHeight     =   1095
      ScaleWidth      =   6975
      TabIndex        =   0
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Pasword:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Repeat password:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim InitSettingsFound As Boolean
Dim ResetCounter As Integer

Private Sub cmdAbout_Click()
    About.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdHelp_Click()
    LoginHelp.Show
End Sub

Private Sub cmdLogin_Click()
    If txtPw1.Text = "" Or txtPw1.Text = " " Then
        MsgBox "Your password cannot be blank or a space character.", vbOKOnly, "Password error"
    ElseIf txtPw1.Text = txtPw2.Text Or InitSettingsFound Then
        Dim SetPath As String
        If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
            SetPath = "C:\Users\Public\exam_settings.txt"
        Else
            SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
        End If
        If FileExists(SetPath) = False Then
            frmWait.Show
            frmWait.Refresh
            Call InitializeEncoder(txtPw1.Text, False)
            Unload frmWait
            InitSettings.Show
            Unload Me
        Else
            frmWait.Show
            frmWait.Refresh
            Call InitializeEncoder(txtPw1.Text, False)
            Unload frmWait
            Open SetPath For Input As #1
            Dim T As String
            Dim Ex As Boolean
            Line Input #1, T
            Close #1
            T = Decode(T)
            Dim DecRes As Integer
            DecRes = CheckDecRes()
            If T = "False" Then Ex = True Else Ex = True
            If DecRes < 0 Then
                txtPw1.Text = ""
                txtPw2.Text = ""
                txtPw1.SetFocus
                Exit Sub
            End If
            If T <> "QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm" And DecRes >= 0 Then
                MsgBox "The settings file did not decode properly, but what's weird is that the decoder module did not notice this. That's almost impossible. Anyway, this message is here to say that you might have entered the password wrong, or the settings file is corrupted. You will not be allowed entry until you fix this :)", vbOKOnly, "A very rare error, congratulations!"
            Else
                Entry.Show
                Unload Me
            End If
        End If
    Else
        MsgBox "The two passwords you entered do not match.", vbOKOnly, "Log-in error"
        txtPw1.Text = ""
        txtPw2.Text = ""
        txtPw1.SetFocus
    End If
    ResetCounter = 0
End Sub

Private Sub cmdReset_Click()
    If MsgBox("This will delete all program configurations, question lists," & Chr(10) & "and exam results. This is useful if you forgot your password" & Chr(10) & "and want to set a new password. Continue?", vbYesNo, "Reset Program") = vbYes Then
        If MsgBox("Are you sure? This will delete" & Chr(10) & "E-V-E-R-Y-T-H-I-N-G-!" & Chr(10) & "(Of course, if you have a backup copy, they will not be affected)", vbYesNo, "Reset Program") = vbYes Then
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

Private Sub Form_Load()
    Call EncoderSetExit(True, 3)
    Dim SetPath As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    If FileExists(SetPath) = False Then
        txtPw2.Visible = True
        Label2.Visible = True
        InitSettingsFound = False
    Else
        txtPw2.Visible = False
        Label2.Visible = False
        InitSettingsFound = True
    End If
End Sub

Private Sub Picture1_DblClick()
    If ResetCounter = 0 Or ResetCounter = 3 Then
        ResetCounter = ResetCounter + 1
    Else
        ResetCounter = 0
    End If
End Sub

Private Sub Picture2_DblClick()
    If ResetCounter = 1 Or ResetCounter = 2 Or ResetCounter = 4 Or ResetCounter = 5 Then
        ResetCounter = ResetCounter + 1
    Else
        ResetCounter = 0
    End If
    If ResetCounter = 6 Then
        ResetCounter = 0
        Call cmdReset_Click
    End If
End Sub

Private Sub txtPw1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdLogin_Click
End Sub

Private Sub txtPw2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdLogin_Click
End Sub

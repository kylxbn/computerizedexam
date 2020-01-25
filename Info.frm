VERSION 5.00
Begin VB.Form Info 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration"
   ClientHeight    =   2415
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   9150
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   9150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   22
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtDOBD 
      Height          =   285
      Left            =   2520
      MaxLength       =   3
      TabIndex        =   21
      Text            =   "Day"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   1920
      Width           =   1215
   End
   Begin VB.ComboBox cboDOBM 
      Height          =   315
      ItemData        =   "Info.frx":00D2
      Left            =   1200
      List            =   "Info.frx":00FA
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
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
      Left            =   7920
      TabIndex        =   18
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox txtTel 
      Height          =   285
      Left            =   1200
      TabIndex        =   17
      Top             =   1560
      Width           =   7815
   End
   Begin VB.TextBox txtAdd 
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Top             =   1200
      Width           =   7815
   End
   Begin VB.ComboBox cboCivil 
      Height          =   315
      ItemData        =   "Info.frx":0160
      Left            =   7800
      List            =   "Info.frx":016D
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtNation 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   840
      Width           =   5535
   End
   Begin VB.TextBox txtPOB 
      Height          =   285
      Left            =   5160
      TabIndex        =   9
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox txtDOBY 
      Height          =   285
      Left            =   3120
      MaxLength       =   4
      TabIndex        =   7
      Text            =   "Year"
      Top             =   480
      Width           =   495
   End
   Begin VB.ComboBox cboSex 
      Height          =   315
      ItemData        =   "Info.frx":018B
      Left            =   7920
      List            =   "Info.frx":0195
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtMI 
      Height          =   285
      Left            =   6360
      MaxLength       =   4
      TabIndex        =   3
      Text            =   "Middle Initial"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtFN 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Text            =   "First Name"
      Top             =   120
      Width           =   3255
   End
   Begin VB.TextBox txtLN 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "Last Name"
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblContact 
      Caption         =   "Contact No.:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblAddress 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblCivil 
      Caption         =   "Civil Status:"
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblNation 
      Caption         =   "Nationality:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lblPOB 
      Caption         =   "Place of Birth:"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblDOB 
      Caption         =   "Date of Birth:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblSex 
      Caption         =   "Sex:"
      Height          =   255
      Left            =   7440
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CompleteInfo As Boolean

Private Function CheckCompleteData(Complete As Boolean) As Boolean
    Dim Res As Boolean
    Res = True
    If txtLN.Text = "Last Name" Then Res = False
    If txtFN.Text = "First Name" Then Res = False
    If txtMI.Text = "Middle Initial" Then Res = False
    If Complete Then
        If cboSex.Text = "" Then Res = False
        If txtPOB.Text = "" Then Res = False
        If txtNation.Text = "" Then Res = False
        If cboCivil.Text = "" Then Res = False
        If txtAdd.Text = "" Then Res = False
        If txtTel.Text = "" Then Res = False
    End If
    CheckCompleteData = Res
End Function

Private Function CheckValidDate() As Boolean
    If cboDOBM.Text = "" Then
        CheckValidDate = False
        Exit Function
    ElseIf Val(txtDOBD.Text) = 0 Then
        CheckValidDate = False
        Exit Function
    ElseIf Val(txtDOBY.Text) = 0 Then
        CheckValidDate = False
        Exit Function
    End If
    If cboDOBM.Text <> "February" Then
        If cboDOBM.Text = "January" Or cboDOBM.Text = "March" Or cboDOBM.Text = "May" Or cboDOBM.Text = "July" Or cboDOBM.Text = "August" Or cboDOBM.Text = "October" Or cboDOBM.Text = "December" Then
            If Val(txtDOBD.Text) > 31 Or Val(txtDOBD.Text) < 0 Then
                CheckValidDate = False
                Exit Function
            End If
        Else
            If Val(txtDOBD.Text) > 30 Or Val(txtDOBD.Text) < 0 Then
                CheckValidDate = False
                Exit Function
            End If
        End If
    Else
        ' 2012 is leap year
        If Val(txtDOBY.Text) Mod 4 = 0 Then
            If Val(txtDOBD.Text) > 29 Or Val(txtDOBD.Text) < 0 Then
                CheckValidDate = False
                Exit Function
            End If
        Else
            If Val(txtDOBD.Text) > 28 Or Val(txtDOBD.Text) < 0 Then
                CheckValidDate = False
                Exit Function
            End If
        End If
    End If
    CheckValidDate = True
End Function

Private Sub cmdCancel_Click()
    Call UnloadEncoder
    Login.Show
    Unload Me
End Sub

Private Sub cmdClear_Click()
    txtLN.Text = ""
    txtFN.Text = ""
    txtMI.Text = ""
    cboSex.Text = "N/A"
    cboDOBM.Text = "Month"
    txtDOBD.Text = "Day"
    txtDOBY.Text = "Year"
    txtPOB.Text = ""
    txtNation.Text = ""
    cboCivil.Text = "N/A"
    txtAdd.Text = ""
    txtTel.Text = ""
End Sub

Private Sub cmdSubmit_Click()
    ' Verify date
    If Not CheckValidDate() And CompleteInfo Then
        MsgBox "Invalid date entered.", vbOKOnly, "Error"
    ElseIf Not CheckCompleteData(CompleteInfo) Then
        MsgBox "Some details were missing.", vbOKOnly, "Error"
    Else
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
        Open Folder + "exam_studentdata.txt" For Output As #1
        Print #1, Encode(txtLN.Text)
        Print #1, Encode(txtFN.Text)
        Print #1, Encode(txtMI.Text)
        Print #1, Encode(cboSex.Text)
        Print #1, Encode(cboDOBM.Text)
        Print #1, Encode(txtDOBD.Text)
        Print #1, Encode(txtDOBY.Text)
        Print #1, Encode(txtPOB.Text)
        Print #1, Encode(txtNation.Text)
        Print #1, Encode(cboCivil.Text)
        Print #1, Encode(txtAdd.Text)
        Print #1, Encode(txtTel.Text)
        Close #1
        Exam.Show
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Call SetError("Please inform the examiner.")
    Dim SetPath As String
    If GetWindowsVersion() = "Windows 7/Server 2008 R2" Then
        SetPath = "C:\Users\Public\exam_settings.txt"
    Else
        SetPath = "C:\Documents and Settings\All Users\Documents\exam_settings.txt"
    End If
    Open SetPath For Input As #1
    Dim T As String
    Line Input #1, T
    Line Input #1, T
    Line Input #1, T
    Close #1
    T = Decode(T)
    If T = "False" Then
        CompleteInfo = False
    Else
        CompleteInfo = True
    End If
    If Not CompleteInfo Then
        lblSex.Visible = False
        cboSex.Visible = False
        lblDOB.Visible = False
        cboDOBM.Visible = False
        txtDOBD.Visible = False
        txtDOBY.Visible = False
        lblPOB.Visible = False
        txtPOB.Visible = False
        lblNation.Visible = False
        txtNation.Visible = False
        lblCivil.Visible = False
        cboCivil.Visible = False
        lblAddress.Visible = False
        txtAdd.Visible = False
        lblContact.Visible = False
        txtTel.Visible = False
    End If
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = 44 Or KeyAscii = 32 Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDOBD_GotFocus()
    If txtDOBD.Text = "Day" Then
        txtDOBD.Text = ""
        txtDOBD.MaxLength = 2
    End If
End Sub


Private Sub txtDOBD_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDOBY_GotFocus()
    If txtDOBY.Text = "Year" Then txtDOBY.Text = ""
End Sub

Private Sub txtDOBD_LostFocus()
    If txtDOBD.Text = "" Then
        txtDOBD.MaxLength = 3
        txtDOBD.Text = "Day"
    End If
End Sub

Private Sub txtDOBY_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtDOBY_LostFocus()
    If txtDOBY.Text = "" Then txtDOBY.Text = "Year"
End Sub

Private Sub txtFN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = 32 Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 241 Or KeyAscii = 209 Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtLN_GotFocus()
    If txtLN.Text = "Last Name" Then txtLN.Text = ""
End Sub

Private Sub txtFN_GotFocus()
    If txtFN.Text = "First Name" Then txtFN.Text = ""
End Sub

Private Sub txtLN_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = 32 Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 241 Or KeyAscii = 209 Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtMI_GotFocus()
    If txtMI.Text = "Middle Initial" Then txtMI.Text = ""
End Sub

Private Sub txtLN_LostFocus()
    If txtLN.Text = "" Then txtLN.Text = "Last Name"
End Sub

Private Sub txtFN_LostFocus()
    If txtFN.Text = "" Then txtFN.Text = "First Name"
End Sub

Private Sub txtMI_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = 32 Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtMI_LostFocus()
    If txtMI.Text = "" Then txtMI.Text = "Middle Initial"
End Sub

Private Sub txtNation_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = 32 Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtPOB_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or KeyAscii = 44 Or KeyAscii = 32 Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Or (KeyAscii >= 48 And KeyAscii <= 57) Then
        Exit Sub
    Else
        KeyAscii = 0
    End If
End Sub

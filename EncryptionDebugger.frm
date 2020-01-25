VERSION 5.00
Begin VB.Form EncryptionDebugger 
   Caption         =   "Encryption Debugger"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Decode"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encode"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2130
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1095
   End
End
Attribute VB_Name = "EncryptionDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call InitializeEncoder(Text2.Text, True)
    Text1.Text = Encode(Text1.Text)
End Sub

Private Sub Command2_Click()
    Call InitializeEncoder(Text2.Text, True)
    Text1.Text = Decode(Text1.Text)
End Sub

Private Sub Form_Load()
    Text1.Text = "For God so loved the world that He gave His own begotten Son that whosoever believes in Him should not perish but have everlasting life."
End Sub

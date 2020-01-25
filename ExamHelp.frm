VERSION 5.00
Begin VB.Form ExamHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exam help"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   Icon            =   "ExamHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Caption         =   "In True or False or Multiple choice questions, just choose your answer then click [Submit]!"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   $"ExamHelp.frx":00D2
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "In Identification questions, just type your answer on the box, then press [Submit]."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "ExamHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please wait..."
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2295
   ControlBox      =   0   'False
   Enabled         =   0   'False
   Icon            =   "frmWait.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   2295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Caption         =   "is being configured."
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait while encryption "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


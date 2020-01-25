VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmChooseDirectory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose directory"
   ClientHeight    =   4185
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5385
   ControlBox      =   0   'False
   Icon            =   "frmChooseDirectory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   3360
      Width           =   1155
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "Choose"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Top             =   3360
      Width           =   1155
   End
   Begin MSComctlLib.TreeView treDirectory 
      Height          =   2355
      Left            =   300
      TabIndex        =   0
      Top             =   420
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4154
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
   End
   Begin VB.Menu PopMenu 
      Caption         =   "pop menu"
      Begin VB.Menu CreateFolder 
         Caption         =   "Create New Folder"
      End
   End
End
Attribute VB_Name = "frmChooseDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsCancel As Boolean
Public TargetDirectory As String
Public WindowTitle As String

Private FileSystem As FileSystemObject
Private InScan As Boolean
Private IsDrive As Boolean
Private IsInitialLoad As Boolean
Private IsInExit As Boolean

Private Sub cmdCancel_Click()
    TargetDirectory = ""
    IsCancel = True
    IsInExit = True
    Me.Hide
End Sub

Private Sub cmdChoose_Click()
    TargetDirectory = treDirectory.SelectedItem.Key
    IsCancel = False
    IsInExit = True
    Me.Hide
End Sub

Private Sub CreateFolder_Click()
    Dim CreateForm As frmNewFolder
    Dim NewFolder As String
    Dim aNode As Node

    ' always get a new instance to have a clean fomr
    Set CreateForm = New frmNewFolder
    CreateForm.TargetFolder = treDirectory.SelectedItem.Key

    ' allow the user to enter the name
    CreateForm.Show vbModal

    If Not CreateForm.IsCancel Then
        NewFolder = CreateForm.TargetFolder & CreateForm.NewFolder

        ' if the name is duplicated or is an illegal name
        On Error GoTo NO_CREATE
        MkDir NewFolder

        ' add it to the tree, make sure the user can see it, and give it a blue color
        Set aNode = treDirectory.Nodes.Add(treDirectory.SelectedItem.Key, tvwChild, NewFolder & "\", CreateForm.NewFolder)
        aNode.EnsureVisible
        aNode.ForeColor = RGB(0, 0, 255)
    End If

    On Error GoTo 0

    Unload CreateForm

    Exit Sub

NO_CREATE:
    MsgBox "Error creating folder: " & vbCrLf & vbCrLf & NewFolder & vbCrLf & vbCrLf & Err.Description
    On Error GoTo 0

    Unload CreateForm
End Sub

Private Sub Form_Activate()
    Dim aNode As Node

    Me.Caption = WindowTitle

    If Not IsDrive Then
        MsgBox "Cannot locate any active drives!", vbOKOnly + vbCritical
        IsInExit = True
        Unload Me
        Exit Sub
    End If

    IsInExit = False

    If Not IsInitialLoad Then
        LoadDriveFolders
    End If

    If TargetDirectory <> "" Then
        For Each aNode In treDirectory.Nodes
            DoEvents

            If IsInExit Then
                Exit For
            End If

            If aNode.Key = TargetDirectory Then
                ' this will defeat re-scanning a directory which is already scanned
                InScan = True
                aNode.Selected = True
                aNode.EnsureVisible
                InScan = False
                Exit For
            End If
        Next
    End If
End Sub

Private Sub Form_Load()
    Dim aDrive As Drive
    Dim aFolder As Folder
    Dim FirstNode As Node
    Dim aNode As Node
    Dim DriveName As String

    InScan = False
    IsDrive = False
    IsInitialLoad = False

    PopMenu.Visible = False
    Set FileSystem = CreateObject("Scripting.FileSystemObject")
    treDirectory.Nodes.Add , , "TOP", "Drives"

    Set FirstNode = Nothing

    For Each aDrive In FileSystem.Drives
        DoEvents

        Select Case aDrive.DriveType
                ' removable, fixed, network
            Case 1, 2, 3:
                If aDrive.IsReady Then
                    IsDrive = True
                    DriveName = UCase(aDrive.DriveLetter) & ": (" & aDrive.VolumeName & ")"

                    Set aNode = treDirectory.Nodes.Add("TOP", tvwChild, aDrive.DriveLetter & ":\", DriveName)

                    If FirstNode Is Nothing Then
                        Set FirstNode = aNode
                        FirstNode.EnsureVisible
                    End If
                End If
        End Select
    Next

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdCancel_Click
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If

    If Me.Width < 5000 Then
        Me.Width = 5000
    End If

    If Me.Height < 5000 Then
        Me.Height = 5000
    End If

    cmdChoose.Top = Me.ScaleHeight - (cmdCancel.Height + 50)
    cmdChoose.Left = Me.ScaleWidth - (cmdCancel.Width + 50)

    cmdCancel.Top = cmdCancel.Top
    cmdCancel.Left = cmdCancel.Left - (cmdChoose.Width + 50)

    treDirectory.Top = 0
    treDirectory.Left = 0
    treDirectory.Width = Me.ScaleWidth
    treDirectory.Height = cmdCancel.Top - 50
End Sub

Private Sub GetSubFolders(ParentNode As Node, ScanDepth As Integer)
    Dim aFolder As Folder
    Dim aNode As Node
    Dim NewKey As String

    If IsInExit Or Not InScan Then
        Exit Sub
    End If

    If Not ParentNode Is Nothing Then
        If ParentNode.Key <> "TOP" Then
            On Error GoTo NEXT_FOLDER

            For Each aFolder In FileSystem.GetFolder(ParentNode.Key).SubFolders
                DoEvents

                If IsInExit Or Not InScan Then
                    Exit For
                End If

                NewKey = ParentNode.Key & aFolder.Name & "\"

                For Each aNode In treDirectory.Nodes
                    DoEvents

                    If aNode.Key = NewKey Or IsInExit Then
                        Exit For
                    End If
                Next

                If aNode Is Nothing Then
                    Set aNode = treDirectory.Nodes.Add(ParentNode.Key, tvwChild, NewKey, aFolder.Name)
                End If

                ' only scan down one level from where we started from
                If ScanDepth > 0 Then
                    GetSubFolders aNode, ScanDepth - 1
                End If

                aNode.Sorted = True

NEXT_FOLDER:
            Next
        End If
    End If
End Sub

Private Sub LoadDriveFolders()
    Dim aNode As Node
    Dim nodeIndex As Integer

    Set aNode = treDirectory.Nodes("TOP").Child

    nodeIndex = aNode.FirstSibling.Index
    GetSubFolders aNode.FirstSibling, 1

    While nodeIndex <> aNode.LastSibling.Index
        DoEvents

        If IsInExit Then
            Exit Sub
        End If

        nodeIndex = aNode.Next.Index
        GetSubFolders treDirectory.Nodes(nodeIndex), 1
    Wend

    IsInitialLoad = True
End Sub

Private Sub treDirectory_Expand(ByVal Node As MSComctlLib.Node)
    If Not InScan Then
        InScan = True
        GetSubFolders Node, 1
        InScan = False
    End If
End Sub

Private Sub treDirectory_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If Not InScan Then
            InScan = True
            GetSubFolders treDirectory.SelectedItem, 1
            InScan = False
        End If

    ElseIf Button = vbRightButton Then
        If treDirectory.SelectedItem.Key <> "TOP" Then
            PopupMenu PopMenu
        End If
    End If
End Sub

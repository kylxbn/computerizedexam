Attribute VB_Name = "FileUtils"
Public Function FileExists(strPath As String) As Boolean
    If Dir$(strPath) <> "" Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Public Function KillFolder(ByVal FullPath As String) As Boolean
    On Error Resume Next
    Dim Fso As New Scripting.FileSystemObject
    If Right(FullPath, 1) = "\" Then FullPath = Left(FullPath, Len(FullPath) - 1)
    If Fso.FolderExists(FullPath) Then
        Fso.DeleteFolder FullPath, True
        KillFolder = Fso.FolderExists(FullPath) = False
    End If
End Function

Public Function FolderExists(ByVal FullPath As String) As Boolean
    Dim Fso As New Scripting.FileSystemObject
    FolderExists = Fso.FolderExists(FullPath)
End Function

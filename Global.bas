Attribute VB_Name = "Global"
' Version history:
'   1.0     First working release
'   1.0.1   Added debug mode, and the functions must now be initialized before use.
'   1.1     Fixed possible security leak where if a repeating message such as "ffffffffff" is encoded,
'           a repeating message like "%syqqqqqqq" is also possible.
'   1.2     Fixed possible security leak where if a message is decoded with a very similar password
'           from the original, then the message is partially readable.
'   1.3     Now able to detect if the used password is wrong.
'   1.4     Now able to detect if the data was altered before decryption (tampered). This routine
'           is also triggered if the password is wrong.
'   1.5     Changed API a bit for password entry--is now not passed multiple times like before.

Option Explicit
Dim IsDebug As Boolean
Dim IsInitiated As Boolean
Dim P As String

Public Sub InitializeEncoder(Pass As String, D As Boolean)
    IsDebug = D
    P = Pass
    IsInitiated = True
End Sub

Public Sub UnloadEncoder()
    P = ""
    IsInitiated = False
End Sub

Public Function Encode(S As String) As String
    If Not IsInitiated Then
        MsgBox "Encoder not yet initiated.", vbOKOnly, "Encoder error"
        End
    End If
    
    Dim I, J, K As Integer
    
    ' Convert String to ASCII while inverting it
    Dim SArray(1000) As Integer
    Dim Original As String
    Original = S
    S = " " + S
    K = Len(S)
    For I = 1 To Len(S)
        SArray(I) = Asc(Mid$(S, K, 1)) - 32
        K = K - 1
    Next I
    
    ' Calculate checksum
    Dim Checksum As Integer
    Checksum = 0
    For I = 1 To Len(S)
        Checksum = Checksum + SArray(I)
    Next I
    Checksum = (Checksum Mod 95) + 32
    
    ' Convert key to ASCII
    Dim PArray(100) As Integer
    For I = 1 To Len(P)
        PArray(I) = Asc(Mid$(P, I, 1)) - 32
    Next I
    
    ' Encode Data
    For I = 1 To Len(S)
        For J = 1 To Len(P)
            SArray(I + J - 1) = (SArray(I + J - 1) + PArray(J)) Mod 95
        Next J
    Next I
    
    ' Normalize Data
    For I = 1 To Len(S)
        SArray(I) = SArray(I) + 32
    Next I
    
    ' Convert Data to String and add checksum
    Dim T As String
    For I = 1 To Len(S)
        T = T + Chr$(SArray(I))
    Next I
    T = Chr$(Checksum) + T
    
    ' Debug
    If IsDebug Then
        If (Len(Original)) <> (Len(T) - 2) Then MsgBox "Length has changed!", vbOKOnly, "Encoder debug"
    End If
    
    Encode = T
    
End Function

Public Function Decode(S As String) As String
    If Not IsInitiated Then
        MsgBox "Encoder not yet initiated.", vbOKOnly, "Encoder error"
        End
    End If
    
    Dim I, J As Integer
    
    ' Convert String to ASCII
    Dim SArray(1000) As Integer
    For I = 1 To Len(S)
        SArray(I) = Asc(Mid$(S, I, 1)) - 32
    Next I
    
    
    
    ' Convert key to ASCII
    Dim PArray(100) As Integer
    For I = 1 To Len(P)
        PArray(I) = Asc(Mid$(P, I, 1)) - 32
    Next I
    
    ' Decode Data
    For I = 2 To Len(S)
        For J = 1 To Len(P)
            If (SArray(I + J - 1) - PArray(J)) < 0 Then SArray(I + J - 1) = SArray(I + J - 1) + 95
            SArray(I + J - 1) = (SArray(I + J - 1) - PArray(J))
        Next J
    Next I
    
    ' Calculate checksum
    Dim Checksum As Integer
    Checksum = 0
    For I = 2 To Len(S)
        Checksum = Checksum + SArray(I)
    Next I
    Checksum = (Checksum Mod 95) + 32
    
    ' Normalize Data
    For I = 1 To Len(S)
        SArray(I) = SArray(I) + 32
    Next I
    
    ' Convert Data to String while inverting it
    Dim T As String
    For I = 2 To Len(S) - 1
        T = Chr$(SArray(I)) + T
    Next I
    
    If IsDebug Then
        If Len(S) <> (Len(T) + 2) Then MsgBox "Length has changed from " + Str$(Len(S)) + " to " + Str$(Len(T) + 2), vbOKOnly, "Encoder debug"
        If SArray(Len(S)) <> 32 Then MsgBox "Wrong password detected " + Str$(SArray(Len(S))), vbOKOnly, "Encoder debug"
        If SArray(1) <> Checksum Then MsgBox "Wrong checksum detected! Data was most probably altered, or you used a wrong password. (expected " + Str$(Checksum) + " but got " + Str$(SArray(1)), vbOKOnly, "Encoder debug"
    Else
        If SArray(Len(S)) <> 32 Then
            MsgBox "Wrong password detected." + Chr(10) + "This means that the password used to create the data files is not the same as the password you are using now." + Chr(10) + "This may be a simple human mistake but may also mean that someone may have edited the data files without permission.", vbOKOnly, "Decoder error"
        ElseIf SArray(1) <> Checksum Then
            MsgBox "Wrong checksum detected." + Chr(10) + "This means that the data files were edited in a text editor (such as Notepad) before they were accessed now." + Chr(10) + "This almost always mean that someone edited the data files without permission.", vbOKOnly, "Decoder error"
        End If
    End If
        
    
    Decode = T
End Function




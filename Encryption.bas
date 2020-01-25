Attribute VB_Name = "Encryption"
' Version history:
'   1.0     First working release
'   1.0.1   Added debug mode, and the functions must now be initialized before use.
'   1.1     Fixed possible security leak where if a repeating message such as "ffffffffff" is encoded,
'           a repeating message like "%syqqqqqqq" is also possible. However, despite that repeating letters
'           is now solved, the effectivity depends on password strength. For example, if the password used is
'           very short (e.g. 4 characters), then encoding "fffffffffffffff" might result in "r7wW9rW9rW9rW9r"
'           so repeating is still present on groups of characters.
'   1.2     Fixed possible security leak where if a message is decoded with a very similar password
'           from the original, then the message is partially readable.
'   1.3     Now able to detect if the used password is wrong.
'   1.4     Now able to detect if the data was altered before decryption (tampered). This routine
'           is also triggered if the password is wrong.
'   1.5     Changed API a bit for password entry--is now not passed multiple times like before.
'   1.6     More secure encryption. Also, several wrong passwords are allowed before forcefully exiting.
'           This counter is set with EncoderSetExit()
'   2.0     After researching more about cryptography and encryption, I made efforts to improve security.
'           Now, the password is NOT used directly--instead, a 448-bit (7 bits * 64) encryption key
'           is derived from the original password and is used for encryption instead of the original password.
'           Because of this, 0-length passwords are NO LONGER ALLOWED even in debug mode (this is a technical
'           limitation and not just my whim). At least 1 character now required. This also means that
'           the problem in version 1.1 is still unsolved, but due to the very long password, the repetition
'           will only occur every 64 characters, which is very unlikely especially when the original
'           message is not repetitive.
'
'   "Scramble" algorithm by Kyle Alexander Buan (tar.shoduze@gmail.com)
'   Latest version released on January 14, 2015

Option Explicit
Dim IsDebug As Boolean
Dim IsInitiated As Boolean
Dim P As String
Dim HasPWError As Boolean
Dim HasCSError As Boolean
Dim ExtError As String
Dim Ex As Boolean
Dim RemainError As Integer
Dim PreviousKey As String

Dim DecRes As Integer

Public Function CheckDecRes() As Integer
    CheckDecRes = DecRes
End Function

Public Sub EncoderSetExit(E As Boolean, C As Integer)
    Ex = E
    RemainError = C - 1
End Sub

Private Function Random() As Integer
    Random = Int(96 * Rnd)
End Function


Private Function Strengthen(P As String) As String
    Dim I, J, L As Integer
    L = Len(P)
    
    ' Convert Password to ASCII
    Dim OldPArray(1000) As Integer
    For I = 1 To L
        OldPArray(I) = Asc(Mid$(P, I, 1)) - 32
    Next I
    
    ' Derive NewPArray from OldPArray
    Dim NewPArray(65) As Integer
    ' Get new key
    For I = 1 To 50000
        Randomize OldPArray(((I - 1) Mod L) + 1)
        For J = 1 To 64
            NewPArray(J) = (NewPArray(J) + Random()) Mod 95
        Next J
    Next I
    
    ' Normalize Data
    For I = 1 To 65
        NewPArray(I) = NewPArray(I) + 32
    Next I
    
    ' Convert Data to String and add checksum
    Dim T As String
    For I = 1 To 65
        T = T + Chr$(NewPArray(I))
    Next I
    
    If IsDebug Then MsgBox PreviousKey & Chr(10) & T, vbOKOnly, "Generated key"
    PreviousKey = T
    Strengthen = T
End Function

Public Sub InitializeEncoder(Pass As String, D As Boolean)
    If Len(Pass) < 1 Then
        MsgBox "Password too short! At least one character required.", vbOKOnly, "Encrypt Init error"
        End
    End If
    IsDebug = D
    P = Strengthen(Pass)
    ExtError = ""
    HasPWError = False
    HasCSError = False
    IsInitiated = True
End Sub

Public Sub UnloadEncoder()
    P = ""
    IsInitiated = False
End Sub

Public Function SetError(E As String)
    ExtError = Chr(10) + E
End Function

Public Function Encode(ByVal S As String) As String
    If Not IsInitiated Then
        MsgBox "Encoder not yet initiated.", vbOKOnly, "Encoder error"
        End
    End If
    
    Dim I, J, K As Integer
    
    If Len(S) = 0 Then
        Encode = ""
        Exit Function
    End If
        
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
    For K = 0 To PArray(Len(P))
        For I = 1 To Len(S) Step ((PArray(2) Mod 3) + 3)
            For J = 1 To Len(P)
                SArray(I + J - 1) = (SArray(I + J - 1) + PArray(J)) Mod 95
            Next J
        Next I
    Next K
    
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

Public Function Decode(ByVal S As String) As String
    If Not IsInitiated Then
        MsgBox "Encoder not yet initiated.", vbOKOnly, "Encoder error"
        End
    End If
    
    Dim I, J, K As Integer
    
    If Len(S) = 0 Then
        Decode = ""
        Exit Function
    End If
    
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
    For K = 0 To PArray(Len(P))
        For I = 2 To Len(S) Step ((PArray(2) Mod 3) + 3)
            For J = 1 To Len(P)
                If (SArray(I + J - 1) - PArray(J)) < 0 Then SArray(I + J - 1) = SArray(I + J - 1) + 95
                SArray(I + J - 1) = (SArray(I + J - 1) - PArray(J))
            Next J
        Next I
    Next K
    
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
    
    DecRes = 0
    If IsDebug Then
        If Len(S) <> (Len(T) + 2) Then MsgBox "Length has changed from " + Str$(Len(S)) + " to " + Str$(Len(T) + 2), vbOKOnly, "Encoder debug"
        If SArray(Len(S)) <> 32 Then MsgBox "Wrong password detected " + Str$(SArray(Len(S))), vbOKOnly, "Encoder debug"
        If SArray(1) <> Checksum Then MsgBox "Wrong checksum detected! Data was most probably altered, or you used a wrong password. (expected " + Str$(Checksum) + " but got " + Str$(SArray(1)), vbOKOnly, "Encoder debug"
    Else
        If SArray(Len(S)) <> 32 Then
            DecRes = -1
            MsgBox "Wrong password detected." + Chr(10) + "This means that the password used to create the data files is not the same as the password you are using now." + Chr(10) + "This may be a simple human mistake but may also mean that someone may have edited the data files without permission." + ExtError, vbOKOnly, "Decoder error"
            If Ex Then
                If RemainError = 0 Then
                    End
                Else
                    RemainError = RemainError - 1
                End If
            End If
        ElseIf SArray(1) <> Checksum Then
            DecRes = -1
            MsgBox "Wrong checksum detected." + Chr(10) + "This means that the data files were edited in a text editor (such as Notepad) before they were accessed now." + Chr(10) + "This may indicate that someone edited the data files without permission," + Chr(10) + "or maybe the password used was just wrong." + ExtError, vbOKOnly, "Decoder error"
            If Ex Then
                If RemainError = 0 Then
                    End
                Else
                    RemainError = RemainError - 1
                End If
            End If
        End If
    End If
    
    Decode = T
End Function




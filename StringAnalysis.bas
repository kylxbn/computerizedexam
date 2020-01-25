Attribute VB_Name = "StringAnalysis"
Option Explicit

Public Function BasicDistance(A As String, B As String) As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Cost As Integer
    Dim Min1 As Integer
    Dim Min2 As Integer
    Dim Min3 As Integer
    If Len(A) = 0 Then
        BasicDistance = Len(B)
        Exit Function
    End If
    If Len(B) = 0 Then
        BasicDistance = Len(A)
        Exit Function
    End If
    Dim D(200, 200) As Integer
    For i = 0 To Len(A)
        D(i, 0) = i
    Next i
    For j = 0 To Len(B)
        D(0, j) = j
    Next j
    For i = 1 To Len(A)
        For j = 1 To Len(B)
            If Mid$(A, i, 1) = Mid$(B, j, 1) Then
                Cost = 0
            Else
                Cost = 1
            End If
            Min1 = D(i - 1, j) + 1
            Min2 = D(i, j - 1) + 1
            Min3 = D(i - 1, j - 1) + Cost
            If Min1 <= Min2 And Min1 <= Min3 Then
                D(i, j) = Min1
            ElseIf Min2 <= Min1 And Min2 <= Min3 Then
                D(i, j) = Min2
            Else
                D(i, j) = Min3
            End If
        Next j
    Next i
    BasicDistance = D(Len(A), Len(B))
End Function

Public Function OSADistance(A As String, B As String) As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Cost As Integer
    Dim Min1 As Integer
    Dim Min2 As Integer
    Dim Min3 As Integer
    If Len(A) = 0 Then
        OSADistance = Len(B)
        Exit Function
    End If
    If Len(B) = 0 Then
        OSADistance = Len(A)
        Exit Function
    End If
    Dim D(200, 200) As Integer
    For i = 0 To Len(A)
        D(i, 0) = i
    Next i
    For j = 0 To Len(B)
        D(0, j) = j
    Next j
    For i = 1 To Len(A)
        For j = 1 To Len(B)
            If Mid$(A, i, 1) = Mid$(B, j, 1) Then
                Cost = 0
            Else
                Cost = 1
            End If
            Min1 = D(i - 1, j) + 1
            Min2 = D(i, j - 1) + 1
            Min3 = D(i - 1, j - 1) + Cost
            If Min1 <= Min2 And Min1 <= Min3 Then
                D(i, j) = Min1
            ElseIf Min2 <= Min1 And Min2 <= Min3 Then
                D(i, j) = Min2
            Else
                D(i, j) = Min3
            End If
            If (i > 1) And (j > 1) And Mid$(A, i + 1, 1) = Mid$(B, j, 1) And Mid$(A, i, 1) = Mid$(A, j + 1, 1) Then
                Min1 = D(i, j)
                Min2 = D(i - 2, j - 2) + Cost
                If Min1 < Min2 Then D(i, j) = Min1 Else D(i, j) = Min2
            End If
        Next j
    Next i
    OSADistance = D(Len(A), Len(B))
End Function

Function ATDistance(s1 As String, s2 As String, Optional limit As Variant, Optional result As Variant) As Integer
    Dim diagonal As Integer
    Dim horizontal As Integer
    Dim vertical As Integer
    Dim swap As Integer
    Dim final As Integer
    If IsMissing(limit) Then
        limit = Len(s1) + Len(s2)
    End If
    If IsMissing(result) Then
        Dim i, j As Integer
        ReDim result(Len(s1), Len(s2)) As Integer
    End If
    If result(Len(s1), Len(s2)) < 1 Then
        If Abs(Len(s1) - Len(s2)) >= limit Then
            final = limit
        Else
            If Len(s1) = 0 Or Len(s2) = 0 Then
                final = Len(s1) + Len(s2)
            Else
                If Mid(s1, 1, 1) = Mid(s2, 1, 1) Then
                    final = ATDistance(Mid(s1, 2), Mid(s2, 2), limit, result)
                Else
                    If Mid(s1, 1, 1) = Mid(s2, 2, 1) And Mid(s1, 2, 1) = Mid(s2, 1, 1) Then
                        swap = ATDistance(Mid(s1, 3), Mid(s2, 3), limit - 1, result)
                        final = 1 + swap
                    Else
                        diagonal = ATDistance(Mid(s1, 2), Mid(s2, 2), limit - 1, result)
                        horizontal = ATDistance(Mid(s1, 2), s2, diagonal, result)
                        vertical = ATDistance(s1, Mid(s2, 2), horizontal, result)
                        final = 1 + vertical
                    End If
                End If
            End If
        End If
    Else
        final = result(Len(s1), Len(s2)) - 1
    End If
    If final < limit Then
        ATDistance = final
        result(Len(s1), Len(s2)) = final + 1
    Else
        ATDistance = limit
    End If
End Function

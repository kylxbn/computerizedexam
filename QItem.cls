VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const TYPE_ENUM = 1
Private Const TYPE_MULTI = 2
Private Const TYPE_TF = 3
Private Const TYPE_ESSAY = 4
Private Const TYPE_IDENT = 5

Private QuestionTextV    As String
Private CorrectAnswerV   As String
Private AnswerText(20)   As String
Private AnswerCountV     As Integer
Private CaseSensitiveV   As Boolean
Private StrictSpellingV  As Boolean
Private InOrderV         As Boolean
Private QTypeV           As Integer

Property Get QuestionText() As String
    QuestionText = QuestionTextV
End Property

Property Let QuestionText(ByVal Q As String)
    QuestionTextV = Q
End Property

Property Get CorrectAnswer() As String
    CorrectAnswer = CorrectAnswerV
End Property

Property Let CorrectAnswer(ByVal Q As String)
    CorrectAnswerV = Q
End Property

Property Get AnswerCount() As Integer
    AnswerCount = AnswerCountV
End Property

Property Get CaseSensitive() As Boolean
    CaseSensitive = CaseSensitiveV
End Property

Property Let CaseSensitive(ByVal Q As Boolean)
    CaseSensitiveV = Q
End Property

Property Get StrictSpelling() As Boolean
    StrictSpelling = StrictSpellingV
End Property

Property Let StrictSpelling(I As Boolean)
    StrictSpellingV = I
End Property

Property Get InOrder() As Boolean
    InOrder = InOrderV
End Property

Property Let InOrder(I As Boolean)
    InOrderV = I
End Property

Property Get QType() As Integer
    QType = QTypeV
End Property

Property Let QType(ByVal Q As Integer)
    If Q = TYPE_MULTI Or Q = TYPE_TF Then
        StrictSpellingV = True
        CaseSensitiveV = True
    End If
    QTypeV = Q
End Property

Public Sub Init()
    QuestionText = "New question"
    CorrectAnswer = "New answer"
    AnswerCountV = 0
    CaseSensitive = False
    StrictSpelling = True
    InOrder = False
    QType = TYPE_IDENT
End Sub

Public Sub AddAnswer(ByVal A As String)
    AnswerCountV = AnswerCountV + 1
    If AnswerCount <= 20 Then
        AnswerText(AnswerCount) = A
    Else
        MsgBox "Add answer: array is full"
    End If
End Sub

Public Sub DeleteAnswer(I As Integer)
    If I > 0 Then
        For I = I To AnswerCount
            AnswerText(I) = AnswerText(I + 1)
        Next I
        AnswerCountV = AnswerCountV - 1
    Else
        MsgBox "Deleting answer: out of bounds"
    End If
End Sub

Public Function GetAnswer(ByVal A As Integer) As String
    If A > AnswerCount Then
        MsgBox "Error accessing fake answers: index out of bounds"
    Else
        GetAnswer = AnswerText(A)
    End If
End Function

Public Function SetAnswer(A As Integer, ByVal S As String)
    If A > AnswerCount Then
        MsgBox "Error accessing fake answers: index out of bounds"
    Else
        AnswerText(A) = S
    End If
End Function

Public Sub WriteQ(C As Integer)
    Print #C, Encode(Replace$(QuestionText, Chr(13) & Chr(10), "\n"))
    Print #C, Encode(CorrectAnswer)
    Print #C, Encode(AnswerCount)
    Dim I As Integer
    For I = 1 To AnswerCount
        Print #C, Encode(AnswerText(I))
    Next I
    If CaseSensitive Then Print #C, Encode("True") Else Print #C, Encode("False")
    If StrictSpelling Then Print #C, Encode("True") Else Print #C, Encode("False")
    If InOrder Then Print #C, Encode("True") Else Print #C, Encode("False")
    Print #C, Encode(Str$(QType))
End Sub

Public Sub WriteQU(C As Integer)
    Print #C, Replace$(QuestionText, Chr(13) & Chr(10), "\n")
    Print #C, CorrectAnswer
    Print #C, AnswerCount
    Dim I As Integer
    For I = 1 To AnswerCount
        Print #C, AnswerText(I)
    Next I
    If CaseSensitive Then Print #C, "True" Else Print #C, "False"
    If StrictSpelling Then Print #C, "True" Else Print #C, "False"
    If InOrder Then Print #C, "True" Else Print #C, "False"
    Print #C, Str$(QType)
End Sub
    
Public Sub ReadQ(C As Integer)
    Dim Temp As String
    
    Line Input #C, Temp
    QuestionText = Decode(Replace$(Temp, "\n", Chr(13) & Chr(10)))
    Line Input #C, Temp
    CorrectAnswer = Decode(Temp)
    Line Input #C, Temp
    AnswerCountV = Int(Decode(Temp))
    Dim I As Integer
    For I = 1 To AnswerCount
        Line Input #C, Temp
        AnswerText(I) = Decode(Temp)
    Next I
    
    Line Input #C, Temp
    Temp = Decode(Temp)
    If Temp = "True" Then CaseSensitive = True Else CaseSensitive = False
    
    Line Input #C, Temp
    Temp = Decode(Temp)
    If Temp = "True" Then StrictSpelling = True Else StrictSpelling = False
    
    Line Input #C, Temp
    Temp = Decode(Temp)
    If Temp = "True" Then InOrder = True Else InOrder = False
    
    Line Input #C, Temp
    QType = Val(Decode(Temp))
End Sub

Public Sub ReadQU(C As Integer)
    Dim Temp As String
    
    Line Input #C, Temp
    QuestionText = Replace$(Temp, "\n", Chr(13) & Chr(10))
    Line Input #C, Temp
    CorrectAnswer = Temp
    Line Input #C, Temp
    AnswerCountV = Int(Temp)
    Dim I As Integer
    For I = 1 To AnswerCount
        Line Input #C, Temp
        AnswerText(I) = Temp
    Next I
    
    Line Input #C, Temp
    Temp = Temp
    If Temp = "True" Then CaseSensitive = True Else CaseSensitive = False
    
    Line Input #C, Temp
    Temp = Temp
    If Temp = "True" Then StrictSpelling = True Else StrictSpelling = False
    
    Line Input #C, Temp
    Temp = Temp
    If Temp = "True" Then InOrder = True Else InOrder = False
    
    Line Input #C, Temp
    QType = Val(Temp)
End Sub


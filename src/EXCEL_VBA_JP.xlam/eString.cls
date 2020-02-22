VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_String As String

Public Property Get Name() As String
    Name = m_String
End Property

Public Property Let Name(Value As String)
    m_String = Value
End Property

Public Property Get forJPfunc() As String
    forJPfunc = m_String
End Property


Public Property Get ToVBA() As String
    ToVBA = m_String
End Property
Public Property Get str() As String
    str = m_String
End Property

Public Property Get Length() As Long
    Length = Len(m_String)
End Property

Public Property Get LeftSide(LengthFromLeftSide As Long) As eString
    Dim w_String As New eString
    w_String = Left(m_String, LengthFromLeftSide)
    Set LeftSide = w_String
    Set w_String = Nothing
End Property

Public Property Get RightSide(LengthFromRightSide As Long) As eString
    Dim w_String As New eString
    w_String = Right(m_String, LengthFromRightSide)
    Set RightSide = w_String
    Set w_String = Nothing
End Property

Public Function StartWith(StartingEString As eString) As Boolean
    Dim l As Long
    If Len(m_String) < Len(StartingEString.ToVBA) Then
        StartWith = False
    Else
        Dim subStr As String
        If Left(m_String, Len(StartingEString.ToVBA)) = StartingEString.ToVBA Then
            StartWith = True
        Else
            StartWith = False
        End If
    End If
End Function

Public Function StartWithString(StartingString As String) As Boolean
    Dim l As Long
    If Len(m_String) < Len(StartingString) Then
        StartWithString = False
    Else
        Dim subStr As String
        If Left(m_String, Len(StartingString)) = StartingString Then
            StartWithString = True
        Else
            StartWithString = False
        End If
    End If
End Function

Public Function EndWith(EndingEString As eString) As Boolean
    Dim l As Long
    If Len(m_String) < Len(EndingEString.ToVBA) Then
        EndWith = False
    Else
        Dim subStr As String
        If Right(m_String, Len(m_String) - Len(EndingEString.ToVBA) + 1) = EndingEString.ToVBA Then
            EndWith = True
        Else
            EndWith = False
        End If
    End If
End Function

Public Function EndWithString(EndingString As String) As Boolean
    Dim l As Long
    If Len(m_String) < Len(EndingString) Then
        EndWithString = False
    Else
        Dim subStr As String
        If Right(m_String, Len(m_String) - Len(EndingString) + 1) = EndingString Then
            EndWithString = True
        Else
            EndWithString = False
        End If
    End If
End Function

Public Function IsBlank() As Boolean
    If Trim(m_String) = "" Then
        IsBlank = True
    Else
        IsBlank = False
    End If
End Function

Public Function PartialMatch(PartialMatchEString As eString) As Boolean
    PartialMatch = InStr(1, m_String, PartialMatchEString.ToVBA) > 0
End Function

Public Function StringPartialMatch(PartialMatchString As String) As Boolean
    StringPartialMatch = InStr(1, m_String, PartialMatchString) > 0
End Function

Public Function ReplaceEString(Target As eString, ReplacingEString As eString) As eString
    m_String = Replace(m_String, Target.ToVBA, ReplacingEString.ToVBA)
    Set ReplaceEString = Me
End Function

Public Function ReplaceSring(Target As String, ReplacingString As String) As String
    m_String = Replace(m_String, Target, ReplacingString)
    ReplaceSring = Me.ToVBA
End Function

Public Function Add(AddingString As String) As String
    m_String = m_String & AddingString
    Add = m_String
End Function






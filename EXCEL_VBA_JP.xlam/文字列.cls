VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "文字列"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_文字列 As String

Public Property Get Name() As String
    Name = m_文字列
End Property

Public Property Let Name(値 As String)
    m_文字列 = 値
End Property

Public Property Get JP関数用() As String
    JP関数用 = m_文字列
End Property


Public Property Get VBAに渡す() As String
    VBAに渡す = m_文字列
End Property
Public Property Get str() As String
    str = m_文字列
End Property

Public Property Get 長さ() As Long
    長さ = Len(m_文字列)
End Property

Public Property Get 右側(右側からの文字数 As Long) As 文字列
    Dim w_文字列 As New 文字列
    w_文字列 = Left(m_文字列, 右側からの文字数)
    Set 右側 = w_文字列
    Set w_文字列 = Nothing
End Property

Public Property Get 左側(左側からの文字数 As Long) As 文字列
    Dim w_文字列 As New 文字列
    w_文字列 = Right(m_文字列, 左側からの文字数)
    Set 左側 = w_文字列
    Set w_文字列 = Nothing
End Property

Public Function 始まっている(始まりの文字列 As 文字列) As Boolean
    Dim l As Long
    If Len(m_文字列) < Len(始まりの文字列.VBAに渡す) Then
        始まっている = False
    Else
        Dim subStr As String
        If Left(m_文字列, Len(始まりの文字列.VBAに渡す)) = 始まりの文字列.VBAに渡す Then
            始まっている = True
        Else
            始まっている = False
        End If
    End If
End Function

Public Function Stringで始まっている(始まりの文字列 As String) As Boolean
    Dim l As Long
    If Len(m_文字列) < Len(始まりの文字列) Then
        Stringで始まっている = False
    Else
        Dim subStr As String
        If Left(m_文字列, Len(始まりの文字列)) = 始まりの文字列 Then
            Stringで始まっている = True
        Else
            Stringで始まっている = False
        End If
    End If
End Function

Public Function 終わっている(終わりの文字列 As 文字列) As Boolean
    Dim l As Long
    If Len(m_文字列) < Len(終わりの文字列.VBAに渡す) Then
        終わっている = False
    Else
        Dim subStr As String
        If Right(m_文字列, Len(m_文字列) - Len(終わりの文字列.VBAに渡す) + 1) = 終わりの文字列.VBAに渡す Then
            終わっている = True
        Else
            終わっている = False
        End If
    End If
End Function

Public Function Stringで終わっている(終わりの文字列 As String) As Boolean
    Dim l As Long
    If Len(m_文字列) < Len(終わりの文字列) Then
        Stringで終わっている = False
    Else
        Dim subStr As String
        If Right(m_文字列, Len(m_文字列) - Len(終わりの文字列) + 1) = 終わりの文字列 Then
            Stringで終わっている = True
        Else
            Stringで終わっている = False
        End If
    End If
End Function

Public Function 空白か() As Boolean
    If Trim(m_文字列) = "" Then
        空白か = True
    Else
        空白か = False
    End If
End Function

Public Function 部分一致(部分一致する文字列 As 文字列) As Boolean
    部分一致 = InStr(1, m_文字列, 部分一致する文字列.VBAに渡す) > 0
End Function

Public Function String部分一致(部分一致する文字列 As String) As Boolean
    String部分一致 = InStr(1, m_文字列, 部分一致する文字列) > 0
End Function

Public Function 置換(置換対象文字列 As 文字列, 置換後文字列 As 文字列) As 文字列
    m_文字列 = Replace(m_文字列, 置換対象文字列.VBAに渡す, 置換後文字列.VBAに渡す)
    Set 置換 = Me
End Function

Public Function Sring置換(置換対象文字列 As String, 置換後文字列 As String) As String
    m_文字列 = Replace(m_文字列, 置換対象文字列, 置換後文字列)
    Sring置換 = Me.VBAに渡す
End Function

Public Function 追加(追加文字列 As String) As String
    m_文字列 = m_文字列 & 追加文字列
    追加 = m_文字列
End Function





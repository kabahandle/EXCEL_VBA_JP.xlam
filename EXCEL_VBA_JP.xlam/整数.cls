VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "整数"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_整数 As Long
Private m_Prefix As String
Private m_Postfix As String
Private m_増分 As Long


Private Sub Class_Initialize()
    'ここに処理を書く
    m_整数 = 0
    m_増分 = 0
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
End Sub

Public Property Get Name() As Long
    Name = m_整数 + m_増分
End Property

Public Property Let Name(値 As Long)
    m_整数 = 値
End Property

Public Sub セット(値 As Long)
    m_整数 = 値
End Sub

Public Sub 値を設定(値 As Long)
    m_整数 = 値
End Sub

Public Property Get JP関数用() As Long
    JP関数用 = m_整数 + m_増分
End Property

Public Property Get VBAに渡す() As Long
    VBAに渡す = m_整数 + m_増分
End Property

Public Sub 増やす(Optional 増分 As Long = 1)
    m_増分 = m_増分 + 増分
End Sub

Public Function 文字列化() As String
    文字列化 = m_Prefix & CStr(m_整数 + m_増分) & m_Postfix
End Function

Public Sub 接頭辞を設定(Optional 接頭辞 As String = "")
    m_Prefix = 接頭辞
End Sub

Public Sub 接尾辞を設定(Optional 接尾辞 As String = "")
    m_Postfix = 接尾辞
End Sub





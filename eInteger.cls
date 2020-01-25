VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eInteger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Integer As Long
Private m_Prefix As String
Private m_Postfix As String
Private m_Increment As Long


Private Sub Class_Initialize()
    'Ç±Ç±Ç…èàóùÇèëÇ≠
    m_Integer = 0
    m_Increment = 0
End Sub

Private Sub Class_Terminate()
    'Ç±Ç±Ç…èàóùÇèëÇ≠
End Sub

Public Property Get Name() As Long
    Name = m_Integer + m_Increment
End Property

Public Property Let Name(Value As Long)
    m_Integer = Value
End Property

Public Sub ToSet(Value As Long)
    m_Integer = Value
End Sub

Public Sub SetValue(Value As Long)
    m_Integer = Value
End Sub

Public Property Get forJPfunc() As Long
    forJPfunc = m_Integer + m_Increment
End Property

Public Property Get ToVBA() As Long
    ToVBA = m_Integer + m_Increment
End Property

Public Sub Increment(Optional AddingValue As Long = 1)
    m_Increment = m_Increment + AddingValue
End Sub

Public Function ToString() As String
    ToString = m_Prefix & CStr(m_Integer + m_Increment) & m_Postfix
End Function

Public Sub SetPrefix(Optional Prefix As String = "")
    m_Prefix = Prefix
End Sub

Public Sub SetPostfix(Optional Postfix As String = "")
    m_Postfix = Postfix
End Sub






VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "通貨"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_通貨 As Currency

Public Property Get Name() As Currency
    Name = m_通貨
End Property

Public Property Let Name(値 As Currency)
    m_通貨 = 値
End Property

Public Property Get JP関数用() As Currency
    JP関数用 = m_通貨
End Property

Public Property Get VBAに渡す() As Currency
    VBAに渡す = m_通貨
End Property

Public Sub セット(値 As Currency)
    m_通貨 = 値
End Sub




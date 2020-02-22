VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Currency As Currency

Public Property Get Name() As Currency
    Name = m_Currency
End Property

Public Property Let Name(Value As Currency)
    m_Currency = Value
End Property

Public Property Get forJPfunc() As Currency
    forJPfunc = m_Currency
End Property

Public Property Get ToVBA() As Currency
    ToVBA = m_Currency
End Property

Public Sub ToSet(Value As Currency)
    m_Currency = Value
End Sub





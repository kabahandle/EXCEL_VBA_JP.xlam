VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "�ʉ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_�ʉ� As Currency

Public Property Get Name() As Currency
    Name = m_�ʉ�
End Property

Public Property Let Name(�l As Currency)
    m_�ʉ� = �l
End Property

Public Property Get JP�֐��p() As Currency
    JP�֐��p = m_�ʉ�
End Property

Public Property Get VBA�ɓn��() As Currency
    VBA�ɓn�� = m_�ʉ�
End Property

Public Sub �Z�b�g(�l As Currency)
    m_�ʉ� = �l
End Sub




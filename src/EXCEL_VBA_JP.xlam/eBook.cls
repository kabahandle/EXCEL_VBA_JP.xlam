VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Book As Workbook

Private Sub Class_Initialize()
    '�����ɏ���������
    Set m_Book = ActiveWorkbook
End Sub

Private Sub Class_Terminate()
    '�����ɏ���������
    Set m_Book = Nothing
End Sub

Public Property Set ToSet(book As Workbook)
    Set m_Book = book
End Property

Public Property Get GetNames() As Names
'    Dim wNames As New Names
    Set GetNames = m_Book.Names
'    Set ���O�z�� = w���O�z��
'    Set w���O�z�� = Nothing
End Property







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
    'ここに処理を書く
    Set m_Book = ActiveWorkbook
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_Book = Nothing
End Sub

Public Property Set ToSet(book As Workbook)
    Set m_Book = book
End Property

Public Property Get GetNames() As Names
'    Dim wNames As New Names
    Set GetNames = m_Book.Names
'    Set 名前配列 = w名前配列
'    Set w名前配列 = Nothing
End Property








VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ブック"
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

Public Property Set セット(book As Workbook)
    Set m_Book = book
End Property

Public Property Get 名前配列() As 名前配列
    Dim w名前配列 As New 名前配列
    Set w名前配列.セット = m_Book.Names
    Set 名前配列 = w名前配列
    Set w名前配列 = Nothing
End Property







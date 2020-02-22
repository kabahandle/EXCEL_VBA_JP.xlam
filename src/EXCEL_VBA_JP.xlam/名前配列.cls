VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "名前配列"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Names As Names

Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_Names = Nothing
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_Names = Nothing
End Sub

Public Property Set セット(名前配列 As Names)
    Set m_Names = 名前配列
End Property

Public Property Get 名前(Optional 順番 As Long = 1) As Name
    Set 名前 = m_Names(順番)
End Property






VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "罫線"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Border As Border

Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_Border = ActiveSheet.Range("A1").Offset(0, 0).Borders(xlEdgeTop)
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_Border = Nothing
End Sub


Public Property Set セット(罫線 As Border)
    Set m_Border = 罫線
End Property

Public Property Get 線種() As 罫線線種
    線種 = m_Border.LineStyle
End Property

Public Property Let 線種(線種値 As 罫線線種)
    m_Border.LineStyle = 線種値
End Property

Public Function 線種連鎖(線種値 As 罫線線種) As 罫線
    m_Border.LineStyle = 線種値
    Set 線種連鎖 = Me
End Function

Public Property Get 線の太さ() As 罫線の太さ
    線の太さ = m_Border.Weight
End Property

Public Property Let 線の太さ(線の太さ値 As 罫線の太さ)
    m_Border.Weight = 線の太さ値
End Property

Public Function 線の太さ連鎖(線の太さ値 As 罫線の太さ) As 罫線
    m_Border.Weight = 線の太さ値
    Set 線の太さ連鎖 = Me
End Function


'--ColorIndex
Public Property Let 色インデックス(ByVal インデックス値 As Long)
    m_Border.ColorIndex = インデックス値
End Property

Public Property Get 色インデックス() As Long
    色インデックス = m_Border.ColorIndex
End Property

Public Function 色インデックス連鎖(ByVal インデックス値 As Long) As 罫線
    m_Border.ColorIndex = インデックス値
    Set 色インデックス連鎖 = Me
End Function


Public Property Let 色インデックスパターン(ByVal インデックス値 As カラーインデックスパターン)
    m_Border.ColorIndex = インデックス値
End Property

Public Function 色インデックスパターン連鎖(ByVal インデックス値 As カラーインデックスパターン) As 罫線
    m_Border.ColorIndex = インデックス値
    Set 色インデックスパターン連鎖 = Me
End Function


'----Color
Public Sub 色設定(赤 As Long, 緑 As Long, 青 As Long)
    m_Border.Color = RGB(赤, 緑, 青)
End Sub
Public Property Let 色(色値 As Variant)
    m_Border.Color = 色値
End Property
Public Property Get 色() As Variant
    色 = m_Border.Color
End Property
Public Function 色連鎖(色値 As Variant) As 罫線
    m_Border.Color = 色値
    Set 色連鎖 = Me
End Function

Function ToVBA() As Border
    Set ToVBA = m_Border
End Function




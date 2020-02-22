VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "書式"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Interior As Interior

Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_Interior = ActiveCell.Interior
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_Interior = Nothing
End Sub

Public Property Set セット(書式 As Interior)
    Set m_Interior = 書式
End Property

'---- ColorIndex
Public Property Get 色番号() As Variant
    色番号 = m_Interior.ColorIndex
End Property

Public Property Let 色番号(色番号値 As Variant)
    m_Interior.ColorIndex = 色番号値
End Property

'---- Color
Public Property Get 色() As Variant
    色 = m_Interior.Color
End Property

Public Property Let 色(色値 As Variant)
    m_Interior.Color = 色値
End Property

Public Function 色連鎖(色値 As Variant) As 書式
    m_Interior.Color = 色値
    Set 色連鎖 = Me
End Function

'---- Pattern
Public Property Get 背景パターン() As セル背景色パターン
    背景パターン = m_Interior.Pattern
End Property

Public Property Let 背景パターン(背景パターン値 As セル背景色パターン)
    m_Interior.Pattern = 背景パターン値
End Property

Public Function 背景パターン連鎖(背景パターン値 As セル背景色パターン) As 書式
    m_Interior.Pattern = 背景パターン値
    Set 背景パターン連鎖 = Me
End Function

'---- PatternColorIndex
Public Property Get 背景パターン色パターン() As カラーインデックスパターン
    背景パターン色パターン = m_Interior.PatternColorIndex
End Property

Public Property Let 背景パターン色パターン(背景パターン色パターン値 As カラーインデックスパターン)
    m_Interior.Color = 色インデックス(背景パターン色パターン値)
End Property

Public Function 背景パターン色パターン連鎖(背景パターン色パターン値 As カラーインデックスパターン) As 書式
    m_Interior.Color = 色インデックス(背景パターン色パターン値)
    Set 背景パターン色パターン連鎖 = Me
End Function


'---- PatternColorIndex
Public Property Get 背景パターン色() As Variant
    背景パターン色 = m_Interior.PatternColor
End Property

Public Property Let 背景パターン色(背景パターン色値 As Variant)
    m_Interior.PatternColor = 背景パターン色値
End Property

Public Function 背景パターン色連鎖(背景パターン色値 As Variant) As 書式
    m_Interior.PatternColor = 背景パターン色値
    Set 背景パターン色連鎖 = Me
End Function


Private Function 色インデックス(idx As Variant) As Variant
    Select Case idx
    Case 1
        色インデックス = RGB(0, 0, 0)
    Case 2
        色インデックス = RGB(255, 255, 255)
    Case 3
        色インデックス = RGB(255, 0, 0)
    Case 4
        色インデックス = RGB(0, 255, 0)
    Case 5
        色インデックス = RGB(0, 0, 255)
    Case 6
        色インデックス = RGB(255, 255, 0)
    Case 7
        色インデックス = RGB(255, 0, 255)
    Case 8
        色インデックス = RGB(0, 255, 255)
    Case 9
        色インデックス = RGB(128, 0, 0)
    Case 10
        色インデックス = RGB(0, 128, 0)
    Case 11
        色インデックス = RGB(0, 0, 128)
    Case 12
        色インデックス = RGB(128, 128, 0)
    Case 13
        色インデックス = RGB(128, 0, 128)
    Case 14
        色インデックス = RGB(0, 128, 128)
    Case 15
        色インデックス = RGB(192, 192, 192)
    Case 16
        色インデックス = RGB(128, 128, 128)
    Case 17
        色インデックス = RGB(153, 153, 255)
    Case 18
        色インデックス = RGB(153, 51, 102)
    Case 19
        色インデックス = RGB(255, 255, 204)
    Case 20
        色インデックス = RGB(204, 255, 255)
    Case 21
        色インデックス = RGB(102, 0, 102)
    Case 22
        色インデックス = RGB(255, 128, 128)
    Case 23
        色インデックス = RGB(0, 102, 204)
    Case 24
        色インデックス = RGB(204, 204, 255)
    Case 25
        色インデックス = RGB(0, 0, 128)
    Case 26
        色インデックス = RGB(255, 0, 255)
    Case 27
        色インデックス = RGB(255, 255, 0)
    Case 28
        色インデックス = RGB(0, 255, 255)
    Case 29
        色インデックス = RGB(128, 0, 128)
    Case 30
        色インデックス = RGB(128, 0, 0)
    Case 31
        色インデックス = RGB(0, 128, 128)
    Case 32
        色インデックス = RGB(0, 0, 255)
    Case 33
        色インデックス = RGB(0, 204, 255)
    Case 34
        色インデックス = RGB(204, 255, 255)
    Case 35
        色インデックス = RGB(204, 255, 204)
    Case 36
        色インデックス = RGB(255, 255, 153)
    Case 37
        色インデックス = RGB(153, 204, 255)
    Case 38
        色インデックス = RGB(255, 153, 204)
    Case 39
        色インデックス = RGB(204, 153, 255)
    Case 40
        色インデックス = RGB(255, 204, 153)
    Case 41
        色インデックス = RGB(51, 102, 255)
    Case 42
        色インデックス = RGB(51, 204, 204)
    Case 43
        色インデックス = RGB(153, 204, 0)
    Case 44
        色インデックス = RGB(255, 204, 0)
    Case 45
        色インデックス = RGB(255, 153, 0)
    Case 46
        色インデックス = RGB(255, 102, 0)
    Case 47
        色インデックス = RGB(102, 102, 153)
    Case 48
        色インデックス = RGB(150, 150, 150)
    Case 49
        色インデックス = RGB(0, 51, 102)
    Case 50
        色インデックス = RGB(51, 153, 102)
    Case 51
        色インデックス = RGB(0, 51, 0)
    Case 52
        色インデックス = RGB(51, 51, 0)
    Case 53
        色インデックス = RGB(153, 51, 0)
    Case 54
        色インデックス = RGB(153, 51, 102)
    Case 55
        色インデックス = RGB(51, 51, 153)
    Case 56
        色インデックス = RGB(51, 51, 51)
    End Select

End Function



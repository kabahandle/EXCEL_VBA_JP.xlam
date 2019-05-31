VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eStyle"
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

Public Property Set ToSet(InteriorValue As Interior)
    Set m_Interior = InteriorValue
End Property

'---- ColorIndex
Public Property Get ColorIndex() As Variant
    ColorIndex = m_Interior.ColorIndex
End Property

Public Property Let ColorIndex(ColorIndexValue As Variant)
    m_Interior.ColorIndex = ColorIndexValue
End Property

'---- Color
Public Property Get Color() As Variant
    Color = m_Interior.Color
End Property

Public Property Let Color(ColorValue As Variant)
    m_Interior.Color = ColorValue
End Property

Public Function ColorChain(ColorValue As Variant) As eStyle
    m_Interior.Color = ColorValue
    Set ColorChain = Me
End Function

'---- Pattern
Public Property Get BackgroundPattern() As CellBackgroundPattern
    BackgroundPattern = m_Interior.Pattern
End Property

Public Property Let BackgroundPattern(BackgroundPatternValue As CellBackgroundPattern)
    m_Interior.Pattern = BackgroundPatternValue
End Property

Public Function BackgroundPatternChain(BackgroundPatternValue As CellBackgroundPattern) As eStyle
    m_Interior.Pattern = BackgroundPatternValue
    Set BackgroundPatternChain = Me
End Function

'---- PatternColorIndex
Public Property Get BackgroundColorPattern() As PatternOfColorIndex
    BackgroundColorPattern = m_Interior.PatternColorIndex
End Property

Public Property Let BackgroundColorPattern(BackgroundColorPatternValue As PatternOfColorIndex)
    m_Interior.Color = GetColorFromIndex(BackgroundColorPatternValue)
End Property

Public Function BackgroundColorPatternChain(BackgroundColorPatternValue As PatternOfColorIndex) As eStyle
    m_Interior.Color = GetColorFromIndex(BackgroundColorPatternValue)
    Set BackgroundColorPatternChain = Me
End Function


'---- PatternColorIndex
'Public Property Get 背景パターン色() As Variant
'    背景パターン色 = m_Interior.PatternColor
'End Property
'
'Public Property Let 背景パターン色(背景パターン色値 As Variant)
'    m_Interior.PatternColor = 背景パターン色値
'End Property
'
'Public Function 背景パターン色連鎖(背景パターン色値 As Variant) As 書式
'    m_Interior.PatternColor = 背景パターン色値
'    Set 背景パターン色連鎖 = Me
'End Function


Private Function GetColorFromIndex(idx As Variant) As Variant
    Select Case idx
    Case 1
        GetColorFromIndex = RGB(0, 0, 0)
    Case 2
        GetColorFromIndex = RGB(255, 255, 255)
    Case 3
        GetColorFromIndex = RGB(255, 0, 0)
    Case 4
        GetColorFromIndex = RGB(0, 255, 0)
    Case 5
        GetColorFromIndex = RGB(0, 0, 255)
    Case 6
        GetColorFromIndex = RGB(255, 255, 0)
    Case 7
        GetColorFromIndex = RGB(255, 0, 255)
    Case 8
        GetColorFromIndex = RGB(0, 255, 255)
    Case 9
        GetColorFromIndex = RGB(128, 0, 0)
    Case 10
        GetColorFromIndex = RGB(0, 128, 0)
    Case 11
        GetColorFromIndex = RGB(0, 0, 128)
    Case 12
        GetColorFromIndex = RGB(128, 128, 0)
    Case 13
        GetColorFromIndex = RGB(128, 0, 128)
    Case 14
        GetColorFromIndex = RGB(0, 128, 128)
    Case 15
        GetColorFromIndex = RGB(192, 192, 192)
    Case 16
        GetColorFromIndex = RGB(128, 128, 128)
    Case 17
        GetColorFromIndex = RGB(153, 153, 255)
    Case 18
        GetColorFromIndex = RGB(153, 51, 102)
    Case 19
        GetColorFromIndex = RGB(255, 255, 204)
    Case 20
        GetColorFromIndex = RGB(204, 255, 255)
    Case 21
        GetColorFromIndex = RGB(102, 0, 102)
    Case 22
        GetColorFromIndex = RGB(255, 128, 128)
    Case 23
        GetColorFromIndex = RGB(0, 102, 204)
    Case 24
        GetColorFromIndex = RGB(204, 204, 255)
    Case 25
        GetColorFromIndex = RGB(0, 0, 128)
    Case 26
        GetColorFromIndex = RGB(255, 0, 255)
    Case 27
        GetColorFromIndex = RGB(255, 255, 0)
    Case 28
        GetColorFromIndex = RGB(0, 255, 255)
    Case 29
        GetColorFromIndex = RGB(128, 0, 128)
    Case 30
        GetColorFromIndex = RGB(128, 0, 0)
    Case 31
        GetColorFromIndex = RGB(0, 128, 128)
    Case 32
        GetColorFromIndex = RGB(0, 0, 255)
    Case 33
        GetColorFromIndex = RGB(0, 204, 255)
    Case 34
        GetColorFromIndex = RGB(204, 255, 255)
    Case 35
        GetColorFromIndex = RGB(204, 255, 204)
    Case 36
        GetColorFromIndex = RGB(255, 255, 153)
    Case 37
        GetColorFromIndex = RGB(153, 204, 255)
    Case 38
        GetColorFromIndex = RGB(255, 153, 204)
    Case 39
        GetColorFromIndex = RGB(204, 153, 255)
    Case 40
        GetColorFromIndex = RGB(255, 204, 153)
    Case 41
        GetColorFromIndex = RGB(51, 102, 255)
    Case 42
        GetColorFromIndex = RGB(51, 204, 204)
    Case 43
        GetColorFromIndex = RGB(153, 204, 0)
    Case 44
        GetColorFromIndex = RGB(255, 204, 0)
    Case 45
        GetColorFromIndex = RGB(255, 153, 0)
    Case 46
        GetColorFromIndex = RGB(255, 102, 0)
    Case 47
        GetColorFromIndex = RGB(102, 102, 153)
    Case 48
        GetColorFromIndex = RGB(150, 150, 150)
    Case 49
        GetColorFromIndex = RGB(0, 51, 102)
    Case 50
        GetColorFromIndex = RGB(51, 153, 102)
    Case 51
        GetColorFromIndex = RGB(0, 51, 0)
    Case 52
        GetColorFromIndex = RGB(51, 51, 0)
    Case 53
        GetColorFromIndex = RGB(153, 51, 0)
    Case 54
        GetColorFromIndex = RGB(153, 51, 102)
    Case 55
        GetColorFromIndex = RGB(51, 51, 153)
    Case 56
        GetColorFromIndex = RGB(51, 51, 51)
    End Select

End Function




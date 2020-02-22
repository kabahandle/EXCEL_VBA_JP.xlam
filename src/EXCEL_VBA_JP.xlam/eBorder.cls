VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Border As Border

Private Sub Class_Initialize()
    'Ç±Ç±Ç…èàóùÇèëÇ≠
    Set m_Border = ActiveSheet.Range("A1").Offset(0, 0).Borders(xlEdgeTop)
End Sub

Private Sub Class_Terminate()
    'Ç±Ç±Ç…èàóùÇèëÇ≠
    Set m_Border = Nothing
End Sub


Public Property Set ToSet(BorderValue As Border)
    Set m_Border = BorderValue
End Property

Public Property Get LineStyle() As StyleOfBoderLine
    LineStyle = m_Border.LineStyle
End Property

Public Property Let LineStyle(LineStyleValue As StyleOfBoderLine)
    m_Border.LineStyle = LineStyleValue
End Property

Public Function LineStyleChain(LineStyleValue As StyleOfBoderLine) As eBorder
    m_Border.LineStyle = LineStyleValue
    Set LineStyleChain = Me
End Function

Public Property Get LineWeight() As BorderThickness
    LineWeight = m_Border.Weight
End Property

Public Property Let LineWeight(LineWeightValue As BorderThickness)
    m_Border.Weight = LineWeightValue
End Property

Public Function LineWeightChain(LineWeightValue As BorderThickness) As eBorder
    m_Border.Weight = LineWeightValue
    Set LineWeightChain = Me
End Function


'--ColorIndex
Public Property Let ColorIndex(ByVal ColorIndexValue As Long)
    m_Border.ColorIndex = ColorIndexValue
End Property

Public Property Get ColorIndex() As Long
    ColorIndex = m_Border.ColorIndex
End Property

Public Function ColorIndexChain(ByVal ColorIndexValue As Long) As eBorder
    m_Border.ColorIndex = ColorIndexValue
    Set ColorIndexChain = Me
End Function


Public Property Let ColorIndexPatern(ByVal ColorIndexPaternValue As PatternOfColorIndex)
    m_Border.ColorIndex = ColorIndexPaternValue
End Property

Public Function ColorIndexPaternChain(ByVal ColorIndexPaternValue As PatternOfColorIndex) As eBorder
    m_Border.ColorIndex = ColorIndexPaternValue
    Set ColorIndexPaternChain = Me
End Function


'----Color
Public Sub SetColor(Red As Long, Green As Long, Blue As Long)
    m_Border.Color = RGB(Red, Green, Blue)
End Sub
Public Property Let Color(ColorValue As Variant)
    m_Border.Color = ColorValue
End Property
Public Property Get Color() As Variant
    Color = m_Border.Color
End Property
Public Function ColorChain(ColorValue As Variant) As eBorder
    m_Border.Color = ColorValue
    Set ColorChain = Me
End Function

Function ToVBA() As Border
    Set ToVBA = m_Border
End Function



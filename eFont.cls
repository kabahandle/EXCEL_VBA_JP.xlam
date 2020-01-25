VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Font As Font

Private Sub Class_Initialize()
    'Ç±Ç±Ç…èàóùÇèëÇ≠
    Set m_Font = ActiveCell.Font
End Sub

Private Sub Class_Terminate()
    'Ç±Ç±Ç…èàóùÇèëÇ≠
    Set m_Font = Nothing
End Sub

Public Property Set ToSet(FontValue As Font)
    Set m_Font = FontValue
End Property


'--- Font.Name
Public Property Get FontName() As String
    FontName = m_Font.Name
End Property
'
Public Property Let FontName(ByVal FontNameValue As String)
    m_Font.Name = FontNameValue
End Property

Public Function FontNameChain(ByVal FontNameValue As String) As eFont
    m_Font.Name = FontNameValue
    Set FontNameChain = Me
End Function

'--- 20180825 14:35 ---

'--- Font.Size
Public Property Get Size() As Long
    Size = m_Font.Size
End Property
'
Public Property Let Size(ByVal SizeValue As Long)
    m_Font.Size = SizeValue
End Property

Public Function SizeChain(ByVal SizeValue As Long) As eFont
    m_Font.Size = SizeValue
    Set SizeChain = Me
End Function

'--- Bold
Public Property Let Bold(ByVal BoldValue As Boolean)
    m_Font.Bold = BoldValue
End Property
Public Property Get Bold() As Boolean
    Bold = m_Font.Bold
End Property
Public Function BoldChain(ByVal BoldValue As Boolean) As eFont
    m_Font.Bold = BoldValue
    Set BoldChain = Me
End Function

'--Italic
Public Property Let Italic(ByVal ItalicValue As Boolean)
    m_Font.Italic = ItalicValue
End Property
Public Property Get Italic() As Boolean
    Italic = m_Font.Italic
End Property
Public Function ItalicChain(ByVal ItalicValue As Boolean) As eFont
    m_Font.Italic = ItalicValue
    Set ItalicChain = Me
End Function


'---Underline
Public Property Let Underline(ByVal UnderlineValue As Boolean)
    m_Font.Underline = UnderlineValue
End Property
Public Property Get Underline() As Boolean
    Underline = m_Font.Underline
End Property
Public Function UnderlineChain(ByVal UnderlineValue As Boolean) As eFont
    m_Font.Underline = UnderlineValue
    Set UnderlineChain = Me
End Function


Public Property Let UnderlinePattern(ByVal UnderlinePatternValue As StyleOfUnderLinePattern)
    m_Font.Underline = UnderlinePatternValue
End Property

Public Function UnderlinePatternChain(ByVal UnderlinePatternValue As StyleOfUnderLinePattern) As eFont
    m_Font.Underline = UnderlinePatternValue
    Set UnderlinePatternChain = Me
End Function


'--StrikethroughÅië≈Çøè¡ÇµÅj
Public Property Let Strikethrough(ByVal StrikethroughValue As Boolean)
    m_Font.Strikethrough = StrikethroughValue
End Property
Public Property Get Strikethrough() As Boolean
    Strikethrough = m_Font.Strikethrough
End Property
Public Function StrikethroughChaub(ByVal StrikethroughValue As Boolean) As eFont
    m_Font.Strikethrough = StrikethroughValue
    Set StrikethroughChaub = Me
End Function



'--ColorIndex
Public Property Let ColorIndex(ByVal ColorIndexValue As Long)
    m_Font.ColorIndex = ColorIndexValue
End Property
Public Property Get ColorIndex() As Long
    ColorIndex = m_Font.ColorIndex
End Property
Public Function ColorIndexChain(ByVal ColorIndexValue As Long) As eFont
    m_Font.ColorIndex = ColorIndexValue
    Set ColorIndexChain = Me
End Function



Public Property Let ColorIndexPattern(ByVal ColorIndexValue As PatternOfColorIndex)
    m_Font.ColorIndex = ColorIndexValue
End Property
Public Function ColorIndexPatternChain(ByVal ColorIndexValue As PatternOfColorIndex) As eFont
    m_Font.ColorIndex = ColorIndexValue
    Set ColorIndexPatternChain = Me
End Function


'--SuperScript
Public Property Let SuperScript(ByVal SuperScriptValue As Boolean)
    m_Font.SuperScript = SuperScriptValue
End Property
Public Property Get SuperScript() As Boolean
    SuperScript = m_Font.SuperScript
End Property
Public Function SuperScriptChain(ByVal SuperScriptValue As Boolean) As eFont
    m_Font.SuperScript = SuperScriptValue
    Set SuperScriptChain = Me
End Function

'--SubScript
Public Property Let SubScript(ByVal SubScriptValue As Boolean)
    m_Font.SubScript = SubScriptValue
End Property
Public Property Get SubScript() As Boolean
    SubScript = m_Font.SubScript
End Property
Public Function SubScriptChain(ByVal SubScriptValue As Boolean) As eFont
    m_Font.SubScript = SubScriptValue
    Set SubScriptChain = Me
End Function


'----Color
Public Sub SetColor(Red As Long, Green As Long, Blue As Long)
    m_Font.Color = RGB(Red, Green, Blue)
End Sub
Public Function SetColorChain(Red As Long, Green As Long, Blue As Long) As eFont
    m_Font.Color = RGB(Red, Green, Blue)
    Set SetColorChain = Me
End Function

Public Property Let Color(ColorValue As Variant)
    m_Font.Color = ColorValue
End Property
Public Property Get Color() As Variant
    Color = m_Font.Color
End Property
Public Function ColorChain(ColorValue As Variant) As eFont
    m_Font.Color = ColorValue
    Set ColorChain = Me
End Function

Public Function ToVBA() As Font
    Set ToVBA = m_Font
End Function





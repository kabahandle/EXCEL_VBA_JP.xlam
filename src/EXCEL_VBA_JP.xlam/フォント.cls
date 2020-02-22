VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "フォント"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Font As Font

Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_Font = ActiveCell.Font
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_Font = Nothing
End Sub

Public Property Set セット(フォント As Font)
    Set m_Font = フォント
End Property


'--- Font.Name
Public Property Get フォント名() As String
    フォント名 = m_Font.Name
End Property
'
Public Property Let フォント名(ByVal フォント名値 As String)
    m_Font.Name = フォント名値
End Property

Public Function フォント名連鎖(ByVal フォント名値 As String) As フォント
    m_Font.Name = フォント名値
    Set フォント名連鎖 = Me
End Function

'--- Font.Size
Public Property Get サイズ() As Long
    サイズ = m_Font.Size
End Property
'
Public Property Let サイズ(ByVal サイズ値 As Long)
    m_Font.Size = サイズ値
End Property

Public Function サイズ連鎖(ByVal サイズ値 As Long) As フォント
    m_Font.Size = サイズ値
    Set サイズ連鎖 = Me
End Function

'--- Bold
Public Property Let 太字(ByVal 太字値 As Boolean)
    m_Font.Bold = 太字値
End Property
Public Property Get 太字() As Boolean
    太字 = m_Font.Bold
End Property
Public Function 太字連鎖(ByVal 太字値 As Boolean) As フォント
    m_Font.Bold = 太字値
    Set 太字連鎖 = Me
End Function

'--Italic
Public Property Let イタリック(ByVal イタリック値 As Boolean)
    m_Font.Italic = イタリック値
End Property
Public Property Get イタリック() As Boolean
    イタリック = m_Font.Italic
End Property
Public Function イタリック連鎖(ByVal イタリック値 As Boolean) As フォント
    m_Font.Italic = イタリック値
    Set イタリック連鎖 = Me
End Function


'---Underline
Public Property Let アンダーライン(ByVal アンダーライン値 As Boolean)
    m_Font.Underline = アンダーライン値
End Property
Public Property Get アンダーライン() As Boolean
    アンダーライン = m_Font.Underline
End Property
Public Function アンダーライン連鎖(ByVal アンダーライン値 As Boolean) As フォント
    m_Font.Underline = アンダーライン値
    Set アンダーライン連鎖 = Me
End Function


Public Property Let アンダーラインパターン(ByVal アンダーラインパターン値 As アンダーラインパターン種類)
    m_Font.Underline = アンダーラインパターン値
End Property

Public Function アンダーラインパターン連鎖(ByVal アンダーラインパターン値 As アンダーラインパターン種類) As フォント
    m_Font.Underline = アンダーラインパターン値
    Set アンダーラインパターン連鎖 = Me
End Function


'--Strikethrough（打ち消し）
Public Property Let 打ち消し線(ByVal 打ち消し線値 As Boolean)
    m_Font.Strikethrough = 打ち消し線値
End Property
Public Property Get 打ち消し線() As Boolean
    打ち消し線 = m_Font.Strikethrough
End Property
Public Function 打ち消し線連鎖(ByVal 打ち消し線値 As Boolean) As フォント
    m_Font.Strikethrough = 打ち消し線値
    Set 打ち消し線連鎖 = Me
End Function



'--ColorIndex
Public Property Let 色インデックス(ByVal インデックス値 As Long)
    m_Font.ColorIndex = インデックス値
End Property
Public Property Get 色インデックス() As Long
    色インデックス = m_Font.ColorIndex
End Property
Public Function 色インデックス連鎖(ByVal インデックス値 As Long) As フォント
    m_Font.ColorIndex = インデックス値
    Set 色インデックス連鎖 = Me
End Function



Public Property Let 色インデックスパターン(ByVal インデックス値 As カラーインデックスパターン)
    m_Font.ColorIndex = インデックス値
End Property
Public Function 色インデックスパターン連鎖(ByVal インデックス値 As カラーインデックスパターン) As フォント
    m_Font.ColorIndex = インデックス値
    Set 色インデックスパターン連鎖 = Me
End Function


'--SuperScript
Public Property Let 上付き文字(ByVal 上付き文字値 As Boolean)
    m_Font.SuperScript = 上付き文字値
End Property
Public Property Get 上付き文字() As Boolean
    上付き文字 = m_Font.SuperScript
End Property
Public Function 上付き文字連鎖(ByVal 上付き文字値 As Boolean) As フォント
    m_Font.SuperScript = 上付き文字値
    Set 上付き文字連鎖 = Me
End Function

'--SubScript
Public Property Let 下付き文字(ByVal 下付き文字値 As Boolean)
    m_Font.SubScript = 下付き文字値
End Property
Public Property Get 下付き文字() As Boolean
    下付き文字 = m_Font.SubScript
End Property
Public Function 下付き文字連鎖(ByVal 下付き文字値 As Boolean) As フォント
    m_Font.SubScript = 下付き文字値
    Set 下付き文字連鎖 = Me
End Function


'----Color
Public Sub 色設定(赤 As Long, 青 As Long, 黄 As Long)
    m_Font.Color = RGB(赤, 青, 黄)
End Sub
Public Function 色設定連鎖(赤 As Long, 青 As Long, 黄 As Long) As フォント
    m_Font.Color = RGB(赤, 青, 黄)
    Set 色設定連鎖 = Me
End Function

Public Property Let 色(色値 As Variant)
    m_Font.Color = 色値
End Property
Public Property Get 色() As Variant
    色 = m_Font.Color
End Property
Public Function 色連鎖(色値 As Variant) As フォント
    m_Font.Color = 色値
    Set 色連鎖 = Me
End Function


'--- to vba
Public Function ToVBA() As Font
    Set ToVBA = m_Font
End Function


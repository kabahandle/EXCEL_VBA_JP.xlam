VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "変数取得"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Property Get 文字列(Optional 初期値 As String) As 文字列
    Dim w_文字列 As New 文字列
    w_文字列.Name = 初期値
    Set 文字列 = w_文字列
    Set w_文字列 = Nothing
End Property

Public Property Get 整数(Optional 初期値 As Long = 0, Optional 接頭辞 As String = "", Optional 接尾辞 As String = "") As 整数
    Dim w_整数 As New 整数
    
    w_整数.セット 初期値
    w_整数.接頭辞を設定 接頭辞
    w_整数.接尾辞を設定 接尾辞
    
    Set 整数 = w_整数
    Set w_整数 = Nothing
End Property

Public Property Get 通貨(Optional 初期値 As Currency) As 通貨
    Dim w_通貨 As New 通貨
    
    w_通貨.セット 初期値
    
    Set 通貨 = w_通貨
    Set w_通貨 = Nothing
End Property

Public Property Get 配列(Optional 初期要素数 As Long = 0) As 配列
    Dim w_配列 As New 配列
    
    Call w_配列.要素数指定(初期要素数)
    
    Set 配列 = w_配列
    Set w_配列 = Nothing
    
End Property

Public Property Get スタック() As スタック
    Dim w_スタック As New スタック
    
    Set スタック = w_スタック
    Set w_スタック = Nothing
    
End Property


Public Property Get ハッシュ() As ハッシュ
    Dim w_ハッシュ As New ハッシュ
    
    Set ハッシュ = w_ハッシュ
    Set w_ハッシュ = Nothing
    
End Property





VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "hngハンガリアンセルv1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const DebugMode = False

Private m_Range As Range
Private m_セル名 As String
Private m_シート順番 As Long
Private m_Prev選択範囲 As Range

Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_Range = ActiveSheet.Range("A1")
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_Range = Nothing
End Sub

'--- 日本語関数用
'Public Property Set Range(レンジ As Range)
'    Set m_Range = レンジ
'End Property
'
'Public Property Get Range() As Range
'    Set Range = m_Range
'End Property

Public Property Set JP関数用(レンジ As Range)
    Set m_Range = レンジ
End Property

Public Property Get JP関数用() As Range
    Set JP関数用 = m_Range
End Property

Public Property Get F_関数用() As Range
    Set F_関数用 = m_Range
End Property

'--- end of 日本語間枢要

Public Property Set setセット(rng As Range)
    Set m_Range = rng
End Property

Public Property Let ceruセル名(strgセル名 As String)
    m_セル名 = strgセル名
End Property

Public Property Get ceruセル名() As String
    ceruセル名 = m_セル名
End Property

Public Property Let shtシート順番(ByVal lngシート順番 As Long)
    m_シート順番 = lngシート順番
End Property

Public Property Get ceruセル範囲() As Range
    Set ceruセル範囲 = m_Range
End Property

Public Function soutai相対位置(Optional ByVal 行 As Long = 0, Optional ByVal 列 As Long = 0) As hngハンガリアンセルv1
    Dim wセル As New hngハンガリアンセルv1
    Set wセル.setセット = m_Range.Offset(行, 列)
    Set soutai相対位置 = wセル
    Set wセル = Nothing
End Function


Public Function sentaku選択() As hngハンガリアンセルv1
    Dim w_範囲 As New hngハンガリアンセルv1
    Dim w_range As Range
    
    m_Range.Select
    
    With Selection
        Set w_range = ActiveSheet.Range(.Cells.Address)
        Set w_範囲.setセット = w_range
    End With
    
    Set sentaku選択 = w_範囲
    
    Set w_範囲 = Nothing
    Set w_range = Nothing
    
End Function

Public Sub sounyu挿入(Optional シフト方向値 As シフト方向 = 右方向にシフト)
    m_Range.Insert シフト方向値
End Sub

Public Sub sakujyo削除(Optional シフト方向値 As シフト方向 = 左方向にシフト)
    m_Range.Delete シフト方向値
End Sub

Public Sub syoukyo消去()
    m_Range.Clear
End Sub

Public Sub syoukyo消去内容のみ()
    m_Range.ClearContents
End Sub

Public Sub syoukyo消去書式のみ()
    m_Range.ClearFormats
End Sub

Public Sub syoukyo消去コメントのみ()
    m_Range.ClearComments
End Sub

Public Sub umeru埋める(埋めるセル範囲 As hngハンガリアンセルv1, セルの埋め方 As 埋め方)
    m_Range.AutoFill Destination:=埋めるセル範囲.ceruセル範囲, Type:=セルの埋め方
End Sub

Public Sub kirityori切り取り(Optional 移動先 As hngハンガリアンセルv1 = Nothing)
    If Not 移動先 Is Nothing Then
        m_Range.Cut 移動先.ceruセル範囲
    Else
        m_Range.Cut
    End If
End Sub

Public Sub copyコピー(Optional コピー先 As hngハンガリアンセルv1 = Nothing)
    If Not コピー先 Is Nothing Then
        m_Range.Copy コピー先.ceruセル範囲
    Else
        m_Range.Copy
    End If
End Sub

Public Sub haritsuke形式をして指定して貼り付け(Optional 貼り付け方法値 As 貼り付け方法 = すべて貼り付け, Optional ByVal 行列を変更して貼り付け As Boolean = False)
#If DebugMode = False Then
    On Error Resume Next
#End If
    m_Range.PasteSpecial 貼り付け方法値, , , 行列を変更して貼り付け
    On Error GoTo 0
End Sub

Public Sub haritsuke貼り付け()
    m_Range.Select
    ActiveSheet.Paste
End Sub

Public Sub ketsugo結合()
    m_Range.Merge
End Sub

Public Function syutan終端(Optional 方向 As 終端方向 = 終端方向.下端) As hngハンガリアンセルv1
    Select Case 方向
    Case 終端方向.上橋
        Set m_Range = m_Range.End(xlUp)
    Case 終端方向.下端
        Set m_Range = m_Range.End(xlDown)
    Case 終端方向.右端
        Set m_Range = m_Range.End(xlToRight)
    Case 終端方向.左端
        Set m_Range = m_Range.End(xlToLeft)
    Case Else
        Set m_Range = m_Range.End(xlDown)
    End Select
    
    Set syutan終端 = Me
End Function

Public Function ceru特定セル抽出(Optional 選択方法 As セル選択方法 = セル選択方法.空のセル, Optional 選択条件値 As セル選択条件値 = セル選択条件値.文字) As hngハンガリアンセルv1
    On Error GoTo ERR1
    
    Select Case 選択方法
    Case セル選択方法.数式あり
        Set m_Range = m_Range.SpecialCells(選択方法, 選択条件値)
    Case セル選択方法.定数あり
        Set m_Range = m_Range.SpecialCells(選択方法, 選択条件値)
    Case Else
        Set m_Range = m_Range.SpecialCells(選択方法)
    End Select
    
    Set ceru特定セル抽出 = Me
    
    Exit Function
    
ERR1:
    Set ceru特定セル抽出 = Me
End Function



'---- Value2
Public Property Get atai値() As Variant
    atai値 = m_Range.Value2
End Property

'---- Value2
Public Property Get addressアドレス(Optional 列を絶対指定 As 絶対か相対かアドレス指定 = 相対アドレス, Optional 行を絶対指定 As 絶対か相対かアドレス指定 = 相対アドレス) As Variant
    addressアドレス = m_Range.Address(列を絶対指定, 行を絶対指定)
End Property

Public Property Let atai値(値データ As Variant)
    m_Range.Value2 = 値データ
End Property

'---- Formula
Public Property Get sushiki数式() As Variant
    sushiki数式 = m_Range.Formula
End Property

Public Property Let sushiki数式(数式値 As Variant)
    m_Range.Formula = 数式値
End Property

''---- Rows
'Public Property Get 行全体(行番号 As Variant) As hngハンガリアンセルv1
'    Set m_Range = m_Range.Rows(行番号)
'    Set 行全体 = Me
'End Property
'
''---- Cols
'Public Property Get 列全体(列番号 As Variant) As hngハンガリアンセルv1
'    Set m_Range = m_Range.Columns(列番号)
'    Set 列全体 = Me
'End Property


'---- Interior
Public Property Get syoshiki書式() As 書式
    Dim w書式 As New 書式
    Set w書式.セット = m_Range.Interior
    Set syoshiki書式 = w書式
    Set w書式 = Nothing
End Property

'---- Count
Public Property Get kosu個数() As Long
    kosu個数 = m_Range.Count
End Property

'---- EntireRow
Public Property Get gyou行全体() As hngハンガリアンセルv1
    Dim wセル As New hngハンガリアンセルv1
    Set wセル.setセット = m_Range.EntireRow
    Set gyou行全体 = wセル
    Set wセル = Nothing
End Property

'---- EntireCols
Public Property Get retsu列全体() As hngハンガリアンセルv1
    Dim wセル As New hngハンガリアンセルv1
    Set wセル.setセット = m_Range.EntireColumn
    Set retsu列全体 = wセル
    Set wセル = Nothing
End Property

'---- Name
Public Property Get name名前() As String
    name名前 = m_Range.Name
End Property

Public Property Let name名前(名前値 As String)
    m_Range.Name = 名前値
End Property

Public Property Let komentoコメント(コメント As String)
    If TypeName(m_Range.Comment) = "Comment" Then
        m_Range.ClearComments
    End If
    m_Range.AddComment コメント
End Property

Public Property Get komentoコメント() As String
    komentoコメント = m_Range.Comment
End Property

' セルを非表示にするとエラー
'Public Property Let 非表示(ByVal 非表示 As Boolean)
'    m_Range.Hidden = 非表示
'End Property

Public Property Get takasa高さ() As Variant
    takasa高さ = m_Range.Height
End Property

'セルの高さと幅は変更することができない。（行高さ・列幅の変更で代用すること）
'Public Property Let 高さ(ByVal 高さ値 As Variant)
'    m_Range.Offset(0, 0).Height = 高さ値
'End Property

Public Property Get haba幅() As Variant
    haba幅 = m_Range.Width
End Property

'セルの高さと幅は変更することができない。（行高さ・列幅の変更で代用すること）
'Public Property Let 幅(ByVal 幅値 As Variant)
'    m_Range.Offset(0, 0).Width = 幅値
'End Property


'--- NumberFormatLocal
Public Property Get hyouji表示形式() As Variant
    hyouji表示形式 = m_Range.NumberFormatLocal
End Property

Public Property Let hyouji表示形式(ByVal 表示形式値 As Variant)
    m_Range.NumberFormatLocal = 表示形式値
End Property

' 有効にするとLetがエラー。戻り値の型が違うため。
'Public Property Get 表示形式パターン() As Variant
'    表示形式 = m_Range.NumberFormatLocal
'End Property

Public Property Let hyoji表示形式パターン(ByVal 表示形式パターン値 As 表示形式パターン定数)
    Select Case 表示形式パターン値
    Case 通貨
        m_Range.NumberFormatLocal = "\#,##0;[赤]\#,##0"
    Case 小数点以下1桁
        m_Range.NumberFormatLocal = "0.0"
    Case 小数点以下2桁
        m_Range.NumberFormatLocal = "0.00"
    Case 第4桁まで0埋め
        m_Range.NumberFormatLocal = "0000"
    Case 第8桁まで0埋め
        m_Range.NumberFormatLocal = "00000000"
    Case 西暦
        m_Range.NumberFormatLocal = "yyyy/mm/dd"
    Case 西暦曜日付き
        m_Range.NumberFormatLocal = "yyyy/mm/dd(aaa)"
    Case 和暦
        m_Range.NumberFormatLocal = "gggr""年""m""月""d""日"""
    Case 和暦曜日付き
        m_Range.NumberFormatLocal = "gggr""年""m""月""d""日""(aaa)"
    Case 日時
        m_Range.NumberFormatLocal = "hh:mm"
    Case 日時分
        m_Range.NumberFormatLocal = "hh:mm:ss"
    Case 日時AMPM
        m_Range.NumberFormatLocal = "hh:mm AM/PM"
    End Select
End Property

'--- HorizentalAlignment
Public Property Get yoko横位置() As セル横位置
    yoko横位置 = m_Range.HorizontalAlignment
End Property
'
Public Property Let yoko横位置(ByVal 横位置値 As セル横位置)
    m_Range.HorizontalAlignment = 横位置値
End Property


'--- VertucalAlignment
Public Property Get tate縦位置() As セル縦位置
    tate縦位置 = m_Range.VerticalAlignment
End Property
'
Public Property Let tate縦位置(ByVal 縦位置値 As セル縦位置)
    m_Range.VerticalAlignment = 縦位置値
End Property


'--- Orientation
Public Property Get kaku角度() As Long
    kaku角度 = m_Range.Orientation
End Property
'
Public Property Let kaku角度(ByVal 角度値 As Long)
    If 角度値 > 90 Then
        角度値 = 90
    ElseIf 角度値 < -90 Then
        角度値 = -90
    End If
    m_Range.Orientation = 角度値
End Property

'Public Property Get 角度パターン() As セル角度
'    角度パターン = m_Range.Orientation
'End Property
'
Public Property Let kaku角度パターン(ByVal 角度値 As セル角度)
    m_Range.Orientation = 角度値
End Property


'---- Interior
Public Property Get fontoフォント() As フォント
    Dim w_フォント As New フォント
    Set w_フォント.セット = m_Range.Font
    Set fontoフォント = w_フォント
    Set w_フォント = Nothing
End Property
Public Property Set fontoフォント(フォント As フォント)
    Set m_Range.Font = フォント.ToVBA
End Property


'---- borders
Public Property Get keisen罫線(Optional 罫線位置値 As 罫線位置 = 左端の罫線) As 罫線
    Dim w_罫線 As New 罫線
    Set w_罫線.セット = m_Range.Borders(罫線位置値)
    Set keisen罫線 = w_罫線
    Set w_罫線 = Nothing
End Property

Public Sub keisen罫線囲み(線種値 As 罫線線種, 罫線の太さ値 As 罫線の太さ, 色インデックス値 As カラーインデックスパターン)
    Call m_Range.BorderAround(線種値, 罫線の太さ値, 色インデックス値)
End Sub

'Public Property Let 罫線(罫線位置値 As 罫線位置)
'    m_Border.Borders = 線の太さ値
'End Property


'---- Photonic
Public Property Get hurikanaふりがな() As String
    hurikanaふりがな = Application.GetPhonetic(m_Range.Offset(0, 0).Value2)
End Property

Public Sub hurikanaふりがな表示()
    m_Range.SetPhonetic
    m_Range.Phonetics.Visible = True
End Sub

Public Property Get kuhaku空白() As Boolean
    kuhaku空白 = (Trim(m_Range.Offset(0, 0)).Value2 = "")
End Property

Public Property Get kuhaku空白でない() As Boolean
    kuhaku空白でない = Trim(m_Range.Offset(0, 0).Value2) <> ""
End Property

Public Function bubun部分一致(部分一致文字列 As 文字列) As Boolean
    bubun部分一致 = (InStr(1, Trim(m_Range.Offset(0, 0).Value2), 部分一致文字列.VBAに渡す) > 0)
End Function

Public Function bubun部分一致でない(部分一致文字列 As 文字列) As Boolean
    bubun部分一致でない = (InStr(1, Trim(m_Range.Offset(0, 0).Value2), 部分一致文字列.VBAに渡す) <= 0)
End Function

Public Function bubun部分一致String(部分一致文字列 As String) As Boolean
    bubun部分一致String = (InStr(1, Trim(m_Range.Offset(0, 0).Value2), 部分一致文字列) > 0)
End Function

Public Property Get seisu整数() As 整数
    Dim ret As Long
    Dim w_整数 As New 整数
    
    On Error GoTo ERR数値001
    
    ret = 0
    ret = CLng(m_Range.Offset(0, 0).Value2)
    
ERR数値001:
    On Error GoTo 0
    
    w_整数.Name = ret
    
    Set seisu整数 = w_整数
    Set w_整数 = Nothing
End Property

Public Property Get seisu整数Long() As Long
    Dim ret As Long
    On Error GoTo ERR数値002
    
    ret = 0
    ret = CLng(m_Range.Offset(0, 0).Value2)
    
ERR数値002:
    On Error GoTo 0
    
    seisu整数Long = ret
    
End Property

Public Property Get tsuka通貨() As 通貨
    Dim ret As Currency
    Dim w_通貨 As New 通貨
    
    On Error GoTo ERR通貨001
    
    ret = 0
    ret = CCur(m_Range.Offset(0, 0).Value2)
    
ERR通貨001:
    On Error GoTo 0
    
    w_通貨 = ret
    
    Set tsuka通貨 = w_通貨
    Set w_通貨 = Nothing
End Property

Public Property Get tsuka通貨Curr() As Currency
    Dim ret As Currency
    On Error GoTo ERR通貨002
    
    ret = 0
    ret = CCur(m_Range.Offset(0, 0).Value2)
    
ERR通貨002:
    On Error GoTo 0
    
    tsuka通貨Curr = ret
End Property

Public Property Get sentaku選択範囲(Optional 範囲パターン指定 As 選択範囲パターン指定 = 選択範囲パターン指定.選択範囲パターン指定なし, Optional ステップ数 As Long = -1, Optional y行目 As Long = -1, Optional x列目 As Long = -1, Optional オフセットy As Long = 0, Optional オフセットx As Long = 0) As hngハンガリアンセルv1
    Dim w_range As Range
    Dim w_範囲 As New hngハンガリアンセルv1
    Dim startCol As Long
    Dim endCol As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim c As Long
    Dim r As Long
    
    With Selection
        Set m_Prev選択範囲 = ActiveSheet.Range(.Cells.Address)
    End With
    startCol = Selection(1).Column + オフセットx
    endCol = Selection(Selection.Count).Column + オフセットx
    startRow = Selection(1).Row + オフセットy
    endRow = Selection(Selection.Count).Row + オフセットy
    y行目 = y行目 + オフセットy
    x列目 = x列目 + オフセットx
    
    
    Set w_range = ActiveSheet.Range("A1").Offset(0, 0)
    
    If 範囲パターン指定 = 選択範囲パターン指定.選択範囲パターン偶数行 Then
        Selection.Offset(0, 0).Select
        Selection.Resize(1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For r = startRow To endRow Step 2
            Set w_range = Union(w_range, Range(Cells(r, startCol), Cells(r, endCol)))
        Next
        
            
        Set w_範囲.setセット = w_range
        Set sentaku選択範囲 = w_範囲
    
    ElseIf 範囲パターン指定 = 選択範囲パターン指定.選択範囲パターン奇数行 Then
        Selection.Offset(1, 0).Select
        Selection.Resize(1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For r = startRow + 1 To endRow Step 2
            Set w_range = Union(w_range, Range(Cells(r, startCol), Cells(r, endCol)))
        Next
        
            
        Set w_範囲.setセット = w_range
        Set sentaku選択範囲 = w_範囲
    
    ElseIf 範囲パターン指定 = 選択範囲パターン指定.選択範囲パターン行ステップ And ステップ数 >= 1 Then
        Selection.Offset(0, 0).Select
        Selection.Resize(1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For r = startRow To endRow Step ステップ数
            Set w_range = Union(w_range, Range(Cells(r, startCol), Cells(r, endCol)))
        Next
        
            
        Set w_範囲.setセット = w_range
        Set sentaku選択範囲 = w_範囲
    
    ElseIf 範囲パターン指定 = 選択範囲パターン指定.選択範囲パターン偶数列 Then
        Selection.Offset(0, 0).Select
        Selection.Resize(, 1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For c = startCol To endCol Step 2
            Set w_range = Union(w_range, Range(Cells(startRow, c), Cells(endRow, c)))
        Next
        
            
        Set w_範囲.setセット = w_range
        Set sentaku選択範囲 = w_範囲
    
    ElseIf 範囲パターン指定 = 選択範囲パターン指定.選択範囲パターン奇数列 Then
        Selection.Offset(0, 1).Select
        Selection.Resize(, 1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For c = startCol + 1 To endCol Step 2
            Set w_range = Union(w_range, Range(Cells(startRow, c), Cells(endRow, c)))
        Next
        
            
        Set w_範囲.setセット = w_range
        Set sentaku選択範囲 = w_範囲
    
    ElseIf 範囲パターン指定 = 選択範囲パターン指定.選択範囲パターン列ステップ And ステップ数 >= 1 Then
        Selection.Offset(0, 0).Select
        Selection.Resize(, 1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For c = startCol To endCol Step ステップ数
            Set w_range = Union(w_range, Range(Cells(startRow, c), Cells(endRow, c)))
        Next
        
            
        Set w_範囲.setセット = w_range
        Set sentaku選択範囲 = w_範囲
    
    Else
        
        If y行目 <= 0 And x列目 <= 0 Then
        
        ElseIf (y行目 >= 1 Or オフセットy = 0) And x列目 <= 0 Then
            Selection.Offset(y行目 - 1, 0).Select
            Selection.Resize(1).Select
        ElseIf y行目 <= 0 And (x列目 >= 1 Or オフセットx = 0) Then
            Selection.Offset(0, x列目 - 1).Select
            Selection.Resize(, 1).Select
        ElseIf y行目 >= 1 And x列目 >= 1 Then
            Selection.Offset(y行目 - 1, x列目 - 1).Select
            Selection.Resize(1, 1).Select
        End If
        
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
            Set w_範囲.setセット = w_range
        End With
        Set sentaku選択範囲 = w_範囲
    
    End If
    
    
    m_Prev選択範囲.Select
    
    Set w_range = Nothing
    Set w_範囲 = Nothing
End Property


Public Sub burokkuセルごとにサブルーチン呼び出し(サブルーチン名 As String, _
                Optional 引数1 As Variant, _
                Optional 引数2 As Variant, _
                Optional 引数3 As Variant, _
                Optional 引数4 As Variant, _
                Optional 引数5 As Variant, _
                Optional 引数6 As Variant, _
                Optional 引数7 As Variant, _
                Optional 引数8 As Variant, _
                Optional 引数9 As Variant)
    Dim w_rng As Range
    Dim w_セル As New hngハンガリアンセルv1
    
    For Each w_rng In m_Range
        Set w_セル.setセット = w_rng
        Application.Run ActiveWorkbook.Name & "!" & サブルーチン名, w_セル, 引数1, 引数2, 引数3, 引数4, 引数5, 引数6, 引数7, 引数8, 引数9
    Next
    
    Set w_セル = Nothing
    Set w_rng = Nothing
    
End Sub

Public Sub burokkuセルごとにサブルーチン呼び出し2(ワークブック名付きサブルーチン名 As String, _
                Optional 引数1 As Variant, _
                Optional 引数2 As Variant, _
                Optional 引数3 As Variant, _
                Optional 引数4 As Variant, _
                Optional 引数5 As Variant, _
                Optional 引数6 As Variant, _
                Optional 引数7 As Variant, _
                Optional 引数8 As Variant, _
                Optional 引数9 As Variant)
    Dim w_rng As Range
    Dim w_セル As New hngハンガリアンセルv1
    
    For Each w_rng In m_Range
        Set w_セル.setセット = w_rng
        Application.Run ワークブック名付きサブルーチン名, w_セル, 引数1, 引数2, 引数3, 引数4, 引数5, 引数6, 引数7, 引数8, 引数9
    Next
    
    Set w_セル = Nothing
    Set w_rng = Nothing
    
End Sub


Public Function offsetオフセット選択(Optional オフセットy As Long = 0, Optional オフセットx As Long = 0) As hngハンガリアンセルv1
    'm_Range.Offset(オフセットy, オフセットx).Select
    
    Set m_Range = Range(m_Range.Cells, m_Range.Cells.Offset(オフセットy, オフセットx))
    m_Range.Select
    
    Set offsetオフセット選択 = Me
End Function


Public Function kanni簡易セルへ() As 簡易セル
    Dim w簡易セル As New 簡易セル
    Set w簡易セル.セット = m_Range
    Set kanni簡易セルへ = w簡易セル
    Set w簡易セル = Nothing
End Function

Public Function seruセルへ() As セル
    Dim wセル As New セル
    Set wセル.セット = m_Range
    Set seruセルへ = wセル
    Set wセル = Nothing
End Function

Public Sub syugo集合積(w_セル As hngハンガリアンセルv1)
    Dim w_range As Range
    Set w_range = range集合積(m_Range, w_セル.F_関数用)
    Set m_Range = w_range
    Set w_range = Nothing
End Sub

Public Sub syugo集合和(w_セル As hngハンガリアンセルv1)
    Dim w_range As Range
    Set w_range = range集合和(m_Range, w_セル.F_関数用)
    Set m_Range = w_range
    Set w_range = Nothing
End Sub

Public Sub syugo集合差(w_セル As hngハンガリアンセルv1)
    Dim w_range As Range
    Set w_range = range集合差(m_Range, w_セル.F_関数用)
    Set m_Range = w_range
    Set w_range = Nothing
End Sub

Public Sub syugo集合反転()
    Set m_Range = range集合反転(m_Range)
End Sub



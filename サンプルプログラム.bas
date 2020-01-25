Attribute VB_Name = "サンプルプログラム"
Option Explicit


Private Sub サンプル04()
    g_道具.セル("A5").値 = "け"
End Sub

Private Sub サンプル05()
    Dim x As New 整数
    x = 1
    g_道具.セル("A6").値 = x
End Sub

Private Sub サンプル06()
    Dim strg As New 文字列
    strg = "abcdfe"
    g_道具.セル("A7").値 = strg
End Sub

Private Sub サンプル07()
    Set g_道具.シート.セット = ActiveSheet
    With g_道具.シート
        g_道具.セル("B1", .シート順番).値 = 1
    End With
End Sub

Private Sub テスト01()
    Dim x As New 整数
    x = 1
    x = x + 2
    g_道具.セル("A6").値 = x
    g_道具.現在のセル.値 = x + 1
    g_道具.現在の表.選択
    g_道具.現在のセル.終端(下端).選択
    g_道具.現在のセル.特定セル抽出(数式あり).選択
    g_道具.シート(1).行(2).値 = 10
    g_道具.シート(1).列(11).選択
    g_道具.現在のシート.行(1).名前 = "行だよ"
    g_道具.現在のシート.行(1).選択
    g_道具.現在のセル.値 = g_道具.現在のシート.行(1).個数
    g_道具.現在のセル.書式.色番号 = 17
    g_道具.セル.相対位置(0, 1).値 = g_道具.名前配列.名前

    Dim wセル As セル
    Set wセル = g_道具.セル("G3")
    wセル.値 = "1月"
    wセル.相対位置(1, 0).値 = "2月"
    wセル.相対位置(2, 0).値 = "3月"
    Set wセル = g_道具.セル("G3:G5")
    wセル.埋める g_道具.セル("G3:H5"), 月単位
    g_道具.セル行列指定(2, 2).値 = "987654"
    
End Sub

Private Sub クリアテスト()
    g_道具.セル("A4").挿入 下方向にシフト
    g_道具.セル("A5").削除 上方向にシフト
    g_道具.セル("A1").消去
    g_道具.セル("A2").消去コメントのみ
    g_道具.セル("A3").消去書式のみ
    g_道具.セル("A4").消去内容のみ
End Sub

Private Sub コピペテスト()
    g_道具.セル("A1").コピー
    g_道具.セル("A2").貼り付け
    
    g_道具.セル("A3").切り取り
    g_道具.セル("A4").形式をして指定して貼り付け 値のみ貼り付け
    
    g_道具.セル("A5").コピー g_道具.セル("B2")
    
    g_道具.セル("A6").切り取り g_道具.セル("B3")
    
    g_道具.セル("A7").切り取り
    g_道具.セル("A8").貼り付け
    
End Sub

Private Sub シフトテスト()
    g_道具.セル("A5").削除 上方向にシフト
    g_道具.セル("A1").消去
    g_道具.セル("A2").消去コメントのみ
    g_道具.セル("A3").消去書式のみ
    g_道具.セル("A4").消去内容のみ
End Sub

Private Sub 結合テスト()
    g_道具.セル("C10:C12").結合
End Sub


Private Sub 文字列をセットテスト()
    Dim strg As 文字列
    Set strg = h_変数取得.文字列("abcdfe")
    Set strg = h_変数取得.文字列(strg.VBAに渡す & "1234")
    g_道具.セル("A7").値 = strg.VBAに渡す
    Set strg = Nothing
End Sub

Private Sub コメントテスト()
    g_道具.セル("A1").コメント = "aaaa" & vbCr & vbLf & "bbb"
End Sub

Private Sub 表示非表示テスト()
    g_道具.シート.行非表示 2, True
    
    'エラー
    'g_道具.セル("D4").非表示 = True
    
    g_道具.シート.列非表示 3, True
    
End Sub

Private Sub 行高さ列幅設定テスト()
    Dim tmp As Variant
    g_道具.シート.行高さ設定 3, 50
    g_道具.シート.列幅設定 3, 50
    
    tmp = g_道具.シート.行高さ取得(3)
    g_道具.セル("A1").値 = tmp
    tmp = g_道具.シート.列幅取得(3)
    g_道具.セル("B1").値 = tmp
    
    'セルの高さと幅は変更することができない。（行高さ・列幅の変更で代用すること）
'    g_道具.セル("A4").高さ = 15
'    g_道具.セル("A5").幅 = 20
    g_道具.セル("A6").値 = g_道具.セル("A4").高さ
    g_道具.セル("A7").値 = g_道具.セル("A5").幅
    
    g_道具.シート.行高さ自動設定 1
    g_道具.シート.列幅自動設定 1
    
End Sub

Private Sub 表示形式テスト()
    g_道具.セル("B1").値 = "0.15"
    g_道具.セル("B1").表示形式パターン = 小数点以下1桁
    
    g_道具.セル("B2").値 = "0.215"
    g_道具.セル("B2").表示形式パターン = 小数点以下2桁
    
    g_道具.セル("B3").値 = "2015/8/1"
    g_道具.セル("B3").表示形式パターン = 西暦
    
    g_道具.セル("B4").値 = "2015/8/1"
    g_道具.セル("B4").表示形式パターン = 西暦曜日付き
    
    g_道具.セル("B5").値 = "12"
    g_道具.セル("B5").表示形式パターン = 第4桁まで0埋め
    
    g_道具.セル("B6").値 = "12345"
    g_道具.セル("B6").表示形式パターン = 第8桁まで0埋め
    
    g_道具.セル("B7").値 = "123456789"
    g_道具.セル("B7").表示形式パターン = 通貨
    
    g_道具.セル("B8").値 = "2015/8/2 12:01:02"
    g_道具.セル("B8").表示形式パターン = 日時
    
    g_道具.セル("B9").値 = "2015/8/2 12:01:02"
    g_道具.セル("B9").表示形式パターン = 日時AMPM
    
    g_道具.セル("B10").値 = "2015/8/2 12:01:02"
    g_道具.セル("B10").表示形式パターン = 日時分
    
    g_道具.セル("B11").値 = "2015/8/3 12:01:02"
    g_道具.セル("B11").表示形式パターン = 和暦
    
    g_道具.セル("B12").値 = "2015/8/3 12:01:02"
    g_道具.セル("B12").表示形式パターン = 和暦曜日付き
    
End Sub

Private Sub セル横位置テスト()
    g_道具.セル("C1").値 = "あああ"
    g_道具.セル("C1").横位置 = 横右詰め
    
    g_道具.セル("C2").値 = "あああ"
    g_道具.セル("C2").横位置 = 横均等割り付け
    
    g_道具.セル("C3").値 = "あああ"
    g_道具.セル("C3").横位置 = 横繰り返し
    
    g_道具.セル("C4").値 = "あああ"
    g_道具.セル("C4").横位置 = 横左詰め
    
    g_道具.セル("C5").値 = "あああ"
    g_道具.セル("C5").横位置 = 横選択範囲内で中央
    
    g_道具.セル("C6").値 = "あああ"
    g_道具.セル("C6").横位置 = 横中央揃え
    
    g_道具.セル("C7").値 = "あああ"
    g_道具.セル("C7").横位置 = 横標準
    
    g_道具.セル("C8").値 = "あああ"
    g_道具.セル("C8").横位置 = 横両端揃え
    
    g_道具.表示 "C1=" & g_道具.セル("C1").横位置

End Sub

Private Sub セル縦位置テスト()
    g_道具.シート.行高さ設定 1, 30
    g_道具.セル("D1").値 = "あああ"
    g_道具.セル("D1").縦位置 = 縦下詰め

    g_道具.シート.行高さ設定 2, 30
    g_道具.セル("D2").値 = "あああ"
    g_道具.セル("D2").縦位置 = 縦均等割り付け

    g_道具.シート.行高さ設定 3, 30
    g_道具.セル("D3").値 = "あああ"
    g_道具.セル("D3").縦位置 = 縦上詰め

    g_道具.シート.行高さ設定 4, 30
    g_道具.セル("D4").値 = "あああ"
    g_道具.セル("D4").縦位置 = 縦中央揃え

    g_道具.シート.行高さ設定 5, 30
    g_道具.セル("D5").値 = "あああ"
    g_道具.セル("D5").縦位置 = 縦両端揃え


End Sub

Private Sub セル角度テスト()
    g_道具.セル("E1").値 = "あああ"
    g_道具.セル("E1").角度パターン = 角度30度

    g_道具.セル("E2").値 = "あああ"
    g_道具.セル("E2").角度パターン = 角度45度

    g_道具.セル("E3").値 = "あああ"
    g_道具.セル("E3").角度パターン = 角度60度

    g_道具.セル("E4").値 = "あああ"
    g_道具.セル("E4").角度パターン = 角度90度

    g_道具.セル("E5").値 = "あああ"
    g_道具.セル("E5").角度パターン = 角度マイナス30度

    g_道具.セル("E6").値 = "あああ"
    g_道具.セル("E6").角度パターン = 角度マイナス45度

    g_道具.セル("E7").値 = "あああ"
    g_道具.セル("E7").角度パターン = 角度マイナス60度

    g_道具.セル("E8").値 = "あああ"
    g_道具.セル("E8").角度パターン = 角度マイナス90度

    g_道具.セル("E9").値 = "あああ"
    g_道具.セル("E9").角度パターン = 角度縦方向

    g_道具.セル("E10").値 = "あああ"
    g_道具.セル("E10").角度 = 30


End Sub

Private Sub フォントテスト()
    g_道具.セル("F1").値 = "aaa"
    g_道具.セル("F1").フォント.フォント名 = フォント名Arial
    g_道具.セル("F1").フォント.サイズ = 8
    
    g_道具.セル("F2").値 = "aaa"
    g_道具.セル("F2").フォント.フォント名 = フォント名ArialBlack
    g_道具.セル("F2").フォント.サイズ = 10
    
    g_道具.セル("F3").値 = "aaa"
    g_道具.セル("F3").フォント.フォント名 = フォント名MSPゴシック
    g_道具.セル("F3").フォント.サイズ = 12
    
    g_道具.セル("F4").値 = "aaa"
    g_道具.セル("F4").フォント.フォント名 = フォント名MSP明朝
    g_道具.セル("F4").フォント.サイズ = 14
    
    g_道具.セル("F5").値 = "aaa"
    g_道具.セル("F5").フォント.フォント名 = フォント名MSゴシック
    g_道具.セル("F5").フォント.サイズ = 16
    
    g_道具.セル("F6").値 = "aaa"
    g_道具.セル("F6").フォント.フォント名 = フォント名MS明朝
    g_道具.セル("F6").フォント.サイズ = 18
    
    g_道具.セル("F7").値 = "aaa"
    g_道具.セル("F7").フォント.フォント名 = フォント名メイリオ
    g_道具.セル("F7").フォント.サイズ = 20
    
End Sub

Private Sub フォント形式テスト()
    g_道具.セル("G1").値 = "bbb"
    g_道具.セル("G1").フォント.太字 = True
    g_道具.セル("H1").値 = g_道具.セル("G1").フォント.太字
    
    g_道具.セル("G2").値 = "bbb"
    g_道具.セル("G2").フォント.イタリック = True
    g_道具.セル("H2").値 = g_道具.セル("G2").フォント.イタリック
    
    g_道具.セル("G3").値 = "bbb"
    g_道具.セル("G3").フォント.アンダーライン = True
    g_道具.セル("H3").値 = g_道具.セル("G3").フォント.アンダーライン
    
    g_道具.セル("G4").値 = "bbb"
    g_道具.セル("G4").フォント.アンダーラインパターン = 下線
    
    g_道具.セル("G5").値 = "bbb"
    g_道具.セル("G5").フォント.アンダーラインパターン = 下線なし
    
    g_道具.セル("G6").値 = "bbb"
    g_道具.セル("G6").フォント.アンダーラインパターン = 下線会計
    
    g_道具.セル("G7").値 = "bbb"
    g_道具.セル("G7").フォント.アンダーラインパターン = 二重下線
    
    g_道具.セル("G8").値 = "bbb"
    g_道具.セル("G8").フォント.アンダーラインパターン = 二重下線会計
    g_道具.セル("G8").フォント.色インデックスパターン = インデックス赤
    
    g_道具.セル("G9").値 = "bbb"
    g_道具.セル("G9").フォント.打ち消し線 = True
    g_道具.セル("H9").値 = g_道具.セル("G9").フォント.打ち消し線

End Sub

Private Sub テスト()
    g_道具.セル("A2").表示形式パターン = 和暦曜日付き
End Sub

Private Sub フォントカラーインデックス()
    g_道具.セル("A1").値 = "bbb"
    g_道具.セル("A1").フォント.色インデックスパターン = インデックスオレンジ
    
    g_道具.セル("A2").値 = "bbb"
    g_道具.セル("A2").フォント.色インデックス = 10
    
    g_道具.セル("A3").値 = g_道具.セル("A2").フォント.色インデックス
    
    g_道具.セル("A4").値 = "bbb"
    g_道具.セル("A4").フォント.色インデックスパターン = 色を自動的に設定
        
End Sub

Private Sub テスト上付き下付き文字()
    g_道具.セル("A1").値 = "bbb"
    g_道具.セル("A1").フォント.上付き文字 = True
    
    g_道具.セル("A2").値 = "bbb"
    g_道具.セル("A2").フォント.下付き文字 = True
    
    g_道具.セル("B1").値 = g_道具.セル("A1").フォント.上付き文字
    g_道具.セル("B2").値 = g_道具.セル("A2").フォント.下付き文字
    
End Sub

Private Sub ボーダーテスト()
'    Dim rng As Range
'    Set rng = Worksheets(1).Range.Offset(0, 0)
'    Call rng.Borders(xlDiagonalDown)
'    Set rng = Nothing

    g_道具.セル("A1:C1").罫線(下端の罫線).色インデックスパターン = インデックス黄土色
    g_道具.セル("A1:C1").罫線(下端の罫線).線の太さ = 極細
    g_道具.セル("A1:C1").罫線(下端の罫線).線種 = 一点鎖線
    
    g_道具.セル("A2:C2").罫線(右端の罫線).色インデックスパターン = インデックス紫
    g_道具.セル("A2:C2").罫線(右端の罫線).線の太さ = 細い
    g_道具.セル("A2:C2").罫線(下端の罫線).線種 = 細実線
    
    g_道具.セル("A3:C3").罫線(下端の罫線).色インデックスパターン = インデックス黄土色
    g_道具.セル("A3:C3").罫線(下端の罫線).線の太さ = 太い
    g_道具.セル("A3:C3").罫線(下端の罫線).線種 = 斜め破線
    
    g_道具.セル("A4:C4").罫線(下端の罫線).色設定 256, 256, 0
    g_道具.セル("A4:C4").罫線(下端の罫線).線の太さ = 太い
    g_道具.セル("A4:C4").罫線(下端の罫線).線種 = 斜め破線
    
    g_道具.セル("A6:C6").罫線囲み 二重線, 中, インデックスオレンジ
    
   
End Sub

Private Sub 背景色テスト()
    g_道具.セル("A1").書式.背景パターン = チェック
    g_道具.セル("A1").書式.背景パターン色パターン = インデックス黄土色
    g_道具.セル("B1").値 = g_道具.セル("A1").書式.背景パターン
    
    g_道具.セル("A2").書式.背景パターン = 右下がり斜め細線
    g_道具.セル("A2").書式.背景パターン色パターン = インデックス紫
    g_道具.セル("B2").値 = g_道具.セル("A2").書式.背景パターン色パターン
    
    
    g_道具.セル("A3").書式.背景パターン = チェック
    g_道具.セル("A3").書式.背景パターン色 = RGB(128, 0, 128)
    g_道具.セル("B3").値 = g_道具.セル("A3").書式.背景パターン色
    
End Sub

Private Sub フリガナテスト()
    g_道具.セル("A1").値 = "今日は"
    g_道具.セル("A1").ふりがな表示
    g_道具.セル("B1").値 = g_道具.セル("A1").ふりがな
End Sub

Private Sub セル行列指定()
    g_道具.セル行列指定(10, 1).書式.背景パターン = 右下がり斜め線
    g_道具.セル行列指定(10, 1).書式.背景パターン色パターン = インデックス黄色
    g_道具.セル行列指定(10, 2).値 = g_道具.セル("A1").書式.背景パターン
End Sub

Private Sub 文字列始まってる01()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから始まってる")
    Set strg文字2 = h_変数取得.文字列("これから")
    
    g_道具.表示 strg文字1.始まっている(strg文字2)
    
End Sub

Private Sub 文字列始まってる02()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから始まってる")
    Set strg文字2 = h_変数取得.文字列("これから始まってるよ")
    
    g_道具.表示 strg文字1.始まっている(strg文字2)
    
End Sub

Private Sub 文字列始まってる03()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから始まってる")
    Set strg文字2 = h_変数取得.文字列("これからよ")
    
    g_道具.表示 strg文字1.始まっている(strg文字2)
    
End Sub


Private Sub String始まってる01()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから始まってる")
    Set strg文字2 = h_変数取得.文字列("これから")
    
    g_道具.表示 strg文字1.Stringで始まっている(strg文字2.VBAに渡す)
    
End Sub

Private Sub String始まってる02()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから始まってる")
    Set strg文字2 = h_変数取得.文字列("これから始まってるよ")
    
    g_道具.表示 strg文字1.Stringで始まっている(strg文字2.VBAに渡す)
    
End Sub

Private Sub String始まってる03()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから始まってる")
    Set strg文字2 = h_変数取得.文字列("これからよ")
    
    g_道具.表示 strg文字1.Stringで始まっている(strg文字2.VBAに渡す)
    
    'セルの文字列を取得できる
    strg文字1 = g_道具.セル("C1").値
    
End Sub

Private Sub セルが空白かどうか01()
    Dim strg文字1 As 文字列
    
    Set strg文字1 = h_変数取得.文字列(g_道具.セル("C1").値)

    g_道具.表示 strg文字1.空白か

    Set strg文字1 = h_変数取得.文字列(g_道具.セル("C1000").値)

    g_道具.表示 strg文字1.空白か

    

End Sub

Private Sub セルが空白()
    Dim strg文字1 As 文字列
    
    Set strg文字1 = h_変数取得.文字列(g_道具.セル("C1").値)

    g_道具.表示 g_道具.セルが空白(1, 1)

    g_道具.表示 g_道具.セルが空白(1, 1000)
    
    g_道具.表示 g_道具.セルが空白(1, 1000, 2)
End Sub

Private Sub セルが空白でない()
    Dim strg文字1 As 文字列
    
    Set strg文字1 = h_変数取得.文字列(g_道具.セル("C1").値)

    g_道具.表示 g_道具.セルが空白でない(1, 1)

    g_道具.表示 g_道具.セルが空白でない(1, 1000)
    
    g_道具.表示 g_道具.セルが空白でない(1, 1000, 2)
End Sub

Private Sub 部分一致01()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("あああこれから始まってる")
    Set strg文字2 = h_変数取得.文字列("これから")
    
    g_道具.表示 strg文字1.部分一致(strg文字2)
    
    
End Sub

Private Sub 部分一致02()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("あああこれから始まってる")
    Set strg文字2 = h_変数取得.文字列("これから始まってるよ")
    
    g_道具.表示 strg文字1.部分一致(strg文字2)
    
End Sub

Private Sub 部分一致03()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから始まってる")
    Set strg文字2 = h_変数取得.文字列("これから")
    
    g_道具.表示 strg文字1.部分一致(strg文字2)
    
End Sub

Private Sub String部分一致01()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("あああこれから始まってる")
    Set strg文字2 = h_変数取得.文字列("これから")
    
    g_道具.表示 strg文字1.String部分一致(strg文字2.VBAに渡す)
    
    
End Sub

Private Sub String部分一致02()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("あああこれから始まってる")
    Set strg文字2 = h_変数取得.文字列("これから始まってるよ")
    
    g_道具.表示 strg文字1.String部分一致(strg文字2.VBAに渡す)
    
End Sub

Private Sub String部分一致03()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから始まってる")
    Set strg文字2 = h_変数取得.文字列("これから")
    
    g_道具.表示 strg文字1.String部分一致(strg文字2.VBAに渡す)
    
End Sub


Private Sub 文字列終わっている01()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから終わってる")
    Set strg文字2 = h_変数取得.文字列("終わってる")
    
    g_道具.表示 strg文字1.終わっている(strg文字2)
    
End Sub

Private Sub 文字列終わっている02()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから終わってる")
    Set strg文字2 = h_変数取得.文字列("あああ終わってる")
    
    g_道具.表示 strg文字1.終わっている(strg文字2)
    
End Sub

Private Sub 文字列終わっている03()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから終わってる")
    Set strg文字2 = h_変数取得.文字列("終わってるよ")
    
    g_道具.表示 strg文字1.終わっている(strg文字2)
    
End Sub


Private Sub String終わっている01()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから終わってる")
    Set strg文字2 = h_変数取得.文字列("終わってる")
    
    g_道具.表示 strg文字1.Stringで終わっている(strg文字2.VBAに渡す)
    
End Sub

Private Sub String終わっている02()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから終わってる")
    Set strg文字2 = h_変数取得.文字列("あああ終わってる")
    
    g_道具.表示 strg文字1.Stringで終わっている(strg文字2.VBAに渡す)
    
End Sub

Private Sub String終わっている03()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("これから終わってる")
    Set strg文字2 = h_変数取得.文字列("終わってるよ")
    
    g_道具.表示 strg文字1.Stringで終わっている(strg文字2.VBAに渡す)
    
    'セルの文字列を取得できる
    strg文字1 = g_道具.セル("C1").値
    
End Sub


Private Sub 長さ01()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("１２３")
    Set strg文字2 = h_変数取得.文字列("１２３４５")
    
    g_道具.表示 strg文字1.長さ
    g_道具.表示 strg文字2.長さ
    
End Sub

Private Sub 右側()
    Dim strg文字1 As 文字列
    'Dim strg文字2 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("１２３４５６７")
    'Set strg文字2 = h_変数取得.文字列"１２３４５")
    
    g_道具.表示 strg文字1.右側(4)
    
End Sub


Private Sub 置換()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    Dim strg文字3 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("１２３４５６７")
    Set strg文字2 = h_変数取得.文字列("１２３４５")
    Set strg文字3 = h_変数取得.文字列("ａｂｃｄｅｆ")
    
    g_道具.表示 strg文字1.置換(strg文字2, strg文字3)
    
End Sub

Private Sub String置換()
    Dim strg文字1 As 文字列
    Dim strg文字2 As 文字列
    Dim strg文字3 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("１２３４５６７")
    Set strg文字2 = h_変数取得.文字列("１２３４５")
    Set strg文字3 = h_変数取得.文字列("ａｂｃｄｅｆ")
    
    g_道具.表示 strg文字1.Sring置換(strg文字2.VBAに渡す, strg文字3.VBAに渡す)
    
End Sub



Private Sub セルが部分一致001()
    Dim strg文字1 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("aa")

    g_道具.表示 g_道具.セルが部分一致(strg文字1, 1, 5)

    g_道具.表示 g_道具.セルが部分一致(strg文字1, 1, 1000)
    
    g_道具.表示 g_道具.セルが部分一致(strg文字1, 1, 5, 2)

    g_道具.表示 g_道具.セルが部分一致(strg文字1, 1, 1000, 2)
    
End Sub

Private Sub セルが部分一致002()
    Dim strg文字1 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("aa")

    g_道具.表示 g_道具.セル行列指定(1, 5).部分一致(strg文字1)

    g_道具.表示 g_道具.セル行列指定(1, 1000).部分一致(strg文字1)
    
End Sub

Private Sub セルが部分一致でない001()
    Dim strg文字1 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("aa")

    g_道具.表示 g_道具.セルが部分一致でない(strg文字1, 1, 5)

    g_道具.表示 g_道具.セルが部分一致でない(strg文字1, 1, 1000)
    
    g_道具.表示 g_道具.セルが部分一致でない(strg文字1, 1, 5, 2)

    g_道具.表示 g_道具.セルが部分一致でない(strg文字1, 1, 1000, 2)
    
End Sub

Private Sub セルが部分一致でない002()
    Dim strg文字1 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("aa")

    g_道具.表示 g_道具.セル行列指定(1, 5).部分一致でない(strg文字1)

    g_道具.表示 g_道具.セル行列指定(1, 1000).部分一致でない(strg文字1)
    
End Sub

Private Sub セルが部分一致String()
    Dim strg文字1 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("aa")

    g_道具.表示 g_道具.セルが部分一致String(strg文字1.VBAに渡す, 1, 5)

    g_道具.表示 g_道具.セルが部分一致String(strg文字1.VBAに渡す, 1, 1000)
    
    g_道具.表示 g_道具.セルが部分一致String(strg文字1.VBAに渡す, 1, 5, 2)

    g_道具.表示 g_道具.セルが部分一致String(strg文字1.VBAに渡す, 1, 1000, 2)
    
End Sub

Private Sub セルが部分一致でないString()
    Dim strg文字1 As 文字列
    
    Set strg文字1 = h_変数取得.文字列("aa")

    g_道具.表示 g_道具.セルが部分一致でないString(strg文字1.VBAに渡す, 1, 5)

    g_道具.表示 g_道具.セルが部分一致でないString(strg文字1.VBAに渡す, 1, 1000)
    
    g_道具.表示 g_道具.セルが部分一致でないString(strg文字1.VBAに渡す, 1, 5, 2)

    g_道具.表示 g_道具.セルが部分一致でないString(strg文字1.VBAに渡す, 1, 1000, 2)
    
End Sub

Private Sub 整数取得001()
    g_道具.表示 g_道具.セル行列指定(1, 2).整数
    g_道具.表示 g_道具.セル行列指定(1, 2).整数Long
End Sub

Private Sub 通貨取得001()
    ' VBA の バグ？　実行時エラーが出る↓実装は整数クラスと同じなのに・・・
    'g_道具.表示 g_道具.セル行列指定(1, 2).通貨
    
    g_道具.表示 g_道具.セル行列指定(1, 2).通貨Curr
    
    'g_道具.表示 g_道具.セル行列指定(1, 2).
End Sub

Private Sub 変数取得_文字列01()
    Dim strg As 文字列
    Set strg = h_変数取得.文字列("あああ")
'    strg = "あああ"
    g_道具.表示 strg.VBAに渡す
End Sub

Private Sub 変数取得_文字列02()
    Dim strg As 文字列
    Set strg = h_変数取得.文字列("いいい")
    g_道具.表示 strg.VBAに渡す
End Sub

Private Sub 変数取得_整数01()
    Dim i As 整数
    Set i = h_変数取得.整数(1)
'    i = 1
    g_道具.表示 i.VBAに渡す
End Sub

Private Sub 変数取得_整数02()
    Dim i As 整数
    Set i = h_変数取得.整数(20)
    g_道具.表示 i.VBAに渡す
End Sub

'Private Sub 変数取得_通貨01()
'    Dim c As 通貨
'    Set c = h_変数取得.通貨
'    ' NG
'    ' VBA の バグ？　実行時エラーが出る↓実装は整数クラスと同じなのに・・・
'    c = CCur(2)
'    g_道具.表示 c
'End Sub
'
'Private Sub 変数取得_通貨02()
'    Dim c As 通貨
'    Set c = h_変数取得.通貨(20.2)
'    ' NG
'    ' VBA の バグ？　実行時エラーが出る↓実装は整数クラスと同じなのに・・・
'    g_道具.表示 c
'End Sub

Private Sub アドレス()
    g_道具.表示 g_道具.セル行列指定(1, 1).アドレス
    
    g_道具.表示 g_道具.セル行列指定(1, 1).アドレス(相対アドレス, 絶対アドレス)
    
    g_道具.表示 g_道具.セル行列指定(1, 1).アドレス(絶対アドレス, 相対アドレス)
    
    g_道具.表示 g_道具.セル行列指定(1, 1).アドレス(絶対アドレス, 絶対アドレス)
End Sub
'
'Private Sub 簡易セル()
'
'End Sub

Private Sub f_関数テスト()
    g_道具.表示 F_コサイン度(60)
    g_道具.表示 F_アークコサイン度(F_コサイン度(30))
    g_道具.表示 F_おおもじ("abdcEfg")
    g_道具.表示 F_あいだもじれつ("ABC-0001-DDD", 5, 4)
    g_道具.表示 F_ぜったいち(-123)
    g_道具.表示 F_日付化(F_すうかげつご(F_いま(), 3))
    g_道具.表示 F_十進数から二進数(65535)
    g_道具.表示 F_余り(13, 5)
    'g_道具.表示 f_
End Sub

Private Sub functionテスト()
    g_道具.表示 g_道具.セル("A1").相対位置(0, 0).値
    g_道具.セル("A1").高さ = 1
End Sub

Public Sub 選択範囲の背景色変更()
    g_道具.セル.選択範囲.書式.背景パターン = 塗りつぶし
    g_道具.セル.選択範囲.書式.色 = RGB(100, 100, 100)
End Sub

Public Sub 数値入力ボックス()
    Dim i As Long
    
    i = io_入出力.数値入力ボックス
    g_道具.表示 i & "でした"
    
    i = io_入出力.数値入力ボックス("数字くれ！")
    g_道具.表示 i & "でした"
    
    
End Sub

Public Sub 文字入力ボックス()
    Dim str As String
    
    str = io_入出力.文字入力ボックス
    g_道具.表示 """" & str & """" & "でした"
    
    str = io_入出力.文字入力ボックス("文字くれ！")
    g_道具.表示 """" & str & """" & "でした"
    
End Sub

Public Sub テキストファイル書き込み()
    io_入出力.ファイル書込用開く "test.txt", 現在のフォルダ
    io_入出力.ファイル1行出力 "1 aaa"
    io_入出力.ファイル1行出力 "2 bbb"
    io_入出力.ファイル1行出力 "3 ccc"
    io_入出力.ファイル1行出力 "4 ddd"
    io_入出力.ファイル1行出力 "5 eee"
    io_入出力.ファイル閉じる
    
End Sub

Public Sub テキストファイル読み込み()
    io_入出力.ファイル読込用開く "test1.txt", 現在のフォルダ
    g_道具.表示 io_入出力.ファイル1行読込
    io_入出力.ファイル閉じる
End Sub

Public Sub 選択範囲に増分入力()
    Dim w_lng As New 整数
    Dim a As Long
    
    w_lng.接尾辞を設定
    
    a = io_入出力.数値入力ボックス
    w_lng.値を設定 a
    w_lng.増やす
    w_lng.増やす
    w_lng.増やす
    g_道具.表示 w_lng.文字列化
    
End Sub

Public Sub 選択範囲の背景色変更_行目指定()
    g_道具.セル.選択範囲(, , 2).書式.背景パターン = 塗りつぶし
    g_道具.セル.選択範囲(, , , 2).書式.色 = RGB(100, 100, 100)
End Sub

Public Sub 選択範囲の背景色変更_列目指定()
    g_道具.セル.選択範囲(, , 0, 3).書式.背景パターン = 塗りつぶし
    g_道具.セル.選択範囲(, , 0, 3).書式.色 = RGB(100, 100, 100)
End Sub

Public Sub 選択範囲の背景色変更03()
    g_道具.セル.選択範囲(選択範囲パターン偶数行).書式.背景パターン = 塗りつぶし
    g_道具.セル.選択範囲(選択範囲パターン偶数行).書式.色 = F_色インデックスからRGB色へ変換(インデックスオレンジ)
End Sub

Public Sub 選択範囲の背景色変更04()
    g_道具.セル.選択範囲(選択範囲パターン奇数行).書式.背景パターン = 塗りつぶし
    g_道具.セル.選択範囲(選択範囲パターン奇数行).書式.色 = RGB(100, 100, 100)
End Sub

Public Sub 選択範囲の背景色変更05()
    g_道具.セル.選択範囲(選択範囲パターン偶数列).書式.背景パターン = 塗りつぶし
    g_道具.セル.選択範囲(選択範囲パターン偶数列).書式.色 = RGB(100, 100, 100)
End Sub

Public Sub 選択範囲の背景色変更06()
    g_道具.セル.選択範囲(選択範囲パターン奇数列).書式.背景パターン = 塗りつぶし
    g_道具.セル.選択範囲(選択範囲パターン奇数列).書式.色 = RGB(100, 100, 100)
End Sub

Public Sub 選択範囲の背景色変更07()
    g_道具.セル.選択範囲(選択範囲パターン行ステップ, 3).書式.背景パターン = 塗りつぶし
    g_道具.セル.選択範囲(選択範囲パターン行ステップ, 3).書式.色 = F_色インデックスからRGB色へ変換(インデックス黄色)
End Sub

Public Sub 選択範囲の背景色変更08()
    g_道具.セル.選択範囲(選択範囲パターン列ステップ, 3).書式.背景パターン = 塗りつぶし
    g_道具.セル.選択範囲(選択範囲パターン列ステップ, 3).書式.色 = F_色インデックスからRGB色へ変換(インデックスオレンジ)
End Sub

Public Sub 関数呼び出しテストコード()
    Dim w_整数 As New 整数
    
    w_整数.接頭辞を設定 "第"
    w_整数.接尾辞を設定 "月"
    
    g_道具.セル.選択範囲(, , 2, 0).セルごとにサブルーチン呼び出し "月埋める", w_整数
    
    Set w_整数 = Nothing
End Sub
    
Public Sub 月埋める(w_セル As セル, w_整数 As 整数)
    w_整数.増やす
    w_セル.値 = w_整数.文字列化
End Sub

Public Sub 関数呼び出しテストコード02()
    g_道具.セル.選択範囲(, , 2, 0).セルごとにサブルーチン呼び出し "月埋める", h_変数取得.整数(0, "第", "月")
End Sub

Public Sub 関数呼び出しテストコード03()
    g_道具.セル.選択範囲(選択範囲パターン偶数行).セルごとにサブルーチン呼び出し "月埋める", h_変数取得.整数(0, "第", "月")
End Sub


Public Sub 関数呼び出しテストコード04()
    g_道具.セル.選択範囲(選択範囲パターン奇数列).セルごとにサブルーチン呼び出し "月埋める", h_変数取得.整数(0, "第", "月")
End Sub

Public Sub 関数呼び出しテストコード05()
    g_道具.セル.選択範囲(選択範囲パターン行ステップ, 4).セルごとにサブルーチン呼び出し "月埋める", h_変数取得.整数(0, "第", "月")
End Sub


Public Sub 関数呼び出しテストコード06()
    g_道具.セル.選択範囲(選択範囲パターン列ステップ, 4).セルごとにサブルーチン呼び出し "月埋める", h_変数取得.整数(0, "第", "月")
End Sub

Public Sub 関数呼び出しテストコード07()
    g_道具.セル("A3:G8").選択.セルごとにサブルーチン呼び出し "月埋める", h_変数取得.整数(0, "第", "月")
End Sub


Public Sub 選択範囲の背景色変更09()
    Call g_道具.セル.選択範囲(選択範囲パターン列ステップ, 3).書式.背景パターン連鎖(塗りつぶし) _
        .色連鎖(F_色インデックスからRGB色へ変換(インデックスオレンジ))
End Sub

'============================
' 201801-
'============================

Public Sub 先頭行に第n月挿入()
    Dim w_整数 As New 整数
    
    w_整数.接頭辞を設定 "第"
    w_整数.接尾辞を設定 "月"
    
    g_道具.セル.選択範囲(, , 1, 0).セルごとにサブルーチン呼び出し "月埋める", w_整数
    
    Set w_整数 = Nothing
End Sub
Public Sub 先頭行に第n月挿入_簡略版()
    g_道具.セル.選択範囲(, , 1, 0).セルごとにサブルーチン呼び出し "月埋める", h_変数取得.整数(0, "第", "月")
End Sub

Public Sub 先頭列に第n日挿入()
    Dim w_整数 As New 整数
    
    w_整数.接頭辞を設定 "第"
    w_整数.接尾辞を設定 "日"
    
    g_道具.セル.選択範囲(, , 0, 1).セルごとにサブルーチン呼び出し "月埋める", w_整数
    
    Set w_整数 = Nothing
End Sub
Public Sub 先頭列に第n日挿入_簡略版()
    g_道具.セル.選択範囲(, , 0, 1).セルごとにサブルーチン呼び出し "月埋める", h_変数取得.整数(0, "第", "日")
End Sub

Public Sub 選択範囲の背景色_連鎖変更01()
    g_道具.セル.選択範囲(選択範囲パターン偶数行).書式.背景パターン連鎖(塗りつぶし) _
    .色連鎖 (F_色インデックスからRGB色へ変換(インデックスオレンジ))
End Sub
Public Sub 選択範囲の背景色_連鎖変更02()
    g_道具.セル.選択範囲(選択範囲パターン奇数列).書式.背景パターン連鎖(横細線) _
    .色連鎖 (F_色インデックスからRGB色へ変換(インデックスオレンジ))
End Sub
Public Sub 選択範囲の背景色_連鎖変更03()
    g_道具.セル.選択範囲(選択範囲パターン行ステップ, 3).書式.背景パターン連鎖(塗りつぶし) _
    .色連鎖 (F_色インデックスからRGB色へ変換(インデックスオレンジ))
End Sub
Public Sub 選択範囲の背景色_連鎖変更04()
    g_道具.セル.選択範囲(選択範囲パターン列ステップ, 3).書式.背景パターン連鎖(格子) _
    .色連鎖 (F_色インデックスからRGB色へ変換(インデックスオレンジ))
End Sub




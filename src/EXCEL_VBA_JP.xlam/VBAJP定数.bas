Attribute VB_Name = "VBAJP定数"
Option Explicit

Enum 終端方向
    上橋 = xlUp
    下端 = xlDown
    右端 = xlToRight
    左端 = xlToLeft
End Enum

Enum セル選択方法
    表示形式あり = xlCellTypeAllFormatConditions
    条件設定あり = xlCellTypeAllValidation
    空のセル = xlCellTypeBlanks
    コメントあり = xlCellTypeComments
    定数あり = xlCellTypeConstants
    数式あり = xlCellTypeFormulas
    最後のセル = xlCellTypeLastCell
    同じ表示形式 = xlCellTypeSameFormatConditions
    同じ条件 = xlCellTypeSameValidation
    可視セル = xlCellTypeVisible
End Enum

Enum セル選択条件値
    エラー値 = xlErrors
    論理値 = xlLogical
    数値 = xlNumbers
    文字 = xlTextValues
End Enum

Enum 埋め方
    埋め方標準 = xlFillDefault
    連続データ = xlFillSeries
    コピー = xlFillCopy
    書式のみ = xlFillFormats
    書式なし = xlFillValues
    年単位 = xlFillYears
    月単位 = xlFillMonths
    日単位 = xlFillDays
    週日単位 = xlFillWeekdays
    加算 = xlLinearTrend
    乗算 = xlGrowthTrend
End Enum

Enum シフト方向
    上方向にシフト = xlShiftUp
    下方向にシフト = xlShiftDown
    右方向にシフト = xlShiftToRight
    左方向にシフト = xlShiftToLeft
End Enum

Enum 貼り付け方法
    すべて貼り付け = xlPasteAll
    数式のみ貼り付け = xlPasteFormulas
    値のみ貼り付け = xlPasteValues
    書式のみ貼り付け = xlPasteFormats
    コメントのみ貼り付け = xlPasteComments
    入力規則のみ貼り付け = xlPasteValidation
    罫線を除く全て貼り付け = xlPasteAllExceptBorders
    列幅 = xlPasteColumnWidths
    数式と数値の書式のみ貼り付け = xlPasteFormulasAndNumberFormats
    値と数値の書式のみ貼り付け = xlPasteValuesAndNumberFormats
    すべてのテーマを貼り付け = xlPasteAllUsingSourceTheme
    すべての結合条件付き書式を貼り付け = xlPasteAllMergingConditionalFormats
End Enum

Enum 表示形式パターン定数
    通貨 = 1
    小数点以下1桁 = 2
    小数点以下2桁 = 3
    第4桁まで0埋め = 4
    第8桁まで0埋め = 5
    西暦 = 6
    西暦曜日付き = 7
    和暦 = 8
    和暦曜日付き = 9
    日時 = 10
    日時分 = 11
    日時AMPM = 12
End Enum

Enum セル横位置
    横標準 = xlGeneral
    横左詰め = xlLeft
    横中央揃え = xlCenter
    横右詰め = xlRight
    横繰り返し = xlFill
    横両端揃え = xlJustify
    横選択範囲内で中央 = xlCenterAcrossSelection
    横均等割り付け = xlDistributed
End Enum

Enum セル縦位置
    縦上詰め = xlTop
    縦中央揃え = xlCenter
    縦下詰め = xlBottom
    縦両端揃え = xlJustify
    縦均等割り付け = xlDistributed
End Enum

Enum セル角度
    角度30度 = 30
    角度45度 = 45
    角度60度 = 60
    角度90度 = 90
    角度マイナス30度 = -30
    角度マイナス45度 = -45
    角度マイナス60度 = -60
    角度マイナス90度 = -90
    角度縦方向 = xlVertical
End Enum

'文字列はenumできない
'→定数だとできる
'Enum 標準フォント名
'    MSPゴシック = "ＭＳ Ｐゴシック"
'    MSP明朝 = "ＭＳ Ｐ明朝"
'    MSゴシック = "ＭＳ ゴシック"
'    MS明朝1 = "ＭＳ 明朝"
'    Arial = "Arial"
'    ArialBlack = "Arial Black"
'    メイリオ = "メイリオ"
'End Enum
Public Const フォント名MSPゴシック = "ＭＳ Ｐゴシック"
Public Const フォント名MSP明朝 = "ＭＳ Ｐ明朝"
Public Const フォント名MSゴシック = "ＭＳ ゴシック"
Public Const フォント名MS明朝 = "ＭＳ 明朝"
Public Const フォント名Arial = "Arial"
Public Const フォント名ArialBlack = "Arial Black"
Public Const フォント名メイリオ = "メイリオ"

Enum アンダーラインパターン種類
    下線なし = xlUnderlineStyleNone
    下線 = xlUnderlineStyleSingle
    二重下線 = xlUnderlineStyleDouble
    下線会計 = xlUnderlineStyleSingleAccounting
    二重下線会計 = xlUnderlineStyleDoubleAccounting
End Enum

Enum カラーインデックスパターン
    インデックス黒 = 1
    インデックス白 = 2
    インデックス赤 = 3
    インデックス緑 = 4
    インデックス青 = 5
    インデックス黄色 = 6
    インデックス紫 = 7
    インデックス水色 = 8
    インデックス茶色 = 9
    インデックス深緑 = 10
    インデックス藍色 = 11
    インデックス黄土色 = 12
    インデックス深紫 = 13
    インデックス緑2 = 20
    インデックス灰色 = 15
    インデックス濃い灰色 = 16
    インデックス青紫 = 17
    インデックス紫2 = 18
    インデックス薄い黄色 = 19
    インデックス薄い青 = 20
    インデックス深紫2 = 21
    インデックス肌色 = 22
    インデックス青2 = 23
    インデックス薄い紫 = 24
    インデックス濃い青2 = 25
    インデックス薄い紫2 = 26
    インデックス黄色2 = 27
    インデックス水色2 = 28
    インデックス紫3 = 29
    インデックス茶色2 = 30
    インデックス深緑2 = 31
    インデックス濃い青 = 32
    インデックス青緑 = 33
    インデックス薄い水色 = 34
    インデックス薄い黄緑 = 35
    インデックス薄い黄色2 = 36
    インデックス薄い水色2 = 37
    インデックス薄いピンク = 38
    インデックス薄い紫3 = 39
    インデックス薄い肌色 = 40
    インデックス藍色2 = 41
    インデックス濃い水色 = 42
    インデックス薄い緑 = 43
    インデックス濃い黄色 = 44
    インデックス薄いオレンジ = 45
    インデックスオレンジ = 46
    インデックス藍色3 = 47
    インデックス灰色2 = 48
    インデックス濃い藍色 = 49
    インデックス緑3 = 50
    インデックス濃い灰色2 = 51
    インデックス濃い灰色3 = 52
    インデックス濃いオレンジ = 53
    インデックス濃いピンク = 54
    インデックス濃い青3 = 55
    インデックス濃い灰色4 = 56
    色を自動的に設定 = xlColorIndexAutomatic
    インデックスなし = xlColorIndexNone
End Enum

Enum 罫線位置
    上橋の罫線 = xlEdgeTop
    下端の罫線 = xlEdgeBottom
    左端の罫線 = xlEdgeLeft
    右端の罫線 = xlEdgeRight
    内側の横線 = xlInsideHorizontal
    内側の縦線 = xlInsideVertical
    右下がりの斜め線 = xlDiagonalDown
    右上がりの斜め線 = xlDiagonalUp
End Enum

Enum 罫線線種
    細実線 = xlContinuous
    破線 = xlDash
    一点鎖線 = xlDashDot
    二点鎖線 = xlDashDotDot
    点線 = xlDot
    二重線 = xlDouble
    斜め破線 = xlSlantDashDot
    線なし = xlLineStyleNone
End Enum

Enum 罫線の太さ
    極細 = xlHairline
    細い = xlThin
    中 = xlMedium
    太い = xlThick
End Enum

Enum セル背景色パターン
    塗りつぶし = xlPatternSolid
    灰色75パーセント = xlGray75
    灰色50パーセント = xlGray50
    灰色25パーセント = xlGray25
    灰色16パーセント = xlGray16
    灰色8パーセント = xlGray8
    横線 = xlHorizontal
    縦線 = xlVertical
    右下がり斜め線 = xlDown
    右上がり斜め線 = xlUp
    チェック = xlChecker
    灰色格子 = xlSemiGray75
    横細線 = xlLightHorizontal
    縦細線 = xlLightVertical
    右下がり斜め細線 = xlLightDown
    右上がり斜め細線 = xlLightUp
    格子 = xlGrid
    格子細線 = xlCrissCross
    線形グラデーション = xlPatternLinearGradient
    方形グラデーション = xlPatternRectangularGradient
End Enum

Enum 絶対か相対かアドレス指定
    絶対アドレス = True
    相対アドレス = False
End Enum


Enum ファイル作成フォルダ種別
    現在のフォルダ = 1
    マイドキュメント = 2
    フルパス = 3
    指定なし = 4
End Enum

Enum 選択範囲パターン指定
    選択範囲パターン指定なし = 0
    選択範囲パターン偶数行 = 1
    選択範囲パターン奇数行 = 2
    選択範囲パターン偶数列 = 3
    選択範囲パターン奇数列 = 4
    選択範囲パターン行ステップ = 5
    選択範囲パターン列ステップ = 6
End Enum

Public Enum 文字列区切り文字
    なし = 0
    カンマ = 1
    タブ = 2
    改行 = 3
    Cr = 4
    半角空白 = 5
    その他 = 6
End Enum

Public Enum 配列の値を指定して削除オプション
    全該当要素削除 = 0
    最初の要素だけ削除 = 1
End Enum

Public Enum ファイル書き込み方法
    上書き = 1
    連番 = 2
End Enum

Public Enum テキスト比較方法
    大文字小文字を区別 = vbBinaryCompare
    大文字小文字を区別しない = vbTextCompare
End Enum

Attribute VB_Name = "VBAEN_CONST"
Option Explicit

Enum OrientForEnding '終端方向
    vbeEndUp = xlUp
    vbeEndDown = xlDown
    vbeEndRight = xlToRight
    vbeEndLeft = xlToLeft
End Enum

Enum SelectionMethodForCell 'セル選択方法
    vbeExistsTypeAllFormatConditions = xlCellTypeAllFormatConditions '表示形式あり
    vbeExistsTypeAllValidation = xlCellTypeAllValidation '条件設定あり
    vbeExistsTypeBlanks = xlCellTypeBlanks  '空のセル
    vbeExistsTypeComments = xlCellTypeComments 'コメントあり
    vbeExistsTypeConstants = xlCellTypeConstants '定数あり
    vbeExistsTypeFormulas = xlCellTypeFormulas '数式あり
    vbeExistsTypeLastCell = xlCellTypeLastCell '最後のセル
    vbeExistsTypeSameFormatConditions = xlCellTypeSameFormatConditions '同じ表示形式
    vbeExistsTypeSameValidation = xlCellTypeSameValidation '同じ条件
    vbeExistsTypeVisible = xlCellTypeVisible '可視セル
End Enum

Enum ConditionsForSelectingCells 'セル選択条件値
    vbeValuesOfErrors = xlErrors 'エラー値
    vbeValuesOfLogical = xlLogical '論理値
    vbeValuesOfNumbers = xlNumbers '数値
    vbeValuesOfText = xlTextValues '文字
End Enum

Enum FillMethod '埋め方
    vbeFillMethodDefault = xlFillDefault '埋め方標準
    vbeFillMethodSeries = xlFillSeries '連続データ
    vbeFillMethodCopy = xlFillCopy 'コピー
    vbeFillMethodFormatsOnly = xlFillFormats '書式のみ
    vbeFillMethodValuesOnly = xlFillValues '書式なし
    vbeFillMethodYears = xlFillYears '年単位
    vbeFillMethodMonths = xlFillMonths '月単位
    vbeFillMethodDays = xlFillDays '日単位
    vbeFillMethodWeekdays = xlFillWeekdays  '週日単位
    vbeFillMethodLinearTrend = xlLinearTrend '加算
    vbeFillMethodGrowthTrend = xlGrowthTrend '乗算
End Enum

Enum ShiftOrient 'シフト方向
    vbeShiftForwordToUp = xlShiftUp '上方向にシフト
    vbeShiftForwordToDown = xlShiftDown '下方向にシフト
    vbeShiftForwordToRight = xlShiftToRight '右方向にシフト
    vbeShiftForwordToLeft = xlShiftToLeft '左方向にシフト
End Enum

Enum PasteMethod '貼り付け方法
    vbePasteAll = xlPasteAll 'すべて貼り付け
    vbePasteOnlyFomurals = xlPasteFormulas '数式のみ貼り付け
    vbePasteOnlyValues = xlPasteValues '値のみ貼り付け
    vbePasteOnlyFormats = xlPasteFormats '書式のみ貼り付け
    vbePasteOnlyComments = xlPasteComments 'コメントのみ貼り付け
    vbePasteOnlyValidation = xlPasteValidation '入力規則のみ貼り付け
    vbePasteAllExceptBoders = xlPasteAllExceptBorders '罫線を除く全て貼り付け
    vbePasteOnlyColumnWidths = xlPasteColumnWidths '列幅
    vbePasteOnlyFormulasAndNumberFormats = xlPasteFormulasAndNumberFormats  '数式と数値の書式のみ貼り付け
    vbePasteOnlyValuesAndNumberFormats = xlPasteValuesAndNumberFormats '値と数値の書式のみ貼り付け
    vbePasteAllUsingSourceTheme = xlPasteAllUsingSourceTheme 'すべてのテーマを貼り付け
    vbePasteAllMergingConditionalFormats = xlPasteAllMergingConditionalFormats  'すべての結合条件付き書式を貼り付け
End Enum

Enum VisualFormatPatternForCell '表示形式パターン定数
    vbeCurency = 1 '通貨
    vbeOneDecimalPlace = 2 '小数点以下1桁
    vbeTwoDecimalPlace = 3 '小数点以下2桁
    vbeZeroPaddingUpTo4digit = 4 '第4桁まで0埋め
    vbeZeroPaddingUpTo8digit = 5 '第8桁まで0埋め
    vbeAnnoDomini = 6 '西暦
    vbeAnnoDominiWithDate = 7 '西暦曜日付き
    vbeJapaneseCalendar = 8 '和暦
    vbeJapaneseCalendarWithDate = 9 '和暦曜日付き
    vbeDateAndTime = 10 '日時
    vbeDareAndTimeAndMinutes = 11 '日時分
    vbeDateAndTimeWithAMandPM = 12 '日時AMPM
End Enum

Enum HorizentalPositionForCell 'セル横位置
    vbeGeneral = xlGeneral '横標準
    vbeLeft = xlLeft '横左詰め
    vbeCenter = xlCenter '横中央揃え
    vbeRight = xlRight '横右詰め
    vbeFill = xlFill '横繰り返し
    vbeHorizentalJustify = xlJustify '横両端揃え
    vbeCenterAcrossSelection = xlCenterAcrossSelection '横選択範囲内で中央
    vbeHorizentalDistributed = xlDistributed '横均等割り付け
End Enum

Enum VerticalPositionForCell 'セル縦位置
    vbeTop = xlTop '縦上詰め
    vbeCenter = xlCenter '縦中央揃え
    vbeBottom = xlBottom '縦下詰め
    vbeVerticalJustify = xlJustify '縦両端揃え
    vbeVerticalDistributed = xlDistributed '縦均等割り付け
End Enum

Enum CellDegree  'セル角度
    vbeDegree30 = 30 '角30度
    vbeDegree45 = 45 '角度45度
    vbeDegree60 = 60 '角度60度
    vbeDegree90 = 90 '角度90度
    vbeDegreeMinus30 = -30 '角度マイナス30度
    vbeDegreeMinus45 = -45 '角度マイナス45度
    vbeDegreeMinus60 = -60 '角度マイナス60度
    vbeDegreeMinus90 = -90 '角度マイナス90度
    vbeDegreeVertical = xlVertical '角度縦方向
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
Public Const FontNameMSPGotic = "ＭＳ Ｐゴシック" 'フォント名MSPゴシック
Public Const FontNameMSPMincho = "ＭＳ Ｐ明朝" 'フォント名MSP明朝
Public Const FontNameMSGotic = "ＭＳ ゴシック" 'フォント名MSゴシック
Public Const FontNameMSMincho = "ＭＳ 明朝" 'フォント名MS明朝
Public Const FontNameArial = "Arial" 'フォント名Arial
Public Const FontNameArialBlack = "Arial Black" 'フォント名ArialBlack
Public Const FontNameメイリオ = "メイリオ" 'フォント名メイリオ

Enum StyleOfUnderLinePattern 'アンダーラインパターン種類
    vbeUnderlineStyleNone = xlUnderlineStyleNone '下線なし
    vbeUnderlineStyleSingle = xlUnderlineStyleSingle '下線
    vbeUnderlineStyleDouble = xlUnderlineStyleDouble '二重下線
    vbeUnderlineStyleSingleAccounting = xlUnderlineStyleSingleAccounting '下線会計
    vbeUnderlineStyleDoubleAccounting = xlUnderlineStyleDoubleAccounting '二重下線会計
End Enum

'Enum PatternOfColorIndex 'カラーインデックスパターン
'    IndexBlack = 1 'インデックス黒
'    インデックス白 = 2 'インデックス白
'    インデックス赤 = 3 'インデックス赤
'    インデックス緑 = 4 'インデックス緑
'    インデックス青 = 5 'インデックス青
'    インデックス黄色 = 6 'インデックス黄色
'    インデックス紫 = 7 'インデックス紫
'    インデックス水色 = 8 'インデックス水色
'    インデックス茶色 = 9 'インデックス茶色
'    インデックス深緑 = 10 'インデックス深緑
'    インデックス藍色 = 11 'インデックス藍色
'    インデックス黄土色 = 12 'インデックス黄土色
'    インデックス深紫 = 13 'インデックス深紫
'    インデックス緑2 = 20 'インデックス緑2
'    インデックス灰色 = 15 'インデックス灰色
'    インデックス濃い灰色 = 16 'インデックス濃い灰色
'    インデックス青紫 = 17 'インデックス青紫
'    インデックス紫2 = 18 'インデックス紫2
'    インデックス薄い黄色 = 19 'インデックス薄い黄色
'    インデックス薄い青 = 20 'インデックス薄い青
'    インデックス深紫2 = 21 'インデックス深紫2
'    インデックス肌色 = 22 'インデックス肌色
'    インデックス青2 = 23 'インデックス青2
'    インデックス薄い紫 = 24 'インデックス薄い紫
'    インデックス濃い青2 = 25 'インデックス濃い青2
'    インデックス薄い紫2 = 26 'インデックス薄い紫2
'    インデックス黄色2 = 27 'インデックス黄色2
'    インデックス水色2 = 28 'インデックス水色2
'    インデックス紫3 = 29 'インデックス紫3
'    インデックス茶色2 = 30 'インデックス茶色2
'    インデックス深緑2 = 31 'インデックス深緑2
'    インデックス濃い青 = 32 'インデックス濃い青
'    インデックス青緑 = 33 'インデックス青緑
'    インデックス薄い水色 = 34 'インデックス薄い水色
'    インデックス薄い黄緑 = 35 'インデックス薄い黄緑
'    インデックス薄い黄色2 = 36 'インデックス薄い黄色2
'    インデックス薄い水色2 = 37 'インデックス薄い水色2
'    インデックス薄いピンク = 38 'インデックス薄いピンク
'    インデックス薄い紫3 = 39 'インデックス薄い紫3
'    インデックス薄い肌色 = 40 'インデックス薄い肌色
'    インデックス藍色2 = 41 'インデックス藍色2
'    インデックス濃い水色 = 42 'インデックス濃い水色
'    インデックス薄い緑 = 43 'インデックス薄い緑
'    インデックス濃い黄色 = 44 'インデックス濃い黄色
'    インデックス薄いオレンジ = 45 'インデックス薄いオレンジ
'    インデックスオレンジ = 46 'インデックスオレンジ
'    インデックス藍色3 = 47 'インデックス藍色3
'    インデックス灰色2 = 48 'インデックス灰色2
'    インデックス濃い藍色 = 49 'インデックス濃い藍色
'    インデックス緑3 = 50 'インデックス緑3
'    インデックス濃い灰色2 = 51 'インデックス濃い灰色2
'    インデックス濃い灰色3 = 52 'インデックス濃い灰色3
'    インデックス濃いオレンジ = 53 'インデックス濃いオレンジ
'    インデックス濃いピンク = 54 'インデックス濃いピンク
'    インデックス濃い青3 = 55 'インデックス濃い青3
'    インデックス濃い灰色4 = 56 'インデックス濃い灰色4
'    色を自動的に設定 = xlColorIndexAutomatic '色を自動的に設定
'    インデックスなし = xlColorIndexNone 'インデックスなし
'End Enum

Enum PatternOfColorIndex 'カラーインデックスパターン
   vbeIndexBlack = 1 'インデックス黒
   vbeIndexWhite = 2 'インデックス白
   vbeIndexRed = 3 'インデックス赤
   vbeIndexGreen = 4 'インデックス緑
   vbeIndexBlue = 5 'インデックス青
   vbeIndexYellow = 6 'インデックス黄色
   vbeIndexPurple = 7 'インデックス紫
   vbeIndexLightBlue = 8 'インデックス水色
   vbeIndexBrown = 9 'インデックス茶色
   vbeIndexDarkGreen = 10 'インデックス深緑
   vbeIndexIndigoBlue = 11 'インデックス藍色
   vbeIndexOcher = 12 'インデックス黄土色
   vbeIndexDeepPurple = 13 'インデックス深紫
   vbeIndexGreen2 = 20 'インデックス緑2
   vbeIndexGray = 15 'インデックス灰色
   vbeIndexDarkGray = 16 'インデックス濃い灰色
   vbeIndexBlueViolet = 17 'インデックス青紫
   vbeIndexPurple2 = 18 'インデックス紫2
   vbeIndexLightYellow = 19 'インデックス薄い黄色
   vbeIndexLightBlue2 = 20 'インデックス薄い青
   vbeIndexDeepPurple2 = 21 'インデックス深紫2
   vbeIndexPeach = 22 'インデックス肌色
   vbeIndexBlue2 = 23 'インデックス青2
   vbeIndexLightPurple = 24 'インデックス薄い紫
   vbeIndexDarkBlue2 = 25 'インデックス濃い青2
   vbeIndexLightPurple2 = 26 'インデックス薄い紫2
   vbeIndexYellow2 = 27 'インデックス黄色2
   vbeIndexLightBlue3 = 28 'インデックス水色2
   vbeIndexPurple3 = 29 'インデックス紫3
   vbeIndexBrown2 = 30 'インデックス茶色2
   vbeIndexDeepGreen2 = 31 'インデックス深緑2
   vbeIndexDeepBlue = 32 'インデックス濃い青
   vbeIndexBlueGreen = 33 'インデックス青緑
   vbeIndexLightBlue4 = 34 'インデックス薄い水色
   vbeIndexLightYellowGreen = 35 'インデックス薄い黄緑
   vbeIndexLightYellow2 = 36 'インデックス薄い黄色2
   vbeIndexLightBlue5 = 37 'インデックス薄い水色2
   vbeIndexLightPink = 38 'インデックス薄いピンク
   vbeIndexLightPurple3 = 39 'インデックス薄い紫3
   vbeIndexLightPeach = 40 'インデックス薄い肌色
   vbeIndexIndigoBlue2 = 41 'インデックス藍色2
   vbeIndexDeepBlue2 = 42 'インデックス濃い水色
   vbeIndexLightGreen = 43 'インデックス薄い緑
   vbeIndexDarkYellow = 44 'インデックス濃い黄色
   vbeIndexLightOrange = 45 'インデックス薄いオレンジ
   vbeIndexOrange = 46 'インデックスオレンジ
   vbeIndexIndigoBlue3 = 47 'インデックス藍色3
   vbeIndexGray2 = 48 'インデックス灰色2
   vbeIndexDarkIndigoBlue = 49 'インデックス濃い藍色
   vbeIndexGreen3 = 50 'インデックス緑3
   vbeIndexDarkGray2 = 51 'インデックス濃い灰色2
   vbeIndexDarkGray3 = 52 'インデックス濃い灰色3
   vbeIndexDarkOrange = 53 'インデックス濃いオレンジ
   vbeIndexDarkPink = 54 'インデックス濃いピンク
   vbeIndexDarkBlue3 = 55 'インデックス濃い青3
   vbeIndexDarkGray4 = 56 'インデックス濃い灰色4
   vbeColorIndexAutomatic = xlColorIndexAutomatic '色を自動的に設定
   vbeColorIndexNone = xlColorIndexNone 'インデックスなし
End Enum

Enum BorderPosition '罫線位置
    vbeEdgeTop = xlEdgeTop '上橋の罫線
    vbeEdgeBottom = xlEdgeBottom '下端の罫線
    vbeEdgeLeft = xlEdgeLeft '左端の罫線
    vbeEdgeRight = xlEdgeRight '右端の罫線
    vbeInsideHorizontal = xlInsideHorizontal '内側の横線
    vbeInsideVertical = xlInsideVertical '内側の縦線
    vbeDiagonalDown = xlDiagonalDown '右下がりの斜め線
    vbeDiagonalUp = xlDiagonalUp '右上がりの斜め線
End Enum

Enum StyleOfBoderLine '罫線線種
    vbeContinuous = xlContinuous '細実線
    vbeDash = xlDash '破線
    vbeDashDot = xlDashDot '一点鎖線
    vbeDashDotDot = xlDashDotDot '二点鎖線
    vbeDot = xlDot '点線
    vbeDouble = xlDouble '二重線
    vbeSlantDashDot = xlSlantDashDot '斜め破線
    vbeLineStyleNone = xlLineStyleNone '線なし
End Enum

Enum BorderThickness '罫線の太さ
    vbeHairline = xlHairline '極細
    vbeThin = xlThin '細い
    vbeMedium = xlMedium '中
    vbeThick = xlThick '太い
End Enum

Enum CellBackgroundPattern 'セル背景色パターン
    vbePatternSolid = xlPatternSolid '塗りつぶし
    vbeGray75 = xlGray75 '灰色75パーセント
    vbeGray50 = xlGray50 '灰色50パーセント
    vbeGray25 = xlGray25 '灰色25パーセント
    vbeGray16 = xlGray16 '灰色16パーセント
    vbeGray8 = xlGray8 '灰色8パーセント
    vbeHorizontalLine = xlHorizontal '横線
    vbeVerticalLine = xlVertical '縦線
    vbeDownBackSlash = xlDown '右下がり斜め線
    vbeUpSlash = xlUp '右上がり斜め線
    vbeChecker = xlChecker 'チェック
    vbeSemiGray75 = xlSemiGray75 '灰色格子
    vbeLightHorizontalLine = xlLightHorizontal '横細線
    vbeLightVerticalLine = xlLightVertical '縦細線
    vbeLightDownBackSlash = xlLightDown '右下がり斜め細線
    vbeLightUpSlash = xlLightUp '右上がり斜め細線
    vbeGrid = xlGrid '格子
    vbeCrissCross = xlCrissCross '格子細線
    vbePatternLinearGradient = xlPatternLinearGradient '線形グラデーション
    vbePatternRectangularGradient = xlPatternRectangularGradient '方形グラデーション
End Enum

Enum AddressDesignation '絶対か相対かアドレス指定
    vbeAbsoluteAddress = True
    vbeRelativeAddress = False
End Enum


Enum FolderType 'ファイル作成フォルダ種別
    vbeCurrentFolder = 1 '現在のフォルダ
    vbeMyDocument = 2 'マイドキュメント
    vbeFullPath = 3 'フルパス
    vbeNone = 4 '指定なし
End Enum

Enum SelectionPattern '選択範囲パターン指定
    vbeNonePattern = 0 '選択範囲パターン指定なし
    vbeEvenRows = 1 '選択範囲パターン偶数行
    vbeOddRows = 2 '選択範囲パターン奇数行
    vbeEvenCols = 3 '選択範囲パターン偶数列
    vbeOddCols = 4 '選択範囲パターン奇数列
    vbeRowsByStep = 5 '選択範囲パターン行ステップ
    vbeColsByStep = 6 '選択範囲パターン列ステップ
End Enum

Public Enum SeparatorChar '文字列区切り文字
    vbeNoneChar = 0 'なし
    vbeComma = 1 'カンマ
    vbeTab = 2 'タブ
    vbeReturn = 3 '改行
    vbeCr = 4 'Cr
    vbeSpaceChar = 5 '半角空白
    vbeElseChar = 6 'その他
End Enum

Public Enum DeleteByValueOptionForArrayElement '配列の値を指定して削除オプション
    AllMatchValues = 0 '全該当要素削除
    FirstMatchValueOnly = 1 '最初の要素だけ削除
End Enum

Public Enum FileWriteMethod 'ファイル書き込み方法
    OverWrite = 1 '上書き
    SerealNo = 2 '連番
End Enum

Public Enum TextCompareMode 'テキスト比較方法
    CaseSensitive = vbBinaryCompare
    NoneCaseSensitive = vbTextCompare
End Enum


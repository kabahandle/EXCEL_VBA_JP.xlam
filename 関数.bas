Attribute VB_Name = "関数"
Public Function カスタム1(合計範囲 As Variant) As Variant
    カスタム1 = 切り捨て2(合計(合計範囲), 2)
End Function
Public Function カスタム2(対象セル As Variant) As Variant
    カスタム2 = 商(対象セル, 12)
End Function
'数学のアークコサイン（arccos）を度で返す関数です。
Public Function アークコサイン度(cos値 As Variant) As Variant
    アークコサイン度 = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Acos(cos値))
End Function

'数学のアークサイン（arcsin）を度で返す関数です。
Public Function アークサイン度(sin値 As Variant) As Variant
    アークサイン度 = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Asin(sin値))
End Function

'数学のアークタンジェント（arctan）を、度で返す関数です。
Public Function アークタンジェント度(tan値 As Variant) As Variant
    アークタンジェント度 = Application.WorksheetFunction.Degrees(Atn(tan値))
End Function

'数学のコサイン（cos）を度数から引く関数です。
Public Function コサイン度(度 As Variant) As Variant
    コサイン度 = Cos(Application.WorksheetFunction.Radians(度))
End Function

'数学のサイン（sin）を度数から引く関数です。
Public Function サイン度(度 As Variant) As Variant
    サイン度 = Sin(Application.WorksheetFunction.Radians(度))
End Function

'数学のタンジェント度（tan）を度数から引く関数です。
Public Function タンジェント度(度 As Variant) As Variant
    タンジェント度 = Tan(Application.WorksheetFunction.Radians(度))
End Function

'2進数を10進数に変換します。
Public Function 二進数から十進数(二進数 As Variant) As Variant

    Dim 十進数計算用 As Variant
    Dim j As Long
    Dim x As Long
    
    十進数計算用 = 0

    For j = 1 To Len(二進数)
        If Mid(二進数, Len(二進数) - j + 1, 1) = "1" Then
            x = 2 ^ (j - 1)
            十進数計算用 = 十進数計算用 + x
        End If
    Next j

    二進数から十進数 = 十進数計算用

End Function

' n ÷ m の式の余りを求めます。
Public Function 余り(割られる数n As Variant, 割る数m As Variant) As Variant
    余り = 割られる数n Mod 割る数m
End Function

'16進数を10進数に変換します
Public Function 十六進数から十進数(十六進数 As Variant) As Variant
    十六進数から十進数 = Val("&H" & 十六進数)
End Function

'10進数を2進数に変換します
Public Function 十進数から二進数(十進数 As Variant, Optional 桁数 As Long = 8) As String
    Dim ビットフラグ As Long
    Dim 二進数計算用 As String

    Do Until (十進数 < 2 ^ ビットフラグ)
        If (十進数 And 2 ^ ビットフラグ) <> 0 Then
            二進数計算用 = "1" & 二進数計算用
        Else
            二進数計算用 = "0" & 二進数計算用
        End If

        ビットフラグ = ビットフラグ + 1
    Loop
    
    Dim n As Long
    Dim padding As String
    For n = 1 To 桁数
        padding = padding + "0"
    Next n

    十進数から二進数 = Format(二進数計算用, padding)
End Function

'10進数を16進数に変換します
Public Function 十進数から十六進数(十進数 As Variant, Optional 桁数 As Long = 4) As Variant
    十進数から十六進数 = 十六進数パディング(Hex(十進数), "0", 桁数)
End Function
'機能：指定文字埋め関数
'引数：str　：変換前の文字列
'　　　chr  ：埋める文字(１文字目のみ使用)
'　　　digit：桁数
'戻値：指定文字埋め後の文字列
Private Function 十六進数パディング(ByVal str As String, _
                     ByVal char As String, _
                     ByVal digit As Long) As String
  Dim tmp As String
  tmp = str
  If Len(str) < digit And Len(char) > 0 Then
    tmp = Right(String(digit, char) & str, digit)
  End If
  十六進数パディング = tmp
End Function

'正規表現の置換パターン文字列を指定して、正規表現置換します。
Public Function 正規表現置換(検索対象 As Variant, 置換パターン文字列 As Variant, 置換後の文字列 As Variant, Optional 大文字小文字無視 As Boolean = False, Optional 最初の一致時のみ置換 As Boolean = False)
    r_RegExp.Pattern = 置換パターン文字列
    r_RegExp.IgnoreCase = 大文字小文字無視
    r_RegExp.Global = Not 最初の一致時のみ置換
    If (IsObject(検索対象)) Then
        正規表現置換 = RegEx.Replace(検索対象.Value2, 置換後の文字列)
    Else
        正規表現置換 = RegEx.Replace(検索対象, 置換後の文字列)
    End If
End Function


Public Function 曜日(日付セル As Variant, 種類1から3 As Variant) As Variant
    曜日 = Application.WorksheetFunction.Weekday(日付セル, 種類1から3)
End Function
Public Function 平方根(数値セル As Variant) As Variant
    平方根 = Sqr(数値セル)
End Function
Public Function 平均(平均範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    平均 = Application.WorksheetFunction.Average(平均範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function 文字列長(文字列 As Variant) As Variant
    文字列長 = Len(文字列)
End Function
Public Function 文字置換(置換対象セル As Variant, 置換対象文字列 As Variant, 置換後文字列 As Variant) As Variant
    文字置換 = Application.WorksheetFunction.Substitute(置換対象セル, 置換対象文字列, 置換後文字列)
End Function
Public Function 分散(セル範囲 As Variant) As Variant
    分散 = Application.WorksheetFunction.VarP(セル範囲)
End Function
Public Function 分(日時セル As Variant) As Variant
    分 = Minute(日時セル)
End Function
Public Function 不偏分散(セル範囲 As Variant) As Variant
    不偏分散 = Application.WorksheetFunction.Var(セル範囲)
End Function
Public Function 不偏標準偏差(セル範囲 As Variant) As Variant
    不偏標準偏差 = Application.WorksheetFunction.StDev(セル範囲)
End Function
Public Function 秒(日時セル As Variant) As Variant
    秒 = Second(日時セル)
End Function
Public Function 標準偏差(セル範囲 As Variant) As Variant
    標準偏差 = Application.WorksheetFunction.StDevP(セル範囲)
End Function
Public Function 倍数切り上げ(数値 As Variant, 倍数基準 As Variant) As Variant
    倍数切り上げ = Application.WorksheetFunction.Ceiling(数値, 倍数基準)
End Function
Public Function 倍数切り捨て(数値 As Variant, 倍数基準 As Variant) As Variant
    倍数切り捨て = Application.WorksheetFunction.Floor(数値, 倍数基準)
End Function
Public Function 年(日付セル As Variant) As Variant
    年 = Year(日付セル)
End Function
Public Function 日付変換(年 As Variant, 月 As Variant, 日 As Variant) As Variant
    日付変換 = DateSerial(年, 月, 日)
End Function
Public Function 日付の差(比較単位 As Variant, 日付セル1 As Variant, 日付セル2 As Variant) As Variant
    日付の差 = DateDiff(比較単位, 日付セル1, 日付セル2)
End Function
Public Function 日(日付セル As Variant) As Variant
    日 = Day(日付セル)
End Function
Public Function 中央値(セル範囲 As Variant) As Variant
    中央値 = Application.WorksheetFunction.Median(セル範囲)
End Function
Public Function 大きい方から何番目かの値(セル範囲 As Variant, 順位 As Variant) As Variant
    大きい方から何番目かの値 = Application.WorksheetFunction.Large(セル範囲, 順位)
End Function
Public Function 対数(元の数値 As Variant) As Variant
    対数 = Log(元の数値)
End Function
Public Function 全角文字を半角化(対象文字セル As Variant) As Variant
    全角文字を半角化 = Application.WorksheetFunction.Asc(対象文字セル)
End Function
Public Function 絶対値(数値セル As Variant) As Variant
    絶対値 = Abs(数値セル)
End Function
Public Function 切り上げ(数値 As Variant, 切り上げる桁数 As Variant) As Variant
    切り上げ = Application.WorksheetFunction.RoundUp(数値, 切り上げる桁数)
End Function
Public Function 切り捨て2(数値 As Variant, 切り捨てる桁数｡ As Variant) As Variant
    切り捨て2 = Application.WorksheetFunction.RoundDown(数値, 切り捨てる桁数｡)
End Function
Public Function 切り捨て(数値 As Variant) As Variant
    切り捨て = Int(数値)
End Function
Public Function 数値間ランダム(開始値 As Variant, 終了値 As Variant) As Variant
    数値間ランダム = Application.WorksheetFunction.RandBetween(開始値, 終了値)
End Function
Public Function 数字をローマ数字化(対象セル As Variant) As Variant
    数字をローマ数字化 = Application.WorksheetFunction.Roman(対象セル)
End Function
Public Function 数ヶ月後の月末(開始日 As Variant, 月 As Variant) As Variant
    数ヶ月後の月末 = Application.WorksheetFunction.EoMonth(開始日, 月)
End Function
Public Function 数ヶ月後(開始日 As Variant, 月 As Variant) As Variant
    数ヶ月後 = Application.WorksheetFunction.EDate(開始日, 月)
End Function
Public Function 常用対数(元の数値 As Variant) As Variant
    常用対数 = Application.WorksheetFunction.Log10(元の数値)
End Function
Public Function 小さい方から何番目かの値(セル範囲 As Variant, 順位 As Variant) As Variant
    小さい方から何番目かの値 = Application.WorksheetFunction.Small(セル範囲, 順位)
End Function
Public Function 商(対象数値セル As Variant, 割る数 As Variant) As Variant
    商 = Application.WorksheetFunction.Quotient(対象数値セル, 割る数)
End Function
Public Function 順位(順位調査セル As Variant, セル範囲 As Variant, 順序フラグ As Variant) As Variant
    順位 = Application.WorksheetFunction.Rank(順位調査セル, セル範囲, 順序フラグ)
End Function
Public Function 縦表引(検索値 As Variant, 検索範囲 As Variant, 列番号 As Variant, Optional オプション1 As Variant) As Variant
    縦表引 = Application.WorksheetFunction.VLookup(検索値, 検索範囲, 列番号, オプション1)
End Function
Public Function 自然対数の底eのべき乗(べきとなる数 As Variant) As Variant
    自然対数の底eのべき乗 = Exp(べきとなる数)
End Function
Public Function 自然対数(元の数値 As Variant) As Variant
    自然対数 = Application.WorksheetFunction.Ln(元の数値)
End Function
Public Function 時間変換(時 As Variant, 分 As Variant, 秒 As Variant) As Variant
    時間変換 = TimeSerial(時, 分, 秒)
End Function
Public Function 時(日時セル As Variant) As Variant
    時 = Hour(日時セル)
End Function
Public Function 四捨五入(数値 As Variant, 四捨五入する桁数 As Variant) As Variant
    四捨五入 = Application.WorksheetFunction.Round(数値, 四捨五入する桁数)
End Function
Public Function 最頻値(セル範囲 As Variant) As Variant
    最頻値 = Application.WorksheetFunction.Mode(セル範囲)
End Function
Public Function 最大公約数(数値範囲 As Variant) As Variant
    最大公約数 = Application.WorksheetFunction.Gcd(数値範囲)
End Function
Public Function 最大(検索範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    最大 = Application.WorksheetFunction.max(検索範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function 最小公倍数(数値範囲 As Variant) As Variant
    最小公倍数 = Application.WorksheetFunction.Lcm(数値範囲)
End Function
Public Function 最小(検索範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    最小 = Application.WorksheetFunction.Min(検索範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function 左文字列(文字列 As Variant, 左からの文字数 As Variant) As Variant
    左文字列 = Left(文字列, 左からの文字数)
End Function
Public Function 左右空白文字削除(対象文字セル As Variant) As Variant
    左右空白文字削除 = Application.WorksheetFunction.Trim(対象文字セル)
End Function
Public Function 今() As Variant
    今 = Now
End Function
Public Function 合計(合計範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    合計 = Application.WorksheetFunction.Sum(合計範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function 件数(カウント範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    件数 = Application.WorksheetFunction.Count(カウント範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function 月々積立貯蓄払込額(月利 As Variant, 積立月数 As Variant, 初期金額 As Variant, 目的の積立額 As Variant) As Variant
    月々積立貯蓄払込額 = Application.WorksheetFunction.Pmt(月利, 積立月数, 初期金額, 目的の積立額)
End Function
Public Function 月々ローン返済額中の元金返済額(月利 As Variant, 求めるものは何月目か As Variant, 返済月数 As Variant, 借入金額 As Variant, 最後に残る金額 As Variant) As Variant
    月々ローン返済額中の元金返済額 = Application.WorksheetFunction.PPmt(月利, 求めるものは何月目か, 返済月数, 借入金額, 最後に残る金額)
End Function
Public Function 月々ローン返済額中の金利分額(月利 As Variant, 求めるものは何月目か As Variant, 返済月数 As Variant, 借入額 As Variant, 最後に残る金額 As Variant) As Variant
    月々ローン返済額中の金利分額 = Application.WorksheetFunction.IPmt(月利, 求めるものは何月目か, 返済月数, 借入額, 最後に残る金額)
End Function
Public Function 月々ローン返済額(月利 As Variant, 返済月数 As Variant, 借入額 As Variant, 最後に残る金額 As Variant) As Variant
    月々ローン返済額 = Application.WorksheetFunction.Pmt(月利, 返済月数, 借入額, 最後に残る金額)
End Function
Public Function 月(日付セル As Variant) As Variant
    月 = Month(日付セル)
End Function
Public Function 繰り返し表示(対象文字列 As Variant, 繰り返し回数 As Variant) As Variant
    繰り返し表示 = Application.WorksheetFunction.Rept(対象文字列, 繰り返し回数)
End Function
Public Function 間文字列(文字列 As Variant, 先頭文字番号 As Variant, 抜き出し文字数 As Variant) As Variant
    間文字列 = Mid(文字列, 先頭文字番号, 抜き出し文字数)
End Function
Public Function 角度(ラジアンセル As Variant) As Variant
    角度 = Application.WorksheetFunction.Degrees(ラジアンセル)
End Function
Public Function 階乗(元の数値 As Variant) As Variant
    階乗 = Application.WorksheetFunction.Fact(元の数値)
End Function
Public Function 何周目の日付(日付 As Variant, フラグ1または2 As Variant) As Variant
    何周目の日付 = Application.WorksheetFunction.WeekNum(日付, フラグ1または2)
End Function
Public Function 横表引(検索値 As Variant, 検索範囲 As Variant, 行番号 As Variant, Optional オプション1 As Variant) As Variant
    横表引 = Application.WorksheetFunction.HLookup(検索値, 検索範囲, 行番号, オプション1)
End Function
Public Function 円周率() As Variant
    円周率 = Application.WorksheetFunction.Pi
End Function
Public Function 英単語の先頭文字を大文字化(英単語を含むセル As Variant) As Variant
    英単語の先頭文字を大文字化 = Application.WorksheetFunction.Proper(英単語を含むセル)
End Function
Public Function 英字大文字化(対象文字セル As Variant) As Variant
    英字大文字化 = UCase(対象文字セル)
End Function
Public Function 英字小文字化(対象文字セル As Variant) As Variant
    英字小文字化 = LCase(対象文字セル)
End Function
Public Function 営業日日数(開始日 As Variant, 終了日 As Variant, 祭日を書いたセル範囲 As Variant) As Variant
    営業日日数 = Application.WorksheetFunction.NetworkDays(開始日, 終了日, 祭日を書いたセル範囲)
End Function
Public Function 営業日(開始日 As Variant, 日数 As Variant, 祭日の日付を書いたセル範囲 As Variant) As Variant
    営業日 = Application.WorksheetFunction.WorkDay(開始日, 日数, 祭日の日付を書いたセル範囲)
End Function
Public Function 右文字列(文字列 As Variant, 右からの文字数 As Variant) As Variant
    右文字列 = Right(文字列, 右からの文字数)
End Function
Public Function 一致(検索値 As Variant, 検索範囲 As Variant, 照合の種類 As Variant) As Variant
    一致 = Application.WorksheetFunction.Match(検索値, 検索範囲, 照合の種類)
End Function
Public Function ローン返済額の利子相当分の累計額(月利 As Variant, ローン契約月数 As Variant, 借入額 As Variant, 第n月 As Variant, 第m月 As Variant, 期末払いなら0期首払いなら1 As Variant) As Variant
    ローン返済額の利子相当分の累計額 = Application.WorksheetFunction.CumIPmt(月利, ローン契約月数, 借入額, 第n月, 第m月, 期末払いなら0期首払いなら1)
End Function
Public Function ローン返済額の元金相当分の累計額(月利 As Variant, ローン契約月数 As Variant, 借入額 As Variant, 第n月 As Variant, 第m月 As Variant, 期末払いなら0期首払いなら1 As Variant) As Variant
    ローン返済額の元金相当分の累計額 = Application.WorksheetFunction.CumPrinc(月利, ローン契約月数, 借入額, 第n月, 第m月, 期末払いなら0期首払いなら1)
End Function
Public Function ろーんへんさいがくのりしぶんのるけいいがく(月利 As Variant, ローン契約月数 As Variant, 借入額 As Variant, 第n月 As Variant, 第m月 As Variant, 期末払いなら0期首払いなら1 As Variant) As Variant
    ろーんへんさいがくのりしぶんのるけいいがく = Application.WorksheetFunction.CumIPmt(月利, ローン契約月数, 借入額, 第n月, 第m月, 期末払いなら0期首払いなら1)
End Function
Public Function ろーんへんさいがくのがんきんぶんのるいせきがく(月利 As Variant, ローン契約月数 As Variant, 借入額 As Variant, 第n月 As Variant, 第m月 As Variant, 期末払いなら0期首払いなら1 As Variant) As Variant
    ろーんへんさいがくのがんきんぶんのるいせきがく = Application.WorksheetFunction.CumPrinc(月利, ローン契約月数, 借入額, 第n月, 第m月, 期末払いなら0期首払いなら1)
End Function
Public Function ランダム() As Variant
    ランダム = Rnd
End Function
Public Function ラジアン(角度セル As Variant) As Variant
    ラジアン = Application.WorksheetFunction.Radians(角度セル)
End Function
Public Function よこひょうびき(検索値 As Variant, 検索範囲 As Variant, 行番号 As Variant, 検索の型 As Variant) As Variant
    よこひょうびき = Application.WorksheetFunction.HLookup(検索値, 検索範囲, 行番号, 検索の型)
End Function
Public Function ようび(日付セル As Variant, 種類1から3 As Variant) As Variant
    ようび = Application.WorksheetFunction.Weekday(日付セル, 種類1から3)
End Function
Public Function もっとも近い偶数(対象数値セル As Variant) As Variant
    もっとも近い偶数 = Application.WorksheetFunction.Even(対象数値セル)
End Function
Public Function もっとも近い奇数(対象数値セル As Variant) As Variant
    もっとも近い奇数 = Application.WorksheetFunction.Odd(対象数値セル)
End Function
Public Function もっともちかいぐうすう(対象数値セル As Variant) As Variant
    もっともちかいぐうすう = Application.WorksheetFunction.Even(対象数値セル)
End Function
Public Function もっともちかいきすう(対象数値セル As Variant) As Variant
    もっともちかいきすう = Application.WorksheetFunction.Odd(対象数値セル)
End Function
Public Function もし平均(検索範囲 As Variant, 比較値 As Variant, 平均範囲 As Variant) As Variant
    もし平均 = Application.WorksheetFunction.AverageIf(検索範囲, 比較値, 平均範囲)
End Function
Public Function もし文字列でない(対象セル As Variant) As Variant
    もし文字列でない = Application.WorksheetFunction.IsNonText(対象セル)
End Function
Public Function もし文字列(対象セル As Variant) As Variant
    もし文字列 = Application.WorksheetFunction.IsText(対象セル)
End Function
Public Function もし数値(対象セル As Variant) As Variant
    もし数値 = Application.WorksheetFunction.IsNumber(対象セル)
End Function
Public Function もし合計(検索範囲 As Variant, 比較値 As Variant, 合計範囲 As Variant) As Variant
    もし合計 = Application.WorksheetFunction.SumIf(検索範囲, 比較値, 合計範囲)
End Function
Public Function もし件数(検索範囲 As Variant, 比較値 As Variant) As Variant
    もし件数 = Application.WorksheetFunction.CountIf(検索範囲, 比較値)
End Function
Public Function もし偶数(対象セル As Variant) As Variant
    もし偶数 = Application.WorksheetFunction.IsEven(対象セル)
End Function
Public Function もし空白(対象セル As Variant) As Variant
    もし空白 = IsEmpty(対象セル)
End Function
Public Function もし奇数(対象セル As Variant) As Variant
    もし奇数 = Application.WorksheetFunction.IsOdd(対象セル)
End Function
Public Function もじれつながさ(文字列 As Variant) As Variant
    もじれつながさ = Len(文字列)
End Function
Public Function もしもじれつでない(対象セル As Variant) As Variant
    もしもじれつでない = Application.WorksheetFunction.IsNonText(対象セル)
End Function
Public Function もしもじれつ(対象セル As Variant) As Variant
    もしもじれつ = Application.WorksheetFunction.IsText(対象セル)
End Function
Public Function もしへいきん(検索範囲 As Variant, 比較値 As Variant, 平均範囲 As Variant) As Variant
    もしへいきん = Application.WorksheetFunction.AVERGEIF(検索範囲, 比較値, 平均範囲)
End Function
Public Function もしのっとあさいんど(対象セル As Variant) As Variant
    もしのっとあさいんど = Application.WorksheetFunction.IsNA(対象セル)
End Function
Public Function もじちかん(置換対象セル As Variant, 置換対象文字列 As Variant, 置換後文字列 As Variant) As Variant
    もじちかん = Application.WorksheetFunction.Substitute(置換対象セル, 置換対象文字列, 置換後文字列)
End Function
Public Function もしすうち(対象セル As Variant) As Variant
    もしすうち = Application.WorksheetFunction.IsNumber(対象セル)
End Function
Public Function もしごうけい(検索範囲 As Variant, 比較値 As Variant, 合計範囲 As Variant) As Variant
    もしごうけい = Application.WorksheetFunction.SumIf(検索範囲, 比較値, 合計範囲)
End Function
Public Function もしけんすう(検索範囲 As Variant, 比較値 As Variant) As Variant
    もしけんすう = Application.WorksheetFunction.CountIf(検索範囲, 比較値)
End Function
Public Function もしくうはく(対象セル As Variant) As Variant
    もしくうはく = IsEmpty(対象セル)
End Function
Public Function もしぐうすう(対象セル As Variant) As Variant
    もしぐうすう = Application.WorksheetFunction.IsEven(対象セル)
End Function
Public Function もしきすう(対象セル As Variant) As Variant
    もしきすう = Application.WorksheetFunction.IsOdd(対象セル)
End Function
Public Function もしエラー(対象セル As Variant) As Variant
    もしエラー = Application.WorksheetFunction.IsError(対象セル)
End Function
Public Function もしNA(対象セル As Variant) As Variant
    もしNA = Application.WorksheetFunction.IsNA(対象セル)
End Function
Public Function もし(条件式 As Variant, 真値 As Variant, 偽値 As Variant) As Variant
    もし = IIf(条件式, 真値, 偽値)
End Function
Public Function みぎもじれつ(文字列 As Variant, 右からの文字数 As Variant) As Variant
    みぎもじれつ = Right(文字列, 右からの文字数)
End Function
Public Function または(論理条件1 As Variant, 論理条件2 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    または = Application.WorksheetFunction.Or(論理条件1, 論理条件2, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function べき乗(元の数セル As Variant, べき乗数セル As Variant) As Variant
    べき乗 = Application.WorksheetFunction.Power(元の数セル, べき乗数セル)
End Function
Public Function へいほうこん(数値セル As Variant) As Variant
    へいほうこん = Sqr(数値セル)
End Function
Public Function へいきん(平均範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    へいきん = Application.WorksheetFunction.Average(平均範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function ぶんさん(セル範囲 As Variant) As Variant
    ぶんさん = Application.WorksheetFunction.VarP(セル範囲)
End Function
Public Function ふん(日時セル As Variant) As Variant
    ふん = Minute(日時セル)
End Function
Public Function ふりがな表示(対象文字セル As Variant) As Variant
    ふりがな表示 = Application.WorksheetFunction.Phonetic(対象文字セル)
End Function
Public Function ふりがな(対象文字セル As Variant) As Variant
    ふりがな = Application.WorksheetFunction.Phonetic(対象文字セル)
End Function
Public Function ふへんぶんさん(セル範囲 As Variant) As Variant
    ふへんぶんさん = Application.WorksheetFunction.Var(セル範囲)
End Function
Public Function ふへんひょうじゅんへんさ(セル範囲 As Variant) As Variant
    ふへんひょうじゅんへんさ = Application.WorksheetFunction.StDev(セル範囲)
End Function
Public Function ひょうじゅんへんさ(セル範囲 As Variant) As Variant
    ひょうじゅんへんさ = Application.WorksheetFunction.StDevP(セル範囲)
End Function
Public Function びょう(日時セル As Variant) As Variant
    びょう = Second(日時セル)
End Function
Public Function ひづけへんかん(年 As Variant, 月 As Variant, 日 As Variant) As Variant
    ひづけへんかん = DateSerial(年, 月, 日)
End Function
Public Function ひだりもじれつ(文字列 As Variant, 左からの文字数 As Variant) As Variant
    ひだりもじれつ = Left(文字列, 左からの文字数)
End Function
Public Function ひ(日付セル As Variant) As Variant
    ひ = Day(日付セル)
End Function
Public Function ばいすうきりすて(数値 As Variant, 倍数基準 As Variant) As Variant
    ばいすうきりすて = Application.WorksheetFunction.Floor(数値, 倍数基準)
End Function
Public Function ばいすうきりあげ(数値 As Variant, 倍数基準 As Variant) As Variant
    ばいすうきりあげ = Application.WorksheetFunction.Ceiling(数値, 倍数基準)
End Function
Public Function ねん(日付セル As Variant) As Variant
    ねん = Year(日付セル)
End Function
Public Function なんしゅうめのひづけ(日付 As Variant, フラグ1または2 As Variant) As Variant
    なんしゅうめのひづけ = Application.WorksheetFunction.WeekNum(日付, フラグ1または2)
End Function
Public Function つきづきろーんへんさいがくへんさいがくのがんきんぶん(月利 As Variant, 求めるものは何月目か As Variant, 返済月数 As Variant, 借入金額 As Variant, 最後に残る金額 As Variant) As Variant
    つきづきろーんへんさいがくへんさいがくのがんきんぶん = Application.WorksheetFunction.PPmt(月利, 求めるものは何月目か, 返済月数, 借入金額, 最後に残る金額)
End Function
Public Function つきづきろーんへんさいがくのきんりぶん(月利 As Variant, 求めるものは何月目か As Variant, 返済月数 As Variant, 借入額 As Variant, 最後に残る金額 As Variant) As Variant
    つきづきろーんへんさいがくのきんりぶん = Application.WorksheetFunction.IPmt(月利, 求めるものは何月目か, 返済月数, 借入額, 最後に残る金額)
End Function
Public Function つきづきろーんへんさいがく(月利 As Variant, 返済月数 As Variant, 借入額 As Variant, 最後に残る金額 As Variant) As Variant
    つきづきろーんへんさいがく = Application.WorksheetFunction.Pmt(月利, 返済月数, 借入額, 最後に残る金額)
End Function
Public Function つきづきつみたてちょちくはらいこみがく(月利 As Variant, 積立月数 As Variant, 初期金額 As Variant, 目的の積立額 As Variant) As Variant
    つきづきつみたてちょちくはらいこみがく = Application.WorksheetFunction.Pmt(月利, 積立月数, 初期金額, 目的の積立額)
End Function
Public Function ちゅうおうち(セル範囲 As Variant) As Variant
    ちゅうおうち = Application.WorksheetFunction.Median(セル範囲)
End Function
Public Function ちいさいほうからなんばんめ(セル範囲 As Variant, 順位 As Variant) As Variant
    ちいさいほうからなんばんめ = Application.WorksheetFunction.Small(セル範囲, 順位)
End Function
Public Function タンジェント(数値セル As Variant) As Variant
    タンジェント = Tan(数値セル)
End Function
Public Function たてひょうびき(検索値 As Variant, 検索範囲 As Variant, 列番号 As Variant, 検索の型 As Variant) As Variant
    たてひょうびき = Application.WorksheetFunction.VLookup(検索値, 検索範囲, 列番号, 検索の型)
End Function
Public Function たいすう(元の数値 As Variant) As Variant
    たいすう = Log(元の数値)
End Function
Public Function ぜんかくをはんかくに(対象文字セル As Variant) As Variant
    ぜんかくをはんかくに = Application.WorksheetFunction.Asc(対象文字セル)
End Function
Public Function セル件数(検索範囲 As Variant) As Variant
    セル件数 = Application.WorksheetFunction.CountA(検索範囲)
End Function
Public Function せるけんすう(検索範囲 As Variant) As Variant
    せるけんすう = Application.WorksheetFunction.CountA(検索範囲)
End Function
Public Function ぜったいち(数値セル As Variant) As Variant
    ぜったいち = Abs(数値セル)
End Function
Public Function すうちかんらんだむ(開始値 As Variant, 終了値 As Variant) As Variant
    すうちかんらんだむ = Application.WorksheetFunction.RandBetween(開始値, 終了値)
End Function
Public Function すうじをろーますうじに(対象セル As Variant) As Variant
    すうじをろーますうじに = Application.WorksheetFunction.Roman(対象セル)
End Function
Public Function すうかげつごのげつまつ(開始日 As Variant, 月 As Variant) As Variant
    すうかげつごのげつまつ = Application.WorksheetFunction.EoMonth(開始日, 月)
End Function
Public Function すうかげつご(開始日 As Variant, 月 As Variant) As Variant
    すうかげつご = Application.WorksheetFunction.EDate(開始日, 月)
End Function
Public Function じょうようたいすう(元の数値 As Variant) As Variant
    じょうようたいすう = Application.WorksheetFunction.Log10(元の数値)
End Function
Public Function しょう(対象数値セル As Variant, 割る数 As Variant) As Variant
    しょう = Application.WorksheetFunction.Quotient(対象数値セル, 割る数)
End Function
Public Function じゅんい(順位調査セル As Variant, セル範囲 As Variant, 順序フラグ As Variant) As Variant
    じゅんい = Application.WorksheetFunction.Rank(順位調査セル, セル範囲, 順序フラグ)
End Function
Public Function しぜんたいすうのていのべきじょう(べきとなる数 As Variant) As Variant
    しぜんたいすうのていのべきじょう = Exp(べきとなる数)
End Function
Public Function しぜんたいすう(元の数値 As Variant) As Variant
    しぜんたいすう = Application.WorksheetFunction.Ln(元の数値)
End Function
Public Function ししゃごにゅう(数値 As Variant, 四捨五入する桁数 As Variant) As Variant
    ししゃごにゅう = Application.WorksheetFunction.Round(数値, 四捨五入する桁数)
End Function
Public Function じかんへんかん(時 As Variant, 分 As Variant, 秒 As Variant) As Variant
    じかんへんかん = TimeSerial(時, 分, 秒)
End Function
Public Function じかんのさ(比較単位 As Variant, 日付セル1 As Variant, 日付セル2 As Variant) As Variant
    じかんのさ = DateDiff(比較単位, 日付セル1, 日付セル2)
End Function
Public Function じ(日時セル As Variant) As Variant
    じ = Hour(日時セル)
End Function
Public Function さゆうくうはくさくじょ(対象文字セル As Variant) As Variant
    さゆうくうはくさくじょ = Application.WorksheetFunction.Trim(対象文字セル)
End Function
Public Function サイン(数値セル As Variant) As Variant
    サイン = Sin(数値セル)
End Function
Public Function さいひんち(セル範囲 As Variant) As Variant
    さいひんち = Application.WorksheetFunction.Mode(セル範囲)
End Function
Public Function さいだいこうばいすう(数値範囲 As Variant) As Variant
    さいだいこうばいすう = Application.WorksheetFunction.Gcd(数値範囲)
End Function
Public Function さいだい(検索範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    さいだい = Application.WorksheetFunction.max(検索範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function さいしょうこうばいすう(数値範囲 As Variant) As Variant
    さいしょうこうばいすう = Application.WorksheetFunction.Lcm(数値範囲)
End Function
Public Function さいしょう(検索範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    さいしょう = Application.WorksheetFunction.Min(検索範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function こもじ(対象文字セル As Variant) As Variant
    こもじ = LCase(対象文字セル)
End Function
Public Function コサイン(数値セル As Variant) As Variant
    コサイン = Cos(数値セル)
End Function
Public Function ごうけい(合計範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    ごうけい = Application.WorksheetFunction.Sum(合計範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function けんすう(カウント範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    けんすう = Application.WorksheetFunction.Count(カウント範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function げつ(日付セル As Variant) As Variant
    げつ = Month(日付セル)
End Function
Public Function くりかえしひょうじ(対象文字列 As Variant, 繰り返し回数 As Variant) As Variant
    くりかえしひょうじ = Application.WorksheetFunction.Rept(対象文字列, 繰り返し回数)
End Function
Public Function きりすて2(数値 As Variant, 切り捨てる桁数｡ As Variant) As Variant
    きりすて2 = Application.WorksheetFunction.RoundDown(数値, 切り捨てる桁数｡)
End Function
Public Function きりすて(数値 As Variant) As Variant
    きりすて = Int(数値)
End Function
Public Function きりあげ(数値 As Variant, 切り上げる桁数 As Variant) As Variant
    きりあげ = Application.WorksheetFunction.RoundUp(数値, 切り上げる桁数)
End Function
Public Function かつ(論理条件1 As Variant, 論理条件2 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    かつ = Application.WorksheetFunction.And(論理条件1, 論理条件2, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function かくど(ラジアンセル As Variant) As Variant
    かくど = Application.WorksheetFunction.Degrees(ラジアンセル)
End Function
Public Function かいじょう(元の数値 As Variant) As Variant
    かいじょう = Application.WorksheetFunction.Fact(元の数値)
End Function
Public Function おおもじ(対象文字セル As Variant) As Variant
    おおもじ = UCase(対象文字セル)
End Function
Public Function おおきいほうからなんばんめ(セル範囲 As Variant, 順位 As Variant) As Variant
    おおきいほうからなんばんめ = Application.WorksheetFunction.Large(セル範囲, 順位)
End Function
Public Function えんしゅうりつ() As Variant
    えんしゅうりつ = Application.WorksheetFunction.Pi
End Function
Public Function えいたんごのせんとうもじをおおもじか(英単語を含むセル As Variant) As Variant
    えいたんごのせんとうもじをおおもじか = Application.WorksheetFunction.Proper(英単語を含むセル)
End Function
Public Function えいぎょうびにっすう(開始日 As Variant, 終了日 As Variant, 祭日を書いたセル範囲 As Variant) As Variant
    えいぎょうびにっすう = Application.WorksheetFunction.NetworkDays(開始日, 終了日, 祭日を書いたセル範囲)
End Function
Public Function えいぎょうび(開始日 As Variant, 日数 As Variant, 祭日の日付を書いたセル範囲 As Variant) As Variant
    えいぎょうび = Application.WorksheetFunction.WorkDay(開始日, 日数, 祭日の日付を書いたセル範囲)
End Function
Public Function インデックス(検索範囲 As Variant, 行番号 As Variant, 列番号 As Variant) As Variant
    インデックス = Application.WorksheetFunction.index(検索範囲, 行番号, 列番号)
End Function
Public Function いま() As Variant
    いま = Now
End Function
Public Function いっち(検索値 As Variant, 検索範囲 As Variant, 照合の種類 As Variant) As Variant
    いっち = Application.WorksheetFunction.Match(検索値, 検索範囲, 照合の種類)
End Function
Public Function あいだもじれつ(文字列 As Variant, 先頭文字番号 As Variant, 抜き出し文字数 As Variant) As Variant
    あいだもじれつ = Mid(文字列, 先頭文字番号, 抜き出し文字数)
End Function
Public Function アークタンジェント(x座標 As Variant, y座標 As Variant) As Variant
    アークタンジェント = Application.WorksheetFunction.Atan2(x座標, y座標)
End Function
Public Function アークサイン(元のサインの数値 As Variant) As Variant
    アークサイン = Application.WorksheetFunction.Asin(元のサインの数値)
End Function
Public Function アークコサイン(元のコサインの数値 As Variant) As Variant
    アークコサイン = Application.WorksheetFunction.Acos(元のコサインの数値)
End Function
Public Function nPr何通り(総数セル As Variant, 取り出す分セル As Variant) As Variant
    nPr何通り = Application.WorksheetFunction.Permut(総数セル, 取り出す分セル)
End Function

Public Function nCr何通り(総数セル As Variant, 取り出す分セル As Variant) As Variant
    nCr何通り = Application.WorksheetFunction.Combin(総数セル, 取り出す分セル)
End Function

Public Function 日付化(日付シリアル As Variant) As Date
    日付化 = CDate(日付シリアル)
End Function

Public Function いろ(Optional 赤 As Variant = 255, Optional 緑 As Variant = 255, Optional 青 As Variant = 255) As Variant
    いろ = RGB(赤, 緑, 青)
End Function

Public Function 色(Optional 赤 As Variant = 255, Optional 緑 As Variant = 255, Optional 青 As Variant = 255) As Variant
    色 = RGB(赤, 緑, 青)
End Function

Public Function 色インデックスからRGB色へ変換(idx As カラーインデックスパターン) As Variant
    Select Case idx
    Case 1
        色インデックスからRGB色へ変換 = RGB(0, 0, 0)
    Case 2
        色インデックスからRGB色へ変換 = RGB(255, 255, 255)
    Case 3
        色インデックスからRGB色へ変換 = RGB(255, 0, 0)
    Case 4
        色インデックスからRGB色へ変換 = RGB(0, 255, 0)
    Case 5
        色インデックスからRGB色へ変換 = RGB(0, 0, 255)
    Case 6
        色インデックスからRGB色へ変換 = RGB(255, 255, 0)
    Case 7
        色インデックスからRGB色へ変換 = RGB(255, 0, 255)
    Case 8
        色インデックスからRGB色へ変換 = RGB(0, 255, 255)
    Case 9
        色インデックスからRGB色へ変換 = RGB(128, 0, 0)
    Case 10
        色インデックスからRGB色へ変換 = RGB(0, 128, 0)
    Case 11
        色インデックスからRGB色へ変換 = RGB(0, 0, 128)
    Case 12
        色インデックスからRGB色へ変換 = RGB(128, 128, 0)
    Case 13
        色インデックスからRGB色へ変換 = RGB(128, 0, 128)
    Case 14
        色インデックスからRGB色へ変換 = RGB(0, 128, 128)
    Case 15
        色インデックスからRGB色へ変換 = RGB(192, 192, 192)
    Case 16
        色インデックスからRGB色へ変換 = RGB(128, 128, 128)
    Case 17
        色インデックスからRGB色へ変換 = RGB(153, 153, 255)
    Case 18
        色インデックスからRGB色へ変換 = RGB(153, 51, 102)
    Case 19
        色インデックスからRGB色へ変換 = RGB(255, 255, 204)
    Case 20
        色インデックスからRGB色へ変換 = RGB(204, 255, 255)
    Case 21
        色インデックスからRGB色へ変換 = RGB(102, 0, 102)
    Case 22
        色インデックスからRGB色へ変換 = RGB(255, 128, 128)
    Case 23
        色インデックスからRGB色へ変換 = RGB(0, 102, 204)
    Case 24
        色インデックスからRGB色へ変換 = RGB(204, 204, 255)
    Case 25
        色インデックスからRGB色へ変換 = RGB(0, 0, 128)
    Case 26
        色インデックスからRGB色へ変換 = RGB(255, 0, 255)
    Case 27
        色インデックスからRGB色へ変換 = RGB(255, 255, 0)
    Case 28
        色インデックスからRGB色へ変換 = RGB(0, 255, 255)
    Case 29
        色インデックスからRGB色へ変換 = RGB(128, 0, 128)
    Case 30
        色インデックスからRGB色へ変換 = RGB(128, 0, 0)
    Case 31
        色インデックスからRGB色へ変換 = RGB(0, 128, 128)
    Case 32
        色インデックスからRGB色へ変換 = RGB(0, 0, 255)
    Case 33
        色インデックスからRGB色へ変換 = RGB(0, 204, 255)
    Case 34
        色インデックスからRGB色へ変換 = RGB(204, 255, 255)
    Case 35
        色インデックスからRGB色へ変換 = RGB(204, 255, 204)
    Case 36
        色インデックスからRGB色へ変換 = RGB(255, 255, 153)
    Case 37
        色インデックスからRGB色へ変換 = RGB(153, 204, 255)
    Case 38
        色インデックスからRGB色へ変換 = RGB(255, 153, 204)
    Case 39
        色インデックスからRGB色へ変換 = RGB(204, 153, 255)
    Case 40
        色インデックスからRGB色へ変換 = RGB(255, 204, 153)
    Case 41
        色インデックスからRGB色へ変換 = RGB(51, 102, 255)
    Case 42
        色インデックスからRGB色へ変換 = RGB(51, 204, 204)
    Case 43
        色インデックスからRGB色へ変換 = RGB(153, 204, 0)
    Case 44
        色インデックスからRGB色へ変換 = RGB(255, 204, 0)
    Case 45
        色インデックスからRGB色へ変換 = RGB(255, 153, 0)
    Case 46
        色インデックスからRGB色へ変換 = RGB(255, 102, 0)
    Case 47
        色インデックスからRGB色へ変換 = RGB(102, 102, 153)
    Case 48
        色インデックスからRGB色へ変換 = RGB(150, 150, 150)
    Case 49
        色インデックスからRGB色へ変換 = RGB(0, 51, 102)
    Case 50
        色インデックスからRGB色へ変換 = RGB(51, 153, 102)
    Case 51
        色インデックスからRGB色へ変換 = RGB(0, 51, 0)
    Case 52
        色インデックスからRGB色へ変換 = RGB(51, 51, 0)
    Case 53
        色インデックスからRGB色へ変換 = RGB(153, 51, 0)
    Case 54
        色インデックスからRGB色へ変換 = RGB(153, 51, 102)
    Case 55
        色インデックスからRGB色へ変換 = RGB(51, 51, 153)
    Case 56
        色インデックスからRGB色へ変換 = RGB(51, 51, 51)
    End Select

End Function

Public Function 色の三原色を取得(色 As Long, ByRef 赤 As Long, ByRef 緑 As Long, ByRef 青 As Long)
    赤 = 色 Mod 256
    緑 = Int(色 / 256) Mod 256
    青 = Int(色 / 256 / 256)
End Function

Public Function 今日()
    今日 = Int(今())
End Function

Public Function 今日の日付()
    今日の日付 = Trim(日付化(今日()))
End Function

Public Function 今の日付()
    今の日付 = 日付化(今())
End Function








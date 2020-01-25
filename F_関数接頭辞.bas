Attribute VB_Name = "F_関数接頭辞"
Option Explicit

Public Function Fcst_カスタム1(合計範囲 As Variant) As Variant
    Fcst_カスタム1 = F_切り捨て2(F_合計(合計範囲), 2)
End Function
Public Function Fcst_カスタム2(対象セル As Variant) As Variant
    Fcst_カスタム2 = F_商(対象セル, 12)
End Function
'数学のアークコサイン（arccos）を度で返す関数です。
Public Function Fmath_アークコサイン度(cos値 As Variant) As Variant
    Fmath_アークコサイン度 = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Acos(cos値))
End Function

'数学のアークサイン（arcsin）を度で返す関数です。
Public Function Fmath_アークサイン度(sin値 As Variant) As Variant
    Fmath_アークサイン度 = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Asin(sin値))
End Function

'数学のアークタンジェント（arctan）を、度で返す関数です。
Public Function Fmath_アークタンジェント度(tan値 As Variant) As Variant
    Fmath_アークタンジェント度 = Application.WorksheetFunction.Degrees(Atn(tan値))
End Function

'数学のコサイン（cos）を度数から引く関数です。
Public Function Fmath_コサイン度(度 As Variant) As Variant
    Fmath_コサイン度 = Cos(Application.WorksheetFunction.Radians(度))
End Function

'数学のサイン（sin）を度数から引く関数です。
Public Function Fmath_サイン度(度 As Variant) As Variant
    Fmath_サイン度 = Sin(Application.WorksheetFunction.Radians(度))
End Function

'数学のタンジェント度（tan）を度数から引く関数です。
Public Function Fmath_タンジェント度(度 As Variant) As Variant
    Fmath_タンジェント度 = Tan(Application.WorksheetFunction.Radians(度))
End Function

'2進数を10進数に変換します。
Public Function Fbit_二進数から十進数(二進数 As Variant) As Variant

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

    Fbit_二進数から十進数 = 十進数計算用

End Function

' n ÷ m の式の余りを求めます。
Public Function Fmath_余り(割られる数n As Variant, 割る数m As Variant) As Variant
    Fmath_余り = 割られる数n Mod 割る数m
End Function

'16進数を10進数に変換します
Public Function Fbit_十六進数から十進数(十六進数 As Variant) As Variant
    Fbit_十六進数から十進数 = Val("&H" & 十六進数)
End Function

'10進数を2進数に変換します
Public Function Fbit_十進数から二進数(十進数 As Variant, Optional 桁数 As Long = 8) As String
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

    Fbit_十進数から二進数 = Format(二進数計算用, padding)
End Function

'10進数を16進数に変換します
Public Function Fbit_十進数から十六進数(十進数 As Variant, Optional 桁数 As Long = 4) As Variant
    Fbit_十進数から十六進数 = Fbit_十六進数パディング(Hex(十進数), "0", 桁数)
End Function
'機能：指定文字埋め関数
'引数：str　：変換前の文字列
'　　　chr  ：埋める文字(１文字目のみ使用)
'　　　digit：桁数
'戻値：指定文字埋め後の文字列
Private Function Fbit_十六進数パディング(ByVal str As String, _
                     ByVal char As String, _
                     ByVal digit As Long) As String
  Dim tmp As String
  tmp = str
  If Len(str) < digit And Len(char) > 0 Then
    tmp = Right(String(digit, char) & str, digit)
  End If
  Fbit_十六進数パディング = tmp
End Function

'正規表現の置換パターン文字列を指定して、正規表現置換します。
Public Function Freg_正規表現置換(検索対象 As Variant, 置換パターン文字列 As Variant, 置換後の文字列 As Variant, Optional 大文字小文字無視 As Boolean = False, Optional 最初の一致時のみ置換 As Boolean = False)
    r_RegExp.Pattern = 置換パターン文字列
    r_RegExp.IgnoreCase = 大文字小文字無視
    r_RegExp.Global = Not 最初の一致時のみ置換
    If (IsObject(検索対象)) Then
        Freg_正規表現置換 = r_RegExp.Replace(検索対象.Value2, 置換後の文字列)
    Else
        Freg_正規表現置換 = r_RegExp.Replace(検索対象, 置換後の文字列)
    End If
End Function


Public Function Fdate_曜日(日付セル As Variant, 種類1から3 As Variant) As Variant
    Fdate_曜日 = Application.WorksheetFunction.Weekday(日付セル, 種類1から3)
End Function
Public Function Fmath_平方根(数値セル As Variant) As Variant
    Fmath_平方根 = Sqr(数値セル)
End Function
Public Function Fstat_平均(平均範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    Fstat_平均 = Application.WorksheetFunction.Average(平均範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function Fstr_文字列長(文字列 As Variant) As Variant
    Fstr_文字列長 = Len(文字列)
End Function
Public Function Fstr_文字置換(置換対象セル As Variant, 置換対象文字列 As Variant, 置換後文字列 As Variant) As Variant
    Fstr_文字置換 = Application.WorksheetFunction.Substitute(置換対象セル, 置換対象文字列, 置換後文字列)
End Function
Public Function Fstat_分散(セル範囲 As Variant) As Variant
    Fstat_分散 = Application.WorksheetFunction.VarP(セル範囲)
End Function
Public Function Fdate_分(日時セル As Variant) As Variant
    Fdate_分 = Minute(日時セル)
End Function
Public Function Fstat_不偏分散(セル範囲 As Variant) As Variant
    Fstat_不偏分散 = Application.WorksheetFunction.Var(セル範囲)
End Function
Public Function Fstat_不偏標準偏差(セル範囲 As Variant) As Variant
    Fstat_不偏標準偏差 = Application.WorksheetFunction.StDev(セル範囲)
End Function
Public Function Fdate_秒(日時セル As Variant) As Variant
    Fdate_秒 = Second(日時セル)
End Function
Public Function Fstat_標準偏差(セル範囲 As Variant) As Variant
    Fstat_標準偏差 = Application.WorksheetFunction.StDevP(セル範囲)
End Function
Public Function Fmath_倍数切り上げ(数値 As Variant, 倍数基準 As Variant) As Variant
    Fmath_倍数切り上げ = Application.WorksheetFunction.Ceiling(数値, 倍数基準)
End Function
Public Function Fmath_倍数切り捨て(数値 As Variant, 倍数基準 As Variant) As Variant
    Fmath_倍数切り捨て = Application.WorksheetFunction.Floor(数値, 倍数基準)
End Function
Public Function Fdate_年(日付セル As Variant) As Variant
    Fdate_年 = Year(日付セル)
End Function
Public Function Fdate_日付変換(年 As Variant, 月 As Variant, 日 As Variant) As Variant
    Fdate_日付変換 = DateSerial(年, 月, 日)
End Function
Public Function Fdate_日付の差(比較単位 As Variant, 日付セル1 As Variant, 日付セル2 As Variant) As Variant
    Fdate_日付の差 = DateDiff(比較単位, 日付セル1, 日付セル2)
End Function
Public Function Fdate_日(日付セル As Variant) As Variant
    Fdate_日 = Day(日付セル)
End Function
Public Function Fmath_中央値(セル範囲 As Variant) As Variant
    Fmath_中央値 = Application.WorksheetFunction.Median(セル範囲)
End Function
Public Function Fexcel_大きい方から何番目かの値(セル範囲 As Variant, 順位 As Variant) As Variant
    Fexcel_大きい方から何番目かの値 = Application.WorksheetFunction.Large(セル範囲, 順位)
End Function
Public Function Fmath_対数(元の数値 As Variant) As Variant
    Fmath_対数 = Log(元の数値)
End Function
Public Function Fstr_全角文字を半角化(対象文字セル As Variant) As Variant
    Fstr_全角文字を半角化 = Application.WorksheetFunction.Asc(対象文字セル)
End Function
Public Function Fmath_絶対値(数値セル As Variant) As Variant
    Fmath_絶対値 = Abs(数値セル)
End Function
Public Function Fmath_切り上げ(数値 As Variant, 切り上げる桁数 As Variant) As Variant
    Fmath_切り上げ = Application.WorksheetFunction.RoundUp(数値, 切り上げる桁数)
End Function
Public Function Fmath_切り捨て2(数値 As Variant, 切り捨てる桁数｡ As Variant) As Variant
    Fmath_切り捨て2 = Application.WorksheetFunction.RoundDown(数値, 切り捨てる桁数｡)
End Function
Public Function Fmath_切り捨て(数値 As Variant) As Variant
    Fmath_切り捨て = Int(数値)
End Function
Public Function Fmath_数値間ランダム(開始値 As Variant, 終了値 As Variant) As Variant
    Fmath_数値間ランダム = Application.WorksheetFunction.RandBetween(開始値, 終了値)
End Function
Public Function Fstr_数字をローマ数字化(対象セル As Variant) As Variant
    Fstr_数字をローマ数字化 = Application.WorksheetFunction.Roman(対象セル)
End Function
Public Function Fdate_数ヶ月後の月末(開始日 As Variant, 月 As Variant) As Variant
    Fdate_数ヶ月後の月末 = Application.WorksheetFunction.EoMonth(開始日, 月)
End Function
Public Function Fdate_数ヶ月後(開始日 As Variant, 月 As Variant) As Variant
    Fdate_数ヶ月後 = Application.WorksheetFunction.EDate(開始日, 月)
End Function
Public Function Fmath_常用対数(元の数値 As Variant) As Variant
    Fmath_常用対数 = Application.WorksheetFunction.Log10(元の数値)
End Function
Public Function Fexcel_小さい方から何番目かの値(セル範囲 As Variant, 順位 As Variant) As Variant
    Fexcel_小さい方から何番目かの値 = Application.WorksheetFunction.Small(セル範囲, 順位)
End Function
Public Function Fmath_商(対象数値セル As Variant, 割る数 As Variant) As Variant
    Fmath_商 = Application.WorksheetFunction.Quotient(対象数値セル, 割る数)
End Function
Public Function Fexcel_順位(順位調査セル As Variant, セル範囲 As Variant, 順序フラグ As Variant) As Variant
    Fexcel_順位 = Application.WorksheetFunction.Rank(順位調査セル, セル範囲, 順序フラグ)
End Function
Public Function Fexcel_縦表引(検索値 As Variant, 検索範囲 As Variant, 列番号 As Variant, Optional オプション1 As Variant) As Variant
    Fexcel_縦表引 = Application.WorksheetFunction.VLookup(検索値, 検索範囲, 列番号, オプション1)
End Function
Public Function Fmath_自然対数の底eのべき乗(べきとなる数 As Variant) As Variant
    Fmath_自然対数の底eのべき乗 = Exp(べきとなる数)
End Function
Public Function Fmath_自然対数(元の数値 As Variant) As Variant
    Fmath_自然対数 = Application.WorksheetFunction.Ln(元の数値)
End Function
Public Function Fdate_時間変換(時 As Variant, 分 As Variant, 秒 As Variant) As Variant
    Fdate_時間変換 = TimeSerial(時, 分, 秒)
End Function
Public Function Fdate_時(日時セル As Variant) As Variant
    Fdate_時 = Hour(日時セル)
End Function
Public Function Fmath_四捨五入(数値 As Variant, 四捨五入する桁数 As Variant) As Variant
    Fmath_四捨五入 = Application.WorksheetFunction.Round(数値, 四捨五入する桁数)
End Function
Public Function Fstat_最頻値(セル範囲 As Variant) As Variant
    Fstat_最頻値 = Application.WorksheetFunction.Mode(セル範囲)
End Function
Public Function Fmath_最大公約数(数値範囲 As Variant) As Variant
    Fmath_最大公約数 = Application.WorksheetFunction.Gcd(数値範囲)
End Function
Public Function Fmath_最大(検索範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    Fmath_最大 = Application.WorksheetFunction.max(検索範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function Fmath_最小公倍数(数値範囲 As Variant) As Variant
    Fmath_最小公倍数 = Application.WorksheetFunction.Lcm(数値範囲)
End Function
Public Function Fmath_最小(検索範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    Fmath_最小 = Application.WorksheetFunction.Min(検索範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function Fstr_左文字列(文字列 As Variant, 左からの文字数 As Variant) As Variant
    Fstr_左文字列 = Left(文字列, 左からの文字数)
End Function
Public Function Fstr_左右空白文字削除(対象文字セル As Variant) As Variant
    Fstr_左右空白文字削除 = Application.WorksheetFunction.Trim(対象文字セル)
End Function
Public Function Fdate_今() As Variant
    Fdate_今 = Now
End Function
Public Function Fexcel_合計(合計範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    Fexcel_合計 = Application.WorksheetFunction.Sum(合計範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function Fexcel_件数(カウント範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    Fexcel_件数 = Application.WorksheetFunction.Count(カウント範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function Ffin_月々積立貯蓄払込額(月利 As Variant, 積立月数 As Variant, 初期金額 As Variant, 目的の積立額 As Variant) As Variant
    Ffin_月々積立貯蓄払込額 = Application.WorksheetFunction.Pmt(月利, 積立月数, 初期金額, 目的の積立額)
End Function
Public Function Ffin_月々ローン返済額中の元金返済額(月利 As Variant, 求めるものは何月目か As Variant, 返済月数 As Variant, 借入金額 As Variant, 最後に残る金額 As Variant) As Variant
    Ffin_月々ローン返済額中の元金返済額 = Application.WorksheetFunction.PPmt(月利, 求めるものは何月目か, 返済月数, 借入金額, 最後に残る金額)
End Function
Public Function Ffin_月々ローン返済額中の金利分額(月利 As Variant, 求めるものは何月目か As Variant, 返済月数 As Variant, 借入額 As Variant, 最後に残る金額 As Variant) As Variant
    Ffin_月々ローン返済額中の金利分額 = Application.WorksheetFunction.IPmt(月利, 求めるものは何月目か, 返済月数, 借入額, 最後に残る金額)
End Function
Public Function Ffin_月々ローン返済額(月利 As Variant, 返済月数 As Variant, 借入額 As Variant, 最後に残る金額 As Variant) As Variant
    Ffin_月々ローン返済額 = Application.WorksheetFunction.Pmt(月利, 返済月数, 借入額, 最後に残る金額)
End Function
Public Function Fdate_月(日付セル As Variant) As Variant
    Fdate_月 = Month(日付セル)
End Function
Public Function Fstr_繰り返し表示(対象文字列 As Variant, 繰り返し回数 As Variant) As Variant
    Fstr_繰り返し表示 = Application.WorksheetFunction.Rept(対象文字列, 繰り返し回数)
End Function
Public Function Fstr_間文字列(文字列 As Variant, 先頭文字番号 As Variant, 抜き出し文字数 As Variant) As Variant
    Fstr_間文字列 = Mid(文字列, 先頭文字番号, 抜き出し文字数)
End Function
Public Function Fmath_角度(ラジアンセル As Variant) As Variant
    Fmath_角度 = Application.WorksheetFunction.Degrees(ラジアンセル)
End Function
Public Function Fmath_階乗(元の数値 As Variant) As Variant
    Fmath_階乗 = Application.WorksheetFunction.Fact(元の数値)
End Function
Public Function Fdate_何周目の日付(日付 As Variant, フラグ1または2 As Variant) As Variant
    Fdate_何周目の日付 = Application.WorksheetFunction.WeekNum(日付, フラグ1または2)
End Function
Public Function Fexcel_横表引(検索値 As Variant, 検索範囲 As Variant, 行番号 As Variant, Optional オプション1 As Variant) As Variant
    Fexcel_横表引 = Application.WorksheetFunction.HLookup(検索値, 検索範囲, 行番号, オプション1)
End Function
Public Function Fmath_円周率() As Variant
    Fmath_円周率 = Application.WorksheetFunction.Pi
End Function
Public Function Fstr_英単語の先頭文字を大文字化(英単語を含むセル As Variant) As Variant
    Fstr_英単語の先頭文字を大文字化 = Application.WorksheetFunction.Proper(英単語を含むセル)
End Function
Public Function Fstr_英字大文字化(対象文字セル As Variant) As Variant
    Fstr_英字大文字化 = UCase(対象文字セル)
End Function
Public Function Fstr_英字小文字化(対象文字セル As Variant) As Variant
    Fstr_英字小文字化 = LCase(対象文字セル)
End Function
Public Function Fdate_営業日日数(開始日 As Variant, 終了日 As Variant, 祭日を書いたセル範囲 As Variant) As Variant
    Fdate_営業日日数 = Application.WorksheetFunction.NetworkDays(開始日, 終了日, 祭日を書いたセル範囲)
End Function
Public Function Fdate_営業日(開始日 As Variant, 日数 As Variant, 祭日の日付を書いたセル範囲 As Variant) As Variant
    Fdate_営業日 = Application.WorksheetFunction.WorkDay(開始日, 日数, 祭日の日付を書いたセル範囲)
End Function
Public Function Fstr_右文字列(文字列 As Variant, 右からの文字数 As Variant) As Variant
    Fstr_右文字列 = Right(文字列, 右からの文字数)
End Function
Public Function Fstr_一致(検索値 As Variant, 検索範囲 As Variant, 照合の種類 As Variant) As Variant
    Fstr_一致 = Application.WorksheetFunction.Match(検索値, 検索範囲, 照合の種類)
End Function
Public Function Ffin_ローン返済額の利子相当分の累計額(月利 As Variant, ローン契約月数 As Variant, 借入額 As Variant, 第n月 As Variant, 第m月 As Variant, 期末払いなら0期首払いなら1 As Variant) As Variant
    Ffin_ローン返済額の利子相当分の累計額 = Application.WorksheetFunction.CumIPmt(月利, ローン契約月数, 借入額, 第n月, 第m月, 期末払いなら0期首払いなら1)
End Function
Public Function Ffin_ローン返済額の元金相当分の累計額(月利 As Variant, ローン契約月数 As Variant, 借入額 As Variant, 第n月 As Variant, 第m月 As Variant, 期末払いなら0期首払いなら1 As Variant) As Variant
    Ffin_ローン返済額の元金相当分の累計額 = Application.WorksheetFunction.CumPrinc(月利, ローン契約月数, 借入額, 第n月, 第m月, 期末払いなら0期首払いなら1)
End Function
Public Function Fmath_ランダム() As Variant
    Fmath_ランダム = Rnd
End Function
Public Function Fmath_ラジアン(角度セル As Variant) As Variant
    Fmath_ラジアン = Application.WorksheetFunction.Radians(角度セル)
End Function
Public Function Fmath_もっとも近い偶数(対象数値セル As Variant) As Variant
    Fmath_もっとも近い偶数 = Application.WorksheetFunction.Even(対象数値セル)
End Function
Public Function Fmath_もっとも近い奇数(対象数値セル As Variant) As Variant
    Fmath_もっとも近い奇数 = Application.WorksheetFunction.Odd(対象数値セル)
End Function
Public Function Fexcel_もし平均(検索範囲 As Variant, 比較値 As Variant, 平均範囲 As Variant) As Variant
    Fexcel_もし平均 = Application.WorksheetFunction.AverageIf(検索範囲, 比較値, 平均範囲)
End Function
Public Function Fexcel_もし文字列でない(対象セル As Variant) As Variant
    Fexcel_もし文字列でない = Application.WorksheetFunction.IsNonText(対象セル)
End Function
Public Function Fexcel_もし文字列(対象セル As Variant) As Variant
    Fexcel_もし文字列 = Application.WorksheetFunction.IsText(対象セル)
End Function
Public Function Fexcel_もし数値(対象セル As Variant) As Variant
    Fexcel_もし数値 = Application.WorksheetFunction.IsNumber(対象セル)
End Function
Public Function Fexcel_もし合計(検索範囲 As Variant, 比較値 As Variant, 合計範囲 As Variant) As Variant
    Fexcel_もし合計 = Application.WorksheetFunction.SumIf(検索範囲, 比較値, 合計範囲)
End Function
Public Function Fexcel_もし件数(検索範囲 As Variant, 比較値 As Variant) As Variant
    Fexcel_もし件数 = Application.WorksheetFunction.CountIf(検索範囲, 比較値)
End Function
Public Function Fexcel_もし偶数(対象セル As Variant) As Variant
    Fexcel_もし偶数 = Application.WorksheetFunction.IsEven(対象セル)
End Function
Public Function Fexcel_もし空白(対象セル As Variant) As Variant
    Fexcel_もし空白 = IsEmpty(対象セル)
End Function
Public Function Fexcel_もし奇数(対象セル As Variant) As Variant
    Fexcel_もし奇数 = Application.WorksheetFunction.IsOdd(対象セル)
End Function
Public Function Fexcel_もしエラー(対象セル As Variant) As Variant
    Fexcel_もしエラー = Application.WorksheetFunction.IsError(対象セル)
End Function
Public Function Fexcel_もしNA(対象セル As Variant) As Variant
    Fexcel_もしNA = Application.WorksheetFunction.IsNA(対象セル)
End Function
Public Function Fexcel_もし(条件式 As Variant, 真値 As Variant, 偽値 As Variant) As Variant
    Fexcel_もし = IIf(条件式, 真値, 偽値)
End Function
Public Function Fexcel_または(論理条件1 As Variant, 論理条件2 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    Fexcel_または = Application.WorksheetFunction.Or(論理条件1, 論理条件2, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function Fmath_べき乗(元の数セル As Variant, べき乗数セル As Variant) As Variant
    Fmath_べき乗 = Application.WorksheetFunction.Power(元の数セル, べき乗数セル)
End Function
Public Function Fstr_ふりがな表示(対象文字セル As Variant) As Variant
    Fstr_ふりがな表示 = Application.WorksheetFunction.Phonetic(対象文字セル)
End Function
Public Function Fmath_タンジェント(数値セル As Variant) As Variant
    Fmath_タンジェント = Tan(数値セル)
End Function
Public Function Fexcel_セル件数(検索範囲 As Variant) As Variant
    Fexcel_セル件数 = Application.WorksheetFunction.CountA(検索範囲)
End Function
Public Function Fmath_サイン(数値セル As Variant) As Variant
    Fmath_サイン = Sin(数値セル)
End Function
Public Function Fexcel_かつ(論理条件1 As Variant, 論理条件2 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    Fexcel_かつ = Application.WorksheetFunction.And(論理条件1, 論理条件2, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function Fmath_アークタンジェント(x座標 As Variant, y座標 As Variant) As Variant
    Fmath_アークタンジェント = Application.WorksheetFunction.Atan2(x座標, y座標)
End Function
Public Function Fmath_アークサイン(元のサインの数値 As Variant) As Variant
    Fmath_アークサイン = Application.WorksheetFunction.Asin(元のサインの数値)
End Function
Public Function Fmath_アークコサイン(元のコサインの数値 As Variant) As Variant
    Fmath_アークコサイン = Application.WorksheetFunction.Acos(元のコサインの数値)
End Function
Public Function Fmath_nPr何通り(総数セル As Variant, 取り出す分セル As Variant) As Variant
    Fmath_nPr何通り = Application.WorksheetFunction.Permut(総数セル, 取り出す分セル)
End Function

Public Function Fmath_nCr何通り(総数セル As Variant, 取り出す分セル As Variant) As Variant
    Fmath_nCr何通り = Application.WorksheetFunction.Combin(総数セル, 取り出す分セル)
End Function

Public Function Fdate_日付化(日付シリアル As Variant) As Date
    Fdate_日付化 = CDate(日付シリアル)
End Function

Public Function Fcolor_色(Optional 赤 As Variant = 255, Optional 緑 As Variant = 255, Optional 青 As Variant = 255) As Variant
    Fcolor_色 = RGB(赤, 緑, 青)
End Function

Public Function Fcolor_色インデックスからRGB色へ変換(idx As カラーインデックスパターン) As Variant
    Select Case idx
    Case 1
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 0, 0)
    Case 2
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 255, 255)
    Case 3
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 0, 0)
    Case 4
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 255, 0)
    Case 5
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 0, 255)
    Case 6
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 255, 0)
    Case 7
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 0, 255)
    Case 8
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 255, 255)
    Case 9
        Fcolor_色インデックスからRGB色へ変換 = RGB(128, 0, 0)
    Case 10
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 128, 0)
    Case 11
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 0, 128)
    Case 12
        Fcolor_色インデックスからRGB色へ変換 = RGB(128, 128, 0)
    Case 13
        Fcolor_色インデックスからRGB色へ変換 = RGB(128, 0, 128)
    Case 14
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 128, 128)
    Case 15
        Fcolor_色インデックスからRGB色へ変換 = RGB(192, 192, 192)
    Case 16
        Fcolor_色インデックスからRGB色へ変換 = RGB(128, 128, 128)
    Case 17
        Fcolor_色インデックスからRGB色へ変換 = RGB(153, 153, 255)
    Case 18
        Fcolor_色インデックスからRGB色へ変換 = RGB(153, 51, 102)
    Case 19
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 255, 204)
    Case 20
        Fcolor_色インデックスからRGB色へ変換 = RGB(204, 255, 255)
    Case 21
        Fcolor_色インデックスからRGB色へ変換 = RGB(102, 0, 102)
    Case 22
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 128, 128)
    Case 23
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 102, 204)
    Case 24
        Fcolor_色インデックスからRGB色へ変換 = RGB(204, 204, 255)
    Case 25
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 0, 128)
    Case 26
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 0, 255)
    Case 27
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 255, 0)
    Case 28
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 255, 255)
    Case 29
        Fcolor_色インデックスからRGB色へ変換 = RGB(128, 0, 128)
    Case 30
        Fcolor_色インデックスからRGB色へ変換 = RGB(128, 0, 0)
    Case 31
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 128, 128)
    Case 32
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 0, 255)
    Case 33
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 204, 255)
    Case 34
        Fcolor_色インデックスからRGB色へ変換 = RGB(204, 255, 255)
    Case 35
        Fcolor_色インデックスからRGB色へ変換 = RGB(204, 255, 204)
    Case 36
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 255, 153)
    Case 37
        Fcolor_色インデックスからRGB色へ変換 = RGB(153, 204, 255)
    Case 38
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 153, 204)
    Case 39
        Fcolor_色インデックスからRGB色へ変換 = RGB(204, 153, 255)
    Case 40
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 204, 153)
    Case 41
        Fcolor_色インデックスからRGB色へ変換 = RGB(51, 102, 255)
    Case 42
        Fcolor_色インデックスからRGB色へ変換 = RGB(51, 204, 204)
    Case 43
        Fcolor_色インデックスからRGB色へ変換 = RGB(153, 204, 0)
    Case 44
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 204, 0)
    Case 45
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 153, 0)
    Case 46
        Fcolor_色インデックスからRGB色へ変換 = RGB(255, 102, 0)
    Case 47
        Fcolor_色インデックスからRGB色へ変換 = RGB(102, 102, 153)
    Case 48
        Fcolor_色インデックスからRGB色へ変換 = RGB(150, 150, 150)
    Case 49
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 51, 102)
    Case 50
        Fcolor_色インデックスからRGB色へ変換 = RGB(51, 153, 102)
    Case 51
        Fcolor_色インデックスからRGB色へ変換 = RGB(0, 51, 0)
    Case 52
        Fcolor_色インデックスからRGB色へ変換 = RGB(51, 51, 0)
    Case 53
        Fcolor_色インデックスからRGB色へ変換 = RGB(153, 51, 0)
    Case 54
        Fcolor_色インデックスからRGB色へ変換 = RGB(153, 51, 102)
    Case 55
        Fcolor_色インデックスからRGB色へ変換 = RGB(51, 51, 153)
    Case 56
        Fcolor_色インデックスからRGB色へ変換 = RGB(51, 51, 51)
    End Select

End Function

Public Function Fcolor_色の三原色を取得(色 As Long, ByRef 赤 As Long, ByRef 緑 As Long, ByRef 青 As Long)
    赤 = 色 Mod 256
    緑 = Int(色 / 256) Mod 256
    青 = Int(色 / 256 / 256)
End Function

Public Function Fdate_今日()
    Fdate_今日 = Int(F_今())
End Function

Public Function Fdate_今日の日付()
    Fdate_今日の日付 = Trim(Fdate_日付化(Fdate_今日()))
End Function

Public Function Fdate_今の日付()
    Fdate_今の日付 = Fdate_日付化(Fdate_今())
End Function







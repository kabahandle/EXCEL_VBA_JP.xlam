Attribute VB_Name = "F_関数"
Option Explicit

Public Function F_カスタム1(合計範囲 As Variant) As Variant
    F_カスタム1 = F_切り捨て2(F_合計(合計範囲), 2)
End Function
Public Function F_カスタム2(対象セル As Variant) As Variant
    F_カスタム2 = F_商(対象セル, 12)
End Function
'数学のアークコサイン（arccos）を度で返す関数です。
Public Function F_アークコサイン度(cos値 As Variant) As Variant
    F_アークコサイン度 = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Acos(cos値))
End Function

'数学のアークサイン（arcsin）を度で返す関数です。
Public Function F_アークサイン度(sin値 As Variant) As Variant
    F_アークサイン度 = Application.WorksheetFunction.Degrees(Application.WorksheetFunction.Asin(sin値))
End Function

'数学のアークタンジェント（arctan）を、度で返す関数です。
Public Function F_アークタンジェント度(tan値 As Variant) As Variant
    F_アークタンジェント度 = Application.WorksheetFunction.Degrees(Atn(tan値))
End Function

'数学のコサイン（cos）を度数から引く関数です。
Public Function F_コサイン度(度 As Variant) As Variant
    F_コサイン度 = Cos(Application.WorksheetFunction.Radians(度))
End Function

'数学のサイン（sin）を度数から引く関数です。
Public Function F_サイン度(度 As Variant) As Variant
    F_サイン度 = Sin(Application.WorksheetFunction.Radians(度))
End Function

'数学のタンジェント度（tan）を度数から引く関数です。
Public Function F_タンジェント度(度 As Variant) As Variant
    F_タンジェント度 = Tan(Application.WorksheetFunction.Radians(度))
End Function

'2進数を10進数に変換します。
Public Function F_二進数から十進数(二進数 As Variant) As Variant

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

    F_二進数から十進数 = 十進数計算用

End Function

' n ÷ m の式の余りを求めます。
Public Function F_余り(割られる数n As Variant, 割る数m As Variant) As Variant
    F_余り = 割られる数n Mod 割る数m
End Function

'16進数を10進数に変換します
Public Function F_十六進数から十進数(十六進数 As Variant) As Variant
    F_十六進数から十進数 = Val("&H" & 十六進数)
End Function

'10進数を2進数に変換します
Public Function F_十進数から二進数(十進数 As Variant, Optional 桁数 As Long = 8) As String
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

    F_十進数から二進数 = Format(二進数計算用, padding)
End Function

'10進数を16進数に変換します
Public Function F_十進数から十六進数(十進数 As Variant, Optional 桁数 As Long = 4) As Variant
    F_十進数から十六進数 = F_十六進数パディング(Hex(十進数), "0", 桁数)
End Function
'機能：指定文字埋め関数
'引数：str　：変換前の文字列
'　　　chr  ：埋める文字(１文字目のみ使用)
'　　　digit：桁数
'戻値：指定文字埋め後の文字列
Private Function F_十六進数パディング(ByVal str As String, _
                     ByVal char As String, _
                     ByVal digit As Long) As String
  Dim tmp As String
  tmp = str
  If Len(str) < digit And Len(char) > 0 Then
    tmp = Right(String(digit, char) & str, digit)
  End If
  F_十六進数パディング = tmp
End Function

'正規表現の置換パターン文字列を指定して、正規表現置換します。
Public Function F_正規表現置換(検索対象 As Variant, 置換パターン文字列 As Variant, 置換後の文字列 As Variant, Optional 大文字小文字無視 As Boolean = False, Optional 最初の一致時のみ置換 As Boolean = False)
    r_RegExp.Pattern = 置換パターン文字列
    r_RegExp.IgnoreCase = 大文字小文字無視
    r_RegExp.Global = Not 最初の一致時のみ置換
    If (IsObject(検索対象)) Then
        F_正規表現置換 = r_RegExp.Replace(検索対象.Value2, 置換後の文字列)
    Else
        F_正規表現置換 = r_RegExp.Replace(検索対象, 置換後の文字列)
    End If
End Function


Public Function F_曜日(日付セル As Variant, 種類1から3 As Variant) As Variant
    F_曜日 = Application.WorksheetFunction.Weekday(日付セル, 種類1から3)
End Function
Public Function F_平方根(数値セル As Variant) As Variant
    F_平方根 = Sqr(数値セル)
End Function
Public Function F_平均(平均範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_平均 = Application.WorksheetFunction.Average(平均範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_文字列長(文字列 As Variant) As Variant
    F_文字列長 = Len(文字列)
End Function
Public Function F_文字置換(置換対象セル As Variant, 置換対象文字列 As Variant, 置換後文字列 As Variant) As Variant
    F_文字置換 = Application.WorksheetFunction.Substitute(置換対象セル, 置換対象文字列, 置換後文字列)
End Function
Public Function F_分散(セル範囲 As Variant) As Variant
    F_分散 = Application.WorksheetFunction.VarP(セル範囲)
End Function
Public Function F_分(日時セル As Variant) As Variant
    F_分 = Minute(日時セル)
End Function
Public Function F_不偏分散(セル範囲 As Variant) As Variant
    F_不偏分散 = Application.WorksheetFunction.Var(セル範囲)
End Function
Public Function F_不偏標準偏差(セル範囲 As Variant) As Variant
    F_不偏標準偏差 = Application.WorksheetFunction.StDev(セル範囲)
End Function
Public Function F_秒(日時セル As Variant) As Variant
    F_秒 = Second(日時セル)
End Function
Public Function F_標準偏差(セル範囲 As Variant) As Variant
    F_標準偏差 = Application.WorksheetFunction.StDevP(セル範囲)
End Function
Public Function F_倍数切り上げ(数値 As Variant, 倍数基準 As Variant) As Variant
    F_倍数切り上げ = Application.WorksheetFunction.Ceiling(数値, 倍数基準)
End Function
Public Function F_倍数切り捨て(数値 As Variant, 倍数基準 As Variant) As Variant
    F_倍数切り捨て = Application.WorksheetFunction.Floor(数値, 倍数基準)
End Function
Public Function F_年(日付セル As Variant) As Variant
    F_年 = Year(日付セル)
End Function
Public Function F_日付変換(年 As Variant, 月 As Variant, 日 As Variant) As Variant
    F_日付変換 = DateSerial(年, 月, 日)
End Function
Public Function F_日付の差(比較単位 As Variant, 日付セル1 As Variant, 日付セル2 As Variant) As Variant
    F_日付の差 = DateDiff(比較単位, 日付セル1, 日付セル2)
End Function
Public Function F_日(日付セル As Variant) As Variant
    F_日 = Day(日付セル)
End Function
Public Function F_中央値(セル範囲 As Variant) As Variant
    F_中央値 = Application.WorksheetFunction.Median(セル範囲)
End Function
Public Function F_大きい方から何番目かの値(セル範囲 As Variant, 順位 As Variant) As Variant
    F_大きい方から何番目かの値 = Application.WorksheetFunction.Large(セル範囲, 順位)
End Function
Public Function F_対数(元の数値 As Variant) As Variant
    F_対数 = Log(元の数値)
End Function
Public Function F_全角文字を半角化(対象文字セル As Variant) As Variant
    F_全角文字を半角化 = Application.WorksheetFunction.Asc(対象文字セル)
End Function
Public Function F_絶対値(数値セル As Variant) As Variant
    F_絶対値 = Abs(数値セル)
End Function
Public Function F_切り上げ(数値 As Variant, 切り上げる桁数 As Variant) As Variant
    F_切り上げ = Application.WorksheetFunction.RoundUp(数値, 切り上げる桁数)
End Function
Public Function F_切り捨て2(数値 As Variant, 切り捨てる桁数｡ As Variant) As Variant
    F_切り捨て2 = Application.WorksheetFunction.RoundDown(数値, 切り捨てる桁数｡)
End Function
Public Function F_切り捨て(数値 As Variant) As Variant
    F_切り捨て = Int(数値)
End Function
Public Function F_数値間ランダム(開始値 As Variant, 終了値 As Variant) As Variant
    F_数値間ランダム = Application.WorksheetFunction.RandBetween(開始値, 終了値)
End Function
Public Function F_数字をローマ数字化(対象セル As Variant) As Variant
    F_数字をローマ数字化 = Application.WorksheetFunction.Roman(対象セル)
End Function
Public Function F_数ヶ月後の月末(開始日 As Variant, 月 As Variant) As Variant
    F_数ヶ月後の月末 = Application.WorksheetFunction.EoMonth(開始日, 月)
End Function
Public Function F_数ヶ月後(開始日 As Variant, 月 As Variant) As Variant
    F_数ヶ月後 = Application.WorksheetFunction.EDate(開始日, 月)
End Function
Public Function F_常用対数(元の数値 As Variant) As Variant
    F_常用対数 = Application.WorksheetFunction.Log10(元の数値)
End Function
Public Function F_小さい方から何番目かの値(セル範囲 As Variant, 順位 As Variant) As Variant
    F_小さい方から何番目かの値 = Application.WorksheetFunction.Small(セル範囲, 順位)
End Function
Public Function F_商(対象数値セル As Variant, 割る数 As Variant) As Variant
    F_商 = Application.WorksheetFunction.Quotient(対象数値セル, 割る数)
End Function
Public Function F_順位(順位調査セル As Variant, セル範囲 As Variant, 順序フラグ As Variant) As Variant
    F_順位 = Application.WorksheetFunction.Rank(順位調査セル, セル範囲, 順序フラグ)
End Function
Public Function F_縦表引(検索値 As Variant, 検索範囲 As Variant, 列番号 As Variant, Optional オプション1 As Variant) As Variant
    F_縦表引 = Application.WorksheetFunction.VLookup(検索値, 検索範囲, 列番号, オプション1)
End Function
Public Function F_自然対数の底eのべき乗(べきとなる数 As Variant) As Variant
    F_自然対数の底eのべき乗 = Exp(べきとなる数)
End Function
Public Function F_自然対数(元の数値 As Variant) As Variant
    F_自然対数 = Application.WorksheetFunction.Ln(元の数値)
End Function
Public Function F_時間変換(時 As Variant, 分 As Variant, 秒 As Variant) As Variant
    F_時間変換 = TimeSerial(時, 分, 秒)
End Function
Public Function F_時(日時セル As Variant) As Variant
    F_時 = Hour(日時セル)
End Function
Public Function F_四捨五入(数値 As Variant, 四捨五入する桁数 As Variant) As Variant
    F_四捨五入 = Application.WorksheetFunction.Round(数値, 四捨五入する桁数)
End Function
Public Function F_最頻値(セル範囲 As Variant) As Variant
    F_最頻値 = Application.WorksheetFunction.Mode(セル範囲)
End Function
Public Function F_最大公約数(数値範囲 As Variant) As Variant
    F_最大公約数 = Application.WorksheetFunction.Gcd(数値範囲)
End Function
Public Function F_最大(検索範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_最大 = Application.WorksheetFunction.max(検索範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_最小公倍数(数値範囲 As Variant) As Variant
    F_最小公倍数 = Application.WorksheetFunction.Lcm(数値範囲)
End Function
Public Function F_最小(検索範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_最小 = Application.WorksheetFunction.Min(検索範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_左文字列(文字列 As Variant, 左からの文字数 As Variant) As Variant
    F_左文字列 = Left(文字列, 左からの文字数)
End Function
Public Function F_左右空白文字削除(対象文字セル As Variant) As Variant
    F_左右空白文字削除 = Application.WorksheetFunction.Trim(対象文字セル)
End Function
Public Function F_今() As Variant
    F_今 = Now
End Function
Public Function F_合計(合計範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_合計 = Application.WorksheetFunction.Sum(合計範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_件数(カウント範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_件数 = Application.WorksheetFunction.Count(カウント範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_月々積立貯蓄払込額(月利 As Variant, 積立月数 As Variant, 初期金額 As Variant, 目的の積立額 As Variant) As Variant
    F_月々積立貯蓄払込額 = Application.WorksheetFunction.Pmt(月利, 積立月数, 初期金額, 目的の積立額)
End Function
Public Function F_月々ローン返済額中の元金返済額(月利 As Variant, 求めるものは何月目か As Variant, 返済月数 As Variant, 借入金額 As Variant, 最後に残る金額 As Variant) As Variant
    F_月々ローン返済額中の元金返済額 = Application.WorksheetFunction.PPmt(月利, 求めるものは何月目か, 返済月数, 借入金額, 最後に残る金額)
End Function
Public Function F_月々ローン返済額中の金利分額(月利 As Variant, 求めるものは何月目か As Variant, 返済月数 As Variant, 借入額 As Variant, 最後に残る金額 As Variant) As Variant
    F_月々ローン返済額中の金利分額 = Application.WorksheetFunction.IPmt(月利, 求めるものは何月目か, 返済月数, 借入額, 最後に残る金額)
End Function
Public Function F_月々ローン返済額(月利 As Variant, 返済月数 As Variant, 借入額 As Variant, 最後に残る金額 As Variant) As Variant
    F_月々ローン返済額 = Application.WorksheetFunction.Pmt(月利, 返済月数, 借入額, 最後に残る金額)
End Function
Public Function F_月(日付セル As Variant) As Variant
    F_月 = Month(日付セル)
End Function
Public Function F_繰り返し表示(対象文字列 As Variant, 繰り返し回数 As Variant) As Variant
    F_繰り返し表示 = Application.WorksheetFunction.Rept(対象文字列, 繰り返し回数)
End Function
Public Function F_間文字列(文字列 As Variant, 先頭文字番号 As Variant, 抜き出し文字数 As Variant) As Variant
    F_間文字列 = Mid(文字列, 先頭文字番号, 抜き出し文字数)
End Function
Public Function F_角度(ラジアンセル As Variant) As Variant
    F_角度 = Application.WorksheetFunction.Degrees(ラジアンセル)
End Function
Public Function F_階乗(元の数値 As Variant) As Variant
    F_階乗 = Application.WorksheetFunction.Fact(元の数値)
End Function
Public Function F_何周目の日付(日付 As Variant, フラグ1または2 As Variant) As Variant
    F_何周目の日付 = Application.WorksheetFunction.WeekNum(日付, フラグ1または2)
End Function
Public Function F_横表引(検索値 As Variant, 検索範囲 As Variant, 行番号 As Variant, Optional オプション1 As Variant) As Variant
    F_横表引 = Application.WorksheetFunction.HLookup(検索値, 検索範囲, 行番号, オプション1)
End Function
Public Function F_円周率() As Variant
    F_円周率 = Application.WorksheetFunction.Pi
End Function
Public Function F_英単語の先頭文字を大文字化(英単語を含むセル As Variant) As Variant
    F_英単語の先頭文字を大文字化 = Application.WorksheetFunction.Proper(英単語を含むセル)
End Function
Public Function F_英字大文字化(対象文字セル As Variant) As Variant
    F_英字大文字化 = UCase(対象文字セル)
End Function
Public Function F_英字小文字化(対象文字セル As Variant) As Variant
    F_英字小文字化 = LCase(対象文字セル)
End Function
Public Function F_営業日日数(開始日 As Variant, 終了日 As Variant, 祭日を書いたセル範囲 As Variant) As Variant
    F_営業日日数 = Application.WorksheetFunction.NetworkDays(開始日, 終了日, 祭日を書いたセル範囲)
End Function
Public Function F_営業日(開始日 As Variant, 日数 As Variant, 祭日の日付を書いたセル範囲 As Variant) As Variant
    F_営業日 = Application.WorksheetFunction.WorkDay(開始日, 日数, 祭日の日付を書いたセル範囲)
End Function
Public Function F_右文字列(文字列 As Variant, 右からの文字数 As Variant) As Variant
    F_右文字列 = Right(文字列, 右からの文字数)
End Function
Public Function F_一致(検索値 As Variant, 検索範囲 As Variant, 照合の種類 As Variant) As Variant
    F_一致 = Application.WorksheetFunction.Match(検索値, 検索範囲, 照合の種類)
End Function
Public Function F_ローン返済額の利子相当分の累計額(月利 As Variant, ローン契約月数 As Variant, 借入額 As Variant, 第n月 As Variant, 第m月 As Variant, 期末払いなら0期首払いなら1 As Variant) As Variant
    F_ローン返済額の利子相当分の累計額 = Application.WorksheetFunction.CumIPmt(月利, ローン契約月数, 借入額, 第n月, 第m月, 期末払いなら0期首払いなら1)
End Function
Public Function F_ローン返済額の元金相当分の累計額(月利 As Variant, ローン契約月数 As Variant, 借入額 As Variant, 第n月 As Variant, 第m月 As Variant, 期末払いなら0期首払いなら1 As Variant) As Variant
    F_ローン返済額の元金相当分の累計額 = Application.WorksheetFunction.CumPrinc(月利, ローン契約月数, 借入額, 第n月, 第m月, 期末払いなら0期首払いなら1)
End Function
Public Function F_ろーんへんさいがくのりしぶんのるけいいがく(月利 As Variant, ローン契約月数 As Variant, 借入額 As Variant, 第n月 As Variant, 第m月 As Variant, 期末払いなら0期首払いなら1 As Variant) As Variant
    F_ろーんへんさいがくのりしぶんのるけいいがく = Application.WorksheetFunction.CumIPmt(月利, ローン契約月数, 借入額, 第n月, 第m月, 期末払いなら0期首払いなら1)
End Function
Public Function F_ろーんへんさいがくのがんきんぶんのるいせきがく(月利 As Variant, ローン契約月数 As Variant, 借入額 As Variant, 第n月 As Variant, 第m月 As Variant, 期末払いなら0期首払いなら1 As Variant) As Variant
    F_ろーんへんさいがくのがんきんぶんのるいせきがく = Application.WorksheetFunction.CumPrinc(月利, ローン契約月数, 借入額, 第n月, 第m月, 期末払いなら0期首払いなら1)
End Function
Public Function F_ランダム() As Variant
    F_ランダム = Rnd
End Function
Public Function F_ラジアン(角度セル As Variant) As Variant
    F_ラジアン = Application.WorksheetFunction.Radians(角度セル)
End Function
Public Function F_よこひょうびき(検索値 As Variant, 検索範囲 As Variant, 行番号 As Variant, 検索の型 As Variant) As Variant
    F_よこひょうびき = Application.WorksheetFunction.HLookup(検索値, 検索範囲, 行番号, 検索の型)
End Function
Public Function F_ようび(日付セル As Variant, 種類1から3 As Variant) As Variant
    F_ようび = Application.WorksheetFunction.Weekday(日付セル, 種類1から3)
End Function
Public Function F_もっとも近い偶数(対象数値セル As Variant) As Variant
    F_もっとも近い偶数 = Application.WorksheetFunction.Even(対象数値セル)
End Function
Public Function F_もっとも近い奇数(対象数値セル As Variant) As Variant
    F_もっとも近い奇数 = Application.WorksheetFunction.Odd(対象数値セル)
End Function
Public Function F_もっともちかいぐうすう(対象数値セル As Variant) As Variant
    F_もっともちかいぐうすう = Application.WorksheetFunction.Even(対象数値セル)
End Function
Public Function F_もっともちかいきすう(対象数値セル As Variant) As Variant
    F_もっともちかいきすう = Application.WorksheetFunction.Odd(対象数値セル)
End Function
Public Function F_もし平均(検索範囲 As Variant, 比較値 As Variant, 平均範囲 As Variant) As Variant
    F_もし平均 = Application.WorksheetFunction.AverageIf(検索範囲, 比較値, 平均範囲)
End Function
Public Function F_もし文字列でない(対象セル As Variant) As Variant
    F_もし文字列でない = Application.WorksheetFunction.IsNonText(対象セル)
End Function
Public Function F_もし文字列(対象セル As Variant) As Variant
    F_もし文字列 = Application.WorksheetFunction.IsText(対象セル)
End Function
Public Function F_もし数値(対象セル As Variant) As Variant
    F_もし数値 = Application.WorksheetFunction.IsNumber(対象セル)
End Function
Public Function F_もし合計(検索範囲 As Variant, 比較値 As Variant, 合計範囲 As Variant) As Variant
    F_もし合計 = Application.WorksheetFunction.SumIf(検索範囲, 比較値, 合計範囲)
End Function
Public Function F_もし件数(検索範囲 As Variant, 比較値 As Variant) As Variant
    F_もし件数 = Application.WorksheetFunction.CountIf(検索範囲, 比較値)
End Function
Public Function F_もし偶数(対象セル As Variant) As Variant
    F_もし偶数 = Application.WorksheetFunction.IsEven(対象セル)
End Function
Public Function F_もし空白(対象セル As Variant) As Variant
    F_もし空白 = IsEmpty(対象セル)
End Function
Public Function F_もし奇数(対象セル As Variant) As Variant
    F_もし奇数 = Application.WorksheetFunction.IsOdd(対象セル)
End Function
Public Function F_もじれつながさ(文字列 As Variant) As Variant
    F_もじれつながさ = Len(文字列)
End Function
Public Function F_もしもじれつでない(対象セル As Variant) As Variant
    F_もしもじれつでない = Application.WorksheetFunction.IsNonText(対象セル)
End Function
Public Function F_もしもじれつ(対象セル As Variant) As Variant
    F_もしもじれつ = Application.WorksheetFunction.IsText(対象セル)
End Function
Public Function F_もしへいきん(検索範囲 As Variant, 比較値 As Variant, 平均範囲 As Variant) As Variant
    F_もしへいきん = Application.WorksheetFunction.AVERGEIF(検索範囲, 比較値, 平均範囲)
End Function
Public Function F_もしのっとあさいんど(対象セル As Variant) As Variant
    F_もしのっとあさいんど = Application.WorksheetFunction.IsNA(対象セル)
End Function
Public Function F_もじちかん(置換対象セル As Variant, 置換対象文字列 As Variant, 置換後文字列 As Variant) As Variant
    F_もじちかん = Application.WorksheetFunction.Substitute(置換対象セル, 置換対象文字列, 置換後文字列)
End Function
Public Function F_もしすうち(対象セル As Variant) As Variant
    F_もしすうち = Application.WorksheetFunction.IsNumber(対象セル)
End Function
Public Function F_もしごうけい(検索範囲 As Variant, 比較値 As Variant, 合計範囲 As Variant) As Variant
    F_もしごうけい = Application.WorksheetFunction.SumIf(検索範囲, 比較値, 合計範囲)
End Function
Public Function F_もしけんすう(検索範囲 As Variant, 比較値 As Variant) As Variant
    F_もしけんすう = Application.WorksheetFunction.CountIf(検索範囲, 比較値)
End Function
Public Function F_もしくうはく(対象セル As Variant) As Variant
    F_もしくうはく = IsEmpty(対象セル)
End Function
Public Function F_もしぐうすう(対象セル As Variant) As Variant
    F_もしぐうすう = Application.WorksheetFunction.IsEven(対象セル)
End Function
Public Function F_もしきすう(対象セル As Variant) As Variant
    F_もしきすう = Application.WorksheetFunction.IsOdd(対象セル)
End Function
Public Function F_もしエラー(対象セル As Variant) As Variant
    F_もしエラー = Application.WorksheetFunction.IsError(対象セル)
End Function
Public Function F_もしNA(対象セル As Variant) As Variant
    F_もしNA = Application.WorksheetFunction.IsNA(対象セル)
End Function
Public Function F_もし(条件式 As Variant, 真値 As Variant, 偽値 As Variant) As Variant
    F_もし = IIf(条件式, 真値, 偽値)
End Function
Public Function F_みぎもじれつ(文字列 As Variant, 右からの文字数 As Variant) As Variant
    F_みぎもじれつ = Right(文字列, 右からの文字数)
End Function
Public Function F_または(論理条件1 As Variant, 論理条件2 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_または = Application.WorksheetFunction.Or(論理条件1, 論理条件2, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_べき乗(元の数セル As Variant, べき乗数セル As Variant) As Variant
    F_べき乗 = Application.WorksheetFunction.Power(元の数セル, べき乗数セル)
End Function
Public Function F_へいほうこん(数値セル As Variant) As Variant
    F_へいほうこん = Sqr(数値セル)
End Function
Public Function F_へいきん(平均範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_へいきん = Application.WorksheetFunction.Average(平均範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_ぶんさん(セル範囲 As Variant) As Variant
    F_ぶんさん = Application.WorksheetFunction.VarP(セル範囲)
End Function
Public Function F_ふん(日時セル As Variant) As Variant
    F_ふん = Minute(日時セル)
End Function
Public Function F_ふりがな表示(対象文字セル As Variant) As Variant
    F_ふりがな表示 = Application.WorksheetFunction.Phonetic(対象文字セル)
End Function
Public Function F_ふりがな(対象文字セル As Variant) As Variant
    F_ふりがな = Application.WorksheetFunction.Phonetic(対象文字セル)
End Function
Public Function F_ふへんぶんさん(セル範囲 As Variant) As Variant
    F_ふへんぶんさん = Application.WorksheetFunction.Var(セル範囲)
End Function
Public Function F_ふへんひょうじゅんへんさ(セル範囲 As Variant) As Variant
    F_ふへんひょうじゅんへんさ = Application.WorksheetFunction.StDev(セル範囲)
End Function
Public Function F_ひょうじゅんへんさ(セル範囲 As Variant) As Variant
    F_ひょうじゅんへんさ = Application.WorksheetFunction.StDevP(セル範囲)
End Function
Public Function F_びょう(日時セル As Variant) As Variant
    F_びょう = Second(日時セル)
End Function
Public Function F_ひづけへんかん(年 As Variant, 月 As Variant, 日 As Variant) As Variant
    F_ひづけへんかん = DateSerial(年, 月, 日)
End Function
Public Function F_ひだりもじれつ(文字列 As Variant, 左からの文字数 As Variant) As Variant
    F_ひだりもじれつ = Left(文字列, 左からの文字数)
End Function
Public Function F_ひ(日付セル As Variant) As Variant
    F_ひ = Day(日付セル)
End Function
Public Function F_ばいすうきりすて(数値 As Variant, 倍数基準 As Variant) As Variant
    F_ばいすうきりすて = Application.WorksheetFunction.Floor(数値, 倍数基準)
End Function
Public Function F_ばいすうきりあげ(数値 As Variant, 倍数基準 As Variant) As Variant
    F_ばいすうきりあげ = Application.WorksheetFunction.Ceiling(数値, 倍数基準)
End Function
Public Function F_ねん(日付セル As Variant) As Variant
    F_ねん = Year(日付セル)
End Function
Public Function F_なんしゅうめのひづけ(日付 As Variant, フラグ1または2 As Variant) As Variant
    F_なんしゅうめのひづけ = Application.WorksheetFunction.WeekNum(日付, フラグ1または2)
End Function
Public Function F_つきづきろーんへんさいがくへんさいがくのがんきんぶん(月利 As Variant, 求めるものは何月目か As Variant, 返済月数 As Variant, 借入金額 As Variant, 最後に残る金額 As Variant) As Variant
    F_つきづきろーんへんさいがくへんさいがくのがんきんぶん = Application.WorksheetFunction.PPmt(月利, 求めるものは何月目か, 返済月数, 借入金額, 最後に残る金額)
End Function
Public Function F_つきづきろーんへんさいがくのきんりぶん(月利 As Variant, 求めるものは何月目か As Variant, 返済月数 As Variant, 借入額 As Variant, 最後に残る金額 As Variant) As Variant
    F_つきづきろーんへんさいがくのきんりぶん = Application.WorksheetFunction.IPmt(月利, 求めるものは何月目か, 返済月数, 借入額, 最後に残る金額)
End Function
Public Function F_つきづきろーんへんさいがく(月利 As Variant, 返済月数 As Variant, 借入額 As Variant, 最後に残る金額 As Variant) As Variant
    F_つきづきろーんへんさいがく = Application.WorksheetFunction.Pmt(月利, 返済月数, 借入額, 最後に残る金額)
End Function
Public Function F_つきづきつみたてちょちくはらいこみがく(月利 As Variant, 積立月数 As Variant, 初期金額 As Variant, 目的の積立額 As Variant) As Variant
    F_つきづきつみたてちょちくはらいこみがく = Application.WorksheetFunction.Pmt(月利, 積立月数, 初期金額, 目的の積立額)
End Function
Public Function F_ちゅうおうち(セル範囲 As Variant) As Variant
    F_ちゅうおうち = Application.WorksheetFunction.Median(セル範囲)
End Function
Public Function F_ちいさいほうからなんばんめ(セル範囲 As Variant, 順位 As Variant) As Variant
    F_ちいさいほうからなんばんめ = Application.WorksheetFunction.Small(セル範囲, 順位)
End Function
Public Function F_タンジェント(数値セル As Variant) As Variant
    F_タンジェント = Tan(数値セル)
End Function
Public Function F_たてひょうびき(検索値 As Variant, 検索範囲 As Variant, 列番号 As Variant, 検索の型 As Variant) As Variant
    F_たてひょうびき = Application.WorksheetFunction.VLookup(検索値, 検索範囲, 列番号, 検索の型)
End Function
Public Function F_たいすう(元の数値 As Variant) As Variant
    F_たいすう = Log(元の数値)
End Function
Public Function F_ぜんかくをはんかくに(対象文字セル As Variant) As Variant
    F_ぜんかくをはんかくに = Application.WorksheetFunction.Asc(対象文字セル)
End Function
Public Function F_セル件数(検索範囲 As Variant) As Variant
    F_セル件数 = Application.WorksheetFunction.CountA(検索範囲)
End Function
Public Function F_せるけんすう(検索範囲 As Variant) As Variant
    F_せるけんすう = Application.WorksheetFunction.CountA(検索範囲)
End Function
Public Function F_ぜったいち(数値セル As Variant) As Variant
    F_ぜったいち = Abs(数値セル)
End Function
Public Function F_すうちかんらんだむ(開始値 As Variant, 終了値 As Variant) As Variant
    F_すうちかんらんだむ = Application.WorksheetFunction.RandBetween(開始値, 終了値)
End Function
Public Function F_すうじをろーますうじに(対象セル As Variant) As Variant
    F_すうじをろーますうじに = Application.WorksheetFunction.Roman(対象セル)
End Function
Public Function F_すうかげつごのげつまつ(開始日 As Variant, 月 As Variant) As Variant
    F_すうかげつごのげつまつ = Application.WorksheetFunction.EoMonth(開始日, 月)
End Function
Public Function F_すうかげつご(開始日 As Variant, 月 As Variant) As Variant
    F_すうかげつご = Application.WorksheetFunction.EDate(開始日, 月)
End Function
Public Function F_じょうようたいすう(元の数値 As Variant) As Variant
    F_じょうようたいすう = Application.WorksheetFunction.Log10(元の数値)
End Function
Public Function F_しょう(対象数値セル As Variant, 割る数 As Variant) As Variant
    F_しょう = Application.WorksheetFunction.Quotient(対象数値セル, 割る数)
End Function
Public Function F_じゅんい(順位調査セル As Variant, セル範囲 As Variant, 順序フラグ As Variant) As Variant
    F_じゅんい = Application.WorksheetFunction.Rank(順位調査セル, セル範囲, 順序フラグ)
End Function
Public Function F_しぜんたいすうのていのべきじょう(べきとなる数 As Variant) As Variant
    F_しぜんたいすうのていのべきじょう = Exp(べきとなる数)
End Function
Public Function F_しぜんたいすう(元の数値 As Variant) As Variant
    F_しぜんたいすう = Application.WorksheetFunction.Ln(元の数値)
End Function
Public Function F_ししゃごにゅう(数値 As Variant, 四捨五入する桁数 As Variant) As Variant
    F_ししゃごにゅう = Application.WorksheetFunction.Round(数値, 四捨五入する桁数)
End Function
Public Function F_じかんへんかん(時 As Variant, 分 As Variant, 秒 As Variant) As Variant
    F_じかんへんかん = TimeSerial(時, 分, 秒)
End Function
Public Function F_じかんのさ(比較単位 As Variant, 日付セル1 As Variant, 日付セル2 As Variant) As Variant
    F_じかんのさ = DateDiff(比較単位, 日付セル1, 日付セル2)
End Function
Public Function F_じ(日時セル As Variant) As Variant
    F_じ = Hour(日時セル)
End Function
Public Function F_さゆうくうはくさくじょ(対象文字セル As Variant) As Variant
    F_さゆうくうはくさくじょ = Application.WorksheetFunction.Trim(対象文字セル)
End Function
Public Function F_サイン(数値セル As Variant) As Variant
    F_サイン = Sin(数値セル)
End Function
Public Function F_さいひんち(セル範囲 As Variant) As Variant
    F_さいひんち = Application.WorksheetFunction.Mode(セル範囲)
End Function
Public Function F_さいだいこうばいすう(数値範囲 As Variant) As Variant
    F_さいだいこうばいすう = Application.WorksheetFunction.Gcd(数値範囲)
End Function
Public Function F_さいだい(検索範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_さいだい = Application.WorksheetFunction.max(検索範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_さいしょうこうばいすう(数値範囲 As Variant) As Variant
    F_さいしょうこうばいすう = Application.WorksheetFunction.Lcm(数値範囲)
End Function
Public Function F_さいしょう(検索範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_さいしょう = Application.WorksheetFunction.Min(検索範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_こもじ(対象文字セル As Variant) As Variant
    F_こもじ = LCase(対象文字セル)
End Function
Public Function F_コサイン(数値セル As Variant) As Variant
    F_コサイン = Cos(数値セル)
End Function
Public Function F_ごうけい(合計範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_ごうけい = Application.WorksheetFunction.Sum(合計範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_けんすう(カウント範囲 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_けんすう = Application.WorksheetFunction.Count(カウント範囲, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_げつ(日付セル As Variant) As Variant
    F_げつ = Month(日付セル)
End Function
Public Function F_くりかえしひょうじ(対象文字列 As Variant, 繰り返し回数 As Variant) As Variant
    F_くりかえしひょうじ = Application.WorksheetFunction.Rept(対象文字列, 繰り返し回数)
End Function
Public Function F_きりすて2(数値 As Variant, 切り捨てる桁数｡ As Variant) As Variant
    F_きりすて2 = Application.WorksheetFunction.RoundDown(数値, 切り捨てる桁数｡)
End Function
Public Function F_きりすて(数値 As Variant) As Variant
    F_きりすて = Int(数値)
End Function
Public Function F_きりあげ(数値 As Variant, 切り上げる桁数 As Variant) As Variant
    F_きりあげ = Application.WorksheetFunction.RoundUp(数値, 切り上げる桁数)
End Function
Public Function F_かつ(論理条件1 As Variant, 論理条件2 As Variant, Optional オプション1 As Variant, Optional オプション2 As Variant, Optional オプション3 As Variant, Optional オプション4 As Variant, Optional オプション5 As Variant) As Variant
    F_かつ = Application.WorksheetFunction.And(論理条件1, 論理条件2, オプション1, オプション2, オプション3, オプション4, オプション5)
End Function
Public Function F_かくど(ラジアンセル As Variant) As Variant
    F_かくど = Application.WorksheetFunction.Degrees(ラジアンセル)
End Function
Public Function F_かいじょう(元の数値 As Variant) As Variant
    F_かいじょう = Application.WorksheetFunction.Fact(元の数値)
End Function
Public Function F_おおもじ(対象文字セル As Variant) As Variant
    F_おおもじ = UCase(対象文字セル)
End Function
Public Function F_おおきいほうからなんばんめ(セル範囲 As Variant, 順位 As Variant) As Variant
    F_おおきいほうからなんばんめ = Application.WorksheetFunction.Large(セル範囲, 順位)
End Function
Public Function F_えんしゅうりつ() As Variant
    F_えんしゅうりつ = Application.WorksheetFunction.Pi
End Function
Public Function F_えいたんごのせんとうもじをおおもじか(英単語を含むセル As Variant) As Variant
    F_えいたんごのせんとうもじをおおもじか = Application.WorksheetFunction.Proper(英単語を含むセル)
End Function
Public Function F_えいぎょうびにっすう(開始日 As Variant, 終了日 As Variant, 祭日を書いたセル範囲 As Variant) As Variant
    F_えいぎょうびにっすう = Application.WorksheetFunction.NetworkDays(開始日, 終了日, 祭日を書いたセル範囲)
End Function
Public Function F_えいぎょうび(開始日 As Variant, 日数 As Variant, 祭日の日付を書いたセル範囲 As Variant) As Variant
    F_えいぎょうび = Application.WorksheetFunction.WorkDay(開始日, 日数, 祭日の日付を書いたセル範囲)
End Function
Public Function F_インデックス(検索範囲 As Variant, 行番号 As Variant, 列番号 As Variant) As Variant
    F_インデックス = Application.WorksheetFunction.index(検索範囲, 行番号, 列番号)
End Function
Public Function F_いま() As Variant
    F_いま = Now
End Function
Public Function F_いっち(検索値 As Variant, 検索範囲 As Variant, 照合の種類 As Variant) As Variant
    F_いっち = Application.WorksheetFunction.Match(検索値, 検索範囲, 照合の種類)
End Function
Public Function F_あいだもじれつ(文字列 As Variant, 先頭文字番号 As Variant, 抜き出し文字数 As Variant) As Variant
    F_あいだもじれつ = Mid(文字列, 先頭文字番号, 抜き出し文字数)
End Function
Public Function F_アークタンジェント(x座標 As Variant, y座標 As Variant) As Variant
    F_アークタンジェント = Application.WorksheetFunction.Atan2(x座標, y座標)
End Function
Public Function F_アークサイン(元のサインの数値 As Variant) As Variant
    F_アークサイン = Application.WorksheetFunction.Asin(元のサインの数値)
End Function
Public Function F_アークコサイン(元のコサインの数値 As Variant) As Variant
    F_アークコサイン = Application.WorksheetFunction.Acos(元のコサインの数値)
End Function
Public Function F_nPr何通り(総数セル As Variant, 取り出す分セル As Variant) As Variant
    F_nPr何通り = Application.WorksheetFunction.Permut(総数セル, 取り出す分セル)
End Function

Public Function F_nCr何通り(総数セル As Variant, 取り出す分セル As Variant) As Variant
    F_nCr何通り = Application.WorksheetFunction.Combin(総数セル, 取り出す分セル)
End Function

Public Function F_日付化(日付シリアル As Variant) As Date
    F_日付化 = CDate(日付シリアル)
End Function

Public Function F_いろ(Optional 赤 As Variant = 255, Optional 緑 As Variant = 255, Optional 青 As Variant = 255) As Variant
    F_いろ = RGB(赤, 緑, 青)
End Function

Public Function F_色(Optional 赤 As Variant = 255, Optional 緑 As Variant = 255, Optional 青 As Variant = 255) As Variant
    F_色 = RGB(赤, 緑, 青)
End Function

Public Function F_色インデックスからRGB色へ変換(idx As カラーインデックスパターン) As Variant
    Select Case idx
    Case 1
        F_色インデックスからRGB色へ変換 = RGB(0, 0, 0)
    Case 2
        F_色インデックスからRGB色へ変換 = RGB(255, 255, 255)
    Case 3
        F_色インデックスからRGB色へ変換 = RGB(255, 0, 0)
    Case 4
        F_色インデックスからRGB色へ変換 = RGB(0, 255, 0)
    Case 5
        F_色インデックスからRGB色へ変換 = RGB(0, 0, 255)
    Case 6
        F_色インデックスからRGB色へ変換 = RGB(255, 255, 0)
    Case 7
        F_色インデックスからRGB色へ変換 = RGB(255, 0, 255)
    Case 8
        F_色インデックスからRGB色へ変換 = RGB(0, 255, 255)
    Case 9
        F_色インデックスからRGB色へ変換 = RGB(128, 0, 0)
    Case 10
        F_色インデックスからRGB色へ変換 = RGB(0, 128, 0)
    Case 11
        F_色インデックスからRGB色へ変換 = RGB(0, 0, 128)
    Case 12
        F_色インデックスからRGB色へ変換 = RGB(128, 128, 0)
    Case 13
        F_色インデックスからRGB色へ変換 = RGB(128, 0, 128)
    Case 14
        F_色インデックスからRGB色へ変換 = RGB(0, 128, 128)
    Case 15
        F_色インデックスからRGB色へ変換 = RGB(192, 192, 192)
    Case 16
        F_色インデックスからRGB色へ変換 = RGB(128, 128, 128)
    Case 17
        F_色インデックスからRGB色へ変換 = RGB(153, 153, 255)
    Case 18
        F_色インデックスからRGB色へ変換 = RGB(153, 51, 102)
    Case 19
        F_色インデックスからRGB色へ変換 = RGB(255, 255, 204)
    Case 20
        F_色インデックスからRGB色へ変換 = RGB(204, 255, 255)
    Case 21
        F_色インデックスからRGB色へ変換 = RGB(102, 0, 102)
    Case 22
        F_色インデックスからRGB色へ変換 = RGB(255, 128, 128)
    Case 23
        F_色インデックスからRGB色へ変換 = RGB(0, 102, 204)
    Case 24
        F_色インデックスからRGB色へ変換 = RGB(204, 204, 255)
    Case 25
        F_色インデックスからRGB色へ変換 = RGB(0, 0, 128)
    Case 26
        F_色インデックスからRGB色へ変換 = RGB(255, 0, 255)
    Case 27
        F_色インデックスからRGB色へ変換 = RGB(255, 255, 0)
    Case 28
        F_色インデックスからRGB色へ変換 = RGB(0, 255, 255)
    Case 29
        F_色インデックスからRGB色へ変換 = RGB(128, 0, 128)
    Case 30
        F_色インデックスからRGB色へ変換 = RGB(128, 0, 0)
    Case 31
        F_色インデックスからRGB色へ変換 = RGB(0, 128, 128)
    Case 32
        F_色インデックスからRGB色へ変換 = RGB(0, 0, 255)
    Case 33
        F_色インデックスからRGB色へ変換 = RGB(0, 204, 255)
    Case 34
        F_色インデックスからRGB色へ変換 = RGB(204, 255, 255)
    Case 35
        F_色インデックスからRGB色へ変換 = RGB(204, 255, 204)
    Case 36
        F_色インデックスからRGB色へ変換 = RGB(255, 255, 153)
    Case 37
        F_色インデックスからRGB色へ変換 = RGB(153, 204, 255)
    Case 38
        F_色インデックスからRGB色へ変換 = RGB(255, 153, 204)
    Case 39
        F_色インデックスからRGB色へ変換 = RGB(204, 153, 255)
    Case 40
        F_色インデックスからRGB色へ変換 = RGB(255, 204, 153)
    Case 41
        F_色インデックスからRGB色へ変換 = RGB(51, 102, 255)
    Case 42
        F_色インデックスからRGB色へ変換 = RGB(51, 204, 204)
    Case 43
        F_色インデックスからRGB色へ変換 = RGB(153, 204, 0)
    Case 44
        F_色インデックスからRGB色へ変換 = RGB(255, 204, 0)
    Case 45
        F_色インデックスからRGB色へ変換 = RGB(255, 153, 0)
    Case 46
        F_色インデックスからRGB色へ変換 = RGB(255, 102, 0)
    Case 47
        F_色インデックスからRGB色へ変換 = RGB(102, 102, 153)
    Case 48
        F_色インデックスからRGB色へ変換 = RGB(150, 150, 150)
    Case 49
        F_色インデックスからRGB色へ変換 = RGB(0, 51, 102)
    Case 50
        F_色インデックスからRGB色へ変換 = RGB(51, 153, 102)
    Case 51
        F_色インデックスからRGB色へ変換 = RGB(0, 51, 0)
    Case 52
        F_色インデックスからRGB色へ変換 = RGB(51, 51, 0)
    Case 53
        F_色インデックスからRGB色へ変換 = RGB(153, 51, 0)
    Case 54
        F_色インデックスからRGB色へ変換 = RGB(153, 51, 102)
    Case 55
        F_色インデックスからRGB色へ変換 = RGB(51, 51, 153)
    Case 56
        F_色インデックスからRGB色へ変換 = RGB(51, 51, 51)
    End Select

End Function

Public Function F_色の三原色を取得(色 As Long, ByRef 赤 As Long, ByRef 緑 As Long, ByRef 青 As Long)
    赤 = 色 Mod 256
    緑 = Int(色 / 256) Mod 256
    青 = Int(色 / 256 / 256)
End Function

Public Function F_今日()
    F_今日 = Int(F_今())
End Function

Public Function F_今日の日付()
    F_今日の日付 = Trim(F_日付化(F_今日()))
End Function

Public Function F_今の日付()
    F_今の日付 = F_日付化(F_今())
End Function






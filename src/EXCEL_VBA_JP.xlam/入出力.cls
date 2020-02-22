VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "入出力"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private m_セル As セル
Private m_シート As シート
Private m_ブック As ブック
Private m_out_出力ファイル名スタック As スタック
Private m_in_入力ファイル名スタック As スタック
'Private m_stack_ファイル名スタック As スタック


Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_セル = New セル
    Set m_シート = New シート
    Set m_ブック = New ブック
    Set m_out_出力ファイル名スタック = New スタック
    Set m_in_入力ファイル名スタック = New スタック
'    Set m_stack_ファイル名スタック = New スタック
    
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
'    Set m_stack_ファイル名スタック = Nothing
    Set m_in_入力ファイル名スタック = Nothing
    Set m_out_出力ファイル名スタック = Nothing
    Set m_ブック = Nothing
    Set m_シート = Nothing
    Set m_セル = Nothing
End Sub

Public Function 数値入力ボックス(Optional メッセージ As Variant = "数値を入力してください。") As Integer
    Dim w_int As Integer
    Dim w_str As String
    
    w_int = 0
    
    w_str = InputBox(メッセージ)
    
    On Error GoTo ERR1:
    
    w_int = CInt(w_str)
    
ERR1:

    数値入力ボックス = w_int
    
End Function

Public Function ファイル書込用開く(Optional ファイル名 As String = "text.txt", Optional フォルダ種別 As ファイル作成フォルダ種別 = ファイル作成フォルダ種別.現在のフォルダ, Optional フォルダ名 As String = "", Optional 書き込み方法 As ファイル書き込み方法 = ファイル書き込み方法.上書き) As String
   
    Dim str_folder As String

    Select Case フォルダ種別
    Case ファイル作成フォルダ種別.現在のフォルダ
        str_folder = ActiveWorkbook.Path
        
    Case ファイル作成フォルダ種別.マイドキュメント
        str_folder = vbajp_MyDoc
    Case ファイル作成フォルダ種別.指定なし
        str_folder = フォルダ名
    Case ファイル作成フォルダ種別.フルパス
        str_folder = ""
    End Select


    Dim fullpath As String
    
    If 書き込み方法 = ファイル書き込み方法.上書き Then
        fullpath = フルパスファイル名取得(str_folder, ファイル名)
    Else
        fullpath = ユニークファイル名取得(str_folder, ファイル名)
    End If
    
    Dim No As Long
    
    No = getNoFromFileFullPath(fullpath)
    
    Open fullpath For Output As #No

    vbajp_cur_FileName_for_write = fullpath
'    m_stack_ファイル名スタック.プッシュ fullpath
    m_out_出力ファイル名スタック.プッシュ fullpath

    Set WSH = Nothing
    
    ファイル書込用開く = fullpath
End Function

Public Function ファイル読込用開く(Optional ファイル名 As String = "text.txt", Optional フォルダ種別 As ファイル作成フォルダ種別 = 現在のフォルダ, Optional フォルダ名 As String = "") As String
    Dim str_folder As String

    Select Case フォルダ種別
    Case ファイル作成フォルダ種別.現在のフォルダ
        str_folder = ActiveWorkbook.Path
        
    Case ファイル作成フォルダ種別.マイドキュメント
        str_folder = vbajp_MyDoc
    Case ファイル作成フォルダ種別.指定なし
        str_folder = フォルダ名
    Case ファイル作成フォルダ種別.フルパス
        str_folder = ""
    End Select


    Dim fullpath As String
    
    fullpath = str_folder & "\" & ファイル名
    
    Dim No As Long
    
    No = getNoFromFileFullPath(fullpath)
    
    Open fullpath For Input As #No

    vbajp_cur_FileName_for_read = fullpath
'    m_stack_ファイル名スタック.プッシュ fullpath
    m_in_入力ファイル名スタック.プッシュ fullpath

    Set WSH = Nothing
    
    ファイル読込用開く = fullpath
End Function



Public Property Get ファイル読込終端か(Optional ファイル名 As String = "", Optional フォルダ種別 As ファイル作成フォルダ種別 = 現在のフォルダ)
    
    Dim str_folder As String

    Dim fullpath As String
    
    If ファイル名 = "" Then
        fullpath = vbajp_cur_FileName_for_write
    Else
        Select Case フォルダ種別
        Case ファイル作成フォルダ種別.現在のフォルダ
            str_folder = ActiveWorkbook.Path
            
        Case ファイル作成フォルダ種別.マイドキュメント
            str_folder = vbajp_MyDoc
        Case ファイル作成フォルダ種別.指定なし
            str_folder = フォルダ名
        Case ファイル作成フォルダ種別.フルパス
            str_folder = ""
        End Select
    
    
        
        fullpath = str_folder & "\" & ファイル名
            
    End If


    
    Dim No As Long
    
    No = getNoFromFileFullPath(fullpath)
    
    ファイル読込終端か = EOF(No)
End Property

Public Sub ファイル1行出力(メッセージ As String, Optional ファイル名 As String = "")
    Dim fullpath As String
    
    If ファイル名 = "" Then
        fullpath = vbajp_cur_FileName_for_write
    Else
'        Select Case フォルダ種別
'        Case ファイル作成フォルダ種別.現在のフォルダ
'            str_folder = ActiveWorkbook.Path
'
'        Case ファイル作成フォルダ種別.マイドキュメント
'            str_folder = vbajp_MyDoc
'        Case ファイル作成フォルダ種別.指定なし
'            str_folder = フォルダ名
'        Case ファイル作成フォルダ種別.フルパス
'            str_folder = ""
'        End Select
'
    
        
        fullpath = ファイル名
            
    End If

    Dim No As Long
    No = getNoFromFileFullPathNonePreserve(fullpath)
'    If no <= 0 Then
'        Exit Sub
'    End If

    Print #No, メッセージ
End Sub

Public Function ファイル1行読込(Optional ファイル名 As String = "") As Variant
    Dim fullpath As String
    
    If ファイル名 = "" Then
        fullpath = vbajp_cur_FileName_for_write
    Else
'        Select Case フォルダ種別
'        Case ファイル作成フォルダ種別.現在のフォルダ
'            str_folder = ActiveWorkbook.Path
'
'        Case ファイル作成フォルダ種別.マイドキュメント
'            str_folder = vbajp_MyDoc
'        Case ファイル作成フォルダ種別.指定なし
'            str_folder = フォルダ名
'        Case ファイル作成フォルダ種別.フルパス
'            str_folder = ""
'        End Select
    
    
        
        fullpath = ファイル名
            
    End If

    Dim No As Long
    No = getNoFromFileFullPathNonePreserve(fullpath)
    If No <= 0 Then
        ファイル1行読込 = ""
        Exit Function
    End If

    Dim line As String
    
    Line Input #No, line
    ファイル1行読込 = line
    
End Function


Public Sub ファイル閉じる(Optional ファイル名 As String = "")

    Dim fullpath As String
    
'    With r_RegExp
'        .Pattern = "^[a-zA-z]:"
'        .Global = False
'    End With
    
    If ファイル名 = "" Then
        fullpath = vbajp_cur_FileName_for_write
'    ElseIf r_RegExp.Match(ファイル名) Then
'        fullpath = ファイル名
    Else
'        Select Case フォルダ種別
'        Case ファイル作成フォルダ種別.現在のフォルダ
'            str_folder = ActiveWorkbook.Path
'
'        Case ファイル作成フォルダ種別.マイドキュメント
'            str_folder = vbajp_MyDoc
'        Case ファイル作成フォルダ種別.指定なし
'            str_folder = フォルダ名
'        Case ファイル作成フォルダ種別.フルパス
'            str_folder = ""
'        End Select
    
    
'        If str_folder <> "" Then
'            fullpath = str_folder & "\" & ファイル名
'        Else
'            fullpath = ファイル名
'        End If
            
        'fullpath = str_folder & "\" & ファイル名
        fullpath = ファイル名
    End If

    If m_out_出力ファイル名スタック.見るだけ() = fullpath Then
        vbajp_cur_FileName_for_write = m_out_出力ファイル名スタック.ポップ()
        vbajp_cur_FileName_for_write = m_out_出力ファイル名スタック.見るだけ()
    ElseIf m_in_入力ファイル名スタック.見るだけ() = fullpath Then
        vbajp_cur_FileName_for_read = m_in_入力ファイル名スタック.ポップ()
        vbajp_cur_FileName_for_read = m_in_入力ファイル名スタック.見るだけ()
    End If


    Dim No As Long
    'no = getNoFromFileFullPathNonePreserve(fullpath)
    No = getNoFromFileFullPath(fullpath)
    If No <= 0 Then
        Exit Sub
    End If

    Close #No
    
    
    
End Sub

Private Function getNoFromFileFullPath(fullpath As String) As Long
    Dim i As Long
    Dim w_str As String
    Dim No As Long
    Dim ary() As String
    
    No = 1
    
    On Error Resume Next
    For i = LBound(vbajp_ary_Files) To UBound(vbajp_ary_Files)
        w_str = ""
        w_str = vbajp_ary_Files(i)
        
        ary = Split(w_str, ",")
        
        If UBound(ary) < 1 Then GoTo SKIP1
        
        w_str = ary(1)
        
        If ary(0) = fullpath Then
            No = CLng(w_str)
            getNoFromFileFullPath = No
            Exit Function
        End If
        
SKIP1:
    Next
    On Error GoTo 0
    
    On Error GoTo ERR1
    If UBound(vbajp_ary_Files) > 0 Then
        ReDim Preserve vbajp_ary_Files(UBound(vbajp_ary_Files) + 1)
    Else
        ReDim Preserve vbajp_ary_Files(1)
    End If
    vbajp_ary_Files(UBound(vbajp_ary_Files)) = fullpath & "," & UBound(vbajp_ary_Files)
    getNoFromFileFullPath = UBound(vbajp_ary_Files)
    Exit Function
    
ERR1:
    ReDim Preserve vbajp_ary_Files(1)
    vbajp_ary_Files(UBound(vbajp_ary_Files)) = fullpath & "," & 1
    getNoFromFileFullPath = UBound(vbajp_ary_Files)
    
    
End Function

Private Function getNoFromFileFullPathNonePreserve(fullpath As String) As Long
    Dim i As Long
    Dim w_str As String
    Dim No As Long
    Dim ary() As String
    
    No = 1
    
    On Error Resume Next
    For i = LBound(vbajp_ary_Files) To UBound(vbajp_ary_Files)
        w_str = ""
        w_str = vbajp_ary_Files(i)
        ary = Split(w_str, ",")
        
        If UBound(ary) < 1 Then GoTo SKIP1
        
        w_str = ary(1)
        
        
        If ary(0) = fullpath Then
            No = CLng(w_str)
            getNoFromFileFullPathNonePreserve = No
            'Call erase_array_n(vbajp_ary_Files, i)
            Exit Function
        End If
        
SKIP1:
    Next
    On Error GoTo 0
    
    getNoFromFileFullPathNonePreserve = 0
    
End Function

Private Function erase_array_n(ary() As String, n As Long)
    For i = n To UBound(ary) - 1
        ary(i) = ary(i + 1)
    Next i
    If UBound(ary) - 1 > 0 Then
        ReDim Preserve ary(UBound(ary) - 1)
    Else
        ReDim Preserve ary(0)
    End If
End Function



Public Function 文字入力ボックス(Optional メッセージ As Variant = "文字を入力してください。") As String
    Dim w_str As String
    
    
    w_str = InputBox(メッセージ)
    

    文字入力ボックス = w_str
    
End Function

Private Function フルパスファイル名取得(Optional ディレクトリ名 As String = "", Optional ファイル名 As String = "text.txt") As String

    If ディレクトリ名 = "" Then
        ディレクトリ名 = ActiveWorkbook.Path
    End If

    Dim No As Long
    No = 1
        
    Dim orig_fullpath As String
    orig_fullpath = ディレクトリ名 & "\" & ファイル名
    
    
    フルパスファイル名取得 = orig_fullpath
    Exit Function
    
'    Dim fullpath As String
'    Dim isFirst As Boolean
'
'    If Dir(orig_fullpath) = "" Then
'        フルパスファイル名取得 = orig_fullpath
'        Exit Function
'    End If
'
'    no = 1
'
'    fullpath = ディレクトリ名 & "\" & GetFNameFromFStr(ファイル名) & no & ".txt"
'    If Dir(fullpath) = "" Then
'        フルパスファイル名取得 = fullpath
'        Exit Function
'    End If
'
'    フルパスファイル名取得 = orig_fullpath
End Function

Private Function ユニークファイル名取得(Optional ディレクトリ名 As String = "", Optional ファイル名 As String = "text.txt") As String

    If ディレクトリ名 = "" Then
        ディレクトリ名 = ActiveWorkbook.Path
    End If

    Dim No As Long
    No = 1
        
    Dim orig_fullpath As String
    orig_fullpath = ディレクトリ名 & "\" & ファイル名
    
    
    Dim fullpath As String
    Dim isFirst As Boolean
    
    For No = 0 To 1000
        If No = 0 Then
            If Dir(orig_fullpath) = "" Then
                ユニークファイル名取得 = orig_fullpath
                Exit Function
            End If
            No = 1
        End If
        
        fullpath = ディレクトリ名 & "\" & GetFNameFromFStr(ファイル名) & No & ".txt"
        If Dir(fullpath) = "" Then
            ユニークファイル名取得 = fullpath
            Exit Function
        End If
    Next
    
    ユニークファイル名取得 = orig_fullpath
End Function


'http://blog.livedoor.jp/excelvbamemo/archives/53683974.html
'--------------------------------------------------------------------
' 関数名 : GetFNameFromFStr
' 機能   : ファイル名から拡張子を除いた文字列を取得する
' 引数   : sFileName : ファイル名[I]
' 戻り値 : ファイル名（拡張子を除く）
' 作成日 : 2007/04/12
' 作成者 : lily(http://blog.livedoor.jp/excelvbamemo/)
'--------------------------------------------------------------------
Private Function GetFNameFromFStr(sFileName As String) As String
Dim sFileStr As String
Dim lFindPoint As Long
Dim lStrLen As Long
    
    '文字列の右端から"."を検索し、左端からの位置を取得する
    lFindPoint = InStrRev(sFileName, ".")
    
    '拡張子を除いたファイル名の取得
    sFileStr = Left(sFileName, lFindPoint - 1)

    GetFNameFromFStr = sFileStr
End Function

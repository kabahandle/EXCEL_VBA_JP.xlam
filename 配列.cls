VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "配列"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private m_Array() As Variant

Private Sub Class_Initialize()
    'ここに処理を書く
    ReDim Preserve m_Array(0)
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    'ReDim m_Array(0)
    'set m_Array = Nothing
    'Unload m_Array
    
End Sub

Public Property Get Array取得() As Variant()
    Array取得 = m_Array
End Property




Public Property Set セット(配列 As 配列)
    m_Array = 配列.Array取得
End Property

Public Sub 要素数指定(要素数 As Long)
    
    ReDim Preserve m_Array(要素数)
    
    'Set 要素数指定 = Me
    
End Sub

Public Sub 追加(値 As Variant)
    ReDim Preserve m_Array(UBound(m_Array) + 1)
    
    m_Array(UBound(m_Array)) = 値
    
    'ReDim Preserve m_Array(UBound(m_Array) + 1)
    
    'Set 追加 = Me
End Sub

Public Function 取り出し(Optional 番号 As Long = 1) As Variant
    If isOutBound(番号) Then
        取り出し = ""
        Exit Function
    End If
    
    取り出し = m_Array(番号)
End Function


Public Sub 削除(Optional 番号 As Long = 1)
    If isOutBound(番号) Then
        'Set 削除 = Me
        Exit Sub
    End If
    
    Dim i As Long


    If 番号 = UBound(m_Array) Then
        ReDim Preserve m_Array(UBound(m_Array))
    Else
        For i = 番号 To UBound(m_Array) - 1
            m_Array(i) = m_Array(i + 1)
        Next i
        ReDim Preserve m_Array(UBound(m_Array) - 1)
    End If
    
    'Set 削除 = Me
    
End Sub

Public Sub 文字列追加(文字列 As String, Optional 区切り文字 As 文字列区切り文字 = 文字列区切り文字.なし, Optional その他区切り文字 As String = " ")
    Dim arrStr() As String
    Dim delim As String
    
    Select Case 区切り文字
    Case 文字列区切り文字.なし
        delim = ""
    Case 文字列区切り文字.カンマ
        delim = ","
    Case 文字列区切り文字.タブ
        delim = vbTab
    Case 文字列区切り文字.改行
        delim = vbCrLf
    Case 文字列区切り文字.Cr
        delim = "" & Chr(13)
    Case 文字列区切り文字.半角空白
        delim = " "
    Case 文字列区切り文字.その他
        delim = その他区切り文字
    End Select
    
    If delim <> "" Then
        arrStr = Split(文字列, delim)
    Else
        arrStr = Array(文字列)
    End If
    
    Dim old_max As Long
    old_max = UBound(m_Array)
    
    ReDim m_Array(UBound(m_Array) + UBound(arrStr))
    
    Dim new_max As Long
    
    new_max = UBound(m_Array)
    
    Dim i As Long
    Dim l As Long
    l = 0
    
    For i = old_max + 1 To new_max
        m_Array(i) = arrStr(l)
        l = l + 1
    Next i

    'Set 文字列追加 = Me
End Sub

Public Sub クリア()
    ReDim m_Array(0)
    
    'Set クリア = Me
End Sub

Public Function 文字列化(Optional 区切り文字 As 文字列区切り文字 = 文字列区切り文字.カンマ, Optional その他区切り文字 As String = " ") As String
    Dim retStr As String
    Dim tmpStr As String
    Dim i As Long
    
    
    Dim delim As String
    
    Select Case 区切り文字
    Case 文字列区切り文字.なし
        delim = ""
    Case 文字列区切り文字.カンマ
        delim = ","
    Case 文字列区切り文字.タブ
        delim = vbTab
    Case 文字列区切り文字.改行
        delim = vbCrLf
    Case 文字列区切り文字.Cr
        delim = "" & Chr(13)
    Case 文字列区切り文字.その他
        delim = その他区切り文字
    End Select
    
    retStr = ""
    tmpStr = ""
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        On Error GoTo ERR1
        tmpStr = CStr(m_Array(i))
        retStr = retStr & delim & tmpStr
ERR1:
        On Error GoTo 0
        'i = i + 1
    Next i
    
    If delim <> "" Then
        retStr = Mid(retStr, 2)
    End If
    
    文字列化 = retStr

End Function

Public Sub 値を指定して削除(値 As Variant, Optional 削除オプション As 配列の値を指定して削除オプション = 配列の値を指定して削除オプション.全該当要素削除)
'    Dim i As Long
'    Dim x As Long
'
'    For i = LBound(m_Array) + 1 To UBound(m_Array)
'        If m_Array(i) = 値 Then
'            x = i
'            Me.削除 (x)
'            Exit Function
'        End If
'    Next i

    Select Case 削除オプション
    Case 配列の値を指定して削除オプション.全該当要素削除
        Call deleteItemAll(値)
    Case 配列の値を指定して削除オプション.最初の要素だけ削除
        Call deleteItemFirst(値)
    End Select
            
    'Set 値を指定して削除 = Me

End Sub

Private Sub deleteItemFirst(値 As Variant)
    Dim i As Long
    Dim x As Long
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        If m_Array(i) = 値 Then
            x = i
            Me.削除 (x)
            Exit Sub
        End If
    Next i
End Sub

Private Sub deleteItemAll(値 As Variant)
    Dim i As Long
    Dim x As Long
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        If m_Array(i) = 値 Then
            x = i
            Call Me.削除(x)
            Exit For
        End If
    Next i
    If isExistsItem(値) Then
        Call deleteItemAll(値)
    End If
End Sub

Private Function isExistsItem(値 As Variant) As Boolean
    Dim i As Long
    Dim x As Long
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        If m_Array(i) = 値 Then
            isExistsItem = True
            Exit Function
        End If
    Next i
    isExistsItem = False
End Function

Public Function 最小番号() As Long
    最小番号 = LBound(m_Array) + 1
End Function

Public Function 最大番号() As Long
    最大番号 = UBound(m_Array)
End Function

Private Function isOutBound(番号 As Long) As Boolean
    If LBound(m_Array) + 1 > 番号 Or UBound(m_Array) < 番号 Then
        isOutBound = True
    Else
        isOutBound = False
    End If
End Function

Public Function 番号が範囲内か(番号 As Long) As Boolean
    番号が範囲内か = Not isOutBound(番号)
End Function

Public Sub 要素ごとにメソッド呼び出し(サブルーチン名 As String, _
                Optional 引数1 As Variant, _
                Optional 引数2 As Variant, _
                Optional 引数3 As Variant, _
                Optional 引数4 As Variant, _
                Optional 引数5 As Variant, _
                Optional 引数6 As Variant, _
                Optional 引数7 As Variant, _
                Optional 引数8 As Variant, _
                Optional 引数9 As Variant)
    Dim i As Long
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        Application.Run ActiveWorkbook.Name & "!" & サブルーチン名, m_Array(i), 引数1, 引数2, 引数3, 引数4, 引数5, 引数6, 引数7, 引数8, 引数9
    Next
    
End Sub

Public Sub 要素ごとにメソッド呼び出し2(ワークブック名付きサブルーチン名 As String, _
                Optional 引数1 As Variant, _
                Optional 引数2 As Variant, _
                Optional 引数3 As Variant, _
                Optional 引数4 As Variant, _
                Optional 引数5 As Variant, _
                Optional 引数6 As Variant, _
                Optional 引数7 As Variant, _
                Optional 引数8 As Variant, _
                Optional 引数9 As Variant)
    Dim i As Long
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        Application.Run ワークブック名付きサブルーチン名, m_Array(i), 引数1, 引数2, 引数3, 引数4, 引数5, 引数6, 引数7, 引数8, 引数9
    Next
    
End Sub




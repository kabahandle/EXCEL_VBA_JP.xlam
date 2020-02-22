VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "スタック"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private m_Array() As Variant
Private index As Long


Private Sub Class_Initialize()
    'ここに処理を書く
    index = 0
    ReDim Preserve m_Array(0)
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    'ReDim m_Array(0)
    'set m_Array = Nothing
    'Unload m_Array
    
End Sub

Public Function プッシュ(値 As Variant) As Variant
    index = index + 1
    ReDim Preserve m_Array(index)
    m_Array(index) = 値
    プッシュ = 値
End Function

Public Function ポップ() As Variant
    Dim 返値 As Variant
    
    返値 = m_Array(index)
    
    index = index - 1
    If index < 1 Then
        index = 0
    End If
    
    ReDim Preserve m_Array(index)
    
    ポップ = 返値
End Function

Public Function 見るだけ() As Variant
    Dim 返値 As Variant
    
    返値 = m_Array(index)

    見るだけ = 返値

End Function

Public Function 文字列化(Optional 区切り文字 As 文字列区切り文字 = 文字列区切り文字.カンマ, Optional その他区切り文字 As String = " ") As String
    Dim i As Long
    Dim strg As String
    
    'Dim retStr As String
    Dim tmpStr As String
    
    
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
    
    'retStr = ""
    
    strg = ""
    
    For i = 1 To UBound(m_Array) Step 1
        'strg = strg & delim & m_Array(i)
        tmpStr = ""
        On Error GoTo ERR1
        tmpStr = CStr(m_Array(i))
        
ERR1:
        On Error GoTo 0
        strg = strg & delim & tmpStr
    Next i

    If delim <> "" Then
        strg = Mid(strg, 2)
    End If

    文字列化 = strg

End Function


Public Function 逆順文字列化(Optional 区切り文字 As 文字列区切り文字 = 文字列区切り文字.カンマ, Optional その他区切り文字 As String = " ") As String
    Dim i As Long
    Dim strg As String
    
    'Dim retStr As String
    Dim tmpStr As String
    
    
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
    
    
    strg = ""
    
    For i = UBound(m_Array) To 1 Step -1
        tmpStr = ""
        On Error GoTo ERR1
        tmpStr = CStr(m_Array(i))
        
ERR1:
        On Error GoTo 0
        strg = strg & delim & tmpStr
    Next i

    If delim <> "" Then
        strg = Mid(strg, 2)
    End If
    
    逆順文字列化 = strg

End Function

Public Sub クリア()
    ReDim m_Array(0)
    
    'Set クリア = Me
End Sub

Public Sub 下から順にメソッド呼び出し(サブルーチン名 As String, _
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
    
    For i = 1 To UBound(m_Array)
        Application.Run ActiveWorkbook.Name & "!" & サブルーチン名, m_Array(i), 引数1, 引数2, 引数3, 引数4, 引数5, 引数6, 引数7, 引数8, 引数9
    Next i
    
End Sub

Public Sub 下から順にメソッド呼び出し2(ワークブック名付きサブルーチン名 As String, _
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
    
    For i = 1 To UBound(m_Array)
        Application.Run ワークブック名付きサブルーチン名, m_Array(i), 引数1, 引数2, 引数3, 引数4, 引数5, 引数6, 引数7, 引数8, 引数9
    Next i
    
End Sub

Public Sub 上から順にメソッド呼び出し(サブルーチン名 As String, _
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
    
    For i = UBound(m_Array) To 1 Step -1
        Application.Run ActiveWorkbook.Name & "!" & サブルーチン名, m_Array(i), 引数1, 引数2, 引数3, 引数4, 引数5, 引数6, 引数7, 引数8, 引数9
    Next i
    
End Sub

Public Sub 上から順にメソッド呼び出し2(ワークブック名付きサブルーチン名 As String, _
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
    
    For i = UBound(m_Array) To 1 Step -1
        Application.Run ワークブック名付きサブルーチン名, m_Array(i), 引数1, 引数2, 引数3, 引数4, 引数5, 引数6, 引数7, 引数8, 引数9
    Next i
    
End Sub





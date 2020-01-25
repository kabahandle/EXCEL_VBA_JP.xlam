VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ハッシュ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Const DebugMode = False

'Private m_HashObj As Dictionary 'TODO ->objectにして、"Windows Scripting Runtime"のチェックをはずす
Private m_HashObj As Object 'TODO ->objectにして、"Windows Scripting Runtime"のチェックをはずす

Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_HashObj = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_HashObj = Nothing
End Sub


'--- end of 日本語間枢要
Public Property Get ToVBA() As Object
    ToVBA = m_HashObj
End Property

Public Property Set セット(ハッシュ As ハッシュ)
    Set m_HashObj = ハッシュ.ToVBA
End Property

Public Property Let テキスト比較モード(テキスト比較モード As テキスト比較方法)
    m_HashObj.CompareMode = テキスト比較モード
End Property


Public Function 追加(キー As Variant, 値 As Variant) As ハッシュ
    'If m_HashObj.Exists(キー) Then
    'Else
        m_HashObj.Add キー, 値
    'End If
    Set 追加 = Me
End Function

' 値 -  キーが存在し、値を取得できた
' Empty - キーが存在せず、値を取得できなかった
Public Function 値(キー As Variant) As Variant
    If m_HashObj.Exists(キー) Then
        値 = m_HashObj.Item(キー)
        Exit Function
    Else
        値 = Empty
        Exit Function
    End If
End Function

Public Function 値設定(キー As Variant, 値 As Variant) As ハッシュ
    m_HashObj.Item(キー) = 値
    Set 値設定 = Me
End Function

'Public Function 項目配列取得() As Variant()
'    項目配列取得 = m_HashObj.Items
'End Function

Public Function 項目数() As Long
    項目数 = m_HashObj.Count
End Function

Public Function キーが存在する(キー As Variant) As Boolean
    キーが存在する = m_HashObj.Exists(キー)
End Function

' TRUE -  旧キーが存在し、新キーへ変更できた
' FALSE - 旧キーが存在しない
Public Function キー変更(旧キー As Variant, 新キー As Variant) _
        As Variant
    
    If m_HashObj.Exists(旧キー) Then
        m_HashObj.key(旧キー) = 新キー
        キー変更 = True
        Exit Function
    Else
        キー変更 = False
        Exit Function
    End If
    
End Function

Public Function キーの配列() As Variant()
    キーの配列 = m_HashObj.Keys
End Function

Public Function キーと値を削除(キー As Variant) As ハッシュ
    m_HashObj.Remove (キー)
    Set キーと値を削除 = Me
End Function

Public Function すべてクリア() As ハッシュ
    m_HashObj.RemoveAll
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
    
    Dim v As Variant
    
    For Each v In m_HashObj.Keys
        tmpStr = ""
        
        On Error GoTo ERR1
        
        tmpStr = CStr(Me.値(v))
ERR1:
        On Error GoTo 0
        
        strg = strg & delim & tmpStr
    Next v
    
    If delim <> "" Then
        strg = Mid(strg, 2)
    End If

    文字列化 = strg
    
End Function

Public Sub 要素ごとにサブルーチン呼び出し(サブルーチン名 As String, _
                Optional 引数1 As Variant, _
                Optional 引数2 As Variant, _
                Optional 引数3 As Variant, _
                Optional 引数4 As Variant, _
                Optional 引数5 As Variant, _
                Optional 引数6 As Variant, _
                Optional 引数7 As Variant, _
                Optional 引数8 As Variant, _
                Optional 引数9 As Variant)
    Dim key As Variant
    
    For Each key In m_HashObj.Keys
        Application.Run ActiveWorkbook.Name & "!" & サブルーチン名, key, Me.値(key), 引数1, 引数2, 引数3, 引数4, 引数5, 引数6, 引数7, 引数8, 引数9
    Next key
    
End Sub

Public Sub 要素ごとにサブルーチン呼び出し2(ワークブック名付きサブルーチン名 As String, _
                Optional 引数1 As Variant, _
                Optional 引数2 As Variant, _
                Optional 引数3 As Variant, _
                Optional 引数4 As Variant, _
                Optional 引数5 As Variant, _
                Optional 引数6 As Variant, _
                Optional 引数7 As Variant, _
                Optional 引数8 As Variant, _
                Optional 引数9 As Variant)
    Dim key As Variant
    
    For Each key In m_HashObj.Keys
        Application.Run ワークブック名付きサブルーチン名, key, Me.値(key), 引数1, 引数2, 引数3, 引数4, 引数5, 引数6, 引数7, 引数8, 引数9
    Next key
    
End Sub

Public Function ToJSON(Optional オブジェクト名 As String = "") As String
    Dim strg As String
    Dim key As Variant
    
    If オブジェクト名 <> "" Then
        strg = "var " & オブジェクト名 & " = { " & vbCrLf
    Else
        strg = " { " & vbCrLf
    End If
    
    For Each key In m_HashObj.Keys
        strg = strg & """" & key & """" & ":" & json列値エンコード(Me.値(key)) & "," & vbCrLf
    Next key
    
    strg = strg & " }; "
    
    ToJSON = strg
    
End Function

Public Function FromJSON(JSONコード As String) As ハッシュ
    Dim line() As String
    Dim max As Long
    Dim keyvalue() As String
    Dim i As Long
    
    JSONコード = Replace(JSONコード, vbCrLf, vbLf)
    line = Split(JSONコード, vbLf)
    
    On Error Resume Next
    
    For i = LBound(line) To UBound(line)
        keyvalue = Split(line(i), ":")
        If UBound(keyvalue) = 1 Then
            Dim strgKey As String
            strgKey = Replace(keyvalue(0), """", "")
            Me.追加 strgKey, keyvalue(1)
        End If
    Next i
    
    On Error GoTo 0
    
    Set FromJSON = Me
    
End Function

Public Function json列値エンコード(strg As String) As String
    strg = Replace(strg, "\", "\\")
    strg = Replace(strg, """", "\""")
    strg = Replace(strg, "/", "\/")
    strg = Replace(strg, vbKeyBack, "\b")
    strg = Replace(strg, Asc(12), "\f")
    strg = Replace(strg, vbCr, "\r")
    strg = Replace(strg, vbLf, "\f")
    strg = Replace(strg, vbTab, "\t")
    json列値エンコード = strg
End Function


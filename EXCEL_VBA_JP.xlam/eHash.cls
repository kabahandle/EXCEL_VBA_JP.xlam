VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public Property Set ToSet(Hash As eHash)
    Set m_HashObj = Hash.ToVBA
End Property

Public Property Let SetTextCompateMode(TextCompareMode As TextCompareMode)
    m_HashObj.CompareMode = TextCompareMode
End Property


Public Function Add(key As Variant, Value As Variant) As eHash
    'If m_HashObj.Exists(キー) Then
    'Else
        m_HashObj.Add key, Value
    'End If
    Set Add = Me
End Function

' 値 -  キーが存在し、値を取得できた
' Empty - キーが存在せず、値を取得できなかった
Public Function Value(key As Variant) As Variant
    If m_HashObj.Exists(key) Then
        Value = m_HashObj.Item(key)
        Exit Function
    Else
        Value = Empty
        Exit Function
    End If
End Function

Public Function SetValue(key As Variant, Value As Variant) As eHash
    m_HashObj.Item(key) = Value
    Set SetValue = Me
End Function

'Public Function 項目配列取得() As Variant()
'    項目配列取得 = m_HashObj.Items
'End Function

Public Function Count() As Long
    Count = m_HashObj.Count
End Function

Public Function IsKeyExists(key As Variant) As Boolean
    IsKeyExists = m_HashObj.Exists(key)
End Function

' TRUE -  旧キーが存在し、新キーへ変更できた
' FALSE - 旧キーが存在しない
Public Function ChangeKey(OldKey As Variant, NewKey As Variant) _
        As Variant
    
    If m_HashObj.Exists(OldKey) Then
        m_HashObj.key(OldKey) = NewKey
        ChangeKey = True
        Exit Function
    Else
        ChangeKey = False
        Exit Function
    End If
    
End Function

Public Function Keys() As Variant()
    Keys = m_HashObj.Keys
End Function

Public Function RemoveSetOfKeyAndValueByKey(key As Variant) As eHash
    m_HashObj.Remove key
    Set RemoveSetOfKeyAndValueByKey = Me
End Function

Public Function RemoveAll() As eHash
    m_HashObj.RemoveAll
End Function

Public Function ToString(Optional Separator As SeparatorChar = SeparatorChar.vbeComma, Optional ElseSeparator As String = " ") As String
    Dim i As Long
    Dim strg As String
    
    'Dim retStr As String
    Dim tmpStr As String
    
    
    Dim delim As String
    
    Select Case Separator
    Case SeparatorChar.vbeNoneChar
        delim = ""
    Case SeparatorChar.vbeComma
        delim = ","
    Case SeparatorChar.vbeTab
        delim = vbTab
    Case SeparatorChar.vbeReturn
        delim = vbCrLf
    Case SeparatorChar.vbeCr
        delim = "" & Chr(13)
    Case SeparatorChar.vbeElseChar
        delim = ElseSeparator
    End Select
    
    'retStr = ""
    
    strg = ""
    
    Dim v As Variant
    
    For Each v In m_HashObj.Keys
        tmpStr = ""
        
        On Error GoTo ERR1
        
        tmpStr = CStr(Me.Value(v))
ERR1:
        On Error GoTo 0
        
        strg = strg & delim & tmpStr
    Next v
    
    If delim <> "" Then
        strg = Mid(strg, 2)
    End If

    ToString = strg
    
End Function

Public Sub CallSubroutineByElements(SubRoutineName As String, _
                Optional arg1 As Variant, _
                Optional arg2 As Variant, _
                Optional arg3 As Variant, _
                Optional arg4 As Variant, _
                Optional arg5 As Variant, _
                Optional arg6 As Variant, _
                Optional arg7 As Variant, _
                Optional arg8 As Variant, _
                Optional arg9 As Variant)
    Dim key As Variant
    
    For Each key In m_HashObj.Keys
        Application.Run ActiveWorkbook.Name & "!" & SubRoutineName, key, Me.Value(key), arg1, arg2, arg3, arg4, _
                        arg5, arg6, arg7, arg8, arg9
    Next key
    
End Sub

Public Sub CallSubroutineByElements2(SubRoutineNameWithWorkbookName As String, _
                Optional arg1 As Variant, _
                Optional arg2 As Variant, _
                Optional arg3 As Variant, _
                Optional arg4 As Variant, _
                Optional arg5 As Variant, _
                Optional arg6 As Variant, _
                Optional arg7 As Variant, _
                Optional arg8 As Variant, _
                Optional arg9 As Variant)
    Dim key As Variant
    
    For Each key In m_HashObj.Keys
        Application.Run SubRoutineNameWithWorkbookName, key, Me.Value(key), arg1, arg2, arg3, arg4, _
                        arg5, arg6, arg7, arg8, arg9
    Next key
    
End Sub

Public Function ToJSON(Optional ObjectName As String = "") As String
    Dim strg As String
    Dim key As Variant
    
    If ObjectName <> "" Then
        strg = "var " & ObjectName & " = { " & vbCrLf
    Else
        strg = " { " & vbCrLf
    End If
    
    For Each key In m_HashObj.Keys
        strg = strg & """" & key & """" & ":" & jsonValueEncode(Me.Value(key)) & "," & vbCrLf
    Next key
    
    strg = strg & " }; "
    
    ToJSON = strg
    
End Function

Public Function FromJSON(JSONCode As String) As eHash
    Dim line() As String
    Dim max As Long
    Dim keyvalue() As String
    Dim i As Long
    
    JSONCode = Replace(JSONCode, vbCrLf, vbLf)
    line = Split(JSONCode, vbLf)
    
    On Error Resume Next
    
    For i = LBound(line) To UBound(line)
        keyvalue = Split(line(i), ":")
        If UBound(keyvalue) = 1 Then
            Dim strgKey As String
            strgKey = Replace(keyvalue(0), """", "")
            Me.Add strgKey, keyvalue(1)
        End If
    Next i
    
    On Error GoTo 0
    
    Set FromJSON = Me
    
End Function

Public Function jsonValueEncode(strg As String) As String
    strg = Replace(strg, "\", "\\")
    strg = Replace(strg, """", "\""")
    strg = Replace(strg, "/", "\/")
    strg = Replace(strg, vbKeyBack, "\b")
    strg = Replace(strg, Asc(12), "\f")
    strg = Replace(strg, vbCr, "\r")
    strg = Replace(strg, vbLf, "\f")
    strg = Replace(strg, vbTab, "\t")
    jsonValueEncode = strg
End Function



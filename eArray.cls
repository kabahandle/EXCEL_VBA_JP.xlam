VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eArray"
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

Public Property Get GetArray() As Variant()
    GetArray = m_Array
End Property




Public Property Set ToSet(ArrayValue As eArray)
    m_Array = ArrayValue.GetArray()
End Property

Public Sub SetElementsNum(ElementsNum As Long)
    
    ReDim Preserve m_Array(ElementsNum)
    
    'Set 要素数指定 = Me
    
End Sub

Public Sub Add(Value As Variant)
    ReDim Preserve m_Array(UBound(m_Array) + 1)
    
    m_Array(UBound(m_Array)) = Value
    
    'ReDim Preserve m_Array(UBound(m_Array) + 1)
    
    'Set 追加 = Me
End Sub

Public Function GetValue(Optional ValueNo As Long = 1) As Variant
    If isOutBound(ValueNo) Then
        GetValue = ""
        Exit Function
    End If
    
    GetValue = m_Array(ValueNo)
End Function


Public Sub Delete(Optional ValueNo As Long = 1)
    If isOutBound(ValueNo) Then
        'Set 削除 = Me
        Exit Sub
    End If
    
    Dim i As Long


    If ValueNo = UBound(m_Array) Then
        ReDim Preserve m_Array(UBound(m_Array))
    Else
        For i = ValueNo To UBound(m_Array) - 1
            m_Array(i) = m_Array(i + 1)
        Next i
        ReDim Preserve m_Array(UBound(m_Array) - 1)
    End If
    
    'Set 削除 = Me
    
End Sub

Public Sub AddString(StringValue As String, Optional Separtor As SeparatorChar = SeparatorChar.vbeNoneChar, Optional ElseSeparator As String = " ")
    Dim arrStr() As String
    Dim delim As String
    
    Select Case Separtor
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
    Case SeparatorChar.vbeSpaceChar
        delim = " "
    Case SeparatorChar.vbeElseChar
        delim = ElseSeparator
    End Select
    
    If delim <> "" Then
        arrStr = Split(StringValue, delim)
    Else
        arrStr = Array(StringValue)
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

Public Sub Clear()
    ReDim m_Array(0)
    
    'Set クリア = Me
End Sub

Public Function ToString(Optional Separtor As SeparatorChar = SeparatorChar.vbeComma, Optional ElseSeparator As String = " ") As String
    Dim retStr As String
    Dim tmpStr As String
    Dim i As Long
    
    
    Dim delim As String
    
    Select Case Separtor
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
    
    ToString = retStr

End Function

Public Sub DeleteByValue(Value As Variant, Optional DeleteByValueOption As DeleteByValueOptionForArrayElement = DeleteByValueOptionForArrayElement.AllMatchValues)
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

    Select Case DeleteByValueOption
    Case DeleteByValueOptionForArrayElement.AllMatchValues
        Call deleteItemAll(Value)
    Case DeleteByValueOptionForArrayElement.FirstMatchValueOnly
        Call deleteItemFirst(Value)
    End Select
            
    'Set 値を指定して削除 = Me

End Sub

Private Sub deleteItemFirst(Value As Variant)
    Dim i As Long
    Dim x As Long
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        If m_Array(i) = Value Then
            x = i
            Call Me.Delete(x)
            Exit Sub
        End If
    Next i
End Sub

Private Sub deleteItemAll(Value As Variant)
    Dim i As Long
    Dim x As Long
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        If m_Array(i) = Value Then
            x = i
            Call Me.Delete(x)
            Exit For
        End If
    Next i
    If isExistsItem(Value) Then
        Call deleteItemAll(Value)
    End If
End Sub

Private Function isExistsItem(Value As Variant) As Boolean
    Dim i As Long
    Dim x As Long
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        If m_Array(i) = Value Then
            isExistsItem = True
            Exit Function
        End If
    Next i
    isExistsItem = False
End Function

Public Function MinNo() As Long
    MinNo = LBound(m_Array) + 1
End Function

Public Function MaxNo() As Long
    MaxNo = UBound(m_Array)
End Function

Private Function isOutBound(ItemNo As Long) As Boolean
    If LBound(m_Array) + 1 > ItemNo Or UBound(m_Array) < ItemNo Then
        isOutBound = True
    Else
        isOutBound = False
    End If
End Function

Public Function isInBound(ItemNo As Long) As Boolean
    isInBound = Not isOutBound(ItemNo)
End Function

Public Sub CallSubroutineByElement(SubRoutineName As String, _
                Optional arg1 As Variant, _
                Optional arg2 As Variant, _
                Optional arg3 As Variant, _
                Optional arg4 As Variant, _
                Optional arg5 As Variant, _
                Optional arg6 As Variant, _
                Optional arg7 As Variant, _
                Optional arg8 As Variant, _
                Optional arg9 As Variant)
    Dim i As Long
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        Application.Run ActiveWorkbook.Name & "!" & SubRoutineName, m_Array(i), arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9
    Next
    
End Sub

Public Sub CallSubroutineByElement2(SubRoutineNameWithWorkbookName As String, _
                Optional arg1 As Variant, _
                Optional arg2 As Variant, _
                Optional arg3 As Variant, _
                Optional arg4 As Variant, _
                Optional arg5 As Variant, _
                Optional arg6 As Variant, _
                Optional arg7 As Variant, _
                Optional arg8 As Variant, _
                Optional arg9 As Variant)
    Dim i As Long
    
    For i = LBound(m_Array) + 1 To UBound(m_Array)
        Application.Run SubRoutineNameWithWorkbookName, m_Array(i), arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9
    Next
    
End Sub






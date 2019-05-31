VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private m_Array() As Variant
Private index As Long


Private Sub Class_Initialize()
    'Ç±Ç±Ç…èàóùÇèëÇ≠
    index = 0
    ReDim Preserve m_Array(0)
End Sub

Private Sub Class_Terminate()
    'Ç±Ç±Ç…èàóùÇèëÇ≠
    'ReDim m_Array(0)
    'set m_Array = Nothing
    'Unload m_Array
    
End Sub

Public Function Push(Value As Variant) As Variant
    index = index + 1
    ReDim Preserve m_Array(index)
    m_Array(index) = Value
    Push = Value
End Function

Public Function Pop() As Variant
    Dim retValue As Variant
    
    retValue = m_Array(index)
    
    index = index - 1
    If index < 1 Then
        index = 0
    End If
    
    ReDim Preserve m_Array(index)
    
    Pop = retValue
End Function

Public Function ReadHead() As Variant
    Dim retValue As Variant
    
    retValue = m_Array(index)

    ReadHead = retValue

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

    ToString = strg

End Function


Public Function ToReverseString(Optional Separator As SeparatorChar = SeparatorChar.vbeComma, Optional ElseSeparator As String = " ") As String
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
    
    ToReverseString = strg

End Function

Public Sub Clear()
    ReDim m_Array(0)
    
    'Set ÉNÉäÉA = Me
End Sub

Public Sub CallSubroutineFromBottom(SubrotineName As String, _
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
    
    For i = 1 To UBound(m_Array)
        Application.Run ActiveWorkbook.Name & "!" & SubrotineName, m_Array(i), arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9
    Next i
    
End Sub

Public Sub CallSubroutineFromBottom2(SubrotineNameWithWorkbookName As String, _
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
    
    For i = 1 To UBound(m_Array)
        Application.Run SubrotineNameWithWorkbookName, m_Array(i), arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9
    Next i
    
End Sub

Public Sub CallSubroutineFromTop(SubrotineName As String, _
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
    
    For i = UBound(m_Array) To 1 Step -1
        Application.Run ActiveWorkbook.Name & "!" & SubrotineName, m_Array(i), arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9
    Next i
    
End Sub

Public Sub CallSubroutineFromTop2(SubrotineNameWithWorkbookName As String, _
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
    
    For i = UBound(m_Array) To 1 Step -1
        Application.Run SubrotineNameWithWorkbookName, m_Array(i), arg1, arg2, arg3, arg4, arg5, arg6, arg7, arg8, arg9
    Next i
    
End Sub







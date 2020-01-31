Attribute VB_Name = "Range�W�����Z"
Option Explicit

'�Q�lURL�Fhttps://mohayonao.hatenadiary.org/entry/20080617/1213712469

' �a�W��
' Union2(ParamArray ArgList() As Variant) As Range
'
' �ϏW��
' Intersect2(ParamArray ArgList() As Variant) As Range
'
' ���W��
' Except2(ByRef SourceRange As Variant, ParamArray ArgList() As Variant) As Range
'
' �Z���͈͂̔��]
' Invert2(ByRef SourceRange As Variant) As Range
'
'
'# �����̃Z�� ArgList �̘a�W����Ԃ�
'# Application.Union �̊g���� Nothing �ł�OK
Public Function Union2(ParamArray ArgList() As Variant) As Range

    Dim buf As Range
    
    Dim i As Long
    For i = 0 To UBound(ArgList)
        If TypeName(ArgList(i)) = "Range" Then
            If buf Is Nothing Then
                Set buf = ArgList(i)
            Else
                Set buf = Application.Union(buf, ArgList(i))
            End If
        End If
    Next
    
    Set Union2 = buf

End Function
Public Function range�W���a(ParamArray ArgList() As Variant) As Range

    Dim buf As Range
    
    Dim i As Long
    For i = 0 To UBound(ArgList)
        If TypeName(ArgList(i)) = "Range" Then
            If buf Is Nothing Then
                Set buf = ArgList(i)
            Else
                Set buf = Application.Union(buf, ArgList(i))
            End If
        End If
    Next
    
    Set range�W���a = buf

End Function


'# �����̃Z�� ArgList �̐ϏW����Ԃ�
'# Application.Intersect �̊g���� Nothing �ł�OK
Public Function Intersect2(ParamArray ArgList() As Variant) As Range

    Dim buf As Range
    
    Dim i As Long
    
    For i = 0 To UBound(ArgList)
        If Not TypeName(ArgList(i)) = "Range" Then
            Exit Function
        ElseIf buf Is Nothing Then
            Set buf = ArgList(i)
        Else
            Set buf = Application.Intersect(buf, ArgList(i))
        End If
        
        If buf Is Nothing Then Exit Function
    Next
    
    Set Intersect2 = buf

End Function
Public Function range�W����(ParamArray ArgList() As Variant) As Range

    Dim buf As Range
    
    Dim i As Long
    
    For i = 0 To UBound(ArgList)
        If Not TypeName(ArgList(i)) = "Range" Then
            Exit Function
        ElseIf buf Is Nothing Then
            Set buf = ArgList(i)
        Else
            Set buf = Application.Intersect(buf, ArgList(i))
        End If
        
        If buf Is Nothing Then Exit Function
    Next
    
    Set range�W���� = buf

End Function


'# SourceRange ���� ArgList ���������������W����Ԃ�
'# (SourceRange �� ���]���� ArgList �Ƃ̐ϏW����Ԃ�)
Public Function Except2 _
    (ByRef SourceRange As Variant, ParamArray ArgList() As Variant) As Range

    If TypeName(SourceRange) = "Range" Then
        
        Dim buf As Range
        
        Set buf = SourceRange
        
        Dim i As Long
        
        For i = 0 To UBound(ArgList)
            If TypeName(ArgList(i)) = "Range" Then
                Set buf = Intersect2(buf, Invert2(ArgList(i)))
            End If
        Next
        
        Set Except2 = buf
        
    End If

End Function

Public Function range�W���� _
    (ByRef SourceRange As Variant, ParamArray ArgList() As Variant) As Range

    If TypeName(SourceRange) = "Range" Then
        
        Dim buf As Range
        
        Set buf = SourceRange
        
        Dim i As Long
        
        For i = 0 To UBound(ArgList)
            If TypeName(ArgList(i)) = "Range" Then
                Set buf = Intersect2(buf, Invert2(ArgList(i)))
            End If
        Next
        
        Set range�W���� = buf
        
    End If

End Function


'# SourceRange �̑I��͈͂𔽓]����
Public Function range�W�����](ByRef SourceRange As Variant) As Range
    Set range�W�����] = Invert2(SourceRange)
End Function
Public Function Invert2(ByRef SourceRange As Variant) As Range

    If Not TypeName(SourceRange) = "Range" Then Exit Function
    
    Dim sh As Worksheet
    Set sh = SourceRange.Parent
    
    Dim buf As Range
    Set buf = SourceRange.Parent.Cells
        
    Dim a As Range
    For Each a In SourceRange.Areas
        
        Dim AreaTop    As Long
        Dim AreaBottom As Long
        Dim AreaLeft   As Long
        Dim AreaRight  As Long
        
        AreaTop = a.Row
        AreaBottom = AreaTop + a.Rows.Count - 1
        AreaLeft = a.Column
        AreaRight = AreaLeft + a.Columns.Count - 1
        
        
        '������
        '���~��
        '������  ���̕���
        Dim RangeLeft   As Range
        Set RangeLeft = GetRangeWithPosition(sh, _
            sh.Cells.Row, sh.Cells.Column, sh.Rows.Count, AreaLeft - 1)
        '   Top           Left             Bottom         Right
        
        '������
        '���~��
        '������  ���̕���
        Dim RangeRight  As Range
        Set RangeRight = GetRangeWithPosition(sh, _
            sh.Cells.Row, AreaRight + 1, sh.Rows.Count, sh.Columns.Count)
        '   Top           Left           Bottom         Right
        
        
        '������
        '���~��
        '������  ���̕���
        Dim RangeTop    As Range
        Set RangeTop = GetRangeWithPosition(sh, _
            sh.Cells.Row, AreaLeft, AreaTop - 1, AreaRight)
        '   Top           Left      Bottom       Right
        
        
        '������
        '���~��
        '������  ���̕���
        Dim RangeBottom As Range
        Set RangeBottom = GetRangeWithPosition(sh, _
            AreaBottom + 1, AreaLeft, sh.Rows.Count, AreaRight)
        '   Top              Left      Bottom         Right
        
        
        Set buf = Intersect2(buf, _
            Union2(RangeLeft, RangeRight, RangeTop, RangeBottom))
        
    Next
    
    Set Invert2 = buf

End Function


'# �l�����w�肵�� Range �𓾂�
Private Function GetRangeWithPosition(ByRef sh As Worksheet, _
    ByVal Top As Long, ByVal Left As Long, _
    ByVal Bottom As Long, ByVal Right As Long) As Range
    
    '# ��������
    If Top > Bottom Or Left > Right Then
        Exit Function
    ElseIf Top < 0 Or Left < 0 Then
        Exit Function
    ElseIf Bottom > Cells.Rows.Count Or Right > Cells.Columns.Count Then
        Exit Function
    End If
    
    Set GetRangeWithPosition _
        = sh.Range(sh.Cells(Top, Left), sh.Cells(Bottom, Right))

End Function

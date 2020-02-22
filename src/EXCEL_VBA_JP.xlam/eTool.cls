VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Cell As eCell
Private m_Sheet As eSheet
Private m_Book As eBook


'Public g_���� As New ����
Public Function Inst()
    Inst = g_����
End Function




Private Sub Class_Initialize()
    '�����ɏ���������
    Set m_Cell = New eCell
    Set m_Sheet = New eSheet
    Set m_Book = New eBook
End Sub

Private Sub Class_Terminate()
    '�����ɏ���������
    Set m_Book = Nothing
    Set m_Sheet = Nothing
    Set m_Cell = Nothing
End Sub

'���悭�Ȃ�
'Public Property Get ���Έʒu(ByRef �Z�� As ����, Optional �s As Integer = 0, Optional �� As Integer = 0) As Range
'    Set ���Έʒu = �Z��.Offset(�s, ��)
'End Property

'Public Property Get �Z��01(Optional �Z���� As String = "A1", Optional �V�[�g���� As Integer = 1) As Range
'    Set �Z��01 = Worksheets(�V�[�g����).Range(�Z����).Offset(0, 0)
'End Property

'Public Property Get �Z��02(Optional �Z���� As String = "A1", Optional �V�[�g���� As Integer = 1) As �Z��
'    Set m_Cell.�Z�b�g = Worksheets(�V�[�g����).Range(�Z����).Offset(0, 0)
'    m_Cell.�Z���͈� = �Z����
'    m_Cell.�V�[�g���� = �V�[�g����
'    Set �Z��02 = m_Cell
'End Property

'Public Property Get �Z��03(Optional �Z���� As String = "A1", Optional �V�[�g���� As Integer = 1) As �Z��
'    Set m_Cell.�Z�b�g = Worksheets(�V�[�g����).Range(�Z����).Offset(0, 0)
''    m_Cell.�Z���͈� = �Z����
''    m_Cell.�V�[�g���� = �V�[�g����
'    Set �Z��03 = m_Cell
'End Property

'Public Property Get �Z��04(Optional �Z���� As String = "A1", Optional �V�[�g���� As Integer = 0) As �Z��
'    If �V�[�g���� <> 0 Then
'        Set m_Cell.�Z�b�g = Worksheets(�V�[�g����).Range(�Z����).Offset(0, 0)
'    Else
'        Set m_Cell.�Z�b�g = ActiveSheet.Range(�Z����).Offset(0, 0)
'    End If
''    m_Cell.�Z���͈� = �Z����
''    m_Cell.�V�[�g���� = �V�[�g����
'    Set �Z��04 = m_Cell
'End Property

'---- �Z��04�x�[�X ---------------------------
Public Function Cell(Optional ByVal CellAddress As String = "A1", Optional ByVal SheetNo As Long = 1) As eCell
    Dim wCell As New eCell
    If SheetNo > 0 Then
        Set wCell.ToSet = Worksheets(SheetNo).Range(CellAddress)
    Else
        Set wCell.ToSet = ActiveSheet.Range(CellAddress)
    End If
    Set Cell = wCell
    Set wCell = Nothing
End Function

'Public Property Get �Z���s��w��(�s As Long, �� As Long) As �Z��
'    Dim w�Z�� As New �Z��
'    If �V�[�g���� > 0 Then
'        Set w�Z��.�Z�b�g = Worksheets(�V�[�g����).Cells(�s, ��)
'    Else
'        Set w�Z��.�Z�b�g = ActiveSheet.Cells(�s, ��)
'    End If
'    Set �Z�� = w�Z��
'End Property

Public Property Get CurrentCell() As eCell
    Set m_Cell.ToSet = ActiveCell
    'm_Cell.�Z���͈� =
    m_Cell.SheetNo = 0
    Set CurrentCell = m_Cell
End Property

Public Property Get CurrentTable() As eCell
    Set m_Cell.ToSet = ActiveCell.CurrentRegion
    'm_Cell.�Z���͈� =
    m_Cell.SheetNo = 0
    Set CurrentTable = m_Cell
End Property


Public Function GetSheet(Optional SheetNo As Integer = 0) As eSheet
    If SheetNo <> 0 Then
        Set m_Sheet.ToSet = Worksheets(SheetNo)
    Else
        Set m_Sheet.ToSet = ActiveSheet
    End If
    m_Sheet.SheetNo = SheetNo
    Set GetSheet = m_Sheet
End Function

Public Property Get CurrentSheet() As eSheet
    Set m_Sheet.ToSet = ActiveSheet
    Set CurrentSheet = m_Sheet
End Property


Public Property Get Names() As Names
    Set Names = m_Book.GetNames
End Property

Public Function CellAtRowCol(Optional ByVal Row As Long = 0, Optional ByVal Col As Long = 0) As eCell
    Set m_Cell.ToSet = ActiveSheet.Cells(Row, Col).Offset(0, 0)
    Set CellAtRowCol = m_Cell
End Function


Public Function IsCellBlank(Row As Long, Col As Long, Optional SheetNo As Long = -1) As Boolean
    If SheetNo > 0 Then
        IsCellBlank = (Trim(Worksheets(SheetNo).Cells(Row, Col).Value2) = "")
    Else
        IsCellBlank = (Trim(ActiveSheet.Cells(Row, Col).Value2) = "")
    End If
End Function

Public Function IsNotCellBlank(Row As Long, Col As Long, Optional SheetNo As Long = -1) As Boolean
    If SheetNo > 0 Then
        IsNotCellBlank = Trim(Worksheets(SheetNo).Cells(Row, Col).Value2) <> ""
    Else
        IsNotCellBlank = Trim(ActiveSheet.Cells(Row, Col).Value2) <> ""
    End If
End Function

Public Function IsCellPartialMatch(PartialMatchEString As eString, Row As Long, Col As Long, Optional SheetNo As Long = -1) As Boolean
    If SheetNo > 0 Then
        IsCellPartialMatch = (InStr(1, Trim(Worksheets(SheetNo).Cells(Row, Col).Value2), PartialMatchEString.ToVBA) > 0)
    Else
        IsCellPartialMatch = (InStr(1, Trim(ActiveSheet.Cells(Row, Col).Value2), PartialMatchEString.ToVBA) > 0)
    End If
End Function

Public Function IsNotCellPartialMatch(PartialMatchEString As eString, Row As Long, Col As Long, Optional SheetNo As Long = -1) As Boolean
    If SheetNo > 0 Then
        IsNotCellPartialMatch = (InStr(1, Trim(Worksheets(SheetNo).Cells(Row, Col).Value2), PartialMatchEString.ToVBA) <= 0)
    Else
        IsNotCellPartialMatch = (InStr(1, Trim(ActiveSheet.Cells(Row, Col).Value2), PartialMatchEString.ToVBA) <= 0)
    End If
End Function

Public Function IsCellPartialMatchString(PartialMatchString As String, Row As Long, Col As Long, Optional SheetNo As Long = -1) As Boolean
    If SheetNo > 0 Then
        IsCellPartialMatchString = (InStr(1, Trim(Worksheets(SheetNo).Cells(Row, Col).Value2), PartialMatchString) > 0)
    Else
        IsCellPartialMatchString = (InStr(1, Trim(ActiveSheet.Cells(Row, Col).Value2), PartialMatchString) > 0)
    End If
End Function

Public Function IsNotCellPartialMatchString(PartialMatchString As String, Row As Long, Col As Long, Optional SheetNo As Long = -1) As Boolean
    If SheetNo > 0 Then
        IsNotCellPartialMatchString = (InStr(1, Trim(Worksheets(SheetNo).Cells(Row, Col).Value2), PartialMatchString) <= 0)
    Else
        IsNotCellPartialMatchString = (InStr(1, Trim(ActiveSheet.Cells(Row, Col).Value2), PartialMatchString) <= 0)
    End If
End Function

Public Sub ShowMessage(Message As Variant)
    MsgBox Message
    
End Sub
Function GetWorkBookName(Optional IsWithConnectStyring As Boolean = True) As String
    If IsWithConnectStyring Then
        GetWorkBookName = ActiveWorkbook.Name & "!"
    Else
        GetWorkBookName = ActiveWorkbook.Name
    End If
        
End Function


 









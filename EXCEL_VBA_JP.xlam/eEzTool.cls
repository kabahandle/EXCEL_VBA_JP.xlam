VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eEzTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Cell As eCell
Private m_Sheet As eSheet
Private m_Book As eBook


'Public g_道具 As New 道具
Public Function Inst()
    Inst = g_道具
End Function




Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_Cell = New eCell
    Set m_Sheet = New eSheet
    Set m_Book = New eBook
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_Book = Nothing
    Set m_Sheet = Nothing
    Set m_Cell = Nothing
End Sub

'↓よくない
'Public Property Get 相対位置(ByRef セル As 道具, Optional 行 As Integer = 0, Optional 列 As Integer = 0) As Range
'    Set 相対位置 = セル.Offset(行, 列)
'End Property

'Public Property Get セル01(Optional セル名 As String = "A1", Optional シート順番 As Integer = 1) As Range
'    Set セル01 = Worksheets(シート順番).Range(セル名).Offset(0, 0)
'End Property

'Public Property Get セル02(Optional セル名 As String = "A1", Optional シート順番 As Integer = 1) As セル
'    Set m_Cell.セット = Worksheets(シート順番).Range(セル名).Offset(0, 0)
'    m_Cell.セル範囲 = セル名
'    m_Cell.シート順番 = シート順番
'    Set セル02 = m_Cell
'End Property

'Public Property Get セル03(Optional セル名 As String = "A1", Optional シート順番 As Integer = 1) As セル
'    Set m_Cell.セット = Worksheets(シート順番).Range(セル名).Offset(0, 0)
''    m_Cell.セル範囲 = セル名
''    m_Cell.シート順番 = シート順番
'    Set セル03 = m_Cell
'End Property

'Public Property Get セル04(Optional セル名 As String = "A1", Optional シート順番 As Integer = 0) As セル
'    If シート順番 <> 0 Then
'        Set m_Cell.セット = Worksheets(シート順番).Range(セル名).Offset(0, 0)
'    Else
'        Set m_Cell.セット = ActiveSheet.Range(セル名).Offset(0, 0)
'    End If
''    m_Cell.セル範囲 = セル名
''    m_Cell.シート順番 = シート順番
'    Set セル04 = m_Cell
'End Property

'---- セル04ベース ---------------------------
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
Public Function ezCell(Optional ByVal CellAddress As String = "A1", Optional ByVal SheetNo As Long = 1) As eEzCell
    Dim wCell As New eEzCell
    If SheetNo > 0 Then
        Set wCell.ToSet = Worksheets(SheetNo).Range(CellAddress)
    Else
        Set wCell.ToSet = ActiveSheet.Range(CellAddress)
    End If
    Set ezCell = wCell
    Set wCell = Nothing
End Function

'Public Property Get セル行列指定(行 As Long, 列 As Long) As セル
'    Dim wセル As New セル
'    If シート順番 > 0 Then
'        Set wセル.セット = Worksheets(シート順番).Cells(行, 列)
'    Else
'        Set wセル.セット = ActiveSheet.Cells(行, 列)
'    End If
'    Set セル = wセル
'End Property

Public Property Get CurrentCell() As eCell
    Set m_Cell.ToSet = ActiveCell
    'm_Cell.セル範囲 =
    m_Cell.SheetNo = 0
    Set CurrentCell = m_Cell
End Property
Public Property Get CurrentEzCell() As eEzCell
    Set m_Cell.ToSet = ActiveCell
    'm_Cell.セル範囲 =
    m_Cell.SheetNo = 0
    Set CurrentEzCell = m_Cell
End Property


Public Property Get CurrentTable() As eCell
    Set m_Cell.ToSet = ActiveCell.CurrentRegion
    'm_Cell.セル範囲 =
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

Public Function EzCellAtRowCol(Optional ByVal Row As Long = 0, Optional ByVal Col As Long = 0) As eEzCell
    Set m_Cell.ToSet = ActiveSheet.Cells(Row, Col).Offset(0, 0)
    Set EzCellAtRowCol = m_Cell
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


 











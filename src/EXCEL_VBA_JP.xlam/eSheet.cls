VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Sheet As Worksheet
Private m_SheetNo As Long


Private Sub Class_Initialize()
    'Ç±Ç±Ç…èàóùÇèëÇ≠
    Set m_Sheet = ActiveSheet
End Sub

Private Sub Class_Terminate()
    'Ç±Ç±Ç…èàóùÇèëÇ≠
    Set m_Sheet = Nothing
End Sub

Public Property Set ToSet(sht As Worksheet)
    Set m_Sheet = sht
End Property

Public Property Let SheetNo(lngSheetNo As Long)
    m_SheetNo = lngSheetNo
End Property

Public Property Get SheetNo() As Long
    SheetNo = m_SheetNo
End Property

Public Property Get Name() As Worksheet
    Set Name = m_Sheet
End Property

Public Property Set Name(Sheet As Worksheet)
    Set m_Sheet = Sheet
End Property


'---- Rows
Public Property Get Row(ByVal RowNo As Long) As eCell
    Dim wCell As New eCell
    Set wCell.ToSet = m_Sheet.Rows(RowNo)
    Set Row = wCell
End Property

Public Sub IsRowHidden(ByVal RowNo As Long, ByVal isHidden As Boolean)
    m_Sheet.Rows(RowNo).Hidden = isHidden
End Sub

Public Sub SetRowHeight(ByVal RowNo As Long, ByVal RowHeight As Variant)
    m_Sheet.Rows(RowNo).RowHeight = RowHeight
End Sub

Public Sub SetRowHightAsAutoFit(ByVal RowNo As Long)
    m_Sheet.Rows(RowNo).AutoFit
End Sub

Public Function GetRowHeight(ByVal RowNo As Long) As Variant
    GetRowHeight = m_Sheet.Rows(RowNo).RowHeight
End Function

'---- Cols
Public Property Get Col(ByVal ColNo As Long) As eCell
    Dim wCell As New eCell
    Set wCell.ToSet = m_Sheet.Columns(ColNo)
    Set Col = wCell
    Set wCell = Nothing
End Property

Public Sub ColIsHidden(ByVal ColNo As Long, ByVal isHidden As Boolean)
    m_Sheet.Columns(ColNo).Hidden = isHidden
End Sub

Public Sub SetColWidth(ByVal ColNo As Long, ByVal ColWidth As Variant)
    m_Sheet.Columns(ColNo).ColumnWidth = ColWidth
End Sub

Public Function GetColWidth(ByVal ColNo As Long) As Variant
    GetColWidth = m_Sheet.Columns(ColNo).ColumnWidth
End Function

Public Sub SetColWidthAsAutoFit(ByVal ColNo As Long)
    m_Sheet.Columns(ColNo).AutoFit
End Sub

Public Sub Activate()
    m_Sheet.Activate
End Sub






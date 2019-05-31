VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Const DebugMode = False

Private m_Range As Range
Private m_CellName As String
Private m_SheetNo As Long
Private m_PrevSelection As Range

Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_Range = ActiveSheet.Range("A1")
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_Range = Nothing
End Sub

'--- 日本語関数用
'Public Property Set Range(レンジ As Range)
'    Set m_Range = レンジ
'End Property
'
'Public Property Get Range() As Range
'    Set Range = m_Range
'End Property

Public Property Set ToJpFunc(rngRange As Range)
    Set m_Range = rngRange
End Property

Public Property Get ToJpFunc() As Range
    Set ToJpFunc = m_Range
End Property

Public Property Get ToFFunc() As Range
    Set ToFFunc = m_Range
End Property

'--- end of 日本語間枢要

Public Property Set ToSet(rng As Range)
    Set m_Range = rng
End Property

Public Property Let CellName(strgCellName As String)
    m_CellName = strgCellName
End Property

Public Property Get CellName() As String
    CellName = m_CellName
End Property

Public Property Let SheetNo(ByVal lngSheetNo As Long)
    m_SheetNo = lngSheetNo
End Property

Public Property Get CellRange() As Range
    Set CellRange = m_Range
End Property

Public Function RelativePosition(Optional ByVal RowNo As Long = 0, Optional ByVal ColNo As Long = 0) As eCell
    Dim wCell As New eCell
    Set wCell.ToSet = m_Range.Offset(RowNo, ColNo)
    Set RelativePosition = wCell
    Set wCell = Nothing
End Function

Public Function GetSelection() As eCell
    Dim w_Area As New eCell
    Dim w_range As Range
    
    m_Range.Select
    
    With Selection
        Set w_range = ActiveSheet.Range(.Cells.Address)
        Set w_Area.ToSet = w_range
    End With
    
    Set GetSelection = w_Area
    
    Set w_Area = Nothing
    Set w_range = Nothing
    
End Function

Public Sub Insert(Optional ShiftOrientValue As ShiftOrient = vbeShiftForwordToRight)
    m_Range.Insert ShiftOrientValue
End Sub

Public Sub Delete(Optional ShiftOrientValue As ShiftOrient = vbeShiftForwordToRight)
    m_Range.Delete ShiftOrientValue
End Sub

Public Sub Clear()
    m_Range.Clear
End Sub

Public Sub ClearContents()
    'm_Range.ClearComments
    m_Range.ClearContents
End Sub

Public Sub ClearFormats()
    m_Range.ClearFormats
End Sub

Public Sub ClearComments()
    m_Range.ClearComments
End Sub

Public Sub Fill(CellAreaToFill As eCell, FillMethodValue As FillMethod)
    m_Range.AutoFill Destination:=CellAreaToFill.CellRange, Type:=FillMethodValue
End Sub

Public Sub Cut(Optional MoveTarget As eCell = Nothing)
    If Not MoveTarget Is Nothing Then
        m_Range.Cut MoveTarget.CellRange
    Else
        m_Range.Cut
    End If
End Sub

Public Sub Copy(Optional CopyTarget As eCell = Nothing)
    If Not CopyTarget Is Nothing Then
        m_Range.Copy CopyTarget.CellRange
    Else
        m_Range.Copy
    End If
End Sub

Public Sub PasteSpecial( _
    Optional PasteMethodValue As PasteMethod = vbePasteAll, _
    Optional ByVal PasteToChangeTheMatrix As Boolean = False)
#If DebugMode = False Then
    On Error Resume Next
#End If
    m_Range.PasteSpecial PasteMethodValue, , , PasteToChangeTheMatrix
    On Error GoTo 0
End Sub

Public Sub Paste()
    m_Range.Select
    ActiveSheet.Paste
End Sub

Public Sub Merge()
    m_Range.Merge
End Sub

Public Function ToEnd(Optional Orient As OrientForEnding = vbeEndDown) As eCell
    Select Case Orient
    Case OrientForEnding.vbeEndUp
        Set m_Range = m_Range.End(xlUp)
    Case OrientForEnding.vbeEndDown
        Set m_Range = m_Range.End(xlDown)
    Case OrientForEnding.vbeEndRight
        Set m_Range = m_Range.End(xlToRight)
    Case OrientForEnding.vbeEndLeft
        Set m_Range = m_Range.End(xlToLeft)
    Case Else
        Set m_Range = m_Range.End(xlDown)
    End Select
    
    Set ToEnd = Me
End Function

Public Function GetSpecialCell( _
    Optional SelectionMethodForCellValue As SelectionMethodForCell = xlCellTypeBlanks, _
    Optional ConditionsForSelectingCellsValue As ConditionsForSelectingCells = vbeValuesOfText) _
    As eCell
    On Error GoTo ERR1
    
    Select Case SelectionMethodForCellValue
    Case SelectionMethodForCell.vbeExistsTypeFormulas
        Set m_Range = m_Range.SpecialCells(SelectionMethodForCellValue, ConditionsForSelectingCellsValue)
    Case SelectionMethodForCell.vbeExistsTypeConstants
        Set m_Range = m_Range.SpecialCells(SelectionMethodForCellValue, ConditionsForSelectingCellsValue)
    Case Else
        Set m_Range = m_Range.SpecialCells(SelectionMethodForCellValue)
    End Select
    
    Set GetSpecialCell = Me
    
    Exit Function
    
ERR1:
    Set GetSpecialCell = Me
End Function



'---- Value2
Public Property Get Value() As Variant
    Value = m_Range.Value2
End Property

'---- Value2
Public Property Get Address( _
    Optional ColAddressDesignation As AddressDesignation = vbeRelativeAddress, _
    Optional RowAddressDesignation As AddressDesignation = vbeRelativeAddress) _
    As Variant
    Address = m_Range.Address(ColAddressDesignation, RowAddressDesignation)
End Property

Public Property Let Value(ValueData As Variant)
    m_Range.Value2 = ValueData
End Property

'---- Formula
Public Property Get Formula() As Variant
    Formula = m_Range.Formula
End Property

Public Property Let Formula(FormulaValue As Variant)
    m_Range.Formula = FormulaValue
End Property

''---- Rows
'Public Property Get 行全体(行番号 As Variant) As セル
'    Set m_Range = m_Range.Rows(行番号)
'    Set 行全体 = Me
'End Property
'
''---- Cols
'Public Property Get 列全体(列番号 As Variant) As セル
'    Set m_Range = m_Range.Columns(列番号)
'    Set 列全体 = Me
'End Property


'---- Interior
Public Property Get Style() As eStyle
    Dim wStyle As New eStyle
    Set wStyle.ToSet = m_Range.Interior
    Set Style = wStyle
    Set wStyle = Nothing
End Property

'---- Count
Public Property Get Count() As Long
    Count = m_Range.Count
End Property

'---- EntireRow
Public Property Get EntireRows() As eCell
    Dim wCell As New eCell
    Set wCell.ToSet = m_Range.EntireRow
    Set EntireRows = wCell
    Set wCell = Nothing
End Property

'---- EntireCols
Public Property Get EntireCols() As eCell
    Dim wCell As New eCell
    Set wCell.ToSet = m_Range.EntireColumn
    Set EntireCols = wCell
    Set wCell = Nothing
End Property

'---- Name
Public Property Get Name() As String
    Name = m_Range.Name
End Property

Public Property Let Name(NameValue As String)
    m_Range.Name = NameValue
End Property

Public Property Let Comment(CommentValue As String)
    If TypeName(m_Range.Comment) = "Comment" Then
        m_Range.ClearComments
    End If
    m_Range.AddComment CommentValue
End Property

Public Property Get Comment() As String
    Comment = m_Range.Comment
End Property

' セルを非表示にするとエラー
'Public Property Let 非表示(ByVal 非表示 As Boolean)
'    m_Range.Hidden = 非表示
'End Property

Public Property Get Height() As Variant
    Height = m_Range.Height
End Property

'セルの高さと幅は変更することができない。（行高さ・列幅の変更で代用すること）
'Public Property Let 高さ(ByVal 高さ値 As Variant)
'    m_Range.Offset(0, 0).Height = 高さ値
'End Property

Public Property Get Width() As Variant
    Width = m_Range.Width
End Property

'セルの高さと幅は変更することができない。（行高さ・列幅の変更で代用すること）
'Public Property Let 幅(ByVal 幅値 As Variant)
'    m_Range.Offset(0, 0).Width = 幅値
'End Property


'--- NumberFormatLocal
Public Property Get NumberFormatLocal() As Variant
    NumberFormatLocal = m_Range.NumberFormatLocal
End Property

Public Property Let NumberFormatLocal(ByVal NumberFormatLocalValue As Variant)
    m_Range.NumberFormatLocal = NumberFormatLocalValue
End Property

' 有効にするとLetがエラー。戻り値の型が違うため。
'Public Property Get 表示形式パターン() As Variant
'    表示形式 = m_Range.NumberFormatLocal
'End Property

Public Property Let VisualFormatPattern(ByVal VisualFormatPatternForCellValue As VisualFormatPatternForCell)
    Select Case VisualFormatPatternForCellValue
    Case VisualFormatPatternForCell.vbeCurency
        m_Range.NumberFormatLocal = "\#,##0;[赤]\#,##0"
    Case VisualFormatPatternForCell.vbeOneDecimalPlace
        m_Range.NumberFormatLocal = "0.0"
    Case VisualFormatPatternForCell.vbeTwoDecimalPlace
        m_Range.NumberFormatLocal = "0.00"
    Case VisualFormatPatternForCell.vbeZeroPaddingUpTo4digit
        m_Range.NumberFormatLocal = "0000"
    Case VisualFormatPatternForCell.vbeZeroPaddingUpTo8digit
        m_Range.NumberFormatLocal = "00000000"
    Case VisualFormatPatternForCell.vbeAnnoDomini
        m_Range.NumberFormatLocal = "yyyy/mm/dd"
    Case VisualFormatPatternForCell.vbeAnnoDominiWithDate
        m_Range.NumberFormatLocal = "yyyy/mm/dd(aaa)"
    Case VisualFormatPatternForCell.vbeJapaneseCalendar
        m_Range.NumberFormatLocal = "gggr""年""m""月""d""日"""
    Case VisualFormatPatternForCell.vbeJapaneseCalendarWithDate
        m_Range.NumberFormatLocal = "gggr""年""m""月""d""日""(aaa)"
    Case VisualFormatPatternForCell.vbeDateAndTime
        m_Range.NumberFormatLocal = "hh:mm"
    Case VisualFormatPatternForCell.vbeDareAndTimeAndMinutes
        m_Range.NumberFormatLocal = "hh:mm:ss"
    Case VisualFormatPatternForCell.vbeDateAndTimeWithAMandPM
        m_Range.NumberFormatLocal = "hh:mm AM/PM"
    End Select
End Property

'--- HorizentalAlignment
Public Property Get HorizontalAlignment() As HorizentalPositionForCell
    HorizontalAlignment = m_Range.HorizontalAlignment
End Property
'
Public Property Let HorizontalAlignment(ByVal HorizentalPositionForCellValue As HorizentalPositionForCell)
    m_Range.HorizontalAlignment = HorizentalPositionForCellValue
End Property

'--- VertucalAlignment
Public Property Get VerticalAlignment() As VerticalPositionForCell
    VerticalAlignment = m_Range.VerticalAlignment
End Property
'
Public Property Let VerticalAlignment(ByVal VerticalPositionForCellValue As VerticalPositionForCell)
    m_Range.VerticalAlignment = VerticalPositionForCellValue
End Property


'--- Orientation
Public Property Get Orientation() As Long
    Orientation = m_Range.Orientation
End Property
'
Public Property Let Orientation(ByVal OrientationValue As Long)
    If OrientationValue > 90 Then
        OrientationValue = 90
    ElseIf OrientationValue < -90 Then
        OrientationValue = -90
    End If
    m_Range.Orientation = OrientationValue
End Property

'Public Property Get OrientationPattern() As Variant
'    OrientationPattern = m_Range.Orientation
'End Property
'
Public Property Let OrientationPattern(ByVal CellDegreValue As CellDegree)
    m_Range.Orientation = CellDegreValue
End Property



'---- Interior
Public Property Get Font() As eFont
    Dim w_Font As New eFont
    Set w_Font.ToSet = m_Range.Font
    Set Font = w_Font
    Set w_Font = Nothing
End Property

'---- borders
Public Property Get Border(Optional BorderPositionValue As BorderPosition = BorderPosition.vbeEdgeLeft) As eBorder
    Dim w_Border As New eBorder
    Set w_Border.ToSet = m_Range.Borders(BorderPositionValue)
    Set Border = w_Border
    Set w_Border = Nothing
End Property

Public Sub BorderEnclose( _
        BoderLineValue As StyleOfBoderLine, _
        LineWeight As BorderThickness, _
        ColorIndexValue As PatternOfColorIndex)
    Call m_Range.BorderAround(BoderLineValue, LineWeight, ColorIndexValue)
End Sub

'Public Property Let 罫線(罫線位置値 As 罫線位置)
'    m_Border.Borders = 線の太さ値
'End Property


'---- Photonic
Public Property Get Phonetic() As String
    Phonetic = Application.GetPhonetic(m_Range.Offset(0, 0).Value2)
End Property

Public Sub SetPhoneticVisible()
    m_Range.SetPhonetic
    m_Range.Phonetics.Visible = True
End Sub

Public Property Get IsWhiteSpace() As Boolean
    IsWhiteSpace = (Trim(m_Range.Offset(0, 0)).Value2 = "")
End Property

Public Property Get IsNotWhiteSpace() As Boolean
    IsNotWhiteSpace = Trim(m_Range.Offset(0, 0).Value2) <> ""
End Property

Public Function IsPartialMatch(PartialMatchString As eString) As Boolean
    IsPartialMatch = (InStr(1, Trim(m_Range.Offset(0, 0).Value2), PartialMatchString.ToVBA) > 0)
End Function

Public Function IsNotPartialMatch(PartialMatchString As eString) As Boolean
    IsNotPartialMatch = (InStr(1, Trim(m_Range.Offset(0, 0).Value2), PartialMatchString.ToVBA) <= 0)
End Function

Public Function IsPartialMatchString(PartialMatchString As String) As Boolean
    IsPartialMatchString = (InStr(1, Trim(m_Range.Offset(0, 0).Value2), PartialMatchString) > 0)
End Function

Public Property Get eInteger() As eInteger
    Dim ret As Long
    Dim w_Integer As New eInteger
    
    On Error GoTo ERRinteger001
    
    ret = 0
    ret = CLng(m_Range.Offset(0, 0).Value2)
    
ERRinteger001:
    On Error GoTo 0
    
    w_Integer.Name = ret
    
    Set eInteger = w_Integer
    Set w_Integer = Nothing
End Property

Public Property Get IntegerLong() As Long
    Dim ret As Long
    On Error GoTo ERRIntegerLong002
    
    ret = 0
    ret = CLng(m_Range.Offset(0, 0).Value2)
    
ERRIntegerLong002:
    On Error GoTo 0
    
    IntegerLong = ret
    
End Property

Public Property Get eCurrency() As eCurrency
    Dim ret As Currency
    Dim w_Currency As New eCurrency
    
    On Error GoTo ERRCurrency001
    
    ret = 0
    ret = CCur(m_Range.Offset(0, 0).Value2)
    
ERRCurrency001:
    On Error GoTo 0
    
    w_Currency = ret
    
    Set eCurrency = w_Currency
    Set w_Currency = Nothing
End Property

Public Property Get CurrencyCurr() As Currency
    Dim ret As Currency
    On Error GoTo ERRCurrency003
    
    ret = 0
    ret = CCur(m_Range.Offset(0, 0).Value2)
    
ERRCurrency003:
    On Error GoTo 0
    
    CurrencyCurr = ret
End Property

Public Property Get SelectedCellArea( _
    Optional SelectionPatternValue As SelectionPattern = vbeNonePattern, _
    Optional StepNumber As Long = -1, _
    Optional RowY As Long = -1, _
    Optional ColX As Long = -1, _
    Optional OffsetY As Long = 0, _
    Optional OffsetX As Long = 0) _
    As eCell
    Dim w_range As Range
    Dim w_CellArea As New eCell
    Dim startCol As Long
    Dim endCol As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim c As Long
    Dim r As Long
    
    With Selection
        Set m_PrevSelection = ActiveSheet.Range(.Cells.Address)
    End With
    startCol = Selection(1).Column + OffsetX
    endCol = Selection(Selection.Count).Column + OffsetX
    startRow = Selection(1).Row + OffsetY
    endRow = Selection(Selection.Count).Row + OffsetY
    RowY = RowY + OffsetY
    ColX = ColX + OffsetX
    
    
    Set w_range = ActiveSheet.Range("A1").Offset(0, 0)
    
    If SelectionPatternValue = SelectionPattern.vbeOddRows Then
        Selection.Offset(0, 0).Select
        Selection.Resize(1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For r = startRow To endRow Step 2
            Set w_range = Union(w_range, Range(Cells(r, startCol), Cells(r, endCol)))
        Next
        
            
        Set w_CellArea.ToSet = w_range
        Set SelectedCellArea = w_CellArea
    
    ElseIf SelectionPatternValue = SelectionPattern.vbeEvenRows Then
        Selection.Offset(1, 0).Select
        Selection.Resize(1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For r = startRow + 1 To endRow Step 2
            Set w_range = Union(w_range, Range(Cells(r, startCol), Cells(r, endCol)))
        Next
        
            
        Set w_CellArea.ToSet = w_range
        Set SelectedCellArea = w_CellArea
    
    ElseIf SelectionPatternValue = SelectionPattern.vbeRowsByStep And StepNumber >= 1 Then
        Selection.Offset(0, 0).Select
        Selection.Resize(1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For r = startRow To endRow Step StepNumber
            Set w_range = Union(w_range, Range(Cells(r, startCol), Cells(r, endCol)))
        Next
        
            
        Set w_CellArea.ToSet = w_range
        Set SelectedCellArea = w_CellArea
    
    ElseIf SelectionPatternValue = SelectionPattern.vbeEvenCols Then
        Selection.Offset(0, 0).Select
        Selection.Resize(, 1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For c = startCol To endCol Step 2
            Set w_range = Union(w_range, Range(Cells(startRow, c), Cells(endRow, c)))
        Next
        
            
        Set w_CellArea.ToSet = w_range
        Set SelectedCellArea = w_CellArea
    
    ElseIf SelectionPatternValue = SelectionPattern.vbeOddCols Then
        Selection.Offset(0, 1).Select
        Selection.Resize(, 1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For c = startCol + 1 To endCol Step 2
            Set w_range = Union(w_range, Range(Cells(startRow, c), Cells(endRow, c)))
        Next
        
            
        Set w_CellArea.ToSet = w_range
        Set SelectedCellArea = w_CellArea
    
    ElseIf SelectionPatternValue = SelectionPattern.vbeColsByStep And StepNumber >= 1 Then
        Selection.Offset(0, 0).Select
        Selection.Resize(, 1).Select
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
        End With
        
        For c = startCol To endCol Step StepNumber
            Set w_range = Union(w_range, Range(Cells(startRow, c), Cells(endRow, c)))
        Next
        
            
        Set w_CellArea.ToSet = w_range
        Set SelectedCellArea = w_CellArea
    
    Else
        
        If RowY <= 0 And ColX <= 0 Then
        
        ElseIf (RowY >= 1 Or OffsetY = 0) And ColX <= 0 Then
            Selection.Offset(RowY - 1, 0).Select
            Selection.Resize(1).Select
        ElseIf RowY <= 0 And (ColX >= 1 Or OffsetX = 0) Then
            Selection.Offset(0, ColX - 1).Select
            Selection.Resize(, 1).Select
        ElseIf RowY >= 1 And ColX >= 1 Then
            Selection.Offset(RowY - 1, ColX - 1).Select
            Selection.Resize(1, 1).Select
        End If
        
        With Selection
            Set w_range = ActiveSheet.Range(.Cells.Address)
            Set w_CellArea.ToSet = w_range
        End With
        Set SelectedCellArea = w_CellArea
    
    End If
    
    
    m_PrevSelection.Select
    
    Set w_range = Nothing
    Set w_CellArea = Nothing
End Property


Public Sub CallSubroutineByCell(SubRoutineName As String, _
                Optional arg1 As Variant, _
                Optional arg2 As Variant, _
                Optional arg3 As Variant, _
                Optional arg4 As Variant, _
                Optional arg5 As Variant, _
                Optional arg6 As Variant, _
                Optional arg7 As Variant, _
                Optional arg8 As Variant, _
                Optional arg9 As Variant)
    Dim w_rng As Range
    Dim w_Cell As New eCell
    
    For Each w_rng In m_Range
        Set w_Cell.ToSet = w_rng
        Application.Run ActiveWorkbook.Name & "!" & SubRoutineName, w_Cell, _
                        arg1, arg2, arg3, arg4, _
                        arg5, arg6, arg7, arg8, arg9
    Next
    
    Set w_Cell = Nothing
    Set w_rng = Nothing
    
End Sub

Public Sub CallSubroutineByCell2(SubRoutineNameWithWorkbookName As String, _
                Optional arg1 As Variant, _
                Optional arg2 As Variant, _
                Optional arg3 As Variant, _
                Optional arg4 As Variant, _
                Optional arg5 As Variant, _
                Optional arg6 As Variant, _
                Optional arg7 As Variant, _
                Optional arg8 As Variant, _
                Optional arg9 As Variant)
    Dim w_rng As Range
    Dim w_Cell As New eCell
    

    For Each w_rng In m_Range
        Set w_Cell.ToSet = w_rng
        Application.Run SubRoutineNameWithWorkbookName, w_Cell, _
                        arg1, arg2, arg3, arg4, _
                        arg5, arg6, arg7, arg8, arg9
    Next
    
    Set w_Cell = Nothing
    Set w_rng = Nothing
    
End Sub


Public Function OffsetSelect(Optional OffsetY As Long = 0, Optional OffsetX As Long = 0) As eCell
    'm_Range.Offset(オフセットy, オフセットx).Select
    
    Set m_Range = Range(m_Range.Cells, m_Range.Cells.Offset(OffsetY, OffsetX))
    m_Range.Select
    
    Set OffsetSelect = Me
End Function


Public Function ToEEzCell() As eEzCell
    Dim w_eEzCell As New eEzCell
    Set w_eEzCell.ToSet = m_Range
    Set ToEEzCell = w_eEzCell
    Set w_eEzCell = Nothing
End Function


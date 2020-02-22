VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "道具"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_セル As セル
Private m_シート As シート
Private m_ブック As ブック


'Public g_道具 As New 道具
Public Function Inst()
    Inst = g_道具
End Function




Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_セル = New セル
    Set m_シート = New シート
    Set m_ブック = New ブック
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_ブック = Nothing
    Set m_シート = Nothing
    Set m_セル = Nothing
End Sub

'↓よくない
'Public Property Get 相対位置(ByRef セル As 道具, Optional 行 As Integer = 0, Optional 列 As Integer = 0) As Range
'    Set 相対位置 = セル.Offset(行, 列)
'End Property

'Public Property Get セル01(Optional セル名 As String = "A1", Optional シート順番 As Integer = 1) As Range
'    Set セル01 = Worksheets(シート順番).Range(セル名).Offset(0, 0)
'End Property

'Public Property Get セル02(Optional セル名 As String = "A1", Optional シート順番 As Integer = 1) As セル
'    Set m_セル.セット = Worksheets(シート順番).Range(セル名).Offset(0, 0)
'    m_セル.セル範囲 = セル名
'    m_セル.シート順番 = シート順番
'    Set セル02 = m_セル
'End Property

'Public Property Get セル03(Optional セル名 As String = "A1", Optional シート順番 As Integer = 1) As セル
'    Set m_セル.セット = Worksheets(シート順番).Range(セル名).Offset(0, 0)
''    m_セル.セル範囲 = セル名
''    m_セル.シート順番 = シート順番
'    Set セル03 = m_セル
'End Property

'Public Property Get セル04(Optional セル名 As String = "A1", Optional シート順番 As Integer = 0) As セル
'    If シート順番 <> 0 Then
'        Set m_セル.セット = Worksheets(シート順番).Range(セル名).Offset(0, 0)
'    Else
'        Set m_セル.セット = ActiveSheet.Range(セル名).Offset(0, 0)
'    End If
''    m_セル.セル範囲 = セル名
''    m_セル.シート順番 = シート順番
'    Set セル04 = m_セル
'End Property

'---- セル04ベース ---------------------------
Public Function セル(Optional ByVal セル名 As String = "A1", Optional ByVal シート順番 As Long = 1) As セル
    Dim wセル As New セル
    If シート順番 > 0 Then
        Set wセル.セット = Worksheets(シート順番).Range(セル名)
    Else
        Set wセル.セット = ActiveSheet.Range(セル名)
    End If
    Set セル = wセル
    Set wセル = Nothing
End Function
Public Function 簡易セル(Optional ByVal セル名 As String = "A1", Optional ByVal シート順番 As Long = 1) As 簡易セル
    Dim wセル As New 簡易セル
    If シート順番 > 0 Then
        Set wセル.セット = Worksheets(シート順番).Range(セル名)
    Else
        Set wセル.セット = ActiveSheet.Range(セル名)
    End If
    Set セル = wセル
    Set wセル = Nothing
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

Public Property Get 現在のセル() As セル
    Set m_セル.セット = ActiveCell
    'm_セル.セル範囲 =
    m_セル.シート順番 = 0
    Set 現在のセル = m_セル
End Property

Public Property Get 現在の表() As セル
    Set m_セル.セット = ActiveCell.CurrentRegion
    'm_セル.セル範囲 =
    m_セル.シート順番 = 0
    Set 現在の表 = m_セル
End Property


Public Function シート(Optional シート順番 As Integer = 0) As シート
    If シート順番 <> 0 Then
        Set m_シート.セット = Worksheets(シート順番)
    Else
        Set m_シート.セット = ActiveSheet
    End If
    m_シート.シート順番 = シート順番
    Set シート = m_シート
End Function

Public Property Get 現在のシート() As シート
    Set m_シート.セット = ActiveSheet
    Set 現在のシート = m_シート
End Property


Public Property Get 名前配列() As 名前配列
    Set 名前配列 = m_ブック.名前配列
End Property

Public Function セル行列指定(Optional ByVal 行 As Long = 0, Optional ByVal 列 As Long = 0) As セル
    Set m_セル.セット = ActiveSheet.Cells(行, 列).Offset(0, 0)
    Set セル行列指定 = m_セル
End Function


Public Function セルが空白(行 As Long, 列 As Long, Optional シート番号 As Long = -1) As Boolean
    If シート番号 > 0 Then
        セルが空白 = (Trim(Worksheets(シート番号).Cells(行, 列).Value2) = "")
    Else
        セルが空白 = (Trim(ActiveSheet.Cells(行, 列).Value2) = "")
    End If
End Function

Public Function セルが空白でない(行 As Long, 列 As Long, Optional シート番号 As Long = -1) As Boolean
    If シート番号 > 0 Then
        セルが空白でない = Trim(Worksheets(シート番号).Cells(行, 列).Value2) <> ""
    Else
        セルが空白でない = Trim(ActiveSheet.Cells(行, 列).Value2) <> ""
    End If
End Function

Public Function セルが部分一致(部分一致文字列 As 文字列, 行 As Long, 列 As Long, Optional シート番号 As Long = -1) As Boolean
    If シート番号 > 0 Then
        セルが部分一致 = (InStr(1, Trim(Worksheets(シート番号).Cells(行, 列).Value2), 部分一致文字列.VBAに渡す) > 0)
    Else
        セルが部分一致 = (InStr(1, Trim(ActiveSheet.Cells(行, 列).Value2), 部分一致文字列.VBAに渡す) > 0)
    End If
End Function

Public Function セルが部分一致でない(部分一致文字列 As 文字列, 行 As Long, 列 As Long, Optional シート番号 As Long = -1) As Boolean
    If シート番号 > 0 Then
        セルが部分一致でない = (InStr(1, Trim(Worksheets(シート番号).Cells(行, 列).Value2), 部分一致文字列.VBAに渡す) <= 0)
    Else
        セルが部分一致でない = (InStr(1, Trim(ActiveSheet.Cells(行, 列).Value2), 部分一致文字列.VBAに渡す) <= 0)
    End If
End Function

Public Function セルが部分一致String(部分一致文字列 As String, 行 As Long, 列 As Long, Optional シート番号 As Long = -1) As Boolean
    If シート番号 > 0 Then
        セルが部分一致String = (InStr(1, Trim(Worksheets(シート番号).Cells(行, 列).Value2), 部分一致文字列) > 0)
    Else
        セルが部分一致String = (InStr(1, Trim(ActiveSheet.Cells(行, 列).Value2), 部分一致文字列) > 0)
    End If
End Function

Public Function セルが部分一致でないString(部分一致文字列 As String, 行 As Long, 列 As Long, Optional シート番号 As Long = -1) As Boolean
    If シート番号 > 0 Then
        セルが部分一致でないString = (InStr(1, Trim(Worksheets(シート番号).Cells(行, 列).Value2), 部分一致文字列) <= 0)
    Else
        セルが部分一致でないString = (InStr(1, Trim(ActiveSheet.Cells(行, 列).Value2), 部分一致文字列) <= 0)
    End If
End Function

Public Sub 表示(メッセージ As Variant)
    MsgBox メッセージ
    
End Sub
Function ワークブック名(Optional 連結文字つける As Boolean = True) As String
    If 連結文字つける Then
        ワークブック名 = ActiveWorkbook.Name & "!"
    Else
        ワークブック名 = ActiveWorkbook.Name
    End If
        
End Function
Public Function hngハンガリアンセル(Optional ByVal セル名 As String = "A1", Optional ByVal シート順番 As Long = 1) As hngハンガリアンセルv1

    Dim wセル As New hngハンガリアンセルv1
    If シート順番 > 0 Then
        Set wセル.setセット = Worksheets(シート順番).Range(セル名)
    Else
        Set wセル.setセット = ActiveSheet.Range(セル名)
    End If
    Set hngハンガリアンセル = wセル
    Set wセル = Nothing
End Function
 

Public Function hngハンガリアンセルv1(Optional ByVal セル名 As String = "A1", Optional ByVal シート順番 As Long = 1) As hngハンガリアンセルv1
    Dim wセル As New hngハンガリアンセルv1
    If シート順番 > 0 Then
        Set wセル.setセット = Worksheets(シート順番).Range(セル名)
    Else
        Set wセル.setセット = ActiveSheet.Range(セル名)
    End If
    Set hngハンガリアンセルv1 = wセル
    Set wセル = Nothing
End Function
 






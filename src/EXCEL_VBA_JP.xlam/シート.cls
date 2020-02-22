VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "シート"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Sheet As Worksheet
Private m_シート順番 As Long


Private Sub Class_Initialize()
    'ここに処理を書く
    Set m_Sheet = ActiveSheet
End Sub

Private Sub Class_Terminate()
    'ここに処理を書く
    Set m_Sheet = Nothing
End Sub

Public Property Set セット(sht As Worksheet)
    Set m_Sheet = sht
End Property

Public Property Let シート順番(lngシート順番 As Long)
    m_シート順番 = lngシート順番
End Property

Public Property Get シート順番() As Long
    シート順番 = m_シート順番
End Property

Public Property Get Name() As Worksheet
    Set Name = m_Sheet
End Property

Public Property Set Name(シート As Worksheet)
    Set m_Sheet = シート
End Property

'---- Rows
Public Property Get 行(ByVal 行番号 As Long) As セル
    Dim wセル As New セル
    Set wセル.セット = m_Sheet.Rows(行番号)
    Set 行 = wセル
End Property

Public Sub 行非表示(ByVal 行番号 As Long, ByVal 非表示 As Boolean)
    m_Sheet.Rows(行番号).Hidden = 非表示
End Sub

Public Sub 行高さ設定(ByVal 行番号 As Long, ByVal 行高さ As Variant)
    m_Sheet.Rows(行番号).RowHeight = 行高さ
End Sub

Public Sub 行高さ自動設定(ByVal 行番号 As Long)
    m_Sheet.Rows(行番号).AutoFit
End Sub

Public Function 行高さ取得(ByVal 行番号 As Long) As Variant
    行高さ取得 = m_Sheet.Rows(行番号).RowHeight
End Function

'---- Cols
Public Property Get 列(ByVal 列番号 As Long) As セル
    Dim wセル As New セル
    Set wセル.セット = m_Sheet.Columns(列番号)
    Set 列 = wセル
    Set wセル = Nothing
End Property

Public Sub 列非表示(ByVal 列番号 As Long, ByVal 非表示 As Boolean)
    m_Sheet.Columns(列番号).Hidden = 非表示
End Sub

Public Sub 列幅設定(ByVal 列番号 As Long, ByVal 列幅 As Variant)
    m_Sheet.Columns(列番号).ColumnWidth = 列幅
End Sub

Public Function 列幅取得(ByVal 列番号 As Long) As Variant
    列幅取得 = m_Sheet.Columns(列番号).ColumnWidth
End Function

Public Sub 列幅自動設定(ByVal 列番号 As Long)
    m_Sheet.Columns(列番号).AutoFit
End Sub

Public Sub 選択()
    m_Sheet.Activate
End Sub





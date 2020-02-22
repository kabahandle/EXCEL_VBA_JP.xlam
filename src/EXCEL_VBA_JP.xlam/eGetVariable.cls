VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eGetVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Property Get eString(Optional InitialValue As String) As eString
    Dim w_eString As New eString
    w_eString.Name = InitialValue
    Set eString = w_eString
    Set w_eString = Nothing
End Property

Public Property Get eInteger(Optional InitialValue As Long = 0, Optional Prefix As String = "", Optional Postfix As String = "") As eInteger
    Dim w_eInteger As New eInteger
    
    w_eInteger.ToSet InitialValue
    w_eInteger.SetPrefix Prefix
    w_eInteger.SetPostfix Postfix
    
    Set eInteger = w_eInteger
    Set w_eInteger = Nothing
End Property

Public Property Get eCurrency(Optional InitialValue As Currency = 0) As eCurrency
    Dim w_eCurrency As New eCurrency
    
    w_eCurrency.ToSet InitialValue
    
    Set eCurrency = w_eCurrency
    Set w_eCurrency = Nothing
End Property

Public Property Get eArray(Optional InitialElementsNumber As Long = 0) As eArray
    Dim w_eArray As New eArray
    
    Call w_eArray.SetElementsNum(InitialElementsNumber)
    
    Set eArray = w_eArray
    Set w_eArray = Nothing
    
End Property

Public Property Get eStack() As eStack
    Dim w_eStack As New eStack
    
    Set eStack = w_eStack
    Set w_eStack = Nothing
    
End Property





Attribute VB_Name = "open_close"
Private Sub Auto_Open()
    Set VBAJP_FSO = CreateObject("Scripting.FileSystemObject")
    ReDim vbajp_ary_Files(1)
    Set WSH = CreateObject("WScript.Shell")
    vbajp_MyDoc = WSH.SpecialFolders("MyDocuments")
    Set r_RegExp = CreateObject("VBScript.RegExp")
End Sub

Private Sub Auto_close()
    Set g_���� = Nothing
    Set eg_eTool = Nothing
    Set k_�ȈՓ��� = Nothing
    Set ek_eEzTool = Nothing
    Set h_�ϐ��擾 = Nothing
    Set eh_eGetValuable = Nothing
    Set io_���o�� = Nothing
    Set eio_eIO = Nothing
    Set VBAJP_FSO = Nothing
    Set vbajp_MyDoc = Nothing
    Set WSH = Nothing
    Set r_RegExp = Nothing
End Sub


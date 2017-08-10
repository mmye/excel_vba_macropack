Attribute VB_Name = "PDF�v�����^�["
Option Explicit
' # �A�N�e�B�u�V�[�g��PDF�Ńf�X�N�g�b�v�ɕۑ�����

Sub PDF�v�����^�[()
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim fileName As String
    On Error Resume Next
    fileName = FSO.GetBaseName(ActiveWorkbook.Name)
        
    Dim Path As String, WSH As Variant
    Set WSH = CreateObject("WScript.Shell")
    Path = WSH.SpecialFolders("Desktop") & "\"
    
    Dim stName
    stName = ActiveSheet.Name
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        fileName:=Path & stName & ".pdf", _
        quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

    Set FSO = Nothing
    Set WSH = Nothing
    
End Sub





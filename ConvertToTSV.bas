Attribute VB_Name = "ConvertToTSV"
Option Explicit

Sub ConvertToTSV()
'   ### �J�����u�b�N��txt�ɕϊ�����X�N���v�g
'   ### �ۑ��O�ɑS�u�b�N�łP��ڂ��폜���鏈���������Ă���

    Dim wkb As Workbook

    For Each wkb In Workbooks
        Dim wks As Worksheet
        
        Set wks = wkb.Sheets(1)
        wks.Columns(2).EntireColumn.Delete
        wkb.SaveAs fileName:=wkb.Path & "\txt\en_src\" & wkb.Name & "_en.txt", _
           FileFormat:=xlCurrentPlatformText
    Next wkb

End Sub

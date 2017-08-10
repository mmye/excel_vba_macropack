Attribute VB_Name = "ConvertToTSV"
Option Explicit

Sub ConvertToTSV()
'   ### 開いたブックをtxtに変換するスクリプト
'   ### 保存前に全ブックで１列目を削除する処理が入っている

    Dim wkb As Workbook

    For Each wkb In Workbooks
        Dim wks As Worksheet
        
        Set wks = wkb.Sheets(1)
        wks.Columns(2).EntireColumn.Delete
        wkb.SaveAs fileName:=wkb.Path & "\txt\en_src\" & wkb.Name & "_en.txt", _
           FileFormat:=xlCurrentPlatformText
    Next wkb

End Sub

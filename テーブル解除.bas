Attribute VB_Name = "�e�[�u������"
Option Explicit

Sub �e�[�u������()
    Dim Tbl As ListObject
    Dim wks As Worksheet
    Const myWks As String = "DFVN-F5 man. IPK (2)"
    
    For Each wks In ThisWorkbook.Sheets
        If wks.Name <> myWks Then
            For Each Tbl In ActiveSheet.ListObjects
                Tbl.Unlist
            Next Tbl
        End If
    Next wks
    
End Sub

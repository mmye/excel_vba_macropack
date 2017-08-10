Attribute VB_Name = "CopyColaToColb"
Option Explicit

Sub CopyColaToColb()
    Dim i As Long
    Dim wkb As Workbook
    Dim wks As Worksheet
    
    Performance.OptimizeOn
    
    For Each wkb In Workbooks
        Set wks = wkb.Sheets(1)
        
        Dim r As Range
        Set r = wks.UsedRange
        Dim LastRow As Long
        LastRow = r.Rows.Count
        
        For i = 5 To LastRow
            wks.Cells(i, 2).Value = wks.Cells(i, 1).Value
        Next i
    Next wkb
    
    Set wkb = Nothing
    Set wks = Nothing
    
    BulkBookUtil.SaveAllBooks
    Performance.OptimizeOff
    BulkBookUtil.CloseAllBooks
End Sub

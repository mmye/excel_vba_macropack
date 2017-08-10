Attribute VB_Name = "ページセットアップツール"
Option Explicit

Sub 改ページプレビューに切替()
Attribute 改ページプレビューに切替.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 改ページプレビューに切替 Macro
'
Dim wks As Worksheet

For Each wks In ActiveWorkbook.Sheets
    wks.Activate
'
    ActiveWindow.View = xlPageBreakPreview
Next wks
End Sub
Sub 通常モードに切替()
Attribute 通常モードに切替.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 通常モードに切替 Macro
Dim wks As Worksheet

For Each wks In ActiveWorkbook.Sheets
    wks.Activate
'
    ActiveWindow.View = xlNormalView
Next wks
End Sub
Sub フッターにページ番号とシート名を記載()
    
    Dim wks As Worksheet

    Application.PrintCommunication = False
    
    For Each wks In ActiveWorkbook.Sheets
        With wks
            .PageSetup.RightHeader = "&A"
            .PageSetup.LeftHeader = "&F | &D"
            .PageSetup.LeftFooter = "&P/&N"
        End With
    Next wks

    Application.PrintCommunication = True
End Sub
Sub 拡大率を100にする()
Attribute 拡大率を100にする.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 拡大率を100にする Macro
'
Dim wks As Worksheet

For Each wks In ActiveWorkbook.Sheets
    wks.Activate
'
    ActiveWindow.Zoom = 100
Next wks

End Sub

Sub 行高を増やす()
    Dim i As Long
    Dim lEndRow As Long
    lEndRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lEndRow
    DoEvents
        Rows(i).EntireRow.RowHeight = Rows(i).EntireRow.RowHeight * 1.5
    Next i

End Sub

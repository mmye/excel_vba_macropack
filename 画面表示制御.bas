Attribute VB_Name = "画面表示制御"
Option Explicit

Sub 枠線表示切り替え()
Attribute 枠線表示切り替え.VB_Description = "シートの枠線表示/非表示を切り替える"
Attribute 枠線表示切り替え.VB_ProcData.VB_Invoke_Func = "G\n14"

    ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines
End Sub
Sub ブック枠線表示切り替え()

    Dim myWindow As Window
    
    For Each myWindow In ActiveWorkbook.Windows
        myWindow.DisplayGridlines = Not myWindow.DisplayGridlines
    Next
    
End Sub


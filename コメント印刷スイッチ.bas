Attribute VB_Name = "コメント印刷スイッチ"
Option Explicit

Sub コメント印刷スイッチ()
    Dim cmt
    Dim st As Worksheet
    
    Set st = ActiveSheet
    
    For Each cmt In st.Comments
        cmt.Visible = Not cmt.Visible
    Next cmt
    
    If st.PageSetup.PrintComments = xlPrintNoComments Then
        st.PageSetup.PrintComments = xlPrintInPlace
        Set st = Nothing
        Exit Sub
    End If
    
    If st.PageSetup.PrintComments = xlPrintInPlace Then
        st.PageSetup.PrintComments = xlPrintNoComments
        Set st = Nothing
        Exit Sub
    End If
    
End Sub

Sub コメント表示切り替え()
    Dim cmt
    Dim st As Worksheet
    
    Set st = ActiveSheet
    
    For Each cmt In st.Comments
        cmt.Visible = Not cmt.Visible
    Next cmt
    Set st = Nothing
End Sub
    


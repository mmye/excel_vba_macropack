Attribute VB_Name = "�y�[�W�Z�b�g�A�b�v�c�[��"
Option Explicit

Sub ���y�[�W�v���r���[�ɐؑ�()
Attribute ���y�[�W�v���r���[�ɐؑ�.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ���y�[�W�v���r���[�ɐؑ� Macro
'
Dim wks As Worksheet

For Each wks In ActiveWorkbook.Sheets
    wks.Activate
'
    ActiveWindow.View = xlPageBreakPreview
Next wks
End Sub
Sub �ʏ탂�[�h�ɐؑ�()
Attribute �ʏ탂�[�h�ɐؑ�.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �ʏ탂�[�h�ɐؑ� Macro
Dim wks As Worksheet

For Each wks In ActiveWorkbook.Sheets
    wks.Activate
'
    ActiveWindow.View = xlNormalView
Next wks
End Sub
Sub �t�b�^�[�Ƀy�[�W�ԍ��ƃV�[�g�����L��()
    
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
Sub �g�嗦��100�ɂ���()
Attribute �g�嗦��100�ɂ���.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �g�嗦��100�ɂ��� Macro
'
Dim wks As Worksheet

For Each wks In ActiveWorkbook.Sheets
    wks.Activate
'
    ActiveWindow.Zoom = 100
Next wks

End Sub

Sub �s���𑝂₷()
    Dim i As Long
    Dim lEndRow As Long
    lEndRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lEndRow
    DoEvents
        Rows(i).EntireRow.RowHeight = Rows(i).EntireRow.RowHeight * 1.5
    Next i

End Sub

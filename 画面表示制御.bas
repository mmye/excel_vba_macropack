Attribute VB_Name = "��ʕ\������"
Option Explicit

Sub �g���\���؂�ւ�()
Attribute �g���\���؂�ւ�.VB_Description = "�V�[�g�̘g���\��/��\����؂�ւ���"
Attribute �g���\���؂�ւ�.VB_ProcData.VB_Invoke_Func = "G\n14"

    ActiveWindow.DisplayGridlines = Not ActiveWindow.DisplayGridlines
End Sub
Sub �u�b�N�g���\���؂�ւ�()

    Dim myWindow As Window
    
    For Each myWindow In ActiveWorkbook.Windows
        myWindow.DisplayGridlines = Not myWindow.DisplayGridlines
    Next
    
End Sub


Attribute VB_Name = "���݂̃t�@�C�������擾"
Option Explicit

Sub GetActiveWorkbookName()
    Dim buf2 As String
    Dim CB As New DataObject
    Dim buf
    
    buf = ActiveWorkbook.Name
    
    With CB
        .SetText buf        ''�ϐ��̃f�[�^��DataObject�Ɋi�[����
        .PutInClipboard     ''DataObject�̃f�[�^���N���b�v�{�[�h�Ɋi�[����
        .GetFromClipboard   ''�N���b�v�{�[�h����DataObject�Ƀf�[�^���擾����
        buf2 = .GetteXt     ''DataObject�̃f�[�^��ϐ��Ɏ擾����
    End With
    MsgBox "�t�@�C�������N���b�v�{�[�h�ɃR�s�[���܂����B" & vbCrLf & buf
    
    Set CB = Nothing
End Sub

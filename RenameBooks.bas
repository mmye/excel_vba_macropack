Attribute VB_Name = "RenameBooks"
Option Explicit

Sub RenameBooks()

'### �w��t�H���_�ɂ���u�b�N�����g�p���ĕʃt�H���_�̃t�@�C������ύX����
' ## �t�@�C���擾�̏��Ԃ����������Ȃ�@��̘A�Ԃ��������Ȃ�Ȃ����@1,10,11,�݂�����

Dim src_path As String
src_path = "C:\000_Kaketsuken_Japanese_HMI\source\"

'�e�L�X�g�����������u�b�N�̕ۑ���
Dim tgt_path As String
tgt_path = "C:\000_Kaketsuken_Japanese_HMI\combined\"

Dim new_name As Variant
new_name = GetFileNames.GetFileNames(src_path)

Dim old_name As Variant
old_name = GetFileNames.GetFileNames(tgt_path)

Dim v

For Each v In new_name
    Name tgt_path & old_name As tgt_path & new_name
Next v


End Sub

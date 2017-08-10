VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM�^�[�~�i�� 
   Caption         =   "���Ϗ����J��..."
   ClientHeight    =   1335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3855
   OleObjectBlob   =   "FRM�^�[�~�i��.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FRM�^�[�~�i��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mLists As Collection
Dim IndexFile As String
Const UboundArr As Long = 6
Const IndexPath As String = "\\LS410D760\share\������\�s1�t���ρE����\1. ���Ϗ�" '�����̃Z�p���[�^�[�͂Ȃ��Ŏw�肷��
Const strInstruction As String = "���Ϗ��ԍ�����͂���Enter�������Ă�������"

Enum RevisionStatus
    FirstRevision = 0
    HasRevisedBefore = 1
End Enum

Private Sub txtTerminal_Change()
    Dim buf As String
    buf = txtTerminal.Text
    txtTerminal.Text = UCase(buf)
    If txtStatus.Caption = "" Then txtStatus.Caption = strInstruction
    If txtStatus.Caption <> "" Then
        txtStatus.Caption = Empty
        txtStatus.Caption = strInstruction
    End If
End Sub

Private Sub UserForm_Click()
3355    txtTerminal.SetFocus
End Sub

Sub LoadWks()
    Set mLists = LoadIndexLists
    If mLists Is Nothing Then
        MsgBox "�C���f�b�N�X����ł�"
        Exit Sub
    End If
End Sub


Private Sub txtTerminal_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim Key As String
    Dim v As Variant
    Dim Path As String, ShName As String

    Debug.Print KeyCode
    Select Case KeyCode
        Case 27
            Unload Me
            Exit Sub
        Case 13
            If LenB(txtTerminal.Text) = 0 Then Exit Sub
            Key = StrConv(txtTerminal.Text, vbNarrow)
            
            Select Case Right$(Key, 2)
                Case "-R"
                    Key = Left$(Key, Len(Key) - 3)
                    Dim IsRevision As Boolean
                    IsRevision = True
                Case "-W"
                    Dim Writable As Boolean
                    Writable = True
                    Key = Left$(Key, Len(Key) - 3)
                End Select
                Key = Trim$(Key)
                Key = UCase(Key)
    End Select
    
    If Len(Key) = 0 Then Exit Sub
    On Error Resume Next
    v = Split(mLists(Key), vbTab)
    If IsEmpty(v) Then
        txtStatus.Caption = Empty
        txtStatus.Caption = "���Ϗ���������܂���B"
        Exit Sub
    End If
    Path = v(6) & v(5)
    ShName = v(7)
    
    Dim File As String, bk As Workbook
    File = Dir(Path)
    Select Case File
        Case ""
            '���Ϗ��u�b�N��������Ȃ������ꍇ
        Case Else
            On Error Resume Next
            Select Case Writable
                Case True
                    Set bk = Workbooks.Open(fileName:=Path, ReadOnly:=False)
                Case False
                    Set bk = Workbooks.Open(fileName:=Path, ReadOnly:=True)
            End Select
            
            If IsRevision Then
                If ShName <> "" Then
                    Dim st As Worksheet
                    Application.DisplayAlerts = False
                    Set st = bk.Sheets(ShName).Copy(After:=bk.Sheets(ShName))
                    Application.DisplayAlerts = False
                    
                    '���r�W�����̃V�[�g�������邽�߂����Ⴒ�������Ă�
                    Dim stName As String
                    stName = bk.Sheets(ShName).Name
                    If (Right$(stName, 2) Like "R[0-9]+") Then
                        stName = Replace(st.Name, "R", "")
                        st.Name = stName & "R" & (Right$(stName, 1) + 1)
                    Else
                        '���łɂ�������Ԃ����m����K�v������
                        'st.Name = st.Name & "R1"
                    End If
                    st.Activate
                    Range("a1").Select
                End If
            Else
                If ShName <> "" Then
                    bk.Sheets(ShName).Activate
                    Range("a1").Select
                End If
                txtStatus.Caption = Empty
                txtStatus.Caption = "���Ϗ�" & ShName & "���J���܂����B"
                Unload Me
            End If
    End Select
    On Error GoTo 0
    Set bk = Nothing
End Sub

Private Sub UserForm_Initialize()
    LoadWks
    txtStatus.Caption = strInstruction
End Sub

Private Function LoadIndexLists() As Collection
    Dim v As Variant
    Dim a As String, l As String, z As String
    Dim Lists As New Collection
    Dim List As Variant

    IndexFile = IndexPath & "\" & "index.txt"
    If Dir(IndexFile) = "" Then MsgBox "�C���f�b�N�X�t�@�C����������܂���B", vbCritical

    On Error Resume Next
    Open IndexFile For Input As #1
    Do Until EOF(1)
        ERR.Number = 0
        Line Input #1, a
        List = Split(a, vbTab)
        Dim Key As String, Content As String
        Key = List(0)
        Content = GetContent(List) 'Content ���w���BVariant�z���Ԃ��ׂ����H�H
        On Error Resume Next
        Lists.Add Content, Key
    Loop
    On Error GoTo 0
    Close #1
    Set LoadIndexLists = Lists
Exit Function
ERR:
Close #1

End Function
Private Function GetContent(List As Variant) As String
    Dim v As Variant
    Dim buf As String
    For Each v In List
        buf = buf & vbTab & v
    Next v
    GetContent = buf
End Function

Private Sub UserForm_Terminate()
    Set mLists = Nothing
End Sub


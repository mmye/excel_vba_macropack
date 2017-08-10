VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "�u��"
   ClientHeight    =   1770
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7680
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Counter As Long
Dim bookNum As String

Private Sub CommandButton1_Click() '�u�����s�{�^��

Application.ScreenUpdating = False
Call ActivateTargetBook
Call ResetCounter
Call CheckBeforeReplacement
Call NormalReplacement
Call ShowReplacementCounts
Application.ScreenUpdating = True
End Sub

Private Sub CommandButton4_Click() '�u����v�{�^��
Unload Me
End Sub

Private Sub CommandButton2_Click() '�u�V�K�v�{�^��

Call CreateListBook '�V�K���X�g�u�b�N�쐬

ComboBox1.Clear
Call GetListBook '�ēǂݍ���
Call WriteListBook
End Sub
Sub WriteListBook()
ComboBox1.Text = GetNumListbook
End Sub
Private Sub CommandButton3_Click() '���X�g�u�b�N�I���{�^��

Dim bk As Object

    '���X�g�u�b�N���I������Ă��Ȃ��ꍇ
    If ComboBox1.Text = "" Then
        Dim wb As Workbook
        Dim bName As String
        bName = GetNumListbook
        ComboBox1.AddItem (bName)
        Workbooks(bName).Activate
    Else
    '���X�g�u�b�N���I������Ă���ꍇ
        
        '���X�g�u�b�N�����݂��邩�`�F�b�N
        For Each wb In Workbooks
            If wb.Name = ComboBox1.Text Then
                Dim flag As Boolean
                flag = True
            End If
        Next
            '���݂��Ȃ��ꍇ
            If flag = False Then
                ComboBox1.Clear
                Call GetListBook
                MsgBox "���̃��X�g�u�b�N�͊J����Ă��Ȃ��悤�ł�"
            Else
            '���݂���ꍇ
                Workbooks(ComboBox1.Text).Activate
                
            End If
    End If
        
End Sub
Private Sub CommandButton5_Click()

   '�u������u�b�N���I������Ă��Ȃ��ꍇ
    If ComboBox2.Text = "" Then
        Exit Sub
    Else
    
    '�I������Ă���ꍇ
       
    '�u�b�N�����݂��邩�`�F�b�N
        Dim wb As Workbook
        For Each wb In Workbooks
            If wb.Name = ComboBox2.Text Then
                Dim flag As Boolean
                flag = True
            End If
        Next
            '���ʁA���݂��Ȃ��ꍇ
            If flag = False Then
                ComboBox2.Clear
                Call GetOpenbook
                MsgBox "���̃u�b�N�͊J����Ă��Ȃ��悤�ł�"
            Else
            '���݂���ꍇ
                VBA.AppActivate Excel.Application.Caption
                'Windows(ComboBox2.Text).Activate
            End If
        End If
    End Sub

Private Sub UserForm_Initialize()

ComboBox1.SetFocus '�R�}���h�{�^���Ƀt�H�[�J�X

'TIPS
ComboBox1.ControlTipText = "�u���p���̃u�b�N"
ComboBox2.ControlTipText = "�u���������u�b�N"
CommandButton1.ControlTipText = "�u�������s"
OptionButton1.ControlTipText = "�������Ԃ�"
OptionButton2.ControlTipText = "��������"
OptionButton3.ControlTipText = "�������΂�"
OptionButton4.ControlTipText = "������̐F���̂܂�"

StartUpPosition = 1 '�t�H�[�����G�N�Z���̒����ɕ\��

'�^�u�I�[�_�[
ComboBox1.TabIndex = 0
ComboBox2.TabIndex = 1
CommandButton2.TabIndex = 2
OptionButton1.TabIndex = 3
OptionButton2.TabIndex = 4
OptionButton3.TabIndex = 5
CommandButton1.TabIndex = 6

'�R���{�{�b�N�X�X�^�C��
ComboBox1.Style = fmStyleDropDownList
ComboBox2.Style = fmStyleDropDownList

'�I�v�V�����{�b�N�X�����I��
OptionButton1 = True

Call ClearCombobox '�R���{�{�b�N�X���N���A
Call GetOpenbook '�J���Ă���u�b�N�����ׂĎ��
Call GetListBook '���X�g�u�b�N���擾

End Sub
Sub ShowReplacementCounts()

MsgBox Counter & "���u�����܂���"
End Sub
Sub ResetCounter()
Counter = 0 '���W���[���ϐ����N���A�B���ꂪ�Ȃ��ƒu���������ݐς���Ă��܂��B
End Sub
Sub ActivateTargetBook()

'�u������u�b�N���A�N�e�B�u�ɂ��Ȃ��ƒu������Ȃ�
If ComboBox2.Text <> "" Then
    Workbooks(ComboBox2.Text).Activate
End If
End Sub

Sub ClearCombobox()
    
    ComboBox1.Clear
    ComboBox2.Clear
End Sub
Sub GetOpenbook()
    
'�J���Ă���u�b�N�����ׂĎ擾�B�R���{�{�b�N�X�Q�ɓ����B
       
    Dim wb As Workbook
    For Each wb In Workbooks
        With wb.Sheets(1)
            If .Cells(1, 1) <> "�������镶����" And .Cells(1, 2) <> "�u����̕�����" Then
                ComboBox2.AddItem wb.Name
            End If
        End With
    Next wb
End Sub
Sub GetListBook()

'���X�g���擾�B�J�n�s�̕�����Ŕ���B�R���{�{�b�N�X�P�ɓ����B
    Dim i As Long, wb As Workbook
   
    For Each wb In Workbooks
        With wb.Sheets(1)
            If .Cells(1, 1) = "�������镶����" And .Cells(1, 2) = "�u����̕�����" Then
                ComboBox1.AddItem wb.Name
            End If
        End With
    Next wb
End Sub
Function GetNumListbook()

'�J���Ă���u�b�N�𐔂��A�ŏI�u�b�N����Ԃ�
Dim bkNum As Integer
bkNum = Workbooks.Count
Dim bkName As String
GetNumListbook = Workbooks(bkNum).Name

End Function
Sub CheckBeforeReplacement()

Dim flag As Boolean
flag = cDocFalse '�֐��Ō�����u�b�N�I�������邩����
    If flag Then
        Exit Sub '���肠��ŒE�o
    End If

Dim TargetSt As Worksheet
Set TargetSt = Workbooks(ComboBox2.Text).Sheets(1)
Dim n As Long
n = CountCells(TargetSt) '��ł͂Ȃ��Z���̌���Ԃ�

Dim flag2 As Boolean
flag2 = JudgeCellCounts(n)

    If flag2 = True Then
        Call PromptUsertoResume
    End If
End Sub
Sub PromptUsertoResume()

Dim flag As Integer
flag = MsgBox("�����Ɏ��Ԃ�������\��������܂��B" & vbCrLf _
            & "�����܂����H" _
        , vbYesNo + vbQuestion, "�������s�O�̒���")

    If flag = vbYes Then
        Exit Sub
    Else
        End
    End If
End Sub

Sub NormalReplacement()
'���X�g�ƃ^�[�Q�b�g�u�b�N��ϐ��Ɋi�[���A
'����������ƒu����̌���ϐ��Ɋi�[����
'�����āA�u�����s�v���V�[�W���ɓn��
'�߂��̓C�x���g�v���V�[�W���i�R�}���h�{�^���N���b�N�j

Dim listBook As Workbook
Dim TargetBook As Workbook
Dim strWhat As String
Dim strReplacement As String
Dim i As Integer

On Error GoTo ErrHandler

Set listBook = Workbooks(ComboBox1.Text)
Set TargetBook = Workbooks(ComboBox2.Text)

    With listBook.Sheets(1)
        
        For i = 2 To .Range("A10000").End(xlUp).Row
            strWhat = .Cells(i, 1).Value
            strReplacement = .Cells(i, 2).Value
            Call doReplace(strWhat, strReplacement, TargetBook) '�u�����s
        Next i
    End With
Exit Sub

ErrHandler:
    Dim myMsg As String
    myMsg = "�G���[�ԍ��F" & ERR.Number & vbCrLf & _
                "�G���[���e�F" & ERR.Description
    MsgBox myMsg
End Sub
Sub doReplace(ByRef strWhat As String, strReplacement As String, TargetBook As Workbook)

    With TargetBook.ActiveSheet
        
        Dim r As Range
        Dim start As Integer
        Dim ExplorerRange As Range
        Set ExplorerRange = .Range("A1", ActiveCell.SpecialCells(xlLastCell)) '�f�[�^������Z���͈͂��擾
    
        Dim FontColor As Integer
        FontColor = GetFontColor '�I�v�V�����{�^��������ׂ�
            
        For Each r In ExplorerRange
        
            start = 1
            Do While InStr(start, r.Value, strWhat) > 0
                start = InStr(start, r.Value, strWhat)
                r.Characters(start, Len(strWhat)).Delete
                r.Characters(start, 0).insert strReplacement
                r.Characters(start, Len(strReplacement)).Font.ColorIndex = FontColor
                Counter = Counter + 1
                start = start + Len(strReplacement)
                'Debug.Print Counter
            Loop
        Next
    End With
End Sub
Function GetFontColor()

    If OptionButton1 Then
        GetFontColor = 3 '��
    End If
    If OptionButton2 Then
        GetFontColor = 5 '��
    End If
    If OptionButton3 Then
        GetFontColor = 50 '��
    End If
    If OptionButton4 Then
        GetFontColor = 1 '��
    End If
End Function
Function cDocFalse()

Dim List As String
Dim bk As String

List = ComboBox1.Text
bk = ComboBox2.Text

    If List = bk Then '�P�j����̕������I������Ă���ꍇ
        MsgBox "�u�����X�g�ƃ^�[�Q�b�g�������ł��B"
        cDocFalse = True
        Exit Function
    End If
    
    If List = "" And bk = "" Then '�Q�j�����ꂩ�̕������I������Ă��Ȃ��ꍇ
        MsgBox "�u�����X�g�ƒu������u�b�N���I������Ă��܂���B"
        cDocFalse = True
        Exit Function
    End If
    
    If List <> "" And bk = "" Then '�Q�j�����ꂩ�̕������I������Ă��Ȃ��ꍇ
        MsgBox "�u������u�b�N���I������Ă��܂���B"
        cDocFalse = True
        Exit Function
    End If
    
    If List = "" And bk <> "" Then '�Q�j�����ꂩ�̕������I������Ă��Ȃ��ꍇ
        MsgBox "�u�����X�g���I������Ă��܂���B"
        cDocFalse = True
        Exit Function
    End If
    
    
End Function

Sub CreateListBook()

Dim wb As Workbook
Application.ScreenUpdating = False

Set wb = Workbooks.Add

    With wb.Sheets(1)
        .Cells(1, 1) = "�������镶����"
        .Cells(1, 2) = "�u����̕�����"
        .Cells(2, 1).Select
        .Name = "1"
    End With

Call ResizeBook(wb)
Call LockPane(wb)

bookNum = GetNumListbook

Application.ScreenUpdating = True

End Sub

Sub ResizeBook(ByRef wb As Workbook)

'�Ώۂ̃��[�N�u�b�N��ʏ�\���ɂ��܂�
    wb.Activate
    ActiveWindow.WindowState = xlNormal
'���{�����\��
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
'�X�N���[���͈͂̐���
    wb.Sheets(1).ScrollArea = "A1:B1000"
'���[�N�u�b�N�̃E�B���h�E�T�C�Y��ύX���܂�
    ActiveWindow.Height = 200
    ActiveWindow.Width = 300
    ActiveWindow.Top = 250
    ActiveWindow.Left = 250

    Call LayoutBook(wb)
End Sub
Sub LockPane(ByRef wb As Workbook)

'�E�C���h�E�g��擪�s�ŌŒ�
    wb.Activate
    ActiveWindow.FreezePanes = True
End Sub
Sub LayoutBook(ByRef wb As Workbook)

'�t�H���g
    With Sheets(1).Range("A1:B1")
        .Font.Name = "MS �S�V�b�N"
        .Font.Size = 10
        .Font.Bold = True
        .Font.ColorIndex = 2
'�h��Ԃ�
        .Interior.ColorIndex = 1
'���x����������
        .HorizontalAlignment = xlCenter
    End With
    
    With Sheets(1).Range("A2:B100")
        .Font.Name = "MS �S�V�b�N"
        .Font.Size = 10
    End With

'�r��
    With Sheets(1).Range("A1:B100")
        .Borders.LineStyle = True
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Weight = xlHairline
    End With

'��
    With Sheets(1).Columns("A:B")
        .ColumnWidth = 23
    End With
End Sub

Function CountCells(st As Worksheet) As Long

'�V�[�g���̃f�[�^���������Z������Ԃ�

With Workbooks(ComboBox2.Text).Sheets(1)

    Dim ActiveRng As Range
    Set ActiveRng = .Range("a1", ActiveCell.SpecialCells(xlLastCell))
    
    Dim rng As Range
    For Each rng In ActiveRng
        If rng.Value <> "" Then
            Dim n As Long
            n = n + 1
        End If
    Next
End With
    
CountCells = n

End Function

Function JudgeCellCounts(n As Long) As Boolean

Dim MaxLimit As Long
MaxLimit = 1500 '���ӏ��

If n > MaxLimit Then
    JudgeCellCounts = True
End If

End Function

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRM�Ζ�\�c�[�� 
   Caption         =   "�Ζ�\�c�[��"
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5010
   OleObjectBlob   =   "FRM�Ζ�\�c�[��.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "FRM�Ζ�\�c�[��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    
    If cmbOriginalWkb.Text = cmbTranslatedWkb Then
        MsgBox "�����Ɩ󕶂͕ʂ̃u�b�N��I�����Ă�������", vbOKOnly, "���m�点"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Call �Ζ�\�쐬
End Sub

Private Sub CommandButton2_Click()
 Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim wkb As Workbook
    
    With cmbOriginalWkb
        For Each wkb In Workbooks
            .AddItem wkb.Name
        Next wkb
    End With
    
    With cmbTranslatedWkb
        For Each wkb In Workbooks
            .AddItem wkb.Name
        Next wkb
    End With
End Sub

Private Sub �Ζ�\�쐬()
    Dim wkbOriginal As Workbook
    Dim wkbTranslation As Workbook
    Dim wkbNew As Workbook
    Dim wksCompareTBL As Worksheet
    Dim buf As String
    Dim r As Long
    Dim c As Long
    Dim lEndRow As Long
    Dim lEndCol As Long
    Dim wks As Worksheet
    Dim wksIndex As Long
    Dim lStart As Long
    Dim lEnd As Long
    Dim i As Long
    Dim sOriginalWkb As String
    Dim sTranslatedWkb As String
    Const mySeparator As String = ""
    Const GrayFontColor As Long = 5855577
    
    Application.ScreenUpdating = False
    sOriginalWkb = cmbOriginalWkb.Text
    sTranslatedWkb = cmbTranslatedWkb.Text

    Set wkbOriginal = Workbooks(sOriginalWkb)
    Set wkbTranslation = Workbooks(sTranslatedWkb)
    
    Set wkbNew = Workbooks.Add
    
    For Each wks In wkbOriginal.Sheets
        'If i > 1 Then Exit Sub
        wksIndex = wksIndex + 1
        Application.StatusBar = wksIndex & ":" & wks.Name & "�̑Ζ���쐬���Ă��܂�..."
        If wksIndex > 1 Then
            
            '�V�[�g�������ꍇ�͔�΂�
            On Error GoTo ERR
            wkbTranslation.Sheets(wksIndex).Copy After:=wkbNew.Sheets(wkbNew.Sheets.Count)
            Set wksCompareTBL = wkbNew.Sheets(wkbNew.Sheets.Count)
            
            On Error Resume Next '���Ɏg�p����Ă���V�[�g���������
            wksCompareTBL.Name = wkbTranslation.Sheets(wksIndex).Name
            If ERR.Number <> 0 Then
                wksCompareTBL.Name = "foo_" & wkbTranslation.Sheets(wksIndex).Name
            End If
            ERR.Clear
            On Error GoTo 0
        
        Else
            wkbTranslation.Sheets(wksIndex).Copy After:=wkbNew.Sheets(wkbNew.Sheets.Count)
            Set wksCompareTBL = wkbNew.Sheets(wkbNew.Sheets.Count)
            wksCompareTBL.Name = wkbTranslation.Sheets(wksIndex).Name
        End If
        
        lEndRow = wks.UsedRange.Item(wks.UsedRange.Count).Row
        lEndCol = wks.UsedRange.Item(wks.UsedRange.Count).column
        For r = 1 To lEndRow
            For c = 1 To lEndCol
                With wksCompareTBL
                DoEvents
                    With .Cells(r, c)
                        If .Value <> "" Then
                            .Font.Name = "Arial" '������̃t�H���g�����낦��
                            '�������ꔻ��FJA�u�b�N�̌����͒ǋL���Ȃ�
                            If .Value <> wkbOriginal.Sheets(wksIndex).Cells(r, c).Value Then
                                lStart = Len(.Value) + Len(mySeparator) + 1
                                buf = wkbOriginal.Sheets(wksIndex).Cells(r, c).Value
                                .Value = .Value & vbCrLf & mySeparator & buf
                                lEnd = Len(.Value)
                                With .Characters(lStart, lEnd).Font
                                    If wksCompareTBL.Cells(r, c).Interior.color <> vbWhite Then
                                        If wks.Cells(r, c).Interior.color = 13434828 Then
                                            .color = GrayFontColor
                                        Else
                                            .color = GrayFontColor ' vbWhite
                                        End If
                                    Else
                                        .color = GrayFontColor '�D�F
                                    End If
                                    .Name = "�l�r �o�S�V�b�N"
                                    .Size = 0.9 * wkbOriginal.Sheets(wksIndex).Cells(r, c).Font.Size
                                End With
                            End If
        '                    i = i + 1
                        End If
                    End With
                End With
            Next c
        Next r
    Next wks
    
    '�]���ȃV�[�g���폜���đI����A1�ɂ���
    With wkbNew.Sheets(1)
        Application.DisplayAlerts = False
        .Delete
        Application.DisplayAlerts = True
        wkbNew.Sheets(1).Activate
        Range("a1").Select
    End With
    
    Application.ScreenUpdating = True
    
    Set wkbNew = Nothing
    Set wkbOriginal = Nothing
    Set wkbTranslation = Nothing
    Application.StatusBar = False
    MsgBox "�������܂���", vbOKOnly + vbInformation, "���m�点"
    
    Exit Sub

ERR:
    'wksIndex���s���Ȓl�������ꍇ�̃G���[�����i���݂��Ȃ��V�[�g�C���f�b�N�X�̎Q�Ɓj
    If ERR.Number = 9 Then Resume Next
End Sub

Sub WksNumberring()
    Dim c As Long
    Dim wks As Worksheet
    
    c = 0
    For Each wks In ActiveWorkbook.Sheets

        wks.Name = c & "_" & wks.Name
            c = c + 1
    Next wks


End Sub

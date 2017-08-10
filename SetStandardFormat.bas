Attribute VB_Name = "SetStandardFormat"
Option Explicit
'�������E�����ɂ���A�x�^�ł�������F�A�v�Z�������ɂ���
'���o����͍�����
'�G���[�͐Ԏ�
'

Type CellItem
    Address As Long
    Item As String
End Type

Sub StandardTableFomatting()
Attribute StandardTableFomatting.VB_Description = "�\�����₷���t�H�[�}�b�g����B"
Attribute StandardTableFomatting.VB_ProcData.VB_Invoke_Func = "F\n14"
    Dim r As Range
    Dim v As Variant
'    v = GetTargetRange
    Set r = Selection
    If r.Rows.Count < 2 Then Exit Sub
'    MsgBox "�Z�����̔z���ԁF" & IsArrayEx(v)
    Call ScreenUpdatingSwitch
    Call InitialTextCentering(r)
    Call DrawBordersInsideHorizontalHairline(r)
    Call RightCenteringNumerbers(r)
    Call FontSetting(r)
    Call RowHeight18pt(r)
    Call ColumnWidthAutofit(r)
    Call TableShade(r)
    Call LabelCentering(r)
    Call ScreenUpdatingSwitch
    
    Set r = Nothing
End Sub
Private Function GetTargetRange() As Variant
    Dim rng As Range
    Dim Row As Long, col As Long
    Dim v As Variant
    Set rng = Selection
    If rng.Rows.Count < 2 Then Exit Function
    v = rng
    
    Dim Lists() As CellItem
    Dim AddressLists() As String
    Dim StrLists() As String
    Dim c As Long, i As Long
    Dim LeftTop As Range
    Dim StartRow As Long, StartCol As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim buf As String
    StartRow = rng.Item(1).Row
    StartCol = rng.Item(1).column
    rowCount = rng.Rows.Count
    colCount = rng.Columns.Count
    
    For Row = StartRow To rowCount
        For col = StartCol To colCount
            
            If Not IsError(Cells(Row, col)) Then
                If Len(Cells(Row, col)) > 0 Then
                    ReDim Preserve AddressLists(c) As String
                    ReDim Preserve StrLists(c) As String
                    buf = Cells(Row, col)
                    AddressLists(c) = Cells(Row, col).Address
                    StrLists(c) = buf
                End If
            End If
        Next col
    Next Row
    
    ReDim Lists(UBound(StrLists)) As CellItem
    For i = LBound(Lists) To UBound(StrLists)
        Lists(i).Address = AddressLists(i)
        Lists(i).Item = StrLists(i)
    Next
    
    '�z��ɂł��Ȃ�������A�񎟔z��ɓ��{�̈ꎞ�z������[�v���ē������@���g����OK
'    GetTargetRange = Lists
End Function

Private Sub InitialTextCentering(rng As Range)
    rng.HorizontalAlignment = xlRight
End Sub

Private Sub LabelCentering(rng As Range)
    Dim HeaderRow As Range
    Dim HeaderCol   As Range
    
    Set HeaderRow = Range(rng.Cells(1, 1), rng.Cells(1, rng.Columns.Count))
    Set HeaderCol = Range(rng.Cells(1, 1), rng.Cells(rng.Rows.Count, 1))
    
    HeaderRow.HorizontalAlignment = xlRight
    HeaderRow.Font.color = vbBlack
    
    HeaderCol.HorizontalAlignment = xlLeft
    HeaderCol.Font.color = vbBlack
    Set HeaderRow = Nothing
    Set HeaderCol = Nothing

End Sub
Private Sub DrawBordersInsideHorizontalHairline(rng As Range)

    rng.Borders.LineStyle = True
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlLineStyleNone
'        .LineStyle = xlContinuous
'        .Weight = xlHairline
    End With
    
    With rng
        .Borders(xlInsideVertical).LineStyle = xlLineStyleNone
        .Borders(xlEdgeLeft).LineStyle = xlLineStyleNone
        .Borders(xlEdgeRight).LineStyle = xlLineStyleNone
    End With
End Sub

Private Sub RightCenteringNumerbers(rng As Range)

    Dim r As Range, rng2 As Range
    Dim rConstants As Range, rFormulas As Range
    Dim rNumeric As Range
    rng.Font.color = vbBlack
    On Error Resume Next
    
    Set rFormulas = rng.SpecialCells(xlCellTypeFormulas) '�t���O�Ŕ��肪�K�v����
    Set rConstants = rng.SpecialCells(xlCellTypeConstants)
    If rConstants Is Nothing And rFormulas Is Nothing Then Exit Sub

    If Not rFormulas Is Nothing Then
        Set rng2 = Union(rFormulas, rConstants)
    Else
        Set rng2 = rConstants
    End If
    On Error GoTo 0
    If rng2 Is Nothing Then End
    For Each r In rng2
        Select Case IsError(r.Value)
            Case False
                If IsNumeric(r) Then
                    If Not rNumeric Is Nothing Then
                        Set rNumeric = Union(rNumeric, r)
                    Else
                        Set rNumeric = r
                    End If
                End If
            Case True
        End Select
    Next r
    
    If Not rNumeric Is Nothing Then
        rNumeric.HorizontalAlignment = xlRight
        rNumeric.Font.color = vbBlack
        If Not rFormulas Is Nothing Then rFormulas.Font.color = vbBlue
    End If
    
    Set rNumeric = Nothing
    Set rng2 = Nothing
    Set rConstants = Nothing
    Set rFormulas = Nothing
End Sub

Private Sub RowHeight18pt(rng As Range)
    rng.RowHeight = 18
End Sub
Private Sub ColumnWidthAutofit(rng As Range)
    rng.Columns.AutoFit
End Sub
Private Sub TableShade(rng)
    Const DefaultColor As String = &HF2F2F2
    Dim InteriorColor As Long
    Dim r As Integer
    Dim lastCol As Integer
    Dim FirstRow As Integer
    Dim firstCol As Integer
    Dim LastRow As Integer

    '�I��͈͂̍s�����݂ɔ��D�œh��Ԃ�
    rng.Interior.ColorIndex = xlNone '�܂��h��Ԃ��Ȃ�

    InteriorColor = DefaultColor
    FirstRow = rng(1).Row
    firstCol = rng(1).column
    LastRow = rng(rng.Count).Row
    lastCol = rng(rng.Count).column
    
    For r = FirstRow + 1 To LastRow Step 2
        Range(Cells(r, firstCol), _
        Cells(r, lastCol)).Interior.color = InteriorColor '�萔�œh��Ԃ�������`
    Next r
End Sub

Private Sub HeaderRowFormatting(FirstRow)
    FirstRow.Font.Bold = True
End Sub
Private Sub FontSetting(rng)
    Dim i As Long
    Dim r As Range
    Dim r1 As Range, r2 As Range, r3 As Range, rError As Range, rAlphaNumeric As Range
    Dim rngHasData As Range
    Dim objChar As Object
    
    rng.Font.Size = 10
    On Error Resume Next '�v�Z�����܂ރZ�����Ȃ��ƃG���[�ɂȂ�
    Set r1 = rng.SpecialCells(xlCellTypeConstants)
    Set r2 = rng.SpecialCells(xlCellTypeFormulas)
    
    r1.Font.color = vbBlack
    If Not r2 Is Nothing Then r2.Font.color = vbBlue
    
    If Not r2 Is Nothing Then
        Set rngHasData = Union(r1, r2, Range("a1"))
    Else
        Set rngHasData = Union(r1, Range("a1"))
    End If
    If rngHasData Is Nothing Then End
    On Error GoTo 0
    For Each r In rngHasData
        Select Case IsError(r.Value)
            Case True
                If rError Is Nothing Then Set rError = r Else: Set rError = Union(rError, r)
            Case False
                DoEvents
                On Error Resume Next
                Select Case IsNumeric(r)
                    Case True
                        r.Font.Name = "Arial"
                    Case False
'                       �S�p���܂ނ��ǂ����]������
'                       �Z���ɑS�p�������܂܂�Ă��邩�ǂ����𔻒肵�������򂳂���B
'                       True�F�ꕶ�����Õ]������    False�F�Z����Arial�ɓK�p����
                        Dim HasJAChar As Boolean
                        HasJAChar = IsZenkaku(r.Value)
                        If HasJAChar Then
'�{�{�{�{�{�{�{�{�{�{�{�{�{�{�a���Ɖp���ɕʂ̃t�H���g��K�p����ӏ��{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{
'                            For i = 1 To r.Characters.Count
'                            Set objChar = r.Characters(i, 1)
'�@                              �O�҂̏����݂̂ł͐����ɑS�p�ݒ肪�K�p����Ă��܂�
'                                If LenB(StrConv(objChar.Text, vbFromUnicode)) = 1 Then
'                                    objChar.Font.Name = "Arial"
'                                Else
'                                    objChar.Font.Name = "�l�r �o�S�V�b�N"
'                                End If
'                            Next i
'�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{�{
                        Else
                            If rAlphaNumeric Is Nothing Then Set rAlphaNumeric = r Else: Set rAlphaNumeric = Union(rAlphaNumeric, r)
                        End If
                End Select
        End Select
    Next r
    If Not rError Is Nothing Then rError.Font.color = vbRed
    If Not rAlphaNumeric Is Nothing Then rAlphaNumeric.Font.Name = "Arial"
    
    Set rngHasData = Nothing
    Set rError = Nothing
    Set rAlphaNumeric = Nothing
    Set r1 = Nothing
    Set r2 = Nothing
    Set objChar = Nothing
    
End Sub
Private Sub ScreenUpdatingSwitch()
    Application.ScreenUpdating = Not Application.ScreenUpdating
End Sub
'***********************************************************
' �@�\   : �������z�񂩔��肵�A�z��̏ꍇ�͋󂩂ǂ��������肷��
' ����   : varArray  �z��
' �߂�l : ���茋�ʁi1:�z��/0:��̔z��/-1:�z�񂶂�Ȃ��j
'***********************************************************
Private Function IsArrayEx(varArray As Variant) As Long
    On Error GoTo ERROR_

    If IsArray(varArray) Then
        IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function

ERROR_:
    If ERR.Number = 9 Then
        IsArrayEx = 0
    End If
End Function

'��IsZenkaku
'���@�\�F������ɑS�p�������܂܂�Ă��邩���ׂ�B
'�������FValue ���ׂ�Ώۂ̕�����B
'���߂�l�F�S�p�������܂܂�Ă���ꍇ��True�A�����łȂ��ꍇFalse�B
Private Function IsZenkaku(ByVal Value As String) As Boolean
Dim ByteLength As Long

ByteLength = LenB(StrConv(Value, vbFromUnicode))

If Len(Value) <> ByteLength Then

IsZenkaku = True

End If

End Function

Attribute VB_Name = "QTA_InsertHeader"
#If VBA7 Then
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Option Explicit

Dim mbCancel As Boolean
Dim mlLastCol As Long
Dim mlLastRow As Long
Dim mbCancelEvent As Boolean

'---------------------------------------------------------------------------------------
' Method : InsertHeader
' Author : temporary3
' Date   : 2016/02/10
' Purpose: �A�N�e�B�u�V�[�g�Ƀw�b�_�[��}������B
'---------------------------------------------------------------------------------------
Sub InsertHeader()
    Dim CurrentActivecell   As Range
    Dim myLastRow As Long
    Dim myPageCount As Long
    Dim wks As Worksheet
    Dim myHeaderSpace As Range
    
    Set wks = ActiveWorkbook.ActiveSheet
    
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    mlLastRow = GetLastRow(wks) '�V�[�g�̈�������őւ���
    mlLastCol = GetLastCol(wks)
    myPageCount = GetPageBreak(wks)
    
    If TypeName(Selection) = "Range" Then Set CurrentActivecell = Selection
    
    Set myHeaderSpace = InsertRows(myPageCount)
    
    Call DrawLines(myPageCount)
    Call InsertDocNo(myPageCount)
    Call PasteLabels(myPageCount)
    Call InsertWincklerText(myPageCount)
    Call AddWincklerlogo(myPageCount)
    Call BreakPages(myPageCount)                '�w�b�_�[�}����̉��y�[�W�ݒ�
    Call RowsHeightAdjustment(myPageCount)
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    If Not CurrentActivecell Is Nothing Then CurrentActivecell.Select
    
    MsgBox "�w�b�_�[��" & myPageCount & "�y�[�W�܂Œǉ����܂���", vbOKOnly + vbInformation, "�w�b�_�[�}������"

End Sub

'---------------------------------------------------------------------------------------
' Method : GetLastRow
' Author : mokoo
' Date   : 2016/02/20
' Purpose: �V�[�g�̍ŉ��s���擾����
'---------------------------------------------------------------------------------------
Private Function GetLastRow(ByVal st As Worksheet) As Long
    Dim rngPrintArea As Range
    
    On Error GoTo ErrHandler
    Set rngPrintArea = st.Range(st.PageSetup.PrintArea)
    On Error GoTo 0
    GetLastRow = rngPrintArea.Item(rngPrintArea.Count).Row

Exit Function
 
ErrHandler:
 
 mbCancelEvent = True
 
 End Function
 
 Private Function GetLastCol(ByVal st As Worksheet) As Long
    Dim rngPrintArea As Range

    On Error GoTo ErrHandler
    Set rngPrintArea = st.Range(st.PageSetup.PrintArea)
    On Error GoTo 0
    GetLastCol = rngPrintArea.Item(rngPrintArea.Count).column

Exit Function
ErrHandler:
    MsgBox "�w�b�_�[������܂���"
    
End Function
'---------------------------------------------------------------------------------------
' Method : GetPageCount
' Author : temporary3
' Date   : 2016/04/07
' Purpose: �A�N�e�B�u�V�[�g�̌��Ϗ��Ƃ��Ẵy�[�W�����J�E���g����i60�s/�y�[�W�Ōv�Z�j
'---------------------------------------------------------------------------------------
Private Function GetPageCount() As Long
    Dim lPageCount As Long
    Dim rowCount As Long
    Dim lPageMargin As Long
    
        
    rowCount = ActiveWorkbook.ActiveSheet.UsedRange.Rows.Count

    If rowCount < 60 Then
        MsgBox "1�y�[�W��������܂���"
        End    '�s����1�y�[�W�����Ȃ�I��
    End If
    
    '�K��s�����傤�ǂ��܂ރy�[�W��
    lPageCount = (rowCount / 60)
    '�]��y�[�W
    If (rowCount Mod 60) <> 0 Then lPageMargin = 1
    
    '���v�y�[�W��
    GetPageCount = lPageCount + lPageMargin
    
End Function

Private Function GetPageBreak(ByVal wks As Worksheet) As Long
    
    GetPageBreak = wks.HPageBreaks.Count

End Function

'---------------------------------------------------------------------------------------
' Method : InsertRows
' Author : temporary3
' Date   : 2016/02/10
' Purpose: �w�b�_�[�p�X�y�[�X�ƂȂ�s��ǉ�����B�i12�s/�w�b�_�[�j
'---------------------------------------------------------------------------------------
Private Function InsertRows(ByVal myPageCount As Long) As Range

    Dim i As Long
    Dim iRowCountHeader As Long
    Dim rngRows As Range
    
    iRowCountHeader = 12    '�w�b�_�[�̍s��

    Set rngRows = Rows("1000:1012")

    For i = 60 To myPageCount * 60 Step 60    '�ŏI�Z���̔ԍ����疖�������߂�
        DoEvents
        rngRows.Copy
        Rows(i).insert
        
        If InsertRows Is Nothing Then
             Set InsertRows = Rows(i & ":" & i + 11)
        Else
            Set InsertRows = Union(InsertRows, Rows(i & ":" & i + 11))
        End If
    Next i
                                                                                                                                              
End Function
'---------------------------------------------------------------------------------------
' Method : DrawLines
' Author : temporary3
' Date   : 2016/02/10
' Purpose: �w�b�_�[�̌r���������B
'---------------------------------------------------------------------------------------
Private Sub DrawLines(ByVal myPageCount As Long)

    Dim i As Long

    For i = 64 To myPageCount * 64 Step 60

        '�w�b�_�[�̌r����4�{�Ђ�
        '�����A�ׁA�ׁA��

        With Range(Cells(i, 1), Cells(i, mlLastCol)).Borders(xlEdgeBottom)  '���S�̉��̂͂��߂̌r���͑���
            .LineStyle = xlContinuous
            .Weight = xlMedium    '�����炢
            .ColorIndex = xlAutomatic
        End With
        With Range(Cells(i, 1).offset(1, 0), Cells(i, mlLastCol).offset(1, 0)).Borders(xlEdgeBottom)  '��Ԗڂ̌r���ׂ͍�
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Range(Cells(i, 1).offset(5, 0), Cells(i, mlLastCol).offset(5, 0)).Borders(xlEdgeBottom)  '�O�Ԗڂ̌r�����ׂ�
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Range(Cells(i, 1).offset(7, 0), Cells(i, mlLastCol).offset(7, 0)).Borders(xlEdgeBottom)    '��ԉ��̌r�����ׂ�
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next i

End Sub

'---------------------------------------------------------------------------------------
' Method : InsertDocNo
' Author : temporary3
' Date   : 2016/02/10
' Purpose: �w�b�_�[���ƂɌ��Ϗ��̃y�[�W�ԍ�����������
'---------------------------------------------------------------------------------------
Private Sub InsertDocNo(ByVal myPageCount As Long)

    Dim i As Long
    Dim Pages As Long
    Dim rngLabel As Range
    Dim Border1 As Border
    Const strLeftEdgePageNo As String = "CN"
    Const strRightEdgePageNo As String = "DF"
    Dim rngQuotationNo As Range
    Dim rngLabelQTNo As Range
    Dim sPageNumbering As String
    
    Pages = 2    '�y�[�W�̊J�n��2

    For i = 61 To (myPageCount * 60 + 1) Step 60
        '�y�[�W�ԍ������Ϗ��ԍ��̃Z���͈�
        Set rngLabel = Range(Cells(i, strLeftEdgePageNo), Cells(i, strRightEdgePageNo))

    '   1�y�[�W�ɂ��錩�Ϗ��ԍ���Range�ɃZ�b�g
        On Error GoTo NoNotFound
        Dim r
        Set r = Rows("1:6").Find("Nagoya")
        Set rngQuotationNo = Cells(r.Row - 1, r.column).Value
        Set rngLabelQTNo = Cells(i, strLeftEdgePageNo)
        
        rngQuotationNo.Copy Destination:=rngLabelQTNo
        rngQuotationNo.HorizontalAlignment = xlLeft
        
        On Error GoTo 0
      
        '���Ϗ��ԍ���2��E�Ƀy�[�W��������
        sPageNumbering = "'" & CStr(Pages) & "/" & CStr(myPageCount + 1)
        Cells(i, strRightEdgePageNo).offset(0, -2).Value = sPageNumbering
        Pages = Pages + 1
        
        '�t�H���g�̐ݒ�
        With rngLabel.Font
            .Name = "MS ����"
            .Size = 10
        End With

        '�r��
        Set Border1 = rngLabel.Borders(xlEdgeBottom)
        With Border1
            .LineStyle = xlDash
            .Weight = xlThin
        End With
    Next i
Exit Sub

NoNotFound:
    MsgBox "���Ϗ��ԍ����擾�ł��܂���ł����B", vbOKOnly + vbCritical

End Sub

Private Sub PasteLabels(ByVal myPageCount As Long)
'�w�b�_�[�̃e�L�X�g�����܂��B

    Dim i As Long
    '�e�L�X�g�͈̔͂��w�肷��ϐ�
    Dim Item As Range
    Dim Description As Range
    Dim POS As Range
    Dim QTY As Range
    Dim UnitPrice As Range
    Dim TotalPrice As Range
    Dim Komoku As Range
    Dim Naiyo As Range
    Dim Designation As Range
    Dim rngHeaderWidth As Range
    Dim rngHeaderArea As Range

    Dim myCol_Item As Long
    Dim myCol_POS As Long
    Dim myCol_Total As Long
    
    '���Ϗ��̊e�v�f�̗�ԍ����擾����
    myCol_Item = ���Ϗ����ڂ̗�ԍ����擾(myCol_POS, myCol_Total)
    
    For i = 67 To (myPageCount * 60 + 7) Step 60
 
    ''�擾������ԍ����g�p���Č��o���e�L�X�g�͈̔͂��w��
    Set Item = Cells(i, myCol_Item)
    Set POS = Cells(i, myCol_POS)
    Set TotalPrice = Cells(i, myCol_Total)
    Set Komoku = Cells(i, "C")
    Set Naiyo = Cells(i, "BJ")
    Set Description = Cells(i, "bj").offset(1, 0)
    Set QTY = Cells(i, "BQ").offset(4, 0)
    Set UnitPrice = Cells(i, "BY").offset(4, 0)
    Set Designation = Cells(i, "S").offset(4, 0)


    '�w�b�_�e�L�X�g�̃t�H���g�ƃT�C�Y
    Set rngHeaderWidth = Range(Cells(i - 7, "A"), Cells(i - 7, mlLastCol))
    Set rngHeaderArea = Range(rngHeaderWidth, rngHeaderWidth.offset(11, 0))
    
    With rngHeaderArea
        '.HorizontalAlignment = xlLeft
        .Font.Name = "�l�r �o����"
        .Font.Size = 10
        .Interior.color = RGB(238, 238, 238)
    End With

    '�w�b�_�e�L�X�g�̏������݂ƈʒu����
    '�@
    With Komoku
        .Value = "���@�@�@��"
        .Font.Bold = False
    End With

    Range(Cells(i, "C"), Cells(i, "U")).HorizontalAlignment = xlCenterAcrossSelection
    '�A
    With Item
        .Value = "Item"
        .Font.Bold = False
    End With
    Range(Cells(i, "C").offset(1, 0), Cells(i, "U").offset(1, 0)).HorizontalAlignment = xlCenterAcrossSelection
    
    '�B
    With Naiyo
        .Value = "���@�@�@�@�e"
        .Font.Bold = False
    End With
    Range(Cells(i, "BJ"), Cells(i, "CV")).HorizontalAlignment = xlCenterAcrossSelection
    
    '�C
    With Description
        .Value = "Description"
        .Font.Bold = False
    End With
    Range(Cells(i, "BJ").offset(1, 0), Cells(i, "CV").offset(1, 0)).HorizontalAlignment = xlCenterAcrossSelection
    
    '�D
    With POS
        .Value = "Pos"
        .Font.Bold = False
    End With
    Range(Cells(i, "C").offset(4, 0), Cells(i, "F").offset(4, 0)).HorizontalAlignment = xlCenterAcrossSelection
    '�E
    With Designation
        .Value = "�i�@  �@��"
        .Font.Bold = False
    End With
    Range(Cells(i + 4, "X"), Cells(i + 4, "AF")).HorizontalAlignment = xlCenterAcrossSelection
    '�F
    With QTY
        .Value = "���@��"
        .Font.Bold = False
    End With
    Range(Cells(i + 4, "BQ"), Cells(i + 4, "BW")).HorizontalAlignment = xlCenterAcrossSelection
    '�G
    With UnitPrice
        .Value = "�P�@�@��"
        .Font.Bold = False
    End With
    Range(Cells(i + 4, "BY"), Cells(i + 4, "CO")).HorizontalAlignment = xlCenterAcrossSelection
    '�H
    With TotalPrice
        .Value = "���@�@�i"
        .Font.Bold = False
    End With
    Range(Cells(i + 4, "CR"), Cells(i + 4, "DK")).HorizontalAlignment = xlCenterAcrossSelection

    Next i

End Sub

'---------------------------------------------------------------------------------------
' Method : InsertWincklerText
' Author : temporary3
' Date   : 2016/02/10
' Purpose: ��Ж��̃e�L�X�g����������
'---------------------------------------------------------------------------------------
Private Sub InsertWincklerText(ByVal myPageCount As Long)

    Dim i As Long

    For i = 61 To (myPageCount * 61) Step 60
        With Cells(i, "M")
            .Value = "�E�C���N�����������"
            .Font.Size = 14
            .Font.Name = "�l�r �o����"
            .Font.Bold = False
        End With
        With Cells(i, "M").offset(1, 0)
            .Value = "WINCKLER & CO, LTD"
            .Font.Size = 14
            .Font.Name = "�l�r �o����"
            .Font.Bold = False
        End With
    Next i

End Sub

'---------------------------------------------------------------------------------------
' Method : RowsHeightAdjustment
' Author : temporary3
' Date   : 2016/02/10
' Purpose: �w�b�_�[�̍s����ݒ肷��
'---------------------------------------------------------------------------------------
Private Sub RowsHeightAdjustment(ByVal myPageCount As Long)

    Dim i As Long

    For i = 60 To (myPageCount * 60) Step 60
        Rows(i).RowHeight = 12.5
        Rows(i + 1).RowHeight = 15.5    '�u�E�C���N����������Ёv
        Rows(i + 2).RowHeight = 15    'Winckler & Co, Ltd.
        Rows(i + 3).RowHeight = 12.5
        Rows(i + 4).RowHeight = 12.5
        Rows(i + 5).RowHeight = 3
        Rows(i + 6).RowHeight = 9.5
        Rows(i + 7).RowHeight = 14    '����
        Rows(i + 8).RowHeight = 14    'Item
        Rows(i + 9).RowHeight = 9.5
        Rows(i + 10).RowHeight = 12.5
        Rows(i + 11).RowHeight = 12    'Pos
    Next i

End Sub
'---------------------------------------------------------------------------------------
' Method : AddWincklerlogo
' Author : temporary3
' Date   : 2016/02/10
' Purpose: �E�C���N�����̃��S��\��t����
'---------------------------------------------------------------------------------------
Private Sub AddWincklerlogo(ByVal myPageCount As Long)

'�E�C���N�����̃��S������̈ʒu�ɓ\��t���܂��B

    Dim myFileName As String
    Dim myPic As Shape
    Dim i As Long
    Dim myFileSheet As Worksheet
    Dim spLogo As Shape
    Dim sShapePositionLEFT As Double
    Dim sShapePositionTOP As Double
    
    '���S�̃t�@�C����
    myFileName = "winckler_logo"
    '���S�̃p�X
    Set myPic = ThisWorkbook.Sheets("pictures").Shapes("winckler_logo")
    'Set myFileSheet = ThisWorkbook.Worksheets("pictures")    '�ҏW���̃u�b�N�Ɠ����f�B���N�g���ɂ���摜�t�@�C�����w�肵�Ă��܂��B
    
    '���S���N���b�v�{�[�h�ɓ����
    myPic.Copy
    'With myFileSheet
    '    Set spLogo = .Shapes(myFileName)
    '    spLogo.Copy
    'End With

    On Error Resume Next    '�Ȃ����G���[���ł邩��B
    For i = 61 To (myPageCount * 61) Step 60
        
        With ActiveWorkbook.ActiveSheet
            .Cells(i, "C").Select
            On Error GoTo ErrHandler
TryAgain:
            .Pastespecial Format:="�} (PNG)", Link:=False, DisplayAsIcon:=False '���s���G���[�P�O�O�S������
            
            '���������ܓ\��t�����}�̈ʒu���擾���āA�ύX����
            sShapePositionLEFT = .Shapes(.Shapes.Count).Left
            sShapePositionTOP = .Shapes(.Shapes.Count).Top
        
            With .Shapes(.Shapes.Count) '�}�̈ʒu�������
                .Left = sShapePositionLEFT - 4.5
                .Top = sShapePositionTOP + 2
            End With
        End With
        On Error GoTo 0
    Next i
    
Exit Sub

ErrHandler:
    If ERR.Number = 1004 Then GoTo TryAgain

End Sub
'---------------------------------------------------------------------------------------
' Method : BreakPages
' Author : temporary3
' Date   : 2016/02/10
' Purpose: �w�b�_�[�}����̉��y�[�W�ݒ�
'---------------------------------------------------------------------------------------
Private Sub BreakPages(ByVal myPageCount As Long)
    Dim i As Long
    Dim Pages As Long
    Dim rng As Range
    Dim RowsPerPage As Long

    RowsPerPage = 60    '1�y�[�W�̍s��
    Pages = 0
    Application.PrintCommunication = False
    With ActiveSheet
        '���݂̉��y�[�W�����ׂă��Z�b�g
        .ResetAllPageBreaks
        For i = 60 To (myPageCount * 60) Step RowsPerPage
            .HPageBreaks.Add before:=.Cells(i, 1)
        Next i
    End With
    Application.PrintCommunication = True
End Sub

Private Function ���Ϗ����ڂ̗�ԍ����擾(ByRef myPosCol As Long, myTotalCol As Long) As Long
    Dim bkTemp As Workbook
    Dim wksQuotation As Worksheet
    Dim wksTemp As Worksheet
    Dim wksEval As Worksheet
    Dim myArray As Variant
    Dim myItemCol As Long
    
    Const bkName As String = "Temp"
    
    Set wksQuotation = ActiveWorkbook.ActiveSheet   '�A�N�e�B�u�V�[�g���Ώی��Ϗ��Ƃ���
    Set bkTemp = Workbooks.Add                          '���̓f�[�^��W�J����ꎞ�V�[�g������
    Set wksTemp = bkTemp.Sheets(1)
    wksTemp.Name = bkName
    
    Set wksEval = �f�[�^�������o�ƃ����N�t��(wksQuotation)
    If mbCancel Then Exit Function '�A�N�e�B�u�V�[�g�����Ϗ��łȂ���ΏI������
    myArray = GetExtractedCols(wksEval)  '�f�[�^�����̏�ʗ�����o��
    ���Ϗ����ڂ̗�ԍ����擾 = GetItemCol(wksTemp) '�i���������ׂ�
    myPosCol = GetPosCol(wksTemp) 'Pos�������ׂ�
    myTotalCol = GetTotalCol(wksQuotation, wksTemp)
    
    MsgBox "�i���̗�ԍ���" & ���Ϗ����ڂ̗�ԍ����擾 & "�ł��B"
    MsgBox "POS���" & myPosCol & "�ł��B"
    MsgBox "Total���" & myTotalCol & "�ł��B"
    
    Application.DisplayAlerts = False
    bkTemp.Close
    Set bkTemp = Nothing
    Application.DisplayAlerts = True
    
    Set wksQuotation = Nothing
    Set wksTemp = Nothing
    
End Function

Private Function �f�[�^�������o�ƃ����N�t��(ByRef wks As Worksheet) As Worksheet
 '�w��͈͂̃f�[�^��������������
    
    Dim lCol As Long
    Dim myRow As Long
    Dim lEndCol As Long
    Dim lFirstRow As Long
    Dim lEndRow As Long
    Dim myCountA As Long
    Dim myPrintRange As Range
    Dim lCount As Long
    Const lStartRow As Long = 36
    
    '���Ϗ��̈���͈͂��擾
    On Error GoTo ErrHandler
    Set myPrintRange = wks.Range(wks.PageSetup.PrintArea)
    On Error GoTo 0
    lEndCol = myPrintRange.Item(myPrintRange.Count).column

    
    '�ꎞ�V�[�g�ɒ��o���ʂ���������
    myRow = 1
    For lCol = 1 To lEndCol
        
        '�f�[�^����������ׂ�
        myCountA = WorksheetFunction.CountA(myPrintRange.Range(myPrintRange.Cells(lStartRow, _
                        lCol), myPrintRange.Cells(lEndRow, lCol)))

        If myCountA > 0 Then
            With Workbooks(Workbooks.Count).Worksheets(1)
                .Cells(1, 1).offset(myRow, 0).Value = lCol
                .Cells(1, 1).offset(myRow, 1).Value = myCountA
            End With
            myRow = myRow + 1
        End If
        
        '���ϕ��������𒲂ׂ�
        If myCountA > 0 Then Call �����������ς��v�Z(lCol, wks)
    Next lCol
    
'   �ꎞ�V�[�g�̏������ƂƂ̂���
    With Workbooks(Workbooks.Count).Worksheets(1)
        .Cells(1, 1).Value = "���"
        .Cells(1, 2).Value = "�f�[�^����"
        .Cells(1, 3).Value = "���ϕ�������"
        .Columns("A:C").EntireColumn.AutoFit
        .Range("A1").CurrentRegion.Sort _
            key1:=Range("b2"), order1:=xlDescending
    End With

    Set �f�[�^�������o�ƃ����N�t�� = Workbooks(Workbooks.Count).Worksheets(1)

Exit Function

ErrHandler:
MsgBox "�����ȃV�[�g�ł�"
mbCancel = True
End Function

Private Function GetExtractedCols(ByVal wks As Worksheet) As Variant
    Dim myArray() As Long
    Dim i As Long
    Dim myRow As Long
    
    Const myExtractingRows As Long = 5 '���o�����
    Const strCol As String = "B"            '�H�H�H
    
    ReDim myArray(myExtractingRows) As Long
    
    With wks
        i = 0
        myRow = 2  '���o���s������
        For i = 0 To myExtractingRows
            myArray(i) = .Cells(myRow, strCol).Value
            myRow = myRow + 1
        Next i
    End With

    GetExtractedCols = myArray

End Function

Private Function GetItemCol(ByRef wks As Worksheet) As Long
    Dim i As Long
    Dim rngAveStringLen As Range
    Dim rng As Range
    Dim lMaxStringLen As Long
    
    Set rngAveStringLen = wks.Range(wks.Cells(2, 3), wks.Cells(2, 3).End(xlDown))
    For Each rng In rngAveStringLen
        If rng.Value > lMaxStringLen Then
            lMaxStringLen = rng.Value
            GetItemCol = rng.offset(0, -2) '���ϕ����������Œ��̗�̔ԍ���Ԃ�
        End If
    Next rng

    Set rngAveStringLen = Nothing
    Set rng = Nothing
    
End Function

Private Function GetPosCol(ByRef wks As Worksheet)
    Dim i As Long
    Dim rngColNum As Range
    Dim rng As Range
    Dim lMinColNum As Long
    Dim buf As Long
    
    Set rngColNum = wks.Range(wks.Cells(2, 1), wks.Cells(2, 1).End(xlDown))
    
    lMinColNum = 10
    For Each rng In rngColNum
        If rng.Value < lMinColNum Then
            lMinColNum = rng.Value
            GetPosCol = rng.Value '�ŏ��̗�ԍ���Ԃ�
        End If
    Next rng
    
    Set rngColNum = Nothing
    Set rng = Nothing
    
End Function

Private Function GetTotalCol(ByRef wksQuotation As Worksheet, ByRef wksTemp As Worksheet)
    Dim i As Long
    Dim rngColNum As Range
    Dim rng As Range
    Dim lMaxColNum As Long
    Dim buf As Long
    Dim rngPrintArea As Range
    Dim myPrintAreaRightEdgeCol As Long
    
    '����͈͂̉E�[�̐؂�ڂ̗�
    Set rngPrintArea = wksQuotation.Range(wksQuotation.PageSetup.PrintArea)
    myPrintAreaRightEdgeCol = rngPrintArea.Item(rngPrintArea.Count).column
    
    Set rngColNum = wksTemp.Range(wksTemp.Cells(2, 1), wksTemp.Cells(2, 1).End(xlDown))
    
    For Each rng In rngColNum
        If rng.Value > lMaxColNum And rng.Value < myPrintAreaRightEdgeCol Then
            lMaxColNum = rng.Value
            GetTotalCol = rng.Value '����͈͓��ōő�̗�ԍ���Ԃ��B�f�[�^�����͍l���Ȃ��B
        End If
    Next rng
    
    Set rngPrintArea = Nothing
    Set rngColNum = Nothing
    Set rng = Nothing
    
End Function

Private Sub �Z���̕\���`��������ׂ�()
'�\���`��������ׂ�B�P���ƍ��v�́A
'�����̕\���`���������ł��邱�ƂŔ���ł���Ǝv���B
    
    Dim i As Long
    Dim myFormat As String
    myFormat = ActiveCell.NumberFormat
    
End Sub
Private Sub �����������ς��v�Z(ByVal lCol As Long, ByRef wksQuotation As Worksheet)
'CountA�Ŋ���o�����񂠂���̃f�[�^����ʐ����̂����A_
'�����Ƃ����ϕ��������������񂪕i���̗񂾁B
    Dim i As Long
    Dim lRow As Long
    Dim lEndRow As Long
    Dim lEndRowTemp As Long
    Dim myItemCntUsedRange As Long
    Dim rng As Range
    Dim sum As Long
    Dim myAverageLen As Double
    Dim myEndRow As Long
    
    '�ŏI�s
    Dim UsedRng
    Set UsedRng = wksQuotation.UsedRange
    lEndRow = UsedRng.UsedRange.Item(UsedRng.UsedRange.Count).Row
    myItemCntUsedRange = UsedRng.Count
    
    
'   ���߂͌v�Z�Ɋ܂߂Ȃ��F���v���z���ȉ��̍s�̃f�[�^�͖�������
    myEndRow = GetGrandTotalRow(wksQuotation)  '���v���z�̍s���擾
            
    For lRow = 39 To lEndRow
        If Len(wksQuotation.Cells(lRow, lCol).Value) > 0 And lRow < myEndRow Then  '�����񂪂����č��v���z�s����̍s�̂݌v�Z����
            sum = Len(wksQuotation.Cells(lRow, lCol)) + sum
            i = i + 1
        End If
    Next lRow
    
    If i = 0 Then Exit Sub '��̗�Ȃ玟��
    
    myAverageLen = sum / i  '���������̍��v�������ł��
    
    Dim a As Worksheet
    Set a = ActiveSheet
'   �ꎞ�V�[�g���Q��
    lEndRowTemp = a.Cells(Rows.Count, 1).End(xlUp).Row + 1
    Set rng = a.Range(a.Cells(2, 1), a.Cells(lEndRowTemp, 1)).Find(lCol).offset(0, 2) '�Y����́u���ϕ��������v�̓��̓Z�����擾
    rng.Value = myAverageLen
    
    Debug.Print lCol & "�̕��ϕ��������F" & myAverageLen
    Set rng = Nothing
    Set a = Nothing

End Sub

Private Function GetGrandTotalRow(ByRef wksQuotation As Worksheet) As Long
    Dim rngNumerics As Range
    Dim rng As Range
    Dim lMaxNum As Long
    
    With wksQuotation
        Set rngNumerics = .Cells(1, 1).SpecialCells(xlCellTypeFormulas, 1)
        For Each rng In rngNumerics
            '�͈͓��ł̍ő�l
            If rng.Value > lMaxNum Then
                GetGrandTotalRow = rng.Row '�ő�l������s��Ԃ�
            End If
        Next rng
    End With
    
    Set rngNumerics = Nothing
    Set rng = Nothing

End Function

Private Sub CountMergeColumn()
    Dim i As Long
    Dim buf As String
    
    '�i���ȊO�͂����炭�Z������������Ă���̂Ń��x���̈ʒu���͂�����Ɠ���ł���
    With ActiveCell.MergeArea
        buf = buf & "�Z���̌��F" & .Columns.Count & vbCrLf
        buf = buf & "���[�̗�F" & .Item(1).column & vbCrLf
        buf = buf & "�E�[�̗�F" & .Item(.Count).column
    End With
    
    MsgBox buf
    
End Sub

Attribute VB_Name = "�I��͈͓���S�p��"
Option Explicit

Sub �I��͈͓���S�p��()

    '�I��͈͂Ɋ܂܂�锼�p������S�p������
    '�A���t�@�x�b�g�a�����Ƃ��B
    Dim rngTarget As Range
    Dim rng As Range

    Set rngTarget = Selection

    For Each rng In rngTarget.Cells
        rng.Value = StrConv(rng, vbWide)
    Next rng

End Sub

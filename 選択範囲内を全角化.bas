Attribute VB_Name = "選択範囲内を全角化"
Option Explicit

Sub 選択範囲内を全角化()

    '選択範囲に含まれる半角文字を全角化する
    'アルファベット和文字とも。
    Dim rngTarget As Range
    Dim rng As Range

    Set rngTarget = Selection

    For Each rng In rngTarget.Cells
        rng.Value = StrConv(rng, vbWide)
    Next rng

End Sub

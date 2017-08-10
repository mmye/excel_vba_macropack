VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FRMターミナル 
   Caption         =   "見積書を開く..."
   ClientHeight    =   1335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3855
   OleObjectBlob   =   "FRMターミナル.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "FRMターミナル"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mLists As Collection
Dim IndexFile As String
Const UboundArr As Long = 6
Const IndexPath As String = "\\LS410D760\share\◆事務\《1》見積・注文\1. 見積書" '末尾のセパレーターはなしで指定する
Const strInstruction As String = "見積書番号を入力してEnterを押してください"

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
        MsgBox "インデックスが空です"
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
        txtStatus.Caption = "見積書が見つかりません。"
        Exit Sub
    End If
    Path = v(6) & v(5)
    ShName = v(7)
    
    Dim File As String, bk As Workbook
    File = Dir(Path)
    Select Case File
        Case ""
            '見積書ブックが見つからなかった場合
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
                    
                    'リビジョンのシート名をつけるためごちゃごちゃやってる
                    Dim stName As String
                    stName = bk.Sheets(ShName).Name
                    If (Right$(stName, 2) Like "R[0-9]+") Then
                        stName = Replace(st.Name, "R", "")
                        st.Name = stName & "R" & (Right$(stName, 1) + 1)
                    Else
                        'すでにある改訂番を検知する必要がある
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
                txtStatus.Caption = "見積書" & ShName & "を開きました。"
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
    If Dir(IndexFile) = "" Then MsgBox "インデックスファイルが見つかりません。", vbCritical

    On Error Resume Next
    Open IndexFile For Input As #1
    Do Until EOF(1)
        ERR.Number = 0
        Line Input #1, a
        List = Split(a, vbTab)
        Dim Key As String, Content As String
        Key = List(0)
        Content = GetContent(List) 'Content がヘン。Variant配列を返すべきか？？
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


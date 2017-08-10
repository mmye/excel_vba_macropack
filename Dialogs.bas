Attribute VB_Name = "Dialogs"
Option Explicit
Function SelectSingleFileDialog(Optional title As String, Optional buttonName As String, _
    Optional initialFileName As String, Optional filters As Collection = Nothing, _
    Optional filterIndex As Long = -1) As String
'title - the title for the dialog
'buttonName - name of the action button. Warning does not always work
'initialFileName - initial file path for the dialog e.g. C:\
'filters - file dialog filters
'filterIndex - index of initially selected filterIndex
    Dim fDialog As FileDialog, result As Integer, it As Variant
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.AllowMultiSelect = False
    'Properties
    If Not (filters Is Nothing) Then
    fDialog.filters.Clear
        For Each it In filters
            fDialog.filters.Add it(0), it(1)
        Next it
        If filterIndex <> -1 Then fDialog.filterIndex = filterIndex
    End If
    If buttonName <> vbNullString Then fDialog.buttonName = buttonName
    If initialFileName <> vbNullString Then fDialog.initialFileName = initialFileName
    If title <> vbNullString Then fDialog.title = title
    'Show
    If fDialog.Show = -1 Then
        SelectSingleFileDialog = fDialog.SelectedItems(1)
    End If
End Function
Function SelectMultiFileDialog(Optional title As String, Optional buttonName As String, _
    Optional initialFileName As String, Optional filters As Collection = Nothing, _
    Optional filterIndex As Long = -1) As FileDialogSelectedItems
'title - the title for the dialog
'buttonName - name of the action button. Warning does not always work
'initialFileName - initial file path for the dialog e.g. C:\
'filters - file dialog filters
'filterIndex - index of initially selected filterIndex
    Dim fDialog As FileDialog, result As Integer, it As Variant
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    fDialog.AllowMultiSelect = True
    'Properties
    If Not (filters Is Nothing) Then
    fDialog.filters.Clear
        For Each it In filters
            fDialog.filters.Add it(0), it(1)
        Next it
        If filterIndex <> -1 Then fDialog.filterIndex = filterIndex
    End If
    If buttonName <> vbNullString Then fDialog.buttonName = buttonName
    If initialFileName <> vbNullString Then fDialog.initialFileName = initialFileName
    If title <> vbNullString Then fDialog.title = title
    'Show
    If fDialog.Show = -1 Then
        Set SelectMultiFileDialog = fDialog.SelectedItems
    End If
End Function
Function SelectFolderDialog(Optional title As String, Optional buttonName As String, _
    Optional initialFileName As String) As String
'title - the title for the dialog
'buttonName - name of the action button. Warning does not always work
'initialFileName - initial folder path for the dialog e.g. C:\
    Dim fDialog As FileDialog, result As Integer, it As Variant
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    'Properties
    If buttonName <> vbNullString Then fDialog.buttonName = buttonName
    If initialFileName <> vbNullString Then fDialog.initialFileName = initialFileName
    If title <> vbNullString Then fDialog.title = title
    'Show
    If fDialog.Show = -1 Then
        SelectFolderDialog = fDialog.SelectedItems(1)
    End If
End Function
Function OpenFileDialog(Optional title As String, Optional buttonName As String, _
    Optional initialFileName As String, Optional filters As Collection = Nothing, _
    Optional filterIndex As Long = -1) As String
'title - the title for the dialog
'buttonName - name of the action button. Warning does not always work
'initialFileName - initial file path for the dialog e.g. C:\
'filters - file dialog filters
'filterIndex - index of initially selected filterIndex
    Dim fDialog As FileDialog, result As Integer, it As Variant
    Set fDialog = Application.FileDialog(msoFileDialogOpen)
    'Properties
    If Not (filters Is Nothing) Then
    fDialog.filters.Clear
        For Each it In filters
            fDialog.filters.Add it(0), it(1)
        Next it
        If filterIndex <> -1 Then fDialog.filterIndex = filterIndex
    End If
    If buttonName <> vbNullString Then fDialog.buttonName = buttonName
    If initialFileName <> vbNullString Then fDialog.initialFileName = initialFileName
    If title <> vbNullString Then fDialog.title = title
    'Show
    If fDialog.Show = -1 Then
        OpenFileDialog = fDialog.SelectedItems(1)
    End If
End Function
Function SaveAsDialog(Optional title As String, Optional buttonName As String, _
    Optional initialFileName As String, Optional filters As Collection = Nothing, _
    Optional filterIndex As Long = -1) As String
'title - the title for the dialog
'buttonName - name of the action button. Warning does not always work
'initialFileName - initial file path for the dialog e.g. C:\
'filters - file dialog filters
'filterIndex - index of initially selected filterIndex
    Dim fDialog As FileDialog, result As Integer, it As Variant
    Set fDialog = Application.FileDialog(msoFileDialogOpen)
    'Properties
    If Not (filters Is Nothing) Then
    fDialog.filters.Clear
        For Each it In filters
            fDialog.filters.Add it(0), it(1)
        Next it
        If filterIndex <> -1 Then fDialog.filterIndex = filterIndex
    End If
    If buttonName <> vbNullString Then fDialog.buttonName = buttonName
    If initialFileName <> vbNullString Then fDialog.initialFileName = initialFileName
    If title <> vbNullString Then fDialog.title = title
    'Show
    If fDialog.Show = -1 Then
        SaveAsDialog = fDialog.SelectedItems(1)
    End If
End Function

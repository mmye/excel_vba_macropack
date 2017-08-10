Attribute VB_Name = "BulkBookUtil"
Option Explicit

Public Sub SaveAllBooks()
Dim wkb As Workbook

Application.DisplayAlerts = False
For Each wkb In Workbooks
    wkb.Save
Next wkb

Application.DisplayAlerts = True
End Sub

Public Sub CloseAllBooks()
Dim wkb As Workbook

Application.DisplayAlerts = False
For Each wkb In Workbooks
    wkb.Close
Next wkb

Application.DisplayAlerts = True
End Sub

Public Function OpenAllBooks(Files, Path)
Dim f
For Each f In Files
    Debug.Print f
    Dim FilePath
    FilePath = Path & f
    Workbooks.Open (FilePath)
Next f

End Function


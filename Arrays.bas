Attribute VB_Name = "Arrays"
Function ArrayLength(arr, Optional rank As Long = 1)
    ArrayLength = UBound(arr, rank) - LBound(arr, rank) + 1
End Function
Function CreateBooleanArray(delimitedValues As String, Optional delimiter As String = ";") As Boolean()
'Define an array using a delimited string of values
    Dim vals() As String, Val, Index As Long, arr() As Boolean
    vals = Split(delimitedValues, delimiter)
    ReDim arr(ArrayLength(vals) - 1) As Boolean
    Index = LBound(arr)
    For Each Val In vals
        arr(Index) = CBool(Val)
        Index = Index + 1
    Next Val
    CreateBooleanArray = arr
End Function
Function CreateByteArray(delimitedValues As String, Optional delimiter As String = ";") As Byte()
'Define an array using a delimited string of values
    Dim vals() As String, Val, Index As Long, arr() As Byte
    vals = Split(delimitedValues, delimiter)
    ReDim arr(ArrayLength(vals) - 1) As Byte
    Index = LBound(arr)
    For Each Val In vals
        arr(Index) = CByte(Val)
        Index = Index + 1
    Next Val
    CreateByteArray = arr
End Function
Function CreateDateArray(delimitedValues As String, Optional delimiter As String = ";") As Date()
'Define an array using a delimited string of values
    Dim vals() As String, Val, Index As Long, arr() As Date
    vals = Split(delimitedValues, delimiter)
    ReDim arr(ArrayLength(vals) - 1) As Date
    Index = LBound(arr)
    For Each Val In vals
        arr(Index) = CDate(Val)
        Index = Index + 1
    Next Val
    CreateDateArray = arr
End Function
Function CreateSingleArray(delimitedValues As String, Optional delimiter As String = ";") As Single()
'Define an array using a delimited string of values
    Dim vals() As String, Val, Index As Long, arr() As Single
    vals = Split(delimitedValues, delimiter)
    ReDim arr(ArrayLength(vals) - 1) As Single
    Index = LBound(arr)
    For Each Val In vals
        arr(Index) = CDate(Val)
        Index = Index + 1
    Next Val
    CreateSingleArray = arr
End Function
Function CreateDoubleArray(delimitedValues As String, Optional delimiter As String = ";") As Double()
'Define an array using a delimited string of values
    Dim vals() As String, Val, Index As Long, arr() As Double
    vals = Split(delimitedValues, delimiter)
    ReDim arr(ArrayLength(vals) - 1) As Double
    Index = LBound(arr)
    For Each Val In vals
        arr(Index) = CDbl(Val)
        Index = Index + 1
    Next Val
    CreateDoubleArray = arr
End Function
Function CreateIntegerArray(delimitedValues As String, Optional delimiter As String = ";") As Integer()
'Define an array using a delimited string of values
    Dim vals() As String, Val, Index As Long, arr() As Integer
    vals = Split(delimitedValues, delimiter)
    ReDim arr(ArrayLength(vals) - 1) As Integer
    Index = LBound(arr)
    For Each Val In vals
        arr(Index) = CInt(Val)
        Index = Index + 1
    Next Val
    CreateIntegerArray = arr
End Function
Function CreateLongArray(delimitedValues As String, Optional delimiter As String = ";") As Long()
'Define an array using a delimited string of values
    Dim vals() As String, Val, Index As Long, arr() As Long
    vals = Split(delimitedValues, delimiter)
    ReDim arr(ArrayLength(vals) - 1) As Long
    Index = LBound(arr)
    For Each Val In vals
        arr(Index) = CLng(Val)
        Index = Index + 1
    Next Val
    CreateLongArray = arr
End Function
Function CreateStringArray(delimitedValues As String, Optional delimiter As String = ";") As String()
'Define an array using a delimited string of values
    CreateStringArray = Split(delimitedValues, delimiter)
End Function
Function MergeArrays(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
'Merge 2 1dimensional arrays
        Dim tmpArr As Variant, upper1 As Long, upper2 As Long
        Dim higherUpper As Long, i As Long, newIndex As Long
        upper1 = UBound(arr1) + 1: upper2 = UBound(arr2) + 1
        higherUpper = IIf(upper1 >= upper2, upper1, upper2)
        ReDim tmpArr(upper1 + upper2 - 1)
        For i = 0 To higherUpper
            If i < upper1 Then
                tmpArr(newIndex) = arr1(i)
                newIndex = newIndex + 1
            End If
        Next i
        
        For i = 0 To higherUpper
            If i < upper2 Then
                tmpArr(newIndex) = arr2(i)
                newIndex = newIndex + 1
            End If
        Next i
        MergeArrays = tmpArr
End Function
Function CompareArrays(ByVal arr1 As Variant, ByVal arr2 As Variant) As Boolean
'Compare 2 1dimensional arrays. Returns true if all items are identical
    Dim i As Long
    For i = LBound(arr1) To UBound(arr1)
       If arr1(i) <> arr2(i) Then
          CompareArrays = False
          Exit Function
       End If
    Next i
    CompareArrays = True
End Function
Function ResizeBooleanArray(ByRef arr() As Boolean, newSize As Long, Optional preserveArray As Boolean = True)
    If preserveArray Then
        ReDim Preserve arr(newSize - 1) As Boolean
    Else
        ReDim arr(newSize - 1) As Boolean
    End If
End Function
Function ResizeByteArray(ByRef arr() As Byte, newSize As Long, Optional preserveArray As Boolean = True)
    If preserveArray Then
        ReDim Preserve arr(newSize - 1) As Byte
    Else
        ReDim arr(newSize - 1) As Byte
    End If
End Function
Function ResizeDateArray(ByRef arr() As Date, newSize As Long, Optional preserveArray As Boolean = True)
    If preserveArray Then
        ReDim Preserve arr(newSize - 1) As Date
    Else
        ReDim arr(newSize - 1) As Date
    End If
End Function
Function ResizeSingleArray(ByRef arr() As Single, newSize As Long, Optional preserveArray As Boolean = True)
    If preserveArray Then
        ReDim Preserve arr(newSize - 1) As Single
    Else
        ReDim arr(newSize - 1) As Single
    End If
End Function
Function ResizeDoubleArray(ByRef arr() As Double, newSize As Long, Optional preserveArray As Boolean = True)
    If preserveArray Then
        ReDim Preserve arr(newSize - 1) As Double
    Else
        ReDim arr(newSize - 1) As Double
    End If
End Function
Function ResizeIntegerArray(ByRef arr() As Integer, newSize As Long, Optional preserveArray As Boolean = True)
    If preserveArray Then
        ReDim Preserve arr(newSize - 1) As Integer
    Else
        ReDim arr(newSize - 1) As Integer
    End If
End Function
Function ResizeLongArray(ByRef arr() As Long, newSize As Long, Optional preserveArray As Boolean = True)
    If preserveArray Then
        ReDim Preserve arr(newSize - 1) As Long
    Else
        ReDim arr(newSize - 1) As Long
    End If
End Function
Function ResizeStringArray(ByRef arr() As String, newSize As Long, Optional preserveArray As Boolean = True)
    If preserveArray Then
        ReDim Preserve arr(newSize - 1) As String
    Else
        ReDim arr(newSize - 1) As String
    End If
End Function

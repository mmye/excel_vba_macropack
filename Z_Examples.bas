Attribute VB_Name = "Z_Examples"
'*******ProgressBar*******
Sub ProgressBarExamples()
    Dim prog As ProgressBar
    Set prog = New ProgressBar
    Call prog.Initialize("My Progress", 100)
    For i = 0 To 99
        prog.AddProgress
        Application.Wait Now + TimeValue("00:00:01")
    Next i
End Sub
'*******Arrays*******
Sub ArraysExample()
    Dim arr1() As Long, arr2() As Long, arr3 As Variant
    arr1 = Arrays.CreateLongArray("1;2;3;4")
    arr2 = Arrays.CreateLongArray("1,5,3,4", ",")
    Debug.Print "Array arr1 is long: " & Arrays.ArrayLength(arr1)
    Debug.Print "Are arrays identical?: " & Arrays.CompareArrays(arr1, arr2)
    arr2(1) = 2
    Debug.Print "Are arrays identical now?: " & Arrays.CompareArrays(arr1, arr2)
    arr3 = Arrays.MergeArrays(arr1, arr2)
    Debug.Print "Merged arrays arr1 and arr2. Arr3 length is: " & Arrays.ArrayLength(arr3)
    Call Arrays.ResizeLongArray(arr1, 10, False)
    Debug.Print "Resized arr1 to 10 items. Arr1 length is: " & Arrays.ArrayLength(arr1)
End Sub
'*******Files*******
Sub FilesExamples()
    Dim fileName As String: fileName = "testFile.txt"
    Call Files.CreateTextFile(fileName, True)
    Debug.Print "Created file " & fileName
    Debug.Print "File exists?: " & Files.PathExists(fileName)
    Files.DeleteFile (fileName)
    Debug.Print "File deleted?: " & Not (Files.PathExists(fileName))
    'Merging text files - requires actual files and folders
    Dim fileNames(0 To 1) As String
    fileNames(0) = "C:\somefolder\test.csv"
    fileNames(1) = "C:\somefolder\test1.csv"
         
    MergeFiles fileNames, "C:\Merged.csv", True, False
    MergeFilesInFolder "C:\somefolder\*.csv", "C:\MergedFolder.csv", True, False
    MergeFilesInSubFolders "C:\somefolder\", "C:\MergedSubFolders.csv", True, False
End Sub
'*******Performance*******
Sub PerformanceExamples()
    'Optimizing...
    Performance.OptimizeOn
    Dim i As Long, x As Long
    Performance.StartLowResPerformanceTimer
    For i = 0 To 100000000
        x = i Mod 9999
    Next i
    Debug.Print "Code took (Low res): " & Performance.StopLowResPerformanceTimer & " seconds to execute"
    Performance.StartHighResPerformanceTimer
    For i = 0 To 100000000
        x = i Mod 9999
    Next i
    Debug.Print "Code took (High res): " & Performance.StopHighResPerformanceTimer & " milliseconds to execute"
    
    Debug.Print "Creating thread..."
    Dim threadId As Long
    threadId = Performance.RunThread(AddressOf PerformanceThreadExample)
    Application.Wait Now + TimeValue("00:00:01")
    Performance.KillThread threadId
    'Finished VBA execution
    Performance.OptimizeOff
End Sub
Sub PerformanceThreadExample()
    Debug.Print "I am a new thread!"
End Sub
'*******RegularExpressions*******
Sub RegularExpressionsExamples()
    Dim str As String: str = "122 322 344"
    Debug.Print "How many 3 digit sequences in " & str & "?: " & RegularExpressions.RegexCountMatches(str, "[0-9]{3}")
    Debug.Print "First match using Get: " & RegularExpressions.RegexExecuteGet("122 322 344", "([0-9]{3})", 0, 0)
    Debug.Print "Second match using Execute: " & RegularExpressions.RegexExecute("122 322 344", "([0-9]{3})", False)(1).SubMatches(0)
    Debug.Print "Before replace: [" & str & "] After replace:[" & RegularExpressions.RegexReplace(str, "\d+", "replace") & "]"
End Sub
'*******Strings*******
Sub StringsExamples()
    Debug.Print "Month in English: " & Strings.LocaleFormat("2015/10", "mmm", English_United_States)
    Debug.Print "Month in French: " & Strings.LocaleFormat("2015/10", "mmm", French_France)
    Debug.Print "DOG to lower case: " & Strings.ToLowerCase("DOG")
    Debug.Print "Split, second item: " & Strings.SplitString("Dog;Cat", ";")(1)
    Dim strArray(0 To 1) As String: strArray(0) = "Dog": strArray(1) = "Cat"
    Debug.Print "Join strings: " & Strings.JoinStrings(strArray, " ")
    Debug.Print "Contains word 'dog'?: " & Strings.Contains("catdoghorse", "dog")
    Debug.Print "Replace string (cats with dogs) in 'cats are mans best fried': " & Strings.ReplaceString("cats are mans best friend", "cats", "dogs")
End Sub
'*******Structures*******
Sub StructuresExamples()
    Dim dict As Object, Items, Keys
    Set dict = Structures.CreateDictionary
    dict.Add "SomeKey1", "SomeValue1"
    dict.Add "SomeKey2", "SomeValue2"
    Items = dict.Items
    Keys = dict.Keys
    Debug.Print "Dict key1: " & Keys(0)
    Debug.Print "Dict item1: " & Items(0)
End Sub
'*******Timers*******
Sub TimersExamples()
    Timers.WaitForSeconds 1
    Debug.Print "Waited for 1 sec using WaitForSeconds"
    Timers.WaitUntilTime Now + TimeValue("00:00:01")
    Debug.Print "Waited for 1 sec using WaitUntilTIme"
    Timers.Sleep 500
    Debug.Print "Slept for 0.5 sec using Sleep"
    Timers.RunMacroInSeconds 1, "HelloTimer"
End Sub
Sub HelloTimer()
    Debug.Print "Hello from RunMacroInSeconds!"
End Sub
'*******SQL*********
Sub SQLExample()
    '`Example Table` is a hidden worksheet in the Time Saver Workbook
    
    'Get RecordSet
    Set rs = GetRecords("SELECT * FROM [Example Table$]")
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        Do Until rs.EOF = True
            Debug.Print rs![Column 1] & " " & rs![Column 2] & " " & rs![Column 3]
            rs.MoveNext
        Loop
    End If
    
    'Get Array of Variants
    varArr = GetRecordsArray("SELECT * FROM [Example Table$]")
    For i = 0 To Arrays.ArrayLength(varArr, 1) - 1
        For j = 0 To Arrays.ArrayLength(varArr, 2) - 1
            Debug.Print "Row: " & i & ", Column: " & j & ", Value: " & varArr(i, j)
        Next j
    Next i
End Sub
'*******Translations*********
Sub TranslateExample()
    Debug.Print Translations.Translate("After finishing the dish, she immediately started to eat another one.", Translations.English, Translations.Japanese)
End Sub

Sub TranslateExampleWithQuotationJAtoEN()
    Dim r As Range
    For Each r In ActiveSheet.UsedRange
        On Error Resume Next
        If r.Value <> "" Then r.Value = Translations.Translate(r.Value, Translations.Japanese, Translations.English)
    Next
    On Error GoTo 0
End Sub
Sub TranslateExampleWithQuotationENtoJA()
    Dim r As Range
    For Each r In ActiveSheet.UsedRange
        On Error Resume Next
        If r.Value <> "" Then r.Value = Translations.Translate(r.Value, Translations.English, Translations.Japanese)
    Next
    On Error GoTo 0
End Sub

'*******Dialogs**************
Sub DialogsExamples()
    Dim filters As Collection, res As String, resIt As FileDialogSelectedItems
    Set filters = New Collection
    filters.Add Array("All", "*.*")
    filters.Add Array("Excel", "*.xlsx;*.xls")
    
    'Single file select
    res = Dialogs.SelectSingleFileDialog("Select a single file", filters:=filters)
    If res = vbNullString Then
        MsgBox "You have not selected any file"
    Else
        MsgBox "You have selected """ & res & """"
    End If
    
    'Multi file select
    Set resIt = Dialogs.SelectMultiFileDialog("Select multiple files", filters:=filters)
    If resIt Is Nothing Then
        MsgBox "You have not selected any file"
    Else
        res = "You have selected" & vbNewLine
        For Each it In resIt
            res = res & it & vbNewLine
        Next it
        MsgBox res
    End If
    
    'Open file dialog
    res = Dialogs.OpenFileDialog("Select a file to open")
    If res = vbNullString Then
        MsgBox "You have not selected any file"
    Else
        MsgBox "You have selected """ & res & """"
    End If
    
    'Select a folder dialog
    Debug.Print Dialogs.SelectFolderDialog("Select a folder to open")
    If res = vbNullString Then
        MsgBox "You have not selected any file"
    Else
        MsgBox "You have selected """ & res & """"
    End If
    
    'Save as dialog
    Debug.Print Dialogs.SaveAsDialog("Save as")
    If res = vbNullString Then
        MsgBox "You have not selected any file"
    Else
        MsgBox "You have selected """ & res & """"
    End If
End Sub
'*******Validate**************
Sub ValidateExamples()
    Debug.Print Validate.IsEmail("someemail@some.com")
    Debug.Print Validate.IsDomainName("analystcave.com")
    Debug.Print Validate.IsURL("http://analystcave.com")
    Debug.Print Validate.IsSSN("333-22-4444")
    Debug.Print Validate.IsCreditCardNumber("1234-1234-1234-1234")
    Debug.Print Validate.IsUSZIP("34545-2367")
    Debug.Print Validate.IsInternationalPhoneNumber("+123(44)123-456-123 ")
    Debug.Print Validate.IsUSPhoneNumber("(572)8841234")
End Sub









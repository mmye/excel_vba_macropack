Attribute VB_Name = "Files"
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
#If VBA7 Then
    Private Declare PtrSafe Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function InternetReadBinaryFile Lib "wininet.dll" Alias "InternetReadFile" (ByVal hfile As Long, ByRef bytearray_firstelement As Byte, ByVal lNumBytesToRead As Long, ByRef lNumberOfBytesRead As Long) As Integer
    Private Declare PtrSafe Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
    Private Declare PtrSafe Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
#Else
    Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
    Private Declare Function InternetReadBinaryFile Lib "wininet.dll" Alias "InternetReadFile" (ByVal hfile As Long, ByRef bytearray_firstelement As Byte, ByVal lNumBytesToRead As Long, ByRef lNumberOfBytesRead As Long) As Integer
    Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
    Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
#End If

Function ReadTextFileEOF(fileName As String) As String
'Read an entire text file
    Dim textData As String, fileNo As Integer
    fileNo = FreeFile
    Open fileName For Input As #fileNo
    textData = Input$(LOF(fileNo), fileNo)
    Close #fileNo
    ReadTextFileEOF = textData
End Function
Function ReadXML(fileName As String) As Object
'Read an XML to DOMDocument object
    Dim XDoc As Object:  Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    XDoc.Load fileName
    Set ReadXML = XDoc
End Function
Function ReadCSVRecordSet(fileName As String, Optional includesColumnNames As Boolean = True) As Object
'Read entire CSV file to RecordSet
'Uses default system delimiter
    Dim pathName As String
    pathName = Left(fileName, InStrRev(fileName, "\"))
    fileName = Right(fileName, Len(fileName) - InStrRev(fileName, "\"))
    Set rs = CreateObject("ADODB.Recordset")
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pathName & ";" _
    & "Extended Properties=""text;HDR=" & IIf(includesColumnNames, "Yes", "No") & ";FMT=Delimited"";"
    strSQL = "SELECT * FROM " & fileName
    rs.Open strSQL, strCon, 3, 3
    Set ReadCSVRecordSet = rs
End Function
Sub CreateTextFile(fileName As String, Optional overwrite As Boolean = True)
'Create text file
    Set fs = CreateObject("Scripting.FileSystemObject")
    fs.CreateTextFile fileName, overwrite
    Set fs = Nothing
End Sub
Sub CreateDirectory(directoryName As String)
    MkDir directoryName
End Sub
Sub AppendTextFile(fileName As String, textData As String)
'Append text to existing text file
    Dim fileNo As Integer
    fileNo = FreeFile
    Open fileName For Output As #fileNo
    Print #fileNo, textData
    Close #fileNo
End Sub
Function PathExists(Path As String) As Boolean
'Check if path/files exists
    Dim fileName As String
    fileName = Dir(Path)
    If fileName <> vbNullString Then
        PathExists = True
    Else
        PathExists = False
    End If
End Function
Sub DeleteFile(fileName As String)
'Deletes any file.
    If Len(Dir(fileName)) > 0 Then
        SetAttr fileName, vbNormal
        Kill fileName
    End If
End Sub
Sub DownloadFile(sUrl As String, FilePath As String, Optional overWriteFile As Boolean)
'Download file from provided URL. Provides updates in status bar
  Dim hInternet, hSession, lngDataReturned As Long, sBuffer() As Byte, totalRead As Long
  Const bufSize = 128
  ReDim sBuffer(bufSize)
  hSession = InternetOpen("", 0, vbNullString, vbNullString, 0)
  If hSession Then hInternet = InternetOpenUrl(hSession, sUrl, vbNullString, 0, INTERNET_FLAG_NO_CACHE_WRITE, 0)
  Set oStream = CreateObject("ADODB.Stream")
  oStream.Open
  oStream.Type = 1
  If hInternet Then
    iReadFileResult = InternetReadBinaryFile(hInternet, sBuffer(0), UBound(sBuffer) - LBound(sBuffer), lngDataReturned)
    ReDim Preserve sBuffer(lngDataReturned - 1)
    oStream.Write sBuffer
    ReDim sBuffer(bufSize)
    totalRead = totalRead + lngDataReturned
    Application.StatusBar = "Downloading file. " & CLng(totalRead / 1024) & " KB downloaded"
    DoEvents
    Do While lngDataReturned <> 0
      iReadFileResult = InternetReadBinaryFile(hInternet, sBuffer(0), UBound(sBuffer) - LBound(sBuffer), lngDataReturned)
      If lngDataReturned = 0 Then Exit Do
      ReDim Preserve sBuffer(lngDataReturned - 1)
      oStream.Write sBuffer
      ReDim sBuffer(bufSize)
      totalRead = totalRead + lngDataReturned
      Application.StatusBar = "Downloading file. " & CLng(totalRead / 1024) & " KB downloaded"
      DoEvents
    Loop
    Application.StatusBar = "Download complete"
    oStream.SaveToFile FilePath, IIf(overWriteFile, 2, 1)
    oStream.Close
  End If
  Call InternetCloseHandle(hInternet)
End Sub
Sub MergeFiles(fileNames() As String, newFileName As String, Optional headers As Boolean = True, Optional addNewLine As Boolean = False)
'Merge list of text files
'fileNames - array of file paths
'newFileName - the name of the new consolidated file
'headers - if true assuming first row contains headers
'addNewLine - if true a vbNewLine character will be added before each next file
    Dim fileName As Variant, textData As String, fileNo As Integer, result As String, firstHeader As Boolean
    firstHeader = True
    For Each fileName In fileNames
        fileNo = FreeFile
        Open fileName For Input As #fileNo
        textData = Input$(LOF(fileNo), fileNo)
        Close #fileNo
        If headers Then
            result = result & IIf(addNewLine, vbNewLine, "") & IIf(firstHeader, textData, Right(textData, Len(textData) - InStr(textData, vbNewLine)))
            firstHeader = False
        Else
            result = result & IIf(addNewLine, vbNewLine, "") & textData
        End If
    Next fileName
    fileNo = FreeFile
    Open newFileName For Output As #fileNo
    Print #fileNo, result
    Close #fileNo
End Sub
Sub MergeFilesInFolder(FolderName As String, newFileName As String, Optional headers As Boolean = True, Optional addNewLine As Boolean = False)
'Merge text files within folder
'folderName - folder which is to be searched for files (accepts Dir wildcards)
'newFileName - the name of the new consolidated file
'headers - if true assuming first row contains headers
'addNewLine - if true a vbNewLine character will be added before each next file
    Dim fileName As Variant, textData As String, fileNo As Integer, result As String, firstHeader As Boolean
    firstHeader = True
    fileName = Dir(FolderName)
    Do Until fileName = ""
        fileNo = FreeFile
        Open (Left(FolderName, InStrRev(FolderName, "\")) & fileName) For Input As #fileNo
        textData = Input$(LOF(fileNo), fileNo)
        Close #fileNo
        If headers Then
            result = result & IIf(addNewLine, vbNewLine, "") & IIf(firstHeader, textData, Right(textData, Len(textData) - InStr(textData, vbNewLine)))
            firstHeader = False
        Else
            result = result & IIf(addNewLine, vbNewLine, "") & textData
        End If
        fileName = Dir
    Loop
    fileNo = FreeFile
    Open newFileName For Output As #fileNo
    Print #fileNo, result
    Close #fileNo
End Sub
Sub MergeFilesInSubFolders(FolderName As String, pattern As String, newFileName As String, Optional headers As Boolean = True, Optional addNewLine As Boolean = False)
'Merge files in folder and subfolders
'folderName - folder to be traversed for files (does not accept wildcards)
'pattern - file patten using Dir wildcards
'newFileName - the name of the new consolidated file
'headers - if true assuming first row contains headers
'addNewLine - if true a vbNewLine character will be added before each next file
    Dim fileName As Variant, folder As Variant, textData As String, fileNo As Integer, result As String, firstHeader As Boolean
    firstHeader = True
    Dim col As Collection
    Set col = New Collection
    col.Add FolderName
    TraversePath FolderName, col
    For Each folder In col
        fileName = Dir(folder & pattern)
        Do Until fileName = ""
            fileNo = FreeFile
            Open (Left(folder, InStrRev(folder, "\")) & fileName) For Input As #fileNo
            textData = Input$(LOF(fileNo), fileNo)
            Close #fileNo
            If headers Then
                result = result & IIf(addNewLine, vbNewLine, "") & IIf(firstHeader, textData, Right(textData, Len(textData) - InStr(textData, vbNewLine)))
                firstHeader = False
            Else
                result = result & IIf(addNewLine, vbNewLine, "") & textData
            End If
            fileName = Dir
        Loop
    Next folder
    fileNo = FreeFile
    Open newFileName For Output As #fileNo
    Print #fileNo, result
    Close #fileNo
End Sub
Function TraversePath(Path As Variant, allDirCollection As Collection)
'Returns a collection of all directories and subdirectories within the provided path
'path - path to be traversed
'allDirCollection - a Collection object (see Structures module) in which all directories will be stored as strings
    Dim currentPath As String, directory As Variant
    Dim dirCollection As Collection
    Set dirCollection = New Collection
      
    currentPath = Dir(Path, vbDirectory)
    'Explore current directory
    Do Until currentPath = vbNullString
        If Left(currentPath, 1) <> "." And Left(currentPath, 2) <> ".." And _
            (GetAttr(Path & currentPath) And vbDirectory) = vbDirectory Then
            dirCollection.Add Path & currentPath & "\"
            allDirCollection.Add Path & currentPath & "\"
        End If
        currentPath = Dir()
    Loop
      
    'Explore subsequent directories
    For Each directory In dirCollection
        TraversePath directory, allDirCollection
    Next directory
End Function

Function CountFileNumber(Path As String)
'return count of files in a specified directory

    Dim i As Long, FSO As Object, f As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    CountFileNumber = FSO.GetFolder(Path).Files.Count
End Function

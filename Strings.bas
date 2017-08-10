Attribute VB_Name = "Strings"
Function ToUpperCase(str As String)
'Transform all letters to UPPERCASE
    ToUpperCase = StrConv(str, vbUpperCase)
End Function
Function ToLowerCase(str As String)
'Transform all letters to lowercase
    ToLowerCase = StrConv(str, vbLowerCase)
End Function
Function ToProperCase(str As String)
'Transform all letters to Propercase
    ToProperCase = StrConv(str, vbProperCase)
End Function
Function LocaleFormat(str As String, strFormat As String, Optional language As Locale)
'Format a string using a specific locale
    LocaleFormat = WorksheetFunction.Text(Date, IIf(IsMissing(language), "", "[$-" & Hex(language) & "]") & strFormat)
End Function
Function SubString(str As String, start As Long, Optional length As Variant) As String
'Get Substring, starting at "start" index. Optionally specify length of substring
    If IsMissing(length) Then
        SubString = Mid(str, start)
        Exit Function
    Else
        SubString = Mid(str, start, length)
    End If
End Function
Function SplitString(str As String, Optional delimiter As String = " ", Optional limit As Long = -1, Optional compareMethod As VbCompareMethod = VbCompareMethod.vbBinaryCompare) As String()
'Split string using default or specified delimiter. Optionally limit the amount of resulting substrings
    SplitString = Split(str, delimiter, limit, compareMethod)
End Function
Function JoinStrings(strArray() As String, Optional delimiter As String = "") As String
'Join strings within an array of strings. Optionally include a specified delimiter between each string
    JoinStrings = Join(strArray, delimiter)
End Function
Function RepeatString(str As String, repeatTimes As Long) As String
'Repeat and append string repeatTimes number of times
    If Len(str) = 1 Then
        RepeatString = String(repeatTimes, str)
    Else
        Dim i As Long: For i = 1 To repeatTimes: RepeatString = RepeatString & str: Next i
    End If
End Function
Function ReplaceString(str As String, findStr As String, replaceStr As String, Optional replaceLimit As Long = -1, Optional ignoreCase As Boolean = False, Optional startAt As Long = 1) As String
'Replace strings findStr in string str with strings replaceStr
'#####Use RegularExpressions.RegexReplace to replace using Regex#####
'replaceLimit - by default unlimited (-1). Providing value will limit the number of performed replacements
'ignoreCase - ignore letter case
'startAt - start looking for string starting at index. First character starts at 1
    ReplaceString = Replace(str, findStr, replaceStr, startAt, replaceLimit, IIf(ignoreCase, vbTextCompare, vbBinaryCompare))
End Function
Function Contains(str As String, containsStr As String, Optional ignoreCase As Boolean = False, Optional startAt As Long = 1) As Boolean
'Returns true if str string contains containsStr string
'ignoreCase - ignore letter case
'startAt - start looking for string starting at index. First character starts at 1
    Contains = InStr(startAt, str, containsStr, IIf(ignoreCase, vbTextCompare, vbBinaryCompare)) > 0
End Function

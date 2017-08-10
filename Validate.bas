Attribute VB_Name = "Validate"
'Default VBA Validation functions:
'IsArray
'IsDate
'HasData
'IsError
'IsMissing
'IsNull
'IsNumeric
'IsObject
'---Additional Common Validation functions---
'***************INTERNET**************
Function IsEmail(str As String) As Boolean
'str - email to validate
    IsEmail = RegularExpressions.RegexTest(str, "(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|""(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*"")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])")
End Function
Function IsDomainName(str As String) As Boolean
'str - domain to validate
    IsDomainName = RegularExpressions.RegexTest(str, "^[a-zA-Z0-9\-\.]+\.(com|org|net|mil|edu|COM|ORG|NET|MIL|EDU)$")
End Function
Function IsIPAddress(str As String) As Boolean
'str - ip to validate
    IsIPAddress = RegularExpressions.RegexTest(str, "^(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9])\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[1-9]|0)\.(25[0-5]|2[0-4][0-9]|[0-1]{1}[0-9]{2}|[1-9]{1}[0-9]{1}|[0-9])$")
End Function
Function IsURL(str As String) As Boolean
'str - url to validate
    IsURL = RegularExpressions.RegexTest(str, "(http|ftp|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?")
End Function
'***************NUMBERS VALIDATION************
Function IsSSN(str As String) As Boolean
'str - Social Security Number (hyphen-separated) to validate
    IsSSN = RegularExpressions.RegexTest(str, "^\d{3}-\d{2}-\d{4}$")
End Function
Function IsCreditCardNumber(str As String) As Boolean
'str - Credit Card (16 digit) to validate
    IsCreditCardNumber = RegularExpressions.RegexTest(str, "^(\d{4}[- ]){3}\d{4}|\d{16}$")
End Function
Function IsUSZIP(str As String) As Boolean
'str - US ZIP (5 or 5+4 digits) to validate
    IsUSZIP = RegularExpressions.RegexTest(str, "^\d{5}$|^\d{5}-\d{4}$")
End Function
Function IsInternationalPhoneNumber(str As String) As Boolean
'str - phone number to validate
    Dim valStr As String
    valStr = Replace(Replace(Replace(str, " ", ""), "-", ""), ".", "")
    IsInternationalPhoneNumber = RegularExpressions.RegexTest(str, "^(\+[1-9][0-9]*(\([0-9]*\)|-[0-9]*-))?[0]?[1-9][0-9\- ]*")
End Function
Function IsUSPhoneNumber(str As String) As Boolean
'str - phone number to validate
    IsUSPhoneNumber = RegularExpressions.RegexTest(str, "^(?:(?:\+?1\s*(?:[.-]\s*)?)?(?:\(\s*([2-9]1[02-9]|[2-9][02-8]1|[2-9][02-8][02-9])\s*\)|([2-9]1[02-9]|[2-9][02-8]1|[2-9][02-8][02-9]))\s*(?:[.-]\s*)?)?([2-9]1[02-9]|[2-9][02-9]1|[2-9][02-9]{2})\s*(?:[.-]\s*)?([0-9]{4})(?:\s*(?:#|x\.?|ext\.?|extension)\s*(\d+))?$")
End Function


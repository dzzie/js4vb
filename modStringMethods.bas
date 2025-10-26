Attribute VB_Name = "modStringMethods"
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT

Option Explicit

'------------------------------------------------------------
' Professional implementation using native VB6 functions
'------------------------------------------------------------

' Check if a name is a valid string method
Public Function IsStringMethod(methodName As String) As Boolean
    Select Case methodName
        Case "charAt", "charCodeAt", "indexOf", "lastIndexOf", _
             "substring", "substr", "slice", "toLowerCase", "toUpperCase", _
             "trim", "split", "replace", "replaceAll", _
             "startsWith", "endsWith", "includes", "repeat", _
             "padStart", "padEnd", "concat"
            IsStringMethod = True
        Case Else:
            IsStringMethod = False
    End Select
End Function

' Execute a string method and return result as CValue
Public Function CallStringMethod(strValue As String, methodName As String, args As Collection) As CValue
    Dim result As New CValue
    
    Select Case methodName
        Case "charAt":
            result.vType = vtString
            result.strVal = str_charAt(strValue, GetArg(args, 1, 0))
            
        Case "charCodeAt":
            result.vType = vtNumber
            result.numVal = str_charCodeAt(strValue, GetArg(args, 1, 0))
            
        Case "indexOf":
            result.vType = vtNumber
            result.numVal = str_indexOf(strValue, GetArgStr(args, 1, ""), GetArg(args, 2, 0))
            
        Case "lastIndexOf":
            result.vType = vtNumber
            result.numVal = str_lastIndexOf(strValue, GetArgStr(args, 1, ""), GetArg(args, 2, -1))
            
        Case "substring":
            result.vType = vtString
            result.strVal = str_substring(strValue, GetArg(args, 1, 0), GetArg(args, 2, Len(strValue)))
            
        Case "substr":
            result.vType = vtString
            result.strVal = str_substr(strValue, GetArg(args, 1, 0), GetArg(args, 2, Len(strValue)))
            
        Case "slice":
            result.vType = vtString
            result.strVal = str_slice(strValue, GetArg(args, 1, 0), GetArg(args, 2, Len(strValue)))
            
        Case "toLowerCase":
            result.vType = vtString
            result.strVal = LCase$(strValue)
            
        Case "toUpperCase":
            result.vType = vtString
            result.strVal = UCase$(strValue)
            
        Case "trim":
            result.vType = vtString
            result.strVal = Trim$(strValue)
            
        Case "split":
            result.vType = vtArray
            Set result.arrayVal = str_split(strValue, GetArgStr(args, 1, ""), GetArg(args, 2, -1))
            
        Case "replace":
            result.vType = vtString
            result.strVal = str_replace(strValue, GetArgStr(args, 1, ""), GetArgStr(args, 2, ""))
            
        Case "replaceAll":
            result.vType = vtString
            result.strVal = Replace(strValue, GetArgStr(args, 1, ""), GetArgStr(args, 2, ""))
            
        Case "startsWith":
            result.vType = vtBoolean
            result.boolVal = str_startsWith(strValue, GetArgStr(args, 1, ""), GetArg(args, 2, 0))
            
        Case "endsWith":
            result.vType = vtBoolean
            result.boolVal = str_endsWith(strValue, GetArgStr(args, 1, ""), GetArg(args, 2, Len(strValue)))
            
        Case "includes":
            result.vType = vtBoolean
            result.boolVal = str_includes(strValue, GetArgStr(args, 1, ""), GetArg(args, 2, 0))
            
        Case "repeat":
            result.vType = vtString
            result.strVal = str_repeat(strValue, GetArg(args, 1, 0))
            
        Case "padStart":
            result.vType = vtString
            result.strVal = str_padStart(strValue, GetArg(args, 1, 0), GetArgStr(args, 2, " "))
            
        Case "padEnd":
            result.vType = vtString
            result.strVal = str_padEnd(strValue, GetArg(args, 1, 0), GetArgStr(args, 2, " "))
            
        Case "concat":
            result.vType = vtString
            result.strVal = str_concat(strValue, args)
            
        Case Else:
            result.vType = vtUndefined
    End Select
    
    Set CallStringMethod = result
End Function

' Helper: Get numeric argument
Private Function GetArg(args As Collection, index As Long, defaultVal As Long) As Long
    If args.count >= index Then
        Dim val As CValue
        Set val = args(index)
        GetArg = CLng(val.ToNumber())
    Else
        GetArg = defaultVal
    End If
End Function

' Helper: Get string argument
Private Function GetArgStr(args As Collection, index As Long, defaultVal As String) As String
    If args.count >= index Then
        Dim val As CValue
        Set val = args(index)
        GetArgStr = val.ToString()
    Else
        GetArgStr = defaultVal
    End If
End Function

' ============================================
' STRING METHOD IMPLEMENTATIONS
' ============================================

Private Function str_charAt(s As String, index As Long) As String
    If index >= 0 And index < Len(s) Then
        str_charAt = Mid$(s, index + 1, 1)
    Else
        str_charAt = ""
    End If
End Function

Private Function str_charCodeAt(s As String, index As Long) As Double
    If index >= 0 And index < Len(s) Then
        str_charCodeAt = Asc(Mid$(s, index + 1, 1))
    Else
        str_charCodeAt = 0  ' Should be NaN
    End If
End Function

Private Function str_indexOf(s As String, searchValue As String, fromIndex As Long) As Long
    Dim pos As Long
    pos = InStr(fromIndex + 1, s, searchValue)
    If pos > 0 Then
        str_indexOf = pos - 1
    Else
        str_indexOf = -1
    End If
End Function

Private Function str_lastIndexOf(s As String, searchValue As String, fromIndex As Long) As Long
    Dim pos As Long
    Dim searchStart As Long
    
    If fromIndex < 0 Then
        searchStart = Len(s)
    Else
        searchStart = fromIndex + 1
    End If
    
    pos = InStrRev(s, searchValue, searchStart)
    If pos > 0 Then
        str_lastIndexOf = pos - 1
    Else
        str_lastIndexOf = -1
    End If
End Function

Private Function str_substring(s As String, startIndex As Long, endIndex As Long) As String
    Dim actualStart As Long
    Dim actualEnd As Long
    Dim strLen As Long
    
    strLen = Len(s)
    
    ' Clamp indices
    actualStart = IIf(startIndex < 0, 0, IIf(startIndex > strLen, strLen, startIndex))
    actualEnd = IIf(endIndex < 0, 0, IIf(endIndex > strLen, strLen, endIndex))
    
    ' Swap if needed
    If actualStart > actualEnd Then
        Dim temp As Long
        temp = actualStart
        actualStart = actualEnd
        actualEnd = temp
    End If
    
    If actualStart >= actualEnd Then
        str_substring = ""
    Else
        str_substring = Mid$(s, actualStart + 1, actualEnd - actualStart)
    End If
End Function

Private Function str_substr(s As String, startIndex As Long, length As Long) As String
    Dim strLen As Long
    Dim actualStart As Long
    Dim actualLength As Long
    
    strLen = Len(s)
    
    ' Handle negative start
    If startIndex < 0 Then
        actualStart = strLen + startIndex
        If actualStart < 0 Then actualStart = 0
    Else
        actualStart = startIndex
    End If
    
    actualLength = IIf(length < 0, 0, length)
    
    If actualStart >= strLen Or actualLength = 0 Then
        str_substr = ""
    Else
        str_substr = Mid$(s, actualStart + 1, actualLength)
    End If
End Function

Private Function str_slice(s As String, startIndex As Long, endIndex As Long) As String
    Dim strLen As Long
    Dim actualStart As Long
    Dim actualEnd As Long
    
    strLen = Len(s)
    
    ' Handle negative start
    If startIndex < 0 Then
        actualStart = strLen + startIndex
        If actualStart < 0 Then actualStart = 0
    Else
        actualStart = IIf(startIndex > strLen, strLen, startIndex)
    End If
    
    ' Handle end
    If endIndex < 0 Then
        actualEnd = strLen + endIndex
        If actualEnd < 0 Then actualEnd = 0
    Else
        actualEnd = IIf(endIndex > strLen, strLen, endIndex)
    End If
    
    If actualStart >= actualEnd Then
        str_slice = ""
    Else
        str_slice = Mid$(s, actualStart + 1, actualEnd - actualStart)
    End If
End Function

Private Function str_split(s As String, separator As String, limit As Long) As Collection
    Dim result As New Collection
    Dim parts() As String
    Dim i As Long
    Dim maxParts As Long
    Dim charVal As CValue
    Dim partVal As CValue
    
    If separator = "" Then
        ' Split into individual characters
        For i = 1 To Len(s)
            Set charVal = New CValue
            charVal.vType = vtString
            charVal.strVal = Mid$(s, i, 1)
            result.add charVal
            If limit > 0 And result.count >= limit Then Exit For
        Next
    Else
        ' Split by separator
        parts = VBA.Split(s, separator)
        
        If limit < 0 Then
            maxParts = UBound(parts) + 1
        Else
            maxParts = IIf(limit > UBound(parts) + 1, UBound(parts) + 1, limit)
        End If
        
        For i = 0 To maxParts - 1
            Set partVal = New CValue
            partVal.vType = vtString
            partVal.strVal = parts(i)
            result.add partVal
        Next
    End If
    
    Set str_split = result
End Function

Private Function str_replace(s As String, searchValue As String, replaceValue As String) As String
    ' Replace only first occurrence
    Dim pos As Long
    pos = InStr(1, s, searchValue)
    
    If pos > 0 Then
        str_replace = Left$(s, pos - 1) & replaceValue & Mid$(s, pos + Len(searchValue))
    Else
        str_replace = s
    End If
End Function

Private Function str_startsWith(s As String, searchString As String, position As Long) As Boolean
    Dim checkStr As String
    checkStr = Mid$(s, position + 1, Len(searchString))
    str_startsWith = (checkStr = searchString)
End Function

Private Function str_endsWith(s As String, searchString As String, length As Long) As Boolean
    Dim checkLen As Long
    Dim checkStr As String
    
    checkLen = IIf(length > Len(s), Len(s), length)
    
    If checkLen < Len(searchString) Then
        str_endsWith = False
    Else
        checkStr = Mid$(s, checkLen - Len(searchString) + 1, Len(searchString))
        str_endsWith = (checkStr = searchString)
    End If
End Function

Private Function str_includes(s As String, searchString As String, position As Long) As Boolean
    str_includes = (InStr(position + 1, s, searchString) > 0)
End Function

Private Function str_repeat(s As String, count As Long) As String
    Dim i As Long
    Dim result As String
    
    If count < 0 Then
        Err.Raise vbObjectError + 2000, "modStringMethods", "repeat count must be non-negative"
    End If
    
    result = ""
    For i = 1 To count
        result = result & s
    Next
    
    str_repeat = result
End Function

Private Function str_padStart(s As String, targetLength As Long, padString As String) As String
    Dim currentLen As Long
    Dim padLen As Long
    Dim padding As String
    
    currentLen = Len(s)
    
    If currentLen >= targetLength Then
        str_padStart = s
        Exit Function
    End If
    
    padLen = targetLength - currentLen
    
    ' Build padding
    padding = ""
    Do While Len(padding) < padLen
        padding = padding & padString
    Loop
    
    str_padStart = Left$(padding, padLen) & s
End Function

Private Function str_padEnd(s As String, targetLength As Long, padString As String) As String
    Dim currentLen As Long
    Dim padLen As Long
    Dim padding As String
    
    currentLen = Len(s)
    
    If currentLen >= targetLength Then
        str_padEnd = s
        Exit Function
    End If
    
    padLen = targetLength - currentLen
    
    ' Build padding
    padding = ""
    Do While Len(padding) < padLen
        padding = padding & padString
    Loop
    
    str_padEnd = s & Left$(padding, padLen)
End Function

Private Function str_concat(s As String, args As Collection) As String
    Dim result As String
    Dim i As Long
    
    result = s
    
    For i = 1 To args.count
        Dim val As CValue
        Set val = args(i)
        result = result & val.ToString()
    Next
    
    str_concat = result
End Function


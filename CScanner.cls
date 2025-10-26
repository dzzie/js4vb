VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT

'------------------------------------------------------------
' CScanner.cls - JavaScript Lexical Scanner (Tokenizer)
'
' "Today is a good day to die!" - Worf, probably
'
' This is the lexer. It turns JavaScript source into tokens.
' Character by character. In VB6. Because we can.
'------------------------------------------------------------

Option Explicit

Private source As String
Private length As Long
Private index As Long
Private LineNumber As Long
Private LineStart As Long

' Lookahead buffer
Private buffer As String

' Token reference for keyword checking
Private tokenHelper As CToken
Private m_lastTokenType As TokenType
Private m_lastTokenValue As String

Private Sub Class_Initialize()
    Set tokenHelper = New CToken
End Sub

Public Sub init(code As String)
    source = code
    length = Len(code)
    index = 0
    LineNumber = 1
    LineStart = 0
    buffer = ""
End Sub

' ============================================
' Character Classification
' ============================================

Private Function IsWhiteSpace(ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim c As Integer
    c = Asc(ch)
    ' Space, tab, VT, FF, NBSP, BOM
    IsWhiteSpace = (c = 32 Or c = 9 Or c = 11 Or c = 12 Or c = 160)
End Function

Private Function IsLineTerminator(ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim c As Integer
    c = Asc(ch)
    ' LF, CR
    IsLineTerminator = (c = 10 Or c = 13)
End Function

Private Function IsIdentifierStart(ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim c As Integer
    c = Asc(ch)
    ' a-z, A-Z, $, [underscore]
    If (c >= 97 And c <= 122) Then
        IsIdentifierStart = True
    ElseIf (c >= 65 And c <= 90) Then
        IsIdentifierStart = True
    ElseIf c = 36 Or c = 95 Then
        IsIdentifierStart = True
    Else
        IsIdentifierStart = False
    End If
End Function

Private Function IsIdentifierPart(ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    If IsIdentifierStart(ch) Then
        IsIdentifierPart = True
        Exit Function
    End If
    ' Also allow digits 0-9
    Dim c As Integer
    c = Asc(ch)
    If c >= 48 And c <= 57 Then
        IsIdentifierPart = True
    Else
        IsIdentifierPart = False
    End If
End Function

Private Function IsDecimalDigit(ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim c As Integer
    c = Asc(ch)
    IsDecimalDigit = (c >= 48 And c <= 57) ' 0-9
End Function

Private Function IsHexDigit(ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    Dim c As Integer
    c = Asc(ch)
    ' 0-9, a-f, A-F
    If (c >= 48 And c <= 57) Then
        IsHexDigit = True
    ElseIf (c >= 97 And c <= 102) Then
        IsHexDigit = True
    ElseIf (c >= 65 And c <= 70) Then
        IsHexDigit = True
    Else
        IsHexDigit = False
    End If
End Function

' ============================================
' Character Access
' ============================================

Private Function CharAt(idx As Long) As String
    If idx >= length Then
        CharAt = ""
    Else
        CharAt = Mid$(source, idx + 1, 1)
    End If
End Function

Private Function CharCodeAt(idx As Long) As Integer
    Dim ch As String
    ch = CharAt(idx)
    If Len(ch) = 0 Then
        CharCodeAt = 0
    Else
        CharCodeAt = Asc(ch)
    End If
End Function

Private Function Peek() As String
    Peek = CharAt(index)
End Function

Private Function PeekAhead(offset As Long) As String
    PeekAhead = CharAt(index + offset)
End Function

' ============================================
' Skip Whitespace and Comments
' ============================================

Private Sub SkipSpaces()
    Do While index < length
        Dim ch As String
        ch = CharAt(index)
        
        If IsWhiteSpace(ch) Then
            index = index + 1
        ElseIf IsLineTerminator(ch) Then
            index = index + 1
            If ch = vbCr And CharAt(index) = vbLf Then
                index = index + 1
            End If
            LineNumber = LineNumber + 1
            LineStart = index
        Else
            Exit Sub
        End If
    Loop
End Sub

Private Function SkipSingleLineComment() As String
    Dim Start As Long
    Dim comment As String
    
    Start = index
    index = index + 2 ' Skip //
    
    Do While index < length
        Dim ch As String
        ch = CharAt(index)
        
        index = index + 1
        
        If IsLineTerminator(ch) Then
            If ch = vbCr And CharAt(index) = vbLf Then
                index = index + 1
            End If
            LineNumber = LineNumber + 1
            LineStart = index
            Exit Do
        End If
    Loop
    
    comment = Mid$(source, Start + 1, index - Start)
    SkipSingleLineComment = comment
End Function

Private Function SkipMultiLineComment() As String
    Dim Start As Long
    Dim comment As String
    
    Start = index
    index = index + 2 ' Skip /*
    
    Do While index < length
        Dim ch As String
        ch = CharAt(index)
        
        If IsLineTerminator(ch) Then
            If ch = vbCr And CharAt(index + 1) = vbLf Then
                index = index + 1
            End If
            LineNumber = LineNumber + 1
            LineStart = index + 1
            index = index + 1
        ElseIf ch = "*" Then
            ' Check for */
            If CharAt(index + 1) = "/" Then
                index = index + 2
                Exit Do
            End If
            index = index + 1
        Else
            index = index + 1
        End If
    Loop
    
    comment = Mid$(source, Start + 1, index - Start)
    SkipMultiLineComment = comment
End Function

Private Sub SkipComment()
    Do While index < length
        Dim ch As String
        ch = CharAt(index)
        
        If ch = "/" Then
            Dim next_ As String
            next_ = CharAt(index + 1)
            
            If next_ = "/" Then
                SkipSingleLineComment
            ElseIf next_ = "*" Then
                SkipMultiLineComment
            Else
                Exit Sub
            End If
        ElseIf IsWhiteSpace(ch) Then
            index = index + 1
        ElseIf IsLineTerminator(ch) Then
            index = index + 1
            If ch = vbCr And CharAt(index) = vbLf Then
                index = index + 1
            End If
            LineNumber = LineNumber + 1
            LineStart = index
        Else
            Exit Sub
        End If
    Loop
End Sub

' ============================================
' Scan Tokens
' ============================================

Private Function ScanIdentifier() As CToken
    Dim Start As Long
    Dim id As String
    Dim token As New CToken
    
    Start = index
    
    ' Must start with valid identifier start
    If Not IsIdentifierStart(CharAt(index)) Then
        Err.Raise vbObjectError + 1000, "CScanner", "Invalid identifier start"
    End If
    
    index = index + 1
    
    ' Continue while valid identifier part
    Do While index < length
        If Not IsIdentifierPart(CharAt(index)) Then
            Exit Do
        End If
        index = index + 1
    Loop
    
    id = Mid$(source, Start + 1, index - Start)
    
    ' Check if it's a keyword
    If tokenHelper.IsKeyword(id) Then
        If tokenHelper.IsBooleanLiteral(id) Then
            token.tType = BooleanLiteral
        ElseIf tokenHelper.IsNullLiteral(id) Then
            token.tType = NullLiteral
        Else
            token.tType = Keyword
        End If
    Else
        token.tType = Identifier
    End If
    
    token.Value = id
    token.LineNumber = LineNumber
    token.LineStart = LineStart
    token.Start = Start
    token.End_ = index
    
    Set ScanIdentifier = token
End Function

Private Function ScanPunctuator() As CToken
    Dim Start As Long
    Dim token As New CToken
    Dim ch As String
    Dim ch2 As String
    Dim ch3 As String
    
    Start = index
    ch = CharAt(index)
    
    token.tType = Punctuator
    token.LineNumber = LineNumber
    token.LineStart = LineStart
    token.Start = Start
    
    ' Try 3-character punctuators first
    ch2 = CharAt(index + 1)
    ch3 = CharAt(index + 2)
    
    If ch = "=" And ch2 = "=" And ch3 = "=" Then
        index = index + 3
        token.Value = "==="
    ElseIf ch = "!" And ch2 = "=" And ch3 = "=" Then
        index = index + 3
        token.Value = "!=="
    ElseIf ch = ">" And ch2 = ">" And ch3 = ">" Then
        index = index + 3
        token.Value = ">>>"
    ElseIf ch = "<" And ch2 = "<" And ch3 = "=" Then
        index = index + 3
        token.Value = "<<="
    ElseIf ch = ">" And ch2 = ">" And ch3 = "=" Then
        index = index + 3
        token.Value = ">>="
    ' 2-character punctuators
    ElseIf ch = "=" And ch2 = "=" Then
        index = index + 2
        token.Value = "=="
    ElseIf ch = "!" And ch2 = "=" Then
        index = index + 2
        token.Value = "!="
    ElseIf ch = "<" And ch2 = "=" Then
        index = index + 2
        token.Value = "<="
    ElseIf ch = ">" And ch2 = "=" Then
        index = index + 2
        token.Value = ">="
    ElseIf ch = "+" And ch2 = "+" Then
        index = index + 2
        token.Value = "++"
    ElseIf ch = "-" And ch2 = "-" Then
        index = index + 2
        token.Value = "--"
    ElseIf ch = "<" And ch2 = "<" Then
        index = index + 2
        token.Value = "<<"
    ElseIf ch = ">" And ch2 = ">" Then
        index = index + 2
        token.Value = ">>"
    ElseIf ch = "&" And ch2 = "&" Then
        index = index + 2
        token.Value = "&&"
    ElseIf ch = "|" And ch2 = "|" Then
        index = index + 2
        token.Value = "||"
    ElseIf ch = "+" And ch2 = "=" Then
        index = index + 2
        token.Value = "+="
    ElseIf ch = "-" And ch2 = "=" Then
        index = index + 2
        token.Value = "-="
    ElseIf ch = "*" And ch2 = "=" Then
        index = index + 2
        token.Value = "*="
    ElseIf ch = "/" And ch2 = "=" Then
        index = index + 2
        token.Value = "/="
    ElseIf ch = "%" And ch2 = "=" Then
        index = index + 2
        token.Value = "%="
    ElseIf ch = "&" And ch2 = "=" Then
        index = index + 2
        token.Value = "&="
    ElseIf ch = "|" And ch2 = "=" Then
        index = index + 2
        token.Value = "|="
    ElseIf ch = "^" And ch2 = "=" Then
        index = index + 2
        token.Value = "^="
    ' Single character punctuators
    ElseIf ch = "{" Or ch = "}" Or ch = "(" Or ch = ")" Or _
           ch = "[" Or ch = "]" Or ch = ";" Or ch = "," Or _
           ch = ":" Or ch = "?" Or ch = "." Or ch = "~" Or _
           ch = "!" Or ch = "+" Or ch = "-" Or ch = "*" Or _
           ch = "/" Or ch = "%" Or ch = "<" Or ch = ">" Or _
           ch = "=" Or ch = "&" Or ch = "|" Or ch = "^" Then
        index = index + 1
        token.Value = ch
    Else
        Err.Raise vbObjectError + 1001, "CScanner", "Unexpected character: " & ch
    End If
    
    token.End_ = index
    Set ScanPunctuator = token
End Function

'============================================
' BIGINT CONVERSION RULES
'============================================
' This engine uses TRUE 64-BIT INTEGERS (Currency-backed Int64)
' NOT JavaScript's fake "BigInt" that's just arbitrary precision!
'
' We leverage VB6's Currency type (8-byte scaled integer) via ULong64.cls
' which wraps utypes.dll for proper 64-bit math operations.
'
' Hex numbers are automatically promoted to Int64 based on size:
'
' +-----------------------------------------------------------------------------+
' | Hex Value       | Decimal          | Type     | Reason                      |
' +-----------------+------------------+----------+-----------------------------+
' | 0x10            | 16               | number   | Fits in 32-bit              |
' | 0x7FFFFFFF      | 2,147,483,647    | number   | Max signed 32-bit           |
' | 0x80000000      | 2,147,483,648    | number   | Needs unsigned, fits Double |
' | 0xFFFFFFFF      | 4,294,967,295    | number   | Max unsigned 32-bit         |
' | 0x100000000     | 4,294,967,296    | bigint   | Requires > 32-bit (hi!=0)   |
' | 0x10n           | 16               | bigint   | Explicit 'n' suffix         |
' +-----------------------------------------------------------------------------+
'
' Decimal numbers are promoted to Int64 when:
' +------------------------------------------------------------------------+
' | Decimal Value        | Type     | Reason                               |
' +----------------------+----------+--------------------------------------+
' | 123                  | number   | Small integer                        |
' | 4294967295           | number   | Fits in 32-bit unsigned              |
' | 4294967296           | bigint   | Exceeds 32-bit (auto-promoted)       |
' | 9007199254740992     | bigint   | Large integer (auto-promoted)        |
' | 123n                 | bigint   | Explicit 'n' suffix                  |
' +------------------------------------------------------------------------+
'
' REAL Int64 OPERATIONS (via utypes.dll):
' - Addition, subtraction, multiplication, division, modulus
' - Bitwise: AND, OR, XOR, shifts (<<, >>)
' - Comparisons: <, >, <=, >=, ==, !=
' - All operations maintain 64-bit precision
' - Range: -9,223,372,036,854,775,808 to 18,446,744,073,709,551,615
'
' Scanner Logic:
' - HEX: 8+ hex digits OR first digit 8-F (MSB set) -> Int64Literal
' - DECIMAL: > 10 digits OR value > 4294967295 -> Int64Literal
' - EXPLICIT: Any number with 'n' suffix -> Int64Literal
'
' Parser Logic:
' - Int64Literal tokens get "n" marker appended to value
' - NumericLiteral with negative lo gets converted: lo + 4294967296#
'
' Interpreter Logic:
' - Values ending in "n" are parsed as vtInt64 using ULong64
' - typeof vtInt64 returns "bigint"
' - Binary ops on Int64 use EvaluateInt64Arithmetic/Comparison
'
' JavaScript's BigInt vs Our Int64:
' - JS BigInt: Arbitrary precision, slow, heap-allocated objects
' - Our Int64: Fixed 64-bit, FAST native CPU operations via DLL
' - We're more like C's int64_t or uint64_t - REAL integers!
'============================================

Private Function ScanNumericLiteral() As CToken
    Dim token As CToken
    Dim Start As Long
    Dim buf As String
    Dim isHex As Boolean
    Dim needsInt64 As Boolean
    
    Start = index
    Set token = New CToken
    buf = ""
    isHex = False
    needsInt64 = False
    
    ' Check for hex prefix (0x or 0X)
    If Peek() = "0" And (PeekAhead(1) = "x" Or PeekAhead(1) = "X") Then
        isHex = True
        buf = buf & Peek() & PeekAhead(1)  ' Include "0x"
        index = index + 2
        
        ' Scan hex digits
        Do While IsHexDigit(Peek())
            buf = buf & Peek()
            index = index + 1
        Loop
        
        ' Check if hex value requires Int64 (more than 8 hex digits = > 32 bit)
        Dim hexDigits As Long
        hexDigits = Len(buf) - 2  ' Subtract "0x"
        
        ' If more than 8 hex digits OR value > 2^31-1, needs Int64
        If hexDigits > 8 Then
            needsInt64 = True
        ElseIf hexDigits = 8 Then
            ' Check if the value exceeds 2^31-1 (0x7FFFFFFF)
            Dim hexPart As String
            hexPart = Mid$(buf, 3)  ' Skip "0x"
            
            ' Convert to test magnitude
            On Error Resume Next
            Dim testVal As Double
            testVal = CDbl("&H" & hexPart)
            If Err.Number <> 0 Or testVal > 2147483647# Then
                needsInt64 = True
            End If
            On Error GoTo 0
        End If
        
    Else
        ' Standard decimal number
        
        ' Scan digits before decimal point
        Do While IsDecimalDigit(Peek())
            buf = buf & Peek()
            index = index + 1
        Loop
        
        ' Check for decimal point
        If Peek() = "." And IsDecimalDigit(PeekAhead(1)) Then
            buf = buf & Peek()
            index = index + 1
            
            ' Scan digits after decimal point
            Do While IsDecimalDigit(Peek())
                buf = buf & Peek()
                index = index + 1
            Loop
        End If
        
        ' Check for exponent (e or E)
        If Peek() = "e" Or Peek() = "E" Then
            buf = buf & Peek()
            index = index + 1
            
            ' Optional + or -
            If Peek() = "+" Or Peek() = "-" Then
                buf = buf & Peek()
                index = index + 1
            End If
            
            ' Scan exponent digits
            Do While IsDecimalDigit(Peek())
                buf = buf & Peek()
                index = index + 1
            Loop
        End If
        
        ' For decimal: Check if integer part exceeds 32-bit range
        ' Only check if no decimal point and no exponent
        If InStr(buf, ".") = 0 And InStr(LCase(buf), "e") = 0 Then
            On Error Resume Next
            Dim decVal As Double
            decVal = CDbl(buf)
            ' JavaScript safe integer range: 2^53 - 1, but we'll use 2^31-1 for VB6 Long
            ' Actually, let's be more lenient and only force Int64 if > 2^32
            If Err.Number <> 0 Or Abs(decVal) > 4294967295# Then
                needsInt64 = True
            End If
            On Error GoTo 0
        End If
    End If
    
    ' Check for explicit BigInt suffix 'n'
    Dim hasExplicitN As Boolean
    hasExplicitN = False
    
    If Peek() = "n" Then
        hasExplicitN = True
        needsInt64 = True
        index = index + 1  ' Consume 'n'
    End If
    
    ' Determine token type
    If needsInt64 Then
        token.tType = Int64Literal
        token.Value = buf  ' Store without the 'n' - we'll add it in the interpreter
    Else
        token.tType = NumericLiteral
        token.Value = buf
    End If
    
    token.LineNumber = LineNumber
    token.LineStart = LineStart
    token.Start = Start
    token.End_ = index
    
    Set ScanNumericLiteral = token
End Function


Private Function ScanStringLiteral() As CToken
    Dim Start As Long
    Dim token As New CToken
    Dim quote As String
    Dim str As String
    
    Start = index
    quote = CharAt(index)
    index = index + 1
    
    str = ""
    
    Do While index < length
        Dim ch As String
        ch = CharAt(index)
        
        If ch = quote Then
            index = index + 1
            Exit Do
        ElseIf ch = "\" Then
            ' Escape sequence
            index = index + 1
            If index >= length Then
                Err.Raise vbObjectError + 1002, "CScanner", "Unterminated string"
            End If
            ch = CharAt(index)
            index = index + 1
            
            ' Handle common escapes
            Select Case ch
                Case "n": str = str & vbLf
                Case "r": str = str & vbCr
                Case "t": str = str & vbTab
                Case "b", "f", "v": str = str ' Ignore these
                Case "\": str = str & "\"
                Case "'": str = str & "'"
                Case """": str = str & """"
                Case "x"
                    ' Hex escape \xNN
                    Dim hex As String
                    hex = Mid$(source, index + 1, 2)
                    If Len(hex) = 2 And IsHexDigit(Mid$(hex, 1, 1)) And IsHexDigit(Mid$(hex, 2, 1)) Then
                        str = str & Chr$("&H" & hex)
                        index = index + 2
                    Else
                        str = str & ch
                    End If
                Case "u"
                    ' Unicode escape \uNNNN - simplified
                    Dim uni As String
                    uni = Mid$(source, index + 1, 4)
                    If Len(uni) = 4 Then
                        ' Try to convert (will fail on high unicode in VB6)
                        On Error Resume Next
                        str = str & ChrW$("&H" & uni)
                        If Err.Number <> 0 Then
                            str = str & "?" ' Placeholder for unsupported unicode
                        End If
                        On Error GoTo 0
                        index = index + 4
                    Else
                        str = str & ch
                    End If
                Case Else
                    str = str & ch
            End Select
        ElseIf IsLineTerminator(ch) Then
            Err.Raise vbObjectError + 1002, "CScanner", "Unterminated string"
        Else
            str = str & ch
            index = index + 1
        End If
    Loop
    
    token.tType = StringLiteral
    token.Value = str
    token.LineNumber = LineNumber
    token.LineStart = LineStart
    token.Start = Start
    token.End_ = index
    
    Set ScanStringLiteral = token
End Function

' ============================================
' Main Scanning Function
' ============================================

Public Function Lex() As CToken
    SkipComment
    
    If index >= length Then
        Dim eofToken As New CToken
        eofToken.tType = EOF_TOKEN
        eofToken.LineNumber = LineNumber
        eofToken.LineStart = LineStart
        eofToken.Start = index
        eofToken.End_ = index
        Set Lex = eofToken
        m_lastTokenType = EOF_TOKEN
        Exit Function
    End If
    
    Dim ch As String
    ch = CharAt(index)
    
    Dim resultToken As CToken
    
    ' Identifier or keyword
    If IsIdentifierStart(ch) Then
        Set resultToken = ScanIdentifier()
    ' Number
    ElseIf IsDecimalDigit(ch) Then
        Set resultToken = ScanNumericLiteral()
    ' String
    ElseIf ch = "'" Or ch = """" Then
        Set resultToken = ScanStringLiteral()
    ' Check for regex vs division
    ElseIf ch = "/" Then
        ' Regex appears after: =, (, [, {, ,, ;, :, !, &, |, ^, ~, ?, +, -, *, /, %, <, >, return, throw, etc.
        ' Division appears after: identifier, ), ], number, string
        
        If CanStartRegex() Then
            Set resultToken = ScanRegexLiteral()
        Else
            Set resultToken = ScanPunctuator()
        End If
    ' Punctuator
    Else
        Set resultToken = ScanPunctuator()
    End If

    m_lastTokenType = resultToken.tType
    m_lastTokenValue = resultToken.Value
    Set Lex = resultToken
    
End Function

Private Function CanStartRegex() As Boolean
    Select Case m_lastTokenType
        Case EOF_TOKEN
            CanStartRegex = True
            
        Case Punctuator
            ' Check the actual punctuator
            Select Case m_lastTokenValue
                Case ")", "]"
                    ' After closing parens/brackets, it's division
                    CanStartRegex = False
                Case "++", "--"
                    ' After increment/decrement, it's division
                    CanStartRegex = False
                Case Else
                    ' After =, (, [, {, ,, ;, :, operators, etc. - it's regex
                    CanStartRegex = True
            End Select
            
        Case Keyword
            ' After keywords like return, throw, new - it's regex
            ' But after 'this' - it's division
            If m_lastTokenValue = "this" Then
                CanStartRegex = False
            Else
                CanStartRegex = True
            End If
            
        Case Identifier, NumericLiteral, StringLiteral, BooleanLiteral, NullLiteral, RegularExpression
            ' After these, / is division
            CanStartRegex = False
            
        Case Else
            ' Default to division
            CanStartRegex = False
    End Select
End Function


Public Function IsEOF() As Boolean
    IsEOF = (index >= length)
End Function

Private Function ScanRegexLiteral() As CToken
    Dim Start As Long
    Dim token As New CToken
    Dim pattern As String
    
    Start = index
    index = index + 1  ' Skip opening /
    
    pattern = ""
    Dim inCharClass As Boolean
    inCharClass = False
    
    ' Scan the regex pattern
    Do While index < length
        Dim ch As String
        ch = CharAt(index)
        
        If ch = "[" And Not inCharClass Then
            ' Start of character class
            inCharClass = True
            pattern = pattern & ch
            index = index + 1
        ElseIf ch = "]" And inCharClass Then
            ' End of character class
            inCharClass = False
            pattern = pattern & ch
            index = index + 1
        ElseIf ch = "/" And Not inCharClass Then
            ' End of regex pattern (only if not in character class)
            index = index + 1
            Exit Do
        ElseIf ch = "\" Then
            ' Escape sequence
            pattern = pattern & ch
            index = index + 1
            If index < length Then
                pattern = pattern & CharAt(index)
                index = index + 1
            End If
        ElseIf IsLineTerminator(ch) Then
            ' Regex cannot contain unescaped line terminators
            Err.Raise vbObjectError + 1003, "CScanner", "Unterminated regex at line " & LineNumber
        Else
            pattern = pattern & ch
            index = index + 1
        End If
    Loop
    
    ' Check if we properly closed the regex
    If index > length Or CharAt(index - 1) <> "/" Then
        Err.Raise vbObjectError + 1003, "CScanner", "Unterminated regex at line " & LineNumber & " - pattern so far: /" & pattern
    End If
    
    ' Scan flags (g, i, m, y, u, s)
    Dim flags As String
    flags = ""
    Do While index < length
        ch = CharAt(index)
        If ch = "g" Or ch = "i" Or ch = "m" Or ch = "y" Or ch = "u" Or ch = "s" Then
            flags = flags & ch
            index = index + 1
        Else
            Exit Do
        End If
    Loop
    
    token.tType = RegularExpression
    token.Value = "/" & pattern & "/" & flags
    token.LineNumber = LineNumber
    token.LineStart = LineStart
    token.Start = Start
    token.End_ = index
    
    Set ScanRegexLiteral = token
End Function

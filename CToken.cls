VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT

'------------------------------------------------------------
' CToken.cls - JavaScript Token
'
' Because VB6 is awesome and deserves it.
'------------------------------------------------------------

' Token types
Public Enum TokenType
    EOF_TOKEN = 0
    
    ' Literals
    BooleanLiteral = 1
    NullLiteral = 2
    NumericLiteral = 3
    StringLiteral = 4
    RegularExpression = 5
    
    ' Identifiers and Keywords
    Identifier = 6
    Keyword = 7
    
    ' Punctuators
    Punctuator = 8
    
    ' Comments (preserved for analysis)
    LineComment = 9
    BlockComment = 10
    
    ' Int64 (BigInt)
    Int64Literal = 11       ' NEW: BigInt literals (123n)
End Enum

' Token data
Public tType As TokenType
Public value As String
Public LineNumber As Long
Public LineStart As Long
Public Start As Long
Public End_ As Long  ' "End" is reserved in VB

Public Function IsKeyword(word As String) As Boolean
    ' JavaScript keywords are case-sensitive
    ' "function" is a keyword, "Function" is not - we can not use collection key lookup
    Select Case word
        Case "break", "case", "catch", "continue", "debugger", _
             "default", "delete", "do", "else", "finally", _
             "for", "function", "if", "in", "instanceof", _
             "new", "return", "switch", "this", "throw", _
             "try", "typeof", "var", "void", "while", "with", _
             "null", "true", "false"
            IsKeyword = True
        Case Else
            IsKeyword = False
    End Select
End Function

Public Function IsBooleanLiteral(word As String) As Boolean
    IsBooleanLiteral = (word = "true" Or word = "false")
End Function

Public Function IsNullLiteral(word As String) As Boolean
    IsNullLiteral = (word = "null")
End Function

Public Function ToString() As String
    ToString = "Token{" & Me.tType & ", '" & Me.value & "', line:" & Me.LineNumber & "}"
End Function

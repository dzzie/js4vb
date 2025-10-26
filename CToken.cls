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
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
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
Public Value As String
Public LineNumber As Long
Public LineStart As Long
Public Start As Long
Public End_ As Long  ' "End" is reserved in VB

' Keywords lookup
Private keywords As Collection

Private Sub Class_Initialize()
    Set keywords = New Collection
    
    ' ES5 keywords
    On Error Resume Next
    keywords.add "break", "break"
    keywords.add "case", "case"
    keywords.add "catch", "catch"
    keywords.add "continue", "continue"
    keywords.add "debugger", "debugger"
    keywords.add "default", "default"
    keywords.add "delete", "delete"
    keywords.add "do", "do"
    keywords.add "else", "else"
    keywords.add "finally", "finally"
    keywords.add "for", "for"
    keywords.add "function", "function"
    keywords.add "if", "if"
    keywords.add "in", "in"
    keywords.add "instanceof", "instanceof"
    keywords.add "new", "new"
    keywords.add "return", "return"
    keywords.add "switch", "switch"
    keywords.add "this", "this"
    keywords.add "throw", "throw"
    keywords.add "try", "try"
    keywords.add "typeof", "typeof"
    keywords.add "var", "var"
    keywords.add "void", "void"
    keywords.add "while", "while"
    keywords.add "with", "with"
    
    ' Literals
    keywords.add "null", "null"
    keywords.add "true", "true"
    keywords.add "false", "false"
    On Error GoTo 0
End Sub

Public Function IsKeyword(word As String) As Boolean
    On Error Resume Next
    Dim X As String
    X = keywords(word)
    IsKeyword = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function IsBooleanLiteral(word As String) As Boolean
    IsBooleanLiteral = (word = "true" Or word = "false")
End Function

Public Function IsNullLiteral(word As String) As Boolean
    IsNullLiteral = (word = "null")
End Function

Public Function ToString() As String
    ToString = "Token{" & Me.tType & ", '" & Me.Value & "', line:" & Me.LineNumber & "}"
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT

'====================================================================
' CReference.cls - Reference to a symbol (variable/function part of code analyzer)
'====================================================================

Option Explicit

Public Enum ReferenceType
    RefRead = 1         ' Reading the value (x in: y = x)
    RefWrite = 2        ' Writing the value (x in: x = 5)
    RefCall = 3         ' Calling as function (foo in: foo())
    RefDeclaration = 4  ' Declaration site (x in: var x)
    RefReadWrite = 5    ' Read-write (+=, etc.)  ' NEW
End Enum

' Location
Public LineNumber As Long
Public ColumnStart As Long
Public ColumnEnd As Long
Public Node As CNode                ' AST node of reference

' Context
Public refType As ReferenceType
Public ParentFunction As String     ' Name of containing function (or "" for global)
Public Context As String            ' Brief context string

' NEW: Enhanced context
Public CodeSnippet As String        ' Line of code for this reference
Public FileName As String           ' For multi-file analysis


Public Function ToString() As String
    Dim typeStr As String
    Select Case refType
        Case RefRead: typeStr = "READ"
        Case RefWrite: typeStr = "WRITE"
        Case RefCall: typeStr = "CALL"
        Case RefDeclaration: typeStr = "DECL"
        Case RefReadWrite: typeStr = "R/W "
    End Select
    
    Dim location As String
    If Len(FileName) > 0 Then
        location = FileName & ":" & LineNumber
    Else
        location = "line " & LineNumber
    End If
    
    Dim scope As String
    scope = IIf(Len(ParentFunction) > 0, ParentFunction, "[global]")
    
    ToString = "[" & typeStr & "] " & location & " in " & scope
    
    ' Add code snippet if available
    If Len(CodeSnippet) > 0 Then
        ToString = ToString & vbCrLf & "      " & Trim$(CodeSnippet)
    End If
End Function

Public Function ToJSON() As String
    ToJSON = "{" & _
        """type"": """ & refType & """," & _
        """line"": " & LineNumber & "," & _
        """function"": """ & EscapeJSON(ParentFunction) & """," & _
        """snippet"": """ & EscapeJSON(CodeSnippet) & """" & _
        "}"
End Function

Private Function EscapeJSON(s As String) As String
    EscapeJSON = Replace(Replace(s, "\", "\\"), """", "\""")
End Function


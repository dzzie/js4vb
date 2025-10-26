VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'====================================================================
' CVariable.cls - Variable metadata and references (part of code analyzer)
'====================================================================
Option Explicit

' Metadata
Public Name As String
Public DeclaredAt As Long           ' Line number where declared
Public declNode As CNode            ' AST node of declaration
Public ScopeLevel As Long           ' 0=global, 1=function, 2=nested, etc.
Public ScopeName As String          ' Name of containing function/scope
Public IsParameter As Boolean       ' Is this a function parameter?
Public IsGlobal As Boolean          ' Is this a global variable?

' References
Public References As Collection     ' Collection of CReference objects
Public WriteCount As Long           ' Number of writes
Public ReadCount As Long            ' Number of reads

' Type inference (best-effort)
Public InferredType As String       ' "number", "string", "object", "function", "unknown"
Public InitialValue As Variant      ' Initial value if literal

Public Function IsUnused() As Boolean
    IsUnused = (References.count = 0)
End Function

Public Function IsDeadCode() As Boolean
    ' Written but never read
    IsDeadCode = (WriteCount > 0 And ReadCount = 0)
End Function

Public Function IsWriteOnly() As Boolean
    IsWriteOnly = (WriteCount > 0 And ReadCount = 0)
End Function

Public Function IsReadOnly() As Boolean
    IsReadOnly = (WriteCount = 0 And ReadCount > 0)
End Function

Public Function TotalReferences() As Long
    TotalReferences = References.count
End Function

Private Sub Class_Initialize()
    Set References = New Collection
    WriteCount = 0
    ReadCount = 0
    InferredType = "unknown"
End Sub

Public Function ToString() As String
    ToString = "Variable '" & Name & "' (" & InferredType & ") - " & _
               References.count & " refs (" & ReadCount & " reads, " & WriteCount & " writes)"
End Function

Public Function GetReferenceReport() As String
    Dim report As String
    report = "=== REFERENCES FOR '" & Name & "' ===" & vbCrLf & vbCrLf
    
    report = report & "Type: " & InferredType & vbCrLf
    report = report & "Scope: " & ScopeName & " (level " & ScopeLevel & ")" & vbCrLf
    report = report & "Declared at line: " & DeclaredAt & vbCrLf
    report = report & "Total references: " & TotalReferences() & vbCrLf  ' ? Function
    report = report & "Reads: " & ReadCount & vbCrLf
    report = report & "Writes: " & WriteCount & vbCrLf
    
    ' Use computed functions
    If IsUnused() Then
        report = report & "? UNUSED!" & vbCrLf
    End If
    
    If IsDeadCode() Then
        report = report & "? DEAD CODE (written but never read)!" & vbCrLf
    End If
    
    report = report & vbCrLf & "References:" & vbCrLf
    
    Dim i As Long
    For i = 1 To References.count
        Dim ref As CReference
        Set ref = References(i)
        report = report & "  " & ref.ToString() & vbCrLf
    Next
    
    GetReferenceReport = report
End Function

' Get just the cross-references (compatibility)
Public Function GetXRefSummary() As String
    GetXRefSummary = GetReferenceReport()
End Function


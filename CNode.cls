VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT
Option Explicit


'------------------------------------------------------------
' CNode.cls - AST Node Base Class
'
' "To infinity and beyond!" - Buzz Lightyear
'
' Every JavaScript construct becomes a node in the AST.
' This is the base class. We're building a TREE. In VB6.
'------------------------------------------------------------

Public Enum NodeType
    ' Program
    Program_Node = 1
    
    ' Statements
    EmptyStatement_Node = 10
    BlockStatement_Node = 11
    ExpressionStatement_Node = 12
    IfStatement_Node = 13
    WhileStatement_Node = 14
    DoWhileStatement_Node = 15
    ForStatement_Node = 16
    ForInStatement_Node = 17
    ContinueStatement_Node = 18
    BreakStatement_Node = 19
    ReturnStatement_Node = 20
    WithStatement_Node = 21
    SwitchStatement_Node = 22
    ThrowStatement_Node = 23
    TryStatement_Node = 24
    
    ' Declarations
    VariableDeclaration_Node = 30
    VariableDeclarator_Node = 31
    FunctionDeclaration_Node = 32
    
    ' Expressions
    ThisExpression_Node = 40
    ArrayExpression_Node = 41
    ObjectExpression_Node = 42
    Property_Node = 43
    FunctionExpression_Node = 44
    SequenceExpression_Node = 45
    UnaryExpression_Node = 46
    BinaryExpression_Node = 47
    AssignmentExpression_Node = 48
    UpdateExpression_Node = 49
    LogicalExpression_Node = 50
    ConditionalExpression_Node = 51
    CallExpression_Node = 52
    NewExpression_Node = 53
    MemberExpression_Node = 54
    
    ' Literals
    Identifier_Node = 60
    Literal_Node = 61
    
    ' Clauses
    SwitchCase_Node = 70
    CatchClause_Node = 71
End Enum

' Base properties all nodes have
Public tType As NodeType
Public LineNumber As Long
Public startPos As Long
Public endPos As Long

' Node-specific data (stored as Variant to handle different types)
' Program
Public Body As Collection           ' Collection of statements

' Identifier
Public Name As String

' Literal
Public value As Variant
Public Raw As String

' Binary/Logical/Assignment Expression
Public operator As String
Public Left As CNode
Public Right As CNode

' Unary/Update Expression
Public Argument As CNode
Public Prefix As Boolean

' Conditional Expression
Public Test As CNode
Public Consequent As CNode
Public Alternate As CNode

' Member Expression
Public Object As CNode
Public prop As CNode
Public Computed As Boolean

' Call/New Expression
Public Callee As CNode
Public Arguments As Collection

' Function Declaration/Expression
Public id As CNode                  ' Identifier (can be Nothing for anonymous)
Public Params As Collection         ' Collection of CNode (Identifiers)
Public FunctionBody As CNode        ' BlockStatement

' Variable Declaration
Public Declarations As Collection   ' Collection of VariableDeclarator
Public Kind As String              ' "var", "let", "const"

' Variable Declarator
Public VarID As CNode              ' Identifier
Public Init As CNode               ' Expression (can be Nothing)

' If Statement
Public IfTest As CNode
Public IfConsequent As CNode
Public IfAlternate As CNode        ' Can be Nothing

' While/DoWhile Statement
Public WhileTest As CNode
Public WhileBody As CNode

' For Statement
Public ForInit As CNode
Public ForTest As CNode
Public ForUpdate As CNode
Public ForBody As CNode

' Return/Throw Statement
Public ReturnArgument As CNode     ' Can be Nothing for bare return

' Object Expression
Public Properties As Collection    ' Collection of Property nodes

' Sequence Expression
Public Expressions As Collection   ' Collection of expressions


' Property (in object literal)
Public PropKey As CNode
Public propValue As CNode
Public propKind As String          ' "init", "get", "set"

' Try Statement
Public TryBlock As CNode
Public TryHandler As CNode         ' CatchClause
Public TryFinalizer As CNode       ' Can be Nothing

' Catch Clause
Public param As CNode              ' Identifier
Public CatchBody As CNode          ' BlockStatement

' Switch Statement
Public discriminant As CNode
Public Cases As Collection         ' Collection of SwitchCase

' Switch Case
Public CaseTest As CNode           ' Can be Nothing for default
Public CaseConsequent As Collection ' Collection of statements

Private Sub Class_Initialize()
    ' Initialize all object references to Nothing explicitly
    Set Test = Nothing
    Set Consequent = Nothing
    Set Alternate = Nothing
    Set Init = Nothing
    Set ForUpdate = Nothing
    Set Body = Nothing
    Set Object = Nothing
    Set prop = Nothing
    Set Callee = Nothing
    Set Left = Nothing
    Set Right = Nothing
    
    ' Initialize collections
    Set Body = New Collection
    Set Arguments = New Collection
    Set Params = New Collection
    Set Declarations = New Collection
    Set Properties = New Collection
    Set Cases = New Collection
    Set CaseConsequent = New Collection
    Set Expressions = New Collection
    
End Sub

Private Function NodeTypeName(t As NodeType) As String
    Select Case t
        Case Program_Node: NodeTypeName = "Program"
        Case Identifier_Node: NodeTypeName = "Identifier"
        Case Literal_Node: NodeTypeName = "Literal"
        Case BinaryExpression_Node: NodeTypeName = "BinaryExpression"
        Case LogicalExpression_Node: NodeTypeName = "LogicalExpression"
        Case AssignmentExpression_Node: NodeTypeName = "AssignmentExpression"
        Case UnaryExpression_Node: NodeTypeName = "UnaryExpression"
        Case UpdateExpression_Node: NodeTypeName = "UpdateExpression"
        Case MemberExpression_Node: NodeTypeName = "MemberExpression"
        Case CallExpression_Node: NodeTypeName = "CallExpression"
        Case NewExpression_Node: NodeTypeName = "NewExpression"
        Case FunctionDeclaration_Node: NodeTypeName = "FunctionDeclaration"
        Case FunctionExpression_Node: NodeTypeName = "FunctionExpression"
        Case VariableDeclaration_Node: NodeTypeName = "VariableDeclaration"
        Case VariableDeclarator_Node: NodeTypeName = "VariableDeclarator"
        Case BlockStatement_Node: NodeTypeName = "BlockStatement"
        Case ExpressionStatement_Node: NodeTypeName = "ExpressionStatement"
        Case IfStatement_Node: NodeTypeName = "IfStatement"
        Case WhileStatement_Node: NodeTypeName = "WhileStatement"
        Case DoWhileStatement_Node: NodeTypeName = "DoWhileStatement"
        Case ForStatement_Node: NodeTypeName = "ForStatement"
        Case ReturnStatement_Node: NodeTypeName = "ReturnStatement"
        Case BreakStatement_Node: NodeTypeName = "BreakStatement"
        Case ContinueStatement_Node: NodeTypeName = "ContinueStatement"
        Case ThrowStatement_Node: NodeTypeName = "ThrowStatement"
        Case TryStatement_Node: NodeTypeName = "TryStatement"
        Case CatchClause_Node: NodeTypeName = "CatchClause"
        Case SwitchStatement_Node: NodeTypeName = "SwitchStatement"
        Case SwitchCase_Node: NodeTypeName = "SwitchCase"
        Case ObjectExpression_Node: NodeTypeName = "ObjectExpression"
        Case ArrayExpression_Node: NodeTypeName = "ArrayExpression"
        Case Property_Node: NodeTypeName = "Property"
        Case ThisExpression_Node: NodeTypeName = "ThisExpression"
        Case ConditionalExpression_Node: NodeTypeName = "ConditionalExpression"
        Case SequenceExpression_Node: NodeTypeName = "SequenceExpression"
        Case EmptyStatement_Node: NodeTypeName = "EmptyStatement"
        Case WithStatement_Node: NodeTypeName = "WithStatement"
        Case ForInStatement_Node: NodeTypeName = "ForInStatement"
        Case Else: NodeTypeName = "Unknown"
    End Select
End Function

Private Function FormatValue(v As Variant) As String
    If IsNull(v) Then
        FormatValue = "null"
    ElseIf VarType(v) = vbBoolean Then
        FormatValue = IIf(v, "true", "false")
    ElseIf VarType(v) = vbString Then
        FormatValue = """" & EscapeJSON(CStr(v)) & """"
    ElseIf IsNumeric(v) Then
        FormatValue = CStr(v)
    Else
        FormatValue = "null"
    End If
End Function

Public Function ToJSON(Optional Indent As Long = 0) As String
    Dim json As String
    Dim spaces As String
    Dim nextIndent As String
    Dim i As Long
    Dim myProp As CNode
     
    ' Create indentation strings
    spaces = Space$(Indent)
    nextIndent = Space$(Indent + 2)
    
    ' Handle Nothing case
    If Me Is Nothing Then
        ToJSON = "null"
        Exit Function
    End If
    
    Select Case Me.tType
        Case Program_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""Program""," & vbCrLf
            json = json & nextIndent & """body"": [" & vbCrLf
            
            If Not Body Is Nothing Then
                For i = 1 To Body.count
                    Dim stmt As CNode
                    Set stmt = Body(i)
                    json = json & stmt.ToJSON(Indent + 4)
                    If i < Body.count Then json = json & ","
                    json = json & vbCrLf
                Next i
            End If
            
            json = json & nextIndent & "]" & vbCrLf
            json = json & spaces & "}"
            
        Case Identifier_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""Identifier""," & vbCrLf
            json = json & nextIndent & """name"": """ & EscapeJSON(Name) & """" & vbCrLf
            json = json & spaces & "}"
            
        Case Literal_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""Literal""," & vbCrLf
            json = json & nextIndent & """value"": " & FormatLiteralValue(value) & "," & vbCrLf
            json = json & nextIndent & """raw"": """ & EscapeJSON(Raw) & """" & vbCrLf
            json = json & spaces & "}"
            
        Case FunctionDeclaration_Node, FunctionExpression_Node
            Dim funcType As String
            If tType = FunctionDeclaration_Node Then
                funcType = "FunctionDeclaration"
            Else
                funcType = "FunctionExpression"
            End If
            
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": """ & funcType & """," & vbCrLf
            
            ' ID (can be null for anonymous functions)
            json = json & nextIndent & """id"": "
            If Not id Is Nothing Then
                json = json & id.ToJSON(0)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            
            ' Parameters
            json = json & nextIndent & """params"": ["
            If Not Params Is Nothing Then
                If Params.count > 0 Then
                    json = json & vbCrLf
                    For i = 1 To Params.count
                        Dim param As CNode
                        Set param = Params(i)
                        json = json & param.ToJSON(Indent + 4)
                        If i < Params.count Then json = json & ","
                        json = json & vbCrLf
                    Next i
                    json = json & nextIndent
                End If
            End If
            json = json & "]," & vbCrLf
            
            ' Body
            json = json & nextIndent & """body"": "
            If Not FunctionBody Is Nothing Then
                json = json & FunctionBody.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case BlockStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""BlockStatement""," & vbCrLf
            json = json & nextIndent & """body"": ["
            
            If Not Body Is Nothing Then
                If Body.count > 0 Then
                    json = json & vbCrLf
                    For i = 1 To Body.count
                        Set stmt = Body(i)
                        json = json & stmt.ToJSON(Indent + 4)
                        If i < Body.count Then json = json & ","
                        json = json & vbCrLf
                    Next i
                    json = json & nextIndent
                End If
            End If
            
            json = json & "]" & vbCrLf
            json = json & spaces & "}"

        Case ExpressionStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""ExpressionStatement""," & vbCrLf
            json = json & nextIndent & """expression"": "
            If Not Test Is Nothing Then  ' ? Change Expression to Test
                json = json & Test.ToJSON(Indent + 2)  ' ? Change Expression to Test
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case VariableDeclaration_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""VariableDeclaration""," & vbCrLf
            json = json & nextIndent & """declarations"": ["
            
            If Not Declarations Is Nothing Then
                If Declarations.count > 0 Then
                    json = json & vbCrLf
                    For i = 1 To Declarations.count
                        Dim decl As CNode
                        Set decl = Declarations(i)
                        json = json & decl.ToJSON(Indent + 4)
                        If i < Declarations.count Then json = json & ","
                        json = json & vbCrLf
                    Next i
                    json = json & nextIndent
                End If
            End If
            
            json = json & "]," & vbCrLf
            json = json & nextIndent & """kind"": ""var""" & vbCrLf
            json = json & spaces & "}"
            
        Case VariableDeclarator_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""VariableDeclarator""," & vbCrLf
            json = json & nextIndent & """id"": "
            If Not VarID Is Nothing Then
                json = json & VarID.ToJSON(0)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """init"": "
            If Not Init Is Nothing Then
                json = json & Init.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case BinaryExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""BinaryExpression""," & vbCrLf
            json = json & nextIndent & """operator"": """ & EscapeJSON(operator) & """," & vbCrLf
            json = json & nextIndent & """left"": "
            If Not Left Is Nothing Then
                json = json & Left.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """right"": "
            If Not Right Is Nothing Then
                json = json & Right.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case LogicalExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""LogicalExpression""," & vbCrLf
            json = json & nextIndent & """operator"": """ & EscapeJSON(operator) & """," & vbCrLf
            json = json & nextIndent & """left"": "
            If Not Left Is Nothing Then
                json = json & Left.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """right"": "
            If Not Right Is Nothing Then
                json = json & Right.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case AssignmentExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""AssignmentExpression""," & vbCrLf
            json = json & nextIndent & """operator"": """ & EscapeJSON(operator) & """," & vbCrLf
            json = json & nextIndent & """left"": "
            If Not Left Is Nothing Then
                json = json & Left.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """right"": "
            If Not Right Is Nothing Then
                json = json & Right.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case UnaryExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""UnaryExpression""," & vbCrLf
            json = json & nextIndent & """operator"": """ & EscapeJSON(operator) & """," & vbCrLf
            json = json & nextIndent & """prefix"": " & IIf(Prefix, "true", "false") & "," & vbCrLf
            json = json & nextIndent & """argument"": "
            If Not Argument Is Nothing Then
                json = json & Argument.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case UpdateExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""UpdateExpression""," & vbCrLf
            json = json & nextIndent & """operator"": """ & EscapeJSON(operator) & """," & vbCrLf
            json = json & nextIndent & """prefix"": " & IIf(Prefix, "true", "false") & "," & vbCrLf
            json = json & nextIndent & """argument"": "
            If Not Argument Is Nothing Then
                json = json & Argument.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case ConditionalExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""ConditionalExpression""," & vbCrLf
            json = json & nextIndent & """test"": "
            If Not Test Is Nothing Then
                json = json & Test.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """consequent"": "
            If Not Consequent Is Nothing Then
                json = json & Consequent.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """alternate"": "
            If Not Alternate Is Nothing Then
                json = json & Alternate.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case CallExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""CallExpression""," & vbCrLf
            json = json & nextIndent & """callee"": "
            If Not Callee Is Nothing Then
                json = json & Callee.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """arguments"": ["
            
            If Not Arguments Is Nothing Then
                If Arguments.count > 0 Then
                    json = json & vbCrLf
                    For i = 1 To Arguments.count
                        Dim arg As CNode
                        Set arg = Arguments(i)
                        json = json & arg.ToJSON(Indent + 4)
                        If i < Arguments.count Then json = json & ","
                        json = json & vbCrLf
                    Next i
                    json = json & nextIndent
                End If
            End If
            
            json = json & "]" & vbCrLf
            json = json & spaces & "}"
            
        Case NewExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""NewExpression""," & vbCrLf
            json = json & nextIndent & """callee"": "
            If Not Callee Is Nothing Then
                json = json & Callee.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """arguments"": ["
            
            If Not Arguments Is Nothing Then
                If Arguments.count > 0 Then
                    json = json & vbCrLf
                    For i = 1 To Arguments.count
                        Set arg = Arguments(i)
                        json = json & arg.ToJSON(Indent + 4)
                        If i < Arguments.count Then json = json & ","
                        json = json & vbCrLf
                    Next i
                    json = json & nextIndent
                End If
            End If
            
            json = json & "]" & vbCrLf
            json = json & spaces & "}"
            
        Case MemberExpression_Node
            'Stop
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""MemberExpression""," & vbCrLf
            json = json & nextIndent & """object"": "
            If Not Object Is Nothing Then
                json = json & Object.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """property"": "
            If Not prop Is Nothing Then
                json = json & prop.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """computed"": " & IIf(Computed, "true", "false") & vbCrLf
            json = json & spaces & "}"
            
        Case ArrayExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""ArrayExpression""," & vbCrLf
            json = json & nextIndent & """elements"": ["
            
            If Not Body Is Nothing Then
                If Body.count > 0 Then
                    json = json & vbCrLf
                    For i = 1 To Body.count
                        Dim elem As CNode
                        Set elem = Body(i)
                        If elem Is Nothing Then
                            json = json & Space$(Indent + 4) & "null"
                        Else
                            json = json & elem.ToJSON(Indent + 4)
                        End If
                        If i < Body.count Then json = json & ","
                        json = json & vbCrLf
                    Next i
                    json = json & nextIndent
                End If
            End If
            
            json = json & "]" & vbCrLf
            json = json & spaces & "}"
            
        Case ObjectExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""ObjectExpression""," & vbCrLf
            json = json & nextIndent & """properties"": ["
            
            If Not Properties Is Nothing Then
                If Properties.count > 0 Then
                    json = json & vbCrLf
                    For i = 1 To Properties.count
                        Set myProp = Properties(i)
                        json = json & myProp.ToJSON(Indent + 4)
                        If i < Properties.count Then json = json & ","
                        json = json & vbCrLf
                    Next i
                    json = json & nextIndent
                End If
            End If
            
            json = json & "]" & vbCrLf
            json = json & spaces & "}"
            
        Case Property_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""Property""," & vbCrLf
            json = json & nextIndent & """key"": "
            If Not PropKey Is Nothing Then
                json = json & PropKey.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """value"": "
            If Not propValue Is Nothing Then
                json = json & propValue.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """kind"": """ & propKind & """" & vbCrLf
            json = json & spaces & "}"
            
        Case ThisExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""ThisExpression""" & vbCrLf
            json = json & spaces & "}"
            
        Case SequenceExpression_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""SequenceExpression""," & vbCrLf
            json = json & nextIndent & """expressions"": ["
            
            If Not Expressions Is Nothing Then
                If Expressions.count > 0 Then
                    json = json & vbCrLf
                    For i = 1 To Expressions.count
                        Dim expr As CNode
                        Set expr = Expressions(i)
                        json = json & expr.ToJSON(Indent + 4)
                        If i < Expressions.count Then json = json & ","
                        json = json & vbCrLf
                    Next i
                    json = json & nextIndent
                End If
            End If
            
            json = json & "]" & vbCrLf
            json = json & spaces & "}"
            
        Case IfStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""IfStatement""," & vbCrLf
            json = json & nextIndent & """test"": "
            If Not IfTest Is Nothing Then
                json = json & IfTest.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """consequent"": "
            If Not IfConsequent Is Nothing Then
                json = json & IfConsequent.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """alternate"": "
            If Not IfAlternate Is Nothing Then
                json = json & IfAlternate.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case WhileStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""WhileStatement""," & vbCrLf
            json = json & nextIndent & """test"": "
            If Not WhileTest Is Nothing Then
                json = json & WhileTest.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """body"": "
            If Not WhileBody Is Nothing Then
                json = json & WhileBody.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case DoWhileStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""DoWhileStatement""," & vbCrLf
            json = json & nextIndent & """body"": "
            If Not WhileBody Is Nothing Then
                json = json & WhileBody.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """test"": "
            If Not WhileTest Is Nothing Then
                json = json & WhileTest.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case ForStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""ForStatement""," & vbCrLf
            json = json & nextIndent & """init"": "
            If Not ForInit Is Nothing Then
                json = json & ForInit.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """test"": "
            If Not ForTest Is Nothing Then
                json = json & ForTest.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """update"": "
            If Not ForUpdate Is Nothing Then
                json = json & ForUpdate.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """body"": "
            If Not ForBody Is Nothing Then
                json = json & ForBody.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case ForInStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""ForInStatement""," & vbCrLf
            json = json & nextIndent & """left"": "
            If Not Left Is Nothing Then
                json = json & Left.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """right"": "
            If Not Right Is Nothing Then
                json = json & Right.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """body"": "
            If Not ForBody Is Nothing Then
                json = json & ForBody.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case ReturnStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""ReturnStatement""," & vbCrLf
            json = json & nextIndent & """argument"": "
            If Not ReturnArgument Is Nothing Then
                json = json & ReturnArgument.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case BreakStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""BreakStatement""," & vbCrLf
            json = json & nextIndent & """label"": null" & vbCrLf
            json = json & spaces & "}"
            
        Case ContinueStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""ContinueStatement""," & vbCrLf
            json = json & nextIndent & """label"": null" & vbCrLf
            json = json & spaces & "}"
            
        Case ThrowStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""ThrowStatement""," & vbCrLf
            json = json & nextIndent & """argument"": "
            If Not ReturnArgument Is Nothing Then
                json = json & ReturnArgument.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case TryStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""TryStatement""," & vbCrLf
            json = json & nextIndent & """block"": "
            If Not TryBlock Is Nothing Then
                json = json & TryBlock.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """handler"": "
            If Not TryHandler Is Nothing Then
                json = json & TryHandler.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """finalizer"": "
            If Not TryFinalizer Is Nothing Then
                json = json & TryFinalizer.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case CatchClause_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""CatchClause""," & vbCrLf
            json = json & nextIndent & """param"": "
            If Not param Is Nothing Then
                json = json & param.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """body"": "
            If Not CatchBody Is Nothing Then
                json = json & CatchBody.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case SwitchStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""SwitchStatement""," & vbCrLf
            json = json & nextIndent & """discriminant"": "
            If Not discriminant Is Nothing Then
                json = json & discriminant.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """cases"": ["
            
            If Not Cases Is Nothing Then
                If Cases.count > 0 Then
                    json = json & vbCrLf
                    For i = 1 To Cases.count
                        Dim caseNode As CNode
                        Set caseNode = Cases(i)
                        json = json & caseNode.ToJSON(Indent + 4)
                        If i < Cases.count Then json = json & ","
                        json = json & vbCrLf
                    Next i
                    json = json & nextIndent
                End If
            End If
            
            json = json & "]" & vbCrLf
            json = json & spaces & "}"
            
        Case SwitchCase_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""SwitchCase""," & vbCrLf
            json = json & nextIndent & """test"": "
            If Not CaseTest Is Nothing Then
                json = json & CaseTest.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """consequent"": ["
            
            If Not CaseConsequent Is Nothing Then
                If CaseConsequent.count > 0 Then
                    json = json & vbCrLf
                    For i = 1 To CaseConsequent.count
                        Set stmt = CaseConsequent(i)
                        json = json & stmt.ToJSON(Indent + 4)
                        If i < CaseConsequent.count Then json = json & ","
                        json = json & vbCrLf
                    Next i
                    json = json & nextIndent
                End If
            End If
            
            json = json & "]" & vbCrLf
            json = json & spaces & "}"
            
        Case EmptyStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""EmptyStatement""" & vbCrLf
            json = json & spaces & "}"
            
        Case WithStatement_Node
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""WithStatement""," & vbCrLf
            json = json & nextIndent & """object"": "
            If Not Object Is Nothing Then
                json = json & Object.ToJSON(Indent + 2)
            Else
                json = json & "null"
            End If
            json = json & "," & vbCrLf
            json = json & nextIndent & """body"": "
            If Not Body Is Nothing Then
                If Body.count > 0 Then
                    Set stmt = Body(1)
                    json = json & stmt.ToJSON(Indent + 2)
                Else
                    json = json & "null"
                End If
            Else
                json = json & "null"
            End If
            json = json & vbCrLf
            json = json & spaces & "}"
            
        Case Else
            ' Unknown node type
            json = spaces & "{" & vbCrLf
            json = json & nextIndent & """type"": ""Unknown""," & vbCrLf
            json = json & nextIndent & """nodeType"": " & tType & vbCrLf
            json = json & spaces & "}"
    End Select
    
    ToJSON = json
End Function

' Helper function to escape JSON strings
Private Function EscapeJSON(text As String) As String
    Dim result As String
    Dim i As Long
    Dim ch As String
    
    result = ""
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        Select Case ch
            Case """"
                result = result & "\"""
            Case "\"
                result = result & "\\"
            Case vbCr
                result = result & "\r"
            Case vbLf
                result = result & "\n"
            Case vbTab
                result = result & "\t"
            Case Else
                result = result & ch
        End Select
    Next i
    
    EscapeJSON = result
End Function

' Helper function to format literal values
Private Function FormatLiteralValue(val As Variant) As String
    If IsNull(val) Then
        FormatLiteralValue = "null"
    ElseIf VarType(val) = vbString Then
        FormatLiteralValue = """" & EscapeJSON(CStr(val)) & """"
    ElseIf VarType(val) = vbBoolean Then
        FormatLiteralValue = IIf(val, "true", "false")
    ElseIf IsNumeric(val) Then
        FormatLiteralValue = CStr(val)
    Else
        FormatLiteralValue = """" & EscapeJSON(CStr(val)) & """"
    End If
End Function


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBeautifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'====================================================================
' CBeautifier.cls - Professional JavaScript Code Beautifier
' "Make it gorgeous!"
'====================================================================

Option Explicit

' Formatting options
Public IndentSize As Long
Public IndentChar As String
Public MaxLineLength As Long
Public PreserveBlankLines As Boolean

' Spacing options
Public SpaceBeforeFunctionParen As Boolean    ' function foo () vs function foo()
Public SpaceInFunctionCall As Boolean         ' foo( x ) vs foo(x)
Public SpaceInParens As Boolean               ' ( x ) vs (x)
Public SpaceBeforeBinaryOp As Boolean         ' x + y vs x+y
Public SpaceAfterBinaryOp As Boolean
Public SpaceBeforeComma As Boolean            ' (a , b) vs (a, b)
Public SpaceAfterComma As Boolean

' Brace style
Public Enum BraceStyle
    Brace_1TBS = 1          ' One True Brace Style: function() {
    Brace_Allman = 2        ' Allman: function()\n{
    Brace_KR = 3            ' K&R (same as 1TBS for JS)
    Brace_Stroustrup = 4    ' Stroustrup: }\nelse {
End Enum

Public BraceStyleOption As BraceStyle

' Line break options
Public BlankLineBeforeFunction As Boolean
Public BlankLineAfterFunction As Boolean
Public BlankLineBeforeBlock As Boolean
Public BlankLineAfterBlock As Boolean

' Semicolon handling
Public Enum SemicolonStyle
    Semi_Always = 1         ' Always add semicolons
    Semi_Never = 2          ' Remove all semicolons (ASI)
    Semi_Preserve = 3       ' Keep as-is
End Enum

Public SemicolonStyleOption As SemicolonStyle

' Quote style
Public Enum QuoteStyle
    Quote_Double = 1        ' "string"
    Quote_Single = 2        ' 'string'
    Quote_Preserve = 3      ' Keep original
End Enum

Public QuoteStyleOption As QuoteStyle

' Internal state
Private m_currentIndent As Long
Private m_output As String

Private Sub Class_Initialize()
    ' Default settings (conservative)
    IndentSize = 2
    IndentChar = " "
    MaxLineLength = 80
    PreserveBlankLines = True
    
    SpaceBeforeFunctionParen = False
    SpaceInFunctionCall = False
    SpaceInParens = False
    SpaceBeforeBinaryOp = True
    SpaceAfterBinaryOp = True
    SpaceBeforeComma = False
    SpaceAfterComma = True
    
    BraceStyleOption = Brace_1TBS
    SemicolonStyleOption = Semi_Always
    QuoteStyleOption = Quote_Double
    
    BlankLineBeforeFunction = True
    BlankLineAfterFunction = False
    BlankLineBeforeBlock = False
    BlankLineAfterBlock = False
End Sub

' ====================================================================
' PRESET STYLES
' ====================================================================

Public Sub LoadPreset(presetName As String)
    Select Case LCase$(presetName)
        Case "airbnb"
            ' Airbnb JavaScript Style Guide
            IndentSize = 2
            SpaceBeforeFunctionParen = True
            SpaceBeforeBinaryOp = True
            SpaceAfterBinaryOp = True
            SpaceAfterComma = True
            BraceStyleOption = Brace_1TBS
            SemicolonStyleOption = Semi_Always
            QuoteStyleOption = Quote_Single
            BlankLineBeforeFunction = True
            
        Case "google"
            ' Google JavaScript Style Guide
            IndentSize = 2
            SpaceBeforeFunctionParen = False
            SpaceBeforeBinaryOp = True
            SpaceAfterBinaryOp = True
            SpaceAfterComma = True
            BraceStyleOption = Brace_1TBS
            SemicolonStyleOption = Semi_Always
            QuoteStyleOption = Quote_Single
            BlankLineBeforeFunction = False
            
        Case "standard"
            ' StandardJS
            IndentSize = 2
            SpaceBeforeFunctionParen = True
            SpaceBeforeBinaryOp = True
            SpaceAfterBinaryOp = True
            SpaceAfterComma = True
            BraceStyleOption = Brace_1TBS
            SemicolonStyleOption = Semi_Never
            QuoteStyleOption = Quote_Single
            BlankLineBeforeFunction = False
            
        Case "compact"
            ' Compact/minified-like (but readable)
            IndentSize = 0
            SpaceBeforeFunctionParen = False
            SpaceBeforeBinaryOp = False
            SpaceAfterBinaryOp = False
            SpaceAfterComma = True
            BraceStyleOption = Brace_1TBS
            SemicolonStyleOption = Semi_Always
            QuoteStyleOption = Quote_Double
            BlankLineBeforeFunction = False
            
        Case "allman", "bsd"
            ' Allman/BSD style (braces on new line)
            IndentSize = 4
            SpaceBeforeFunctionParen = False
            SpaceBeforeBinaryOp = True
            SpaceAfterBinaryOp = True
            SpaceAfterComma = True
            BraceStyleOption = Brace_Allman
            SemicolonStyleOption = Semi_Always
            QuoteStyleOption = Quote_Double
            BlankLineBeforeFunction = True
            
        Case "kr", "k&r"
            ' K&R style
            IndentSize = 4
            SpaceBeforeFunctionParen = False
            SpaceBeforeBinaryOp = True
            SpaceAfterBinaryOp = True
            SpaceAfterComma = True
            BraceStyleOption = Brace_KR
            SemicolonStyleOption = Semi_Always
            QuoteStyleOption = Quote_Double
            BlankLineBeforeFunction = False
    End Select
End Sub

' ====================================================================
' MAIN BEAUTIFY FUNCTION
' ====================================================================

Public Function BeautifyStr(ByVal script As String) As String
    On Error Resume Next
    Dim parser As New CParser, ast As CNode
    Set ast = parser.ParseScript(script)
    BeautifyStr = Beautify(ast)
End Function

Public Function Beautify(ast As CNode) As String
    m_currentIndent = 0
    m_output = ""
    
    ' Emit the AST with beautification
    EmitNode ast
    
    ' Post-process
    Beautify = PostProcess(m_output)
End Function

Private Function PostProcess(code As String) As String
    Dim result As String
    result = code
    
    ' Apply quote style conversion
    If QuoteStyleOption <> Quote_Preserve Then
        result = ConvertQuotes(result)
    End If
    
    ' Remove excessive blank lines
    result = RemoveExcessiveBlankLines(result)
    
    ' Trim trailing whitespace
    result = TrimTrailingWhitespace(result)
    
    PostProcess = result
End Function

' ====================================================================
' AST EMISSION WITH BEAUTIFICATION
' ====================================================================

Private Sub EmitNode(Node As CNode)
    If Node Is Nothing Then Exit Sub
    
    Select Case Node.tType
        Case Program_Node
            EmitProgram Node
            
        Case FunctionDeclaration_Node
            EmitFunctionDeclaration Node
            
        Case FunctionExpression_Node
            EmitFunctionExpression Node
            
        Case VariableDeclaration_Node
            EmitVariableDeclaration Node
            
        Case BlockStatement_Node
            EmitBlockStatement Node
            
        Case ExpressionStatement_Node
            EmitExpressionStatement Node
            
        Case ReturnStatement_Node
            EmitReturnStatement Node
            
        Case IfStatement_Node
            EmitIfStatement Node
            
        Case WhileStatement_Node
            EmitWhileStatement Node
            
        Case ForStatement_Node
            EmitForStatement Node
            
        Case SwitchStatement_Node
            EmitSwitchStatement Node
            
        Case BreakStatement_Node
            Indent
            Append "break"
            If SemicolonStyleOption = Semi_Always Then Append ";"
            AppendLine ""
            
        Case ContinueStatement_Node
            Indent
            Append "continue"
            If SemicolonStyleOption = Semi_Always Then Append ";"
            AppendLine ""
            
        Case BinaryExpression_Node, LogicalExpression_Node
            EmitBinaryExpression Node
            
        Case UnaryExpression_Node
            EmitUnaryExpression Node
            
        Case UpdateExpression_Node
            EmitUpdateExpression Node
            
        Case AssignmentExpression_Node
            EmitAssignmentExpression Node
            
        Case CallExpression_Node
            EmitCallExpression Node
            
        Case MemberExpression_Node
            EmitMemberExpression Node
            
        Case Identifier_Node
            Append Node.Name
            
        Case Literal_Node
            EmitLiteral Node
            
        Case ArrayExpression_Node
            EmitArrayExpression Node
            
        Case ObjectExpression_Node
            EmitObjectExpression Node
    End Select
End Sub

Private Sub EmitProgram(Node As CNode)
    Dim i As Long
    For i = 1 To Node.Body.count
        Dim stmt As CNode
        Set stmt = Node.Body(i)
        
        ' Add blank line before functions
        If BlankLineBeforeFunction And stmt.tType = FunctionDeclaration_Node Then
            If i > 1 Then AppendLine ""
        End If
        
        EmitNode stmt
        
        ' Add blank line after functions
        If BlankLineAfterFunction And stmt.tType = FunctionDeclaration_Node Then
            If i < Node.Body.count Then AppendLine ""
        End If
    Next
End Sub

Private Sub EmitFunctionDeclaration(Node As CNode)
    Indent
    Append "function"
    
    If Not Node.id Is Nothing Then
        Append " "
        Append Node.id.Name
    End If
    
    ' Then handle paren spacing
    If SpaceBeforeFunctionParen Then Append " "
    
    Append "("
    EmitParameters Node.Params
    Append ")"
    
    ' Brace style
    Select Case BraceStyleOption
        Case Brace_1TBS, Brace_KR
            Append " {"
            AppendLine ""
            
        Case Brace_Allman
            AppendLine ""
            Indent
            Append "{"
            AppendLine ""
    End Select
    
    ' Function body
    m_currentIndent = m_currentIndent + 1
    EmitNode Node.FunctionBody
    m_currentIndent = m_currentIndent - 1
    
    ' Closing brace
    Indent
    Append "}"
    AppendLine ""
End Sub

Private Sub EmitFunctionExpression(Node As CNode)
    Append "function"
    
    If Not Node.id Is Nothing Then
        If SpaceBeforeFunctionParen Then Append " "
        Append Node.id.Name
    End If
    
    Append "("
    EmitParameters Node.Params
    Append ")"
    
    Append " {"
    AppendLine ""
    
    m_currentIndent = m_currentIndent + 1
    EmitNode Node.FunctionBody
    m_currentIndent = m_currentIndent - 1
    
    Indent
    Append "}"
End Sub

Private Sub EmitParameters(Params As Collection)
    If Params Is Nothing Then Exit Sub
    
    Dim i As Long
    For i = 1 To Params.count
        Dim param As CNode
        Set param = Params(i)
        
        If SpaceInParens And i = 1 Then Append " "
        
        Append param.Name
        
        If i < Params.count Then
            If SpaceBeforeComma Then Append " "
            Append ","
            If SpaceAfterComma Then Append " "
        End If
        
        If SpaceInParens And i = Params.count Then Append " "
    Next
End Sub

Private Sub EmitVariableDeclaration(Node As CNode)
    Indent
    Append Node.Kind & " "
    
    Dim i As Long
    For i = 1 To Node.Declarations.count
        Dim decl As CNode
        Set decl = Node.Declarations(i)
        
        Append decl.VarID.Name
        
        If Not decl.Init Is Nothing Then
            If SpaceBeforeBinaryOp Then Append " "
            Append "="
            If SpaceAfterBinaryOp Then Append " "
            EmitNode decl.Init
        End If
        
        If i < Node.Declarations.count Then
            Append ","
            If SpaceAfterComma Then Append " "
        End If
    Next
    
    ' Semicolon
    If SemicolonStyleOption = Semi_Always Then
        Append ";"
    End If
    
    AppendLine ""
End Sub

Private Sub EmitBlockStatement(Node As CNode)
    ' Body is already indented by parent
    Dim i As Long
    For i = 1 To Node.Body.count
        EmitNode Node.Body(i)
    Next
End Sub

Private Sub EmitExpressionStatement(Node As CNode)
    Indent
    EmitNode Node.Test
    
    If SemicolonStyleOption = Semi_Always Then
        Append ";"
    End If
    
    AppendLine ""
End Sub

Private Sub EmitReturnStatement(Node As CNode)
    Indent
    Append "return"
    
    If Not Node.ReturnArgument Is Nothing Then
        Append " "
        EmitNode Node.ReturnArgument
    End If
    
    If SemicolonStyleOption = Semi_Always Then
        Append ";"
    End If
    
    AppendLine ""
End Sub

Private Sub EmitIfStatement(Node As CNode)
    Indent
    Append "if"
    
    If SpaceBeforeFunctionParen Then Append " "
    
    Append "("
    If SpaceInParens Then Append " "
    EmitNode Node.IfTest
    If SpaceInParens Then Append " "
    Append ")"
    
    ' Consequent
    If Node.IfConsequent.tType = BlockStatement_Node Then
        Append " {"
        AppendLine ""
        m_currentIndent = m_currentIndent + 1
        EmitNode Node.IfConsequent
        m_currentIndent = m_currentIndent - 1
        Indent
        Append "}"
    Else
        AppendLine ""
        m_currentIndent = m_currentIndent + 1
        EmitNode Node.IfConsequent
        m_currentIndent = m_currentIndent - 1
    End If
    
    ' Alternate
    If Not Node.IfAlternate Is Nothing Then
        Append " else"
        
        If Node.IfAlternate.tType = IfStatement_Node Then
            Append " "
            ' Remove indent for else-if
            Dim savedIndent As Long
            savedIndent = m_currentIndent
            m_currentIndent = 0
            EmitNode Node.IfAlternate
            m_currentIndent = savedIndent
        ElseIf Node.IfAlternate.tType = BlockStatement_Node Then
            Append " {"
            AppendLine ""
            m_currentIndent = m_currentIndent + 1
            EmitNode Node.IfAlternate
            m_currentIndent = m_currentIndent - 1
            Indent
            Append "}"
            AppendLine ""
        Else
            AppendLine ""
            m_currentIndent = m_currentIndent + 1
            EmitNode Node.IfAlternate
            m_currentIndent = m_currentIndent - 1
        End If
    Else
        AppendLine ""
    End If
End Sub

Private Sub EmitWhileStatement(Node As CNode)
    Indent
    Append "while"
    
    If SpaceBeforeFunctionParen Then Append " "
    
    Append "("
    If SpaceInParens Then Append " "
    EmitNode Node.WhileTest
    If SpaceInParens Then Append " "
    Append ") {"
    AppendLine ""
    
    m_currentIndent = m_currentIndent + 1
    EmitNode Node.WhileBody
    m_currentIndent = m_currentIndent - 1
    
    Indent
    Append "}"
    AppendLine ""
End Sub

Private Sub EmitForStatement(Node As CNode)
    Indent
    Append "for"
    
    If SpaceBeforeFunctionParen Then Append " "
    
    Append "("
    
    ' Init
    If Not Node.ForInit Is Nothing Then
        ' Don't add newline for init
        Dim savedIndent As Long
        savedIndent = m_currentIndent
        m_currentIndent = 0
        EmitNode Node.ForInit
        m_currentIndent = savedIndent
        ' Remove the trailing newline that EmitNode added
        m_output = Left$(m_output, Len(m_output) - Len(vbCrLf))
    End If
    
    Append ";"
    If SpaceAfterComma Then Append " "
    
    ' Test
    If Not Node.ForTest Is Nothing Then
        EmitNode Node.ForTest
    End If
    
    Append ";"
    If SpaceAfterComma Then Append " "
    
    ' Update
    If Not Node.ForUpdate Is Nothing Then
        EmitNode Node.ForUpdate
    End If
    
    Append ") {"
    AppendLine ""
    
    m_currentIndent = m_currentIndent + 1
    EmitNode Node.ForBody
    m_currentIndent = m_currentIndent - 1
    
    Indent
    Append "}"
    AppendLine ""
End Sub

Private Sub EmitBinaryExpression(Node As CNode)
    EmitNode Node.Left
    
    If SpaceBeforeBinaryOp Then Append " "
    Append Node.operator
    If SpaceAfterBinaryOp Then Append " "
    
    EmitNode Node.Right
End Sub

Private Sub EmitUnaryExpression(Node As CNode)
    Append Node.operator
    
    ' No space for operators like -, +, !
    ' Space for 'typeof', 'delete', etc.
    If Len(Node.operator) > 1 Then Append " "
    
    EmitNode Node.Argument
End Sub

Private Sub EmitUpdateExpression(Node As CNode)
    If Node.Prefix Then
        Append Node.operator
        EmitNode Node.Argument
    Else
        EmitNode Node.Argument
        Append Node.operator
    End If
End Sub

Private Sub EmitAssignmentExpression(Node As CNode)
    EmitNode Node.Left
    
    If SpaceBeforeBinaryOp Then Append " "
    Append Node.operator
    If SpaceAfterBinaryOp Then Append " "
    
    EmitNode Node.Right
End Sub

Private Sub EmitCallExpression(Node As CNode)
    EmitNode Node.Callee
    
    Append "("
    
    Dim i As Long
    For i = 1 To Node.Arguments.count
        If SpaceInFunctionCall And i = 1 Then Append " "
        
        EmitNode Node.Arguments(i)
        
        If i < Node.Arguments.count Then
            If SpaceBeforeComma Then Append " "
            Append ","
            If SpaceAfterComma Then Append " "
        End If
        
        If SpaceInFunctionCall And i = Node.Arguments.count Then Append " "
    Next
    
    Append ")"
End Sub

Private Sub EmitMemberExpression(Node As CNode)
    EmitNode Node.Object
    
    If Node.Computed Then
        Append "["
        EmitNode Node.prop
        Append "]"
    Else
        Append "."
        EmitNode Node.prop
    End If
End Sub

Private Sub EmitLiteral(Node As CNode)
    If VarType(Node.value) = vbString Then
        ' Handle string with quote style
        Dim quote As String
        Select Case QuoteStyleOption
            Case Quote_Double: quote = """"
            Case Quote_Single: quote = "'"
            Case Quote_Preserve
                ' Use original
                Append Node.Raw
                Exit Sub
        End Select
        
        Append quote & CStr(Node.value) & quote
    Else
        Append CStr(Node.value)
    End If
End Sub

Private Sub EmitArrayExpression(Node As CNode)
    Append "["
    
    Dim i As Long
    For i = 1 To Node.Body.count
        If SpaceInParens And i = 1 Then Append " "
        
        EmitNode Node.Body(i)
        
        If i < Node.Body.count Then
            If SpaceBeforeComma Then Append " "
            Append ","
            If SpaceAfterComma Then Append " "
        End If
        
        If SpaceInParens And i = Node.Body.count Then Append " "
    Next
    
    Append "]"
End Sub

Private Sub EmitObjectExpression(Node As CNode)
    Append "{"
    
    If Node.Properties.count > 0 Then
        AppendLine ""
        m_currentIndent = m_currentIndent + 1
        
        Dim i As Long
        For i = 1 To Node.Properties.count
            Indent
            
            Dim prop As CNode
            Set prop = Node.Properties(i)
            
            ' Key
            EmitNode prop.PropKey
            
            Append ":"
            If SpaceAfterComma Then Append " "
            
            ' Value
            EmitNode prop.propValue  ' ? Fixed!
            
            If i < Node.Properties.count Then
                Append ","
            End If
            
            AppendLine ""
        Next
        
        m_currentIndent = m_currentIndent - 1
        Indent
    End If
    
    Append "}"
End Sub

Private Sub EmitSwitchStatement(Node As CNode)
    Indent
    Append "switch"
    Append " "
    Append "("
    EmitNode Node.discriminant
    Append ")"
    Append " "
    Append "{"
    AppendLine ""
    
    m_currentIndent = m_currentIndent + 1
    
    Dim i As Long
    For i = 1 To Node.Cases.count
        Dim switchCase As CNode
        Set switchCase = Node.Cases(i)
        EmitSwitchCase switchCase
    Next i
    
    m_currentIndent = m_currentIndent - 1
    
    Indent
    Append "}"
    AppendLine ""
End Sub

Private Sub EmitSwitchCase(Node As CNode)
    Indent
    
    If Node.CaseTest Is Nothing Then
        Append "default"
    Else
        Append "case"
        Append " "
        EmitNode Node.CaseTest
    End If
    
    Append ":"
    AppendLine ""
    
    ' Emit consequent statements
    ' Each EmitNode handles its own indentation, semicolons, and newlines
    If Node.CaseConsequent.count > 0 Then
        m_currentIndent = m_currentIndent + 1
        
        Dim i As Long
        For i = 1 To Node.CaseConsequent.count
            Dim stmt As CNode
            Set stmt = Node.CaseConsequent(i)
            EmitNode stmt
        Next i
        
        m_currentIndent = m_currentIndent - 1
    End If
End Sub

' ====================================================================
' OUTPUT HELPERS
' ====================================================================

Private Sub Indent()
    Dim i As Long
    For i = 1 To m_currentIndent * IndentSize
        m_output = m_output & IndentChar
    Next
End Sub

Private Sub Append(text As String)
    m_output = m_output & text
End Sub

Private Sub AppendLine(text As String)
    m_output = m_output & text & vbCrLf
End Sub

' ====================================================================
' POST-PROCESSING
' ====================================================================

Private Function ConvertQuotes(code As String) As String
    ' Simple quote conversion (naive - doesn't handle escapes perfectly)
    Dim result As String
    result = code
    
    Select Case QuoteStyleOption
        Case Quote_Double
            ' Convert 'string' to "string"
            ' TODO: Handle escaped quotes properly
            
        Case Quote_Single
            ' Convert "string" to 'string'
            ' TODO: Handle escaped quotes properly
    End Select
    
    ConvertQuotes = result
End Function

Private Function RemoveExcessiveBlankLines(code As String) As String
    Dim result As String
    result = code
    
    ' Replace 3+ blank lines with 2
    Do While InStr(result, vbCrLf & vbCrLf & vbCrLf & vbCrLf) > 0
        result = Replace(result, vbCrLf & vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf & vbCrLf)
    Loop
    
    RemoveExcessiveBlankLines = result
End Function

Private Function TrimTrailingWhitespace(code As String) As String
    Dim lines() As String
    lines = Split(code, vbCrLf)
    
    Dim i As Long
    For i = 0 To UBound(lines)
        lines(i) = RTrim$(lines(i))
    Next
    
    TrimTrailingWhitespace = Join(lines, vbCrLf)
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com> + Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'todo: barfing on parsing regex literals, and named for loops like outer: for(){ (see es5.txt


'------------------------------------------------------------
' CParser.cls - JavaScript Recursive Descent Parser
'
' "To infinity and beyond!" - Still Buzz
'
' This is where the magic happens. Tokens ? AST.
' Recursive descent parsing. In VB6. We're INSANE.
'------------------------------------------------------------

Option Explicit

Private scanner As CScanner
Private lookahead As CToken
Private lastToken As CToken

Private Sub Class_Initialize()
    Set scanner = New CScanner
End Sub

Public Function ParseScript(source As String) As CNode
    On Error GoTo ErrorHandler
    
    scanner.init source
    Set lookahead = scanner.Lex()
    
    Dim program As New CNode
    Dim stmt As CNode
    
    ' Force initialization by accessing a property first
    program.LineNumber = 1
    program.tType = Program_Node
    
    ' Parse all top-level statements
    Do While lookahead.tType <> EOF_TOKEN
        
        Set stmt = ParseStatementListItem()
        
        If stmt Is Nothing Then
            Err.Raise vbObjectError + 2001, "CParser", "ParseStatementListItem returned Nothing"
        End If
        
        program.Body.add stmt
    Loop
    
    Set ParseScript = program
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR in ParseScript: " & Err.Description
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Current token: " & lookahead.Value & " (type " & lookahead.tType & ")"
    Err.Raise Err.Number, Err.source, Err.Description
End Function

' ============================================
' Token Management
' ============================================

Private Function Expect(tokenValue As String) As CToken
    Dim token As CToken
    Set token = lookahead
    
    If lookahead.tType <> Punctuator And lookahead.tType <> Keyword Then
        ThrowUnexpectedToken lookahead
    End If
    
    If lookahead.Value <> tokenValue Then
        ThrowUnexpectedToken lookahead
    End If
    
    Set lastToken = lookahead
    Set lookahead = scanner.Lex()
    
    Set Expect = token
End Function

Private Function match(tokenValue As String) As Boolean
    match = (lookahead.tType = Punctuator Or lookahead.tType = Keyword) _
            And lookahead.Value = tokenValue
End Function

Private Function MatchKeyword(kw As String) As Boolean
    MatchKeyword = (lookahead.tType = Keyword) And (lookahead.Value = kw)
End Function

Private Sub ThrowUnexpectedToken(token As CToken)
    Dim column As Long
    column = token.Start - token.LineStart + 1  ' +1 for 1-based column
    Err.Raise vbObjectError + 2000, "CParser", _
              "Unexpected token: " & token.Value & " at line " & token.LineNumber & ", column " & column
End Sub

' ============================================
' Statement Parsing
' ============================================

Private Function ParseStatementListItem() As CNode
    On Error GoTo ErrorHandler
    
    If lookahead.tType = Keyword Then
        Select Case lookahead.Value
            Case "function"
                Set ParseStatementListItem = ParseFunctionDeclaration()
            Case Else
                Set ParseStatementListItem = ParseStatement()
        End Select
    Else
        Set ParseStatementListItem = ParseStatement()
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR in ParseStatementListItem"
    Debug.Print "  Token: '" & lookahead.Value & "' (type " & lookahead.tType & ")"
    Err.Raise Err.Number, "ParseStatementListItem", Err.Description
End Function

Private Function ParseStatement() As CNode
    On Error GoTo ErrorHandler
    
    Dim stmt As CNode
    Dim startLine As Long
    
    startLine = lookahead.LineNumber  ' CAPTURE LINE NUMBER
    
    Select Case lookahead.tType
        Case Punctuator
            Select Case lookahead.Value
                Case ";"
                    Set stmt = ParseEmptyStatement()
                Case "{"
                    Set stmt = ParseBlockStatement()
                Case Else
                    Set stmt = ParseExpressionStatement()
            End Select
            
        Case Keyword
            Select Case lookahead.Value
                Case "var"
                    Set stmt = ParseVariableStatement()
                Case "if"
                    Set stmt = ParseIfStatement()
                Case "while"
                    Set stmt = ParseWhileStatement()
                Case "for"
                    Set stmt = ParseForStatement()
                Case "return"
                    Set stmt = ParseReturnStatement()
                Case "break"
                    Set stmt = ParseBreakStatement()
                Case "continue"
                    Set stmt = ParseContinueStatement()
                Case "throw"
                    Set stmt = ParseThrowStatement()
                Case "try"
                    Set stmt = ParseTryStatement()
                Case "switch"
                    Set stmt = ParseSwitchStatement()
                Case "do"
                    Set stmt = ParseDoWhileStatement()
                Case "debugger"
                    Set stmt = ParseDebuggerStatement()  ' ADD THIS
                Case "with"
                    Set stmt = ParseWithStatement()
                Case Else
                    Set stmt = ParseExpressionStatement()
            End Select
            
        Case Else
            Set stmt = ParseExpressionStatement()
    End Select
    
    ' SET LINE NUMBER ON NODE
    If Not stmt Is Nothing Then
        stmt.LineNumber = startLine
    End If
    
    Set ParseStatement = stmt
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR in ParseStatement"
    Debug.Print "  Token: '" & lookahead.Value & "' (type " & lookahead.tType & ")"
    Err.Raise Err.Number, "ParseStatement", Err.Description
End Function

Private Function ParseWithStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line of identifier
    
    Expect "with"
    Expect "("
    
    Dim Node As New CNode
    Node.tType = EmptyStatement_Node  ' Just skip for navigation
    Node.LineNumber = startLine
    
    ' Parse the object expression (but ignore it)
    Dim obj As CNode
    Set obj = ParseExpression()
    
    Expect ")"
    
    ' Parse the body (but ignore it for navigation)
    Dim Body As CNode
    Set Body = ParseStatement()
    
    ' Just return empty - 'with' is bad practice anyway
    Set ParseWithStatement = Node
End Function

Private Function ParseDebuggerStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line of identifier
    
    Expect "debugger"
    If match(";") Then Expect ";"
    
    Dim Node As New CNode
    Node.tType = EmptyStatement_Node  ' Just treat as empty for navigation
    Node.LineNumber = startLine
    Set ParseDebuggerStatement = Node
End Function

Private Function ParseEmptyStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line of identifier
    
    Expect ";"
    
    Dim Node As New CNode
    Node.tType = EmptyStatement_Node
    Node.LineNumber = startLine
    Set ParseEmptyStatement = Node
End Function

Private Function ParseBlockStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line of identifier
    
    Expect "{"
    
    Dim Node As New CNode
    Node.tType = BlockStatement_Node
    Node.LineNumber = startLine
    
    Do While Not match("}")
        Node.Body.add ParseStatementListItem()
    Loop
    
    Expect "}"
    
    Set ParseBlockStatement = Node
End Function

Private Function ParseExpressionStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line of identifier
    
    Dim expr As CNode
    Set expr = ParseExpression()
    
    ' Consume optional semicolon
    If match(";") Then
        Expect ";"
    End If
    
    Dim Node As New CNode
    Node.tType = ExpressionStatement_Node
    Node.LineNumber = startLine
    Set Node.Test = expr ' Reuse Test field for expression
    
    Set ParseExpressionStatement = Node
End Function

Private Function ParseVariableStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line of identifier
    
    Expect "var"
    
    Dim Node As New CNode
    Node.tType = VariableDeclaration_Node
    Node.Kind = "var"
    Node.LineNumber = startLine
    
    ' Parse declarators
    Do
        Node.Declarations.add ParseVariableDeclarator()
        
        If Not match(",") Then Exit Do
        Expect ","
    Loop
    
    ' Consume optional semicolon
    If match(";") Then
        Expect ";"
    End If
    
    Set ParseVariableStatement = Node
End Function

Private Function ParseVariableDeclarator() As CNode

    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line of identifier
    
    Dim Node As New CNode
    Node.tType = VariableDeclarator_Node
    Node.LineNumber = startLine
    
    ' Parse identifier
    If lookahead.tType <> Identifier Then
        ThrowUnexpectedToken lookahead
    End If
    
    Set Node.VarID = ParsePrimaryExpression() ' Gets identifier
    
    ' Parse initializer if present
    If match("=") Then
        Expect "="
        Set Node.init = ParseAssignmentExpression()
    End If
    
    Set ParseVariableDeclarator = Node
End Function

Private Function ParseIfStatement() As CNode
    Dim ln As Long
    ln = lookahead.LineNumber
    
    Expect "if"
    Expect "("
    
    Dim Node As New CNode
    Node.tType = IfStatement_Node
    Node.LineNumber = ln
    Set Node.IfTest = ParseExpression()
    
    Expect ")"
    
    Set Node.IfConsequent = ParseStatement()
    
    ' Check for else clause
    If MatchKeyword("else") Then
        Expect "else"
        Set Node.IfAlternate = ParseStatement()
    End If
    
    Set ParseIfStatement = Node
End Function

Private Function ParseWhileStatement() As CNode
    Dim ln As Long
    ln = lookahead.LineNumber
    
    Expect "while"
    Expect "("
    
    Dim Node As New CNode
    Node.tType = WhileStatement_Node
    Node.LineNumber = ln
    Set Node.WhileTest = ParseExpression()
    
    Expect ")"
    
    Set Node.WhileBody = ParseStatement()
    
    Set ParseWhileStatement = Node
End Function

Private Function ParseDoWhileStatement() As CNode
    Dim ln As Long
    ln = lookahead.LineNumber
    
    Expect "do"
    
    Dim Node As New CNode
    Node.tType = DoWhileStatement_Node
    Node.LineNumber = ln
    Set Node.WhileBody = ParseStatement()
    
    Expect "while"
    Expect "("
    Set Node.WhileTest = ParseExpression()
    Expect ")"
    
    If match(";") Then Expect ";"
    
    Set ParseDoWhileStatement = Node
End Function

Private Function ParseForStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line of "for"
    
    Expect "for"
    Expect "("
    
    Dim Node As New CNode
    Dim varDecl As CNode
    Dim declNode As CNode
    Dim seqNode As CNode
    
    If MatchKeyword("var") Then
        Dim varStartLine As Long
        varStartLine = lookahead.LineNumber  ' Capture line of "var"
        
        ' Parse var declaration (without semicolon)
        Expect "var"
        
        Set varDecl = New CNode
        varDecl.tType = VariableDeclaration_Node
        varDecl.Kind = "var"
        varDecl.LineNumber = varStartLine
        
        ' Parse the declarator
        If lookahead.tType <> Identifier Then
            ThrowUnexpectedToken lookahead
        End If
        
        Dim declStartLine As Long
        declStartLine = lookahead.LineNumber  ' Capture line of identifier
        
        Set declNode = New CNode
        declNode.tType = VariableDeclarator_Node
        declNode.LineNumber = declStartLine
        Set declNode.VarID = ParsePrimaryExpression()
        
        ' Now check what comes next
        If MatchKeyword("in") Then
            ' for-in with var
            varDecl.Declarations.add declNode
            
            Node.tType = ForInStatement_Node
            Node.LineNumber = startLine
            Set Node.Left = varDecl
            Expect "in"
            Set Node.Right = ParseExpression()
            Expect ")"
            Set Node.ForBody = ParseStatement()
            Set ParseForStatement = Node
            Exit Function
        ElseIf match("=") Then
            ' Has initializer, must be regular for
            Expect "="
            Set declNode.init = ParseAssignmentExpression()
            varDecl.Declarations.add declNode
            
            ' Continue with regular for loop logic
            Do While match(",")
                Expect ","
                varDecl.Declarations.add ParseVariableDeclarator()
            Loop
            
            Set Node.ForInit = varDecl
        Else
            ' No initializer, no 'in' - regular for with just declaration
            varDecl.Declarations.add declNode
            
            Do While match(",")
                Expect ","
                varDecl.Declarations.add ParseVariableDeclarator()
            Loop
            
            Set Node.ForInit = varDecl
        End If
        
    ElseIf Not match(";") Then
        ' Could be:
        ' - for (i = 0, j = 1; ...) - comma-separated assignments
        ' - for (i in obj) - for-in loop
        ' - for (expr; ...) - single expression
        
        ' Parse first expression - use UnaryExpression to avoid consuming 'in'
        Dim firstExpr As CNode
        Set firstExpr = ParseUnaryExpression()  ' Stops before 'in' operator
        
        ' Check if this is for-in
        If MatchKeyword("in") Then
            ' for-in without var
            Node.tType = ForInStatement_Node
            Node.LineNumber = startLine
            Set Node.Left = firstExpr
            Expect "in"
            Set Node.Right = ParseExpression()
            Expect ")"
            Set Node.ForBody = ParseStatement()
            Set ParseForStatement = Node
            Exit Function
        End If
        
        ' Not for-in - we may have only parsed part of the expression
        ' We need to continue parsing to get the full expression
        
        ' Check for assignment operators
        If lookahead.tType = Punctuator And IsAssignmentOperator(lookahead.Value) Then
            Dim assignStartLine As Long
            assignStartLine = firstExpr.LineNumber  ' Use line from left expr
            
            Dim op As String
            op = lookahead.Value
            Expect op
            
            Dim assignNode As New CNode
            assignNode.tType = AssignmentExpression_Node
            assignNode.Operator = op
            assignNode.LineNumber = assignStartLine
            Set assignNode.Left = firstExpr
            Set assignNode.Right = ParseAssignmentExpression()
            Set firstExpr = assignNode
        End If
        
        ' Check for comma-separated expressions
        If match(",") Then
            Dim seqStartLine As Long
            seqStartLine = firstExpr.LineNumber  ' Use line from first expr
            
            ' Multiple expressions: i = 0, j = 1, ...
            Dim exprs As New Collection
            exprs.add firstExpr
            
            Do While match(",")
                Expect ","
                exprs.add ParseAssignmentExpression()
            Loop
            
            ' Create sequence expression
            Set seqNode = New CNode
            seqNode.tType = SequenceExpression_Node
            seqNode.LineNumber = seqStartLine
            Set seqNode.Body = exprs
            Set Node.ForInit = seqNode
        Else
            ' Single expression
            Set Node.ForInit = firstExpr
        End If
    End If
    
    Expect ";"
    
    If Not match(";") Then
        Set Node.ForTest = ParseExpression()
    End If
    
    Expect ";"
    
    If Not match(")") Then
        Set Node.ForUpdate = ParseExpression()
    End If
    
    Expect ")"
    
    Node.tType = ForStatement_Node
    Node.LineNumber = startLine  ' Set line for for statement
    Set Node.ForBody = ParseStatement()
    
    Set ParseForStatement = Node
End Function

Private Function ParseReturnStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber
    
    Expect "return"
    
    Dim Node As New CNode
    Node.tType = ReturnStatement_Node
    Node.LineNumber = startLine
    
    ' Check for return value
    If Not match(";") And lookahead.tType <> EOF_TOKEN Then
        Set Node.ReturnArgument = ParseExpression()
    End If
    
    If match(";") Then Expect ";"
    
    Set ParseReturnStatement = Node
End Function

Private Function ParseBreakStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber
    Expect "break"
    If match(";") Then Expect ";"
    
    Dim Node As New CNode
    Node.tType = BreakStatement_Node
    Node.LineNumber = startLine
    Set ParseBreakStatement = Node
End Function

Private Function ParseContinueStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber
    Expect "continue"
    If match(";") Then Expect ";"
    
    Dim Node As New CNode
    Node.tType = ContinueStatement_Node
    Node.LineNumber = startLine
    Set ParseContinueStatement = Node
End Function

Private Function ParseThrowStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber
    
    Expect "throw"
    
    Dim Node As New CNode
    Node.tType = ThrowStatement_Node
    Node.LineNumber = startLine
    Set Node.ReturnArgument = ParseExpression()
    
    If match(";") Then Expect ";"
    
    Set ParseThrowStatement = Node
End Function

Private Function ParseTryStatement() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line of "try"
    
    Expect "try"
    
    Dim Node As New CNode
    Node.tType = TryStatement_Node
    Node.LineNumber = startLine       ' Set line for try statement
    Set Node.TryBlock = ParseBlockStatement()
    
    If MatchKeyword("catch") Then
        Dim catchStartLine As Long
        catchStartLine = lookahead.LineNumber  ' Capture line of "catch"
        
        Expect "catch"
        Expect "("
        
        Dim handler As New CNode
        handler.tType = CatchClause_Node
        handler.LineNumber = catchStartLine    ' Set line for catch clause
        Set handler.param = ParsePrimaryExpression()
        
        Expect ")"
        Set handler.CatchBody = ParseBlockStatement()
        Set Node.TryHandler = handler
    End If
    
    If MatchKeyword("finally") Then
        Expect "finally"
        Set Node.TryFinalizer = ParseBlockStatement()
    End If
    
    Set ParseTryStatement = Node
End Function

Private Function ParseSwitchStatement() As CNode
    
    Dim startLine As Long
    startLine = lookahead.LineNumber
    
    Expect "switch"
    Expect "("
    
    Dim Node As New CNode
    Dim switchCase As New CNode
    
    Node.tType = SwitchStatement_Node
    Node.LineNumber = startLine       ' Set line for switch statement
    
    Set Node.discriminant = ParseExpression()
    
    Expect ")"
    Expect "{"
    
    Do While Not match("}")
    
        Dim caseStartLine As Long
        caseStartLine = lookahead.LineNumber  ' Capture line of "case" or "default"

        Set switchCase = New CNode
        switchCase.tType = SwitchCase_Node
        switchCase.LineNumber = caseStartLine  ' Set line for each case
   
        If MatchKeyword("default") Then
            Expect "default"
            Expect ":"
            ' Test is Nothing for default
        Else
            Expect "case"
            Set switchCase.CaseTest = ParseExpression()
            Expect ":"
        End If
        
        ' Parse consequent statements
        Do While Not match("}") And Not MatchKeyword("case") And Not MatchKeyword("default")
            switchCase.CaseConsequent.add ParseStatementListItem()
        Loop
        
        Node.Cases.add switchCase
    Loop
    
    Expect "}"
    
    Set ParseSwitchStatement = Node
End Function

' ============================================
' Expression Parsing - The Fun Part!
' ============================================

Private Function ParseExpression() As CNode
    ' Sequence expression (comma operator)
    Dim expr As CNode, startLine As Long
    Set expr = ParseAssignmentExpression()
    
    If match(",") Then
        Dim exprs As New Collection
        exprs.add expr
        startLine = lookahead.LineNumber
    
        Do While match(",")
            Expect ","
            exprs.add ParseAssignmentExpression()
        Loop
        
        Dim Node As New CNode
        Node.tType = SequenceExpression_Node
        Node.LineNumber = startLine
        Set Node.Body = exprs
        Set expr = Node
    End If
    
    Set ParseExpression = expr
End Function

Private Function ParseAssignmentExpression() As CNode
    Dim expr As CNode, startLine As Long
    Set expr = ParseConditionalExpression()
    
    ' Check for assignment operators
    If lookahead.tType = Punctuator Then
        If IsAssignmentOperator(lookahead.Value) Then
            Dim op As String
            op = lookahead.Value
            startLine = lookahead.LineNumber
    
            Expect op
            
            Dim Node As New CNode
            Node.tType = AssignmentExpression_Node
            Node.Operator = op
            Node.LineNumber = startLine
            Set Node.Left = expr
            Set Node.Right = ParseAssignmentExpression()
            Set expr = Node
        End If
    End If
    
    Set ParseAssignmentExpression = expr
End Function

Private Function IsAssignmentOperator(op As String) As Boolean
    IsAssignmentOperator = (op = "=" Or op = "+=" Or op = "-=" Or _
                           op = "*=" Or op = "/=" Or op = "%=" Or _
                           op = "<<=" Or op = ">>=" Or op = ">>>=" Or _
                           op = "&=" Or op = "|=" Or op = "^=")
End Function

Private Function ParseConditionalExpression() As CNode
    Dim expr As CNode, startLine As Long
    Set expr = ParseLogicalORExpression()
    
    ' Ternary operator: test ? consequent : alternate
    If match("?") Then
        startLine = lookahead.LineNumber
        Expect "?"
        
        Dim Consequent As CNode
        Set Consequent = ParseAssignmentExpression()
        
        Expect ":"
        
        Dim Alternate As CNode
        Set Alternate = ParseAssignmentExpression()
        
        Dim Node As New CNode
        Node.tType = ConditionalExpression_Node
        Set Node.Test = expr
        Node.LineNumber = startLine
        Set Node.Consequent = Consequent
        Set Node.Alternate = Alternate
        Set expr = Node
    End If
    
    Set ParseConditionalExpression = expr
End Function

' ============================================
' Binary Expression Precedence Climbing
' ============================================

Private Function ParseLogicalORExpression() As CNode
    Dim expr As CNode, Node As New CNode, startLine As Long
    Set expr = ParseLogicalANDExpression()
    
    Do While match("||")
        Dim op As String
        op = lookahead.Value
        startLine = lookahead.LineNumber
        Expect op
        
        Set Node = New CNode
        Node.tType = LogicalExpression_Node
        Node.LineNumber = startLine
        Node.Operator = op
        Set Node.Left = expr
        Set Node.Right = ParseLogicalANDExpression()
        Set expr = Node
    Loop
    
    Set ParseLogicalORExpression = expr
End Function

Private Function ParseLogicalANDExpression() As CNode
    Dim expr As CNode, Node As New CNode, startLine As Long
    Set expr = ParseBitwiseORExpression()
    
    Do While match("&&")
        Dim op As String
        op = lookahead.Value
        startLine = lookahead.LineNumber
        Expect op
        
        Set Node = New CNode
        Node.tType = LogicalExpression_Node
        Node.Operator = op
        Node.LineNumber = startLine
        Set Node.Left = expr
        Set Node.Right = ParseBitwiseORExpression()
        Set expr = Node
    Loop
    
    Set ParseLogicalANDExpression = expr
End Function

Private Function ParseBitwiseORExpression() As CNode
    Dim expr As CNode, Node As New CNode, startLine As Long
    Set expr = ParseBitwiseXORExpression()
    
    Do While match("|")
        Dim op As String
        op = lookahead.Value
        startLine = lookahead.LineNumber
        Expect op
        
        Set Node = New CNode
        Node.tType = BinaryExpression_Node
        Node.Operator = op
        Node.LineNumber = startLine
        Set Node.Left = expr
        Set Node.Right = ParseBitwiseXORExpression()
        Set expr = Node
    Loop
    
    Set ParseBitwiseORExpression = expr
End Function

Private Function ParseBitwiseXORExpression() As CNode
    Dim expr As CNode, Node As New CNode, startLine As Long
    Set expr = ParseBitwiseANDExpression()
    
    Do While match("^")
        Dim op As String
        op = lookahead.Value
        startLine = lookahead.LineNumber
        Expect op
        
        Set Node = New CNode
        Node.tType = BinaryExpression_Node
        Node.Operator = op
        Node.LineNumber = startLine
        Set Node.Left = expr
        Set Node.Right = ParseBitwiseANDExpression()
        Set expr = Node
    Loop
    
    Set ParseBitwiseXORExpression = expr
End Function

Private Function ParseBitwiseANDExpression() As CNode
    Dim expr As CNode, Node As New CNode, startLine As Long
    Set expr = ParseEqualityExpression()
    
    Do While match("&")
        Dim op As String
        startLine = lookahead.LineNumber
        op = lookahead.Value
        Expect op
        
        Set Node = New CNode
        Node.tType = BinaryExpression_Node
        Node.Operator = op
        Node.LineNumber = startLine
        Set Node.Left = expr
        Set Node.Right = ParseEqualityExpression()
        Set expr = Node
    Loop
    
    Set ParseBitwiseANDExpression = expr
End Function

Private Function ParseEqualityExpression() As CNode
    Dim expr As CNode, Node As New CNode, startLine As Long
    Set expr = ParseRelationalExpression()
    
    Do While match("==") Or match("!=") Or match("===") Or match("!==")
        Dim op As String
        startLine = lookahead.LineNumber
        op = lookahead.Value
        Expect op
        
        Set Node = New CNode
        Node.tType = BinaryExpression_Node
        Node.Operator = op
        Node.LineNumber = startLine
        Set Node.Left = expr
        Set Node.Right = ParseRelationalExpression()
        Set expr = Node
    Loop
    
    Set ParseEqualityExpression = expr
End Function

Private Function ParseRelationalExpression() As CNode
    Dim expr As CNode, Node As New CNode, startLine As Long
    Set expr = ParseShiftExpression()
    
    Do While match("<") Or match(">") Or match("<=") Or match(">=") Or _
          MatchKeyword("instanceof") Or MatchKeyword("in")
        Dim op As String
        startLine = lookahead.LineNumber
        op = lookahead.Value
        Expect op
        
        Set Node = New CNode
        Node.tType = BinaryExpression_Node
        Node.Operator = op
        Node.LineNumber = startLine
        Set Node.Left = expr
        Set Node.Right = ParseShiftExpression()
        Set expr = Node
    Loop
    
    Set ParseRelationalExpression = expr
End Function

Private Function ParseShiftExpression() As CNode
    Dim expr As CNode, Node As New CNode, startLine As Long
    Set expr = ParseAdditiveExpression()
    
    Do While match("<<") Or match(">>") Or match(">>>")
        Dim op As String
        op = lookahead.Value
        startLine = lookahead.LineNumber
        Expect op
        
        Set Node = New CNode
        Node.tType = BinaryExpression_Node
        Node.Operator = op
        Node.LineNumber = startLine
        Set Node.Left = expr
        Set Node.Right = ParseAdditiveExpression()
        Set expr = Node
    Loop
    
    Set ParseShiftExpression = expr
End Function

Private Function ParseAdditiveExpression() As CNode
    Dim expr As CNode, Node As CNode, startLine As Long
    Set expr = ParseMultiplicativeExpression()
    
    Do While match("+") Or match("-")
        Dim op As String
        startLine = lookahead.LineNumber
        op = lookahead.Value
        Expect op
        
        Set Node = New CNode
        Node.tType = BinaryExpression_Node
        Node.Operator = op
        Node.LineNumber = startLine
        Set Node.Left = expr
        Set Node.Right = ParseMultiplicativeExpression()
        Set expr = Node
    Loop
    
    Set ParseAdditiveExpression = expr
End Function

Private Function ParseMultiplicativeExpression() As CNode

    Dim expr As CNode, Node As CNode, startLine As Long
    Dim op As String
    
    Set expr = ParseUnaryExpression()
    
    Do While match("*") Or match("/") Or match("%")
        startLine = lookahead.LineNumber  ' Capture BEFORE Expect
        op = lookahead.Value
        Expect op
        Set Node = New CNode
        Node.tType = BinaryExpression_Node
        Node.Operator = op
        Node.LineNumber = startLine
        Set Node.Left = expr
        Set Node.Right = ParseUnaryExpression()
        Set expr = Node
    Loop
    
    Set ParseMultiplicativeExpression = expr
End Function

' ============================================
' Unary and Postfix Expressions
' ============================================

Private Function ParseUnaryExpression() As CNode
    Dim expr As CNode, Node As CNode, startLine As Long
    
    ' Prefix operators
    If lookahead.tType = Punctuator Or lookahead.tType = Keyword Then
    
        startLine = lookahead.LineNumber
        
        If match("++") Or match("--") Then
            Set Node = New CNode
            Node.tType = UpdateExpression_Node
            Node.Operator = lookahead.Value
            Node.LineNumber = startLine
            Node.Prefix = True
            Expect Node.Operator
            Set Node.Argument = ParseUnaryExpression()
            Set expr = Node
            Set ParseUnaryExpression = expr
            Exit Function
        ElseIf match("+") Or match("-") Or match("~") Or match("!") Then
            Set Node = New CNode
            Node.tType = UnaryExpression_Node
            Node.Operator = lookahead.Value
            Node.LineNumber = startLine
            Node.Prefix = True
            Expect Node.Operator
            Set Node.Argument = ParseUnaryExpression()
            Set expr = Node
            Set ParseUnaryExpression = expr
            Exit Function
        ElseIf MatchKeyword("typeof") Or MatchKeyword("void") Or MatchKeyword("delete") Then
            Set Node = New CNode
            Node.tType = UnaryExpression_Node
            Node.Operator = lookahead.Value
            Node.LineNumber = startLine
            Node.Prefix = True
            Expect Node.Operator
            Set Node.Argument = ParseUnaryExpression()
            Set expr = Node
            Set ParseUnaryExpression = expr
            Exit Function
        End If
    End If
    
    ' Otherwise parse postfix
    Set expr = ParsePostfixExpression()
    Set ParseUnaryExpression = expr
End Function

Private Function ParsePostfixExpression() As CNode
    Dim expr As CNode, Node As CNode, startLine As Long
    
    Set expr = ParseLeftHandSideExpression()
    
    ' Postfix ++ or --
    If lookahead.tType = Punctuator Then
        If lookahead.Value = "++" Or lookahead.Value = "--" Then
            startLine = lookahead.LineNumber  ' Capture BEFORE consuming
            Set Node = New CNode
            Node.tType = UpdateExpression_Node
            Node.Operator = lookahead.Value
            Node.LineNumber = startLine
            Node.Prefix = False
            Set Node.Argument = expr
            Expect Node.Operator
            Set expr = Node
        End If
    End If
    
    Set ParsePostfixExpression = expr
End Function

' ============================================
' Left-Hand Side Expressions
' ============================================

Private Function ParseLeftHandSideExpression() As CNode
    Dim expr As CNode
    
    If MatchKeyword("new") Then
        Set expr = ParseNewExpression()
        
        ' ?? ADD: Allow member/call access after 'new'
        ' Ejemplo: new Date().toString()
        Do While match(".") Or match("[") Or match("(")
            If match(".") Then
                Set expr = ParseStaticMember(expr)
            ElseIf match("[") Then
                Set expr = ParseComputedMember(expr)
            ElseIf match("(") Then
                ' Convertir a CallExpression si hay ()
                Dim callNode As New CNode
                callNode.tType = CallExpression_Node
                callNode.LineNumber = expr.LineNumber
                Set callNode.Callee = expr
                Set callNode.Arguments = ParseArguments()
                Set expr = callNode
            End If
        Loop
    Else
        Set expr = ParseCallExpression()
    End If
    
    Set ParseLeftHandSideExpression = expr
End Function

Private Function ParseNewExpression() As CNode
    Dim X As Long
    X = lookahead.LineNumber
    
    Expect "new"
    
    Dim Node As New CNode
    Node.tType = NewExpression_Node
    Node.LineNumber = X
    Set Node.Callee = ParseMemberExpression()
    
    ' Arguments are optional for new
    If match("(") Then
        Set Node.Arguments = ParseArguments()
    End If
    
    Set ParseNewExpression = Node
End Function

Private Function ParseCallExpression() As CNode
    Dim expr As CNode           ' ? REMOVE "node As New CNode" from here
    Dim Node As CNode           ' ? Just declare it
    Dim X As Long
    X = lookahead.LineNumber
    
    Set expr = ParseMemberExpression()
    
    ' Handle call and member chains
    Do While True
        If match("(") Then
            Set Node = New CNode  ' ? This is fine now
            Node.tType = CallExpression_Node
            Node.LineNumber = X
            Set Node.Callee = expr
            Set Node.Arguments = ParseArguments()
            Set expr = Node
        ElseIf match("[") Then
            Set expr = ParseComputedMember(expr)
        ElseIf match(".") Then
            'Set expr = ParseStaticMember(expr)
            'Debug.Print ">>> Before ParseStaticMember"
            Dim tempExpr As CNode
            Set tempExpr = ParseStaticMember(expr)
            'Debug.Print ">>> After ParseStaticMember, prop Is Nothing = " & (tempExpr.prop Is Nothing)
            If Not tempExpr.prop Is Nothing Then
                'Debug.Print ">>> prop.Name = " & tempExpr.prop.Name
            End If
            Set expr = tempExpr
        Else
            Exit Do
        End If
    Loop
    
    Set ParseCallExpression = expr
End Function


Private Function ParseMemberExpression() As CNode
    Dim expr As CNode
    Set expr = ParsePrimaryExpression()
    
    Do While match(".") Or match("[")
        If match("[") Then
            Set expr = ParseComputedMember(expr)
        Else
            Set expr = ParseStaticMember(expr)
        End If
    Loop
    
    Set ParseMemberExpression = expr
End Function

Private Function ParseComputedMember(obj As CNode) As CNode
    Dim X As Long
    X = lookahead.LineNumber
    
    Expect "["
    
    Dim Node As New CNode
    Node.tType = MemberExpression_Node
    Node.LineNumber = X
    Set Node.Object = obj
    Set Node.prop = ParseExpression()
    Node.Computed = True
    
    Expect "]"
    
    Set ParseComputedMember = Node
End Function

Private Function ParseStaticMember(obj As CNode) As CNode
    Dim X As Long
    X = lookahead.LineNumber
    
    Expect "."
    
    Dim Node As CNode           ' ? REMOVE "As New"
    Set Node = New CNode        ' ? CREATE NEW
    Node.tType = MemberExpression_Node
    Node.LineNumber = X
    Set Node.Object = obj
    
    ' After dot, keywords can be used as property names
    If lookahead.tType <> Identifier And lookahead.tType <> Keyword Then
        ThrowUnexpectedToken lookahead
    End If
    
    ' Create identifier node for the property name (even if it's a keyword)
    Dim propNode As CNode       ' ? REMOVE "As New"
    Set propNode = New CNode    ' ? CREATE NEW
    propNode.tType = Identifier_Node
    propNode.name = lookahead.Value
    propNode.LineNumber = X
    Set lookahead = scanner.Lex()
    
    Set Node.prop = propNode
    'Debug.Print ">>> ParseStaticMember: Set prop to " & propNode.Name & ", prop Is Nothing = " & (node.prop Is Nothing)
    Node.Computed = False
    
    'Stop
    
    Set ParseStaticMember = Node
End Function

Private Function ParseArguments() As Collection
    Expect "("
    
    Dim args As New Collection
    
    If Not match(")") Then
        Do While True
            args.add ParseAssignmentExpression()
            If Not match(",") Then Exit Do
            Expect ","
        Loop
    End If
    
    Expect ")"
    
    Set ParseArguments = args
End Function

' ============================================
' Primary Expressions
' ============================================

Private Function ParsePrimaryExpression() As CNode

    Dim Node As New CNode
    Dim startLine As Long
    
    startLine = lookahead.LineNumber
    Node.LineNumber = lookahead.LineNumber
      
    Select Case lookahead.tType
        Case Identifier
            'Set Node = New CNode
            Node.tType = Identifier_Node
            Node.name = lookahead.Value
            ' Consume the identifier token
            Set lookahead = scanner.Lex()
            
        Case StringLiteral, NumericLiteral, Int64Literal
            'Set Node = New CNode
            Node.tType = Literal_Node
            Node.Raw = lookahead.Value
        
            If lookahead.tType = NumericLiteral Then
                Dim checkVal As String
                checkVal = lookahead.Value
                
                ' Check if it's hex
                If Len(checkVal) > 2 And (Left$(LCase(checkVal), 2) = "0x") Then
                    ' It's hex - ALWAYS use ULong64 to parse
                    Dim u64Parse As New ULong64
                    u64Parse.mode = mUnsigned
                    
                    If u64Parse.fromString(checkVal, mHex) Then
                        Dim hi As Long, lo As Long
                        u64Parse.GetLongs hi, lo
                        
                        ' Check if it needs Int64 storage
                        If hi <> 0 Then
                            ' Definitely needs Int64 - store as string (no "n" - let EvaluateLiteral decide)
                            Node.Value = checkVal
                        ElseIf lo < 0 Then
                            ' Lo is negative (0x80000000-0xFFFFFFFF range)
                            ' Store as string so EvaluateLiteral can auto-promote to BigInt
                            Node.Value = checkVal
                        Else
                            ' Small positive value (0x00000000-0x7FFFFFFF)
                            ' Safe to store as Double
                            Node.Value = CDbl(lo)
                        End If
                    Else
                        ' Parse failed - store as 0
                        Node.Value = CDbl(0)
                    End If
                Else
                    ' Regular decimal number
                    Node.Value = CDbl(checkVal)
                End If
                
            ElseIf lookahead.tType = Int64Literal Then
                ' Already marked as Int64 by scanner
                Node.Value = lookahead.Value & "n"
                
            Else
                ' StringLiteral - keep as string
                Node.Value = lookahead.Value
            End If
            
            ' Consume the literal token
            Set lookahead = scanner.Lex()
            
        Case BooleanLiteral
            'Set Node = New CNode
            Node.tType = Literal_Node
            Node.Raw = lookahead.Value
            Node.Value = (lookahead.Value = "true")
            ' Consume the boolean token
            Set lookahead = scanner.Lex()
            
        Case NullLiteral
            'Set Node = New CNode
            Node.tType = Literal_Node
            Node.Raw = "null"
            Node.Value = Null
            ' Consume the null token
            Set lookahead = scanner.Lex()
            
        Case Keyword
            If MatchKeyword("this") Then
                Expect "this"
                'Set Node = New CNode
                Node.tType = ThisExpression_Node
            ElseIf MatchKeyword("function") Then
                Set Node = ParseFunctionExpression()
                Node.LineNumber = startLine
            Else
                ThrowUnexpectedToken lookahead
            End If
            
        Case Punctuator
            Select Case lookahead.Value
                Case "["
                    Set Node = ParseArrayExpression()
                    Node.LineNumber = startLine
                Case "{"
                    Set Node = ParseObjectExpression()
                    Node.LineNumber = startLine
                Case "("
                    Set Node = ParseGroupExpression()
                    Node.LineNumber = startLine
                Case Else
                    ThrowUnexpectedToken lookahead
            End Select
            
        Case RegularExpression
            'Set Node = New CNode
            Node.tType = Literal_Node
            Node.Raw = lookahead.Value
            Node.Value = lookahead.Value
            Set lookahead = scanner.Lex()
            
        Case Else
            ThrowUnexpectedToken lookahead
    End Select
    
    Set ParsePrimaryExpression = Node
End Function

Private Function ParseArrayExpression() As CNode
    Expect "["
    
    Dim Node As New CNode
    Node.tType = ArrayExpression_Node
    
    Do While Not match("]")
        If match(",") Then
            ' Sparse array - add Nothing
            Node.Body.add Nothing
            Expect ","
        Else
            Node.Body.add ParseAssignmentExpression()
            If Not match("]") Then
                Expect ","
            End If
        End If
    Loop
    
    Expect "]"
    
    Set ParseArrayExpression = Node
End Function

Private Function ParseObjectExpression() As CNode
    Dim startLine As Long
    Dim funcNode As CNode
    Dim keyNode As CNode
    Dim prop As CNode
    Dim getSetToken As CToken
    
    startLine = lookahead.LineNumber
    
    Expect "{"
    
    Dim Node As New CNode
    Node.tType = ObjectExpression_Node
    Node.LineNumber = startLine
    
    Do While Not match("}")
        Set prop = New CNode
        prop.tType = Property_Node
        prop.LineNumber = lookahead.LineNumber
        
        ' Check for getter/setter
        ' Must be careful: {get: 5} is a regular property, but {get foo()} is a getter
        If lookahead.tType = Identifier And (lookahead.Value = "get" Or lookahead.Value = "set") Then
            ' Save the get/set keyword
            Dim propKind As String
            propKind = lookahead.Value
            
            ' Consume 'get' or 'set'
            Set getSetToken = lookahead
            Set lookahead = scanner.Lex()
            
            ' Now check: is next token an identifier followed by '('?
            ' If yes, it's a getter/setter. If no, it's a regular property named 'get' or 'set'
            If lookahead.tType = Identifier Then
                ' This looks like: get name() or set name()
                prop.propKind = propKind
                
                ' Parse property name
                Set prop.PropKey = ParsePrimaryExpression()
                
                ' Must be followed by (
                If Not match("(") Then
                    ' Oops, not a getter/setter after all - this is malformed
                    ThrowUnexpectedToken lookahead
                End If
                
                Expect "("
                
                ' Parse as function
                Set funcNode = New CNode
                funcNode.tType = FunctionExpression_Node
                funcNode.LineNumber = getSetToken.LineNumber
                
                ' Setters have one parameter, getters have none
                If propKind = "set" Then
                    If Not match(")") Then
                        If lookahead.tType = Identifier Then
                            funcNode.params.add ParsePrimaryExpression()
                        End If
                    End If
                End If
                
                Expect ")"
                
                Set funcNode.FunctionBody = ParseBlockStatement()
                Set prop.propValue = funcNode
                
                Node.Properties.add prop
                
                If Not match("}") Then
                    If match(",") Then Expect ","
                End If
                
                GoTo NextProperty
                
            ElseIf match(":") Then
                ' This is a regular property: {get: value} or {set: value}
                ' We already consumed 'get'/'set', so treat it as the key
                prop.propKind = "init"
                
                Set keyNode = New CNode
                keyNode.tType = Identifier_Node
                keyNode.name = propKind
                keyNode.LineNumber = getSetToken.LineNumber
                Set prop.PropKey = keyNode
                
                Expect ":"
                Set prop.propValue = ParseAssignmentExpression()
                
                Node.Properties.add prop
                
                If Not match("}") Then
                    Expect ","
                End If
                
                GoTo NextProperty
            Else
                ' Unexpected
                ThrowUnexpectedToken lookahead
            End If
        End If
        
        ' Regular property
        prop.propKind = "init"
        
        ' Parse key
        If lookahead.tType = Identifier Then
            Set prop.PropKey = ParsePrimaryExpression()
        ElseIf lookahead.tType = StringLiteral Or lookahead.tType = NumericLiteral Then
            Set prop.PropKey = ParsePrimaryExpression()
        Else
            ThrowUnexpectedToken lookahead
        End If
        
        Expect ":"
        Set prop.propValue = ParseAssignmentExpression()
        
        Node.Properties.add prop
        
        If Not match("}") Then
            Expect ","
        End If
        
NextProperty:
    Loop
    
    Expect "}"
    
    Set ParseObjectExpression = Node
End Function


Private Function ParseGroupExpression() As CNode
    Expect "("
    Dim expr As CNode
    Set expr = ParseExpression()
    Expect ")"
    Set ParseGroupExpression = expr
End Function

' ============================================
' Function Parsing
' ============================================

Private Function ParseFunctionDeclaration() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line number
    
    Expect "function"
    
    Dim Node As New CNode
    Node.tType = FunctionDeclaration_Node
    Node.LineNumber = startLine  ' SET IT HERE
    
    ' Function must have a name
    If lookahead.tType <> Identifier Then
        ThrowUnexpectedToken lookahead
    End If
    
    Set Node.id = ParsePrimaryExpression()
    
    Expect "("
    
    ' Parse parameters
    If Not match(")") Then
        Do While True
            If lookahead.tType <> Identifier Then
                ThrowUnexpectedToken lookahead
            End If
            Node.params.add ParsePrimaryExpression()
            
            If Not match(",") Then Exit Do
            Expect ","
        Loop
    End If
    
    Expect ")"
    
    ' Parse body - ParseBlockStatement will consume the {
    Set Node.FunctionBody = ParseBlockStatement()
    
    Set ParseFunctionDeclaration = Node
End Function

Private Function ParseFunctionExpression() As CNode
    Dim startLine As Long
    startLine = lookahead.LineNumber  ' Capture line number
    
    Expect "function"
    
    Dim Node As New CNode
    Node.tType = FunctionExpression_Node
    Node.LineNumber = startLine  ' SET IT HERE
    
    ' Function name is optional for expressions
    If lookahead.tType = Identifier Then
        Set Node.id = ParsePrimaryExpression()
    End If
    
    Expect "("
    
    ' Parse parameters
    If Not match(")") Then
        Do While True
            If lookahead.tType <> Identifier Then
                ThrowUnexpectedToken lookahead
            End If
            Node.params.add ParsePrimaryExpression()
            
            If Not match(",") Then Exit Do
            Expect ","
        Loop
    End If
    
    Expect ")"
    
    ' Parse body - ParseBlockStatement will consume the {
    Set Node.FunctionBody = ParseBlockStatement()
    
    Set ParseFunctionExpression = Node
End Function

Private Function ParseVariableDeclaration() As CNode
    Expect "var"
    
    Dim Node As New CNode
    Node.tType = VariableDeclaration_Node
    Node.Kind = "var"
    
    ' Parse declarators
    Do
        Node.Declarations.add ParseVariableDeclarator()
        
        If Not match(",") Then Exit Do
        Expect ","
    Loop
    
    Set ParseVariableDeclaration = Node
End Function

' ============================================
' WE DID IT! A COMPLETE PARSER! IN VB6!
' "To infinity... and beyond!"
' ============================================


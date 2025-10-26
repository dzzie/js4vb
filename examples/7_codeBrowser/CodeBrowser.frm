VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCodeBrowser 
   Caption         =   "JavaScript Code Browser"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse File"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox txtFilePath 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5760
      Width           =   8535
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   8760
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin MSComctlLib.TreeView tvCode 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9763
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "File Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   1335
   End
End
Attribute VB_Name = "frmCodeBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'------------------------------------------------------------
' frmCodeBrowser.frm - JavaScript Code Browser
'
' "I can show you the world... of JavaScript ASTs!" - Aladdin, probably
'
' Real-world application: Parse JS files and show structure
' in a TreeView for easy navigation and analysis.
'------------------------------------------------------------

Option Explicit

Private parser As CParser
Private m_recursionDepth As Long

Private Const debug_mode As Boolean = True

Private Sub Form_Load()
    Set parser = New CParser
    
    ' Setup TreeView
    tvCode.Nodes.Clear
    
    ' Setup icons FIRST (or disable them)
    SetupIcons
    
    ' IMPORTANT: If you don't have icons, remove the ImageList reference!
    ' Set tvCode.ImageList = Nothing
    
    ' Example file path
    txtFilePath.Text = App.Path & "\test.js"
    cmdParse_Click
    
    Dim n As node
    For Each n In tvCode.Nodes
        n.Expanded = True
    Next
    
End Sub

Private Sub SetupIcons()
    ' ============================================
    ' OPTION 1: NO ICONS (Recommended for testing)
    ' ============================================
    ' Just disconnect the ImageList from TreeView
    Set tvCode.ImageList = Nothing
    
    ' Now the TreeView will work fine without icons!
    ' You can add icons later if you want
    
    Exit Sub
    
    ' ============================================
    ' OPTION 2: Load from .ico files (use if you have icons)
    ' ============================================
    ' Uncomment this section if you have icon files
    '
    ' On Error Resume Next
    ' ImageList1.ListImages.Clear
    '
    ' ' Load icons from files (create an \icons subfolder)
    ' ImageList1.ListImages.Add 1, "program", LoadPicture(App.Path & "\icons\file.ico")
    ' ImageList1.ListImages.Add 2, "function", LoadPicture(App.Path & "\icons\function.ico")
    ' ImageList1.ListImages.Add 3, "param", LoadPicture(App.Path & "\icons\param.ico")
    ' ImageList1.ListImages.Add 4, "vardecl", LoadPicture(App.Path & "\icons\variable.ico")
    ' ImageList1.ListImages.Add 5, "var", LoadPicture(App.Path & "\icons\var.ico")
    ' ImageList1.ListImages.Add 6, "if", LoadPicture(App.Path & "\icons\branch.ico")
    ' ImageList1.ListImages.Add 7, "loop", LoadPicture(App.Path & "\icons\loop.ico")
    ' ImageList1.ListImages.Add 8, "switch", LoadPicture(App.Path & "\icons\switch.ico")
    ' ImageList1.ListImages.Add 9, "case", LoadPicture(App.Path & "\icons\case.ico")
    ' ImageList1.ListImages.Add 10, "trycatch", LoadPicture(App.Path & "\icons\error.ico")
    ' ImageList1.ListImages.Add 11, "return", LoadPicture(App.Path & "\icons\return.ico")
    ' ImageList1.ListImages.Add 12, "throw", LoadPicture(App.Path & "\icons\throw.ico")
    ' ImageList1.ListImages.Add 13, "break", LoadPicture(App.Path & "\icons\break.ico")
    ' ImageList1.ListImages.Add 14, "expression", LoadPicture(App.Path & "\icons\expression.ico")
    ' ImageList1.ListImages.Add 15, "object", LoadPicture(App.Path & "\icons\object.ico")
    ' ImageList1.ListImages.Add 16, "array", LoadPicture(App.Path & "\icons\array.ico")
    '
    ' On Error GoTo 0
    '
    ' ' If icons didn't load, disconnect ImageList
    ' If ImageList1.ListImages.Count = 0 Then
    '     Set tvCode.ImageList = Nothing
    ' End If
End Sub

Private Sub cmdBrowse_Click()
    ' Use Common Dialog to browse for file
    Dim dlg As Object
    On Error Resume Next
    Set dlg = CreateObject("MSComDlg.CommonDialog")
    
    If Err.Number = 0 Then
        dlg.Filter = "JavaScript Files (*.js)|*.js|All Files (*.*)|*.*"
        dlg.ShowOpen
        If dlg.FileName <> "" Then
            txtFilePath.Text = dlg.FileName
        End If
    Else
        ' Fallback: manual entry
        MsgBox "Common Dialog not available. Please enter path manually.", vbInformation
    End If
End Sub

Private Sub cmdParse_Click()
    If Trim(txtFilePath.Text) = "" Then
        MsgBox "Please enter a file path", vbExclamation
        Exit Sub
    End If
    
    If Not FileExists(txtFilePath.Text) Then
        MsgBox "File not found: " & txtFilePath.Text, vbExclamation
        Exit Sub
    End If
    
    ' Read file
    Dim Code As String
    Code = ReadTextFile(txtFilePath.Text)
    
    If Code = "" Then
        MsgBox "File is empty or could not be read", vbExclamation
        Exit Sub
    End If
    
    ' Parse it!
    Me.MousePointer = vbHourglass
    tvCode.Nodes.Clear
    
    On Error GoTo ParseError
    
    Dim ast As CNode
    Set ast = parser.ParseScript(Code)
    
    ' Build the tree
    BuildTreeView ast
    
    ' OUTPUT TEXT TREE TO DEBUG WINDOW
    Debug.Print "========================================"
    Debug.Print "JavaScript Code Structure"
    Debug.Print "========================================"
    PrintTreeToDebug
    Debug.Print "========================================"
    Debug.Print "Total nodes: " & tvCode.Nodes.Count
    Debug.Print "========================================"
    
    Me.MousePointer = vbDefault
    'MsgBox "Parsing complete! Found " & tvCode.Nodes.Count & " nodes." & vbCrLf & vbCrLf & "See Immediate Window (Ctrl+G) for text output.", vbInformation
    Exit Sub
    
ParseError:
    Me.MousePointer = vbDefault
    MsgBox "Parse error: " & Err.Description, vbCritical
End Sub

' ============================================
' Print tree structure to Debug window
' ============================================

Private Sub PrintTreeToDebug()
    Dim node As MSComctlLib.node
    For Each node In tvCode.Nodes
        ' Calculate indent based on level
        Dim level As Long
        level = 0
        Dim tempNode As MSComctlLib.node
        Set tempNode = node
        Do While Not tempNode.Parent Is Nothing
            level = level + 1
            Set tempNode = tempNode.Parent
        Loop
        
        Dim indent As String
        indent = String$(level * 2, " ")
        
        If node.Children > 0 Then
            Debug.Print indent & "+ " & node.Text
        Else
            Debug.Print indent & "- " & node.Text
        End If
    Next node
End Sub

' ============================================
' Build TreeView from AST
' ============================================

Private Sub BuildTreeView(ast As CNode)
    ' Add root node
    ' Reset recursion depth counter
    m_recursionDepth = 0
    
    Dim rootNode As node
    Set rootNode = tvCode.Nodes.Add(, , "root", "?? JavaScript Program (" & ast.Body.Count & " statements)") ', 1)
    
    ' Process each top-level statement
    Dim i As Long
    For i = 1 To ast.Body.Count
        Dim stmt As CNode
        Set stmt = ast.Body(i)
        ProcessNode stmt, rootNode
    Next i
    
    ' Expand root and first level
    rootNode.Expanded = True
    
    ' Expand all function nodes so you can see their contents
    Dim node As node
    For Each node In tvCode.Nodes
        If InStr(node.Text, "function") > 0 Then
            node.Expanded = True
        End If
    Next node
End Sub

Private Sub ProcessNode(node As CNode, parentNode As node)

    Debug.Print "Node: " & node.Name
    
    If node Is Nothing Then Exit Sub
    
    ' Declare ALL variables at the top (VB6 requirement)
    Dim nodeText As String
    Dim nodeKey As String
    Dim newNode As MSComctlLib.node
    Dim j As Long
    Dim k As Long
    Dim paramsText As String
    Dim param As CNode
    Dim decl As CNode           ' ? ADD THIS
    Dim varText As String       ' ? ADD THIS
    Dim varNode As MSComctlLib.node  ' ? ADD THIS
    Dim elseNode As MSComctlLib.node
    Dim switchCase As CNode
    Dim caseText As String
    Dim caseNode As MSComctlLib.node
    Dim tryNode As MSComctlLib.node
    Dim catchNode As MSComctlLib.node
    Dim catchText As String
    Dim finallyNode As MSComctlLib.node
    Dim prop As CNode
    Dim propText As String
    Dim propNode As MSComctlLib.node
    
    ' Generate unique key
    nodeKey = "node_" & tvCode.Nodes.Count
    
    ' Prevent infinite recursion
    m_recursionDepth = m_recursionDepth + 1
    If m_recursionDepth > 100 Then
        Debug.Print "WARNING: Max recursion depth reached in ProcessNode"
        m_recursionDepth = m_recursionDepth - 1
        Exit Sub
    End If

    Select Case node.tType
        ' ============================================
        ' FUNCTION DECLARATIONS - THE IMPORTANT ONES!
        ' ============================================
        Case FunctionDeclaration_Node
            If debug_mode Then Debug.Print "CASE: FunctionDeclaration_Node"
            nodeText = "function " & node.id.Name & "()"
            If node.LineNumber > 0 Then
                nodeText = nodeText & " [Line " & node.LineNumber & "]"
            End If
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 2)
            newNode.Bold = True
            
            ' Add parameters
            If node.Params.Count > 0 Then
                paramsText = "Parameters: "
                For j = 1 To node.Params.Count
                    Set param = node.Params(j)
                    paramsText = paramsText & param.Name
                    If j < node.Params.Count Then paramsText = paramsText & ", "
                Next j
                tvCode.Nodes.Add newNode, tvwChild, nodeKey & "_params", paramsText ', 3)
            End If
            
            ' Process function body
            If Not node.FunctionBody Is Nothing Then
                ProcessNode node.FunctionBody, newNode
            End If
            
        ' ============================================
        ' FUNCTION EXPRESSIONS (Anonymous functions)
        ' ============================================
        Case FunctionExpression_Node
            If debug_mode Then Debug.Print "CASE: FunctionExpression_Node"
            If node.id Is Nothing Then
                nodeText = "function (anonymous)"
            Else
                nodeText = "function " & node.id.Name & "()"
            End If
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 2)
            
            ' Process function body
            If Not node.FunctionBody Is Nothing Then
                ProcessNode node.FunctionBody, newNode
            End If
            
        ' ============================================
        ' VARIABLE DECLARATIONS
        ' ============================================
        Case VariableDeclaration_Node
            If debug_mode Then Debug.Print "CASE: VariableDeclaration_Node"
            nodeText = node.Kind & " declaration"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 4)
            
            ' Add each variable
            For j = 1 To node.Declarations.Count
                Set decl = node.Declarations(j)
                varText = decl.VarID.Name
                
                ' Check if initializer is a function expression
                If Not decl.Init Is Nothing Then
                    If decl.Init.tType = FunctionExpression_Node Then
                        ' This is a function assigned to a variable!
                        varText = varText & " = function()"
                        Set varNode = tvCode.Nodes.Add(newNode, tvwChild, nodeKey & "_var" & j, varText) ', 2)
                        varNode.Bold = True
                        
                        ' Add parameters if any
                        If decl.Init.Params.Count > 0 Then
                            paramsText = "Parameters: "
                            For k = 1 To decl.Init.Params.Count

                                Set param = decl.Init.Params(k)
                                paramsText = paramsText & param.Name
                                If k < decl.Init.Params.Count Then paramsText = paramsText & ", "
                            Next k
                            tvCode.Nodes.Add varNode, tvwChild, nodeKey & "_var" & j & "_params", paramsText ', 3)
                        End If
                        
                        ' Process function body
                        If Not decl.Init.FunctionBody Is Nothing Then
                            ProcessNode decl.Init.FunctionBody, varNode
                        End If
                    Else
                        ' Regular variable with non-function initializer
                        varText = varText & " = " & GetExpressionPreview(decl.Init)
                        tvCode.Nodes.Add newNode, tvwChild, nodeKey & "_var" & j, varText ', 5)
                    End If
                Else
                    ' Variable with no initializer
                    tvCode.Nodes.Add newNode, tvwChild, nodeKey & "_var" & j, varText ', 5)
                End If
            Next j
            
        ' ============================================
        ' CONTROL FLOW STATEMENTS
        ' ============================================
        Case IfStatement_Node
            If debug_mode Then Debug.Print "CASE: IfStatement_Node"
            nodeText = "if statement"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 6)
            
            ' Process consequent
            If Not node.IfConsequent Is Nothing Then
                ProcessNode node.IfConsequent, newNode
            End If
            
            ' Process alternate (else)
            If Not node.IfAlternate Is Nothing Then
                Set elseNode = tvCode.Nodes.Add(newNode, tvwChild, nodeKey & "_else", "else") ', 6)
                ProcessNode node.IfAlternate, elseNode
            End If
            
        Case WhileStatement_Node
            If debug_mode Then Debug.Print "CASE: WhileStatement_Node"
            nodeText = "while loop"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 7)
            If Not node.WhileBody Is Nothing Then
                ProcessNode node.WhileBody, newNode
            End If
            
        Case DoWhileStatement_Node
            If debug_mode Then Debug.Print "CASE: DoWhileStatement_Node"
            nodeText = "do-while loop"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 7)
            If Not node.WhileBody Is Nothing Then
                ProcessNode node.WhileBody, newNode
            End If
            
        Case ForStatement_Node
            If debug_mode Then Debug.Print "CASE: ForStatement_Node"
            nodeText = "for loop"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 7)
            If Not node.ForBody Is Nothing Then
                ProcessNode node.ForBody, newNode
            End If
            
        Case ForInStatement_Node
            If debug_mode Then Debug.Print "CASE: ForInStatement_Node"
            nodeText = "for-in loop"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 7)
            If Not node.ForBody Is Nothing Then
                ProcessNode node.ForBody, newNode
            End If
            
        Case SwitchStatement_Node
            If debug_mode Then Debug.Print "CASE: SwitchStatement_Node"
            nodeText = "switch statement"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 8)
            
            ' Process cases
            For j = 1 To node.Cases.Count
                Set switchCase = node.Cases(j)
                If switchCase.CaseTest Is Nothing Then
                    caseText = "default:"
                Else
                    caseText = "case " & GetExpressionPreview(switchCase.CaseTest) & ":"
                End If
                Set caseNode = tvCode.Nodes.Add(newNode, tvwChild, nodeKey & "_case" & j, caseText) ', 9)
                
                ' Process case statements
                For k = 1 To switchCase.CaseConsequent.Count
                    ProcessNode switchCase.CaseConsequent(k), caseNode
                Next k
            Next j
            
        Case TryStatement_Node
            If debug_mode Then Debug.Print "CASE: TryStatement_Node"
            nodeText = "try-catch"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 10)
            
            ' Try block
            If Not node.TryBlock Is Nothing Then
                Set tryNode = tvCode.Nodes.Add(newNode, tvwChild, nodeKey & "_try", "try") ', 10)
                ProcessNode node.TryBlock, tryNode
            End If
            
            ' Catch block
            If Not node.TryHandler Is Nothing Then
                catchText = "catch (" & node.TryHandler.param.Name & ")"
                Set catchNode = tvCode.Nodes.Add(newNode, tvwChild, nodeKey & "_catch", catchText) ', 10)
                ProcessNode node.TryHandler.CatchBody, catchNode
            End If
            
            ' Finally block
            If Not node.TryFinalizer Is Nothing Then
                Set finallyNode = tvCode.Nodes.Add(newNode, tvwChild, nodeKey & "_finally", "finally") ', 10)
                ProcessNode node.TryFinalizer, finallyNode
            End If
            
        ' ============================================
        ' RETURN/THROW STATEMENTS
        ' ============================================
        Case ReturnStatement_Node
            If debug_mode Then Debug.Print "CASE: ReturnStatement_Node"
            If node.ReturnArgument Is Nothing Then
                nodeText = "return"
            Else
                nodeText = "return " & GetExpressionPreview(node.ReturnArgument)
            End If
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 11)
            
        Case ThrowStatement_Node
            If debug_mode Then Debug.Print "CASE: ThrowStatement_Node"
            nodeText = "throw " & GetExpressionPreview(node.ReturnArgument)
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 12)
            
        Case BreakStatement_Node
            If debug_mode Then Debug.Print "CASE: BreakStatement_Node"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, "break") ', 13)
            
        Case ContinueStatement_Node
            If debug_mode Then Debug.Print "CASE: ContinueStatement_Node"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, "continue") ', 13)
            
        ' ============================================
        ' BLOCK STATEMENTS (process children)
        ' ============================================
        Case BlockStatement_Node
            If debug_mode Then Debug.Print "CASE: BlockStatement_Node"
            ' Don't add a node for blocks, just process children
            Debug.Print "BlockStatement_Node calling for:" & node.Body.Count
            For j = 1 To node.Body.Count
                ProcessNode node.Body(j), parentNode
            Next j
            
        ' ============================================
        ' EXPRESSION STATEMENTS
        ' ============================================
        ' ============================================
        ' EXPRESSION STATEMENTS
        ' ============================================
        Case ExpressionStatement_Node
            If debug_mode Then Debug.Print "CASE: ExpressionStatement_Node"
            ' Check if this is a Call Expression (could be an IIFE)
            If Not node.Test Is Nothing Then
                If node.Test.tType = CallExpression_Node Then
                    ' Check if callee is a function expression (IIFE pattern)
                    If Not node.Test.Callee Is Nothing Then
                        If node.Test.Callee.tType = FunctionExpression_Node Then
                            ' This is an IIFE!
                            nodeText = "(function()...)()"
                            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText)
                            newNode.Bold = True
                            
                            ' Process the IIFE function body
                            If Not node.Test.Callee.FunctionBody Is Nothing Then
                                ProcessNode node.Test.Callee.FunctionBody, newNode
                            End If
                            m_recursionDepth = m_recursionDepth - 1
                            Exit Sub
                        End If
                    End If
                End If
                
                ' Check if this is an assignment expression
                If node.Test.tType = AssignmentExpression_Node Then
                    ' Get the left side for display
                    Dim leftSide As String
                    leftSide = GetExpressionPreview(node.Test.Left)
                    
                    ' Check if right side is a function
                    If Not node.Test.Right Is Nothing Then
                        If node.Test.Right.tType = FunctionExpression_Node Then
                            ' This is assigning a function!
                            nodeText = leftSide & " = function()"
                            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText)
                            newNode.Bold = True
                            
                            ' Add parameters
                            If node.Test.Right.Params.Count > 0 Then
                                paramsText = "Parameters: "
                                For j = 1 To node.Test.Right.Params.Count
                                    Set param = node.Test.Right.Params(j)
                                    paramsText = paramsText & param.Name
                                    If j < node.Test.Right.Params.Count Then paramsText = paramsText & ", "
                                Next j
                                tvCode.Nodes.Add newNode, tvwChild, nodeKey & "_params", paramsText
                            End If
                            
                            ' Process function body
                            If Not node.Test.Right.FunctionBody Is Nothing Then
                                ProcessNode node.Test.Right.FunctionBody, newNode
                            End If
                            m_recursionDepth = m_recursionDepth - 1
                            Exit Sub
                            
                        ElseIf node.Test.Right.tType = ObjectExpression_Node Then
                            ' This is assigning an object literal!
                            nodeText = leftSide & " = {...}"
                            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText)
                            newNode.Bold = True
                            
                            ' Process object properties directly (don't call ProcessNode on the whole object)
                            If node.Test.Right.Properties.Count > 0 Then
                                For j = 1 To node.Test.Right.Properties.Count
                                    Set prop = node.Test.Right.Properties(j)
                                    
                                    ' DEBUG: See what the property key actually is
                                    If debug_mode Then Debug.Print "Property " & j & " - PropKey.Name: " & prop.PropKey.Name & ", PropKey.tType: " & prop.PropKey.tType
       
                                    propText = GetExpressionPreview(prop.PropKey) & ": "
                                    
                                    ' Check if property value is a function (method)
                                    If Not prop.PropValue Is Nothing Then
                                        If prop.PropValue.tType = FunctionExpression_Node Then
                                            propText = propText & "function()"
                                            Set propNode = tvCode.Nodes.Add(newNode, tvwChild, nodeKey & "_prop" & j, propText)
                                            propNode.Bold = True
                                            
                                            ' Add parameters
                                            If prop.PropValue.Params.Count > 0 Then
                                                paramsText = "Parameters: "
                                                For k = 1 To prop.PropValue.Params.Count
                                                    Set param = prop.PropValue.Params(k)
                                                    paramsText = paramsText & param.Name
                                                    If k < prop.PropValue.Params.Count Then paramsText = paramsText & ", "
                                                Next k
                                                tvCode.Nodes.Add propNode, tvwChild, nodeKey & "_prop" & j & "_params", paramsText
                                            End If
                                            
                                            ' Process function body
                                            If Not prop.PropValue.FunctionBody Is Nothing Then
                                                ProcessNode prop.PropValue.FunctionBody, propNode
                                            End If
                                        Else
                                            propText = propText & GetExpressionPreview(prop.PropValue)
                                            tvCode.Nodes.Add newNode, tvwChild, nodeKey & "_prop" & j, propText
                                        End If
                                    End If
                                Next j
                            End If
                            
                            m_recursionDepth = m_recursionDepth - 1
                            Exit Sub
                        End If
                    End If
                    
                    ' For any other assignment, show it
                    nodeText = leftSide & " = " & GetExpressionPreview(node.Test.Right)
                    Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText)
                    m_recursionDepth = m_recursionDepth - 1
                    Exit Sub
                End If
            End If
            
            ' Regular expression statement
            nodeText = GetExpressionPreview(node.Test)
            If Len(nodeText) > 50 Then
                nodeText = Left(nodeText, 47) & "..."
            End If
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText)
            
        ' ============================================
        ' OBJECT/ARRAY EXPRESSIONS
        ' ============================================
        Case ObjectExpression_Node
            If debug_mode Then Debug.Print "CASE: ObjectExpression_Node"
            nodeText = "{ object literal }"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 15)
            
            ' Process object properties - check for methods!
            If node.Properties.Count > 0 Then
                For j = 1 To node.Properties.Count
                    Set prop = node.Properties(j)
                    
                    propText = GetExpressionPreview(prop.PropKey) & ": "
                    
                    ' Check if property value is a function (method)
                    If Not prop.PropValue Is Nothing Then
                        If prop.PropValue.tType = FunctionExpression_Node Then
                            propText = propText & "function()"
                            Set propNode = tvCode.Nodes.Add(newNode, tvwChild, nodeKey & "_prop" & j, propText) ', 2)
                            propNode.Bold = True
                            
                            ' Add parameters
                            If prop.PropValue.Params.Count > 0 Then
                                paramsText = "Parameters: "
                                For k = 1 To prop.PropValue.Params.Count
                                    Set param = prop.PropValue.Params(k)
                                    paramsText = paramsText & param.Name
                                    If k < prop.PropValue.Params.Count Then paramsText = paramsText & ", "
                                Next k
                                tvCode.Nodes.Add propNode, tvwChild, nodeKey & "_prop" & j & "_params", paramsText ', 3)
                            End If
                            
                            ' Process function body
                            If Not prop.PropValue.FunctionBody Is Nothing Then
                                ProcessNode prop.PropValue.FunctionBody, propNode
                            End If
                        Else
                            propText = propText & GetExpressionPreview(prop.PropValue)
                            tvCode.Nodes.Add newNode, tvwChild, nodeKey & "_prop" & j, propText ', 5)
                        End If
                    End If
                Next j
            End If
            
        Case ArrayExpression_Node
            If debug_mode Then Debug.Print "CASE: ArrayExpression_Node"
            nodeText = "[ array literal ]"
            Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, nodeText) ', 16)
            
    End Select
    
    ' Decrement recursion counter
    m_recursionDepth = m_recursionDepth - 1
End Sub



' ============================================
' Helper: Get preview of expression
' ============================================
Private Function GetExpressionPreview(node As CNode, Optional depth As Long = 0) As String
    ' Prevent infinite recursion
    If depth > 10 Then
        GetExpressionPreview = "..."
        Exit Function
    End If
    
    If node Is Nothing Then
        GetExpressionPreview = "(empty)"
        Exit Function
    End If
    
    Select Case node.tType
        Case Identifier_Node
            If debug_mode Then Debug.Print "GetExpressionPreview: Identifier_Node (depth=" & depth & ")"
            GetExpressionPreview = node.Name
            
        Case Literal_Node
            If debug_mode Then Debug.Print "GetExpressionPreview: Literal_Node (depth=" & depth & ")"
            If VarType(node.Value) = vbString Then
                GetExpressionPreview = """" & node.Value & """"
            Else
                GetExpressionPreview = CStr(node.Value)
            End If
            
        Case BinaryExpression_Node, LogicalExpression_Node
            If debug_mode Then Debug.Print "GetExpressionPreview: BinaryExpression_Node/LogicalExpression_Node (depth=" & depth & ")"
            GetExpressionPreview = GetExpressionPreview(node.Left, depth + 1) & " " & node.Operator & " " & GetExpressionPreview(node.Right, depth + 1)
            
        Case CallExpression_Node
            If debug_mode Then Debug.Print "GetExpressionPreview: CallExpression_Node (depth=" & depth & ")"
            GetExpressionPreview = GetExpressionPreview(node.Callee, depth + 1) & "()"
            
        Case MemberExpression_Node
            If debug_mode Then Debug.Print "GetExpressionPreview: MemberExpression_Node (depth=" & depth & ")"
            Dim objPreview As String
            objPreview = GetExpressionPreview(node.Object, depth + 1)
            ' If we already hit the limit, don't add more
            If objPreview = "..." Then
                GetExpressionPreview = "..."
            ElseIf node.Computed Then
                GetExpressionPreview = objPreview & "[...]"
            Else
                GetExpressionPreview = objPreview & "." & GetExpressionPreview(node.prop, depth + 1)
            End If
            
        Case NewExpression_Node
            If debug_mode Then Debug.Print "GetExpressionPreview: NewExpression_Node (depth=" & depth & ")"
            GetExpressionPreview = "new " & GetExpressionPreview(node.Callee, depth + 1)
            
        Case FunctionExpression_Node
            If debug_mode Then Debug.Print "GetExpressionPreview: FunctionExpression_Node (depth=" & depth & ")"
            GetExpressionPreview = "function(...)"
            
        Case ObjectExpression_Node
            If debug_mode Then Debug.Print "GetExpressionPreview: ObjectExpression_Node (depth=" & depth & ")"
            GetExpressionPreview = "{...}"
            
        Case ArrayExpression_Node
            If debug_mode Then Debug.Print "GetExpressionPreview: ArrayExpression_Node (depth=" & depth & ")"
            GetExpressionPreview = "[...]"
            
        Case Else
            If debug_mode Then Debug.Print "GetExpressionPreview: Case Else (unknown type, depth=" & depth & ")"
            GetExpressionPreview = "(expression)"
    End Select
End Function



' ============================================
' File I/O Helpers
' ============================================

Private Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
    On Error GoTo 0
End Function

Private Function ReadTextFile(filePath As String) As String
    Dim fileNum As Integer
    Dim fileContent As String
    
    On Error GoTo ReadError
    
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    
    If LOF(fileNum) > 0 Then
        fileContent = Space$(LOF(fileNum))
        Get #fileNum, , fileContent
    End If
    
    Close #fileNum
    
    ReadTextFile = fileContent
    Exit Function
    
ReadError:
    If fileNum > 0 Then Close #fileNum
    ReadTextFile = ""
End Function

Private Sub tvCode_NodeClick(ByVal node As MSComctlLib.node)
    ' You could add code here to jump to line numbers,
    ' show more details, etc.
    Debug.Print "Selected: " & node.Text
End Sub

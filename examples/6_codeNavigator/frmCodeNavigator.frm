VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmCodeNavigator 
   Caption         =   "Form2"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13440
   LinkTopic       =   "Form2"
   ScaleHeight     =   6735
   ScaleWidth      =   13440
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10575
      Top             =   5850
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   5550
      Left            =   5490
      TabIndex        =   6
      Top             =   45
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   9790
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   5e6
      TextRTF         =   $"frmCodeNavigator.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdExportJSON 
      Caption         =   "Export JSON"
      Height          =   465
      Left            =   11745
      TabIndex        =   5
      Top             =   5715
      Width           =   1455
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   465
      Left            =   585
      TabIndex        =   2
      Top             =   6075
      Width           =   1455
   End
   Begin VB.TextBox txtFilePath 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5640
      Width           =   5340
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse File"
      Height          =   495
      Left            =   2430
      TabIndex        =   0
      Top             =   6075
      Width           =   1335
   End
   Begin MSComctlLib.TreeView tvCode 
      Height          =   5535
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5370
      _ExtentX        =   9472
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
      Left            =   0
      TabIndex        =   4
      Top             =   5400
      Width           =   1335
   End
End
Attribute VB_Name = "frmCodeNavigator"
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
Private txtSource As String

Private Const debug_mode As Boolean = False

Private Sub Form_Load()
    Set parser = New CParser
    
    ' Setup TreeView
    tvCode.Nodes.Clear
    
    ' Setup icons FIRST (or disable them)
    SetupIcons
    
    ' IMPORTANT: If you don't have icons, remove the ImageList reference!
    ' Set tvCode.ImageList = Nothing
    
    ' Example file path
    'txtFilePath.Text = App.Path & "\tests\test.js" 'ok!
    'txtFilePath.Text = App.Path & "\tests\for-in2.txt" 'ok
    'txtFilePath.Text = App.Path & "\tests\for-in.txt"
    'txtFilePath.Text = App.Path & "\tests\es5.txt" 'mostly ok see commented out sections
    txtFilePath.Text = App.Path & "\jquery.js"  'OK!
    
    Me.Visible = True
    DoEvents
    
    cmdParse_Click
    
    Dim n As node
    For Each n In tvCode.Nodes
        n.Expanded = True
    Next
        
    
    'End
    
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
    Dim code As String
    txtSource = Empty
    rtf.Text = Empty
    
    code = ReadTextFile(txtFilePath.Text)
    
    If code = "" Then
        MsgBox "File is empty or could not be read", vbExclamation
        Exit Sub
    End If
    
    rtf.Text = code
    txtSource = code
    ' Parse it!
    Me.MousePointer = vbHourglass
    tvCode.Nodes.Clear
    
    On Error GoTo ParseError
    
    Dim startTime As Single
    startTime = Timer
    
    Dim ast As CNode
    Set ast = parser.ParseScript(code)
    
    Dim parseTime As Single
    parseTime = Timer - startTime
    
    Debug.Print "Parse time: " & Format$(parseTime, "0.00") & " seconds"
    
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
    MsgBox "Parsing complete! Found " & tvCode.Nodes.Count & " nodes." & vbCrLf & vbCrLf & "Parsed in " & Format$(parseTime, "0.00") & " seconds", vbInformation
    Exit Sub
    
ParseError:
    Me.MousePointer = vbDefault
    
    ' Enhanced error reporting
    Dim errMsg As String
    errMsg = "Parse error at line " & Err.Description & vbCrLf & vbCrLf
    errMsg = errMsg & "Error: " & Err.Description & vbCrLf & vbCrLf
    
    ' Try to show context
    Dim lines() As String
    lines = Split(code, vbLf)
    
    ' Extract line number from error message
    Dim lineNum As Long
    Dim matches As Object
    On Error Resume Next
    Set matches = CreateObject("VBScript.RegExp")
    matches.Pattern = "line (\d+)"
    matches.Global = False
    Dim match As Object
    Set match = matches.Execute(Err.Description)
    If match.Count > 0 Then
        lineNum = CLng(match(0).SubMatches(0))
        
        If lineNum > 0 And lineNum <= UBound(lines) + 1 Then
            errMsg = errMsg & "Context:" & vbCrLf
            If lineNum > 1 Then errMsg = errMsg & "Line " & (lineNum - 1) & ": " & lines(lineNum - 2) & vbCrLf
            errMsg = errMsg & "Line " & lineNum & ": " & lines(lineNum - 1) & " <<<" & vbCrLf
            If lineNum < UBound(lines) + 1 Then errMsg = errMsg & "Line " & (lineNum + 1) & ": " & lines(lineNum) & vbCrLf
        End If
    End If
    On Error GoTo 0
    
    'MsgBox errMsg, vbCritical, "Parse Error"
    
    Debug.Print "PARSE ERROR:"
    Debug.Print Err.Description
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


Private Sub BuildTreeView(ast As CNode)
    ' Add root node
    m_recursionDepth = 0
    
    Dim rootNode As node
    Set rootNode = tvCode.Nodes.Add(, , "root", "?? " & ExtractFileName(txtFilePath.Text))
    
    ' First pass: Find all top-level functions and constructors
    Dim i As Long
    For i = 1 To ast.Body.Count
        Dim stmt As CNode
        Set stmt = ast.Body(i)
        ProcessTopLevel stmt, rootNode
    Next i
    
    ' Expand root
    rootNode.Expanded = True
    
    ' Expand all class/object nodes
    Dim node As node
    For Each node In tvCode.Nodes
        If node.Children > 0 And node.Parent Is rootNode Then
            node.Expanded = True
        End If
    Next node
End Sub

Private Function ExtractFileName(fullPath As String) As String
    Dim pos As Long
    pos = InStrRev(fullPath, "\")
    If pos > 0 Then
        ExtractFileName = Mid$(fullPath, pos + 1)
    Else
        ExtractFileName = fullPath
    End If
End Function

Private Sub ProcessTopLevel(node As CNode, parentNode As node)
    If node Is Nothing Then Exit Sub
    
    Dim nodeKey As String
    nodeKey = "node_" & tvCode.Nodes.Count
    
    Select Case node.tType
        Case FunctionDeclaration_Node
            ' Check if this looks like a constructor (starts with capital letter)
            If IsConstructor(node.id.Name) Then
                ' Look ahead to see if there's a prototype assignment coming
                ' For now, just mark it as a class
                AddClassConstructor node, parentNode, nodeKey
            Else
                ' Regular top-level function
                AddFunctionNode node, parentNode, nodeKey
            End If
            
        Case VariableDeclaration_Node
            ' Check each variable declaration
            Dim j As Long
            For j = 1 To node.Declarations.Count
                Dim decl As CNode
                Set decl = node.Declarations(j)
                
                If Not decl.init Is Nothing Then
                    If decl.init.tType = FunctionExpression_Node Then
                        ' var myFunc = function() {}
                        AddFunctionNode decl.init, parentNode, nodeKey & "_" & j, decl.VarID.Name
                        
                    ElseIf decl.init.tType = ObjectExpression_Node Then
                        ' var myObj = { ... }
                        AddObjectNode decl.VarID.Name, decl.init, parentNode, nodeKey & "_" & j
                    End If
                End If
            Next j
            
        Case ExpressionStatement_Node
            If Not node.Test Is Nothing Then
                ' Check for IIFE: (function(){...})()
                If node.Test.tType = CallExpression_Node Then
                    If Not node.Test.Callee Is Nothing Then
                        If node.Test.Callee.tType = FunctionExpression_Node Then
                            ' This is an IIFE - process its body!
                            If Not node.Test.Callee.FunctionBody Is Nothing Then
                                If node.Test.Callee.FunctionBody.tType = BlockStatement_Node Then
                                    ' Process all statements inside the IIFE
                                    Dim k As Long
                                    For k = 1 To node.Test.Callee.FunctionBody.Body.Count
                                        ProcessTopLevel node.Test.Callee.FunctionBody.Body(k), parentNode
                                    Next k
                                End If
                            End If
                            
                            ' ALSO check the arguments being passed to this IIFE
                            ' jQuery passes a factory function as an argument!
                            If Not node.Test.Arguments Is Nothing Then
                                Dim m As Long
                                For m = 1 To node.Test.Arguments.Count
                                    Dim arg As CNode
                                    Set arg = node.Test.Arguments(m)
                                    
                                    ' Check if this argument is a function
                                    If arg.tType = FunctionExpression_Node Then
                                        ' Process the function's body too!
                                        If Not arg.FunctionBody Is Nothing Then
                                            If arg.FunctionBody.tType = BlockStatement_Node Then
                                                Dim n As Long
                                                For n = 1 To arg.FunctionBody.Body.Count
                                                    ProcessTopLevel arg.FunctionBody.Body(n), parentNode
                                                Next n
                                            End If
                                        End If
                                    End If
                                Next m
                            End If
                        End If
                    End If
                ElseIf node.Test.tType = AssignmentExpression_Node Then
                    ProcessAssignment node.Test, parentNode, nodeKey
                End If
            End If
            
        Case BlockStatement_Node
            ' Process children of blocks
            For j = 1 To node.Body.Count
                ProcessTopLevel node.Body(j), parentNode
            Next j
    End Select
End Sub

Private Function IsConstructor(funcName As String) As Boolean
    ' Check if function name starts with capital letter (constructor convention)
    If Len(funcName) = 0 Then
        IsConstructor = False
        Exit Function
    End If
    
    Dim firstChar As String
    firstChar = Left$(funcName, 1)
    IsConstructor = (firstChar >= "A" And firstChar <= "Z")
End Function


Private Sub ProcessAssignment(assignNode As CNode, parentNode As node, nodeKey As String)
    If assignNode Is Nothing Then Exit Sub
    
    Dim leftSide As String
    leftSide = GetExpressionPreview(assignNode.Left, 0)
    
    ' Check for prototype assignments (class methods)
    If InStr(leftSide, ".prototype") > 0 Then
        If Not assignNode.Right Is Nothing Then
            If assignNode.Right.tType = ObjectExpression_Node Then
                ' ClassName.prototype = { methods... }
                Dim className As String
                className = ExtractClassName(leftSide)
                AddClassPrototype className, assignNode.Right, parentNode, nodeKey
            End If
        End If
    ElseIf Not assignNode.Right Is Nothing Then
        ' Regular assignment
        If assignNode.Right.tType = FunctionExpression_Node Then
            ' obj.method = function() {}
            AddFunctionNode assignNode.Right, parentNode, nodeKey, leftSide
        ElseIf assignNode.Right.tType = ObjectExpression_Node Then
            ' obj = { ... }
            AddObjectNode leftSide, assignNode.Right, parentNode, nodeKey
        End If
    End If
End Sub

Private Function ExtractClassName(protoString As String) As String
    ' "Calculator.prototype" -> "Calculator"
    Dim pos As Long
    pos = InStr(protoString, ".prototype")
    If pos > 0 Then
        ExtractClassName = Left$(protoString, pos - 1)
    Else
        ExtractClassName = protoString
    End If
End Function

Private Sub AddClassConstructor(funcNode As CNode, parentNode As node, nodeKey As String)
    ' Don't add the constructor as a separate node
    ' It will be shown when we find the .prototype assignment
    ' Just skip it for now
End Sub

Private Sub AddClassPrototype(className As String, objNode As CNode, parentNode As node, nodeKey As String)
    ' Add class node (line number from first property if available)
    Dim lineInfo As String
    lineInfo = ""
    
    Dim classNode As MSComctlLib.node
    Set classNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, "[Class] " & className & lineInfo)
    classNode.Bold = True
    
    ' Add methods
    If objNode.Properties.Count > 0 Then
        Dim j As Long
        For j = 1 To objNode.Properties.Count
            Dim prop As CNode
            Set prop = objNode.Properties(j)
            
            If Not prop.propValue Is Nothing Then
                If prop.propValue.tType = FunctionExpression_Node Then
                    ' This is a method
                    Dim methodName As String
                    methodName = GetExpressionPreview(prop.PropKey, 0)
                    
                    Dim params As String
                    params = GetParamList(prop.propValue)
                    
                    ' Add line number if available
                    Dim methodLine As String
                    methodLine = ""
                    If prop.propValue.LineNumber > 0 Then
                        methodLine = " [Line " & prop.propValue.LineNumber & "]"
                    End If
                    
                    tvCode.Nodes.Add classNode, tvwChild, nodeKey & "_m" & j, "    " & methodName & "(" & params & ")" & methodLine
                End If
            End If
        Next j
    End If
End Sub

Private Sub AddFunctionNode(funcNode As CNode, parentNode As node, nodeKey As String, Optional customName As String = "")
    Dim funcName As String
    
    If customName <> "" Then
        funcName = customName
    ElseIf Not funcNode.id Is Nothing Then
        funcName = funcNode.id.Name
    Else
        funcName = "(anonymous)"
    End If
    
    Dim params As String
    params = GetParamList(funcNode)
    
    ' DEBUG: Check what line number we have
    If debug_mode Then Debug.Print "AddFunctionNode: " & funcName & " - LineNumber=" & funcNode.LineNumber
    
    ' Add line number if available
    Dim lineInfo As String
    lineInfo = ""
    If funcNode.LineNumber > 0 Then
        lineInfo = " [Line " & funcNode.LineNumber & "]"
    End If
    
    Dim newNode As MSComctlLib.node
    Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, "[Function] " & funcName & "(" & params & ")" & lineInfo)
    newNode.Bold = True
End Sub

Private Sub AddObjectNode(objName As String, objNode As CNode, parentNode As node, nodeKey As String)
    ' Add object node
    Dim lineInfo As String
    lineInfo = ""
    If objNode.LineNumber > 0 Then
        lineInfo = " [Line " & objNode.LineNumber & "]"
    End If
    
    Dim newNode As MSComctlLib.node
    Set newNode = tvCode.Nodes.Add(parentNode, tvwChild, nodeKey, "[Object] " & objName & lineInfo)
    newNode.Bold = True
    
    ' Add properties/methods
    If objNode.Properties.Count > 0 Then
        Dim j As Long
        For j = 1 To objNode.Properties.Count
            Dim prop As CNode
            Set prop = objNode.Properties(j)
            
            Dim propName As String
            propName = GetExpressionPreview(prop.PropKey, 0)
            
            If Not prop.propValue Is Nothing Then
                If prop.propValue.tType = FunctionExpression_Node Then
                    ' Method
                    Dim params As String
                    params = GetParamList(prop.propValue)
                    
                    ' Add line number if available
                    Dim methodLine As String
                    methodLine = ""
                    If prop.propValue.LineNumber > 0 Then
                        methodLine = " [Line " & prop.propValue.LineNumber & "]"
                    End If
                    
                    tvCode.Nodes.Add newNode, tvwChild, nodeKey & "_p" & j, "    " & propName & "(" & params & ")" & methodLine
                    
                ElseIf prop.propValue.tType = ObjectExpression_Node Then
                    ' Nested object
                    AddObjectNode propName, prop.propValue, newNode, nodeKey & "_p" & j
                Else
                    ' Regular property
                    tvCode.Nodes.Add newNode, tvwChild, nodeKey & "_p" & j, "    " & propName
                End If
            End If
        Next j
    End If
End Sub


Private Function GetParamList(funcNode As CNode) As String
    If funcNode.params.Count = 0 Then
        GetParamList = ""
        Exit Function
    End If
    
    Dim params As String
    Dim j As Long
    For j = 1 To funcNode.params.Count
        Dim param As CNode
        Set param = funcNode.params(j)
        params = params & param.Name
        If j < funcNode.params.Count Then params = params & ", "
    Next j
    
    GetParamList = params
End Function


Private Sub cmdExportJSON_Click()
    'On Error GoTo ErrorHandler
    
    Dim source As String
    Dim parser As New CParser
    Dim ast As CNode
    Dim json As String
    
    ' Get source code
    source = txtSource
    If Len(txtSource) = 0 Then
        MsgBox "We have not parsed a file yet", vbInformation
        Exit Sub
    End If
    
    ' Parse
    Me.MousePointer = vbHourglass
    Set ast = parser.ParseScript(source)
    
    ' Generate JSON
    json = ast.ToJSON(0)
    
    ' Display
    rtf.Text = json
    
    Me.MousePointer = vbDefault
    MsgBox "JSON generated successfully!", vbInformation
    Exit Sub
    
ErrorHandler:
    Me.MousePointer = vbDefault
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Private Sub cmdSaveJSON_Click()
    Dim fileNum As Integer
    Dim filePath As String
    
    ' Show save dialog
    With CommonDialog1
        .Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        .DefaultExt = "json"
        .ShowSave
        filePath = .FileName
    End With
    
    If Len(filePath) > 0 Then
        fileNum = FreeFile
        Open filePath For Output As #fileNum
        Print #fileNum, rtf.Text
        Close #fileNum
        
        MsgBox "JSON saved to: " & filePath, vbInformation
    End If
End Sub

Private Sub cmdPrettyPrint_Click()
    ' JSON is already pretty-printed by ToJSON!
    MsgBox "JSON is already formatted with proper indentation!", vbInformation
End Sub



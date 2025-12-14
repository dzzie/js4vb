VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2668C1EA-1D34-42E2-B89F-6B92F3FF627B}#5.0#0"; "scivb2.ocx"
Begin VB.UserControl ucJSDebugger 
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14400
   ScaleHeight     =   8400
   ScaleWidth      =   14400
   Begin sci2.SciSimple scivb 
      Height          =   4500
      Left            =   90
      TabIndex        =   0
      Top             =   450
      Width           =   10800
      _ExtentX        =   19050
      _ExtentY        =   7938
   End
   Begin MSComctlLib.ListView lstVariables 
      Height          =   3195
      Left            =   90
      TabIndex        =   1
      Top             =   5040
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   5636
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Variable"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstCallStack 
      Height          =   3195
      Left            =   5490
      TabIndex        =   2
      Top             =   5040
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   5636
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Function"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Line"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbarDebug 
      Height          =   330
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Object.ToolTipText     =   "Run"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Start Debugger"
            Object.ToolTipText     =   "Start Debugger"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Break"
            Object.ToolTipText     =   "Break"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Toggle Breakpoint"
            Object.ToolTipText     =   "Toggle Breakpoint"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clear All Breakpoints"
            Object.ToolTipText     =   "Clear All Breakpoiunts"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step In"
            Object.ToolTipText     =   "Step In"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step Over"
            Object.ToolTipText     =   "Step Over"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step Out"
            Object.ToolTipText     =   "Step Out"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run to Cursor"
            Object.ToolTipText     =   "Run to Cursor"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilToolbars_Disabled 
      Left            =   7380
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0000
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":010C
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0218
            Key             =   "Run to Cursor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0324
            Key             =   "Clear All Breakpoints"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":042E
            Key             =   "Step Over"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":053A
            Key             =   "Step Out"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0646
            Key             =   "Step In"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0752
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":085E
            Key             =   "Start Debugger"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0968
            Key             =   "Toggle Breakpoint"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilToolbar 
      Left            =   6705
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0A72
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0B7E
            Key             =   "Start Debugger"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0C88
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0D92
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0E9C
            Key             =   "Toggle Breakpoint"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0FA6
            Key             =   "Clear All Breakpoints"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":10B0
            Key             =   "Step In"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":11BA
            Key             =   "Step Over"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":12C4
            Key             =   "Step Out"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":13CE
            Key             =   "Run to Cursor"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":14D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: Ready"
      Height          =   285
      Left            =   4095
      TabIndex        =   3
      Top             =   135
      Width           =   3225
   End
End
Attribute VB_Name = "ucJSDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'' In your java.hilighter or custom highlighter, add these style definitions:
'
'' Style for hex numbers (0x...)
'Style 15 = fore:#007ACC, back:#FFFFFF, bold
'
'' Style for BigInt literals (...n)
'Style 16 = fore:#098658, back:#FFFFFF, bold
'
'' You'll need to update the lexer to recognize these patterns
'' This is typically done in the highlighter definition file

' Scintilla marker constants
Const SC_MARK_CIRCLE = 0
Const SC_MARK_ARROW = 2
Const SC_MARK_BACKGROUND = 22

' Interpreter with events
Private WithEvents interp As CInterpreter
Attribute interp.VB_VarHelpID = -1

' State
Private m_currentFile As String
Private m_isRunning As Boolean
Private m_lastEIPLine As Long
Private m_isInitialized As Boolean
Private m_executionStartTime As Double
Private objCache As New Collection

' Events for host
Event output(msg As String)
Event ErrorOccurred(line As Long, msg As String)

'I will clean these up proper later
Event StateChanged(IsRunning As Boolean, IsPaused As Boolean)
Event ScriptStart()
Event ScriptTerminate()

' ============================================
' PUBLIC INTERFACE
' ============================================

Function SetNextStatement(target As Long)
    SetNextStatement = interp.SetNextStatement(target)
    If SetNextStatement Then UpdateUI
End Function

Sub Abort()
    interp.Abort
End Sub

'only have to configure this once per instance unless you reset
Public Function AddObject(obj As Object, Name As String) As Boolean
    Dim o As CCachedObj
    
    If IsRunning Then Exit Function
    
    For Each o In objCache
        If o.Name = Name Then Exit Function
    Next
    
    Set o = New CCachedObj
    Set o.obj = obj
    o.Name = Name
    objCache.Add o
    AddObject = True
    
End Function

Public Property Get sci() As Object
    Set sci = scivb
End Property

Public Property Get text() As String
    text = scivb.text
End Property

Public Property Let text(v As String)
    scivb.text = v
End Property

Public Property Get IsRunning() As Boolean
    IsRunning = m_isRunning
End Property

Public Property Get CurrentFile() As String
    CurrentFile = m_currentFile
End Property

Public Function LoadFile(fpath As String) As Boolean
    If Not FileExists(fpath) Then Exit Function
    m_currentFile = fpath
    LoadFile = scivb.LoadFile(fpath)
    InitializeScintillaMarkers
End Function

Public Function SaveFile(fpath As String) As Boolean
    SaveFile = scivb.SaveFile(fpath)
    m_currentFile = fpath
End Function

Private Sub SetToolBarIcons(Optional forceDisable As Boolean = False)
    Dim b As Button
    
    If forceDisable Then
        For Each b In tbarDebug.Buttons
            b.Enabled = False
        Next
        Set tbarDebug.ImageList = Nothing
        Set tbarDebug.ImageList = ilToolbars_Disabled
        Exit Sub
    End If
    
    Set tbarDebug.ImageList = Nothing
    Set tbarDebug.ImageList = IIf(m_isRunning, ilToolbar, ilToolbars_Disabled)
    
    For Each b In tbarDebug.Buttons
        If Len(b.Key) > 0 Then
            b.Image = b.Key
            b.ToolTipText = b.Key
            If b.Key <> "Run" And b.Key <> "Start Debugger" And InStr(b.Key, "Breakpoint") < 1 Then
                b.Enabled = m_isRunning
            End If
        End If
    Next
    
End Sub

Private Sub UpdateToolbarState()
    Dim btn As Button
    
    For Each btn In tbarDebug.Buttons
        Select Case btn.Key
            Case "Run"
                ' Run without debugger - only enabled when NOT debugging
                btn.Enabled = Not m_isRunning
                btn.ToolTipText = "Run (no debugger)"
                
            Case "Start Debugger"
                If m_isRunning Then
                    ' Change to "Continue" when debugging
                    btn.ToolTipText = "Continue (F5)"
                    'If interp Is Nothing Then
                    '    btn.Enabled = False
                    'Else
                    '    btn.Enabled = interp.IsPaused  ' Only when paused
                    'End If
                Else
                    ' "Start Debugger" when not running
                    btn.ToolTipText = "Start Debugger (F5)"
                    btn.Enabled = True
                End If
                
            Case "Break"
                If interp Is Nothing Then
                    btn.Enabled = False
                Else
                    btn.Enabled = m_isRunning And Not interp.IsPaused
                End If
                
            Case "Stop"
                btn.Enabled = m_isRunning
                
            Case "Step In", "Step Over", "Step Out"
                ' Only enabled when paused
                If interp Is Nothing Then
                    btn.Enabled = False
                Else
                    btn.Enabled = m_isRunning And interp.IsPaused
                End If
                
            Case "Continue", "Run to Cursor"  ' ADD THESE
                If interp Is Nothing Then
                    btn.Enabled = False
                Else
                    btn.Enabled = m_isRunning And interp.IsPaused
                End If
                
            Case "Toggle Breakpoint", "Clear All Breakpoints"
                btn.Enabled = True
        End Select
    Next
    
    SetToolBarIcons
End Sub

' ============================================
' EXECUTION CONTROL
' ============================================

Public Sub RunScript()
    
    Dim o As Object
    
    If m_isRunning Then Exit Sub
    
    m_isRunning = True
    lblStatus.Caption = "Status: Running..."
    UpdateToolbarState
    SetToolBarIcons
   ' lblStatus = "Status: " & IIf(withDebugger, "Debugging...", "Running...")

    Set interp = New CInterpreter
    
    For Each o In objCache
         interp.AddCOMObject o.Name, o.obj
    Next
    
    On Error Resume Next
    RaiseEvent ScriptStart
    interp.AddCode scivb.text
    
    If Err.Number <> 0 Then
        RaiseEvent ErrorOccurred(0, Err.Description)
        lblStatus.Caption = "Status: Error"
        m_isRunning = False
        UpdateToolbarState
        Exit Sub
    End If
    
    lblStatus.Caption = "Status: Complete"
    m_isRunning = False
    UpdateToolbarState
    RaiseEvent StateChanged(False, False)
    RaiseEvent ScriptTerminate
End Sub

Public Sub StartDebugging()
    Dim o As Object
    
    If m_isRunning Then Exit Sub
    
    m_isRunning = True
    m_executionStartTime = Timer  ' Record start time
    lblStatus.Caption = "Status: Debugging..."
    UpdateToolbarState
    scivb.ReadOnly = True
    
    Set interp = New CInterpreter
        
    For Each o In objCache
         interp.AddCOMObject o.Name, o.obj
    Next
    
    ' Copy breakpoints from UI to interpreter
    Dim i As Long
    For i = 0 To scivb.TotalLines - 1
        If HasBreakpointMarker(i) Then
            interp.SetBreakpoint i + 1  ' 1-based in interpreter
        End If
    Next
    
    On Error Resume Next
    RaiseEvent ScriptStart
    interp.Execute scivb.text, True
    
    If Err.Number <> 0 Then
        RaiseEvent ErrorOccurred(0, Err.Description)
        StopDebugging
        Exit Sub
    End If
    
    ' Execution paused at first line or breakpoint
    UpdateUI
    RaiseEvent ScriptTerminate
End Sub

Public Sub StopDebugging()
    On Error Resume Next  ' ADD ERROR HANDLING
    
    If interp Is Nothing Then Exit Sub
    
    If Not m_isRunning Then Exit Sub
    
    
    If interp.IsDebugging Then interp.StopDebug
    interp.Abort
    
    Set interp = Nothing
     
    ClearCurrentLineMarker
    scivb.ReadOnly = False
    m_isRunning = False
    
    lblStatus.Caption = "Status: Ready"
    UpdateToolbarState
    RaiseEvent StateChanged(False, False)
    scivb.SelLength = 0
    
End Sub

Public Sub StepInto()
    If Not m_isRunning Then Exit Sub
    If interp Is Nothing Then Exit Sub
    interp.StepInto
End Sub

Public Sub StepOver()
    If Not m_isRunning Then Exit Sub
    If interp Is Nothing Then Exit Sub
    interp.StepOver
End Sub

Public Sub StepOut()
    If Not m_isRunning Then Exit Sub
    If interp Is Nothing Then Exit Sub
    interp.StepOut
End Sub

Public Sub Continue()
    If Not m_isRunning Then Exit Sub
    If interp Is Nothing Then Exit Sub
    
    lblStatus.Caption = "Status: Running..."
    UpdateToolbarState
    
    interp.Run  ' This sets smNone and unpauses
End Sub

Public Sub Cleanup()
    On Error Resume Next
    
    ' Stop any running script
    If m_isRunning Then
        StopDebugging
    End If
    
    ' Clear the editor
    If Not scivb Is Nothing Then
        scivb.text = ""
    End If
    
    ' Release interpreter
    Set interp = Nothing
End Sub

' ============================================
' BREAKPOINT MANAGEMENT
' ============================================

Public Sub ToggleBreakpoint(Optional line As Long = -1)
    If line = -1 Then line = scivb.currentLine
    
    If HasBreakpointMarker(line) Then
        scivb.DeleteMarker line, 2
        If Not interp Is Nothing Then
            interp.ClearBreakpoint line + 1
        End If
    Else
        scivb.SetMarker line, 2
        If Not interp Is Nothing Then
            interp.SetBreakpoint line + 1
        End If
    End If
End Sub

Public Sub ClearAllBreakpoints()
    scivb.DeleteAllMarkers 2
    If Not interp Is Nothing Then
        interp.ClearAllBreakpoints
    End If
End Sub

Private Function HasBreakpointMarker(line As Long) As Boolean
    Dim markers As Long
    markers = scivb.DirectSCI.MarkerGet(line)
    HasBreakpointMarker = ((markers And 4) <> 0)  ' Bit 2 for marker 2
End Function

' ============================================
' UI UPDATES
' ============================================

Private Sub UpdateUI()

    'Debug.Print ">>> UpdateUI called"  ' ADD THIS
    
    If interp Is Nothing Then
        'Debug.Print ">>> interp is Nothing!"  ' ADD THIS
        Exit Sub
    End If
    
    If Not interp.IsDebugging Then
        'Debug.Print ">>> Not debugging!"  ' ADD THIS
        Exit Sub
    End If
    
    If interp Is Nothing Then Exit Sub
    If Not interp.IsDebugging Then Exit Sub
    
    ' Update current line
    ClearCurrentLineMarker
    
    Dim curLine As Long
    curLine = interp.currentLine - 1  ' 0-based for Scintilla
    
    'Debug.Print ">>> Setting markers on line: " & curLine  ' ADD THIS
    
    scivb.SetMarker curLine, 1  ' Arrow
    scivb.SetMarker curLine, 3  ' Background highlight
    m_lastEIPLine = curLine
    
    scivb.GotoLineCentered curLine + 1, False
    
    ' Update variables
    UpdateVariablesView
    
    ' Update call stack
    UpdateCallStackView
    
    ' Update status
    If interp.IsPaused Then
        lblStatus.Caption = "Status: Paused at line " & interp.currentLine
    End If
    
    DoEvents
End Sub

Private Sub ClearCurrentLineMarker()
    If m_lastEIPLine >= 0 Then
        scivb.DeleteMarker m_lastEIPLine, 1
        scivb.DeleteMarker m_lastEIPLine, 3
        
        ' Force refresh
        Dim StartPos As Long, EndPos As Long
        StartPos = scivb.PositionFromLine(m_lastEIPLine)
        EndPos = scivb.PositionFromLine(m_lastEIPLine + 1)
        scivb.DirectSCI.Colourise StartPos, EndPos
    End If
End Sub

Private Sub UpdateVariablesView()
    lstVariables.ListItems.Clear
    
    If interp Is Nothing Then Exit Sub
    
    Dim vars As Collection, li As ListItem, vi As CVarItem
    Set vars = interp.GetCurrentScopeVariables() 'of cvaritem

    For Each vi In vars
        Set li = lstVariables.ListItems.Add(, , vi.Name)
        li.SubItems(1) = vi.value.ToString()
        li.SubItems(2) = vi.value.GetTypeName()
    Next
    
End Sub

Private Sub UpdateCallStackView()

    lstCallStack.ListItems.Clear
    
    If interp Is Nothing Then Exit Sub
    
    Dim stack() As String
    stack = interp.GetCallStackStrings()
    
    Dim i As Long
    For i = LBound(stack) To UBound(stack)
        Dim funcName As String, lineNum As String
        ParseCallStackItem stack(i), funcName, lineNum
        lstCallStack.ListItems.Add , , funcName
        lstCallStack.ListItems(lstCallStack.ListItems.Count).SubItems(1) = lineNum
    Next
    
End Sub


Private Sub ParseCallStackItem(item As String, ByRef funcName As String, ByRef lineNum As String)
    ' Format: "functionName (line 123)"
    Dim pos As Long
    pos = InStr(item, " (line ")
    If pos > 0 Then
        funcName = Left$(item, pos - 1)
        lineNum = Mid$(item, pos + 7)
        lineNum = Left$(lineNum, Len(lineNum) - 1)  ' Remove trailing )
    Else
        funcName = item
        lineNum = ""
    End If
End Sub



Private Sub interp_ConsoleLog(ByVal msg As String)
    RaiseEvent output(msg)
End Sub

' ============================================
' INTERPRETER EVENT HANDLERS
' ============================================

Private Sub interp_OnBreakpoint(ByVal LineNumber As Long, ByVal sourceCode As String)
    UpdateUI
    RaiseEvent StateChanged(True, True)
End Sub

Private Sub interp_OnError(ByVal ErrorMessage As String, ByVal LineNumber As Long, ByVal source As String, ByVal col As Long)
    RaiseEvent ErrorOccurred(LineNumber, ErrorMessage)
    StopDebugging
End Sub

Private Sub interp_OnStep(ByVal LineNumber As Long, ByVal sourceCode As String)
    UpdateUI
    RaiseEvent StateChanged(True, True)
End Sub

Private Sub interp_OnCallStackChanged()
    UpdateCallStackView
End Sub

Private Sub interp_OnVariablesChanged()
    UpdateVariablesView
End Sub

 

Private Sub interp_OnExecutionComplete()

    Dim elapsed As Double
    elapsed = Timer - m_executionStartTime
    
    lblStatus.Caption = "Status: Complete (" & Format$(elapsed, "0.000") & "s)"
    
    ' Show completion with timing
    RaiseEvent output(vbCrLf & "=== Script execution completed in " & _
                     Format$(elapsed, "0.000") & " seconds ===" & vbCrLf)

    StopDebugging
    
End Sub



' ============================================
' SCINTILLA EVENT HANDLERS
' ============================================

Private Sub scivb_KeyDown(KeyCode As Long, Shift As Long)
    Select Case KeyCode
        Case vbKeyF2: ToggleBreakpoint
        Case vbKeyF5: If m_isRunning Then Continue Else StartDebugging
        Case vbKeyF7: StepInto
        Case vbKeyF8: StepOver
        Case vbKeyF9: StepOut
    End Select
End Sub

Private Sub scivb_MarginClick(lline As Long, position As Long, margin As Long, modifiers As Long)
    If margin = 1 Then  ' Margin 1 is for breakpoints
        ToggleBreakpoint lline
    End If
End Sub

' ============================================
' TOOLBAR HANDLERS
' ============================================
Private Sub tbarDebug_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Run"
            ' Always run without debugger
            RunScript
            
        Case "Start Debugger"
            ' Dual purpose: Start debugging OR Continue
            If m_isRunning Then
                ' Already debugging - act as Continue
                Continue
            Else
                ' Not running - start debugging
                StartDebugging
            End If
            
        Case "Break"
            ' Pause execution (if running without breakpoint)
            If m_isRunning And Not interp Is Nothing Then
                interp.PauseExecution
                lblStatus.Caption = "Status: Paused (manual break)"
                UpdateUI
                UpdateToolbarState
            End If
            
        Case "Stop"
            StopDebugging
            
        Case "Step In"
            StepInto
            
        Case "Step Over"
            StepOver
            
        Case "Step Out"
            StepOut
            
        Case "Run to Cursor"
            ' TODO: Implement run to cursor
            Continue  ' For now, just continue
            
        Case "Toggle Breakpoint"
            ToggleBreakpoint
            
        Case "Clear All Breakpoints"
            ClearAllBreakpoints
    End Select
End Sub

' ============================================
' INITIALIZATION
' ============================================

Private Sub UserControl_Initialize()
    SetToolBarIcons
    m_lastEIPLine = -1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode And Not m_isInitialized Then
        InitializeScintillaMarkers
        scivb.ReadOnly = False
        m_isInitialized = True
    End If
End Sub

Private Sub InitializeScintillaMarkers()
    ' Breakpoint marker (red circle)
    scivb.DirectSCI.MarkerDefine 2, SC_MARK_CIRCLE
    scivb.DirectSCI.MarkerSetFore 2, vbRed
    scivb.DirectSCI.MarkerSetBack 2, vbRed
    
    ' Current line arrow (yellow)
    scivb.DirectSCI.MarkerDefine 1, SC_MARK_ARROW
    scivb.DirectSCI.MarkerSetFore 1, vbBlack
    scivb.DirectSCI.MarkerSetBack 1, vbYellow
    
    ' Current line background
    scivb.DirectSCI.MarkerDefine 3, SC_MARK_BACKGROUND
    scivb.DirectSCI.MarkerSetFore 3, vbBlack
    scivb.DirectSCI.MarkerSetBack 3, vbYellow
    
    ' Setup margins
    scivb.DirectSCI.SetMarginWidthN 0, 40  ' Line numbers
    scivb.DirectSCI.SetMarginWidthN 1, 16  ' Breakpoints
    scivb.DirectSCI.SetMarginMaskN 1, 4    ' Marker 2 (breakpoints)
    scivb.DirectSCI.SetMarginSensitiveN 1, True
    
    ' IDA address marker (blue circle)
    scivb.DirectSCI.MarkerDefine 4, SC_MARK_CIRCLE
    scivb.DirectSCI.MarkerSetFore 4, vbBlue
    scivb.DirectSCI.MarkerSetBack 4, vbBlue
    
    scivb.LineNumbers = True
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    ' Resize Scintilla
    scivb.Width = UserControl.Width - scivb.Left - 100
    scivb.Height = (UserControl.Height - scivb.Top) * 0.55
    
    ' Position bottom panels
    Dim panelTop As Long
    panelTop = scivb.Top + scivb.Height + 90
    
    lstVariables.Top = panelTop
    lstVariables.Width = (UserControl.Width - 200) / 2
    lstVariables.Height = UserControl.Height - panelTop - 90
    
    lstCallStack.Left = lstVariables.Left + lstVariables.Width + 90
    lstCallStack.Top = panelTop
    lstCallStack.Width = lstVariables.Width
    lstCallStack.Height = lstVariables.Height
End Sub

Private Function FileExists(fpath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir$(fpath) <> "")
End Function

Private Sub UserControl_Terminate()
    On Error Resume Next
    
    ' Stop debugging if still running
    If m_isRunning Then
        StopDebugging
    End If
    
    ' Detach subclassing
    'If Not SC Is Nothing Then
    '    SC.UnSubAll
    '    Set SC = Nothing
    'End If
    
    ' Cleanup interpreter
    Set interp = Nothing
    
End Sub

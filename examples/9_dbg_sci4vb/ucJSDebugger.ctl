VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{61ECB1F7-F440-419B-A75C-7EFDF0889185}#4.0#0"; "sci4vb.ocx"
Begin VB.UserControl ucJSDebugger 
   ClientHeight    =   8400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14400
   ScaleHeight     =   8400
   ScaleWidth      =   14400
   ToolboxBitmap   =   "ucJSDebugger.ctx":0000
   Begin sci4vb.SciWrapper scivb 
      Height          =   4395
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7752
   End
   Begin MSComctlLib.ListView lstVariables 
      Height          =   3195
      Left            =   90
      TabIndex        =   0
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
      Height          =   105
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   185
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
            Picture         =   "ucJSDebugger.ctx":0312
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":041E
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":052A
            Key             =   "Run to Cursor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0636
            Key             =   "Clear All Breakpoints"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0740
            Key             =   "Step Over"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":084C
            Key             =   "Step Out"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0958
            Key             =   "Step In"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0A64
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0B70
            Key             =   "Start Debugger"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0C7A
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
            Picture         =   "ucJSDebugger.ctx":0D84
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0E90
            Key             =   "Start Debugger"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":0F9A
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":10A4
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":11AE
            Key             =   "Toggle Breakpoint"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":12B8
            Key             =   "Clear All Breakpoints"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":13C2
            Key             =   "Step In"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":14CC
            Key             =   "Step Over"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":15D6
            Key             =   "Step Out"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":16E0
            Key             =   "Run to Cursor"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucJSDebugger.ctx":17EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: Ready"
      Height          =   285
      Left            =   4095
      TabIndex        =   2
      Top             =   135
      Width           =   3225
   End
End
Attribute VB_Name = "ucJSDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=============================================================================
' JavaScript Debugger UserControl - DROP-IN REPLACEMENT
' Ported from scivb2.ocx to sci4vb.ocx
'=============================================================================
Option Explicit

' Member variables
Private WithEvents interp As CInterpreter   ' JavaScript interpreter
Attribute interp.VB_VarHelpID = -1
Private m_isRunning As Boolean              ' True when script is executing
Private m_isInitialized As Boolean          ' Initialization guard
Private m_executionStartTime As Double      ' For timing
Private m_currentFile As String             ' Current file path
Private objCache As New Collection          ' Cached COM objects

' Events
Public Event output(ByVal msg As String)
Public Event StateChanged(ByVal isDebugging As Boolean, ByVal isPaused As Boolean)
Public Event ErrorOccurred(ByVal LineNumber As Long, ByVal ErrorMessage As String)
Public Event ScriptStart()
Public Event ScriptTerminate()

'=============================================================================
' PUBLIC INTERFACE
'=============================================================================

' Expose underlying Scintilla control
Public Property Get sci() As Object
    Set sci = scivb
End Property

' Add COM object to interpreter
Public Function AddObject(obj As Object, Name As String) As Boolean
    Dim o As CCachedObj
    
    If m_isRunning Then Exit Function
    
    ' Check if already exists
    On Error Resume Next
    For Each o In objCache
        If o.Name = Name Then Exit Function
    Next
    On Error GoTo 0
    
    Set o = New CCachedObj
    Set o.obj = obj
    o.Name = Name
    objCache.Add o
    AddObject = True
End Function

' Set next statement to execute
Public Function SetNextStatement(target As Long) As Boolean
    SetNextStatement = interp.SetNextStatement(target)
    If SetNextStatement Then UpdateUI
End Function

' Abort execution
Public Sub Abort()
    On Error Resume Next
    If Not interp Is Nothing Then interp.Abort
End Sub

' Get running state
Public Property Get IsRunning() As Boolean
    IsRunning = m_isRunning
End Property

' Get/Set current file
Public Property Get CurrentFile() As String
    CurrentFile = m_currentFile
End Property

' Load file into editor
Public Function LoadFile(fpath As String) As Boolean
    On Error GoTo ErrorHandler
    If Not FileExists(fpath) Then Exit Function
    
    m_currentFile = fpath
    LoadFile = scivb.doc.LoadFile(fpath)
    Exit Function
    
ErrorHandler:
    LoadFile = False
End Function

' Save editor to file
Public Function SaveFile(fpath As String) As Boolean
    On Error GoTo ErrorHandler
    SaveFile = scivb.doc.SaveFile(fpath)
    m_currentFile = fpath
    Exit Function
    
ErrorHandler:
    SaveFile = False
End Function

Public Property Get Text() As String
    Text = scivb.doc.Text
End Property

Public Property Let Text(ByVal newText As String)
    scivb.doc.Text = newText
    scivb.Style.Colorise 0, scivb.doc.TextLength
End Property

Public Sub LoadScript(ByVal sourceCode As String)
    scivb.doc.Text = sourceCode
    scivb.Style.Colorise 0, scivb.doc.TextLength
    lblStatus.Caption = "Status: Ready"
End Sub

'straight run no debugger
Public Sub RunScript()
    Dim o As CCachedObj

    If Len(Trim$(scivb.doc.Text)) = 0 Then
        MsgBox "No script to execute.", vbExclamation
        Exit Sub
    End If
    
    If m_isRunning Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    m_isRunning = True
    m_executionStartTime = Timer
    lblStatus.Caption = "Status: Running..."
    UpdateToolbarState
    
    Set interp = New CInterpreter
    
    For Each o In objCache
        interp.AddCOMObject o.Name, o.obj
    Next
    
    interp.AddCode "function x(){alert(1)}"
    
    RaiseEvent ScriptStart
    interp.Execute scivb.doc.Text
    
    Dim elapsed As Double
    elapsed = Timer - m_executionStartTime
    lblStatus.Caption = "Status: Complete (" & Format$(elapsed, "0.000") & "s)"
    RaiseEvent output(vbCrLf & "Script execution completed in " & Format$(elapsed, "0.000") & " seconds" & vbCrLf)
    
    m_isRunning = False
    UpdateToolbarState
    RaiseEvent StateChanged(False, False)
    RaiseEvent ScriptTerminate
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Status: Error"
    RaiseEvent output("ERROR: " & Err.Description & vbCrLf)
    RaiseEvent ErrorOccurred(0, Err.Description)
    m_isRunning = False
    UpdateToolbarState
    MsgBox "Error executing script: " & Err.Description, vbCritical
End Sub

Public Sub StartDebugging()
    Dim o As CCachedObj
    Dim i As Long
    Dim bps As Collection, bp
    
    If Len(Trim$(scivb.doc.Text)) = 0 Then
        MsgBox "No script to debug.", vbExclamation
        Exit Sub
    End If
    
    If m_isRunning Then Exit Sub
    
    On Error GoTo ErrorHandler
    
    m_executionStartTime = Timer
    lblStatus.Caption = "Status: Debugging..."
    m_isRunning = True
    scivb.doc.ReadOnly = True
    
    Set interp = New CInterpreter
    
    For Each o In objCache
        interp.AddCOMObject o.Name, o.obj
    Next
    
    'sci4vb keeps the master list of breakpoints
    'and auto handles margin clicks
    For Each bp In scivb.breakpoints.getAll()
        interp.SetBreakpoint bp + 1
    Next

    UpdateToolbarState
    RaiseEvent ScriptStart
    interp.Execute scivb.doc.Text, True
    
    UpdateUI
    RaiseEvent StateChanged(True, True)
    RaiseEvent ScriptTerminate
    
    Exit Sub
    
ErrorHandler:
    lblStatus.Caption = "Status: Error"
    RaiseEvent output("ERROR: " & Err.Description & vbCrLf)
    RaiseEvent ErrorOccurred(0, Err.Description)
    MsgBox "Error starting debugger: " & Err.Description, vbCritical
    StopDebugging
End Sub

Public Sub StopDebugging()
    On Error Resume Next
    
    If Not m_isRunning Then Exit Sub
    
    If Not interp Is Nothing Then
        interp.StopDebug
    End If
    
    m_isRunning = False
    scivb.doc.ClearLastExec
    scivb.doc.ReadOnly = False
    
    lstVariables.ListItems.Clear
    lstCallStack.ListItems.Clear
    
    lblStatus.Caption = "Status: Stopped"
    UpdateToolbarState
    RaiseEvent StateChanged(False, False)
End Sub

Public Sub StepInto()
    If Not m_isRunning Then StartDebugging: Exit Sub
    On Error Resume Next
    interp.StepInto
End Sub

Public Sub StepOver()
    If Not m_isRunning Then StartDebugging: Exit Sub
    On Error Resume Next
    interp.StepOver
End Sub

Public Sub StepOut()
    If Not m_isRunning Then StartDebugging: Exit Sub
    On Error Resume Next
    interp.StepOut
End Sub

Public Sub Continue()
    If Not m_isRunning Then StartDebugging: Exit Sub
    On Error Resume Next
    interp.Run
End Sub

Public Sub Cleanup()
    On Error Resume Next
    
    ' Stop any running script
    If m_isRunning Then
        StopDebugging
    End If
    
    ' Clear the editor
    If Not scivb Is Nothing Then
        scivb.doc.Text = ""
    End If
    
    ' Release interpreter
    Set interp = Nothing
End Sub

'this is coming from the toolbar or F2 key,
'do not do this on marginclick the scivb now handles this internally..
Public Sub ToggleBreakpoint(Optional LineNumber As Long = -1)
    Dim line As Long
    
    If LineNumber = -1 Then
        line = scivb.sel.currentLine
    Else
        line = LineNumber
    End If

    If scivb.breakpoints.isSet(line) Then
        scivb.breakpoints.Remove line
        If Not interp Is Nothing Then interp.ClearBreakpoint line + 1 'if we are running
    Else
        scivb.breakpoints.Add line
        If Not interp Is Nothing Then interp.SetBreakpoint line + 1 'if we are running
    End If

End Sub

Public Sub ClearAllBreakpoints()
    scivb.breakpoints.clearAll
    If Not interp Is Nothing Then interp.ClearAllBreakpoints
End Sub

'=============================================================================
' PRIVATE HELPERS
'=============================================================================

Private Sub UpdateUI()
    On Error Resume Next
    
    If interp Is Nothing Then Exit Sub
    If Not interp.isDebugging Then Exit Sub
    
    Dim currentLine As Long
    currentLine = interp.currentLine - 1
    
    If currentLine >= 0 Then
       scivb.doc.SetExecLine currentLine
    End If
    
    UpdateVariablesView
    UpdateCallStackView
    
    If interp.isPaused Then
        lblStatus.Caption = "Status: Paused at line " & interp.currentLine
    End If
    
    DoEvents
End Sub

Private Sub UpdateVariablesView()
    On Error Resume Next
    
    lstVariables.ListItems.Clear
    
    If interp Is Nothing Then Exit Sub
    
    Dim vars As Collection
    Set vars = interp.GetCurrentScopeVariables()
    
    If vars Is Nothing Then Exit Sub
    
    Dim varInfo As String
    Dim i As Long
    Dim item As ListItem
    
    For i = 1 To vars.Count
        varInfo = vars(i)
        Set item = lstVariables.ListItems.Add(, , ParseVarName(varInfo))
        item.SubItems(1) = ParseVarValue(varInfo)
        item.SubItems(2) = ParseVarType(varInfo)
    Next i
End Sub

Private Function ParseVarName(varInfo As String) As String
    Dim pos As Long
    pos = InStr(varInfo, " = ")
    If pos > 0 Then
        ParseVarName = Left$(varInfo, pos - 1)
    Else
        ParseVarName = varInfo
    End If
End Function

Private Function ParseVarValue(varInfo As String) As String
    Dim pos1 As Long, pos2 As Long
    pos1 = InStr(varInfo, " = ")
    pos2 = InStr(varInfo, " (")
    If pos1 > 0 And pos2 > pos1 Then
        ParseVarValue = Mid$(varInfo, pos1 + 3, pos2 - pos1 - 3)
    Else
        ParseVarValue = ""
    End If
End Function

Private Function ParseVarType(varInfo As String) As String
    Dim pos1 As Long, pos2 As Long
    pos1 = InStr(varInfo, " (")
    pos2 = InStrRev(varInfo, ")")
    If pos1 > 0 And pos2 > pos1 Then
        ParseVarType = Mid$(varInfo, pos1 + 2, pos2 - pos1 - 2)
    Else
        ParseVarType = ""
    End If
End Function

Private Sub UpdateCallStackView()
    On Error Resume Next
    
    lstCallStack.ListItems.Clear
    
    If interp Is Nothing Then Exit Sub
    
    Dim stack() As String
    stack = interp.GetCallStackStrings()
    
    Dim i As Long
    Dim item As ListItem
    
    For i = LBound(stack) To UBound(stack)
        Dim funcName As String, lineNum As String
        ParseCallStackItem stack(i), funcName, lineNum
        
        Set item = lstCallStack.ListItems.Add(, , funcName)
        item.SubItems(1) = lineNum
    Next i
End Sub

Private Sub ParseCallStackItem(item As String, ByRef funcName As String, ByRef lineNum As String)
    Dim pos As Long
    pos = InStr(item, " (line ")
    If pos > 0 Then
        funcName = Left$(item, pos - 1)
        lineNum = Mid$(item, pos + 7)
        lineNum = Left$(lineNum, Len(lineNum) - 1)
    Else
        funcName = item
        lineNum = ""
    End If
End Sub


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
        If Len(b.key) > 0 Then
            b.Image = b.key
            b.ToolTipText = b.key
            If b.key <> "Run" And b.key <> "Start Debugger" And InStr(b.key, "Breakpoint") < 1 Then
                b.Enabled = m_isRunning
            End If
        End If
    Next
    
End Sub

Private Sub UpdateToolbarState()
    Dim btn As Button
    
    For Each btn In tbarDebug.Buttons
        Select Case btn.key
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
                    btn.Enabled = m_isRunning And Not interp.isPaused
                End If
                
            Case "Stop"
                btn.Enabled = m_isRunning
                
            Case "Step In", "Step Over", "Step Out"
                ' Only enabled when paused
                If interp Is Nothing Then
                    btn.Enabled = False
                Else
                    btn.Enabled = m_isRunning And interp.isPaused
                End If
                
            Case "Continue", "Run to Cursor"  ' ADD THESE
                If interp Is Nothing Then
                    btn.Enabled = False
                Else
                    btn.Enabled = m_isRunning And interp.isPaused
                End If
                
            Case "Toggle Breakpoint", "Clear All Breakpoints"
                btn.Enabled = True
        End Select
    Next
    
    SetToolBarIcons
End Sub

'=============================================================================
' INTERPRETER EVENT HANDLERS
'=============================================================================

Private Sub interp_ConsoleLog(ByVal msg As String)
    RaiseEvent output(msg)
End Sub

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
    
    RaiseEvent output(vbCrLf & "=== Script execution completed in " & _
                     Format$(elapsed, "0.000") & " seconds ===" & vbCrLf)

    StopDebugging
End Sub



Private Sub scivb_KeyPressed(key As Long, modifiers As Long, handled As Boolean)
    'Debug.Print "scivb_KeyPressed: " & key & " mod: " & modifiers
    Select Case key
        Case vbKeyF2: ToggleBreakpoint
        Case vbKeyF5: If m_isRunning Then Continue Else StartDebugging
        Case vbKeyF7: StepInto
        Case vbKeyF8: StepOver
        Case vbKeyF9: StepOut
    End Select
End Sub

'user set a breakpoint by margin click in scivb
'we only care if the script is running, scivb keeps the master list
Private Sub scivb_UserBreakpointToggle(line As Long, isAdding As Boolean, Cancel As Boolean)
    If interp Is Nothing Then Exit Sub
    If isAdding Then
        interp.SetBreakpoint line + 1
    Else
        interp.ClearBreakpoint line + 1
    End If
End Sub

'=============================================================================
' TOOLBAR HANDLERS
'=============================================================================

Private Sub tbarDebug_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "Run"
            RunScript
        Case "Start Debugger"
            If m_isRunning Then Continue Else StartDebugging
        Case "Break"
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
            Continue
        Case "Toggle Breakpoint"
            ToggleBreakpoint
        Case "Clear All Breakpoints"
            ClearAllBreakpoints
            
    End Select
End Sub

'=============================================================================
' INITIALIZATION
'=============================================================================

Private Sub UserControl_Initialize()
    SetToolBarIcons
End Sub
 
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "uc kd: " & KeyCode
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    If Ambient.UserMode And Not m_isInitialized Then
        'InitializeScintilla
        scivb.doc.ReadOnly = False
        m_isInitialized = True
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    
    scivb.Width = UserControl.Width - scivb.Left - 100
    scivb.Height = (UserControl.Height - scivb.Top) * 0.55
    
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
    
    If m_isRunning Then
        StopDebugging
    End If
    
    Set interp = Nothing
End Sub



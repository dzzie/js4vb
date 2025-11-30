VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDebugger 
   Caption         =   "Form2"
   ClientHeight    =   11085
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   12855
   LinkTopic       =   "Form2"
   ScaleHeight     =   11085
   ScaleWidth      =   12855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   420
      Left            =   9900
      TabIndex        =   4
      Top             =   10620
      Width           =   1050
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Height          =   420
      Left            =   11205
      TabIndex        =   3
      Top             =   10620
      Width           =   1185
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   8775
      Width           =   12480
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   7470
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSetNext 
      Caption         =   "Set Next"
      Height          =   375
      Left            =   10890
      TabIndex        =   1
      Top             =   135
      Width           =   1005
   End
   Begin Project1.ucJSDebugger js 
      Height          =   8430
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   12390
      _ExtentX        =   21855
      _ExtentY        =   16140
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "frmDebugger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim last As String

Private Sub cmdClear_Click()
    txtOut.Text = Empty
End Sub

Private Sub cmdCopy_Click()
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText txtOut.Text
End Sub

Private Sub cmdSetNext_Click()
    Dim x As Long
    On Error Resume Next
    x = InputBox("Set next statement: ")
    If Err.Number = 0 Then
       If Not js.SetNextStatement(x) Then MsgBox "Failed to set line to " & x, vbExclamation
    Else
        MsgBox "Not number? " & Err.Description
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Cleanup debugger before closing
    On Error Resume Next

    ' Ask user if script is running
    If js.IsRunning Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Script is still running. Stop and exit?", _
                         vbYesNo + vbQuestion, "Confirm Exit")
        
        If response = vbNo Then
            Cancel = True  ' Cancel the unload
            Exit Sub
        End If
        
        ' Force stop
        js.StopDebugging
    End If
    
    js.Abort
    
End Sub

 

Private Sub js_ScriptStart()
    txtOut = Empty
End Sub

Private Sub js_output(ByVal msg As String)
     txtOut = txtOut & msg & vbCrLf
End Sub

 


Private Sub Form_Load()

    On Error Resume Next
    
    last = App.path & "\lastScript.txt"
    If FileExists(last) Then
        js.Text = ReadFile(last)
    End If
    
    js.AddObject Me, "frmDebugger"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    WriteFile last, js.Text
    js.Cleanup
    DoEvents
End Sub


Function FileExists(path As String) As Boolean
  On Error GoTo hell
    
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function

Function ReadFile(filename)
  Dim f As Long, temp
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Sub WriteFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub


Private Sub js_StateChanged(ByVal isDebugging As Boolean, ByVal isPaused As Boolean)
    On Error Resume Next
    WriteFile last, js.Text
End Sub

Private Sub mnuOpen_Click()
    Dim f As String
    dlg.InitDir = App.path
    dlg.ShowOpen
    f = dlg.filename
    If FileExists(f) Then js.Text = ReadFile(f)
End Sub

Private Sub mnuSaveAs_Click()
    Dim f As String
    dlg.InitDir = App.path
    dlg.ShowSave
    f = dlg.filename
    WriteFile f, js.Text
End Sub

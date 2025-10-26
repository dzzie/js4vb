VERSION 5.00
Begin VB.Form frmParserDemo 
   Caption         =   "JavaScript Parser Demo"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15960
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   15960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse JavaScript"
      Height          =   495
      Left            =   13095
      TabIndex        =   4
      Top             =   6435
      Width           =   2400
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6090
      Left            =   4905
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   225
      Width           =   10575
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6045
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   240
      Width           =   4725
   End
   Begin VB.Label Label2 
      Caption         =   "Output (AST JSON or Error):"
      Height          =   255
      Left            =   5175
      TabIndex        =   2
      Top             =   0
      Width           =   2520
   End
   Begin VB.Label Label1 
      Caption         =   "JavaScript Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmParserDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' frmParserDemo - JavaScript Parser Demo Form
'
' Simple demo showing the parser in action
'------------------------------------------------------------

Option Explicit

Private Sub Form_Load()
    ' Set default example code
    txtInput.Text = "// Try some JavaScript!" & vbCrLf & _
                    "var x = 5;" & vbCrLf & _
                    "var y = 10;" & vbCrLf & _
                    "var sum = x + y;" & vbCrLf & _
                    "console.log(sum);"
    
    txtOutput.Text = "Click 'Parse JavaScript' to see the AST..."
End Sub

Private Sub cmdParse_Click()
    On Error GoTo ErrorHandler
    
    Dim parser As New CParser
    Dim ast As CNode
    Dim json As String
    
    ' Clear output
    txtOutput.Text = ""
    txtOutput.ForeColor = vbBlack
    
    ' Show we're processing
    Me.MousePointer = vbHourglass
    cmdParse.Enabled = False
    DoEvents
    
    ' Parse the JavaScript
    Set ast = parser.ParseScript(txtInput.Text)
    
    ' Generate pretty-printed JSON
    json = ast.ToJSON(0)
    
    ' Show success
    txtOutput.Text = "✓ Parse Successful!" & vbCrLf & vbCrLf & json
    txtOutput.ForeColor = RGB(0, 128, 0)  ' Green
    
    Me.MousePointer = vbDefault
    cmdParse.Enabled = True
    Exit Sub
    
ErrorHandler:
    ' Show error
    txtOutput.Text = "✗ Parse Error:" & vbCrLf & vbCrLf & Err.Description
    txtOutput.ForeColor = vbRed
    
    Me.MousePointer = vbDefault
    cmdParse.Enabled = True
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Ctrl+Enter to parse
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then
        cmdParse_Click
    End If
End Sub

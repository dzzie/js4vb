VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT
'------------------------------------------------------------
' CFunction.cls - Runtime function representation
' "Functions are first-class citizens!"
'------------------------------------------------------------

Option Explicit

Public Params As Collection    ' Collection of parameter names (strings)
Public Body As CNode          ' Function body (BlockStatement)
Public closure As CScope      ' Closure scope (where function was defined)
Public Name As String         ' Function name (can be empty for anonymous)

Private Sub Class_Initialize()
    Set Params = New Collection
End Sub

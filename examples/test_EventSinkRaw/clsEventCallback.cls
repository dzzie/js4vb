VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEventCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Make this the default method via Tools -> Procedure Attributes -> Procedure ID: (Default)
Public Sub OnEvent(ParamArray args() As Variant)
Attribute OnEvent.VB_UserMemId = 0
    Dim i As Long
    
    Form1.d "=== EVENT RECEIVED ==="
    For i = LBound(args) To UBound(args)
         Form1.d "  args(" & i & ") = " & TypeName(args(i))
    Next i
    Form1.d "======================"
    
    
End Sub

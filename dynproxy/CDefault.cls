VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 
Private m_data(0 To 2) As Variant

Private Sub Class_Initialize()
    m_data(0) = "Item0"
    m_data(1) = "Item1"
    m_data(2) = "Item2"
End Sub

' Default member - MUST be marked in the .cls file attributes
Public Function Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    Debug.Print ">>> CTestObject.Item called with index=" & index
    Item = m_data(index)
End Function

Public Property Get Count() As Long
    Count = 3
End Property

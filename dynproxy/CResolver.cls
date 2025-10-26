VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class: CResolver
Option Explicit

Private m_children As Object ' Scripting.Dictionary to hold name -> object

Private Sub Class_Initialize()
    Set m_children = CreateObject("Scripting.Dictionary")
End Sub

Public Function ResolveGetID(ByVal name As String) As Long
    Select Case LCase$(name)
        Case "kitty": ResolveGetID = -30000 - (Abs(CLng(CRC32_Ansi(name))) Mod 10000) ' stable nonzero id
        Case "add":   ResolveGetID = -20001   ' non-zero ? override this name
        Case "hello": ResolveGetID = -20002   ' (optional)
        Case Else:    ResolveGetID = 0        ' 0 means “don’t override”
    End Select
End Function

Public Function ResolveInvoke(ByVal name As String, ByVal flags As Long, args() As Variant, ByVal argC As Long) As Variant

    Select Case True
        Case (flags And 1) <> 0  ' DISPATCH_METHOD
            Debug.Print "Method call:", name
        Case (flags And 2) <> 0  ' DISPATCH_PROPERTYGET
            Debug.Print "Property Get:", name
        Case (flags And 4) <> 0  ' DISPATCH_PROPERTYPUT
            Debug.Print "Property Let:", name
        Case (flags And 8) <> 0  ' DISPATCH_PROPERTYPUTREF
            Debug.Print "Property Set:", name
    End Select
    
    If (flags And &H2) <> 0 Then
        If name = "kitty" Then
            If Not m_children.Exists(name) Then
                ' make a new proxy object that uses this resolver as its behavior engine
                ' CreateProxyForObjectRaw returns a raw pointer; convert to Object via ObjectFromPtr
                Dim p As Long
                p = CreateProxyForObjectRaw(0&, ObjPtr(Me)) ' inner=0, resolver = this instance
                Dim obj As Object
                Set obj = ObjectFromPtr(p)   ' VB now owns the object ref
                ' (do NOT call ReleaseDispatchRaw p - VB owns it)
                m_children.Add name, obj
            End If
            Set ResolveInvoke = m_children(name)
            Exit Function
        End If
    End If
        
    Select Case name
        Case "Hello"
            ResolveInvoke = "Hi from resolver!"
        Case "Add"
            Debug.Print "In ResolveInvoke add(" & args(0) & "," & args(1) & ")"
            ResolveInvoke = args(0) + args(1)
        Case Else
            Err.Raise 438, , "No such member: " & name
    End Select
End Function


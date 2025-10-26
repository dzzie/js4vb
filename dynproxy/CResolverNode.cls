VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResolverNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' CResolverNode.cls
Option Explicit

Private m_name      As String                 ' (optional) for debugging
Private m_children  As Object                 ' Dictionary: name -> Object (proxied child)
Private m_props     As Object                 ' Dictionary: name -> Variant (stored values)

Public Sub Init(ByVal nodeName As String)
    m_name = nodeName
    Set m_children = CreateObject("Scripting.Dictionary")
    Set m_props = CreateObject("Scripting.Dictionary")
End Sub

' Tell the proxy which names we want to handle here.
' Non-zero = we handle it; 0 = let inner handle it.
Public Function ResolveGetID(ByVal name As String) As Long
    Select Case LCase$(name)
        Case "kitty", "meow"
            ResolveGetID = -30000            ' any non-zero stable id
        Case Else
            ResolveGetID = 0
    End Select
End Function

' Core dispatcher
' flags: METHOD=1, GET=2, PUT=4, PUTREF=8
Public Function ResolveInvoke(ByVal name As String, ByVal flags As Long, args() As Variant, ByVal argC As Long) As Variant
    Dim lname As String: lname = LCase$(name)

    ' --- PROPERTY GETS ---
    If (flags And 2) <> 0 Then
        Select Case lname
            Case "kitty"
                ' lazy-create child proxy node on first access
                If Not m_children.Exists("kitty") Then
                    Dim childNode As CResolverNode
                    Set childNode = New CResolverNode
                    childNode.Init m_name & ".kitty"

                    Dim p As Long: p = CreateProxyForObjectRaw(0&, PtrFromObject(childNode))
                    Dim childObj As Object: Set childObj = ObjectFromPtr(p)
                    ' NOTE: VB now owns the proxy ref; do NOT call ReleaseDispatchRaw

                    m_children.Add "kitty", childObj
                End If
                Set ResolveInvoke = m_children("kitty")
                Exit Function

            Case "meow"
                If m_props.Exists("meow") Then
                    ResolveInvoke = m_props("meow")
                Else
                    ResolveInvoke = Empty
                End If
                Exit Function
        End Select
    End If

    ' --- PROPERTY PUT (Let/Set) ---
    If (flags And 4) <> 0 Or (flags And 8) <> 0 Then
        Select Case lname
            Case "meow"
                m_props("meow") = args(0)     ' store value
                Exit Function                 ' Sub-like; no return
        End Select
    End If

    ' --- METHODS (not used in this demo, included for completeness) ---
    If (flags And 1) <> 0 Then
        Select Case lname
            Case "hello": ResolveInvoke = "Hi from resolver!": Exit Function
        End Select
    End If

    Err.Raise 438, , "Unmocked member: " & name & " on " & m_name
End Function


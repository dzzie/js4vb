VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3900
      Left            =   180
      TabIndex        =   2
      Top             =   855
      Width           =   7080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Meow"
      Height          =   420
      Left            =   3375
      TabIndex        =   1
      Top             =   225
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   270
      TabIndex        =   0
      Top             =   135
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hLib As Long

Private Sub Form_Load()
    
    Dim p(), x
    p = Array("\dynproxy.dll", "\release\dynproxy.dll", "\debug\dynproxy.dll")
    
    For Each x In p
        hLib = LoadLibrary(App.Path & x)
        If hLib <> 0 Then
            List1.AddItem "loaded: " & x & " = 0x" & Hex(hLib)
            List1.AddItem "ComTypeName(me) = " & ComTypeName(Me)
            Exit For
        End If
    Next
    
    If hLib = 0 Then
        List1.AddItem "Could not find dynproxy.dll"
        Command1.Enabled = False
        Command2.Enabled = False
    End If
    
    'TestDefaultProp 'works
    
End Sub

Sub TestDefaultProp()
    
    Dim d As New CDefault
    Dim r As New CResolver
    Dim proxied As Object
    Dim p As Long
    
    'MsgBox d(0)  'confirmed working as default access..
      
    p = CreateProxyForObjectRaw(ObjPtr(d), 0&)   ' No resolver, just forward
    Set proxied = ObjectFromPtr(p)
    
    Debug.Print "=== Proxied Test 1 ==="
    Debug.Print proxied(0)
    
    
End Sub


Private Sub Command1_Click()
    Dim inner As Object
    Dim r As New CResolver
    Dim p As Long
    Dim o As Object
     
    List1.Clear
    Set inner = CreateObject("Scripting.Dictionary")
    p = CreateProxyForObjectRaw(ObjPtr(inner), ObjPtr(r))
    'p = CreateProxyForObjectRawEx(ObjPtr(inner), ObjPtr(r), 1)  ' resolver-first
    
    ' wrap into VB object (VB now owns one ref; DO NOT also call ReleaseDispatchRaw)
    Set o = ObjectFromPtr(p)
    
    o.Add "k", "v" 'call default dictionary add sub
     
    ' default (inner-first): Add/Exists go to Dictionary, Hello goes to resolver
    List1.AddItem "o.Exists(k) = " & o.Exists("k")   ' inner
    List1.AddItem "o.Hello = " & o.Hello        ' resolver

    ' Toggle resolver-first and clear cache so "Add" will now route to resolver
    'SetProxyResolverWins p, 1
    'ClearProxyNameCache p
    
    SetProxyOverride p, StrPtr("Add"), -20001
    List1.AddItem "o.Add(3, 3) = " & o.Add(3, 3)          ' now resolver even if resolverWins=0

    ClearProxyNameCache p
    List1.AddItem "o.Add(j, z) = " & o.Add("j", "z") ' now hits default  Add which is a sub no return value
    List1.AddItem "o.Exists(j) = " & o.Exists("j")   ' inner
    
    ' Toggle back to inner-first (optional)
    SetProxyResolverWins p, 0
    ClearProxyNameCache p
    
    Set o = Nothing            ' VB releases the proxy; don't call ReleaseDispatchRaw

End Sub


Private Sub Command2_Click()

    Dim root As New CResolverNode
    Dim p As Long
    Dim o As Object
    
    List1.Clear
    root.Init "root"

    p = CreateProxyForObjectRaw(0&, PtrFromObject(root))   ' inner = NULL, resolver = root
 
    Set o = ObjectFromPtr(p)   ' VB owns the proxy ref now

    ' --- your exact scenario ---
    o.kitty.meow = 12          ' PUT on child "kitty"
    List1.AddItem "o.kitty.meow = " & o.kitty.meow   ' -> 12 (GET)

    ' Optional sanity:
    o.kitty.meow = 99
    List1.AddItem "o.kitty.meow = " & o.kitty.meow   ' -> 99

    Set o = Nothing            ' cleanup: don't call ReleaseDispatchRaw

End Sub



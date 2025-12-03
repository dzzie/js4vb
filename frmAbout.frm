VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About js4vb"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   8040
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   0
      Width           =   7965
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   6960
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblURL 
      Caption         =   "https://sandsprite.com"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Index           =   0
      Left            =   60
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   2
      Top             =   3120
      Width           =   4335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT
 
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Enum shellOpenState
    so_Hidden = 0
    so_Min = 2
    so_Max = 3
    so_Norm = 4
End Enum

Private Sub cmdOK_Click()
      Unload Me
End Sub

Public Function LaunchForm()

     txtDesc = CompileVersionInfo() & vbCrLf & vbCrLf & _
            Replace("js4vb is a Javascript ECMA Script 5 engine written almost\nentirly in VB6." & _
            "It is an easy to use ActiveX dll with\na couple C dlls for lower level details." & _
            vbCrLf & vbCrLf & _
            "js4vb supports debugging, 64 bit numbers and native COM access." & _
            vbCrLf & vbCrLf & _
            "js4vb was created by David Zimmer with the help of claude.ai", "\n", vbCrLf)
        
     Me.Visible = True
     
End Function
 

Private Sub lblURL_Click(index As Integer)
        ShellExec lblURL(index).Caption
End Sub



Public Function ShellExec(path As String, _
                         Optional ByVal action As String = "Open", _
                         Optional ByVal Params As String = vbNullString, _
                         Optional ByVal Directory As String = vbNullString, _
                         Optional ByVal State As shellOpenState = so_Norm) As Boolean
                         
    ShellExec = (ShellExecute(0, action, path, Params, Directory, State) >= 33)
    
End Function

Public Function CompileVersionInfo() As String

    On Error Resume Next
    
    Dim dllVer As String
    Dim dllPath As String
    Dim jsvbPath As String
    Dim ret() As String
    Dim hIndex As Long
    Dim hlNames As String
    Dim i As Long
    
    jsvbPath = GetDllPath("js4vb.dll")
    push ret, pad("js4vb:") & jsvbPath
    push ret, pad("") & App.Major & "." & App.Minor & "." & App.Revision & "  " & FileSize(jsvbPath) & ", " & GetCompileTime(jsvbPath)
    
    
    If Not ensureDynProxy() Then
        push ret, "Dynproxy: not found"
    Else
        dllPath = GetDllPath("dynproxy.dll")
        If FileExists(dllPath) Then
            dllVer = GetFileVersion(dllPath)
            push ret, pad("Dynproxy:") & dllPath
            push ret, pad(" ") & Trim(dllVer & "    " & FileSize(dllPath) & ", " & GetCompileTime(dllPath))
        End If
    End If
    
    If Not ensureUTypes() Then
        push ret, "UTypes: not found"
    Else
        dllPath = GetDllPath("utypes.dll")
        If FileExists(dllPath) Then
            dllVer = GetFileVersion(dllPath)
            push ret, pad("UTypes:") & dllPath
            push ret, pad("") & Trim(dllVer & "    " & FileSize(dllPath) & ", " & GetCompileTime(dllPath))
        End If
    End If
    
    CompileVersionInfo = Join(ret, vbCrLf)
    
End Function

Public Function GetDllPath(dll As String) As String
     Dim h As Long, ret As String
     ret = Space(500)
     h = GetModuleHandle(dll)
     h = GetModuleFileName(h, ret, 500)
     If h > 0 Then ret = Mid(ret, 1, h)
     GetDllPath = ret
End Function

Public Function GetFileVersion(Optional ByVal PathWithFilename As String) As String
    ' return file-properties of given file  (EXE , DLL , OCX)
    'http://support.microsoft.com/default.aspx?scid=kb;en-us;160042
    
    If Len(PathWithFilename) = 0 Then Exit Function
    
    Dim lngBufferlen As Long
    Dim lngDummy As Long
    Dim lngRc As Long
    Dim lngVerPointer As Long
    Dim lngHexNumber As Long
    Dim b() As Byte
    Dim b2() As Byte
    Dim strBuffer As String
    Dim strLangCharset As String
    Dim strTemp As String
    Dim n As Long
    
    ReDim b2(500)
    
    lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
    If lngBufferlen <= 0 Then Exit Function
    
    ReDim b(lngBufferlen)
    lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, b(0))
    If lngRc = 0 Then Exit Function
    
    lngRc = VerQueryValue(b(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
    If lngRc = 0 Then Exit Function
    
    MoveMemory b2(0), lngVerPointer, lngBufferlen
    lngHexNumber = b2(2) + b2(3) * &H100 + b2(0) * &H10000 + b2(1) * &H1000000
    strLangCharset = Right("0000000" & hex(lngHexNumber), 8)
    
    strBuffer = String$(800, 0)
    strTemp = "\StringFileInfo\" & strLangCharset & "\FileVersion"
    lngRc = VerQueryValue(b(0), strTemp, lngVerPointer, lngBufferlen)
    If lngRc = 0 Then Exit Function
    
    lstrcpy strBuffer, lngVerPointer
    n = InStr(strBuffer, Chr(0)) - 1
    If n > 0 Then
        strBuffer = Mid$(strBuffer, 1, n)
        GetFileVersion = strBuffer
    End If
   
End Function

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo Init
    Dim X
       
    X = UBound(ary)
    ReDim Preserve ary(X + 1)
    
    If isObject(value) Then
        Set ary(X + 1) = value
    Else
        ary(X + 1) = value
    End If
    
    Exit Sub
Init:
    ReDim ary(0)
    If isObject(value) Then
        Set ary(0) = value
    Else
        ary(0) = value
    End If
End Sub

Function pad(v, Optional L As Long = 13, Optional char As String = " ", Optional padRight As Boolean = True)
    On Error GoTo hell
    Dim X As Long
    X = Len(v)
    If X < L Then
        If padRight Then
             pad = v & String(L - X, char)
        Else
             pad = String(L - X, char) & v
        End If
    Else
hell:
        pad = v
    End If
End Function

Public Function FileSize(fPath As String) As String
    Dim fsize As Long
    Dim szName As String
    On Error GoTo hell
    
    fsize = FileLen(fPath)
    
    szName = " bytes"
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Kb"
    End If
    
    If fsize > 1024 Then
        fsize = fsize / 1024
        szName = " Mb"
    End If
    
    FileSize = fsize & szName
    
    Exit Function
hell:
    
End Function

Function GetCompileTime(Optional ByVal exe As String, Optional fmt As String = "mm.dd.yy") As String
    
    Dim f As Long, i As Integer
    Dim stamp As Long, e_lfanew As Long
    Dim base As Date, compiled As Date

    On Error GoTo errExit
    
    If Len(exe) = 0 Then
        exe = App.path & "" & App.EXEName & ".exe"
    End If
    
    FileLen exe 'throw error if not exist
    
    f = FreeFile
    Open exe For Binary Access Read As f
    Get f, , i
    
    If i <> &H5A4D Then GoTo errExit 'MZ check
     
    Get f, 60 + 1, e_lfanew
    Get f, e_lfanew + 1, i
    
    If i <> &H4550 Then GoTo errExit 'PE check
    
    Get f, e_lfanew + 9, stamp
    Close f
    
    base = DateSerial(1970, 1, 1)
    compiled = DateAdd("s", stamp, base)
    GetCompileTime = Format(compiled, fmt) '"ddd, mmm d yyyy, h:nn:ss ")
    
    Exit Function
    
errExit:
    On Error Resume Next
    Close f
        
End Function


Function FileExists(path) As Boolean
  On Error GoTo hell
    
  '.(0), ..(0) etc cause dir to read it as cwd!
  If Len(path) = 0 Then Exit Function
  If Right(path, 1) = "\" Then Exit Function
  If InStr(path, Chr(0)) > 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
  
  Exit Function
hell: FileExists = False
End Function


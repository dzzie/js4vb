Attribute VB_Name = "Module1"
Option Explicit

Public Declare Sub IPCDebugMode Lib "dynproxy.dll" (Optional ByVal enabled As Long = 1)
Public Declare Function CreateProxyForProgIDRaw Lib "dynproxy.dll" (ByVal progId As Long, ByVal resolverDisp As Long) As Long
Public Declare Function CreateProxyForObjectRaw Lib "dynproxy.dll" (ByVal innerDispPtr As Long, ByVal resolverDispPtr As Long) As Long
Public Declare Sub ReleaseDispatchRaw Lib "dynproxy.dll" (ByVal pDisp As Long)
Public Declare Sub SetProxyResolverWins Lib "dynproxy.dll" (ByVal proxyPtr As Long, ByVal enable As Long)
Public Declare Sub ClearProxyNameCache Lib "dynproxy.dll" (ByVal proxyPtr As Long)
Public Declare Function CreateProxyForObjectRawEx Lib "dynproxy.dll" (ByVal innerPtr As Long, ByVal resolverPtr As Long, ByVal resolverWins As Long) As Long
Public Declare Sub SetProxyOverride Lib "dynproxy.dll" (ByVal proxyPtr As Long, ByVal nameBSTR As Long, ByVal dispid As Long)
Public Declare Function ComTypeName Lib "dynproxy.dll" (ByVal obj As IUnknown) As Variant
    
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef src As Any, ByVal cb As Long)
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

'' Clears both instance and class caches
'Public Declare Sub ComTypeNameClearCache Lib "dynproxy.dll" ()
'
'' Returns S_OK (0) on success
'Public Declare Function ComTypeNameClearCacheEx Lib "dynproxy.dll" (ByVal flags As Long) As Long
'
'' Flags (mirror the C constants)
'Public Const CTN_CLEAR_INSTANCE As Long = &H1
'Public Const CTN_CLEAR_CLASS    As Long = &H2
'Public Const CTN_CLEAR_ALL      As Long = CTN_CLEAR_INSTANCE Or CTN_CLEAR_CLASS
'
'' Usage examples:
''   Call ComTypeNameClearCache                  ' clear everything
''   Call ComTypeNameClearCacheEx(CTN_CLEAR_ALL) ' same as above
''   Call ComTypeNameClearCacheEx(CTN_CLEAR_CLASS)



' CRC-32 (IEEE 802.3) polynomial 0xEDB88320, init 0xFFFFFFFF, final xor 0xFFFFFFFF
' Works with VB6 (32-bit). Long is signed, so mask when showing hex.
Private Const CRC32_INIT As Long = &HFFFFFFFF
Private Const CRC32_XOR_OUT As Long = &HFFFFFFFF

Private CRC32_Table(0 To 255) As Long
Private m_TableReady As Boolean

Public Function ObjectFromPtr(ByVal p As Long) As Object
    RtlMoveMemory ObjectFromPtr, p, 4&
End Function

Public Function PtrFromObject(ByVal o As Object) As Long
    PtrFromObject = ObjPtr(o)
End Function


Private Sub BuildTableIfNeeded()
    Dim i As Long, j As Long, crc As Long
    If m_TableReady Then Exit Sub
    For i = 0 To 255
        crc = i
        For j = 0 To 7
            If (crc And 1) <> 0 Then
                crc = &HEDB88320 Xor ((crc And &H7FFFFFFF) \ 2) ' arithmetic shift workaround
            Else
                crc = (crc And &H7FFFFFFF) \ 2
            End If
        Next
        CRC32_Table(i) = crc
    Next
    m_TableReady = True
End Sub

' Core update over bytes
Public Function CRC32_Update(ByVal crc_in As Long, ByRef bytes() As Byte, ByVal offset As Long, ByVal Count As Long) As Long
    BuildTableIfNeeded
    Dim i As Long, b As Long, crc As Long
    crc = crc_in
    For i = 0 To Count - 1
        b = bytes(offset + i) And &HFF&
        crc = CRC32_Table((crc Xor b) And &HFF&) Xor ((crc And &HFFFFFFFF) \ &H100)
    Next
    CRC32_Update = crc
End Function

' Convenience: CRC32 over a whole byte array
Public Function CRC32_Bytes(ByRef bytes() As Byte) As Long
    Dim n As Long, crc As Long
    n = UBound(bytes) - LBound(bytes) + 1
    If n <= 0 Then
        CRC32_Bytes = (CRC32_INIT Xor CRC32_XOR_OUT)
        Exit Function
    End If
    crc = CRC32_Update(CRC32_INIT, bytes, LBound(bytes), n)
    CRC32_Bytes = (crc Xor CRC32_XOR_OUT)
End Function

' CRC32 of an ANSI string: uses StrConv to ANSI bytes
Public Function CRC32_Ansi(ByVal s As String) As Long
    Dim b() As Byte
    If LenB(s) = 0 Then
        CRC32_Ansi = (CRC32_INIT Xor CRC32_XOR_OUT)
        Exit Function
    End If
    b = StrConv(s, vbFromUnicode) ' ANSI bytes
    CRC32_Ansi = (CRC32_Update(CRC32_INIT, b, LBound(b), UBound(b) - LBound(b) + 1) Xor CRC32_XOR_OUT)
End Function

' CRC32 of UTF-16LE (VB6 native BSTR bytes). Stable across locales.
Public Function CRC32_Utf16(ByVal s As String) As Long
    Dim nChars As Long, nBytes As Long
    nChars = Len(s)
    If nChars = 0 Then
        CRC32_Utf16 = (CRC32_INIT Xor CRC32_XOR_OUT)
        Exit Function
    End If
    nBytes = nChars * 2
    Dim arr() As Byte
    ReDim arr(0 To nBytes - 1) As Byte
    ' copy raw BSTR bytes into arr
    RtlMoveMemory arr(0), ByVal StrPtr(s), nBytes
    CRC32_Utf16 = (CRC32_Update(CRC32_INIT, arr, 0, nBytes) Xor CRC32_XOR_OUT)
End Function

' Hex helper (unsigned display)
Public Function CRC32_Hex(ByVal crc As Long) As String
    CRC32_Hex = Right$("00000000" & Hex$(crc And &HFFFFFFFF), 8)
End Function



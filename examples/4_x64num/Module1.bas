Attribute VB_Name = "Module1"
Option Explicit

'this demo requires utypes.dll in the same dir as main dll
'you can upload this file along with the debug output to claude for analysis..everything is passing..
Sub main()
    
    'TestDebug
    'TestAutoPromotion
    'TestRollovers
    'TestClean
    'TestSimpleString
    TestEverything
    
    'Debug.Print IsHexString("0xFFFFFFFF + 1 =") 'false
    End
    
    TestHexStr
    TestHexEdgeCases
    QuickHexTest
    
End Sub

Sub TestEverything()
    Dim interp As New CInterpreter
    
    Debug.Print "=== AUTO-BIGINT PROMOTION ==="
    interp.Execute "print('typeof 0x7FFFFFFF:', typeof 0x7FFFFFFF);"
    interp.Execute "print('typeof 0x80000000:', typeof 0x80000000);"
    interp.Execute "print('typeof 0xFFFFFFFF:', typeof 0xFFFFFFFF);"
    
    Debug.Print ""
    Debug.Print "=== BIGINT MATH ==="
    interp.Execute "print('0xFFFFFFFF + 1 =', hex(0xFFFFFFFF + 1));"
    interp.Execute "print('0x80000000 + 0x80000000 =', hex(0x80000000 + 0x80000000));"
    interp.Execute "print('0x140000000 + 0xFFFFFFFF =', hex(0x140000000 + 0xFFFFFFFF));"
    
    Debug.Print ""
    Debug.Print "=== IDA SCRIPTING ==="
    interp.Execute "var imageBase = 0x140000000;"
    interp.Execute "var offset = 0x1000;"
    interp.Execute "var addr = imageBase + offset;"
    interp.Execute "print('Image base:', hex(imageBase));"
    interp.Execute "print('Offset:', hex(offset));"
    interp.Execute "print('Address:', hex(addr));"
    interp.Execute "print('Type of address:', typeof addr);"
End Sub

Sub TestSimpleString()
    Dim interp As New CInterpreter
    'interp.Execute "print('First', 'Second', 'Third');"
    interp.Execute "print('0xFFFFFFFF + 1 =');"
End Sub

Sub TestDebug()
    Dim interp As New CInterpreter

    interp.Execute "console.log(typeof 0x80000000);"
End Sub

Sub TestClean()
    Dim interp As New CInterpreter
    ' interp.debug_mode = False  ' Make sure debug is OFF
    
    Debug.Print "=== CLEAN TEST ==="
    
    interp.Execute "print('typeof 0x7FFFFFFF:', typeof 0x7FFFFFFF);"
    interp.Execute "print('typeof 0x80000000:', typeof 0x80000000);"
    interp.Execute "print('typeof 0xFFFFFFFF:', typeof 0xFFFFFFFF);"
    
    interp.Execute "print('0xFFFFFFFF + 1 =', hex(0xFFFFFFFF + 1));"
    interp.Execute "print('0x80000000 + 0x80000000 =', hex(0x80000000 + 0x80000000));"
    
    interp.Execute "var addr = 0x140000000;"
    interp.Execute "var offset = 0x1000;"
    interp.Execute "print('Address:', hex(addr + offset));"
End Sub

Sub TestAutoPromotion()
    Dim interp As New CInterpreter
    
    ' These should all be bigint
    interp.Execute "console.log('0x80000000 type:', typeof 0x80000000);"
    interp.Execute "console.log('0x90000000 type:', typeof 0x90000000);"
    interp.Execute "console.log('0xFFFFFFFF type:', typeof 0xFFFFFFFF);"
    
    ' This should be number
    interp.Execute "console.log('0x7FFFFFFF type:', typeof 0x7FFFFFFF);"
End Sub

Sub TestRollovers()
    Dim interp As New CInterpreter
    
    Debug.Print "=== 32-BIT ROLLOVER TESTS ==="
    
    ' 32-bit Addition Overflow
    interp.Execute "print('0xFFFFFFFF + 1 =', hex(0xFFFFFFFF + 1));"
    interp.Execute "print('0xFFFFFFFF + 0xFFFFFFFF =', hex(0xFFFFFFFF + 0xFFFFFFFF));"
    
    ' 32-bit Multiplication
    interp.Execute "print('0xFFFFFFFF * 2 =', hex(0xFFFFFFFF * 2));"
    
    Debug.Print ""
    Debug.Print "=== 64-BIT ROLLOVER TESTS ==="
    
    ' 64-bit Max + 1 (should wrap to 0)
    interp.Execute "print('MAX64 + 1 =', hex(0xFFFFFFFFFFFFFFFF + 1));"
    interp.Execute "print('MAX64 + 2 =', hex(0xFFFFFFFFFFFFFFFF + 2));"
    
    ' 64-bit 0 - 1 (should wrap to max)
    interp.Execute "print('0 - 1 =', hex(0x0 - 0x1));"
    
    ' 64-bit Multiplication
    interp.Execute "print('MAX64 * 2 =', hex(0xFFFFFFFFFFFFFFFF * 2));"
    
    Debug.Print ""
    Debug.Print "=== BOUNDARY TESTS ==="
    
    ' Signed boundaries
    interp.Execute "print('0x7FFFFFFF + 1 =', hex(0x7FFFFFFF + 1));"
    interp.Execute "print('0x7FFFFFFFFFFFFFFF + 1 =', hex(0x7FFFFFFFFFFFFFFF + 1));"
    
    Debug.Print ""
    Debug.Print "=== BITWISE OPERATIONS ==="
    
    ' Shifts
    interp.Execute "print('MAX64 << 1 =', hex(0xFFFFFFFFFFFFFFFF << 1));"
    interp.Execute "print('MAX64 >> 1 =', hex(0xFFFFFFFFFFFFFFFF >> 1));"
    
    Debug.Print ""
    Debug.Print "=== IDA ADDRESS TESTS ==="
    
    ' Image base + huge offset
    interp.Execute "print('0x140000000 + 0xFFFFFFFF =', hex(0x140000000 + 0xFFFFFFFF));"
    
    ' Near-max wraparound
    interp.Execute "print('0xFFFFFFFFFFFFFF00 + 0x100 =', hex(0xFFFFFFFFFFFFFF00 + 0x100));"
    
    ' RVA calculation
    interp.Execute "print('0x140001234 - 0x140000000 =', hex(0x140001234 - 0x140000000));"
    
    ' Boundary tests
    interp.Execute "print(typeof 0x7FFFFFFF);"        '// "number" - exactly at signed max
    interp.Execute "print(typeof 0x80000000);"        '// "bigint" - just over signed max
    interp.Execute "print(typeof 0xFFFFFFFF);"        '// "bigint" - unsigned 32-bit max
    
    ' Addition overflow
    interp.Execute "print(0x7FFFFFFF + 1);"           '// 2147483648 (Number, fits in double)"
    interp.Execute "print(0x80000000 + 0x80000000);"  '// BigInt + BigInt = works!"
    interp.Execute "print(0xFFFFFFFF + 1);"           '// BigInt + Number = auto-promotes!"
    
    ' Smaller values stay as Number
    interp.Execute "print(typeof 0xFF);"              '// "number"
    interp.Execute "print(typeof 0xFFFF);"            '// "number"
    interp.Execute "print(typeof 0xFFFFFF);"          '// "number"
    interp.Execute "print(typeof 0x7FFFFFF);"         '// "number" - 7 digits
    
    ' Large values are BigInt
    interp.Execute "print(typeof 0x100000000);"       '// "bigint" - 9 digits
    interp.Execute "print(typeof 0x140000000);"       '// "bigint" - IDA typical
    
    ' Edge cases
    interp.Execute "print(typeof 0x0FFFFFFF);"        '// "number" - 7 digits with leading 0
    interp.Execute "print(typeof 0x00000001);"        '// "number" - leading zeros don't matter

End Sub

Sub TestHexStr()

    Dim interp As New CInterpreter
    
    interp.Execute "print(hex(255), hex(65535))"
    interp.Execute "print(hex(4294967295))" 'FFFFFFFF
    interp.Execute "print(hex(0xFFFFFFFFFFFFFFFFn));"
    
    ' parseInt tests
    interp.Execute "print(parseInt('255'))"
    interp.Execute "print(parseInt('FF', 16))"
    interp.Execute "print(parseInt('0xFFFFFFFF'))"
    interp.Execute "print(parseInt('18446744073709551615'))"  ' Max uint64
    
    ' Math operations with large numbers
    interp.Execute "print(4294967295 + 1)"  ' Should this overflow or become BigInt?
    interp.Execute "print(0xFFFFFFFFn + 1n)"  ' BigInt addition
    interp.Execute "print(0xFFFFFFFFFFFFFFFFn + 1n)"  ' Should wrap or error?
    
    ' Bitwise operations
    interp.Execute "print(0xFFFFFFFFn & 0xFFn)"
    interp.Execute "print(0xFFFFFFFFn | 0xFF00n)"
    interp.Execute "print(0xFFFFFFFFn ^ 0xFFFFFFFFn)"
    interp.Execute "print(0xFFFFFFFFn << 8n)"
    interp.Execute "print(0xFFFFFFFFFFFFFFFFn >> 8n)"
    
    ' Comparisons
    interp.Execute "print(0xFFFFFFFFn > 4294967295)"
    interp.Execute "print(0xFFFFFFFFn == 4294967295)"
    
    ' Type conversions
    interp.Execute "print(Number(0xFFFFFFFFn))"
    interp.Execute "print(String(0xFFFFFFFFFFFFFFFFn))"
    
    ' Mixed operations (these are tricky - JS doesn't allow mixing BigInt and Number in math)
    interp.Execute "print(0xFFFFFFFFn + 1)"  ' Should error in real JS

    '// Basic mixing
    interp.Execute "print(0xFFFFFFFFn + 1);"              '// Should: 4294967296n"
    interp.Execute "print(0x140000000n + 0x1000);"        '// Should: 5368713216n (0x140001000n)"
    
    '// Comparisons
    interp.Execute "print(0xFFFFFFFFn > 4294967295);"     '// Should: false"
    interp.Execute "print(0xFFFFFFFFn == 4294967295);"    '// Should: true"
    
    '// Conversions
    interp.Execute "print(Number(0xFFFFFFFFn));"          '// Should: 4294967295"
    interp.Execute "print(hex(0xFFFFFFFFn));"             '// Should: 0xFFFFFFFF or 0x00000000FFFFFFFF"
    
    '// parseInt
    interp.Execute "print(parseInt('255')); "             '// Should: 255"
    interp.Execute "print(parseInt('0xFFFFFFFF'));"       '// Should: 4294967295"
    interp.Execute "print(parseInt('18446744073709551615'));" '// Should: 18446744073709551615 (as BigInt)"

    '// Test automatic promotion
    interp.Execute "print(typeof 0xFF);"                    '// "number"
    interp.Execute "print(typeof 0xFFFFFFFF);"             '// "number"
    interp.Execute "print(typeof 0x100000000);"            '// "bigint" (auto!)
    interp.Execute "print(typeof 0x140000000);"            '// "bigint" (auto!)
    
    '// Test with explicit n (still works)
    interp.Execute "print(typeof 100n);  "                 '// "bigint"
    interp.Execute "print(typeof 0xFFn); "                 '// "bigint"
    
    '// Test output (no more n!)
    interp.Execute "print(0x140000000);"                   '// "5368709120" (clean!)
    interp.Execute "print(hex(0x140000000));"              '// "0x0000000140000000"
    
    Dim X()
    push X, "// Test IDA-style code"
    push X, "var imageBase = 0x140000000;          // Auto-BigInt, no n needed!"
    push X, "var offset = 0x1000;                  // Regular number"
    push X, "var addr = imageBase + offset;        // Mixed math works!"
    push X, "print(hex(addr));                     // Perfect!"
    interp.Execute Join(X, vbCrLf)
 
End Sub



Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim X
       
    X = UBound(ary)
    ReDim Preserve ary(X + 1)
    
    If IsObject(Value) Then
        Set ary(X + 1) = Value
    Else
        ary(X + 1) = Value
    End If
    
    Exit Sub
init:
    ReDim ary(0)
    If IsObject(Value) Then
        Set ary(0) = Value
    Else
        ary(0) = Value
    End If
End Sub

Sub TestHexEdgeCases()
    Dim interp As New CInterpreter
    
    Debug.Print "=== Comprehensive Hex Test ==="
    Debug.Print ""
    
    ' Test 1: Boundary of 32-bit signed (2^31-1)
    Debug.Print "Test 1: 0x7FFFFFFF (max positive signed 32-bit = 2147483647)"
    interp.Execute "var a = 0x7FFFFFFF; console.log(a, typeof a);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 2: Boundary crossing (2^31)
    Debug.Print "Test 2: 0x80000000 (2^31 = 2147483648, needs unsigned)"
    interp.Execute "var b = 0x80000000; console.log(b, typeof b);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 3: Max 32-bit unsigned
    Debug.Print "Test 3: 0xFFFFFFFF (max 32-bit unsigned = 4294967295)"
    interp.Execute "var c = 0xFFFFFFFF; console.log(c, typeof c);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 4: Just over 32-bit boundary
    Debug.Print "Test 4: 0x100000000 (2^32, needs BigInt)"
    interp.Execute "var d = 0x100000000; console.log(d, typeof d);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 5: Explicit BigInt with small value
    Debug.Print "Test 5: 0x10n (explicit BigInt)"
    interp.Execute "var e = 0x10n; console.log(e, typeof e);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 6: Arithmetic mixing hex and decimal
    Debug.Print "Test 6: 0xFF + 256 (hex + decimal)"
    interp.Execute "var f = 0xFF + 256; console.log(f);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 7: Large hex in calculation
    Debug.Print "Test 7: 0xFFFFFFFF - 1"
    interp.Execute "var g = 0xFFFFFFFF - 1; console.log(g);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 8: Small hex values
    Debug.Print "Test 8: Various small hex"
    interp.Execute "console.log(0x0, 0x1, 0xA, 0xF);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 9: Hex in array
    Debug.Print "Test 9: Hex values in array"
    interp.Execute "var h = [0x10, 0xFF, 0x1000]; console.log(h.length, h[1]);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 10: Hex comparison
    Debug.Print "Test 10: Hex comparison"
    interp.Execute "console.log(0xFF > 100, 0xFF < 300);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 11: typeof on various hex values
    Debug.Print "Test 11: typeof checks"
    interp.Execute "console.log(typeof 0x10, typeof 0xFFFFFFFF, typeof 0x100000000);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    ' Test 12: Large decimal (should auto-promote to BigInt)
    Debug.Print "Test 12: Large decimal auto-promotion"
    interp.Execute "var i = 9007199254740992; console.log(i, typeof i);"
    Debug.Print Trim$(interp.GetOutput())
    interp.ClearOutput
    Debug.Print ""
    
    Debug.Print "=== All Tests Complete ==="
End Sub

'this test would be if you wanted to sanity test the ulong64 class itself
' Sub TestULong64Sanity()
'    Debug.Print "=== ULong64 Sanity Check ==="
'    Debug.Print ""
'
'    Dim u64 As New ULong64
'
'    ' Test 1: Small hex
'    Debug.Print "Test 1: Small hex (0x10)"
'    u64.mode = mUnsigned
'    If u64.fromString("0x10", mHex) Then
'        Debug.Print "  Parsed OK"
'        Debug.Print "  Decimal: " & u64.ToString(mUnsigned)
'        Debug.Print "  Hex: " & u64.ToString(mHex)
'        Debug.Print "  Raw: " & u64.rawValue
'    Else
'        Debug.Print "  PARSE FAILED!"
'    End If
'    Debug.Print ""
'
'    ' Test 2: 0xFF
'    Debug.Print "Test 2: 0xFF"
'    Set u64 = New ULong64
'    u64.mode = mUnsigned
'    If u64.fromString("0xFF", mHex) Then
'        Debug.Print "  Parsed OK"
'        Debug.Print "  Decimal: " & u64.ToString(mUnsigned)
'        Debug.Print "  Hex: " & u64.ToString(mHex)
'        Debug.Print "  Raw: " & u64.rawValue
'    Else
'        Debug.Print "  PARSE FAILED!"
'    End If
'    Debug.Print ""
'
'    ' Test 3: 0xFFFFFFFF (the problem child)
'    Debug.Print "Test 3: 0xFFFFFFFF (should be 4294967295)"
'    Set u64 = New ULong64
'    u64.mode = mUnsigned
'    If u64.fromString("0xFFFFFFFF", mHex) Then
'        Debug.Print "  Parsed OK"
'        Debug.Print "  Decimal: " & u64.ToString(mUnsigned)
'        Debug.Print "  Hex: " & u64.ToString(mHex)
'        Debug.Print "  Raw: " & u64.rawValue
'
'        Dim hi As Long, lo As Long
'        u64.GetLongs hi, lo
'        Debug.Print "  Hi: " & hi & ", Lo: " & lo
'    Else
'        Debug.Print "  PARSE FAILED!"
'    End If
'    Debug.Print ""
'
'    ' Test 4: 0x100000000 (2^32)
'    Debug.Print "Test 4: 0x100000000 (2^32, should need 64-bit)"
'    Set u64 = New ULong64
'    u64.mode = mUnsigned
'    If u64.fromString("0x100000000", mHex) Then
'        Debug.Print "  Parsed OK"
'        Debug.Print "  Decimal: " & u64.ToString(mUnsigned)
'        Debug.Print "  Hex: " & u64.ToString(mHex)
'        Debug.Print "  Raw: " & u64.rawValue
'
'        u64.GetLongs hi, lo
'        Debug.Print "  Hi: " & hi & ", Lo: " & lo
'    Else
'        Debug.Print "  PARSE FAILED!"
'    End If
'    Debug.Print ""
'
'    ' Test 5: Create from hi/lo longs
'    Debug.Print "Test 5: Create from SetLongs(0, -1) - should be 0xFFFFFFFF"
'    Set u64 = New ULong64
'    u64.mode = mUnsigned
'    u64.SetLongs 0, -1  ' In VB6, -1 as Long is 0xFFFFFFFF
'    Debug.Print "  Decimal: " & u64.ToString(mUnsigned)
'    Debug.Print "  Hex: " & u64.ToString(mHex)
'    Debug.Print "  Raw: " & u64.rawValue
'    Debug.Print ""
'
'    ' Test 6: Direct Currency assignment
'    Debug.Print "Test 6: Direct rawValue = 4294967295"
'    Set u64 = New ULong64
'    u64.mode = mUnsigned
'    u64.rawValue = 4294967295@  ' @ forces Currency type
'    Debug.Print "  Decimal: " & u64.ToString(mUnsigned)
'    Debug.Print "  Hex: " & u64.ToString(mHex)
'    u64.GetLongs hi, lo
'    Debug.Print "  Hi: " & hi & ", Lo: " & lo
'    Debug.Print ""
'
'    ' Test 7: Arithmetic test
'    Debug.Print "Test 7: 0xFF + 0x01"
'    Dim u64a As New ULong64
'    Dim u64b As New ULong64
'    u64a.mode = mUnsigned
'    u64b.mode = mUnsigned
'
'    u64a.fromString "0xFF", mHex
'    u64b.fromString "0x01", mHex
'
'    Dim result As ULong64
'    Set result = u64a.Add(u64b)
'
'    Debug.Print "  Result Decimal: " & result.ToString(mUnsigned)
'    Debug.Print "  Result Hex: " & result.ToString(mHex)
'    Debug.Print ""
'
'    ' Test 8: Check signed vs unsigned mode
'    Debug.Print "Test 8: 0xFFFFFFFF in signed mode"
'    Set u64 = New ULong64
'    u64.mode = mSigned
'    If u64.fromString("0xFFFFFFFF", mHex) Then
'        Debug.Print "  Parsed OK"
'        Debug.Print "  Signed: " & u64.ToString(mSigned)
'        Debug.Print "  Unsigned: " & u64.ToString(mUnsigned)
'        Debug.Print "  Hex: " & u64.ToString(mHex)
'    Else
'        Debug.Print "  PARSE FAILED!"
'    End If
'    Debug.Print ""
'
'    Debug.Print "=== End Sanity Check ==="
'End Sub


Sub QuickHexTest()
    Dim interp As New CInterpreter
    
    ' Small hex - stays regular number
    interp.Execute "console.log(0x10);"  ' 16
    Debug.Print interp.GetOutput()
    interp.ClearOutput
    
    ' Large hex - auto BigInt
    interp.Execute "console.log(0xFFFFFFFF);"  ' 4294967295n
    Debug.Print interp.GetOutput()
    interp.ClearOutput
    
    ' Explicit BigInt
    interp.Execute "console.log(0x100n);"  ' 256n
    Debug.Print interp.GetOutput()
    interp.ClearOutput
    
    ' Hex arithmetic
    interp.Execute "var x = 0xFF + 0x01; console.log(x);"  ' 256
    Debug.Print interp.GetOutput()
End Sub

Function IsHexString(s As String) As Boolean
    ' Check if string is a hex number (0x... format)
    If Len(s) < 3 Then
        IsHexString = False
        Exit Function
    End If
    
    If LCase(Left$(s, 2)) <> "0x" Then
        IsHexString = False
        Exit Function
    End If
    
    ' Check if rest are hex digits
    Dim i As Long
    For i = 3 To Len(s)
        Dim c As String
        c = UCase(Mid$(s, i, 1))
        If Not ((c >= "0" And c <= "9") Or (c >= "A" And c <= "F")) Then
            IsHexString = False
            Exit Function
        End If
    Next
    
    IsHexString = True
End Function




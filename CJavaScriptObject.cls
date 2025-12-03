VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJavaScriptObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Author:  David Zimmer <dzzie@yahoo.com>
'AI:      Claude.ai
'Site:    http://sandsprite.com
'License: MIT

'============================================================================
' CJavaScriptObject - Dynamic COM Proxy Resolver for JavaScript Objects
'============================================================================
' This class bridges JavaScript objects to VB6 COM objects using dynproxy.dll
' Allows VB6 to access JS objects like: obj.name, obj.age = 30, etc.
'============================================================================

Option Explicit

Private m_jsObject As CValue      ' The JavaScript object (vtObject)
Private m_interp As CInterpreter  ' Reference to interpreter
Private m_name As String          ' Debug name
Private m_nextDispid As Long              ' ? ADD THIS
Private m_nameToDispid As Object         ' ? ADD THIS (Dictionary)

Private m_varName As String       ' ? ADD THIS for objects!
Private m_isVariable As Boolean   ' ? ADD THIS flag

Private Const dbg_mode As Boolean = False

Private Sub Class_Initialize()
    m_nextDispid = -30000
    Set m_nameToDispid = CreateObject("Scripting.Dictionary")
    'Debug.Print ">>> CJavaScriptObject.Initilize " & hex(ObjPtr(Me))
End Sub

Private Sub Class_Terminate()
    'Debug.Print ">>> CJavaScriptObject.Terminate " & hex(ObjPtr(Me))
End Sub

Public Sub InitFromCValue(jsObj As CValue, interp As CInterpreter, Optional debugName As String = "")

    If jsObj.vType <> vtObject And jsObj.vType <> vtfunction And jsObj.vType <> vtArray Then
        Err.Raise 5, "CJavaScriptObject", "Expected JavaScript object, got " & jsObj.vType
    End If
    
    Set m_jsObject = jsObj
    Set m_interp = interp
    m_name = debugName
    m_isVariable = False  ' ? Direct object reference
    
    ' Initialize if not already done
    If m_nameToDispid Is Nothing Then
        Set m_nameToDispid = CreateObject("Scripting.Dictionary")
        m_nextDispid = -30000
    End If
    
    'Debug.Print ">>> CJavaScriptObject.Init: " & m_name
    

End Sub

Public Sub InitFromVarName(varName As String, interp As CInterpreter)
    m_varName = varName
    Set m_interp = interp
    m_name = varName
    m_isVariable = True  ' ? Variable reference
    
    If m_nameToDispid Is Nothing Then
        Set m_nameToDispid = CreateObject("Scripting.Dictionary")
        m_nextDispid = -30000
    End If
    
    If dbg_mode Then Debug.Print ">>> CJavaScriptObject.InitFromVarName: " & varName
End Sub

Private Function GetCurrentObject() As CValue
    If m_isVariable Then
        ' Get from scope
        Set GetCurrentObject = m_interp.GlobalScope.GetVar(m_varName)
    Else
        ' Direct reference
        Set GetCurrentObject = m_jsObject
    End If
End Function

Public Function ResolveGetID(ByVal Name As String) As Long
    ' Get current object
    Dim currentObj As CValue
    Set currentObj = GetCurrentObject()
    
    ' Empty name = default method (DISPID 0)
    If Len(Name) = 0 Then
        ResolveGetID = 0
        Exit Function
    End If
    
    ' Special methods for function objects
    If Not currentObj Is Nothing And currentObj.vType = vtfunction Then
        Select Case LCase$(Name)
            Case "call", "invoke"
                ResolveGetID = -1
                Exit Function
        End Select
    End If
    
    ' Check cache
    If m_nameToDispid.Exists(Name) Then
        ResolveGetID = m_nameToDispid(Name)
    Else
        ' Assign new unique DISPID
        m_nextDispid = m_nextDispid - 1
        ResolveGetID = m_nextDispid
        m_nameToDispid.add Name, ResolveGetID
    End If
End Function

' Core dispatcher - handles ALL property access and method calls
' flags: METHOD=1, PROPERTYGET=2, PROPERTYPUT=4, PROPERTYPUTREF=8
Public Function ResolveInvoke(ByVal Name As String, ByVal flags As Long, args() As Variant, ByVal argC As Long) As Variant
    On Error GoTo ErrorHandler
    
    If dbg_mode Then Debug.Print ">>> CJavaScriptObject.ResolveInvoke('" & Name & "', flags=" & flags & ", argC=" & argC & ")"
    If dbg_mode Then Debug.Print ">>> flags breakdown: METHOD=" & (flags And 1) & " GET=" & (flags And 2) & " PUT=" & (flags And 4) & " PUTREF=" & (flags And 8)
     
    Dim jsArgs As Collection
    Dim i As Long
    Dim result As CValue
    Dim argVal As CValue
    Dim lname As String
    Dim method As CValue
    Dim jsArgs2 As Collection
    Dim result2 As CValue
    Dim index As Long
    Dim elem As CValue
    Dim currentObj As CValue
    
    'handle default method invocation
    If Name = "<DISPID_0>" Then
        'Default indexed access!
        If argC <> 0 Then
            index = CLng(args(0)) 'I am not supporting multi dimensional arrays with this...
            Set currentObj = GetCurrentObject()
            If (flags And 2) <> 0 Then  ' PROPERTYGET
                If currentObj.vType = vtArray Then
                    If index >= 0 And index < currentObj.arrayVal.count Then ' Return array element
                        Set elem = currentObj.arrayVal(index + 1)
                        If elem.vType = vtObject Then
                            Set ResolveInvoke = CValueToVariant(elem)
                        Else
                            ResolveInvoke = CValueToVariant(elem)
                        End If
                        Exit Function
                    End If
                End If
            ElseIf (flags And 4) <> 0 Then  ' PROPERTYPUT
                ' Array assignment: arr(index) = value
                If UBound(args) >= 1 Then  ' Need index AND value (0 based already known not empty)
                    index = CLng(args(0))  ' First arg is the index
                    Set currentObj = GetCurrentObject()

                    If currentObj.vType = vtArray Then
                        ' Convert new value to CValue
                        Dim newVal As CValue
                        Set newVal = VariantToCValue(args(1))  ' Second arg is the value

                        ' Bounds check
                        If index >= 0 And index < currentObj.arrayVal.count Then
                            ' UPDATE IN-PLACE!
                            Dim existingElem As CValue
                            Set existingElem = currentObj.arrayVal(index + 1)

                            ' Copy the new value INTO the existing CValue
                            existingElem.vType = newVal.vType

                            Select Case newVal.vType
                                Case vtNumber
                                    existingElem.numVal = newVal.numVal
                                Case vtString
                                    existingElem.strVal = newVal.strVal
                                Case vtBoolean
                                    existingElem.boolVal = newVal.boolVal
                                Case vtNull, vtUndefined
                                    ' Just the type change
                                Case vtObject
                                    Set existingElem.objectKeys = newVal.objectKeys
                                    Set existingElem.objectProps = newVal.objectProps
                                Case vtArray
                                    Set existingElem.arrayVal = newVal.arrayVal
                                Case vtfunction, vtCOMObject
                                    Set existingElem.objVal = newVal.objVal
                                Case vtInt64
                                    existingElem.int64Val = newVal.int64Val
                                Case Else
                                    existingElem.vType = vtUndefined
                            End Select

                            If dbg_mode Then Debug.Print ">>> Array element " & index & " updated in-place to type " & newVal.vType
                        ElseIf index = currentObj.arrayVal.count Then
                            ' Allow appending (index == length)
                            currentObj.arrayVal.add newVal
                            If dbg_mode Then Debug.Print ">>> Array element " & index & " appended"
                        Else
                            ' Out of bounds - could auto-expand or error
                            If dbg_mode Then Debug.Print ">>> Array index " & index & " out of bounds (length=" & currentObj.arrayVal.count & ")"
                        End If
                    End If
                End If
                Exit Function
            End If
        End If
    End If
'
    ' -----------------------------------------------------------
    ' METHOD CALL
    ' -----------------------------------------------------------
    If (flags And 1) <> 0 Then
        If dbg_mode Then Debug.Print ">>> METHOD: " & Name
        
        ' Get current object
        Set currentObj = GetCurrentObject()
        
        If currentObj Is Nothing Then
            If dbg_mode Then Debug.Print ">>> ERROR: Current object is Nothing!"
            ResolveInvoke = Empty
            Exit Function
        End If

        lname = LCase$(Name)
        
        ' -------------------------------------------------------
        ' CASE 1: Calling .Call() or .Invoke() on a FUNCTION wrapper
        ' -------------------------------------------------------
        If currentObj.vType = vtfunction And (lname = "call" Or lname = "invoke") Then
            If dbg_mode Then Debug.Print ">>> Calling wrapped function via ." & Name & "()"
            
            ' Convert VB args to JavaScript CValues
            Set jsArgs = New Collection
            
            For i = UBound(args) To LBound(args) Step -1
                Set argVal = VariantToCValue(args(i))
                jsArgs.add argVal
            Next
            
            ' Call function (no 'this' for standalone functions)
            Set result = m_interp.CallJSFunction(currentObj, jsArgs, Nothing)
            
            ' Return result
            If result.vType = vtObject Then
                Set ResolveInvoke = CValueToVariant(result)
            Else
                ResolveInvoke = CValueToVariant(result)
            End If
            
            If dbg_mode Then Debug.Print ">>> Function returned: " & result.vType
            Exit Function
        End If
        
        ' -------------------------------------------------------
        ' CASE 2: Calling a METHOD on an OBJECT (person.greet())
        ' -------------------------------------------------------
        If currentObj.vType = vtObject Then
            If dbg_mode Then Debug.Print ">>> Looking for method '" & Name & "' on object"
            
            ' Get the method from the object
            Set method = currentObj.GetProperty(Name)
            
            If method Is Nothing Then
                ' Property doesn't exist
                ' If PROPERTYGET flag is also set, let it handle this
                If (flags And 2) <> 0 Then
                    If dbg_mode Then Debug.Print ">>> Property not found, falling through to PROPERTYGET"
                    ' Don't handle here, let PROPERTYGET section deal with it
                    ' Fall through (don't Exit Function)
                    GoTo propGetHandler
                Else
                    ' Pure method call, property doesn't exist
                    Err.Raise 438, , "Object doesn't support property or method: " & Name
                End If
            End If
            
            If method.vType = vtfunction Then
                If dbg_mode Then Debug.Print ">>> Found method '" & Name & "' - calling with 'this'"
                
                ' Convert VB args to JavaScript CValues
                Set jsArgs2 = New Collection
                
                For i = UBound(args) To LBound(args) Step -1
                    Set argVal = VariantToCValue(args(i))
                    jsArgs2.add argVal
                Next
                
                ' Call method with currentObj as 'this'!
                Set result2 = m_interp.CallJSFunction(method, jsArgs2, currentObj)
                
                ' Return result
                If result2.vType = vtObject Then
                    Set ResolveInvoke = CValueToVariant(result2)
                Else
                    ResolveInvoke = CValueToVariant(result2)
                End If
                
                If dbg_mode Then Debug.Print ">>> Method returned: " & result2.vType
                Exit Function
            Else
                ' Property exists but is NOT a function
                If (flags And 2) <> 0 Then
                    ' PROPERTYGET flag is set - let PROPERTYGET handle it
                    If dbg_mode Then Debug.Print ">>> '" & Name & "' is not a function, falling through to PROPERTYGET"
                    ' DON'T Exit Function - fall through!
                    GoTo propGetHandler
                Else
                    ' Pure METHOD call (flags=1) on non-function - error!
                    Err.Raise 438, , "'" & Name & "' is not a function"
                End If
            End If
        End If
        
        ' -------------------------------------------------------
        ' CASE 3: Unknown case
        ' -------------------------------------------------------
        Err.Raise 438, , "Cannot call method '" & Name & "' on type " & currentObj.vType
        Exit Function
    End If


propGetHandler: 'YEAH..I DID IT...AND ILL DO IT AGAIN!
    ' -----------------------------------------------------------
    ' PROPERTY GET
    ' -----------------------------------------------------------
    If (flags And 2) <> 0 Then
        If dbg_mode Then Debug.Print ">>> PROPERTYGET: " & Name
        
        ' Get current object
        Set currentObj = GetCurrentObject()
        
        If currentObj Is Nothing Then
            If dbg_mode Then Debug.Print ">>> Current object is Nothing!"
            ResolveInvoke = Empty
            Exit Function
        End If
        
        ' Skip invoke/call on functions
        lname = LCase$(Name)
        
        If currentObj.vType = vtfunction And (lname = "invoke" Or lname = "call") Then
            If dbg_mode Then Debug.Print ">>> Skipping PROPERTYGET for invoke/call"
            ResolveInvoke = Empty
            Exit Function
        End If
        
        ' Get property
        Dim propValue As CValue
        Set propValue = currentObj.GetProperty(Name)
        
        If propValue Is Nothing Then
            If dbg_mode Then Debug.Print ">>> Property '" & Name & "' not found"
            ResolveInvoke = Empty
            Exit Function
        End If
        
        ' Return value
        If propValue.vType = vtObject Or propValue.vType = vtArray Then
            Set ResolveInvoke = CValueToVariant(propValue)
        Else
            ResolveInvoke = CValueToVariant(propValue)
        End If

        If dbg_mode Then Debug.Print ">>> PROPERTYGET '" & Name & "' = " & TypeName(ResolveInvoke) & _
                         " (" & propValue.vType & ")" & _
                         " ObjPtr: " & ObjPtr(propValue)
                         
        Exit Function
    End If
    
    ' -----------------------------------------------------------
    ' PROPERTY PUT - Modify existing CValue in-place!
    ' -----------------------------------------------------------
    If (flags And 4) <> 0 Or (flags And 8) <> 0 Then
        If dbg_mode Then Debug.Print ">>> PROPERTYPUT: " & Name
        
        ' Get current object
        Set currentObj = GetCurrentObject()
        
        If currentObj Is Nothing Then
            If dbg_mode Then Debug.Print ">>> Current object is Nothing!"
            Exit Function
        End If
        
        ' Convert new value
        Dim newValue As CValue
        Set newValue = VariantToCValue(args(0))
        
        ' Get existing property
        Dim existingProp As CValue
        Set existingProp = currentObj.GetProperty(Name)
        
        If existingProp Is Nothing Then
            ' Property doesn't exist - add new
            currentObj.SetProperty Name, newValue
            If dbg_mode Then Debug.Print ">>> PROPERTYPUT '" & Name & "' = " & newValue.vType & " (new property)"
        Else
            ' Property exists - UPDATE IN-PLACE!
            If dbg_mode Then Debug.Print ">>> PROPERTYPUT '" & Name & "' updating in-place, old vType=" & existingProp.vType & " new vType=" & newValue.vType
            
            ' Copy the new value INTO the existing CValue object
            existingProp.vType = newValue.vType
            
            Select Case newValue.vType
                Case vtNumber
                    existingProp.numVal = newValue.numVal
                    
                Case vtString
                    existingProp.strVal = newValue.strVal
                    
                Case vtBoolean
                    existingProp.boolVal = newValue.boolVal
                    
                Case vtNull
                    ' Just set vType, no value to copy
                    
                Case vtUndefined
                    ' Just set vType, no value to copy
                    
                Case vtObject
                    ' Copy object references
                    Set existingProp.objectKeys = newValue.objectKeys
                    Set existingProp.objectProps = newValue.objectProps
                    
                Case vtArray
                    Set existingProp.arrayVal = newValue.arrayVal
                    
                Case vtfunction
                    Set existingProp.objVal = newValue.objVal
                    
                Case vtCOMObject
                    Set existingProp.objVal = newValue.objVal
                    
                Case vtInt64
                    existingProp.int64Val = newValue.int64Val
                    
                Case Else
                    ' Unknown type - just set to undefined
                    existingProp.vType = vtUndefined
            End Select
            
            If dbg_mode Then Debug.Print ">>> PROPERTYPUT '" & Name & "' = " & existingProp.vType & " (updated in-place, ObjPtr=" & ObjPtr(existingProp) & ")"
        End If
        
        Exit Function
    End If
    

    
    ' Unknown operation
    Err.Raise 438, , "Unsupported operation on " & Name & " (flags=" & flags & ")"
    Exit Function
    
ErrorHandler:
    Debug.Print ">>> CJavaScriptObject.ResolveInvoke ERROR: " & Err.description & " ERL_Line: " & Erl
    Err.Raise Err.Number, "CJavaScriptObject.ResolveInvoke", Err.description
End Function

' ----------------------------------------------------------------
' HELPER: Convert JavaScript CValue to VB Variant
' ----------------------------------------------------------------
Private Function CValueToVariant(cv As CValue) As Variant
    If cv Is Nothing Then
        CValueToVariant = Empty
        Exit Function
    End If
    
    Select Case cv.vType
        Case vtNumber
            CValueToVariant = cv.numVal
            
        Case vtString
            CValueToVariant = cv.strVal
            
        Case vtBoolean
            CValueToVariant = cv.boolVal
            
        Case vtNull
            CValueToVariant = Null
            
        Case vtUndefined
            CValueToVariant = Empty
            
        Case vtObject, vtArray
            ' Wrap nested JavaScript object in ANOTHER proxy!
            If dbg_mode Then Debug.Print ">>> CValueToVariant: Wrapping nested " & IIf(cv.vType = vtObject, "object", "array")
            
            Dim nestedObj As New CJavaScriptObject
            nestedObj.InitFromCValue cv, m_interp, m_name & ".<nested>"
            
            ' Create proxy for nested object
            Dim p As Long
            p = CreateProxyForObjectRaw(0&, PtrFromObject(nestedObj))
            If dbg_mode Then Debug.Print ">>> Creating nested obj for ObjPtr " & ObjPtr(cv) & " type: " & cv.vType & " asStr: " & cv.ToString()
            ' Return as VB Object
            Set CValueToVariant = ObjectFromPtr(p)
            
        Case vtfunction
            ' Return a callable wrapper? For now, return a marker
            CValueToVariant = "<JavaScript Function>"
            
        Case vtInt64
            ' Convert to string for now
            CValueToVariant = cv.ToString()
            
        Case Else
            CValueToVariant = Empty
    End Select
End Function

' ----------------------------------------------------------------
' HELPER: Convert VB Variant to JavaScript CValue
' ----------------------------------------------------------------
Private Function VariantToCValue(v As Variant) As CValue
    Dim cv As New CValue
    
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            cv.vType = vtNumber
            cv.numVal = CDbl(v)
            
        Case vbString
            cv.vType = vtString
            cv.strVal = CStr(v)
            
        Case vbBoolean
            cv.vType = vtBoolean
            cv.boolVal = CBool(v)
            
        Case vbNull
            cv.vType = vtNull
            
        Case vbEmpty
            cv.vType = vtUndefined
            
        Case vbObject
            ' VB Object -> could wrap as JavaScript object
            ' For now, store as undefined
            cv.vType = vtUndefined
            
        Case Else
            cv.vType = vtUndefined
    End Select
    
    Set VariantToCValue = cv
End Function



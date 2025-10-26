VERSION 5.00
Begin VB.Form frmJsAsCOM 
   Caption         =   "Form2"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmJsAsCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'this sample requires the C dynproxy.dll std dll compiled and in same directory as the main dll

'here we are wrapping JS objects with a COM wrapper and a dynamic resolver so we can access props and methods directly in vb!
'note the late bound prop names using the bang operator are case sensitive

Private Sub Form_Load()

    TestJsonOverCOM
    TestFunctionCalls
    TestBangVariant
    End
    
End Sub

Private Sub TestFunctionCalls()
    Dim interp As New CInterpreter
    
    Debug.Print "=== TEST: Standalone Function ==="
    interp.Execute "function greet(name) { console.log('Hello, ' + name); return 'Hello, ' + name;}"
    
    Dim greet As Object
    Set greet = interp.EvalAsObject("greet")
    
    greet.Invoke "World"          ' "Hello, World"
    Debug.Print "dbg: " & greet.Call("JavaScript")    ' "Hello, JavaScript"
    
    Debug.Print vbCrLf & "=== TEST: Object Method w/this ==="
    interp.Execute "var person = { " & _
        "name: 'John', " & _
        "greet: function() { return 'Hi, I am ' + this.name; } " & _
    "};"
    
    Dim person As Object
    Set person = interp.EvalAsObject("person")
    
    Debug.Print person.Name              ' "John"
    Debug.Print person.greet()           ' Should be: "Hi, I am John"

'    Debug.Print vbCrLf & "=== TEST: Function with Multiple Args ==="
'    interp.Execute "function add(a, b) { return a + b; }"
'
'    Dim oAdd As Object
'    Set oAdd = interp.EvalAsObject("add")
'
'    Debug.Print oAdd.Add(5, 10)             ' 15
'    Debug.Print oAdd.Add(100, 200)          ' 300

End Sub

Sub TestBangVariant()
    
    Dim interp As New CInterpreter
    Dim data As Object
    Dim items As Variant

    'force latebound case - vb ide sucks here..case sensitive
    Const user = 1, profile = 1, Name = 1, email = 1
    Const books = 1, title = 1, price = 1
    
    interp.Execute "var data = {" & _
                                "items: [1, 2, 3]," & _
                                "user: {" & _
                                "    profile: {" & _
                                "        name: 'John'," & _
                                "        email: 'john@example.com'" & _
                                "    }" & _
                                "}};"

    Set data = interp.EvalAsVariant("data")
    Debug.Print "typename(data) = " & TypeName(data) 'Dictionary

    Debug.Print data!user!profile!Name   '  John
    Debug.Print data!user!profile!email  '  john@example.com

    items = data("items")
    Debug.Print "typename(items) = " & TypeName(items) 'Variant()
    Debug.Print "items(2) = " & items(2)               'items(2) = 3
    Debug.Print data("items")(2)                       '3
    'Debug.Print data.items(2) 'error object required
    
 
'    interp.Execute "var data = { " & _
'        "user: { profile: { name: 'John', email: 'j@test.com' } }, " & _
'        "books: [ {title: 'Book', price: 10}, {title: 'Pen', price: 2} ] " & _
'    "};"
'
'    Set data = interp.EvalAsVariant("data")
    
 
    
End Sub


Sub TestJsonOverCOM()

    Dim interp As New CInterpreter
    'Const NAME = ""
    
    ' Create a JavaScript object
    interp.Execute "var person = { name: 'John', age: 30, city: 'NYC' };"

    ' Get it as a VB COM object!
    Dim person As Object
    Set person = interp.EvalAsObject("person")
    
    'NOTE: if your IDE messes with the late bound case it will fail to look up right now..
    
    ' USE IT LIKE A VB OBJECT!
    Debug.Print "Name: " & person.Name       ' "John"  (trying bad case)
    Debug.Print "Age: " & person.age         ' 30
    Debug.Print "City: " & person.city       ' "NYC"

    ' MODIFY IT!
    person.age = 31
    person.Job = "Developer"                 ' Add new property!

    ' Read BACK via proxy (not JavaScript!)
    Debug.Print "Age (read via proxy): " & person.age
    Debug.Print "Job (read via proxy): " & person.Job
    
    ' CHECK IT IN JAVASCRIPT!
    interp.AddCode "console.log('interp.AddCode -> Age: ' + person.age);"      ' 31
    interp.AddCode "console.log('interp.AddCode -> Job: ' + person.job);"      ' "Developer"
    
    ' NESTED OBJECTS!
    interp.Execute "var company = { name: 'Acme', address: { city: 'Boston', zip: '02101' } };"
    
    Dim Company As Object
    Set Company = interp.EvalAsObject("company")
    
    Debug.Print "Company: " & Company.Name                    ' "Acme"
    Debug.Print "City: " & Company.address.city              ' "Boston"
    Debug.Print "Zip: " & Company.address.Zip                ' "02101"
    
    'MsgBox "FOR THE GLORY! IT WORKS!", vbExclamation
    
    'output:
        'Name: John
        'Age: 30
        'City: NYC
        'Age (read via proxy): 31
        'Job (read via proxy): Developer
        'interp.AddCode -> Age: 31
        'interp.AddCode -> Job: Developer
        'Company: Acme
        'City: Boston
        'Zip: 02101

End Sub

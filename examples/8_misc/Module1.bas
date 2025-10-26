Attribute VB_Name = "Module1"

Option Explicit

Dim m As New CInterpreter
 
Sub main()

    m.AddCOMObject "fso", CreateObject("Scripting.FileSystemObject")
    
    Debug.Print m.eval("print(typeof fso)")
    
End Sub


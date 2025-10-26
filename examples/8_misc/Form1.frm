VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents m As CInterpreter
Attribute m.VB_VarHelpID = -1
 
Sub form_load()
    
    Set m = New CInterpreter
    m.AddCOMObject "fso", CreateObject("Scripting.FileSystemObject")
    
    Debug.Print "Eval: " & m.eval("print(typeof fso)")
    
End Sub


 

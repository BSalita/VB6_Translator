VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public v As New Class1
Public vv As New VBRUN.PropertyBag

Function f() As Class1
Set f = New Class1
End Function
Function ff() As VBRUN.PropertyBag
Set ff = New VBRUN.PropertyBag
End Function


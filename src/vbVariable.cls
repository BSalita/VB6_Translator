VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public MemberType As vbext_MemberType
Public varType As vbDataType
Public varSymbol As String
Public varLength As Long
Public varDimensions As Collection
Public varAddress As Long
Attribute varAddress.VB_VarUserMemId = 0
Public varAttributes As VariableAttributes
Public varVariant As Variant
Public varModule As vbModule ' obsolete?
Public varProc As procTable ' obsolete?
Public varLineNumber As Long ' Make this a Type or Class
Public varComponent As VBComponent

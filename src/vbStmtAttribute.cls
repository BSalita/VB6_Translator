VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtAttribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' is this class obsolete?

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
tokens.Remove 1 ' Remove Attribute
Set token = tokens.Item(1)
tokens.Remove 1 ' Remove Attribute type name
If tokens.Item(1).tokString <> "=" Then Err.Raise 1 ' Unexpected Attribute syntax
tokens.Remove 1 ' Remove =
Select Case UCase(token.tokString)
    Case "VB_CREATEABLE"
        currentModule.clsVB_Creatable = tokens.Item(1).tokValue
    Case "VB_EXPOSED"
        currentModule.clsVB_Exposed = tokens.Item(1).tokValue
    Case "VB_GLOBALNAMESPACE"
        currentModule.clsVB_GlobalNameSpace = tokens.Item(1).tokValue
    Case "VB_EXT_KEY"
        currentModule.clsVB_Ext_KEY = tokens.Item(1).tokValue
    Case "VB_NAME"
        currentModule.clsVB_Name = tokens.Item(1).tokValue
    Case "VB_PREDECLAREDID"
        currentModule.clsVB_PredeclaredID = tokens.Item(1).tokValue
End Select
tokens.Remove 1
End Sub

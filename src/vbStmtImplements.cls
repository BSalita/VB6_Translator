VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtImplements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
tokens.Remove 1 ' remove implements
If PassNumber <> 1 Then
    RemoveAllTokens tokens
    Exit Sub
End If
Dim dt As vbDataType
Set dt = getProjectTLibDataType(tokens)
If dt Is Nothing Then Err.Raise 1
Print #99, "Implements: type="; dt.dtType
' fixme: processing can safely be ignored because VB performs Implements validation
currentModule.clsImplements.Add dt
End Sub

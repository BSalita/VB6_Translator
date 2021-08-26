VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Set token = tokens.Item(1)
tokens.Remove 1
token.tokPCode = vbPCodeCloseFile
If Not IsEOL(tokens) Then
    Do
        getFileNumber tokens, output_stack
        token.tokCount = token.tokCount + 1
        If IsEOL(tokens) Then Exit Do
        If tokens.Item(1).tokString <> "," Then Err.Raise 1
        tokens.Remove 1 ' ,
    Loop
End If
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub

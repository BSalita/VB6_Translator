VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Set token = tokens.Item(1)
token.tokPCode = vbPCodeName
tokens.Remove 1 ' Name
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbString
If UCase(tokens.Item(1).tokString) <> "AS" Then Err.Raise 1 ' expecting As
tokens.Remove 1 ' As
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbString
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub




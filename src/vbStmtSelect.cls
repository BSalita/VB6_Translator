VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtSelect"
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
token.tokString = "SelectCase"
token.tokPCode = vbPCodeSelect
tokens.Remove 1
If UCase(tokens.Item(1).tokString) <> "CASE" Then Err.Raise 1 ' expecting Case keyword
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, output_stack, -1
token.tokDataType = output_stack.Item(output_stack.count).tokDataType And Not VT_BYREF
CoerceOperand gOptimizeFlag, output_stack, output_stack.count, token.tokDataType
Dim s As New vbSelect
s.selectDataType = token.tokDataType

If currentProc.procSelects.count = 0 Then currentProc.procSelects.Add s Else currentProc.procSelects.Add s, , 1
If currentProc.procNests.count = 0 Then currentProc.procNests.Add currentProc.procSelects.Item(1) Else currentProc.procNests.Add currentProc.procSelects.Item(1), , 1

output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub


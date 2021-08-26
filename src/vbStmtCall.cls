VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
tokens.Remove 1 ' remove Call keyword
oRPN.RPNize gOptimizeFlag, tokens, output_stack, VT_VOID
Set token = output_stack.Item(output_stack.count)
' fixme: must check that top of stack is a callable thing
Print #99, "StmtLocalCall dt="; token.tokDataType
output_stack.Add New vbToken
output_stack.Item(output_stack.count).tokString = "Call"
output_stack.Item(output_stack.count).tokType = tokstatement
output_stack.Item(output_stack.count).tokDataType = token.tokDataType
output_stack.Item(output_stack.count).tokPCode = vbPCodeCall
currentProc.procStatements.Add output_stack
For Each token In output_stack
  Print #99, "t="; token.tokString; " dt="; token.tokDataType; " tc="; token.tokCount
Next
End Sub


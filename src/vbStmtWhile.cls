VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtWhile"
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
token.tokPCode = vbPCodeWhile
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbBoolean
token.tokDataType = vbBoolean
If currentProc.procWhiles.count = 0 Then currentProc.procWhiles.Add New vbWhile Else currentProc.procWhiles.Add New vbWhile, , 1
Set currentProc.procWhiles.Item(1).stmtTok = token
If currentProc.procNests.count = 0 Then currentProc.procNests.Add currentProc.procWhiles.Item(1) Else currentProc.procNests.Add currentProc.procWhiles.Item(1), , 1
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub


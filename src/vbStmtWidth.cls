VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtWidth"
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
token.tokPCode = vbPCodeWidth
tokens.Remove 1
getFileNumber tokens, output_stack
If getSpecialTypes(tokens.Item(1)) <> special_comma Then Err.Raise 1 ' expecting ,
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub




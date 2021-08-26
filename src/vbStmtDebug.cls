VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
' Parsing Debug object as if it was a statement, Debug is not a public typelib object
Print #99, "in debug 1"
tokens.Remove 1 ' remove Debug
Print #99, "in debug 2"
If tokens.Item(1).tokType <> tokMember Then Err.Raise 1
tokens.Remove 1 ' remove .
Print #99, "in debug 3 s="; tokens.Item(1).tokString
Set token = tokens.Item(1)
tokens.Remove 1 ' remove Debug member
token.tokType = tokstatement
Select Case UCase(token.tokString)
    Case "ASSERT"
        token.tokPCode = vbPCodeDebugAssert
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbBoolean
    Case "PRINT"
        token.tokPCode = vbPCodeDebugPrint
        parsePrintWriteExpression tokens, output_stack
    Case Else
        Err.Raise 1 ' Invalid Debug member name
End Select
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtDo"
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
If IsEOL(tokens) Then
    token.tokPCode = vbPCodeDo
Else
    If UCase(tokens.Item(1).tokString) = "WHILE" Then
        tokens.Remove 1
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbBoolean
        token.tokString = "DoWhile"
        token.tokPCode = vbPCodeDoWhile
    ElseIf UCase(tokens.Item(1).tokString) = "UNTIL" Then
        tokens.Remove 1
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbBoolean
        token.tokString = "DoUntil"
        token.tokPCode = vbPCodeDoUntil
    Else
        Err.Raise 1 ' expecting While or Until keyword
    End If
    token.tokDataType = output_stack.Item(output_stack.count).tokDataType
End If
If currentProc.procDos.count = 0 Then currentProc.procDos.Add New vbDo Else currentProc.procDos.Add New vbDo, , 1
Set currentProc.procDos.Item(1).stmtTok = token
If currentProc.procNests.count = 0 Then currentProc.procNests.Add currentProc.procDos.Item(1) Else currentProc.procNests.Add currentProc.procDos.Item(1), , 1
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub



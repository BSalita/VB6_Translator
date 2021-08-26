VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtLoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection

If checkNesting(currentProc.procDos) Then Err.Raise 1 ' Missing Do statement
Set token = tokens.Item(1)
tokens.Remove 1
Set token.tokDo = currentProc.procDos.Item(1)
currentProc.procDos.Remove 1
currentProc.procNests.Remove 1

If IsEOL(tokens) Then
    If token.tokDo.stmtTok.tokPCode = vbPCodeDo Then
        token.tokPCode = vbPCodeLoopInfinite
    Else
        token.tokPCode = vbPCodeLoop
    End If
Else
    If token.tokDo.stmtTok.tokPCode <> vbPCodeDo Then Err.Raise 1 ' Can't have While/Until on both DO and Loop
    If UCase(tokens.Item(1).tokString) = "WHILE" Then
        tokens.Remove 1
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbBoolean
        token.tokString = "LoopWhile"
        token.tokPCode = vbPCodeLoopWhile
    ElseIf UCase(tokens.Item(1).tokString) = "UNTIL" Then
        tokens.Remove 1
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbBoolean
        token.tokString = "LoopUntil"
        token.tokPCode = vbPCodeLoopUntil
    Else
    If Not IsEOL(tokens) Then Err.Raise 1 ' expecting While or Until keyword
    End If
End If
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub


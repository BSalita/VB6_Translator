VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbCompilerDirectiveElseIf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim output_stack As New Collection
tokens.Remove 1 ' ElseIf
oRPN.ConstantRPNize ProcessCompilerExpression, tokens, output_stack, vbBoolean
If getKeyword(tokens.Item(1)) <> KW_THEN Then Err.Raise 1 ' expecting Then keyword
tokens.Remove 1 ' Then
If CompilerIfStack.Item(CompilerIfStack.count).CompilerIfHaveProcessedElse Then Err.Raise 1 ' #ElseIf may not appear after #Else
If CompilerIfStack.Item(CompilerIfStack.count).CompilerIfHaveProcessedTrue Then
    CompilerIfStack.Item(CompilerIfStack.count).CompilerIfExpressionValue = False
Else
    CompilerIfStack.Item(CompilerIfStack.count).CompilerIfHaveProcessedTrue = CBool(output_stack.Item(1).tokValue)
    CompilerIfStack.Item(CompilerIfStack.count).CompilerIfExpressionValue = CBool(output_stack.Item(1).tokValue)
End If
If CompilerIfStack.count > 1 Then CompilerIfStack.Item(CompilerIfStack.count).CompilerIfExpressionValue = CompilerIfStack.Item(CompilerIfStack.count).CompilerIfExpressionValue And CompilerIfStack.Item(CompilerIfStack.count - 1).CompilerIfExpressionValue
output_stack.Remove 1
End Sub


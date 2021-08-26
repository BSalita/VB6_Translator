VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbCompilerDirectiveElse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
tokens.Remove 1 ' Else
If CompilerIfStack.Item(CompilerIfStack.count).CompilerIfHaveProcessedElse Then Err.Raise 1 ' #Else already processed
CompilerIfStack.Item(CompilerIfStack.count).CompilerIfHaveProcessedElse = True
CompilerIfStack.Item(CompilerIfStack.count).CompilerIfExpressionValue = Not CompilerIfStack.Item(CompilerIfStack.count).CompilerIfHaveProcessedTrue
If CompilerIfStack.count > 1 Then CompilerIfStack.Item(CompilerIfStack.count).CompilerIfExpressionValue = CompilerIfStack.Item(CompilerIfStack.count).CompilerIfExpressionValue And CompilerIfStack.Item(CompilerIfStack.count - 1).CompilerIfExpressionValue
End Sub


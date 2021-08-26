VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtWend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim output_stack As New Collection

If checkNesting(currentProc.procWhiles) Then Err.Raise 1 ' Missing While statement
tokens.Item(1).tokDataType = currentProc.procWhiles.Item(1).stmtTok.tokDataType
currentProc.procWhiles.Remove 1
currentProc.procNests.Remove 1

tokens.Item(1).tokPCode = vbPCodeWend
output_stack.Add tokens.Item(1)
currentProc.procStatements.Add output_stack
tokens.Remove 1
End Sub


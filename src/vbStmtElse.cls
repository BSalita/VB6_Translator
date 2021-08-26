VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtElse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Set token = tokens.Item(1)
tokens.Remove 1
If Not checkNesting(currentProc.procBlockIfs) Then
    token.tokPCode = vbPCodeElse
    If Not currentProc.procBlockIfs.Item(1).ifElse Is Nothing Then Err.Raise 1 ' Else Else - wrong for block Ifs
    Set currentProc.procBlockIfs.Item(1).ifElse = token
ElseIf Not checkNesting(currentProc.procSingleIfs) Then
    ' assign label to single else
    If currentProc.procSingleIfs.Item(1).ifElse Is Nothing Then ' Else Else - OK for single line Ifs
        token.tokPCode = vbPCodeSingleIfElse
    Else
        token.tokString = "EndIfElse"
        token.tokPCode = vbPCodeSingleIfEndIfElse
        currentProc.procNests.Remove 1
        currentProc.procSingleIfs.Remove 1
    End If
    Set currentProc.procSingleIfs.Item(1).ifElse = token
Else
    Err.Raise 1 ' Else without If
End If
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub


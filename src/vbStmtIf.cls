VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtIf"
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
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbBoolean
If getKeyword(tokens.Item(1)) <> KW_THEN Then Err.Raise 1 ' expecting Then keyword
tokens.Remove 1

If tokens.count = 0 Then ' can't use IsEOL(tokens) because Else and : generate Single If
    token.tokPCode = vbPCodeIf
    If currentProc.procBlockIfs.count = 0 Then currentProc.procBlockIfs.Add New vbIf Else currentProc.procBlockIfs.Add New vbIf, , 1
    Set currentProc.procBlockIfs.Item(1).stmtTok = token
    If currentProc.procNests.count = 0 Then currentProc.procNests.Add currentProc.procBlockIfs.Item(1) Else currentProc.procNests.Add currentProc.procBlockIfs.Item(1), , 1
Else
    token.tokPCode = vbPCodeSingleIf
    If currentProc.procSingleIfs.count = 0 Then currentProc.procSingleIfs.Add New vbIf Else currentProc.procSingleIfs.Add New vbIf, , 1
    token.tokString = "SingleIf"
    Set currentProc.procSingleIfs.Item(1).stmtTok = token
    If currentProc.procNests.count = 0 Then currentProc.procNests.Add currentProc.procSingleIfs.Item(1) Else currentProc.procNests.Add currentProc.procSingleIfs.Item(1), , 1
End If

output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub


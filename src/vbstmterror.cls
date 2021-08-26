VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Set token = tokens.Item(1)
token.tokPCode = vbPCodeError
tokens.Remove 1 ' Error
If tokens.Item(1).tokString = ":" Then
    ' Possible bug or misfeature in VB allowing Error to be used as a label.
    Print #99, """Error"" is a statement and may not be used as a label. Please rename label."
    MsgBox """Error"" is a statement and may not be used as a label. Please rename label."
    Err.Raise 1
ElseIf tokens.Item(1).tokString = "." Then ' Error.Raise is a frequent error
    ' Possible bug or misfeature in VB allowing Error as an object reference.
    Print #99, """Error"" is a statement and may not be used as in as an object reference."
    MsgBox """Error"" is a statement and may not be used as in an object reference."
    Err.Raise 1
End If
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbLong ' Doc says VT_UI2
token.tokDataType = vbLong
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub




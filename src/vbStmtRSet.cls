VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtRSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)

Dim oLet As New vbStmtLet
oLet.vbStmt_stmtParse tokens

#If 0 Then

Dim token As vbToken
Dim output_stack As New Collection
Dim LHS_output_stack As New Collection
Dim stmtRSet As vbToken

Set stmtRSet = tokens.Item(1)
stmtRSet.tokPCode = vbPCodeRSet
tokens.Remove 1

Set token = SymbolLookUp(gOptimizeFlag, tokens, LHS_output_stack, INVOKE_PROPERTYPUT)
If token.tokDataType <> vbString Then Err.Raise 1 ' Expecting String variable

' use function to get =?
If tokens.Item(1).tokString <> "=" Then Err.Raise 1
tokens.Remove 1

oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbString

AppendCollection LHS_output_stack, output_stack
output_stack.Add token

stmtRSet.tokDataType = token.tokDataType
output_stack.Add stmtRSet ' add RSet

currentProc.procStatements.Add output_stack
#End If
End Sub



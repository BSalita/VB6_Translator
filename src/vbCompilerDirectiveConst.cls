VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbCompilerDirectiveConst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection

tokens.Remove 1 ' Const

Set token = tokens.Item(1)
tokens.Remove 1 ' =
If token.tokType <> toksymbol Then Err.Raise 1 ' expecting constname

' use function to get =?
If tokens.Item(1).tokString <> "=" Then Err.Raise 1
tokens.Remove 1

oRPN.ConstantRPNize ProcessCompilerExpression, tokens, output_stack, -1
Print #99, "#Const: 1"
token.tokType = output_stack.Item(1).tokType
Print #99, "#Const: 2"
token.tokValue = output_stack.Item(1).tokValue
Print #99, "#Const: 3"
token.tokDataType = varType(output_stack.Item(1).tokValue)
Print #99, "#Const: 4"
output_stack.Remove 1
Print #99, "#Const: 5 s="; token.tokString; " pc.c="; currentModule.cPredefinedConstant.count; " pn="; PassNumber
If PassNumber = 1 Then currentModule.cPredefinedConstant.Add token, UCase(token.tokString)
Print #99, "#Const: 6"

End Sub



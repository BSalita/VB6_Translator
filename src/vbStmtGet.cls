VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

' Get/Put code is almost same
Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim stmtToken As vbToken
Dim output_stack As New Collection
Set stmtToken = tokens.Item(1)
stmtToken.tokPCode = vbPCodeGet ' Only this line is differs from Put
tokens.Remove 1
getFileNumber tokens, output_stack
If IsEOL(tokens) Then Err.Raise 1
Dim special As SpecialTypes
special = getSpecialTypes(tokens.Item(1))
If special <> special_comma Then Err.Raise 1
tokens.Remove 1
special = getSpecialTypes(tokens.Item(1))
If special = special_comma Then
    output_stack.Add New vbToken
                    ' should be assigning to variant
    output_stack.Item(output_stack.count).tokString = "0"
    output_stack.Item(output_stack.count).tokType = tokVariant
    output_stack.Item(output_stack.count).tokValue = 0&
    output_stack.Item(output_stack.count).tokDataType = vbLong
Else
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbLong
    special = getSpecialTypes(tokens.Item(1))
    If special <> special_comma Then Err.Raise 1
End If
tokens.Remove 1
' INVOKE_PROPERTYPUT for fixed string - fstring1.vbp
' don't want to generate _DEFAULT for Variants - using NoInsertObjDefault=True
Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_PROPERTYGET, , , True) ' INVOKE_FUNC or INVOKE_PROPERTYGET
output_stack.Add token
'token.tokDataType = token.tokDataType Or VT_BYREF
stmtToken.tokDataType = token.tokDataType
output_stack.Add stmtToken
currentProc.procStatements.Add output_stack
End Sub



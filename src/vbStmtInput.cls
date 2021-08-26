VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim stmtToken As vbToken
Dim output_stack As New Collection
Set stmtToken = tokens.Item(1)
stmtToken.tokPCode = vbPCodeInput
tokens.Remove 1
getFileNumber tokens, output_stack
Dim special As SpecialTypes
Do
    special = getSpecialTypes(tokens.Item(1))
    If special <> special_comma Then Err.Raise 1
    tokens.Remove 1
    ' INVOKE_PROPERTYPUT for fixed string - fstring1.vbp
    ' don't want to generate _DEFAULT for Variants - using NoInsertObjDefault=True
    Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_PROPERTYGET, , , True) ' INVOKE_FUNC Or INVOKE_PROPERTYGET)
    output_stack.Add token
Loop Until IsEOL(tokens)
output_stack.Add stmtToken
currentProc.procStatements.Add output_stack
End Sub



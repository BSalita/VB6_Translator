VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

' Lock/Unlock code is similar
Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Dim special As SpecialTypes
Set token = tokens.Item(1)
token.tokPCode = vbPCodeLock
tokens.Remove 1
getFileNumber tokens, output_stack
If IsEOL(tokens) Then ' no recordrange
    output_stack.Add New vbToken
    output_stack.Item(output_stack.count).tokString = "0"
    output_stack.Item(output_stack.count).tokType = tokVariant
    output_stack.Item(output_stack.count).tokValue = 0&
    output_stack.Item(output_stack.count).tokDataType = vbLong
    output_stack.Add New vbToken
    output_stack.Item(output_stack.count).tokString = "0"
    output_stack.Item(output_stack.count).tokType = tokVariant
    output_stack.Item(output_stack.count).tokValue = 0&
    output_stack.Item(output_stack.count).tokDataType = vbLong
Else
    special = getSpecialTypes(tokens.Item(1))
    If special <> special_comma Then Err.Raise 1 ' expecting ,
    tokens.Remove 1
    If IsEOL(tokens) Then Err.Raise 1 ' unexpected eos
    If getKeyword(tokens.Item(1)) = KW_TO Then
        output_stack.Add New vbToken
        output_stack.Item(output_stack.count).tokString = "0"
        output_stack.Item(output_stack.count).tokType = tokVariant
        output_stack.Item(output_stack.count).tokValue = 0&
        output_stack.Item(output_stack.count).tokDataType = vbLong
        If getKeyword(tokens.Item(1)) <> KW_TO Then Err.Raise 1 ' expecting TO keyword
        tokens.Remove 1
        If IsEOL(tokens) Then Err.Raise 1 ' expecting "End" numeric expression
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbLong
    Else
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbLong
        If IsEOL(tokens) Then
            output_stack.Add New vbToken
            output_stack.Item(output_stack.count).tokString = "0"
            output_stack.Item(output_stack.count).tokType = tokVariant
            output_stack.Item(output_stack.count).tokValue = 0&
            output_stack.Item(output_stack.count).tokDataType = vbLong
        Else
            If getKeyword(tokens.Item(1)) <> KW_TO Then Err.Raise 1 ' expecting TO keyword
            tokens.Remove 1
            If IsEOL(tokens) Then Err.Raise 1 ' expecting "End" numeric expression
            oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbLong
        End If
    End If
End If
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub


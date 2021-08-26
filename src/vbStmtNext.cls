VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtNext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim tokStmt As vbToken
Dim output_stack As New Collection

Set tokStmt = tokens.Item(1)
tokens.Remove 1

If IsEOL(tokens) Then
    If checkNesting(currentProc.procFors) Then Err.Raise 1 ' Missing For statement
    ' Not sure if NextVar needs to be created. Use different pcode instead?
    Set token = New vbToken
    token.tokType = tokVariable
    Set token.tokVariable = currentProc.procFors.Item(1).forVariable
    token.tokDataType = token.tokVariable.varType.dtDataType
    token.tokString = token.tokVariable.varSymbol
    If currentProc.procFors.Item(1).stmtTok.tokPCode = vbPCodefor Then
        tokStmt.tokPCode = vbPCodeForNext
    ElseIf currentProc.procFors.Item(1).stmtTok.tokPCode = vbPCodeforeach Then
        tokStmt.tokPCode = vbPCodeForEachNext
    Else
        Err.Raise 1 ' internal error
    End If
    output_stack.Add token
    currentProc.procFors.Remove 1
    currentProc.procNests.Remove 1
    tokStmt.tokCount = 1
Else
    Do
        If checkNesting(currentProc.procFors) Then Err.Raise 1 ' Missing For statement
        If currentProc.procFors.Item(1).stmtTok.tokPCode = vbPCodefor Then
' object ref not allowed as For/Next variable, Type member is allowed
            Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET)
            tokStmt.tokPCode = vbPCodeForNextV
        ElseIf currentProc.procFors.Item(1).stmtTok.tokPCode = vbPCodeforeach Then
' object ref required as For Each/Next variable, Type member is allowed
            Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET, True, , True)
            tokStmt.tokPCode = vbPCodeForEachNextV
        Else
            Err.Raise 1 ' internal error
        End If
        If Not token.tokVariable Is currentProc.procFors.Item(1).forVariable Then Err.Raise 1 ' Next variable doesn't match For
        output_stack.Add token
        currentProc.procFors.Remove 1
        currentProc.procNests.Remove 1
        tokStmt.tokCount = tokStmt.tokCount + 1
    Loop While OptionalComma(tokens)
End If

tokStmt.tokDataType = token.tokDataType
output_stack.Add tokStmt
currentProc.procStatements.Add output_stack
End Sub


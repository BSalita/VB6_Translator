VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim stmtTok As vbToken
Dim token As vbToken
Dim output_stack As New Collection

If currentProc.procSelects.count = 0 Then Err.Raise 1 ' Case without Select
'If currentProc.procSelects.Item(1).stmtTok.tokStmtCount Then Err.Raise 1 ' Select not followed by Case

Set stmtTok = tokens.Item(1)
Set stmtTok.tokSelect = currentProc.procSelects.Item(1)
stmtTok.tokDataType = stmtTok.tokSelect.selectDataType
If UCase(tokens.Item(2).tokString) = "ELSE" Then
    tokens.Remove 1
    tokens.Remove 1
    If Not stmtTok.tokSelect.selectElseLabel Is Nothing Then Err.Raise 1 ' Case Else already defined
    stmtTok.tokPCode = vbPCodeCaseElse
    stmtTok.tokString = "CaseElse"
' assign value to case else label
Else
    stmtTok.tokPCode = vbPCodeCase
    Do
        tokens.Remove 1 ' remove CASE or ,
        Set token = New vbToken
        token.tokDataType = stmtTok.tokDataType
        If UCase(tokens.Item(1).tokString) = "IS" Then
            tokens.Remove 1 ' Remove Is
            If tokens.Item(1).tokType <> tokOperator Then Err.Raise 1 ' Expecting comparision operator
            On Error Resume Next
            Set token.tokOperator = cBinaryOperators.Item(UCase(tokens.Item(1).tokString))
            On Error GoTo 0
            If token.tokOperator Is Nothing Then Err.Raise 1 ' Expecting comparison operator
            token.tokPCode = token.tokOperator.oprPCode
            tokens.Remove 1 ' Remove operator
            oRPN.RPNize gOptimizeFlag, tokens, output_stack, stmtTok.tokDataType
            token.tokType = tokCaseIs
            token.tokString = "CaseIs"
        Else
            oRPN.RPNize gOptimizeFlag, tokens, output_stack, stmtTok.tokDataType
            token.tokType = tokCase
            token.tokString = "Case"
            If Not IsEOL(tokens) Then
                If UCase(tokens.Item(1).tokString) = "TO" Then
                    tokens.Remove 1
                    oRPN.RPNize gOptimizeFlag, tokens, output_stack, stmtTok.tokDataType
                    token.tokType = tokCaseTo
                Else
        '                If output_stack.Item(output_stack.count).tokDataType <> vbBoolean Then
                End If
        ' assign value to case label
            End If
        End If
    '    token.tokDataType = output_stack.Item(output_stack.count).tokDataType
        output_stack.Add token
    '    If i > 1 Then
    '        Set token = New vbToken
    '        token.tokType = tokOperator
    '        Set token.tokOperator = cBinaryOperators.Item("OR")
    '        token.tokLHS = output_stack.count - 1
    '        token.tokRHS = output_stack.count
    '        token.tokOperator.oprOutput True, output_stack, token
    '        output_stack.Add token
    '    End If
        If IsEOL(tokens) Then Exit Do
    Loop While tokens.Item(1).tokString = ","
End If
output_stack.Add stmtTok
currentProc.procStatements.Add output_stack
End Sub


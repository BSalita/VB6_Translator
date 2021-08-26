VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Dim special As SpecialTypes
Set token = tokens.Item(1)
tokens.Remove 1
If UCase(tokens.Item(1).tokString) = "LOCAL" Then
' fixme: can Local be safely ignored?
    tokens.Remove 1 ' remove Local
    If UCase(tokens.Item(1).tokString) <> "ERROR" Then Err.Raise 1
End If
If UCase(tokens.Item(1).tokString) = "ERROR" Then
    tokens.Remove 1
    If UCase(tokens.Item(1).tokString) = "RESUME" Then ' On Error Resume Next
        token.tokString = "OnErrorResumeNext"
        token.tokPCode = vbpcodeonerrorresumenext ' -1
        tokens.Remove 1
        If UCase(tokens.Item(1).tokString) <> "NEXT" Then Err.Raise 1 ' Expecting Next keyword
    Else
        If UCase(tokens.Item(1).tokString) <> "GOTO" Then Err.Raise 1 ' Expecting GoTo or Resume keyword
        tokens.Remove 1
        token.tokString = "OnErrorGoTo"
        token.tokPCode = vbPCodeOnErrorLabel
        If tokens.Item(1).tokType = tokVariant Then
            If CInt(tokens.Item(1).tokString) = 0 Then ' Resume 0
                token.tokString = "OnErrorGoTo0"
                token.tokPCode = vbPCodeOnError0 ' -2
            Else ' Resume line number
                AddLabelRef tokens.Item(1), output_stack
            End If
        Else ' Resume label
            AddLabelRef tokens.Item(1), output_stack
        End If
    End If
    tokens.Remove 1
Else
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
    token.tokDataType = output_stack.Item(output_stack.count).tokDataType
    If UCase(tokens.Item(1).tokString) = "GOSUB" Then
        token.tokString = "OnGoSub"
        token.tokPCode = vbPCodeOnGoSub
    ElseIf UCase(tokens.Item(1).tokString) = "GOTO" Then
        token.tokString = "OnGoTo"
        token.tokPCode = vbPCodeOnGoTo
    Else
        Err.Raise 1
    End If
    Do
        tokens.Remove 1
        If IsEOL(tokens) Then Err.Raise 1 ' Expecting label or line number
        AddLabelRef tokens.Item(1), output_stack
        tokens.Remove 1
        If IsEOL(tokens) Then Exit Do
        special = getSpecialTypes(tokens.Item(1))
    Loop While special = special_comma
End If
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub


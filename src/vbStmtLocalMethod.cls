VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtLocalMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

' Note needs Public
Public Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Print #99, "StmtLocalMethod: s="; tokens.Item(1).tokString
Select Case UCase(LastQualifier.tokString)
    Case "PRINT" ' cannot use Write keyword as method name
        Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET, True, False, True)
        token.tokType = tokstatement
        token.tokPCode = vbPCodePrintMethod
        parsePrintWriteExpression tokens, output_stack
        output_stack.Add token
    Case Else
        ' kludgy: wraping statement method with parenthesis to make same syntax as Call
        Dim i As Long
        For i = 1 To tokens.count
            If tokens.Item(i) Is LastQualifier Then Exit For
        Next
        If i > tokens.count Then Err.Raise 1
        tokens.Add New vbToken, , , i
        tokens.Item(i + 1).tokType = tokSpecial
        tokens.Item(i + 1).tokString = "("
        Dim j As Long
        For j = i + 2 To tokens.count
            If IsEOL(tokens, j) Then Exit For
        Next
        tokens.Add New vbToken, , , j - 1
        tokens.Item(j).tokType = tokSpecial
        tokens.Item(j).tokString = ")"
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, VT_VOID, , True ' VT_VOID specifies no return value
        Set token = output_stack.Item(output_stack.count)
        If token.tokPCode = 0 Then
            output_stack.Add New vbToken
            output_stack.Item(output_stack.count).tokType = tokstatement
            output_stack.Item(output_stack.count).tokString = "Call"
            output_stack.Item(output_stack.count).tokDataType = token.tokDataType
            output_stack.Item(output_stack.count).tokPCode = vbPCodeCall
        Else ' Circle, Line, PSet, Scale
            token.tokType = tokstatement
        End If
End Select
Print #99, "StmtLocalMethod: dt="; token.tokDataType
currentProc.procStatements.Add output_stack
For Each token In output_stack
    Print #99, "t="; token.tokString; " dt="; token.tokDataType; " tc="; token.tokCount
Next
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
'Dim token As vbToken
Print #99, "End: t.c="; tokens.count
Dim output_stack As New Collection

'Set token = tokens.Item(1)
'token.tokPCode = vbPCodeEnd

tokens.Remove 1

If IsEOL(tokens) Then
' fill in End stuff!!!!!!!!!!!
Else
' fill in goto stuff!!!!!!!!!!!!!!!!!!!!!!
    Print #99, "End: s="; tokens.Item(1).tokString; " kw="; tokens.Item(1).tokKeyword; " cp="; Not currentProc Is Nothing
    Select Case tokens.Item(1).tokKeyword
        Case KW_FUNCTION
            If tokens.count <> 1 Then Err.Raise 1
            If currentProc.MemberType <> vbext_mt_Method Then Err.Raise 1
            If currentProc.procFunctionResultType Is Nothing Then Err.Raise 1
            EndProc
        Case KW_IF
            If checkNesting(currentProc.procBlockIfs) Then Err.Raise 1 ' Missing If statement
            Set tokens.Item(1).tokIf = currentProc.procBlockIfs.Item(1)
' if stmt count = currentProc.procBlockIfs.stmtTok.tokStmtCount then error ' empty If statement
            currentProc.procBlockIfs.Remove 1
            currentProc.procNests.Remove 1
            tokens.Item(1).tokType = tokstatement
            tokens.Item(1).tokPCode = vbPCodeEndIf
            tokens.Item(1).tokString = "EndIf"
            output_stack.Add tokens.Item(1)
            currentProc.procStatements.Add output_stack
        Case KW_PROPERTY
            Print #99, "mt="; currentProc.MemberType; " ik="; currentProc.InvokeKind; " pf="; Not currentProc.procFunctionResultType Is Nothing
            If tokens.count <> 1 Then Err.Raise 1
            If currentProc.MemberType <> vbext_mt_Property Then Err.Raise 1
            If currentProc.InvokeKind = INVOKE_PROPERTYGET Xor Not currentProc.procFunctionResultType Is Nothing Then Err.Raise 1
            EndProc
        Case KW_SELECT
            If checkNesting(currentProc.procSelects) Then Err.Raise 1 ' Missing Select statement
            Set tokens.Item(1).tokSelect = currentProc.procSelects.Item(1)
            currentProc.procSelects.Remove 1
            currentProc.procNests.Remove 1
            tokens.Item(1).tokType = tokstatement
            tokens.Item(1).tokPCode = vbPCodeEndSelect
            tokens.Item(1).tokString = "EndSelect"
            output_stack.Add tokens.Item(1)
            currentProc.procStatements.Add output_stack
        Case KW_SUB
            If tokens.count <> 1 Then Err.Raise 1
            If currentProc.MemberType <> vbext_mt_Method Then Err.Raise 1
            If Not currentProc.procFunctionResultType Is Nothing Then Err.Raise 1
            EndProc
        Case KW_TYPE
            Err.Raise 1 ' internal error - implemented in Type statement processing
        Case KW_WITH
            If checkNesting(currentProc.procWiths) Then Err.Raise 1 ' Missing With statement
            Set tokens.Item(1).tokWith = currentProc.procWiths.Item(1)
            tokens.Item(1).tokType = tokstatement
            tokens.Item(1).tokPCode = vbPCodeEndWith
            tokens.Item(1).tokString = "EndWith"
            output_stack.Add tokens.Item(1)
            currentProc.procStatements.Add output_stack
            currentProc.procWiths.Remove 1
            currentProc.procNests.Remove 1
        Case Else
            Err.Raise 1 ' Expecting Function, If, Property, Select, Sub, Type or With keyword
    End Select
    tokens.Remove 1
End If

'output_stack.Add token

'currentproc.procStatements.Add output_stack
End Sub




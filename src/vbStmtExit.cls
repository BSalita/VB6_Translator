VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Dim kw As Integer

Set token = tokens.Item(1)
tokens.Remove 1

kw = getKeyword(tokens.Item(1))
' fill in goto stuff!!!!!!!!!!!!!!!!!!!!!!
Select Case tokens.Item(1).tokKeyword
    Case KW_DO
        If currentProc.procDos.count = 0 Then Err.Raise 1 ' Exit without Do
'        Set token.tokDo = currentProc.procDos.Item(1)
        token.tokString = "ExitDo"
        token.tokPCode = vbPCodeExitDo
    Case KW_FOR
        If currentProc.procFors.count = 0 Then Err.Raise 1 ' Exit without For
'        Set token.tokFor = currentProc.procFors.Item(1)
        token.tokString = "ExitFor"
        token.tokPCode = vbPCodeExitFor
    Case KW_FUNCTION
        If currentProc.MemberType <> vbext_mt_Method Then Err.Raise 1
        If currentProc.procFunctionResultType Is Nothing Then Err.Raise 1
        token.tokString = "ExitFunction"
        token.tokPCode = vbPCodeExitFunction
    Case KW_PROPERTY
        If currentProc.MemberType <> vbext_mt_Property Then Err.Raise 1
        ' property procedure may or may not have a return value
        token.tokString = "ExitProperty"
        token.tokPCode = vbPCodeExitProperty
#If 0 Then ' VB doesn't implement
    Case KW_SELECT
        If currentProc.procSelects.count = 0 Then Err.Raise 1 ' Exit without Select
        Set token.tokSelect = currentProc.procSelects.Item(1)
        token.tokString = "ExitSelect"
        token.tokPCode = vbPCodeExitSelect
#End If
    Case KW_SUB
        If currentProc.MemberType <> vbext_mt_Method Then Err.Raise 1
        If Not currentProc.procFunctionResultType Is Nothing Then Err.Raise 1
        token.tokString = "ExitSub"
        token.tokPCode = vbPCodeExitSub
#If 0 Then ' VB doesn't implement
    Case KW_WHILE
        If currentProc.procWhiles.count = 0 Then Err.Raise 1 ' Exit without While
'        Set token.tokWith = currentProc.procWiths.Item(1)
        token.tokString = "ExitWhile"
        token.tokPCode = vbPCodeExitWhile
#End If
End Select
tokens.Remove 1

output_stack.Add token

currentProc.procStatements.Add output_stack
End Sub



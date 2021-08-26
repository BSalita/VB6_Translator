VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtErase"
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
token.tokPCode = vbPCodeErase
tokens.Remove 1
Do
    getArrayVariableAddr tokens, output_stack
    If IsEOL(tokens) Then Exit Do
    special = getSpecialTypes(tokens.Item(1))
    If special = SPECIAL_OP Then
        tokens.Remove 1
        If getSpecialTypes(tokens.Item(1)) <> SPECIAL_CP Then Err.Raise 1 ' Expecting )
        tokens.Remove 1
        If IsEOL(tokens) Then Exit Do
        special = getSpecialTypes(tokens.Item(1))
    End If
    If special <> special_comma Then Err.Raise 1 ' expecting ,
    tokens.Remove 1
Loop Until IsEOL(tokens)
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub


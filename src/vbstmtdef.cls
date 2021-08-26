VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
'Dim output_stack As New Collection
Dim special As SpecialTypes
Dim f As Integer, l As Integer
Set token = tokens.Item(1)
Do
Print #99, "a1"
    tokens.Remove 1
Print #99, "a2"
    If Not isalpha(tokens.Item(1).tokString) Then Err.Raise 1 ' expecting alpha
Print #99, "a3"
    If Len(tokens.Item(1).tokString) <> 1 Then Err.Raise 1 ' expecting single alpha
Print #99, "a4"
    f = asc(UCase(tokens.Item(1).tokString))
Print #99, "a5"
    tokens.Remove 1
Print #99, "a6"
    l = f
Print #99, "a7"
    If Not IsEOL(tokens) Then
Print #99, "b"
        If tokens.Item(1).tokString = "-" Then
Print #99, "c"
            tokens.Remove 1
Print #99, "d"
            If Not isalpha(tokens.Item(1).tokString) Then Err.Raise 1 ' expecting alpha
Print #99, "e"
            If Len(tokens.Item(1).tokString) <> 1 Then Err.Raise 1 ' expecting single alpha
Print #99, "f"
            l = asc(UCase(tokens.Item(1).tokString))
Print #99, "g"
            tokens.Remove 1
Print #99, "h"
            If f > l Then Err.Raise 1 ' expecting ascending range
Print #99, "j"
        End If
    End If
Print #99, "x f="; f; " l="; l
    Dim i As Integer
    For i = f To l
        Dim vt As TliVarType
        vt = 0
        On Error Resume Next
        vt = currentModule.DefaultTypes.Item(Chr(i))
        On Error GoTo 0
        If vt <> 0 Then Err.Raise 1 ' letter range previously used
        currentModule.DefaultTypes.Add VT_R8 ' or whatever is type
    Next
Print #99, "y"
    If IsEOL(tokens) Then Exit Do
    special = getSpecialTypes(tokens.Item(1))
Loop While special = special_comma
'output_stack.Add token
'currentproc.procStatements.Add output_stack
Print #99, "z"
End Sub

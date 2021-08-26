VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtResume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Set token = tokens.Item(1)
tokens.Remove 1
If Not IsEOL(tokens) Then
    If UCase(tokens.Item(1).tokString) = "NEXT" Then ' Resume Next (-1)
        token.tokString = "ResumeNext"
        token.tokPCode = vbPCodeResumeNext
    ElseIf tokens.Item(1).tokType = tokVariant Then
        If CInt(tokens.Item(1).tokString) = 0 Then ' Resume 0 (-2)
            token.tokString = "Resume0"
            token.tokPCode = vbPCodeResume0
        Else ' Resume line number
            token.tokPCode = vbPCodeResumeLabel
            AddLabelRef tokens.Item(1), output_stack
        End If
    Else ' Resume label
        token.tokPCode = vbPCodeResumeLabel
        AddLabelRef tokens.Item(1), output_stack
    End If
    tokens.Remove 1
Else ' Resume (-2)
    token.tokString = "Resume"
    token.tokPCode = vbPCodeResume0
End If
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub



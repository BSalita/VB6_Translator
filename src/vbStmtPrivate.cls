VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtPrivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
If Not currentProc Is Nothing Then Err.Raise 1 ' Private keyword not allowed in procedures
tokens.Remove 1 ' remove Public
Select Case UCase(tokens.Item(1).tokString)
    Case "CONST"
        ParseConst tokens, PROC_ATTR_PRIVATE
    Case "DECLARE"
        ParseDeclare tokens, PROC_ATTR_PRIVATE
    Case "ENUM"
        ParseEnum tokens, PROC_ATTR_PRIVATE
'    Case "EVENT" ' Private Event not possible
    Case "FUNCTION"
        ParseFunction tokens, PROC_ATTR_PRIVATE
    Case "PROPERTY"
        ParseProperty tokens, PROC_ATTR_PRIVATE
    Case "STATIC"
        tokens.Remove 1 ' remove Static
        Select Case UCase(tokens.Item(1).tokString)
'        Case "EVENT" ' Private Event not possible
        Case "FUNCTION"
            ParseFunction tokens, PROC_ATTR_PRIVATE Or PROC_ATTR_Static
        Case "PROPERTY"
            ParseProperty tokens, PROC_ATTR_PRIVATE Or PROC_ATTR_Static
        Case "SUB"
            ParseSub tokens, PROC_ATTR_PRIVATE Or PROC_ATTR_Static
        Case Else
            Err.Raise 1
        End Select
    Case "SUB"
        ParseSub tokens, PROC_ATTR_PRIVATE
    Case "TYPE"
        parseUDT tokens, PROC_ATTR_PRIVATE
    Case Else
        tokens.Add New vbToken, , 1 ' This is stupid, need to rework parseDim?
        parseDim tokens, PROC_ATTR_PRIVATE
End Select
End Sub


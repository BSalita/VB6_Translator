VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Print #99, "in public 1"
If Not currentProc Is Nothing Then Err.Raise 1 ' Public keyword not allowed in procedures
Print #99, "in public 2"
Set token = tokens.Item(1) ' Insert this same code for Private and Static, etc.
tokens.Remove 1 ' remove Public
Print #99, "in public 3 s="; tokens.Item(1).tokString
Select Case UCase(tokens.Item(1).tokString)
    Case "CONST"
        ParseConst tokens, PROC_ATTR_PUBLIC
    Case "DECLARE"
        ParseDeclare tokens, PROC_ATTR_PUBLIC
    Case "ENUM"
        ParseEnum tokens, PROC_ATTR_PUBLIC
    Case "EVENT"
        ParseEvent tokens, PROC_ATTR_PUBLIC
    Case "FUNCTION"
        ParseFunction tokens, PROC_ATTR_PUBLIC
    Case "PROPERTY"
        ParseProperty tokens, PROC_ATTR_PUBLIC
    Case "STATIC"
        tokens.Remove 1 ' remove Static
        Select Case UCase(tokens.Item(1).tokString)
        Case "EVENT"
            ParseEvent tokens, PROC_ATTR_PUBLIC Or PROC_ATTR_Static
        Case "FUNCTION"
            ParseFunction tokens, PROC_ATTR_PUBLIC Or PROC_ATTR_Static
        Case "PROPERTY"
            ParseProperty tokens, PROC_ATTR_PUBLIC Or PROC_ATTR_Static
        Case "SUB"
            ParseSub tokens, PROC_ATTR_PUBLIC Or PROC_ATTR_Static
        Case Else
            Err.Raise 1
        End Select
    Case "SUB"
        ParseSub tokens, PROC_ATTR_PUBLIC
    Case "TYPE"
Print #99, "in public 4"
        parseUDT tokens, PROC_ATTR_PUBLIC
    Case Else
        tokens.Add token, , 1 ' This is stupid, need to rework parseDim?
        parseDim tokens, PROC_ATTR_PUBLIC Or PROC_ATTR_VARIABLE
End Select
End Sub




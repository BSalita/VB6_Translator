VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtFriend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
If Not currentProc Is Nothing Then Err.Raise 1 ' Public keyword not allowed in procedures
tokens.Remove 1 ' remove Friend
Select Case UCase(tokens.Item(1).tokString)
'    Case "EVENT" ' Friend Event not possible
    Case "FUNCTION"
        ParseFunction tokens, PROC_ATTR_Friend
    Case "PROPERTY"
        ParseProperty tokens, PROC_ATTR_Friend
    Case "STATIC"
        tokens.Remove 1 ' remove Static
        Select Case UCase(tokens.Item(1).tokString)
'        Case "EVENT" ' Friend Event not possible
        Case "FUNCTION"
            ParseFunction tokens, PROC_ATTR_Friend Or PROC_ATTR_Static
        Case "PROPERTY"
            ParseProperty tokens, PROC_ATTR_Friend Or PROC_ATTR_Static
        Case "SUB"
            ParseSub tokens, PROC_ATTR_Friend Or PROC_ATTR_Static
        Case Else
            Err.Raise 1
        End Select
    Case "SUB"
        ParseSub tokens, PROC_ATTR_Friend
    Case Else
        Err.Raise 1 ' expecting Function, Static, or Sub keyword
End Select
End Sub



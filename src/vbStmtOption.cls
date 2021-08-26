VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim output_stack As New Collection

' fixme: disable for now. Multiple passes of Option/declare statements is a problem.
#If 0 Then
If currentModule.Consts.count > 0 Then Err.Raise 1 ' Option must appear before declarations
If currentModule.Declares.count > 0 Then Err.Raise 1
If currentModule.Enums.count > 0 Then Err.Raise 1
If currentModule.events.count > 0 Then Err.Raise 1
If currentModule.ModuleVars.count > 0 Then Err.Raise 1
If currentModule.procs.count > 0 Then Err.Raise 1
If currentModule.Types.count > 0 Then Err.Raise 1
#End If

tokens.Remove 1

Select Case UCase(tokens.Item(1).tokString)
    Case "BASE"
        tokens.Remove 1
        If IsEOL(tokens) Then Err.Raise 1 ' Option Base must specify 0 or 1
        If tokens.Item(1).tokType <> tokVariant Then Err.Raise 1 ' Option Base must specify 0 or 1
        If tokens.Item(1).tokString < 0 Or tokens.Item(1).tokString > 1 Then Err.Raise 1 ' Option Base must specify 0 or 1
        currentModule.OptionBase = tokens.Item(1).tokString
        tokens.Remove 1
    Case "COMPARE"
        tokens.Remove 1
        ' Option Compare must specify Binary, DataBase, or Text
        If IsEOL(tokens) Then Err.Raise 1
        Select Case UCase(tokens.Item(1).tokString)
        Case "BINARY"
            currentModule.OptionCompare = 1
        Case "DATABASE"
            currentModule.OptionCompare = 2
        Case "TEXT"
            currentModule.OptionCompare = 3
        Case Else
            Err.Raise 1
        End Select
        tokens.Remove 1
    Case "EXPLICIT"
        tokens.Remove 1
        currentModule.OptionExplicit = True
    Case "PRIVATE"
        tokens.Remove 1
        If IsEOL(tokens) Then Err.Raise 1 ' missing Module keyword
        If UCase(tokens.Item(1).tokString) <> "MODULE" Then Err.Raise 1 ' expecting Module keyword
        tokens.Remove 1
        currentModule.OptionPrivateModule = True ' VB ignores this
    Case Else
        Err.Raise 1 ' unrecognized Option
End Select
End Sub


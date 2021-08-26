VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Dim kw As Keywords
Dim open_access As Integer
Dim open_mode As Integer
Dim open_lock As Integer

Set token = tokens.Item(1)
token.tokPCode = vbPCodeOpen
tokens.Remove 1

oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbString ' Filename expression

If getKeyword(tokens.Item(1)) = KW_FOR Then
    tokens.Remove 1
    Select Case UCase(tokens.Item(1).tokString)
        Case "APPEND"
            open_mode = VB_OPEN_MODE_APPEND
        Case "BINARY"
            open_mode = VB_OPEN_MODE_BINARY
        Case "INPUT"
            open_mode = VB_OPEN_MODE_INPUT
        Case "OUTPUT"
            open_mode = VB_OPEN_MODE_OUTPUT
        Case "RANDOM"
            open_mode = VB_OPEN_MODE_RANDOM
        Case Else
            Err.Raise 1 ' expecting Append, Binary, Input, Output, or Random keyword
    End Select
    tokens.Remove 1
Else
    open_mode = VB_OPEN_MODE_RANDOM
End If

If UCase(tokens.Item(1).tokString) = "ACCESS" Then
    tokens.Remove 1
    Select Case UCase(tokens.Item(1).tokString)
        Case "READ"
            open_access = VB_OPEN_ACCESS_READ
            If UCase(tokens.Item(2).tokString) = "WRITE" Then tokens.Remove 1: open_access = VB_OPEN_ACCESS_READ_WRITE
        Case "WRITE"
            open_access = VB_OPEN_ACCESS_WRITE
        Case Else
            Err.Raise 1 ' expecting read or write keyword
    End Select
    tokens.Remove 1
Else
' Is default READ_WRITE for all open modes?
    open_access = VB_OPEN_ACCESS_READ_WRITE
End If

If UCase(tokens.Item(1).tokString) = "LOCK" Then
    tokens.Remove 1
    Select Case UCase(tokens.Item(1).tokString)
        Case "READ"
            open_lock = VB_OPEN_LOCK_READ
            If UCase(tokens.Item(2).tokString) = "WRITE" Then tokens.Remove 1: open_lock = VB_OPEN_LOCK_READ_WRITE
        Case "WRITE"
            open_lock = VB_OPEN_LOCK_WRITE
        Case Else
            Err.Raise 1 ' expecting read, write, or shared keyword
    End Select
    tokens.Remove 1
ElseIf UCase(tokens.Item(1).tokString) = "SHARED" Then
    tokens.Remove 1
    open_lock = VB_OPEN_LOCK_SHARED
Else
' Need to check if default is VB_OPEN_LOCK_NONE (0)
'    open_lock = VB_OPEN_LOCK_READ
End If

output_stack.Add New vbToken
output_stack.Item(output_stack.count).tokString = "&h" & Hex(open_access Or open_mode Or open_lock)
output_stack.Item(output_stack.count).tokType = tokVariant
output_stack.Item(output_stack.count).tokValue = open_access Or open_mode Or open_lock
output_stack.Item(output_stack.count).tokDataType = vbInteger

kw = getKeyword(tokens.Item(1))
If kw <> KW_AS Then Err.Raise 1 ' expecting As keyword
tokens.Remove 1
getFileNumber tokens, output_stack

If Not IsEOL(tokens) Then
    If UCase(tokens.Item(1).tokString) = "LEN" Then
        tokens.Remove 1
' use function to get =?
        If tokens.Item(1).tokString <> "=" Then Err.Raise 1 ' expecting =
        tokens.Remove 1
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger ' len expression
    Else
        output_stack.Add New vbToken
        output_stack.Item(output_stack.count).tokString = "0"
        output_stack.Item(output_stack.count).tokType = tokVariant
        output_stack.Item(output_stack.count).tokValue = 0
        output_stack.Item(output_stack.count).tokDataType = vbInteger
    End If
Else
    output_stack.Add New vbToken
    output_stack.Item(output_stack.count).tokString = "0"
    output_stack.Item(output_stack.count).tokType = tokVariant
    output_stack.Item(output_stack.count).tokValue = 0
    output_stack.Item(output_stack.count).tokDataType = vbInteger
End If

output_stack.Add token
currentProc.procStatements.Add output_stack

End Sub

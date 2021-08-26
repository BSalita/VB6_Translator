VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtRaiseEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Private Sub vbStmt_stmtParse(ByVal tokens As Collection)
#If 0 Then
Dim token As vbToken
Dim output_stack As New Collection
tokens.Remove 1 ' remove RaiseEvent keyword
oRPN.RPNize gOptimizeFlag, tokens, output_stack, VT_VOID
Set token = output_stack.Item(output_stack.count)
' fixme: must check that top of stack is a callable thing
Print #99, "RaiseEvent dt="; token.tokDataType
output_stack.Add New vbToken
output_stack.Item(output_stack.count).tokString = "RaiseEvent"
output_stack.Item(output_stack.count).tokType = tokstatement
output_stack.Item(output_stack.count).tokDataType = token.tokDataType
output_stack.Item(output_stack.count).tokPCode = vbPCodeRaiseEvent
currentProc.procStatements.Add output_stack
For Each token In output_stack
  Print #99, "t="; token.tokString; " dt="; token.tokDataType; " tc="; token.tokCount
Next
#Else
Dim output_stack As New Collection
Dim special As SpecialTypes
Dim token As vbToken
Dim tokStmt As vbToken
Set tokStmt = tokens.Item(1)
tokens.Remove 1 ' Remove RaiseEvent
tokStmt.tokPCode = vbPCodeRaiseEvent
Set token = tokens.Item(1)
tokens.Remove 1 ' Remove Event name
token.tokType = tokLocalModule
Dim pt As proctable
For Each pt In currentModule.events
Print #99, "pt="; pt.procName; " ik="; pt.InvokeKind
Next
On Error Resume Next
Set token.tokLocalFunction = currentModule.events.Item(SymIK(token.tokString, INVOKE_EVENTFUNC))
On Error GoTo 0
Print #99, "RaiseEvent: ep="; Not token.tokLocalFunction Is Nothing
If token.tokLocalFunction Is Nothing Then
    Print #99, "Unable to process RaiseEvent statement. Event name conflicts with another procedure name: "; token.tokString
    MsgBox "Unable to process RaiseEvent statement. Event name conflicts with another procedure name: " & token.tokString
    Err.Raise 1 ' Event not defined
End If
token.tokPCodeSubType = INVOKE_EVENTFUNC
Dim p As paramTable
Print #99, "pc="; token.tokLocalFunction.procParams.count; " tc="; tokens.count
If token.tokLocalFunction.procParams.count = 0 Then
    If Not IsEOL(tokens) Then
        If tokens.Item(1).tokString = "(" Then
            tokens.Remove 1 ' (
            If tokens.Item(1).tokString <> ")" Then Err.Raise 1
            tokens.Remove 1 ' )
        End If
    End If
Else
Print #99, "1"
    If tokens.Item(1).tokString <> "(" Then Err.Raise 1
Print #99, "2"
    For Each p In token.tokLocalFunction.procParams
Print #99, "3"
        If token.tokCount > 0 Then If tokens.Item(1).tokString <> "," Then Err.Raise 1
Print #99, "4"
        tokens.Remove 1 ' ( or ,
        token.tokCount = token.tokCount + 1 ' parameter count
Print #99, "5"
        Dim dt As TliVarType
        dt = p.paramVariable.varType.dtDataType
        If p.paramVariable.varAttributes And VARIABLE_BYREF Then dt = dt Or VT_BYREF
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, dt
Print #99, "6"
    Next
Print #99, "7"
    If tokens.Item(1).tokString <> ")" Then Err.Raise 1
Print #99, "8"
    tokens.Remove 1 ' )
End If
Print #99, "9"
token.tokDataType = VT_VOID
output_stack.Add token ' Event
output_stack.Add tokStmt
currentProc.procStatements.Add output_stack
#End If
' fixme: kludge: only attempting to add item to CoerceObjects, not outputing to stack
' fixme: CoerceObjects collection should be interface oriented, not class oriented
        Dim CvtToken As vbToken
        Set CvtToken = New vbToken
        CvtToken.tokString = currentModule.EventName & "QI"
        CvtToken.tokType = tokQI_Module
'        CvtToken.tokPCode = vbPCodeQIModule
        CvtToken.tokDataType = vbObject
        Set CvtToken.tokModule = currentModule
        Set CvtToken.tokVariable = output_stack.Item(output_stack.count).tokVariable
'        output_stack.Add CvtToken, , , cnt
'        cnt = cnt + 1
        On Error Resume Next
        currentProject.CoerceObjects.Add CvtToken, currentModule.EventGUID
        On Error GoTo 0
End Sub


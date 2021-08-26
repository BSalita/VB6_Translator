VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Dim LHS_output_stack As New Collection
Dim LHS_count As Long
Dim LHS_tokens As New Collection
Dim stmtSet As vbToken

Print #99, "Set: 1"
If AssignmentOperator Is Nothing Then Err.Raise 1
Print #99, "Set: 2"
If LastQualifier Is Nothing Then Err.Raise 1
'kludge: bad to pass ik in PCodeSubType???
Print #99, "Set: 3"
LastQualifier.tokPCodeSubType = INVOKE_PROPERTYPUTREF
Print #99, "Set: 4"
Set stmtSet = tokens.Item(1)
Print #99, "Set: 5"
tokens.Remove 1

For Each token In tokens
    tokens.Remove 1
    If token Is AssignmentOperator Then Exit For
    LHS_tokens.Add token
Next
Print #99, "Set: 6 tc="; tokens.count
If token Is Nothing Then Err.Raise 1 ' Internal error
Print #99, "Set: 7 s="; tokens.Item(1).tokString
' use function to get =?
'If tokens.Item(1).tokString <> "=" Then Err.Raise 1
'tokens.Remove 1
Print #99, "Set: 8"

' parse RHS
If UCase(tokens.Item(1).tokString) = "NOTHING" Then
Print #99, "Nothing"
    stmtSet.tokString = "SetNothing"
' Is vbPCodeSetNothing needed???
    stmtSet.tokPCode = vbPCodeSetNothing
    Dim NothingToken As vbToken
    Set NothingToken = tokens.Item(1)
    tokens.Remove 1
    output_stack.Add NothingToken
Else
Print #99, "Obj expr "; tokens.Item(1).tokString
'    stmtSet.tokPCode = vbPCodeSet
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, VT_VOID, , False, False, True
    If Not IsVObj(output_stack.Item(output_stack.count).tokDataType And Not VT_BYREF) Then Err.Raise 1 ' expecting Object or Variant expression
End If

' parse LHS
Set token = SymbolLookUp(gOptimizeFlag, LHS_tokens, LHS_output_stack, INVOKE_PROPERTYPUTREF, False, True, True)
Print #99, "Set: LHS count="; LHS_tokens.count; " dt="; token.tokDataType
If Not IsEOL(LHS_tokens) Then Err.Raise 1 ' invalid lvalue
If Not IsVObj(token.tokDataType And Not VT_BYREF) Then Err.Raise 1 ' expecting variable of data type Object or Variant - could be byref variable
Print #99, "Set: s="; token.tokString; " pst="; token.tokPCodeSubType; " tv="; Not token.tokVariable Is Nothing; " tlf="; Not token.tokLocalFunction Is Nothing; " pc="; stmtSet.tokPCode
If token.tokVariable Is Nothing Then Err.Raise 1
Print #99, "set: at="; Hex(token.tokVariable.varAttributes)
#If 1 Then ' New
' fixme: make And test into =
' fixme: implement tokInvokeKind
' didn't work for Dispatch - trying again - zdev\zooom.vbp
' Set col.Item(1) = Nothing - default member must be inserted on LHS
' Set col = Nothing - uses Put to perform Set so INVOKE_PROPERTYPUT is allowed for Set
If Not CBool(token.tokPCodeSubType And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) Then
    token.tokPCodeSubType = 0
    LHS_output_stack.Add token
    Dim cnt As Long
    cnt = LHS_output_stack.count
    If Not CoerceDefaultMember(LHS_output_stack, LHS_output_stack.count, -1, INVOKE_PROPERTYPUTREF) Then Err.Raise 1
    ' fixme: create routine that forces QI? call it CheckQI?
    ' Coerced default could be a variant
    If IsObj(LHS_output_stack(cnt).tokDataType) Then CoerceObject LHS_output_stack, cnt, LHS_output_stack(cnt).tokDataType, LHS_output_stack(LHS_output_stack.count).tokVariable.varType.dtInterfaceInfo, LHS_output_stack(LHS_output_stack.count).tokVariable.varType.dtClass ' force QI
    Set token = LHS_output_stack.Item(LHS_output_stack.count)
    LHS_output_stack.Remove LHS_output_stack.count
End If
token.tokPCodeSubType = INVOKE_PROPERTYPUTREF
Print #99, "Set: dt="; token.tokDataType
stmtSet.tokDataType = token.tokDataType
If Not IsVObj(token.tokDataType And Not VT_BYREF) Then Err.Raise 1 ' expecting variable of data type Object or Variant - could be byref variable
If Not token.tokVariable.varType.dtClassInfo Is Nothing Then stmtSet.tokValue = token.tokVariable.varType.dtInterfaceInfo.GUID
#Else
If Not token.tokVariable Is Nothing Then If Not token.tokVariable.varType.dtClassInfo Is Nothing Then stmtSet.tokValue = token.tokVariable.varType.dtInterfaceInfo.GUID
#End If
Print #99, "GUID="; stmtSet.tokValue
LHS_output_stack.Add token
If stmtSet.tokPCode = vbPCodeSetNothing Then
'    CoerceObject LHS_output_stack, LHS_output_stack.count, Nothing, Nothing
'    NothingToken.tokDataType = token.tokDataType ' And Not VT_BYREF ' could be object or variant
    NothingToken.tokDataType = IIf((token.tokVariable.varAttributes And (VARIABLE_PUTVAL Or VT_BYREF)) = (VARIABLE_PUTVAL Or VT_BYREF), stmtSet.tokDataType Or VT_BYREF, stmtSet.tokDataType And Not VT_BYREF)
    If token.tokType = tokVariable Or token.tokType = tokArrayVariable Or token.tokType = tokUDT Then
        stmtSet.tokString = "SetNothing"
        stmtSet.tokPCode = vbPCodeSetNothing
    End If
Else
'Print #99, "ii="; token.tokInterfaceInfo.Name; " iin="; token.tokVariable.varType.dtInterfaceInfo.Name
    ' fixme: kludge - need to have PUTVAL varType readily available - need varReturnType
    If token.tokVariable.varAttributes And VARIABLE_PUTVAL Then ' LocalFunction
        CoerceOperand gOptimizeFlag, output_stack, output_stack.count, IIf((token.tokVariable.varAttributes And (VARIABLE_PUTVAL Or VT_BYREF)) = (VARIABLE_PUTVAL Or VT_BYREF), stmtSet.tokDataType Or VT_BYREF, stmtSet.tokDataType And Not VT_BYREF), token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.count).paramVariable.varType.dtInterfaceInfo, token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.count).paramVariable.varType.dtClass
'        CoerceOperand output_stack, output_stack.count, IIf((token.tokVariable.varAttributes And (VARIABLE_PUTVAL Or VT_BYREF)) = (VARIABLE_PUTVAL Or VT_BYREF), vbObject Or VT_BYREF, vbObject And Not VT_BYREF), token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.count).paramVariable.varType.dtInterfaceInfo, token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.count).paramVariable.varType.dtClass
    Else
'        CoerceOperand output_stack, output_stack.count, IIf((token.tokVariable.varAttributes And (VARIABLE_PUTVAL Or VT_BYREF)) = (VARIABLE_PUTVAL Or VT_BYREF), stmtSet.tokDataType Or VT_BYREF, stmtSet.tokDataType And Not VT_BYREF), token.tokVariable.varType.dtInterfaceInfo, token.tokVariable.varType.dtClass
        CoerceOperand gOptimizeFlag, output_stack, output_stack.count, stmtSet.tokDataType And Not VT_BYREF, token.tokVariable.varType.dtInterfaceInfo, token.tokVariable.varType.dtClass
'        CoerceOperand output_stack, output_stack.count, IIf((token.tokVariable.varAttributes And (VARIABLE_PUTVAL Or VT_BYREF)) = (VARIABLE_PUTVAL Or VT_BYREF), vbObject Or VT_BYREF, vbObject And Not VT_BYREF), token.tokVariable.varType.dtInterfaceInfo, token.tokVariable.varType.dtClass
    End If
'    If stmtSet.tokDataType = vbObject Then CoerceObject output_stack, output_stack.count, token.tokVariable.varType.dtInterfaceInfo, token.tokVariable.varType.dtClass
'    If token.tokVariable Is Nothing Or Not token.tokLocalFunction Is Nothing Or Not token.tokMemberInfo Is Nothing Then
    stmtSet.tokPCode = vbPCodeSet
''''    stmtSet.tokPCode = vbPCodePropertySet
Print #99, "Set: tt="; token.tokType; " tv="; Not token.tokVariable Is Nothing; " tlf="; Not token.tokLocalFunction Is Nothing; " tii="; Not token.tokInterfaceInfo Is Nothing; " tmi="; Not token.tokMemberInfo Is Nothing
''''If token.tokType = tokVariable Or token.tokType = tokArrayVariable Or token.tokType = tokUDT Then stmtSet.tokPCode = vbPCodeSet
' fixme: PCodePropertyLet should be eliminated? Used to determine whether 1 or 2
'   C operands are stacked. Should be eliminated and code placed into CEmitter
'    If Not token.tokVariable Is Nothing Then If token.tokVariable.MemberType = vbext_mt_Variable And CBool(token.tokVariable.varAttributes And (VARIABLE_FUNCTION Or VARIABLE_PARAMETER Or PROC_ATTR_PRIVATE Or PROC_ATTR_DEFAULT)) Then stmtSet.tokPCode = vbPCodeSet
End If

For Each token In LHS_output_stack
    output_stack.Add token
Next

output_stack.Add stmtSet ' add Set
currentProc.procStatements.Add output_stack

For Each token In output_stack
    Print #99, "set os.ts="; token.tokString; " dt="; token.tokDataType; " r="; token.tokRank
Next

End Sub

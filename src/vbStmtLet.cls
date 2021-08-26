VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtLet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim output_stack As New Collection
Dim LHS_count As Long
Dim LHS_tokens As New Collection
Dim stmtLet As vbToken

If AssignmentOperator Is Nothing Then Err.Raise 1
If LastQualifier Is Nothing Then Err.Raise 1
'kludge: bad to pass ik in PCodeSubType???
LastQualifier.tokPCodeSubType = INVOKE_PROPERTYPUT
Set stmtLet = tokens.Item(1)
tokens.Remove 1

For Each token In tokens
    tokens.Remove 1
    If token Is AssignmentOperator Then Exit For
    LHS_tokens.Add token
Next
If token Is Nothing Then Err.Raise 1 ' Internal error
' use function to get =?
'If tokens.Item(1).tokString <> "=" Then Err.Raise 1
'tokens.Remove 1

oRPN.RPNize gOptimizeFlag, tokens, output_stack, VT_VOID
LHS_count = output_stack.count
Set token = SymbolLookUp(gOptimizeFlag, LHS_tokens, output_stack, INVOKE_PROPERTYPUT, False, True, True)
If Not IsEOL(LHS_tokens) Then Err.Raise 1 ' invalid lvalue
If token.tokDataType = 0 Then Err.Raise 1 ' Invalid lvalue data type
output_stack.Add token
Print #99, "let: dt="; token.tokDataType; " mt="; token.tokVariable.MemberType
If IsObj(token.tokDataType) Then
    If token.tokLocalFunction Is Nothing Then GoTo 20
    Print #99, "let: mt="; token.tokLocalFunction.MemberType; " ik="; token.tokLocalFunction.InvokeKind
    ' don't insert default member if Property Let Object - PropLetObj1.vbp
    If token.tokLocalFunction.MemberType <> vbext_mt_Property Or token.tokLocalFunction.InvokeKind <> INVOKE_PROPERTYPUT Then
20
#If 0 Then
        CoerceOperand gOptimizeFlag, output_stack, LHS_count, -1, output_stack.Item(output_stack.count).tokInterfaceInfo ' force default, if needed
        ' INVOKE_PROPERTYPUT for fixed string - fstring1.vbp
        CoerceOperand gOptimizeFlag, output_stack, output_stack.count, -1, output_stack.Item(output_stack.count).tokInterfaceInfo, , INVOKE_PROPERTYPUT ' force default, if needed
#Else
        ' INVOKE_PROPERTYPUT for fixed string - fstring1.vbp
        CoerceOperand gOptimizeFlag, output_stack, output_stack.count, -1, output_stack.Item(output_stack.count).tokInterfaceInfo, , INVOKE_PROPERTYPUT ' force default, if needed
' new code: if LHS is uncoercable object, then don't coerce RHS - Picture = Picture (default Handle is ik of 0, not Let)
        If Not IsObj(output_stack.Item(output_stack.count).tokDataType) Then CoerceOperand gOptimizeFlag, output_stack, LHS_count, -1, output_stack.Item(output_stack.count).tokInterfaceInfo ' force default, if needed
#End If
    End If
End If
Print #99, "let: ts="; output_stack.Item(output_stack.count).tokString; " pcst="; output_stack.Item(output_stack.count).tokPCodeSubType
If output_stack.Item(output_stack.count).tokPCodeSubType And Not INVOKE_PROPERTYPUT Then
    ' c.Item(1) = "" where Item is only a Func/Get
    If output_stack.Item(output_stack.count).tokDataType <> vbVariant Then Err.Raise 1 ' internal error
    Set token = DispatchDefaultMember(INVOKE_PROPERTYPUT)
    output_stack.Add token
End If
' fixme: implement tokInvokeKind
'If output_stack.Item(output_stack.count).tokPCodeSubType <> INVOKE_PROPERTYPUT Then output_stack.Add DispatchDefaultMember(ik)
#If 0 Then ' cls = cls becomes cls.default = cls.default
' note: the following catches a probable bug (lack of error message) in VB
' fixme: the following test suggests that InvokeKind should be a field in tokVariable
If token.tokDataType = vbObject Then
    If token.tokVariable.varType.dtType = tokProjectClass Then
        If token.tokVariable.MemberType = vbext_mt_Variable Then
            MsgBox "Cannot use LET assignment statement to assign a value to an object variable, use SET instead: " & token.tokString
            Print #99, "Cannot use LET assignment statement to assign a value to an object variable, use SET instead: " & token.tokString
        End If
    End If
End If
#End If
stmtLet.tokPCode = vbPCodeLet
#If 0 Then
' fixme: PCodePropertyLet should be eliminated? Used to determine whether 1 or 2
'   C operands are stacked. Should be eliminated and code placed into CEmitter
Print #99, "Let: tt="; token.tokType; " tv="; Not token.tokVariable Is Nothing; " tlf="; Not token.tokLocalFunction Is Nothing; " tii="; Not token.tokInterfaceInfo Is Nothing; " tmi="; Not token.tokMemberInfo Is Nothing
'If Not token.tokVariable Is Nothing Then If token.tokVariable.MemberType = vbext_mt_Variable And CBool(token.tokVariable.varAttributes And (VARIABLE_FUNCTION Or VARIABLE_PARAMETER Or PROC_ATTR_PRIVATE Or PROC_ATTR_DEFAULT)) Then stmtLet.tokPCode = vbPCodeLet
If output_stack(output_stack.count).tokType = tokVariable Or output_stack(output_stack.count).tokType = tokArrayVariable Or output_stack(output_stack.count).tokType = tokUDT Then
    If output_stack(output_stack.count).tokVariable.Component.Type = vbext_mt_stdmodule Then
        stmtLet.tokPCode = vbPCodeLet
    End If
End If
#End If
'If token.tokVariable Is Nothing Or Not token.tokLocalFunction Is Nothing Or Not token.tokMemberInfo Is Nothing Then
'    stmtLet.tokPCode = vbPCodePropertyLet
'Else
'    stmtLet.tokPCode = vbPCodelet
'End If
Print #99, "let: s="; token.tokString; " v="; Not token.tokVariable Is Nothing; " tlf="; Not token.tokLocalFunction Is Nothing; " pc="; stmtLet.tokPCode
Print #99, "let: at="; Hex(token.tokVariable.varAttributes); " mt="; token.tokVariable.MemberType
'stmtLet.tokDataType = token.tokDataType
stmtLet.tokDataType = output_stack(output_stack.count).tokDataType
' If INVOKE_PROPERTYPUT and BYREF then coerce to pointer
'CoerceOperand output_stack, LHS_count, IIf((token.tokVariable.varAttributes And (VARIABLE_PUTVAL Or VT_BYREF)) = (VARIABLE_PUTVAL Or VT_BYREF), stmtLet.tokDataType Or VT_BYREF, stmtLet.tokDataType And Not VT_BYREF)
' bmpinfo.vbp - Image1.Picture = LoadPicture(ImageFile) - If LHS is an obj (wasn't coerced), coerce QI
If Not token.tokLocalFunction Is Nothing And (token.tokVariable.varAttributes And (VARIABLE_PUTVAL Or VT_BYREF)) = (VARIABLE_PUTVAL Or VT_BYREF) Then
' Test using FStack.vbp, Property1.vbp, Project6.vbp
' AddRef needed for ByRef Property Let assignments as in FStack.vbp
' fixme: testing MemberType = vbt_mt_Property fails. Why? Using Not token.tokLocalFunction is Nothing instead
'''' was stmtlet.tokdatatype, but shouldn't ever change RHS data type - c = mid("",1)
    CoerceOperand gOptimizeFlag, output_stack, LHS_count, stmtLet.tokDataType Or VT_BYREF, output_stack.Item(output_stack.count).tokVariable.varType.dtInterfaceInfo
    ' VT_ARRAY is here so VT_BYREF isn't removed - assumes VT_ARRAY and VT_BYREF are always paired
ElseIf Not token.tokMemberInfo Is Nothing Or CBool(stmtLet.tokDataType And VT_ARRAY) Then
    If Not token.tokMemberInfo Is Nothing Then Print #99, "ik="; token.tokMemberInfo.InvokeKind
    ' RHS must be ByRef if Property Let and (VT_BYREF or PointerLevel = 1) - should already be set in GetReturnType
    ' howtos1r\graphv~1.vbp
    CoerceOperand gOptimizeFlag, output_stack, LHS_count, stmtLet.tokDataType, output_stack.Item(output_stack.count).tokVariable.varType.dtInterfaceInfo
Else
' RHS should always be dereferenced?
'''' was stmtlet.tokdatatype, but shouldn't ever change RHS data type - c = mid("",1)
    CoerceOperand gOptimizeFlag, output_stack, LHS_count, stmtLet.tokDataType And Not VT_BYREF, output_stack.Item(output_stack.count).tokVariable.varType.dtInterfaceInfo
End If
output_stack.Add stmtLet ' add Let
currentProc.procStatements.Add output_stack

For Each token In output_stack
    Print #99, "let: os.ts="; token.tokString; " dt="; token.tokDataType; " pst="; token.tokPCodeSubType
Next

End Sub


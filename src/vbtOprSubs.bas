Attribute VB_Name = "vbtOprSubs"
Option Explicit

Sub UnaryOperatorInit(ByVal cbMe As vbOpr)
Dim o As Variant
On Error Resume Next ' a unary operator might raise an error
For Each o In cVariantTypes
    cbMe.oprLetResultTypeUnary varType(o), varType(cbMe.oprOperateUnary(o))
Next
cbMe.oprLetResultTypeUnary vbVariant, vbVariant
End Sub

Sub BinaryOperatorInit(ByVal cbMe As vbOpr)
Dim l As Variant
Dim r As Variant
On Error Resume Next ' Is operator can raise an error
For Each l In cVariantTypes
    For Each r In cVariantTypes
#If 0 Then
'Print #98, "pcode,l,r="; cbMe.oprPCode; varType(l); varType(r);
        Dim vt As TliVarType
        vt = 0
        vt = varType(cbMe.oprOperateBinary(l, r))
        If vt = 0 Then
            If varType(l) = vbObject Xor varType(r) = vbObject Then
                If varType(l) = vbObject Then
                    vt = varType(cbMe.oprOperateBinary(CVar(1), r))
                Else
                    vt = varType(cbMe.oprOperateBinary(l, CVar(1)))
                End If
            End If
        End If
'Print #98, vt
        cbMe.oprLetResultTypeBinary varType(l), varType(r), vt
#Else
        cbMe.oprLetResultTypeBinary varType(l), varType(r), varType(cbMe.oprOperateBinary(l, r))
#End If
    Next
    cbMe.oprLetResultTypeBinary varType(l), vbVariant, vbVariant
    cbMe.oprLetResultTypeBinary vbVariant, varType(l), vbVariant
Next
cbMe.oprLetResultTypeBinary vbVariant, vbVariant, vbVariant
End Sub

' eliminate Me by using token.tokoperator
Function UnaryOptimize(ByVal output_stack As Collection, ByVal token As vbToken, ByVal cbMe As vbOpr) As Boolean
Print #99, "UnaryOptimize: RHS type="; output_stack.Item(token.tokRHS).tokType; " value="; output_stack.Item(token.tokRHS).tokValue; " dt="; output_stack.Item(token.tokRHS).tokDataType
If HasValue(output_stack.Item(token.tokRHS)) Then
    Dim t As New vbToken
    t.tokType = tokVariant
    t.tokValue = cbMe.oprOperateUnary(output_stack.Item(token.tokRHS).tokValue)
    t.tokString = CStr(t.tokValue) ' used for debugging
    t.tokDataType = varType(t.tokValue)
    output_stack.Add t, , , token.tokRHS
    output_stack.Remove token.tokRHS
    UnaryOptimize = True
End If
End Function

' eliminate Me by using token.tokoperator
Function BinaryOptimize(ByVal output_stack As Collection, ByVal token As vbToken, ByVal cbMe As vbOpr) As Boolean
' fixme: probably should implement sanity check. check that vartype(value) = datatype (in hasvalue?)
'        and vartype(result) = datatype of result
Print #99, "BinaryOptimize: LHS type="; output_stack.Item(token.tokLHS).tokType;
Print #99, " LHS dt="; output_stack.Item(token.tokLHS).tokDataType; "LHS vt="; varType(output_stack.Item(token.tokLHS).tokValue);
If varType(output_stack.Item(token.tokLHS).tokValue) = vbObject Then Print #99, "object" Else Print #99, output_stack.Item(token.tokLHS).tokValue
Print #99, "BinaryOptimize: RHS type="; output_stack.Item(token.tokRHS).tokType;
Print #99, " RHS dt="; output_stack.Item(token.tokRHS).tokDataType; "RHS vt="; varType(output_stack.Item(token.tokRHS).tokValue);
If varType(output_stack.Item(token.tokRHS).tokValue) = vbObject Then Print #99, "object" Else Print #99, output_stack.Item(token.tokRHS).tokValue
If HasValue(output_stack.Item(token.tokLHS)) And HasValue(output_stack.Item(token.tokRHS)) Then
    ' new - don't write over existing token because it munges EnumMemberRPN stack.
    Dim t As New vbToken
    t.tokType = tokVariant
Print #99, "1"
    t.tokValue = cbMe.oprOperateBinary(output_stack.Item(token.tokLHS).tokValue, output_stack.Item(token.tokRHS).tokValue)
Print #99, "2"
    t.tokString = CStr(t.tokValue) ' used for debugging
    ' Optimization could change datatype. Is this OK? - vartype(2 / 2) = vbDouble (not vbInteger)
    t.tokDataType = varType(t.tokValue)
    output_stack.Add t, , , token.tokLHS
    output_stack.Remove token.tokLHS
    output_stack.Remove token.tokRHS
    BinaryOptimize = True
    Print #99, "BinaryOptimize: optimized: dt="; output_stack.Item(token.tokLHS).tokDataType; " value=";
    If varType(output_stack.Item(token.tokLHS).tokValue) = vbObject Then Print #99, "object" Else Print #99, output_stack.Item(token.tokLHS).tokValue
End If
Print #99, "BinaryOpimize: opt="; BinaryOptimize
End Function

' eliminate Me by using token.tokoperator
Sub UnaryOutput(ByVal OptimizeFlag As Integer, ByVal output_stack As Collection, ByVal token As vbToken, ByVal cbMe As vbOpr)
If token.tokRHS < 1 Then Err.Raise 1 ' was < 2 but -1 didn't work
Print #99, token.tokString;
Print #99, "["; output_stack.Item(token.tokRHS).tokDataType; ":"; output_stack.Item(token.tokRHS).tokString; "]"
If OptimizeFlag And OptimizeConstantExpressions Then If UnaryOptimize(output_stack, token, cbMe) Then Exit Sub
token.tokPCodeSubType = cbMe.oprCoerceOperandUnary(OptimizeFlag, output_stack, token)
token.tokDataType = cbMe.oprGetResultTypeUnary(output_stack.Item(token.tokRHS).tokDataType And Not VT_BYREF)
If token.tokDataType = 0 Then Err.Raise 1 ' operator can't handle data type
token.tokPCode = cbMe.oprPCode
output_stack.Add token
End Sub

' eliminate me by using token.tokoperator
Sub BinaryOutput(ByVal OptimizeFlag As Integer, ByVal output_stack As Collection, ByVal token As vbToken, ByVal cbMe As vbOpr)
If token.tokLHS = 0 Or token.tokRHS = 0 Then Err.Raise 1
Print #99, "["; output_stack.Item(token.tokLHS).tokDataType; ":"; output_stack.Item(token.tokLHS).tokString; "]";
Print #99, token.tokString;
Print #99, "["; output_stack.Item(token.tokRHS).tokDataType; ":"; output_stack.Item(token.tokRHS).tokString; "]"
If OptimizeFlag And OptimizeConstantExpressions Then If BinaryOptimize(output_stack, token, cbMe) Then Exit Sub
token.tokPCodeSubType = cbMe.oprCoerceOperandsBinary(OptimizeFlag, output_stack, token)
token.tokDataType = cbMe.oprGetResultTypeBinary(output_stack.Item(token.tokLHS).tokDataType And Not VT_BYREF, output_stack.Item(token.tokRHS).tokDataType And Not VT_BYREF)
If token.tokDataType = 0 Then Err.Raise 1 ' operator can't handle data type
token.tokPCode = cbMe.oprPCode
output_stack.Add token
Print #99, "BinaryOutput: dt="; token.tokDataType
End Sub


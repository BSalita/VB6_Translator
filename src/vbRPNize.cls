VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbRPNize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Sub ConstantRPNize(ByVal OptimizeFlag As Long, ByVal tokens As Collection, ByVal output_stack As Collection, ByVal CoerceDataType As TliVarType, Optional ByVal ik As InvokeKinds = INVOKE_FUNC Or INVOKE_PROPERTYGET, Optional ByVal StmtMethod As Boolean)
Dim osc As Long
Print #99, "ConstantRPNize: of="; OptimizeFlag; " tc="; tokens.count; " osc="; output_stack.count; " dt="; CoerceDataType; " ik="; ik; " sm="; StmtMethod
osc = output_stack.count + 1
' Don't force coercion to variant - it generates xToVar conversion token - just ignore
RPNize OptimizeFlag, tokens, output_stack, IIf(CoerceDataType = vbVariant, -1, CoerceDataType), ik, StmtMethod
If osc <> output_stack.count Then Err.Raise 1 ' not a constant expression - expression has terms
If Not HasValue(output_stack.Item(osc)) Then Err.Raise 1 ' expecting constant value
End Sub

' Needs to be cleaned up
Sub RPNize(ByVal OptimizeFlag As Long, ByVal tokens As Collection, ByVal output_stack As Collection, ByVal CoerceDataType As TliVarType, Optional ByVal ik As InvokeKinds = INVOKE_FUNC Or INVOKE_PROPERTYGET, Optional ByVal StmtMethod As Boolean, Optional ByVal LetStmt As Boolean, Optional ByVal NoInsertObjDefault As Boolean, Optional ByVal DefaultValue As Variant, Optional CoerceInterfaceInfo As InterfaceInfo, Optional CoerceModule As vbModule)
Dim operator_stack As New Collection
Dim token As vbToken
Dim special As SpecialTypes
Dim opop As Boolean
Print #99, "RPNize: of="; OptimizeFlag; " tc="; tokens.count; " osc="; output_stack.count; " dt="; CoerceDataType; " ik="; ik; " sm="; StmtMethod; " oe="; NoInsertObjDefault
Dim original_output_stack_count As Long
original_output_stack_count = output_stack.count
If IsEOL(tokens) Then Err.Raise 1 ' expecting expression
If NoInsertObjDefault Then
    If CoerceDataType <> VT_VARIANT And CoerceDataType <> VT_DISPATCH And CoerceDataType <> VT_VOID And CoerceDataType <> -1 Then Err.Raise 1
Else
' ObjDefault1.vbp
'    Print #99, "RPNize: NoInsertObjDefault set vt="; CoerceDataType
'    NoInsertObjDefault = (CoerceDataType Mod 65536) = vbObject Or (CoerceDataType Mod 65536) = vbVariant
End If
Print #99, "RPNize: oe="; NoInsertObjDefault
Do Until IsEOL(tokens)
Print #99, "RPNize: 1 ts="; tokens.Item(1).tokString; " tc="; tokens.count
    Set token = tokens.Item(1)
    Print #99, "RPNize: tokType="; token.tokType; " tokString="; token.tokString; " opop="; opop
' fixme: "NEW" is kludgy, make all RPN specials into tokxxxx?
    If opop Then If token.tokType = toksymbol Or token.tokType = tokWith Or token.tokType = tokVariant Or UCase(token.tokString) = "NEW" Then Exit Do
    tokens.Remove 1
    Select Case token.tokType
        Case tokKeyword
            Dim t As vbToken
            Select Case getKeyword(token)
                Case KW_NEW
                    Set t = OutputClass(tokens, output_stack)
                    t.tokType = tokNewObject
                    t.tokString = "New"
                    t.tokDataType = vbObject ' fixme: Can New create non-dispatch classes?
                    output_stack.Add t
                    ' add to collection of New Classes
                    Print #99, "New type="; t.tokVariable.varType.dtType
                    Select Case t.tokVariable.varType.dtType
                        Case tokProjectClass, tokFormClass
                            On Error Resume Next
                            currentProject.NewObjects.Add t, t.tokVariable.varType.dtClass.GUID
                            On Error GoTo 0
                        Case tokReferenceClass
                            On Error Resume Next
                            currentProject.NewObjects.Add t, t.tokVariable.varType.dtClassInfo.GUID
                            On Error GoTo 0
                        Case tokIDispatchInterface
                            On Error Resume Next
                            currentProject.NewObjects.Add t, "stdole.IDispatch"
                            On Error GoTo 0
                        Case Else
                            Err.Raise 1
                    End Select
                    Exit Do
                Case KW_LBOUND
                    output_bound OptimizeFlag, tokens, output_stack, tokLBound
                Case KW_UBOUND
                    output_bound OptimizeFlag, tokens, output_stack, tokUBound
                Case KW_TYPEOF
                    ' Object variable/function, New and Nothing not allowed, can't use RPNize because Is will be processed
                ' similar to code below
                    Set t = Scope.ScopeLookUp(OptimizeFlag, tokens, output_stack, INVOKE_PROPERTYGET Or INVOKE_FUNC, False, True)
                    output_stack.Add t
Print #99, "RPNize: TypeOf: 1 s="; t.tokString; " dt="; t.tokDataType
                    If (t.tokDataType And Not VT_BYREF) <> vbVariant And (t.tokDataType And Not VT_BYREF) <> vbObject And t.tokDataType <> vbUserDefinedType Then Err.Raise 1
Print #99, "RPNize: TypeOf: 3"
                    If UCase(tokens.Item(1).tokString) <> "IS" Then Err.Raise 1 ' Expecting IS operator
Print #99, "RPNize: TypeOf: 4"
                    tokens.Remove 1
                    Dim tt As New vbToken
                    Set tt = New vbToken
                    Set tt.tokVariable = New vbVariable
                    Set tt.tokVariable.varType = getDataType(tokens, False)
Print #99, "RPNize: TypeOf: 5 dt="; tt.tokDataType
                    If (t.tokDataType And Not VT_BYREF) <> vbVariant And (t.tokDataType And Not VT_BYREF) <> tt.tokVariable.varType.dtDataType Then Err.Raise 1
                    tt.tokType = tokTypeOf
                    tt.tokString = "TypeOf"
                    tt.tokDataType = vbBoolean
                    output_stack.Add tt
                Case Else
                    If tokens.count = 0 Then tokens.Add token Else tokens.Add token, , 1 ' re-add keyword token
                    Exit Do
            End Select
            opop = True
        Case toksymbol, tokWith
            If tokens.count = 0 Then tokens.Add token Else tokens.Add token, , 1 ' re-add keyword token
            If OptimizeFlag = ProcessCompilerExpression Then ' = is correct, not And
                Set token = Nothing
                On Error Resume Next
                Set token = currentModule.cPredefinedConstant.Item(UCase(tokens.Item(1).tokString))
                On Error GoTo 0
                If token Is Nothing Then
                    On Error Resume Next
                    Set token = g_cPredefinedConstant.Item(UCase(tokens.Item(1).tokString))
                    On Error GoTo 0
                    If token Is Nothing Then
                        Set token = New vbToken
                        token.tokType = tokVariant
                        ' OK to make Empty?
                        ' token.tokValue = Empty - not needed - defaults to Empty
                    End If
                End If
                tokens.Remove 1
                output_stack.Add token
            Else
3
                Set token = Scope.ScopeLookUp(OptimizeFlag, tokens, output_stack, ik, StmtMethod, NoInsertObjDefault)
                output_stack.Add token
'                If token.tokDataType = VT_LPWSTR And CBool(ik And INVOKE_PROPERTYGET) Then CoerceOperand OptimizeFlag, output_stack, output_stack.count, VT_BSTR, Nothing, Nothing, ik
                ' Line Input #1, fs
                ' note: overloading purpose of NoInsertObjDefault
                If Not NoInsertObjDefault And token.tokDataType = VT_LPWSTR Then CoerceOperand OptimizeFlag, output_stack, output_stack.count, VT_BSTR, Nothing, Nothing, ik
            End If
'            If StmtMethod And token Is LastQualifier Then Exit Do ' err.raise (1)
            opop = True
Print #99, "RPNize: 2"
        Case tokNothing
            output_stack.Add token
            opop = True
        Case tokVariant
            Print #99, "RPNize: operand push: tokvariant: vartype="; varType(token.tokValue);
            If varType(token.tokValue) <> vbObject Then Print #99, " value="; token.tokValue; " tokString="; token.tokString;
            Print #99,
'            token.tokDataType = varType(token.tokValue) ' shouldn't need
   ' need to do something with literalpool
  '          If token.tokDataType = vbString Then
  ' this line aborts
   '             If currentProc.procLiteralPool.Item(token.tokValue) Is Nothing Then currentProc.procLiteralPool.Add token, token.tokValue
    '        End If
            output_stack.Add token
            opop = True
        Case tokSpecial
            special = getSpecialTypes(token)
            If special = SPECIAL_OP Then
                Do
                    RPNize OptimizeFlag, tokens, output_stack, VT_VOID
                    special = getSpecialTypes(tokens.Item(1))
                    tokens.Remove 1
                Loop While special = special_comma
                If special <> SPECIAL_CP Then Err.Raise 1
                opop = True
            ElseIf special = SPECIAL_CP Or special = special_comma Or special = special_sc Then ' Or special = special_underline Then
                If tokens.count = 0 Then tokens.Add token Else tokens.Add token, , 1 ' re-add SPECIAL_CP token
                Exit Do
            Else
                Err.Raise 1 ' unexpected special symbol
            End If
        Case tokColon
            If tokens.count = 0 Then tokens.Add token Else tokens.Add token, , 1 ' re-add SPECIAL_CP token
            Exit Do
#If 0 Then
        Case tokWith
#If 0 Then ' moved to ScopeLookUp
            If currentProc.procWiths.count = 0 Then
                Print #99, "RPNize: Missing With statement"
                MsgBox "RPNize: Missing With statement"
                Err.Raise 1
            End If
            Set token = New vbToken
            token.tokType = tokWithValue
            token.tokString = "WithValue(" & CStr(currentProc.procWiths.count) & ")"
            Set token.tokVariable = currentProc.procWiths.Item(1).WithValue.tokVariable
            Set token.tokInterfaceInfo = currentProc.procWiths.Item(1).WithValue.tokInterfaceInfo
            Set token.tokMemberInfo = currentProc.procWiths.Item(1).WithValue.tokMemberInfo
            token.tokCount = currentProc.procWiths.count
            token.tokDataType = currentProc.procWiths.Item(1).WithValue.tokDataType
            ' CreateArgStack only used for leading "!"? - !Whatever = !Whatever --> .("Whatever") = .("Whatever")
            Dim astk As New Collection
            If tokens.Item(1).tokString = "(" Then Set astk = Scope.CreateArgStack(token, tokens, New Collection)
            Do
                Set token = Scope.StackVariable(token, tokens, astk, output_stack, ik, StmtMethod) ', NoInsertObjDefault)
                If IsEOL(tokens) Then Exit Do
                Print #99, "RPNize: t="; tokens.Item(1).tokType; " osc="; output_stack.count
'    If tokens.Item(1).tokType <> tokMember Then Exit Do
'    Set arg_stack = CreateArgStack(ScopeLookUp, tokens, New Collection, StmtMethod) ' ScopeLookUp.tokVariable.varType.dtType = tokIDispatchInterface)
'        output_stack.Add ScopeLookUp
            Loop While tokens.Item(1).tokType = tokMember
#End If
            If tokens.count = 0 Then tokens.Add token Else tokens.Add token, , 1 ' re-add keyword token
            Set token = Scope.ScopeLookUp(OptimizeFlag, tokens, output_stack, ik, StmtMethod, NoInsertObjDefault)
            output_stack.Add token
            If StmtMethod And token Is LastQualifier Then Exit Do ' .Raise (1)
            opop = True
#End If
        Case tokOperator
            If LetStmt Then
                If tokens.count = 0 Then tokens.Add token Else tokens.Add token, , 1 ' re-add operator token
                Exit Do
            End If
            On Error Resume Next
process_operator:
            If opop Then
                Set token.tokOperator = Nothing
                Set token.tokOperator = cBinaryOperators.Item(UCase(token.tokString))
                On Error GoTo 0
                If token.tokOperator Is Nothing Then Err.Raise 1 ' internal error - unknown operator
                Do While operator_stack.count > 0
                    If token.tokOperator.oprPriority > operator_stack.Item(1).tokOperator.oprPriority Then Exit Do
                    operator_stack.Item(1).tokRHS = output_stack.count
                    operator_stack.Item(1).tokOperator.oprOutput OptimizeFlag, output_stack, operator_stack.Item(1)
                    Print #99, "RPNize: binary operator pop: tokType="; operator_stack.Item(1).tokType; " tokString="; token.tokString; " tokPriority="; operator_stack.Item(1).tokOperator.oprPriority
                    operator_stack.Remove 1
                Loop
                If operator_stack.count = 0 Then operator_stack.Add token Else operator_stack.Add token, , 1
                token.tokLHS = output_stack.count
                Print #99, "RPNize: binary operator push: tokType="; operator_stack.Item(1).tokType; " tokString="; token.tokString; " tokPriority="; operator_stack.Item(1).tokOperator.oprPriority
            Else
                Set token.tokOperator = Nothing
                Set token.tokOperator = cUnaryOperators.Item(UCase(token.tokString))
                On Error GoTo 0
                If token.tokOperator Is Nothing Then Err.Raise 1 ' internal error - unknown operator
                If operator_stack.count = 0 Then operator_stack.Add token Else operator_stack.Add token, , 1
                Print #99, "RPNize: unary operator push: tokString="; token.tokString
            End If
            opop = False
        Case Else
            ' VB doesn't catch Error.Raise 1 s/b Err.Raise 1
            MsgBox "Unsupported syntax or syntax error not caught by VB: " & token.tokString
            Print #99, "Unsupported syntax or syntax error not caught by VB: " & token.tokString
            Err.Raise 1
    End Select
10
Print #99, "RPNize: 4"
Loop
Print #99, "RPNize: output_stack="; output_stack.count; " operator_stack="; operator_stack.count
' fixme: following is debug code to stop run away stacking - remove or adjust for production
If output_stack.count > 1024 Or operator_stack.count > 1024 Then Err.Raise 1
For Each token In operator_stack
    Print #99, "RPNize: eoe: operator pop: t="; token.tokType; " ts="; token.tokString; " tp="; token.tokOperator.oprPriority
    token.tokRHS = output_stack.count
    token.tokOperator.oprOutput OptimizeFlag, output_stack, token
    Print #99, "RPNize: 5"
Next
Print #99, "RPNize: 6"
For Each token In output_stack
    Print #99, "RPNize: ot="; token.tokString; " t="; token.tokType; " pc="; token.tokPCode; " pcs="; token.tokPCodeSubType; " tc="; token.tokCount; " r="; token.tokRank; " dt="; token.tokDataType; " v="; Not token.tokVariable Is Nothing
Next
Print #99, "RPNize: 7: CoerceDataType="; CoerceDataType; " oosc="; original_output_stack_count; " osc="; output_stack.count
If original_output_stack_count = output_stack.count Then
    If CoerceDataType <> VT_VOID Then
        Print #99, "RPNize: 8"
        Set token = New vbToken
        If CoerceDataType = vbVariant Or CoerceDataType = -1 Then
            MissingArg token, vbVariant
        Else
            If IsMissing(DefaultValue) Then Err.Raise 1 ' Missing value, no default
            token.tokType = tokVariant
            token.tokValue = DefaultValue
            token.tokString = DefaultValue ' remove when tokstring is fixed
            token.tokDataType = varType(DefaultValue)
        End If
        output_stack.Add token
    End If
ElseIf output_stack.Item(output_stack.count).tokType <> tokArrayVariable Or output_stack.Item(output_stack.count).tokRank <> 0 Then
    Print #99, "RPNize: 9"
    Print #99, "RPNize: v="; Not output_stack.Item(output_stack.count).tokVariable Is Nothing; ; " dt="; output_stack.Item(output_stack.count).tokDataType
    ' If CoerceDataType is a variant then always coerce - debug.print class1
    ' Line Input #1, v - don't want to generate default - using NoInsertObjDefault=True
    CoerceOperand OptimizeFlag, output_stack, output_stack.count, CoerceDataType, Nothing, Nothing, ik, NoInsertObjDefault
    ' Skip processing of Nothing - its tokVariable is Nothing
    If CoerceDataType = vbObject And Not output_stack.Item(output_stack.count).tokVariable Is Nothing Then
        Print #99, "RPNize: t="; output_stack.Item(output_stack.count).tokVariable.varType.dtType
        If output_stack.Item(output_stack.count).tokDataType <> CoerceDataType Then Err.Raise 1
        If output_stack.Item(output_stack.count).tokDataType <> vbObject Then Err.Raise 1
        CoerceObject output_stack, output_stack.count, vbObject, CoerceInterfaceInfo, CoerceModule
    End If
End If
Print #99, "RPNize: done: osc="; output_stack.count
End Sub

Function OutputClass(ByVal tokens As Collection, ByVal output_stack As Collection) As vbToken
Print #99, "OutputClass"
Set OutputClass = New vbToken
Set OutputClass.tokVariable = New vbVariable
OutputClass.tokVariable.MemberType = vbext_mt_Variable
Set OutputClass.tokVariable.varType = getProjectTLibDataType(tokens)
Print #99, "OutputClass: vt="; OutputClass.tokVariable.varType Is Nothing
If OutputClass.tokVariable.varType Is Nothing Then Err.Raise 1 ' Class name not found
OutputClass.tokDataType = vbObject
If Not OutputClass.tokVariable.varType.dtClass Is Nothing Then
    OutputClass.tokString = OutputClass.tokString & " " & OutputClass.tokVariable.varType.dtClass.Name
ElseIf Not OutputClass.tokVariable.varType.dtClassInfo Is Nothing Then
    OutputClass.tokString = OutputClass.tokString & " " & TypeInfoToVBType(OutputClass.tokVariable.varType.dtClassInfo)
End If
End Function

Sub output_bound(ByVal OptimizeFlag As Long, ByVal tokens As Collection, ByVal output_stack As Collection, ByVal bound As vbPCodes)
Dim t As vbToken
Print #99, "output_bound: tc="; tokens.count; " osc="; output_stack.count; " b="; bound
If tokens.Item(1).tokString <> "(" Then Err.Raise 1
tokens.Remove 1 ' (
getArrayVariableAddr tokens, output_stack
If tokens.Item(1).tokString = "," Then
    tokens.Remove 1 ' ,
    RPNize OptimizeFlag, tokens, output_stack, vbLong
Else
    Set t = New vbToken
    t.tokType = tokVariant
    t.tokString = "1"
    t.tokValue = 1&
    t.tokDataType = vbLong
    output_stack.Add t
End If
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1 ' )
CoerceOperand OptimizeFlag, output_stack, output_stack.count, VT_VARIANT
Set t = New vbToken
t.tokType = bound
t.tokString = "Bound"
t.tokDataType = vbLong
output_stack.Add t
End Sub

' fixme: Need to implement circular def detection. Pass Object (vbConst, vbEnumMember) and use Is?
' VB assures circular def doesn't occur.
Function EvalConstRPNStack(ByVal tokens As Collection, ByVal dt As TliVarType, Optional ByVal l As Long = -1) As Variant
Print #99, "EvalConstRPNStack: t.c="; tokens.count; " dt="; dt; " pn="; PassNumber
Dim token As vbToken
Dim output_stack As New Collection
For Each token In tokens
    Print #99, "s="; token.tokString; " t="; token.tokType
    Select Case token.tokType
    Case tokConst
        Dim c As vbConst
        Set c = token.tokConst
        Print #99, "c="; c.ConstName; " e="; IsEmpty(c.ConstValue)
        If IsEmpty(c.ConstValue) Then
            c.ConstValue = EvalConstRPNStack(c.ConstRPN, c.ConstDataType.dtDataType, c.ConstDataType.dtLength)
            If c.ConstDataType.dtDataType = vbVariant Then
                If Not IsEmpty(c.ConstValue) And Not IsNull(c.ConstValue) Then
                    Set c.ConstDataType = New vbDataType
                    c.ConstDataType.dtDataType = varType(c.ConstValue)
                End If
            End If
        End If
        token.tokValue = c.ConstValue
        output_stack.Add token
    Case tokConstantInfo
        output_stack.Add token
    Case tokCvt
'        ' throw away Variant coercions, leave Value data type as is
'        If token.tokDataType <> vbVariant Then CoerceOperand OptimizeFlag, output_stack, output_stack.cnt, token.tokDataType
    Case tokEnumMember
        Dim em As vbEnumMember
        Set em = token.tokEnumMember
        Print #99, "em="; em.enumMemberName; " e="; IsEmpty(em.enumMemberValue)
        If IsEmpty(em.enumMemberValue) Then em.enumMemberValue = EvalConstRPNStack(em.enumMemberRPN, vbLong)
        token.tokValue = em.enumMemberValue
        token.tokDataType = vbLong
        output_stack.Add token
    Case tokVariant
        output_stack.Add token
    Case tokOperator
        token.tokLHS = output_stack.count - 1
        token.tokRHS = output_stack.count
        token.tokOperator.oprOutput OptimizeConstantExpressions, output_stack, token
    Case Else
        Err.Raise 1
    End Select
Next
If output_stack.count <> 1 Then Err.Raise 1 ' Expecting constant expression
EvalConstRPNStack = CoerceConstant(output_stack.Item(1).tokValue, dt, l)
Print #99, "EvalConstRPNStack: v="; EvalConstRPNStack; " vt="; varType(EvalConstRPNStack)
End Function

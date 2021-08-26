VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbtScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Function ScopeLookUp(ByVal OptimizeFlag As Long, ByVal tokens As Collection, ByVal output_stack As Collection, ByVal ik As InvokeKinds, ByVal StmtMethod As Boolean, ByVal NoInsertObjDefault As Boolean, Optional ByVal ReDimFlag As Boolean) As vbToken
Print #99, "ScopeLookUp: s="; tokens.Item(1).tokString; " ik="; ik; " sm="; StmtMethod; " oe="; NoInsertObjDefault; " rdf="; ReDimFlag
Dim token As vbToken
Set token = tokens.Item(1)
tokens.Remove 1
Dim arg_stack As Collection
Set arg_stack = CreateArgStack(token, tokens, New Collection, StmtMethod, ReDimFlag)
Print #99, "asc="; arg_stack.Count
If token.tokType = tokWith Then
    If currentProc.procWiths.Count = 0 Then
        Print #99, "ScopeLookUp: Missing With statement"
        MsgBox "ScopeLookUp: Missing With statement"
        Err.Raise 1
    End If
    Set ScopeLookUp = New vbToken
    ScopeLookUp.tokType = tokWithValue
    ' Not using tokWith for tokWithValue token. Copying to tok fields. Perhaps should???
    Set ScopeLookUp.tokWith = currentProc.procWiths.Item(1)
    ScopeLookUp.tokCount = ScopeLookUp.tokWith.WithCount
    ScopeLookUp.tokDataType = ScopeLookUp.tokWith.WithValue.tokDataType And Not VT_BYREF ' never a VT_BYREF
    ScopeLookUp.tokString = "WithValue(" & CStr(ScopeLookUp.tokCount) & ")"
    Set ScopeLookUp.tokVariable = ScopeLookUp.tokWith.WithValue.tokVariable
    Set ScopeLookUp.tokInterfaceInfo = ScopeLookUp.tokWith.WithValue.tokInterfaceInfo
    Set ScopeLookUp.tokMemberInfo = ScopeLookUp.tokWith.WithValue.tokMemberInfo
    ' CreateArgStack only used for leading "!"? - !Whatever = !Whatever --> .("Whatever") = .("Whatever")
'    If ScopeLookUp.Item(1).tokString = "(" Then Set arg_stack = Scope.CreateArgStack(ScopeLookUp, tokens, New Collection, StmtMethod, ReDimFlag)
Else
    ' fixme: kludge probably need to pass token and tokens to each routine
    If tokens.Count = 0 Then tokens.Add token Else tokens.Add token, , 1
    ' fixme: create collection of LookUps instead of index - use For Each to interate
    Dim i As Integer
    For i = 1 To 5
        Set ScopeLookUp = ScopeLookUpLevel(i, OptimizeFlag, tokens, arg_stack, output_stack, ik, StmtMethod)
        If Not ScopeLookUp Is Nothing Then Exit For
        ' x(1) = "" where x is Get and Set requires Let search fails so do Get+Default
    '    If ik And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then If arg_stack.count > 0 Then Set ScopeLookUp = ScopeLookUpLevel(i, OptimizeFlag, tokens, arg_stack, output_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET, StmtMethod)
        If ik And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then Set ScopeLookUp = ScopeLookUpLevel(i, OptimizeFlag, tokens, arg_stack, output_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET, StmtMethod)
        If Not ScopeLookUp Is Nothing Then Exit For
    Next
    If ScopeLookUp Is Nothing Then
        Set ScopeLookUp = CreateVariant(tokens, ReDimFlag)
        If ReDimFlag Then token.tokRank = arg_stack.Count: GoTo done
    End If
    If arg_stack.Count = 0 Then Set arg_stack = CreateArgStack(ScopeLookUp, tokens, New Collection, StmtMethod, ReDimFlag)
End If
' Module1.f(...)
Print #99, "asc="; arg_stack.Count
Do
    Set ScopeLookUp = StackVariable(ScopeLookUp, tokens, arg_stack, output_stack, ik, StmtMethod, NoInsertObjDefault, ReDimFlag)
    Print #99, "1000"
'        Set ScopeLookUp = Scope.StackVariable(Nothing, tokens, output_stack, ik, StmtMethod, NoInsertObjDefault)
    If IsEOL(tokens) Then Exit Do
    Print #99, "1001"
    Print #99, "t="; tokens.Item(1).tokType; " osc="; output_stack.Count
'    If tokens.Item(1).tokType <> tokMember Then Exit Do
'    Set arg_stack = CreateArgStack(ScopeLookUp, tokens, New Collection, StmtMethod,redimflag) ' ScopeLookUp.tokVariable.varType.dtType = tokIDispatchInterface)
'        output_stack.Add ScopeLookUp
Loop While tokens.Item(1).tokType = tokMember
done:
Print #99, "rdf="; ReDimFlag; " asc="; arg_stack.Count
If ReDimFlag Then
    Dim c As Collection
    For Each c In arg_stack
        Print #99, "c="; c.Count
        For Each token In c
            Print #99, " t="; token.tokString
            output_stack.Add token
        Next
    Next
End If
Print #99, "ScopeLookUp: found="; Not ScopeLookUp Is Nothing
End Function

' Obtain first member of object expression
Function ScopeLookUpLevel(ByVal ScopeLevel As Integer, ByVal OptimizeFlag As Long, ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection, ByVal ik As InvokeKinds, ByVal StmtMethod As Boolean) As vbToken
Print #99, "ScopeLookUpLevel: level="; ScopeLevel; " ts="; tokens.Item(1).tokString; " asc="; arg_stack.Count; " ik="; ik; " sm="; StmtMethod
Select Case ScopeLevel
    Case 1 ' Local scope
        Set ScopeLookUpLevel = LocalScope(tokens, arg_stack, currentProc, StmtMethod)
    Case 2 ' Module scope
        Set ScopeLookUpLevel = ModuleScope(tokens, arg_stack, output_stack, currentModule, ik)
    ' case 2.5 ' Module specifiers
        If ScopeLookUpLevel Is Nothing Then
            If UCase(tokens.Item(1).tokString) = "ME" Then
                If IsAny(currentModule.Component.Type, vbext_ct_StdModule, vbext_ct_ResFile, vbext_ct_RelatedDocument) Then Err.Raise 1 ' Me not allowed in these modules
                Set ScopeLookUpLevel = CreateMe(tokens.Item(1), currentModule)
                tokens.Remove 1 ' remove Me
'            ElseIf UCase(tokens.Item(1).tokString) = "USERCONTROL" And currentModule.Component.Type = vbext_ct_UserControl Then ' axmrquee\axmrquee.vbp
'                Set ScopeLookUpLevel = CreateMe(tokens.Item(1), currentModule)
'                tokens.Remove 1 ' remove UserControl
            ElseIf UCase(tokens.Item(1).tokString) = UCase(currentModule.Name) Then
                ' fixme: If currentModule.Component.Type <> stdmodule,form,global (appobject) class? Then Err.Raise 1 ' current module name not allowed
                Set ScopeLookUpLevel = CreateMe(tokens.Item(1), currentModule)
                tokens.Remove 1 ' remove current module name
            End If
        End If
    Case 3 ' Global scope
        Set ScopeLookUpLevel = GlobalScope(tokens, arg_stack, output_stack, ik)
    Case 4 ' Project name
        If IsCurrentProjectName(tokens) Then
            Set ScopeLookUpLevel = GlobalScope(tokens, arg_stack, output_stack, ik)
            If ScopeLookUpLevel Is Nothing Then Err.Raise 1 ' Specified project contains no declaration for specifed member or module
        End If
    Case 5 ' Is the symbol a TLib name or TLib member?
        Set ScopeLookUpLevel = getTLIFunction(tokens, output_stack, arg_stack, ik, StmtMethod)
    Case Else
        Print #99, "ScopeLookUpLevel: unknown scope level"
        Err.Raise 1
End Select
Print #99, "ScopeLookUpLevel: fnd="; Not ScopeLookUpLevel Is Nothing; " acs="; arg_stack.Count
End Function

Function CreateMe(ByVal token As vbToken, ByVal m As vbModule) As vbToken
Print #99, "CreateMe: ts="; token.tokString; " m="; m.Name; " m.c.t="; m.Component.Type
Select Case m.Component.Type
' vbext_ct_UserControl axmrquee\axmrquee.vbp
Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_PropPage, vbext_ct_UserControl ', vbext_ct_ActiveXDesigner
    Set CreateMe = GetForm(m)
Case Else
    token.tokType = tokme
    token.tokDataType = vbObject
    Set token.tokVariable = New vbVariable
    Set token.tokVariable.varType = New vbDataType
    Set token.tokVariable.varType.dtClass = m
    token.tokVariable.varType.dtType = tokProjectClass
    Set CreateMe = token
End Select
End Function

Function LocalScope(ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal p As proctable, ByVal StmtMethod As Boolean) As vbToken
Print #99, "LocalScope: s="; tokens.Item(1).tokString; " asc="; arg_stack.Count
If p Is Nothing Then Exit Function
Dim varname As String
varname = UCase(tokens.Item(1).tokString)
On Error Resume Next
' Checking variables according to likelyhood of match
Dim v As vbVariable
Set v = p.procLocalVariables.Item(varname)
Print #99, "v="; Not v Is Nothing; " c="; p.procParams.Count; " p="; p.procName
If v Is Nothing Then
    Set v = p.procParams.Item(varname).paramVariable
End If
If v Is Nothing Then
' Is it the name of the containing function/property?
' should varSymbol be made UCase?
    If varname = UCase(p.procFunctionResultType.varSymbol) Then
        If Not p.procFunctionResultType Is Nothing Then
'            If Not StmtMethod Or Not tokens.Item(1) Is LastQualifier Then
'                If tokens.Item(1) Is LastQualifier And arg_stack.count = 0 Then
                If arg_stack.Count = 0 Then
                    Set v = p.procFunctionResultType
                End If
'            End If
'        Else
'            Print #99, "Assignment to Property Let/Set name not allowed"
'            Err.Raise 1
        End If
    End If
End If
If v Is Nothing Then
    Set LocalScope = LookUpConst(p.procConsts, tokens)
Else
    If Not v Is Nothing Then Set LocalScope = CreateVariableToken(tokens, v)
End If
'If Not LocalScope Is Nothing Then tokens.Remove 1
Print #99, "LocalScope: fnd="; Not LocalScope Is Nothing
End Function

#If 0 Then ' unneeded - Could use modified version for getDataTypeInModule
Function LookUpUDT(ByVal types As Collection, ByVal tokens As Collection, Optional pa As procattributes = PROC_ATTR_DEFAULT Or PROC_ATTR_PRIVATE Or PROC_ATTR_PUBLIC) As vbToken
Print #99, "LookUpUDT: types="; Not types Is Nothing; " t.c="; tokens.Count; " pa="; Hex(pa)
Print #99, "sym="; tokens.Item(1).tokString
If types Is Nothing Then
    ' search all modules except current module, which was previously searched
    Dim m As vbModule
    For Each m In currentProject.prjModules
        If Not m Is currentModule Then
            ' Type defaults scope is Public
            Set LookUpUDT = LookUpUDT(m.types, tokens, PROC_ATTR_PUBLIC Or PROC_ATTR_DEFAULT) ' recursive
' fixme: check here for ambigous definitions (same symbol, multiple modules)?
            If Not LookUpUDT Is Nothing Then Exit For
        End If
    Next
Else
    Dim t As vbType
    For Each t In types
        Print #99, "t="; t.typeName; " guid="; t.typeGUID; " m.c="; t.typeMembers.Count
    Next
    On Error Resume Next
    Set t = types.Item(UCase(tokens.Item(1).tokString))
    On Error GoTo 0
    Print #99, "t="; Not t Is Nothing
    If Not t Is Nothing Then
        Print #99, "1"
        If pa And t.typeAttributes Then
            Set LookUpUDT = tokens.Item(1)
            Print #99, "2"
            tokens.Remove 1 ' remove type
            Print #99, "3"
            LookUpUDT.tokType = tokUDT
            Print #99, "4"
            LookUpUDT.tokString = t.typeName
            Print #99, "5"
            LookUpUDT.tokDataType = vbUserDefinedType
            Print #99, "6"
            Set LookUpUDT.tokVariable = New vbVariable
            Print #99, "7"
            Set LookUpUDT.tokVariable.varType = New vbDataType
            Print #99, "8"
            LookUpUDT.tokVariable.varType.dtType = tokUDT
            Print #99, "9"
            Set LookUpUDT.tokVariable.varType.dtUDT = t
            Print #99, "10"
        End If
    End If
    Print #99, "11"
End If
Print #99, "LookUpUDT: fnd="; Not LookUpUDT Is Nothing
End Function
#End If

Function LookUpConst(ByVal Consts As Collection, ByVal tokens As Collection, Optional pa As procattributes = PROC_ATTR_DEFAULT Or PROC_ATTR_PRIVATE Or PROC_ATTR_PUBLIC) As vbToken
Print #99, "LookUpConst: consts="; Not Consts Is Nothing; " t.c="; tokens.Count; " pa="; Hex(pa)
If Consts Is Nothing Then
    ' search all modules except current module, which was previously searched
    Dim m As vbModule
    For Each m In currentProject.prjModules
        If Not m Is currentModule Then
            ' Const default scope is Private
            Set LookUpConst = LookUpConst(m.Consts, tokens, PROC_ATTR_PUBLIC) ' recursive
' fixme: should this check for ambigous definitions (same symbol, multiple modules)
            If Not LookUpConst Is Nothing Then Exit For
        End If
    Next
Else
    Dim varname As String
    varname = UCase(tokens.Item(1).tokString)
    Print #99, "vn="; varname
    Dim c As vbConst
    For Each c In Consts
        Print #99, "c="; c.ConstName; " dt="; c.ConstDataType.dtDataType; " cv.vt="; varType(c.ConstValue)
    Next
    On Error Resume Next
    Set c = Consts.Item(varname)
    On Error GoTo 0
    Print #99, "c="; Not c Is Nothing
    If Not c Is Nothing Then
        Print #99, "1"
        If pa And c.ConstAttributes Then
            Set LookUpConst = tokens.Item(1)
            Print #99, "2"
            tokens.Remove 1 ' remove const
            Print #99, "3"
            LookUpConst.tokType = tokConst
            Print #99, "4"
            LookUpConst.tokString = c.ConstName
            Print #99, "5"
            LookUpConst.tokDataType = c.ConstDataType.dtDataType
            LookUpConst.tokValue = c.ConstValue
            Print #99, "6 dt="; c.ConstDataType.dtDataType; " cv.vt="; varType(c.ConstValue); " e="; IsEmpty(c.ConstValue)
#If 0 Then
            If PassNumber = 2 And IsEmpty(c.ConstValue) Then
            Print #99, "7"
                Dim output_stack As New Collection
                oRPN.ConstantRPNize OptimizeConstantExpressions, c.ConstRPN, output_stack, c.ConstDataType.dtDataType
                Print #99, "8"
                c.ConstValue = output_stack.Item(1).tokValue
                Set output_stack = Nothing
            End If
#End If
            Print #99, "9"
            LookUpConst.tokValue = c.ConstValue
            Set LookUpConst.tokConst = c
            Print #99, "10"
        End If
    End If
    Print #99, "8"
End If
Print #99, "LookUpConst: fnd="; Not LookUpConst Is Nothing
End Function

Function ModuleScope(ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection, ByVal m As vbModule, ByVal ik As InvokeKinds) As vbToken
Print #99, "ModuleScope: tc="; tokens.Count; " asc="; arg_stack.Count; " m="; m Is Nothing
Print #99, "ModuleScope: ts="; tokens.Item(1).tokString; " m="; m.Name
On Error Resume Next
' Checking variables according to likelyhood of match
' fixme: need LookUpVar similar to LookUpUDT
Dim v As vbVariable
Set v = m.ModuleVars.Item(UCase(tokens.Item(1).tokString))
On Error GoTo 0
#If 0 Then
Dim v As vbVariable
For Each v In m.ModuleVars
    Print #99, "vs="; v.varSymbol
Next
Dim p As proctable
For Each p In m.procs
    Print #99, "pn="; p.procName
Next
#End If
If v Is Nothing Then
'    Set ModuleScope = LookUpUDT(m.types, tokens)
    If ModuleScope Is Nothing Then Set ModuleScope = LookUpConst(m.Consts, tokens)
    If ModuleScope Is Nothing Then Set ModuleScope = LookUpEnumMember(m, tokens)
    If ModuleScope Is Nothing Then Set ModuleScope = getDeclare(m.Declares, tokens, arg_stack, output_stack)
    If ModuleScope Is Nothing Then Set ModuleScope = getLocalFunction(m, tokens, arg_stack, output_stack, ik)
    If ModuleScope Is Nothing Then
        Select Case m.Component.Type
            Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_PropPage, vbext_ct_UserControl ', vbext_ct_ActiveXDesigner
                Set ModuleScope = FormSearch(m, tokens, arg_stack, output_stack, ik)
        End Select
    End If
Else
    Set ModuleScope = CreateVariableToken(tokens, v)
End If
Print #99, "ModuleScope: fnd="; Not ModuleScope Is Nothing
End Function

' fixme: shouldn't GlobalScope consider appobject classes to be processed like stdmodules?
' fixme: what is proper search order for globals? Search every module for a single type before going to next type (var, enum, const) repeat search for next type, or complete each module search for all types before going to next?
Function GlobalScope(ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack, ByVal ik As InvokeKinds) As vbToken
Print #99, "GlobalScope: ts="; tokens.Item(1).tokString; " asc="; arg_stack.Count
Dim varname As String
varname = UCase(tokens.Item(1).tokString)
Dim m As vbModule
On Error Resume Next
Set m = currentProject.prjModules.Item(varname)
On Error GoTo 0
' fixme: allow stdmodule,form,global (appobject) classes?
If Not m Is Nothing Then
    Select Case m.Component.Type
    ' Module1.x
    ' Designer.x
    Case vbext_ct_StdModule, vbext_ct_ActiveXDesigner
        tokens.Remove 1 ' remove module name
        If IsEOL(tokens) Then Err.Raise 1 ' Missing .
        If tokens.Item(1).tokType <> tokMember Then Err.Raise 1 ' Expecting .
        tokens.Remove 1 ' remove period
        If IsEOL(tokens) Then Err.Raise 1 ' Missing class name
        If tokens.Item(1).tokType <> toksymbol Then Err.Raise 1 ' expecting module name
        Set GlobalScope = ModuleScope(tokens, arg_stack, output_stack, m, ik)
    ' fixme: Need IsForm(ct)
    ' Form1.x
    Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_PropPage, vbext_ct_UserControl ', vbext_ct_ActiveXDesigner
        Set GlobalScope = GetForm(m)
        tokens.Remove 1
    Case Else
        ' fixme: assumes all classes cannot be referenced (cls.x). OK for appobjects????
        Set m = Nothing
'        Err.Raise 1 ' Reference to component type not implemented
    End Select
End If
If m Is Nothing Then
    Dim v As vbVariable
    Dim mm As vbModule
    For Each mm In currentProject.prjModules
        Print #99, "mm="; mm.Name; " v="; Not v Is Nothing; " n="; varname; " c.c="; mm.Consts.Count
        Dim vv As vbVariable
        If Not mm Is currentModule Then
            If mm.Component.Type = vbext_ct_StdModule Then
                ' fixme: need to make these lookups into funcs
                On Error Resume Next
                Set vv = mm.ModuleVars.Item(varname)
                On Error GoTo 0
            End If
            If Not vv Is Nothing Then If vv.varAttributes And (PROC_ATTR_PRIVATE Or PROC_ATTR_DEFAULT) Then Set vv = Nothing
            Print #99, "vv="; Not vv Is Nothing
            If Not vv Is Nothing Then
                If Not m Is Nothing Then
                    Print #99, vv.varSymbol & " is defined in multiple modules. Needs qualifier: " & m.Name & ", " & mm.Name
                    MsgBox vv.varSymbol & " is defined in multiple modules. Needs qualifier: " & m.Name & ", " & mm.Name
                    Err.Raise 1
                End If
                Set m = mm
                Set v = vv
                Set vv = Nothing
            End If
        End If
    Next
    Print #99, "2 v="; Not v Is Nothing
    If v Is Nothing Then
        Set GlobalScope = LookUpConst(Nothing, tokens)
        If GlobalScope Is Nothing Then Set GlobalScope = LookUpEnumMember(Nothing, tokens)
        If GlobalScope Is Nothing Then Set GlobalScope = getDeclare(Nothing, tokens, arg_stack, output_stack)
        If GlobalScope Is Nothing Then Set GlobalScope = getLocalFunction(Nothing, tokens, arg_stack, output_stack, ik)
    Else
        Set GlobalScope = CreateVariableToken(tokens, v)
    End If
End If
Print #99, "GlobalScope: fnd="; Not GlobalScope Is Nothing
End Function

Function CreateVariableToken(ByVal tokens As Collection, ByVal v As vbVariable) As vbToken
Print #99, "CreateVariableToken: ts="; tokens.Item(1).tokString
Set CreateVariableToken = tokens.Item(1)
tokens.Remove 1
Set CreateVariableToken.tokVariable = v
Print #99, "CreateVariableToken: s="; v.varSymbol; " mt="; v.MemberType; " dims="; Not v.varDimensions Is Nothing; " type="; v.varType.dtType
Select Case v.MemberType ' need to translate other types enum?
    Case vbext_mt_Variable
        CreateVariableToken.tokType = IIf(CreateVariableToken.tokVariable.varDimensions Is Nothing, tokVariable, tokArrayVariable)
' fixme: need to implement tokInvokeKind and assign correct INVOKE type
        CreateVariableToken.tokPCodeSubType = INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF
    Case vbext_mt_Const
        CreateVariableToken.tokType = tokConst
        CreateVariableToken.tokValue = v.varVariant
        CreateVariableToken.tokPCodeSubType = INVOKE_PROPERTYGET
    Case Else
        Err.Raise 1 ' internal error
End Select
CreateVariableToken.tokDataType = v.varType.dtDataType Or IIf(v.varDimensions Is Nothing, 0, VT_ARRAY Or VT_BYREF)
' new
CreateVariableToken.tokDataType = CreateVariableToken.tokDataType Or IIf(v.varAttributes And VARIABLE_BYREF, VT_BYREF, 0)
Set CreateVariableToken.tokModule = v.varModule
' possibly more tok members should be initialized
Set CreateVariableToken.tokInterfaceInfo = v.varType.dtInterfaceInfo
If Not CreateVariableToken.tokInterfaceInfo Is Nothing Then Print #99, "CreateVariableToken: ii.n="; CreateVariableToken.tokInterfaceInfo.Name; " ii.guid="; CreateVariableToken.tokInterfaceInfo.GUID; " ii.am="; Hex(CreateVariableToken.tokInterfaceInfo.AttributeMask)
Print #99, "CreateVariableToken: dt="; CreateVariableToken.tokDataType; " type="; CreateVariableToken.tokType
End Function

' fixme: make other project-wide module scanning lookups implement similar recursive algorithm
Function LookUpEnumMember(ByVal m As vbModule, ByVal tokens As Collection, Optional pa As procattributes = PROC_ATTR_DEFAULT Or PROC_ATTR_PRIVATE Or PROC_ATTR_PUBLIC) As vbToken
Print #99, "LookUpEnumMember: m="; Not m Is Nothing; " t.c="; tokens.Count
If m Is Nothing Then
    ' search all modules except current module, which was previously searched
    For Each m In currentProject.prjModules
        If Not m Is currentModule Then
            ' Enum default scope is Public
            Set LookUpEnumMember = LookUpEnumMember(m, tokens, PROC_ATTR_PUBLIC Or PROC_ATTR_DEFAULT) ' recursive
' fixme: should this check for ambigous definitions (same symbol, multiple modules)
            If Not LookUpEnumMember Is Nothing Then Exit For
        End If
    Next
Else
    Dim e As vbEnum
    Dim em As vbEnumMember
    Dim varname As String
    Print #99, "m="; m.Name; " s="; tokens.Item(1).tokString; " ct="; m.Component.Type
    varname = UCase(tokens.Item(1).tokString)
'    If m.Component.Type = vbext_ct_StdModule Then
        Print #99, "LookUpEnumMember: 1"
        On Error Resume Next
        Set e = m.Enums.Item(varname)
        On Error GoTo 0
        Print #99, "LookUpEnumMember: 2 e="; Not e Is Nothing
        If Not e Is Nothing Then
            Print #99, "LookUpEnumMember: 12 a="; e.enumAttributes
' fixme: this suggests there could/should be separate collections for Public vs. Private?!
            If pa And e.enumAttributes Then
                Print #99, "LookUpEnumMember: 3"
                tokens.Remove 1 ' remove Enum name
                Print #99, "LookUpEnumMember: 4"
                If tokens.Item(1).tokType <> tokMember Then Err.Raise 1
                Print #99, "LookUpEnumMember: 5"
                tokens.Remove 1 ' remove .
                Print #99, "LookUpEnumMember: 6 s="; tokens.Item(1).tokString
                If tokens.Item(1).tokType <> toksymbol Then Err.Raise 1
                Print #99, "LookUpEnumMember: 7"
                On Error Resume Next
                Set em = e.enumMembers.Item(UCase(tokens.Item(1).tokString))
                On Error GoTo 0
                Print #99, "LookUpEnumMember: 8"
                If em Is Nothing Then Err.Raise 1 ' expecting member of e.enumName
            End If
        Else
            Print #99, "LookUpEnumMember: 10"
            For Each e In m.Enums
                Print #99, "LookUpEnumMember: 12 n="; e.enumName; " a="; e.enumAttributes; " mc="; e.enumMembers.Count
' fixme: this suggests there could/should be separate collections for Public vs. Private?!
                If pa And e.enumAttributes Then
                    On Error Resume Next
                    Set em = e.enumMembers.Item(varname)
                    On Error GoTo 0
                    If Not em Is Nothing Then Exit For
                End If
            Next
        End If
        Print #99, "LookUpEnumMember: 12"
        If Not em Is Nothing Then
        Print #99, "LookUpEnumMember: 13"
            tokens.Remove 1 ' remove Enum Member
            Set LookUpEnumMember = New vbToken
            LookUpEnumMember.tokType = tokEnumMember
            LookUpEnumMember.tokString = em.enumMemberName
            LookUpEnumMember.tokDataType = vbLong
            LookUpEnumMember.tokValue = em.enumMemberValue
            Set LookUpEnumMember.tokEnumMember = em
        End If
'    End If
End If
Print #99, "LookUpEnumMember: fnd="; Not LookUpEnumMember Is Nothing
End Function

Function getLocalFunction(ByVal LocalModule As vbModule, ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection, ByVal ik As InvokeKinds) As vbToken '
Print #99, "getLocalFunction: lm="; Not LocalModule Is Nothing; " t.c="; tokens.Count; " as.c="; arg_stack.Count; " os.c="; output_stack.Count; " ik="; ik
If IsEOL(tokens) Then
    Print #99, "Expecting name"
    MsgBox "Expecting name"
    Err.Raise 1
End If
Dim token As vbToken
Set token = tokens.Item(1)
Print #99, "s="; token.tokString; " tt="; token.tokType
If token.tokType <> toksymbol And token.tokType <> tokVariable Then
    Print #99, "Invalid qualifier: "; token.tokType
    MsgBox "Invalid qualifier: " & token.tokType
    Err.Raise 1
End If
Dim ProcNameIK As String
' kludge: tokPCodeSubType holds ik assigned in Stmt Let/Set
' fixme: finish obsoleting tokPCodeSubType
'ProcNameIK = symik(token.tokString,IIf(Not AssignmentOperator Is Nothing And token.tokPCodeSubType <> 0, token.tokPCodeSubType, ik)
ProcNameIK = SymIK(token.tokString, ik)
Print #99, "pnik="; ProcNameIK

If LocalModule Is Nothing Then
    For Each LocalModule In currentProject.prjModules
Print #99, "getLocalFunction: lc="; LocalModule.Name; " ct="; LocalModule.Component.Type
Dim pt As proctable
For Each pt In LocalModule.procs
  Print #99, "pnik="; UCase(pt.procName) & pt.InvokeKind
Next
' ik not implemented - skip SET proc if GET reference
        On Error Resume Next
        If LocalModule.Component.Type = vbext_ct_StdModule Then
            Set token.tokLocalFunction = LocalModule.procs.Item(SymIK(token.tokString, ik))
        Else
' Global class method name space implementation would be put here?
            If LocalModule Is currentModule Then
Print #99, "0"
                Set token.tokLocalFunction = currentModule.procs.Item(SymIK(token.tokString, ik))
Print #99, "00"
            End If
        End If
        On Error GoTo 0
        If Not token.tokLocalFunction Is Nothing Then Exit For
    Next
    If LocalModule Is Nothing Then Set LocalModule = currentModule
Else
Print #99, "n="; LocalModule.Name; " c="; LocalModule.procs.Count
    For Each pt In LocalModule.procs
        Print #99, "pt="; pt.procName; " ik="; pt.InvokeKind;
        If Not pt.procFunctionResultType Is Nothing Then Print #99, " dt="; pt.procFunctionResultType.varType.dtDataType;
        Print #99,
'        If UCase(token.tokString) = UCase(pt.procName) And IIf(Not AssignmentOperator Is Nothing And token.tokPCodeSubType <> 0, token.tokPCodeSubType, ik) And pt.InvokeKind Then
'            Exit For
'        End If
    Next
    On Error Resume Next
    Set token.tokLocalFunction = LocalModule.procs.Item(ProcNameIK)
    On Error GoTo 0
' fixme: need to redo public variable processing.
    If token.tokLocalFunction Is Nothing Then
        On Error Resume Next
        Set token.tokLocalFunction = LocalModule.procs.Item(SymIK(token.tokString, INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF))
        On Error GoTo 0
    ElseIf CBool(ik And INVOKE_PROPERTYPUT) And arg_stack.Count >= token.tokLocalFunction.procParams.Count Then
        Set token.tokLocalFunction = Nothing
    End If
Print #99, "ts="; token.tokString; " tlf="; Not token.tokLocalFunction Is Nothing
End If
Print #99, "1"
If Not token.tokLocalFunction Is Nothing Then
Print #99, "4"
    tokens.Remove 1
Print #99, "5"
    If LocalModule Is currentModule Then
        token.tokType = tokLocalModule ' reference is within module
     Else
        token.tokType = tokGlobalModule ' reference is outside of module
    End If
    Set token.tokVariable = token.tokLocalFunction.procFunctionResultType
Print #99, "5b"
    If token.tokVariable Is Nothing Then
Print #99, "5c"
        Set token.tokVariable = New vbVariable
        Set token.tokVariable.varType = New vbDataType
        token.tokVariable.varType.dtType = tokProjectClass
        Set token.tokVariable.varType.dtClass = LocalModule
    End If
Print #99, "6 type="; token.tokVariable.varType.dtType; " c="; Not token.tokVariable.varType.dtClass Is Nothing
    token.tokPCodeSubType = token.tokLocalFunction.InvokeKind
Print #99, "7 ik="; token.tokPCodeSubType
    ' ugh. token.tokLocalFunction.InvokeKind could be combination of flags
    If token.tokLocalFunction.InvokeKind = INVOKE_PROPERTYPUT Or token.tokLocalFunction.InvokeKind = INVOKE_PROPERTYPUTREF Then
Print #99, "8"
        ' fixme: needs tokVarDataType implemented
        token.tokDataType = token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.Count).paramVariable.varType.dtDataType Or IIf(token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.Count).paramVariable.varDimensions Is Nothing, 0, VT_ARRAY Or VT_BYREF)
    ElseIf Not token.tokVariable Is Nothing Then
Print #99, "9"
        token.tokDataType = token.tokVariable.varType.dtDataType Or IIf(token.tokVariable.varDimensions Is Nothing, 0, VT_ARRAY Or VT_BYREF)
' fixme: hmmm, only needed for form variables (Text1)
        Set token.tokInterfaceInfo = token.tokVariable.varType.dtInterfaceInfo
    End If
    Set getLocalFunction = token
End If
Print #99, "getLocalFunction: fnd="; Not getLocalFunction Is Nothing
End Function

Function getTLIFunction(ByVal tokens As Collection, ByVal output_stack As Collection, ByRef arg_stack As Collection, ByVal ik As InvokeKinds, ByVal StmtMethod As Boolean) As vbToken
'Dim coclassinfo As coclassinfo
Print #99, "getTLIFunction: c="; tokens.Count; " asc="; arg_stack.Count; " sm="; StmtMethod; " ik="; ik
If IsEOL(tokens) Then
    Print #99, "Expecting function name"
    MsgBox "Expecting function name"
    Err.Raise 1
End If
Dim reference As reference
Set reference = getTLib(tokens.Item(1).tokString)
'Set classinfo = getTLIClass(tokens, reference)
Print #99, "getTLIFunction: s="; tokens.Item(1).tokString; " ref="; Not reference Is Nothing
If reference Is Nothing Then
    For Each reference In currentProject.VBProject.References
        If SearchTLI(tokens, reference, output_stack, arg_stack, ik, StmtMethod) Then
            Set getTLIFunction = tokens.Item(1)
            tokens.Remove 1
            Exit For
        End If
    Next
Else
    tokens.Remove 1 ' remove TypeLib
    If IsEOL(tokens) Then
        Print #99, "Expecting period"
        MsgBox "Expecting period"
        Err.Raise 1
    End If
    If tokens.Item(1).tokType <> tokMember Then
        Print #99, "Expecting period"
        MsgBox "Expecting period"
        Err.Raise 1
    End If
    tokens.Remove 1 ' remove period
    If IsEOL(tokens) Then
        Print #99, "Expecting qualifier"
        MsgBox "Expecting qualifier"
        Err.Raise 1
    End If
    If tokens.Item(1).tokType <> toksymbol Then
        Print #99, "Expecting qualifier"
        MsgBox "Expecting qualifier"
        Err.Raise 1
    End If
'    token.tokString = tokens.Item(1).tokString ' a bit kludgy - copy in member name
    If arg_stack.Count > 0 Then Err.Raise 1
    ' fixme: klugy, removing first token, then adding in, then removing if found
    Set getTLIFunction = tokens.Item(1)
    tokens.Remove 1
    Set arg_stack = CreateArgStack(getTLIFunction, tokens, arg_stack, StmtMethod)
    If tokens.Count = 0 Then tokens.Add getTLIFunction Else tokens.Add getTLIFunction, , 1
    If Not SearchTLI(tokens, reference, output_stack, arg_stack, ik, StmtMethod) Then
        Print #99, "member name not found in designated type library"
        MsgBox "member name not found in designated type library"
        Err.Raise 1
    End If
    Set getTLIFunction = tokens.Item(1)
    tokens.Remove 1
End If
Print #99, "getTLIFunction: fnd="; Not getTLIFunction Is Nothing
End Function

Function SearchTLI(ByVal tokens As Collection, ByVal reference As reference, ByVal output_stack As Collection, ByVal arg_stack As Collection, ByVal ik As InvokeKinds, ByVal StmtMethod As Boolean) As Boolean
'On Error GoTo done
Print #99, "SearchTLI: s="; tokens.Item(1).tokString; " asc="; arg_stack.Count; " dt="; tokens.Item(1).tokDataType; " ik="; ik
Print #99, "ref="; reference.Name; " fullpath="; reference.FullPath
' Should LCID be 0?
'       On Error Resume Next
'       Set TLib = Nothing
' Doesn't seem to work for builtin references
'       Set TLib = TypeLibInfoFromRegistry(reference.GUID, reference.Major, reference.Minor, 0)
Dim tlib As TypeLibInfo
Set tlib = TypeLibInfoFromFile(reference.FullPath)
'       On Error GoTo 0
If Not tlib Is Nothing Then
    If SearchForAppObject(tokens, tlib, arg_stack, ik) Then
        SearchTLI = True
        GoTo done
    End If
    Dim DeclarationInfo As DeclarationInfo
    For Each DeclarationInfo In tlib.Declarations
        Print #99, "Declaration="; DeclarationInfo.Name
        If UCase(DeclarationInfo.Name) = UCase(tokens.Item(1).tokString) Then
            tokens.Remove 1 ' remove qualifier
            ' fixme: this code is duplicated from above
            ' fixme: probably should keep DeclarationInfo token, perhaps insert missing ones too
            If IsEOL(tokens) Then
                Print #99, "Expecting period"
                MsgBox "Expecting period"
                Err.Raise 1
            End If
            If tokens.Item(1).tokType <> tokMember Then
                Print #99, "Expecting period"
                MsgBox "Expecting period"
                Err.Raise 1
            End If
            tokens.Remove 1 ' remove period
            If IsEOL(tokens) Then
                Print #99, "Expecting qualifier"
                MsgBox "Expecting qualifier"
                Err.Raise 1
            End If
            If tokens.Item(1).tokType <> toksymbol Then
                Print #99, "Expecting qualifier"
                MsgBox "Expecting qualifier"
                Err.Raise 1
            End If
            If arg_stack.Count > 0 Then Err.Raise 1
            ' fixme: klugy, removing first token, then adding in, then removing if found
            Dim token As vbToken
            Set token = tokens.Item(1)
            tokens.Remove 1
            Set arg_stack = CreateArgStack(token, tokens, arg_stack, StmtMethod)
            If tokens.Count = 0 Then tokens.Add token Else tokens.Add token, , 1
            SearchTLI = SearchDeclarationInfo(tokens.Item(1), output_stack, arg_stack, DeclarationInfo, ik)
            GoTo done
        End If
    Next
#If 1 Then
    For Each DeclarationInfo In tlib.Declarations
' could print first time through
        Print #99, "Declaration="; DeclarationInfo.Name
        If SearchDeclarationInfo(tokens.Item(1), output_stack, arg_stack, DeclarationInfo, ik) Then
            SearchTLI = True
            GoTo done
        End If
#Else
    Dim InterfaceInfo As InterfaceInfo
    For Each InterfaceInfo In tlib.Interfaces
' could print first time through
        Print #99, "InterfaceInfo="; InterfaceInfo.Name; " guid="; InterfaceInfo.Name
        Dim MemberInfo As MemberInfo
        For Each MemberInfo In InterfaceInfo.Members
#End If
    Next
    SearchTLI = getConstant(tokens)
End If
done:
If SearchTLI Then Set tokens.Item(1).tokReference = reference
Print #99, "SearchTLI: fnd="; SearchTLI; " tc="; tokens.Count; " asc="; arg_stack.Count
End Function

Function getConstant(ByVal tokens As Collection) As Boolean
Print #99, "getConstant: tc="; tokens.Count
Dim token As vbToken
Set token = tokens.Item(1)
Print #99, "getConstant: constName="; UCase(token.tokString)
Dim reference As reference
Dim ConstName As String
Dim ConstantInfo As ConstantInfo
Dim MemberInfo As MemberInfo
Dim tlib As TypeLibInfo
ConstName = UCase(token.tokString)
For Each reference In currentProject.VBProject.References
'    MsgBox "ref=" & reference.Name & " desc=" & reference.Description
' Should LCID be 0?
    On Error Resume Next ' Error can occur
    Set tlib = Nothing
' Doesn't seem to work for builtin references
'    Set TLib = TypeLibInfoFromRegistry(reference.GUID, reference.Major, reference.Minor, 0)
    Set tlib = TypeLibInfoFromFile(reference.FullPath)
    On Error GoTo 0
    If Not tlib Is Nothing Then
Print #99, "getConstant: searching tlib:"; tlib.Name
        ' can't use tlib.Constants.NamedItem(constName) because it returns InterfaceInfos too - Why?
        For Each ConstantInfo In tlib.Constants
'Print #99, "getConstant: ci=:"; ConstantInfo.Name
            If ConstName = UCase(ConstantInfo.Name) Then Exit For
        Next
        If Not ConstantInfo Is Nothing Then
            Print #99, "1"
            tokens.Remove 1 ' remove ConstantInfo name
            If IsEOL(tokens) Then
            Print #99, "2"
                token.tokType = tokConstantInfo
                token.tokDataType = vbLong
                Set token.tokVariable = New vbVariable
                Set token.tokVariable.varType = New vbDataType
                token.tokVariable.varType.dtType = tokReferenceClass
                Set token.tokVariable.varType.dtConstantInfo = ConstantInfo
                getConstant = True
            Else
                Print #99, "3"
                If tokens.Item(1).tokString <> "." Then Err.Raise 1
                tokens.Remove 1 ' remove .
                getConstant = getConstantInfoMember(ConstantInfo, tokens.Item(1))
                Print #99, "4"
                If Not getConstant Then Err.Raise 1
            End If
        Else
            Print #99, "5"
            For Each ConstantInfo In tlib.Constants
                getConstant = getConstantInfoMember(ConstantInfo, tokens.Item(1))
                If getConstant Then Exit For
            Next
        End If
    End If
    If getConstant Then Exit For
Next
done:
Print #99, "getConstant: fnd="; getConstant
End Function

Function getConstantInfoMember(ByVal ConstantInfo As ConstantInfo, ByVal token As vbToken) As Boolean
'Print #99, "getConstantInfoMember: ci="; ConstantInfo.Name; " m="; ConstantInfo.Members.count
Dim ConstName As String
ConstName = UCase(token.tokString)
'Print #99, "constName="; ConstName; " ci="; ConstantInfo.Name
Dim MemberInfo As MemberInfo
For Each MemberInfo In ConstantInfo.Members
'    Print #99, "mi="; MemberInfo.Name
    If ConstName = UCase(MemberInfo.Name) Then
        getConstantInfoMember = True
        token.tokType = tokConstantInfo
        token.tokValue = MemberInfo.Value
        token.tokDataType = varType(token.tokValue)
        Set token.tokMemberInfo = MemberInfo
        Set token.tokVariable = New vbVariable
        Set token.tokVariable.varType = New vbDataType
        token.tokVariable.varType.dtType = tokReferenceClass
        Set token.tokVariable.varType.dtConstantInfo = ConstantInfo
        Exit For
    End If
Next
'Print #99, "getConstantInfoMember: fnd="; getConstantInfoMember
End Function

Function SearchDeclarationInfo(ByVal token As vbToken, ByVal output_stack As Collection, ByVal arg_stack As Collection, ByVal DeclarationInfo As DeclarationInfo, ByVal ik As InvokeKinds) As Boolean
Dim funcname As String
Dim BStrName As String
Dim BVarName As String
Dim MemberInfo As MemberInfo
Print #99, "SearchDeclarationInfo: ts="; token.tokString; " asc="; arg_stack.Count; " ik="; ik; " di="; DeclarationInfo.Name
' fixme: need standardized way of selecting ik type - different routines using diff methods
' fixme: perhaps each token symbol needs an InvokeKind set to FUNC/GET?
If Not token Is LastQualifier Or token.tokPCodeSubType = 0 Then ik = INVOKE_FUNC Or INVOKE_PROPERTYGET
funcname = UCase(token.tokString)
' FixMe: How to choose between _B_STR_ and _B_VAR funcs (MSVBVM60.DLL). Need to
' process expression first, then choose vbString or vbVariant type.
If token.tokDataType = vbString Then BStrName = "_B_STR_" & funcname ' don't need to compare unless string
BVarName = "_B_VAR_" & funcname
For Each MemberInfo In DeclarationInfo.Members
'Print #99, "MemberInfo="; MemberInfo.Name; " am="; MemberInfo.AttributeMask; " ik="; MemberInfo.InvokeKind; " pc="; MemberInfo.parameters.count; " poc="; MemberInfo.parameters.OptionalCount; " pdc="; MemberInfo.parameters.defaultCount; " vtbl="; MemberInfo.VTableOffset
    If ik And MemberInfo.InvokeKind Then
        Dim MemberName As String
        MemberName = UCase(MemberInfo.Name)
'Print #99, "MemberName="; MemberName
' fixme: would like to use MatchMemberInfo()
        If (MemberName = funcname Or MemberName = BStrName Or MemberName = BVarName) Then
            If IsArgCountValid(MemberInfo, arg_stack) Then
                Set token.tokMemberInfo = MemberInfo
            Else
                Dim t As vbToken
                Dim rtdt As TliVarType
                Dim rtii As InterfaceInfo
                If MemberInfo.ReturnType.varType = vbObject Then Err.Raise 1 ' Set t = InsertDefaultMember(ii, output_stack, arg_stack, rtdt, rtii)
                If t Is Nothing Then
                    Print #99, "arg count mismatch"
                    MsgBox "SearchDeclarationInfo: arg count mismatch: " & MemberInfo.Name
                    Exit For
                End If
                Set MemberInfo = t.tokMemberInfo
            End If
'                AddTLIDeclarationMember DeclarationInfo, MemberInfo
            token.tokType = tokDeclarationInfo
            Set token.tokVariable = New vbVariable
            Set token.tokVariable.varType = New vbDataType
            Set token.tokDeclarationInfo = DeclarationInfo
            Set token.tokMemberInfo = MemberInfo
            token.tokDataType = InitVarTypeFromVarTypeInfo(MemberInfo.ReturnType, token.tokVariable.varType)
            Set token.tokInterfaceInfo = token.tokVariable.varType.dtInterfaceInfo
            AddTLIDeclarationMember DeclarationInfo, MemberInfo
            SearchDeclarationInfo = True
            Exit For
        End If
    End If
Next
Print #99, "SearchDeclarationInfo: fnd="; SearchDeclarationInfo
End Function

Function IsArgCountValid(ByVal mi As MemberInfo, ByVal arg_stack As Collection) As Boolean
Print #99, "IsArgCountValid: mi="; mi.Name; " am="; mi.AttributeMask; " cc="; mi.CallConv; " dk="; mi.DescKind; " asc="; arg_stack.Count; " vto="; mi.VTableOffset; " ik="; mi.InvokeKind; " pc="; mi.parameters.Count; " pdc="; mi.parameters.defaultCount; " poc="; mi.parameters.OptionalCount
Dim asc As Integer
asc = arg_stack.Count
Dim pi As ParameterInfo
For Each pi In mi.parameters
    Dim i As Integer
    Print #99, "n="; pi.Name; " f="; pi.flags; " i="; CStr(i); " o="; pi.Optional; " d="; pi.Default; " asc="; arg_stack.Count
    If Not CBool(pi.flags And (PARAMFLAG_FLCID Or PARAMFLAG_FRETVAL)) Then
        i = i + 1
        Dim arg As Collection
        Set arg = Nothing
        On Error Resume Next
        Set arg = arg_stack.Item(CStr(i))
        If arg Is Nothing Then Set arg = arg_stack.Item(pi.Name)
        On Error GoTo 0
        If i = mi.parameters.Count And CBool(mi.InvokeKind And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) And (mi.ReturnType.varType = VT_VOID Or mi.ReturnType.varType = VT_HRESULT) Then
            If Not arg Is Nothing Then Err.Raise 1
            IsArgCountValid = True
            GoTo 10
        End If
        Print #99, "i="; i
        If arg Is Nothing Then Set arg = New Collection: arg_stack.Add arg, CStr(i)
        Print #99, "a.c="; arg.Count
        ' TODO: need documentation on why missing args need to be fleshed out.
        If arg.Count = 0 Then
            Dim token As vbToken
            Set token = New vbToken
            Print #99, "1"
            If pi.Default Then
                Print #99, "d="; varType(pi.DefaultValue)
                token.tokType = tokVariant
                Print #99, "2"
                token.tokString = "_Param" & CStr(i)
                Print #99, "3"
                On Error Resume Next
                Print #99, "iserr="; IsError(pi.DefaultValue)
                Print #99, "ise="; IsEmpty(pi.DefaultValue)
                Print #99, "isn="; IsNull(pi.DefaultValue)
                Print #99, "iso="; IsObject(pi.DefaultValue)
                Print #99, "ism="; IsMissing(pi.DefaultValue)
                Print #99, "isnothing="; pi.DefaultValue Is Nothing
                On Error GoTo 0
                ' cbCoolbar.Bands.Add
                ' oleview displays disassembled idl parameter as <unprintable IDispatch*>
                If IsObject(pi.DefaultValue) Then If pi.DefaultValue Is Nothing Then Set token.tokValue = Nothing Else Err.Raise 1 Else token.tokValue = pi.DefaultValue
                Print #99, "4"
                token.tokDataType = varType(pi.DefaultValue)
                Print #99, "5"
                arg.Add token
                Print #99, "6"
            ElseIf pi.Optional Or CBool(pi.flags And PARAMFLAG_FOPT) Then
                arg.Add MissingArg(token, pi.VarTypeInfo.varType)
            Else
                Err.Raise 1 ' non-optional parameter missing
            End If
            Print #99, "ts="; token.tokString; " ac="; arg.Count
        End If
    End If
Next
' if member is a property get, with no parameters, and return type of object then assume default member will gooble args
Dim rtdt As TliVarType
Dim rtii As InterfaceInfo
GetReturnType Nothing, mi, rtdt, rtii
'Print #99, "ik="; mi.InvokeKind; " i="; i; " vt="; mi.ReturnType.varType
Print #99, "ik="; mi.InvokeKind; " i="; i; " vt="; rtdt
'If mi.InvokeKind = INVOKE_PROPERTYGET And i = 0 And (IsObj(mi.ReturnType.varType) Or mi.ReturnType.varType = 0) Then
If mi.InvokeKind = INVOKE_PROPERTYGET And i = 0 And (IsObj(rtdt) Or rtdt = 0) Then
    IsArgCountValid = True
Else
    IsArgCountValid = arg_stack.Count = i Or (arg_stack.Count > i And mi.parameters.OptionalCount = -1)
    If Not IsArgCountValid Then
        While arg_stack.Count > asc
            arg_stack.Remove asc
        Wend
    End If
End If
10
Print #99, "IsArgCountValid: done="; IsArgCountValid; " asc="; arg_stack.Count
End Function

Function SearchForAppObject(ByVal tokens As Collection, ByVal tlib As TypeLibInfo, ByVal arg_stack As Collection, ByVal ik As InvokeKinds) As Boolean
' check for appobject
Dim token As vbToken
Set token = tokens.Item(1)
Print #99, "SearchForAppObject: ts="; token.tokString; " ik="; ik
Print #99, "SearchForAppObject: tlib="; tlib.Name; " GUID="; tlib.GUID
Dim classinfo As CoClassInfo
Dim MemberInfo As MemberInfo
' search for explicitly qualified AppObject - DBEngine.WorkSpaces(0)
For Each classinfo In tlib.CoClasses
    Print #99, "SearchForAppObject: classInfo="; classinfo.Name; " guid="; classinfo.GUID; " amask="; Hex(classinfo.AttributeMask)
    If classinfo.AttributeMask And TYPEFLAG_FAPPOBJECT Then
        Print #99, "SearchForAppObject: appobject found: "; classinfo.Name
        If UCase(token.tokString) = UCase(classinfo.Name) Then
            tokens.Remove 1 ' remove appobject classname
            If tokens.Item(1).tokString <> "." Then Err.Raise 1
            tokens.Remove 1 ' remove .
            Set token = tokens.Item(1)
            tokens.Remove 1 ' control name
            If token.tokType <> toksymbol Then Err.Raise 1
            If arg_stack.Count <> 0 Then Err.Raise 1 ' Invalid arg list
            ' new - DBEngine.OpenDataBase(...)
            Set arg_stack = CreateArgStack(token, tokens, arg_stack)
            If Not MatchMemberInfo(classinfo.DefaultInterface, token, arg_stack, MemberInfo, ik) Then Err.Raise 1 ' appobject explicit but member not found
            If tokens.Count = 0 Then tokens.Add token Else tokens.Add token, , 1
            Exit For
        End If
    End If
Next
If MemberInfo Is Nothing Then
    ' search for unqualified AppObject member - AppName
    For Each classinfo In tlib.CoClasses
        Print #99, "SearchForAppObject: classInfo="; classinfo.Name; " guid="; classinfo.GUID; " amask="; Hex(classinfo.AttributeMask)
        If classinfo.AttributeMask And TYPEFLAG_FAPPOBJECT Then
            Print #99, "SearchForAppObject: appobject found: "; classinfo.Name
            If MatchMemberInfo(classinfo.DefaultInterface, token, arg_stack, MemberInfo, ik) Then Exit For
        End If
    Next
End If
If Not MemberInfo Is Nothing Then
    Print #99, "SearchForAppObject: appobject member found n="; MemberInfo.Name
    Print #99, "SearchForAppObject: 1 cn="; classinfo.Name; " guid="; classinfo.GUID; " ii.n="; classinfo.DefaultInterface.Name
    ' fixme: This is duped - do only once
    Dim rtdt As TliVarType
    Dim rtii As InterfaceInfo
    token.tokType = tokReferenceClass
    Set token.tokReturnType = GetReturnType(classinfo.DefaultInterface, MemberInfo, rtdt, rtii)
    token.tokDataType = rtdt
    Set token.tokInterfaceInfo = classinfo.DefaultInterface
    Set token.tokMemberInfo = MemberInfo
    Set token.tokVariable = New vbVariable
    Set token.tokVariable.varModule = currentModule
    Set token.tokVariable.varProc = currentProc
    token.tokVariable.MemberType = vbext_mt_Variable
    Set token.tokVariable.varType = New vbDataType
    token.tokVariable.varType.dtType = tokReferenceClass
    Set token.tokVariable.varType.dtClassInfo = classinfo
    Set token.tokVariable.varType.dtInterfaceInfo = rtii
    AddTLIInterfaceMember token.tokInterfaceInfo, MemberInfo, token.tokReturnType, MemberInfo.InvokeKind
    SearchForAppObject = True
End If
Print #99, "SearchForAppObject: fnd="; SearchForAppObject
End Function

' Lookup an unqualified name in a forms typelib
Function FormSearch(ByVal m As vbModule, ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection, ik As InvokeKinds) As vbToken
' Form1.method
' Put this coclass form code into module variable m.coclassinfo
Print #99, "FormSearch: vb path="; tliVB.FullPath; " ts="; tokens.Item(1).tokString; " asc="; arg_stack.Count; " ik="; ik
Dim token As vbToken
Set token = tokens.Item(1)
tokens.Remove 1
Dim tlib As TypeLibInfo
Dim classinfo As CoClassInfo
Print #99, "1"
Set tlib = TypeLibInfoFromFile(tliVB.FullPath)
Print #99, "2"
If tlib Is Nothing Then Err.Raise 1 ' Typelibrary error
Print #99, "3 ct="; m.Component.Type
Select Case m.Component.Type
    Case vbext_ct_VBForm
        Set classinfo = tlib.CoClasses.NamedItem("Form")
    Case vbext_ct_VBMDIForm
        Set classinfo = tlib.CoClasses.NamedItem("MDIForm")
    Case vbext_ct_PropPage
        Set classinfo = tlib.CoClasses.NamedItem("PropertyPage") ' axmrquee\axmrquee.vbp
    Case vbext_ct_UserControl
        Set classinfo = tlib.CoClasses.NamedItem("UserControl") ' axmrquee\axmrquee.vbp
        If UCase(token.tokString) = "USERCONTROL" Then
            If tokens.Item(1).tokString <> "." Then Err.Raise 1
            tokens.Remove 1 ' .
            Set token = tokens.Item(1)
            tokens.Remove 1 ' control name
            If token.tokType <> toksymbol Then Err.Raise 1
            If arg_stack.Count <> 0 Then Err.Raise 1 ' Invalid arg list
            ' fixme: UserControl.blah()??? needs to be implemented
            Set arg_stack = CreateArgStack(token, tokens, arg_stack)
            Dim uc As Boolean
            uc = True
        End If
    Case Else
        For Each classinfo In tlib.CoClasses
            Print #99, " ci="; classinfo.Name; " guid="; classinfo.GUID
        Next
        Err.Raise 1 ' unexpected component type
End Select
Print #99, "4";
If classinfo Is Nothing Then Err.Raise 1
Print #99, "5"
Set tlib = Nothing
Print #99, "6"
Dim ii As InterfaceInfo
Set ii = classinfo.DefaultInterface
Print #99, "7"
Dim mi As MemberInfo
If MatchMemberInfo(ii, token, arg_stack, mi, ik) Then
    output_stack.Add GetForm(m)
    Set FormSearch = token
    ' following code should be in subroutine
    Dim rtdt As TliVarType
    Dim rtii As InterfaceInfo
    Set FormSearch.tokReturnType = GetReturnType(ii, mi, rtdt, rtii)
    Set FormSearch.tokVariable = New vbVariable
    Set FormSearch.tokVariable.varType = New vbDataType
    FormSearch.tokType = tokReferenceClass
    FormSearch.tokString = mi.Name
    Set FormSearch.tokInterfaceInfo = ii
    Set FormSearch.tokMemberInfo = mi
    FormSearch.tokVariable.varType.dtType = tokReferenceClass
'    Set FormSearch.tokVariable.varType.dtClass = m
    Set FormSearch.tokVariable.varType.dtInterfaceInfo = rtii
    FormSearch.tokDataType = rtdt
    AddTLIInterfaceMember ii, mi, FormSearch.tokReturnType, mi.InvokeKind
    CoerceObject output_stack, output_stack.Count, vbObject, ii, Nothing
Else
    If uc Then Err.Raise 1 ' Member not found - UserControl.xxx but xxx not found
    If tokens.Count = 0 Then tokens.Add token Else tokens.Add token, , 1
End If
Print #99, "FormSearch: done"
End Function

Sub SearchProjectClass(ByVal token As vbToken, ByVal member As vbToken, ByVal output_stack As Collection, ByRef ik As InvokeKinds)
Print #99, "SearchProjectClass: s="; token.tokString; " t="; token.tokType; " m="; member.tokString; " os="; output_stack.Count; " ik="; ik
If token.tokVariable.varType.dtClass Is Nothing Then Err.Raise 1
Print #99, "Local class="; token.tokVariable.varType.dtClass.Name
Print #99, "Local procs="; Not token.tokVariable.varType.dtClass.procs Is Nothing
Print #99, "Local procs.count=" & token.tokVariable.varType.dtClass.procs.Count
'            If Member Is LastQualifier Then ik = LastQualifier.tokPCodeSubType
Dim proctable As proctable
For Each proctable In token.tokVariable.varType.dtClass.procs
    Print #99, "pt="; proctable.procName; " ik="; proctable.InvokeKind
'    On Error Resume Next
'    Print #99, "g="; token.tokVariable.varType.dtClass.procs.Item(SymIK(proctable.procname, INVOKE_PROPERTYGET).procName
'    Print #99, "l="; token.tokVariable.varType.dtClass.procs.Item(SymIK(proctable.procname, INVOKE_PROPERTYPUT).procName
'    Print #99, "s="; token.tokVariable.varType.dtClass.procs.Item(SymIK(proctable.procname, INVOKE_PROPERTYPUTREF).procName
'    On Error GoTo 0
Next
On Error Resume Next
Set member.tokLocalFunction = token.tokVariable.varType.dtClass.procs.Item(SymIK(member.tokString, ik))
On Error GoTo 0
If member.tokLocalFunction Is Nothing Then
    Print #99, "ik="; ik
    If ik = INVOKE_PROPERTYPUTREF Then
        On Error Resume Next
        ' Is it a really a Function or Property Get with default value? Set col.Item(1) = Nothing
        Set member.tokLocalFunction = token.tokVariable.varType.dtClass.procs.Item(SymIK(member.tokString, INVOKE_FUNC Or INVOKE_PROPERTYGET))
        On Error GoTo 0
    ElseIf ik = INVOKE_PROPERTYPUT Then
        On Error Resume Next
        ' Is it a really a Function or Property Get with default value?
        Set proctable = token.tokVariable.varType.dtClass.procs.Item(SymIK(member.tokString, INVOKE_FUNC Or INVOKE_PROPERTYGET))
        On Error GoTo 0
        If proctable Is Nothing Then
            On Error Resume Next
            Set proctable = token.tokVariable.varType.dtClass.procs.Item(SymIK(member.tokString, INVOKE_PROPERTYPUTREF))
            On Error GoTo 0
            If Not proctable Is Nothing Then
                Print #99, "Missing Set keyword in Property Set assignment statement"
                MsgBox "Missing Set keyword in Property Set assignment statement"
                Err.Raise 1
            End If
        Else
            Set member.tokLocalFunction = proctable
        End If
    End If
    ' is it a public variable (INVOKE_UNKNOWN)?
    If member.tokLocalFunction Is Nothing Then
        On Error Resume Next
        Set member.tokLocalFunction = token.tokVariable.varType.dtClass.procs.Item(SymIK(member.tokString, INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF))
        On Error GoTo 0
    End If
End If
Print #99, "Local lf="; Not member.tokLocalFunction Is Nothing
If Not member.tokLocalFunction Is Nothing Then ' Is it a method?
    ' want dtClass
    Print #99, "Local ik="; ik; member.tokLocalFunction.InvokeKind; " frt="; Not member.tokLocalFunction.procFunctionResultType Is Nothing; " lm="; Not member.tokLocalFunction.procLocalModule Is Nothing
    If Not member.tokLocalFunction.procFunctionResultType Is Nothing Then
        Print #99, "Local m=" & member.tokLocalFunction.procLocalModule.Name
        Print #99, "Local lf.ik=" & member.tokLocalFunction.InvokeKind
        Print #99, "Local res=" & (Not member.tokLocalFunction.procFunctionResultType Is Nothing)
        Print #99, "Local vt=" & (Not member.tokLocalFunction.procFunctionResultType.varType Is Nothing)
        Print #99, "Local type=" & member.tokLocalFunction.procFunctionResultType.varType.dtType
        Print #99, "Local c=" & (Not member.tokLocalFunction.procFunctionResultType.varType.dtClass Is Nothing)
        Print #99, "Local ci=" & (Not member.tokLocalFunction.procFunctionResultType.varType.dtClassInfo Is Nothing)
        Print #99, "Local ii=" & (Not member.tokLocalFunction.procFunctionResultType.varType.dtInterfaceInfo Is Nothing)
        Print #99, "Local a=" & Hex(member.tokLocalFunction.procattributes)
        ' fixme: need to create "varDataType"
        member.tokDataType = member.tokLocalFunction.procFunctionResultType.varType.dtDataType Or IIf(member.tokLocalFunction.procFunctionResultType.varDimensions Is Nothing, 0, VT_ARRAY Or VT_BYREF) ' ByRef OK?
'        print #99, "local is="; Member.tokVariable Is Member.tokLocalFunction.procFunctionResultType
        Set member.tokVariable = member.tokLocalFunction.procFunctionResultType
        Print #99, "type="; member.tokVariable.varType.dtType
        Select Case member.tokVariable.varType.dtType
        Case tokProjectClass ' These should be in dtInterfaceType ENum
            member.tokPCodeSubType = member.tokLocalFunction.InvokeKind
        Case tokReferenceClass
            Set member.tokInterfaceInfo = member.tokVariable.varType.dtInterfaceInfo
            Print #99, " ik="; member.tokPCodeSubType
        Case tokIDispatchInterface
            '
        Case tokFormClass
            '
        Case 0
            ' its an intrinsic type (e.g. Long)
            Set member.tokVariable = token.tokVariable
        Case Else
            Err.Raise 1
        End Select
'        If ik = (INVOKE_FUNC Or INVOKE_PROPERTYGET) Then
'            If Member.tokVariable.MemberType = vbext_mt_Variable Then
'                ik = INVOKE_PROPERTYGET
'            End If
'        End If
'        Set Member.tokVariable.varType.dtClass = token.tokVariable.varType.dtClass
' procfunctionresulttype.vartype needs to be ????????????
    ElseIf member.tokLocalFunction.InvokeKind = INVOKE_PROPERTYPUT Or member.tokLocalFunction.InvokeKind = INVOKE_PROPERTYPUTREF Then
        ' fixme: need to create "varDataType"
        member.tokDataType = member.tokLocalFunction.procParams.Item(member.tokLocalFunction.procParams.Count).paramVariable.varType.dtDataType Or IIf(member.tokLocalFunction.procParams.Item(member.tokLocalFunction.procParams.Count).paramVariable.varDimensions Is Nothing, 0, VT_ARRAY Or VT_BYREF) ' ByRef OK?
        Set member.tokVariable = member.tokLocalFunction.procParams.Item(member.tokLocalFunction.procParams.Count).paramVariable
        Print #99, " sym="; member.tokVariable.varSymbol; " type="; member.tokVariable.varType.dtType
' Hmmm, dtType is 0 when data type is intrinisic (Long)
' fixme: need to create varType2 to hold definition
''''        If member.tokVariable.varType.dtType = 0 Then member.tokVariable.varType.dtType = tokProjectClass
'        If member.tokVariable.varType.dtType = 0 Then
'            Print #99, "SearchProjectClass: dtType is 0 for member:"; member.tokString
'            MsgBox "SearchProjectClass: dtType is 0 for member:" & member.tokString
' fixme: should fix this at source of problem, not here
''''            member.tokVariable.varType.dtType = tokProjectClass
'        End If
' if Let/Set - check if putval is ByRef - coerce will insert ByRef
'                    If Member.tokLocalFunction.procParams.Item(Member.tokLocalFunction.procParams.count).paramVariable.varAttributes And VT_BYREF Then Member.tokDataType = Member.tokDataType Or VT_BYREF
'Print #99, "Local dt="; Member.tokDataType
    Else
        Set member.tokVariable = New vbVariable
        Set member.tokVariable.varType = New vbDataType
        member.tokVariable.varType.dtType = tokProjectClass
        member.tokDataType = VT_VOID
    End If
' if (INVOKE_FUNC or INVOKE_PROPERTYGET) and INVOKE_FUNC?
    ik = ik And member.tokLocalFunction.InvokeKind
    Print #99, "type change="; token.tokVariable.varType.dtType
' This was changing dtType from FormClass to ProjectClass. floodp1r\paladin.prj
'    token.tokVariable.varType.dtType = tokProjectClass
Else ' could be a local variable or inherited member
    Dim vv As vbVariable
    For Each vv In token.tokVariable.varType.dtClass.ModuleVars
        Print #99, "SearchProjectClass: var: n="; vv.varSymbol; " v="; vv.varType.dtDataType
    Next
    Set vv = Nothing
    On Error Resume Next
    Set vv = token.tokVariable.varType.dtClass.ModuleVars.Item(UCase(member.tokString))
    On Error GoTo 0
    Print #99, "Local vv="; Not vv Is Nothing
    If vv Is Nothing Then
#If 0 Then
        ' Local member not found - must be inherited
        Select Case token.tokVariable.varType.dtClass.Component.Type
            Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_PropPage, vbext_ct_UserControl ', vbext_ct_ActiveXDesigner
                ' sort of like inheritance
'                Set member.tokVariable = FormSearch(token.tokVariable.varType.dtClass, tokens, New Collection, output_stack, ik).tokVariable
                Set member.tokVariable = GetForm(token.tokVariable.varType.dtClass).tokVariable
'                GoTo process_form
            Case Else ' no other inheritance allowed
                Print #99, "SearchProjectClass: invalid component type:"; token.tokVariable.varType.dtClass.Component.Type
                Err.Raise 1 ' Expecting a method or variable
        End Select
#End If
    Else
        Set member.tokVariable = vv
        Print #99, "Local vartype="; Not member.tokVariable.varType Is Nothing
        member.tokDataType = member.tokVariable.varType.dtDataType
        ik = ik And Not INVOKE_FUNC
    End If
End If
Print #99, "SearchProjectClass: done: Local dt=" & member.tokDataType & " st=" & member.tokPCodeSubType
End Sub

Function FindInterface(ByVal member As vbToken, ByVal arg_stack As Collection, ByRef ii As InterfaceInfo, ByVal ik As InvokeKinds, ByRef mi As MemberInfo) As Boolean
Print #99, "FindInterface: m="; member.tokString; " asc="; arg_stack.Count; " ii="; ii.Name; " am="; Hex(ii.AttributeMask)
' Checking Dispatch interface before VTable Interface to fill optional/default args
' Always check VTable because it is much more efficient than Dispatch
Dim iiv As InterfaceInfo
On Error Resume Next
Set iiv = ii.VTableInterface
On Error GoTo 0
If Not iiv Is Nothing Then
    If ii.GUID <> iiv.GUID Then
        If FindInterface(member, arg_stack, iiv, ik, mi) Then GoTo 10
    End If
End If
If Not MatchMemberInfo(ii, member, arg_stack, mi, ik) Then
' special case - some ii.mi (Label1.Caption = "") has vtbl
    Dim iivii As InterfaceInfo
    For Each iivii In ii.ImpliedInterfaces
        ' recursive
        If FindInterface(member, arg_stack, iivii, ik, mi) Then
            Print #99, "FindInterface: implied interface: "; iivii.Name; "QI("; ii.Name; ")"
            Set ii = iivii
            Exit For
        End If
    Next
End If
10
FindInterface = Not mi Is Nothing
Print #99, "FindInterface: done: "; FindInterface
End Function

Sub SearchReferenceClass(ByVal token As vbToken, ByVal member As vbToken, ByVal arg_stack As Collection, ByVal output_stack As Collection, ByVal ik As InvokeKinds, ByVal NoInsertObjDefault As Boolean)
Print #99, "SearchReferenceClass: ts="; token.tokString; " dt="; token.tokDataType; " ms="; member.tokString; " asc="; arg_stack.Count; " osc="; output_stack.Count; " ik="; ik; " oe="; NoInsertObjDefault; " ii="; Not token.tokVariable.varType.dtInterfaceInfo Is Nothing
If Not IsObj(token.tokDataType And Not VT_BYREF) Then Err.Raise 1
Print #99, "var="; Not member.tokVariable Is Nothing
If member.tokVariable Is Nothing Then Err.Raise 1
Print #99, "vartype="; Not member.tokVariable Is Nothing
If member.tokVariable.varType Is Nothing Then Err.Raise 1
Print #99, "type="; member.tokVariable.varType.dtType
Print #99, "dt="; member.tokVariable.varType.dtDataType
If member.tokVariable.varType.dtType <> 0 Or member.tokVariable.varType.dtDataType <> 0 Then Err.Raise 1
'Set Member.tokVariable = New vbVariable
'Set Member.tokVariable.varType = New vbDataType
' fixme: Use vbVariable or tokInterfaceInfo????? - eliminate one or other
On Error Resume Next
Print #99, "t="; token.tokType
'Print #99, "dtt="; token.tokReturnTypeType
Print #99, "ii="; token.tokInterfaceInfo.Name
Print #99, "ii.guid="; token.tokInterfaceInfo.GUID
Print #99, "iiv.guid="; token.tokInterfaceInfo.VTableInterface.GUID
Print #99, "ii.am="; Hex(token.tokInterfaceInfo.AttributeMask)
Print #99, "mi="; token.tokMemberInfo.Name
'Print #99, "rtii="; token.tokReturnTypeInterfaceInfo.Name
Print #99, "dtii="; token.tokVariable.varType.dtInterfaceInfo.Name
Dim ii As InterfaceInfo
Dim rtdt As TliVarType
Dim rtii As InterfaceInfo
On Error GoTo 0
#If 0 Then
If token.tokMemberInfo Is Nothing Then
    ' variable typed as TLib interface
    Print #99, "SearchReferenceClass: mi is nothing"
    Set ii = token.tokInterfaceInfo
Else
    ' previous member exists, extract return interface (could be Dispatch)
'    token.tokDataType = InitVarTypeFromVarTypeInfo(token.tokMemberInfo.ReturnType, member.tokVariable.varType)
'    Set rtii = member.tokVariable.varType.dtInterfaceInfo
    Print #99, "SearchReferenceClass: mi="; token.tokMemberInfo.Name
    GetReturnType token.tokInterfaceInfo, token.tokMemberInfo, rtdt, ii
    token.tokDataType = rtdt
End If
#Else
Set ii = token.tokVariable.varType.dtInterfaceInfo
#End If
Dim mi As MemberInfo
If Not ii Is Nothing Then
    If Not FindInterface(member, arg_stack, ii, ik, mi) And arg_stack.Count > 0 Then
        Dim t As vbToken
    '    Dim rtdt As TliVarType
    '    Dim rtii As InterfaceInfo
        Set t = InsertDefaultMember(ii, INVOKE_FUNC Or INVOKE_PROPERTYGET, output_stack, arg_stack, mi, rtdt, rtii)
        If Not t Is Nothing Then
            Print #99, "rtdt="; rtdt; " rtii="; Not rtii Is Nothing; " mi="; Not mi Is Nothing
            CoerceObject output_stack, output_stack.Count, output_stack(output_stack.Count).tokDataType, t.tokInterfaceInfo, Nothing ' force QI
            ' Form1.Controls has rtii is Nothing
            If rtdt = VT_DISPATCH And Not rtii Is Nothing Then
                output_stack.Add t
                FindInterface member, arg_stack, rtii, ik, mi
'                CoerceObject output_stack, output_stack.count,vbobject, rtii, Nothing
            End If
        End If
    End If
End If
If ii Is Nothing Then
    ProcessObjectMember token, member, output_stack
ElseIf mi Is Nothing Then
    Print #99, "SearchReferenceClass: No memberinfo for "; member.tokString; " interface="; ii.Name; " GUID="; ii.GUID; " am="; Hex(ii.AttributeMask)
    If Not ii.AttributeMask And 1 Then ' 1 is hidden attr
        If ii.AttributeMask And TYPEFLAG_FNONEXTENSIBLE Then
            Print #99, "Interface is not extensible"
            Err.Raise 1 ' Name isn't a member of interface, interface in non-extensible
        End If
        If Not ii.AttributeMask And TYPEFLAG_FDISPATCHABLE Then
            Print #99, "Member not found, interface not dispatchable: " & member.tokString
            MsgBox "Member not found, interface not dispatchable: " & member.tokString
            Err.Raise 1 ' Name isn't a member of interface, interface in non-dispatchable
        End If
    End If
    ProcessObjectMember token, member, output_stack
    CoerceObject output_stack, output_stack.Count, vbObject, Nothing, Nothing
Else
    Print #99, "ii="; ii.Name; " guid="; ii.GUID; " mi.ik="; mi.InvokeKind
    Print #99, " vto="; mi.VTableOffset; " am="; Hex(ii.AttributeMask)
'    If IsVTable Eqv (mi.VTableOffset = -1 Or CBool(ii.AttributeMask And 1)) Then Err.Raise 1
    ' fixme: This is duped - do only once
    Dim ReturnType As VarTypeInfo
    Print #99, "SearchReferenceClass: 1"
    Set member.tokReturnType = GetReturnType(ii, mi, rtdt, rtii)
    Print #99, "SearchReferenceClass: 2"
    member.tokDataType = rtdt
    Set member.tokInterfaceInfo = ii
    Set member.tokMemberInfo = mi
    member.tokVariable.varType.dtType = tokReferenceClass
    Set member.tokVariable.varType.dtInterfaceInfo = rtii
    AddTLIInterfaceMember ii, mi, member.tokReturnType, ik ' IIf(token Is LastQualifier, ik, INVOKE_FUNC Or INVOKE_PROPERTYGET)
    CoerceObject output_stack, output_stack.Count, output_stack(output_stack.Count).tokDataType, ii, Nothing
End If
Print #99, "SearchReferenceClass: done: pcst="; member.tokPCodeSubType; " dt="; member.tokDataType
End Sub

Sub ProcessObjectMember(ByVal token As vbToken, ByVal member As vbToken, ByVal output_stack As Collection)
Print #99, "ProcessObjectMember: ts="; token.tokString; " ms="; member.tokString; " osc="; output_stack.Count
member.tokType = tokIDispatchInterface
member.tokDataType = vbVariant
Set member.tokVariable = New vbVariable
Set member.tokVariable.varType = New vbDataType
member.tokVariable.varType.dtType = tokIDispatchInterface
member.tokVariable.varType.dtDataType = vbVariant
Print #99, "ProcessObjectMember: done"
End Sub

Sub SearchUserDefinedType(ByVal token As vbToken, ByVal member As vbToken, ByVal output_stack As Collection, ByVal ik As InvokeKinds, ByVal NoInsertObjDefault As Boolean)
Print #99, "SearchUserDefinedType: ts="; token.tokString; " ms="; member.tokString; " osc="; output_stack.Count; " ik="; ik; " oe="; NoInsertObjDefault; " type="; token.tokType
If token.tokVariable Is Nothing Then Err.Raise 1
If token.tokVariable.varType Is Nothing Then Err.Raise 1
Print #99, "type="; token.tokVariable.varType.dtType
Select Case token.tokVariable.varType.dtType
    Case tokProjectClass ' These should be in dtInterfaceType ENum
        Dim v As vbVariable
        On Error Resume Next
        Set v = token.tokVariable.varType.dtUDT.typeMembers.Item(UCase(member.tokString))
        On Error GoTo 0
        Print #99, "v="; Not v Is Nothing
        If v Is Nothing Then Err.Raise 1 ' Not a typemember
        Print #99, "v.n="; v.varSymbol
        Set member.tokVariable = v
        member.tokDataType = v.varType.dtDataType Or IIf(v.varDimensions Is Nothing, 0, VT_ARRAY Or VT_BYREF)
' don't change dtType,
''''        member.tokVariable.varType.dtType = tokProjectClass
    Case tokReferenceClass, tokFormClass
        Dim dt As VariantTypeConstants
        Set member.tokMemberInfo = token.tokVariable.varType.dtRecordInfo.GetMember(member.tokString)
        member.tokDataType = InitVarTypeFromVarTypeInfo(member.tokMemberInfo.ReturnType, member.tokVariable.varType)
''''        member.tokVariable.varType.dtType = tokReferenceClass
' fixme: varSymbol is seldom initialized, perhaps should eliminate?
        member.tokVariable.varSymbol = member.tokMemberInfo.Name
'        Set member.tokVariable.varType.dtTLI = token.tokReference
    Case tokIDispatchInterface
        member.tokDataType = vbVariant
    Case Else
        Print #99, "SearchUserDefinedType: Invalid type: "; token.tokVariable.varType.dtType
        Err.Raise 1 ' Unexpected qualifer
End Select
member.tokType = tokUDT
member.tokPCode = vbPCodeUDT
'member.tokPCodeSubType = ik
Print #99, "SearchUserDefinedType: done: dt="; member.tokDataType; " vt="; Not member.tokVariable Is Nothing
End Sub

' fixme: implement object to parse variable type expression instead of using tokType/dtType
Function StackVariable(ByVal token As vbToken, ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection, ByVal ikp As InvokeKinds, Optional ByVal StmtMethod As Boolean, Optional ByVal NoInsertObjDefault As Boolean, Optional ByVal ReDimFlag As Boolean) As vbToken
Dim member As vbToken
If Not LastQualifier Is Nothing Then Print #99, "StackVariable: lq="; LastQualifier.tokString
Dim ik As InvokeKinds
ik = INVOKE_FUNC Or INVOKE_PROPERTYGET
3
Print #99, "StackVariable: ts="; token.tokString; " tc="; tokens.Count; " asc="; arg_stack.Count; " ikp="; ikp; " ik="; ik; " sm="; StmtMethod; " osc="; output_stack.Count; " oe="; NoInsertObjDefault; " rdf="; ReDimFlag; " st="; token.tokPCodeSubType
If token Is Nothing Then Err.Raise 1
If token.tokPCodeSubType = 0 Then token.tokPCodeSubType = 14 ' temp until IK is implemented
If arg_stack.Count > 0 Then GoTo 8
If IsEOL(tokens) Then GoTo 10
' NoInsertObjDefault - Set Form1 = Nothing (Form1 is somehow not LastQualifier)
'    ik = IIf((token Is LastQualifier) Or NoInsertObjDefault, ikp, ik) And token.tokPCodeSubType
'    GoTo 10
'        GoTo 105 ' check for default member -> Text1 = Text1
'End If
Print #99, "StackVariable: token(1) s="; tokens.Item(1).tokString; " acs="; arg_stack.Count
If tokens.Item(1).tokType <> tokMember Then
8
'    If tokens.Item(1).tokString = "(" Then Set member = InsertDefaultMember(token)
' fixme: probably should insert "Item" token and process below. What about "?Bound" processing?
#If 0 Then
    If token.tokVariable.varAttributes And VARIABLE_CONTROLARRAY Then
        Print #99, "StackVariable: control array not implemented" ' partial implemented
        If arg_stack.Count <> 1 Then Err.Raise 1 ' Perhaps s/b > 0?
        AppendCollection arg_stack.Item(1), output_stack
        CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, vbLong
        arg_stack.Remove 1
        If token.tokCount <> 0 Then Err.Raise 1
        token.tokCount = 1
    End If
#End If
10
    Print #99, "ii="; Not token.tokInterfaceInfo Is Nothing
    If Not token.tokVariable Is Nothing Then Print #99, "type="; token.tokVariable.varType.dtType; " mt="; token.tokVariable.MemberType
'    If arg_stack.count > 0 Then If token.tokDataType = vbObject Then If token.tokVariable.varType.dtType = tokReferenceClass Then Set member = InsertDefaultMember(token)
    ProcessSubscripts token, arg_stack, output_stack
    If token.tokRank > 0 Then Set arg_stack = CreateArgStack(token, tokens, New Collection, StmtMethod, ReDimFlag) ', token.tokVariable.varType.dtType = tokIDispatchInterface)
'    ik = IIf(token Is LastQualifier, ikp, ik) And token.tokPCodeSubType
    ' CheckBox1 = 0 - needs - And token.tokDataType <> vbObject
    ik = IIf(token Is LastQualifier, ikp, ik) And token.tokPCodeSubType
'    If token.tokVariable.varAttributes And VARIABLE_CONTROLARRAY Then ik = INVOKE_FUNC
    ' fixme: need to implement invoke kind
    ' if Property Let (could be vbObject) then keep ik
    ' Label1 = "" - Label1 s/b INVOKE_FUNC or INOVKE_PROPERTYGET
'    ik = ikp And token.tokPCodeSubType
'    If member Is Nothing Then
    ' new - For Each label in colLabels - don't generate default for label
    If Not NoInsertObjDefault Or arg_stack.Count > 0 Then FunctionArgs token, tokens, arg_stack, output_stack, ik
'    Else
'        output_stack.Add token
'        If token.tokDataType And VT_BYREF Then OutputSubRef output_stack, output_stack.count
'    End If
    GoTo 1000
'        GoTo 1000 ' check for (
End If
tokens.Remove 1 ' remove .
output_stack.Add token
If token.tokDataType And VT_BYREF Then OutputSubRef output_stack, output_stack.Count
If IsEOL(tokens) Then Err.Raise 1 ' Expecting qualifier
Print #99, "StackVariable: token s="; token.tokString; " type="; token.tokType; " dt="; token.tokDataType; " st="; token.tokPCodeSubType; " var="; Not token.tokVariable Is Nothing
If token.tokVariable Is Nothing Then Err.Raise 1
If token.tokVariable.varType Is Nothing Then Err.Raise 1
Set member = tokens.Item(1)
tokens.Remove 1
If arg_stack.Count = 0 Then Set arg_stack = CreateArgStack(member, tokens, New Collection, StmtMethod, ReDimFlag, token) ', token.tokVariable.varType.dtType = tokIDispatchInterface)
Print #99, "StackVariable: m="; member.tokString; " type="; member.tokType; " pcst="; member.tokPCodeSubType; " ik="; ik
If member.tokType <> toksymbol Then Err.Raise 1 ' Invalid qualifier
token.tokPCodeSubType = IIf(token.tokPCodeSubType = 0, ik, token.tokPCodeSubType And ik)
If member Is LastQualifier Then ik = ikp
Set member.tokVariable = New vbVariable
member.tokVariable.MemberType = vbext_mt_Variable
Set member.tokVariable.varType = New vbDataType
Print #99, "StackVariable: Member s="; member.tokString; " type="; member.tokType; " dt="; member.tokDataType; " st="; member.tokPCodeSubType; " var="; member.tokVariable Is Nothing; " ii="; member.tokInterfaceInfo Is Nothing
'ProcessSubscripts member, tokens, output_stack
' kludge - for Move 1,2,3,4 - arg_stack has 4 args so don't wipe it out
'If arg_stack.count = 0 Then Set arg_stack = CreateArgStack(member, tokens, New Collection, StmtMethod) ', token.tokVariable.varType.dtType = tokIDispatchInterface)
Select Case token.tokDataType And Not (VT_BYREF Or VT_ARRAY) ' form control arrays are &h2009
Case vbVariant
    ProcessObjectMember token, member, output_stack
    member.tokType = tokIDispatchInterface
'    member.tokPCode = vbPCodeMember
    GoTo 99
Case vbObject, VT_UNKNOWN
'    CoerceOperand goptimizeflag, output_stack, output_stack.count, vbObject
'        If token.tokType = tokMember Then
'            Member.tokDataType = vbVariant
'            Set Member.tokVariable = token.tokVariable ' don't think this is needed for member
'        ElseIf Not token.tokVariable.varType.dtClass Is Nothing Then
    Print #99, "StackVariable: type="; token.tokVariable.varType.dtType
    member.tokType = token.tokVariable.varType.dtType
    Select Case member.tokType
    Case tokProjectClass ' These should be in dtInterfaceType ENum
        SearchProjectClass token, member, output_stack, ik
    Case tokReferenceClass
        If token.tokVariable.varAttributes And VARIABLE_CONTROLARRAY Then
            Select Case UCase(member.tokString)
            Case "ITEM"
                token.tokPCodeSubType = INVOKE_FUNC
                If arg_stack.Count <> 1 Then Err.Raise 1 ' wrong number of subscripts
                Set member.tokInterfaceInfo = token.tokInterfaceInfo
                Set member.tokMemberInfo = token.tokMemberInfo
                member.tokDataType = vbVariant
                member.tokVariable.varType.dtType = tokReferenceClass
                Set member.tokVariable.varType.dtInterfaceInfo = token.tokInterfaceInfo
                member.tokVariable.varType.dtDataType = vbObject
'                SearchReferenceClass token, member, arg_stack, output_stack, ik, NoInsertObjDefault
#If 1 Then
                For Each token In arg_stack.Item(1)
                    output_stack.Add token
                Next
                arg_stack.Remove 1
                member.tokCount = 1
#End If
    FunctionArgs member, tokens, arg_stack, output_stack, ik
                GoTo 98
            Case "COUNT", "LBOUND", "UBOUND"
                token.tokPCodeSubType = INVOKE_FUNC
                If arg_stack.Count <> 0 Then Err.Raise 1 ' wrong number of subscripts
                member.tokDataType = vbVariant
                GoTo 98
            Case Else
'                If arg_stack.count <> 1 Then Err.Raise 1 ' wrong number of subscripts
            SearchReferenceClass token, member, arg_stack, output_stack, ik, NoInsertObjDefault
#If 0 Then
                Set token = New vbToken
                Set token.tokInterfaceInfo = output_stack.Item(output_stack.Count).tokInterfaceInfo
                Set token.tokMemberInfo = output_stack.Item(output_stack.Count).tokMemberInfo
                token.tokDataType = vbObject
                token.tokVariable.varType.dtType = tokReferenceClass
                Set token.tokVariable.varType.dtInterfaceInfo = token.tokInterfaceInfo
                token.tokVariable.varType.dtDataType = vbObject
'                SearchReferenceClass token, member, arg_stack, output_stack, ik, NoInsertObjDefault
#If 1 Then
                For Each token In arg_stack.Item(1)
                    output_stack.Add token
                Next
                arg_stack.Remove 1
                member.tokCount = 1
#End If
    FunctionArgs token, tokens, arg_stack, output_stack, ik
                GoTo 98
#End If
            End Select
        Else
            SearchReferenceClass token, member, arg_stack, output_stack, ik, NoInsertObjDefault
        End If
    Case tokIDispatchInterface
        ProcessObjectMember token, member, output_stack
    Case tokFormClass
        ' fixme: Needs to search both TLib Form members and Module Members - sort of like inheritance
'        Set member = getLocalFunction(m, tokens, arg_stack, output_stack, ik)
'        Set member = FormSearch(m, tokens, arg_stack, output_stack, ik)
        SearchProjectClass token, member, output_stack, ik
        Print #99, "stackvariable: lf="; Not member.tokLocalFunction Is Nothing
        If member.tokLocalFunction Is Nothing Then
            member.tokType = tokReferenceClass
            Set token.tokInterfaceInfo = token.tokVariable.varType.dtInterfaceInfo
            SearchReferenceClass token, member, arg_stack, output_stack, ik, NoInsertObjDefault
        Else
            member.tokType = tokProjectClass
        End If
    Case Else
        Print #99, "StackVariable: Invalid Class type: "; token.tokVariable.varType.dtType
        Err.Raise 1 ' Unexpected qualifer
    End Select
Print #99, "m.t="; member.tokType; " t.type="; token.tokVariable.varType.dtType; " m.type="; member.tokVariable.varType.dtType
'    If member.tokType <> token.tokVariable.varType.dtType Then Err.Raise 1
''''    If member.tokVariable.varType.dtType = 0 Then Err.Raise 1
    If member.tokVariable.varType.dtType = 0 Then
        ' Module1.i - has dtType of 0 because its first class data type
        Print #99, "mt="; member.tokVariable.MemberType; " at="; Hex(member.tokVariable.varAttributes)
        If member.tokVariable.varAttributes And VARIABLE_PUTVAL Then
            FunctionArgs member, tokens, arg_stack, output_stack, ik
        Else
            ProcessSubscripts member, arg_stack, output_stack
        End If
    Else
'        member.tokPCode = vbPCodeMember
99
        Print #99, "StackVariable: p-8 asc="; arg_stack.Count; " ik="; ik
        FunctionArgs member, tokens, arg_stack, output_stack, ik
        Print #99, "StackVariable: p-9 ik="; ik
    End If
Case vbUserDefinedType
    SearchUserDefinedType token, member, output_stack, ik, NoInsertObjDefault
    ProcessSubscripts member, arg_stack, output_stack
Case Else
    If Not token Is LastQualifier Then
        Print #99, "StackVariable: Invalid DataType: "; token.tokDataType
        Err.Raise 1 ' Unexpected qualifer
    End If
End Select
98
Print #99, "4"
Set token = member
Print #99, "StackVariable: p-30"
1000
Print #99, "StackVariable: p-99"
token.tokPCodeSubType = ik
Set StackVariable = token
Print #99, "StackVariable: done: s="; token.tokString; " tc="; tokens.Count; " dt="; token.tokDataType; " st="; token.tokPCodeSubType; " t="; token.tokType; " asc="; arg_stack.Count; " v="; Not token.tokVariable Is Nothing
If arg_stack.Count > 0 Then Err.Raise 1 ' IDispatch arg_stack might exist
Print #99, "9999"
End Function

Sub ProcessSubscripts(ByVal token As vbToken, ByVal arg_stack As Collection, ByVal output_stack As Collection)
Print #99, "ProcessSubscripts: ts="; token.tokString; " asc="; arg_stack.Count; " tv="; Not token.tokVariable Is Nothing; " dt="; token.tokDataType
' note: arrays don't have to have subscripts, as in passing array args (Erase a)
' don't process functions that return arrays as arrays
If Not token.tokVariable Is Nothing And token.tokLocalFunction Is Nothing Then
    If arg_stack.Count > 0 Then
        Print #99, "r="; token.tokRank; " t="; token.tokVariable.varType.dtType; " dt="; token.tokDataType; " vt.dt="; token.tokVariable.varType.dtDataType; " vd="; Not token.tokVariable.varDimensions Is Nothing
        If Not token.tokVariable.varDimensions Is Nothing Then ' Compiled variable
            If token.tokVariable.varDimensions.Count > 0 And arg_stack.Count <> token.tokVariable.varDimensions.Count Then Err.Raise 1
            token.tokDataType = token.tokDataType And Not (VT_ARRAY Or VT_BYREF)
        ElseIf token.tokDataType And VT_VECTOR Then ' TLib variable
            Print #99, "mi="; Not token.tokMemberInfo Is Nothing
            If token.tokMemberInfo Is Nothing Then Err.Raise 1 ' internal error
            Print #99, "mi.rt="; Not token.tokMemberInfo.ReturnType Is Nothing
            If token.tokMemberInfo.ReturnType Is Nothing Then Err.Raise 1 ' internal error
            Dim bounds() As Long
            Print #99, "asc="; arg_stack.Count; " ab="; token.tokMemberInfo.ReturnType.ArrayBounds(bounds)
            If arg_stack.Count <> token.tokMemberInfo.ReturnType.ArrayBounds(bounds) Then Err.Raise 1
        ElseIf token.tokDataType And VT_ARRAY Then ' TLib Dim d()
            ' new VT_ARRAY needed for ReDimPublicType1.vbp
            Print #99, "mi="; Not token.tokMemberInfo Is Nothing
            If token.tokMemberInfo Is Nothing Then Err.Raise 1 ' internal error
            Print #99, "mi.rt="; Not token.tokMemberInfo.ReturnType Is Nothing
            If token.tokMemberInfo.ReturnType Is Nothing Then Err.Raise 1 ' internal error
        Else
            GoTo 10
        End If
        token.tokRank = arg_stack.Count
        Dim c As Collection
        For Each c In arg_stack
            Dim t As vbToken
            For Each t In c
                output_stack.Add t
            Next
            arg_stack.Remove 1
            CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, vbLong
        Next
    End If
End If
10
Print #99, "ProcessSubscripts: rank="; token.tokRank; " dt="; token.tokDataType; " asc="; arg_stack.Count
End Sub

' fixme: make getDeclare like LookUpEnumMember and LookUpConst
Function getDeclare(ByVal dcls As Collection, ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal output_stack As Collection) As vbToken
Print #99, "getDeclare: 1 dcls="; Not dcls Is Nothing; " tc="; tokens.Count
Print #99, "ts="; tokens.Item(1).tokString
If dcls Is Nothing Then
    On Error Resume Next
    Set tokens.Item(1).tokDeclare = currentModule.Declares.Item(UCase(tokens.Item(1).tokString))
    On Error GoTo 0
    Print #99, "getDeclare: 2 d="; Not tokens.Item(1).tokDeclare Is Nothing
    If tokens.Item(1).tokDeclare Is Nothing Then
        Dim module As vbModule
        For Each module In currentProject.prjModules
            Print #99, "getDeclare: 3 m="; module.Name
            If Not module Is currentModule Then
                If module.Component.Type = vbext_ct_StdModule Then
                    On Error Resume Next
                    Set tokens.Item(1).tokDeclare = module.Declares.Item(UCase(tokens.Item(1).tokString))
                    On Error GoTo 0
                    Print #99, "getDeclare: 4 d="; Not tokens.Item(1).tokDeclare Is Nothing
                    If Not tokens.Item(1).tokDeclare Is Nothing Then
                        Print #99, "getDeclare: 5 a="; tokens.Item(1).tokDeclare.dclAttributes
                        ' Declare default scope is Public
                        If tokens.Item(1).tokDeclare.dclAttributes And (PROC_ATTR_PUBLIC Or PROC_ATTR_DEFAULT) Then Exit For
                        Set tokens.Item(1).tokDeclare = Nothing
                    End If
                End If
                Print #99, "getDeclare: 6"
            End If
        Next
    End If
Else
    On Error Resume Next
    Set tokens.Item(1).tokDeclare = dcls.Item(UCase(tokens.Item(1).tokString))
    On Error GoTo 0
End If
If Not tokens.Item(1).tokDeclare Is Nothing Then
Print #99, "getDeclare: 7"
    Set getDeclare = tokens.Item(1)
    tokens.Remove 1 ' remove procedure name
    getDeclare.tokType = tokDeclare
    Set getDeclare.tokVariable = getDeclare.tokDeclare.dclFunctionResultType
' Hmmm, can Declare return an array? If so, must implement VT_ARRAY flag
    If Not getDeclare.tokDeclare.dclFunctionResultType Is Nothing Then getDeclare.tokDataType = getDeclare.tokDeclare.dclFunctionResultType.varType.dtDataType
Print #99, "getDeclare: 8 dt=" & getDeclare.tokDataType
End If
Print #99, "getDeclare: 9"
End Function

' token must be byref to match lastqualifier
Function MatchMemberInfo(ByVal ii As InterfaceInfo, ByVal token As vbToken, ByVal arg_stack As Collection, ByRef mi As MemberInfo, ByVal ik As InvokeKinds) As Boolean
Print #99, "mn="; token.tokString; " asc="; arg_stack.Count; " ik="; ik
Dim mn As String
mn = UCase(token.tokString)
On Error GoTo 10
For Each mi In ii.Members
    Print #99, "MatchMemberInfo: mi.n="; mi.Name; " mi.ik="; mi.InvokeKind; " mi.Id="; Hex(mi.MemberId); " mi.am="; Hex(mi.AttributeMask); " mi.rt="; Not mi.ReturnType Is Nothing; " mi.pi.c="; mi.parameters.Count; " mi.pi.pd.c="; mi.parameters.defaultCount; " mi.pi.po.c="; mi.parameters.OptionalCount; " mi.vto="; mi.VTableOffset
    If mn = UCase(mi.Name) Then
        Print #99, "mi.ik="; mi.InvokeKind; " mi.am="; Hex(mi.AttributeMask)
        ' (mi.AttributeMask And 1) rejects Unknown and Dispatch methods - attempting to avoid problem of GetTypeInfo appearing twice (IDispatch and _TypeLibInfo in typelib)
        ' Is there a better way? Perhaps searching ImpliedInterfaces last?
        ' INVOKE_UNKNOWN is used for dispinterfaces (default interface of Font) (although non-default IFONT is available)
        'If (mi.AttributeMask And 1) = 0 And ((token Is LastQualifier And (mi.InvokeKind = INVOKE_UNKNOWN Or (ik And mi.InvokeKind))) Or ((Not token Is LastQualifier) And (mi.InvokeKind And (INVOKE_FUNC Or INVOKE_PROPERTYGET)))) Then
    '    If mi.InvokeKind = INVOKE_UNKNOWN Or
        If (mi.AttributeMask And 1) = 0 Then
''''        If ((mi.AttributeMask And 1) = 0 And ((token Is LastQualifier And (ik And mi.InvokeKind))) Or _
''''        ((Not token Is LastQualifier) And (mi.InvokeKind And (INVOKE_FUNC Or INVOKE_PROPERTYGET)))) Then
    '        If arg_stack.count >= (mi.Parameters.count - IIf(mi.InvokeKind And (INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF), 1, 0) - mi.Parameters.OptionalCount) And arg_stack.count <= (mi.Parameters.count - IIf(mi.InvokeKind And (INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF), 1, 0)) Then
            If mi.InvokeKind And ik Then If IsArgCountValid(mi, arg_stack) Then Exit For
'                    Print #99, "match: rt.vt="; IIf(mi.ReturnType Is Nothing, "void", mi.ReturnType.varType)
' new - added INVOKE_FUNC because interface had only a invokekind of INVOKE_FUNC
            If CBool(mi.InvokeKind And (INVOKE_FUNC Or INVOKE_PROPERTYGET)) And CBool(ik And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) Then
                ' Left can either be Form.Left or Left(s,1) function
                Dim miGet As MemberInfo
                Set miGet = mi
            End If
#If 0 Then
            Else
    '            If IsObj(mi.ReturnType.varType) Then Set token = InsertDefaultMember(token)
                If token Is Nothing Then
                    Print #99, "MatchMemberInfo: arg count mismatch: "; mi.Name
                    MsgBox "MatchMemberInfo: arg count mismatch: " & mi.Name
                Else
                    Set mi = token.tokMemberInfo
                End If
            End If
#End If
        End If
    End If
    If False Then
10      Print #99, "Bad typelib member:"; ii.Parent.Name; "."; ii.Name; "."; mi.Name
        Resume 20
    End If
20
Next
If mi Is Nothing And Not miGet Is Nothing Then Set mi = miGet
MatchMemberInfo = Not mi Is Nothing
Print #99, "MatchMemberInfo: fnd="; MatchMemberInfo
End Function

Sub ArgStackRPN(ByVal tokens As Collection, ByVal arg_stack As Collection, ByVal arg_stackTo As Collection, ByRef named As Boolean, ByRef HasTo As Boolean)
Dim temp_stack As New Collection
Dim temp_stackTo As New Collection
Dim arg_id As String
If tokens.Item(1).tokType = tokNamedParameter Then
    arg_id = UCase(tokens.Item(1).tokString)
    tokens.Remove 1 ' remove named parameter
    named = True
End If
If named Then
    If arg_id = "" Then Err.Raise 1 ' Expecting named argument
Else
    arg_id = CStr(arg_stack.Count + 1)
End If
If tokens.Count > 1 Then
    Dim kw As Keywords
    Dim token As vbToken
    ' ByRef in a Call statement doesn't seem to work in VB6, but is documented. So let RPNize error out.
    ' ByVal is used by Declare - ok for Call?
    kw = getKeyword(tokens.Item(1))
    If kw = KW_BYVAL Then
        Dim byV As Boolean
        byV = True
        Set token = tokens.Item(1)
        tokens.Remove 1 ' remove ByVal
        oRPN.RPNize gOptimizeFlag, tokens, temp_stack, VT_VOID ' VT_VOID specifies no coercion
        token.tokType = tokByVal
        token.tokDataType = temp_stack.Item(temp_stack.Count).tokDataType
        temp_stack.Add token
        GoTo 10
    ElseIf kw = KW_ADDRESSOF Then
        Dim addrOf As Boolean
        addrOf = True
        tokens.Remove 1 ' remove AddressOf
        Set token = tokens.Item(1)
        tokens.Remove 1 ' remove std procedure name
        If token.tokType <> toksymbol Then Err.Raise 1
        ' this lookup should be a proc
        Dim s As String
        s = UCase(token.tokString)
        Dim m As vbModule
        For Each m In currentProject.prjModules
            If m.Component.Type = vbext_ct_StdModule Then
                Dim p As proctable
                For Each p In m.procs
                    If s = UCase(p.procName) Then
                        token.tokType = tokStdProcedure
                        Set token.tokModule = m
                        Set token.tokLocalFunction = p
                        token.tokDataType = vbLong
                        temp_stack.Add token
                        GoTo 10
                    End If
                Next
            End If
        Next
        Err.Raise 1 ' can't find std module method name
    End If
End If
oRPN.RPNize gOptimizeFlag, tokens, temp_stack, VT_VOID
If Not IsEOL(tokens) Then
    If UCase(tokens.Item(1).tokString) = "TO" Then
        If named Or byV Or addrOf Then Err.Raise 1
        tokens.Remove 1 ' remove To
        oRPN.RPNize gOptimizeFlag, tokens, temp_stackTo, VT_VOID ' VT_VOID specifies no coercion
        HasTo = True
    End If
End If
10
Print #99, "ArgStackRPN: arg_id="; arg_id
arg_stack.Add temp_stack, arg_id ' fixme: need to check for duplicate Named parameter
arg_stackTo.Add temp_stackTo ' for dimension statements
End Sub

Function CreateArgStack(ByVal token As vbToken, ByVal tokens As Collection, ByVal arg_stack As Collection, Optional ByVal StmtMethod As Boolean, Optional ReDimFlag As Boolean, Optional ByVal ParentToken As vbToken) As Collection ', Optional ByVal DispatchType As Boolean) As Collection
Dim named As Boolean
Dim HasTo As Boolean
Dim IsFormObject As Boolean
Print #99, "CreateArgStack: ts="; token.tokString; " type="; token.tokType; " tc="; tokens.Count; " tv="; Not token.tokVariable Is Nothing; " asc="; arg_stack.Count; " sm="; StmtMethod; " rdf="; ReDimFlag; " pt="; Not ParentToken Is Nothing
' line2.prj
If ParentToken Is Nothing Then
    If IsForm(currentModule.Component.Type) Then
        IsFormObject = True
    Else
        If Not token.tokVariable Is Nothing Then IsFormObject = token.tokVariable.varType.dtType = tokFormClass
    End If
Else
    Print #99, "pii="; Not ParentToken.tokVariable.varType.dtInterfaceInfo Is Nothing
    If Not ParentToken.tokVariable.varType.dtInterfaceInfo Is Nothing Then
        Print #99, "ii="; ParentToken.tokVariable.varType.dtInterfaceInfo.Name
        ' fixme: Should parent library name be verified?
        ' fixme: Should upper-case be forced and tested?
        ' fixme: Too odd, how does VB know when to parse Line, Circle, etc?
        If IsAny(ParentToken.tokVariable.varType.dtInterfaceInfo.Name, "_Form", "_PictureBox", "_Printer", "_PropertyPage", "_UserControl", "_UserDocument") Then IsFormObject = True
    End If
End If
Print #99, "ifo="; IsFormObject
'If Not IsFormObject Or Not IsForm(currentModule.Component.Type) Then GoTo expr
Select Case UCase(token.tokString)
Case "CIRCLE" ' special case - any object using CIRCLE keyword
    If Not IsFormObject Then GoTo expr
    parseCircleExpression tokens, arg_stack
    token.tokPCode = vbpcodeCircle
    If ParentToken Is Nothing Then
        If IsFormObject Then Set token.tokReference = tliVB
    Else
        If Not ParentToken.tokVariable.varType.dtInterfaceInfo Is Nothing Then Set token.tokReference = tliVB
    End If
Case "INPUT", "INPUTB" ' ??? special case - any object using Input keyword
    ' only if Input is unqualified (not preceded by .)
    If Not ParentToken Is Nothing Then GoTo expr
    parseInputExpression tokens, arg_stack
Case "LINE" ' special case - any object using LINE keyword
    If Not StmtMethod Then GoTo expr ' allow o.line = o.line in any type of module
    parseLineExpression tokens, arg_stack
    token.tokPCode = vbpcodeline
    If ParentToken Is Nothing Then
        If IsFormObject Then Set token.tokReference = tliVB
    Else
        If Not ParentToken.tokVariable.varType.dtInterfaceInfo Is Nothing Then Set token.tokReference = tliVB
    End If
Case "PRINT"
    If Not IsFormObject Then
        ' Can't be used as method (what would v.Print 1;2 pass?)
        Print #99, "Invalid method name. Emitter may fail: "; token.tokString
        MsgBox "Invalid method name. Emitter may fail: " & token.tokString
        ' err.raise 1
    End If
    ' do nothing
Case "PSET" ' special case - any object using PSET keyword
    If Not IsFormObject Then GoTo expr
    parsePSetExpression tokens, arg_stack
    token.tokPCode = vbpcodepset
    If ParentToken Is Nothing Then
        If IsFormObject Then Set token.tokReference = tliVB
    Else
        If Not ParentToken.tokVariable.varType.dtInterfaceInfo Is Nothing Then Set token.tokReference = tliVB
    End If
Case "SCALE" ' special case - any object using SCALE keyword
    If Not IsFormObject Then GoTo expr
    parseScaleExpression tokens, arg_stack
    token.tokPCode = vbpcodescale
    If ParentToken Is Nothing Then
        If IsFormObject Then Set token.tokReference = tliVB
    Else
        If Not ParentToken.tokVariable.varType.dtInterfaceInfo Is Nothing Then Set token.tokReference = tliVB
    End If
'Case "WRITE" hmmm, does Write only use special syntax if unqualified? bigtol1g\project1.vbp
'    ' do nothing
Case Else
expr:
Print #99, "expr"
    Dim arg_stackTo As New Collection
    If tokens.Count > 1 Then
        If StmtMethod And token Is LastQualifier Then ' process statement style args
            ' StmtMethod or not, it is possible for tokens.count to be zero,
            ' because arg_stack isn't allowed to be Nothing. Seems safe to leave as is.
            ' fixme: Is this obsolete?
            If tokens.Item(1).tokString = "(" Then ' could be : or Else
                tokens.Remove 1 ' remove (
                If Not IsEOL(tokens) Then
                    If tokens.Item(1).tokString <> ")" Then
                        Do
                            Print #99, "CreateArgStack: p-1"
                            ArgStackRPN tokens, arg_stack, arg_stackTo, named, HasTo
                            Print #99, "CreateArgStack: p-2"
                        Loop While OptionalComma(tokens)
                        If tokens.Item(1).tokString <> ")" Then Err.Raise 1
                    End If
                    tokens.Remove 1 ' remove )
                    If HasTo Then Err.Raise 1 ' To keyword not allowed
                End If
            End If
        Else
            named = False
            Print #99, "CreateArgStack: ts="; tokens.Item(1).tokString
            If tokens.Item(1).tokString = "(" Then ' Exit Do
                If tokens.Item(2).tokString = ")" Then
                    tokens.Remove 1 ' (
                    tokens.Remove 1 ' )
                Else
                    Do Until IsEOL(tokens)
                        Print #99, "CreateArgStack: p-3"
                        tokens.Remove 1 ' remove ( or ,
                        Print #99, "CreateArgStack: p-4"
                        ArgStackRPN tokens, arg_stack, arg_stackTo, named, HasTo
                        Print #99, "CreateArgStack: p-5"
                        If IsEOL(tokens) Then Err.Raise 1
                        Print #99, "CreateArgStack: p-6"
                        If tokens.Item(1).tokString = ")" Then Exit Do
                        Print #99, "CreateArgStack: p-7"
                        If tokens.Item(1).tokString <> "," Then Err.Raise 1
                        Print #99, "CreateArgStack: p-8"
                    Loop
                    tokens.Remove 1 ' remove )
                End If
                If ReDimFlag Then
                    If IsEOL(tokens) Then
                        ArgDimensions arg_stack, arg_stackTo
                    Else
                        If tokens.Item(1).tokString = "." Then
                            If HasTo Then Err.Raise 1 ' To keyword not allowed
                        Else
                            ArgDimensions arg_stack, arg_stackTo
                        End If
                    End If
                Else
                    If HasTo Then Err.Raise 1 ' To keyword not allowed
                End If
            End If
        End If
    End If
End Select
Set CreateArgStack = arg_stack
Print #99, "CreateArgStack: asc="; arg_stack.Count
End Function

Sub ArgDimensions(ByVal arg_stack As Collection, ByVal arg_stackTo As Collection)
Print #99, "ArgDimensions: as.c="; arg_stack.Count; " asTo.c="; arg_stackTo.Count
If arg_stack.Count = 0 Then Err.Raise 1 ' Expecting dimensions
If arg_stack.Count <> arg_stackTo.Count Then Err.Raise 1
Dim c As Collection
Dim cTo As Collection
Dim i As Long
Dim token As vbToken
For i = 1 To arg_stack.Count
    Set c = arg_stack.Item(i)
    Set cTo = arg_stackTo.Item(i)
    Print #99, "i="; i; " c.c="; c.Count; " cTo.c="; cTo.Count
    Dim newC As Collection
    Set newC = New Collection
    If cTo.Count = 0 Then
        Set token = New vbToken
        token.tokString = "0"
        token.tokType = tokVariant
        token.tokValue = 0&
        token.tokDataType = vbLong
        newC.Add token
        For Each token In c
            newC.Add token
            c.Remove 1
        Next
    Else
        For Each token In c
            newC.Add token
            c.Remove 1
        Next
        For Each token In cTo
            newC.Add token
            cTo.Remove 1
        Next
    End If
    Print #99, "newC.c="; newC.Count
    For Each token In newC
        c.Add token
    Next
    For Each token In c
        Print #99, "t="; token.tokType; " s="; token.tokString
    Next
Next
End Sub

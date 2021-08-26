Attribute VB_Name = "vbtDataType"
Option Explicit

Function getOptionalByAsDataType(ByVal tokens As Collection, Optional ByVal AnyAllowed As Boolean) As paramTable
Dim kw As Keywords
Print #99, "getOptionalByAsDataType: 1 s="; tokens.Item(1).tokString
    kw = getKeyword(tokens.Item(1))
Print #99, "getOptionalByAsDataType: kw="; kw
    If kw = KW_OPTIONAL Then
        tokens.Remove 1
Print #99, "getOptionalByAsDataType: 2 s="; tokens.Item(1).tokString
        Set getOptionalByAsDataType = getByAsDataType(tokens, AnyAllowed)
Print #99, "getOptionalByAsDataType: 3 s="; tokens.Item(1).tokString
        getOptionalByAsDataType.paramVariable.varAttributes = getOptionalByAsDataType.paramVariable.varAttributes Or VARIABLE_OPTIONAL
Print #99, "4"
        Dim output_stack As New Collection
        If Not IsEOL(tokens) Then
Print #99, "5"
            If tokens.Item(1).tokString = "=" Then
Print #99, "6"
                tokens.Remove 1
Print #99, "7"
                oRPN.ConstantRPNize OptimizeConstantExpressions, tokens, output_stack, getOptionalByAsDataType.paramVariable.varType.dtDataType
Print #99, "8"
                getOptionalByAsDataType.paramVariable.varVariant = output_stack.Item(1).tokValue
                getOptionalByAsDataType.paramVariable.varAttributes = getOptionalByAsDataType.paramVariable.varAttributes Or VARIABLE_DEFAULTVALUE
Print #99, "10"
            End If
        End If
Print #99, "11"
    ElseIf kw = KW_PARAMARRAY Then ' ByRef or ByVal not allowed
        tokens.Remove 1
        Dim pt As New paramTable
        Dim token As vbToken
        If tokens.Item(1).tokType <> toksymbol Then Err.Raise 1 ' Expecting parameter name
        Set token = tokens.Item(1)
        Set pt.paramVariable = New vbVariable
        pt.paramVariable.MemberType = vbext_mt_Variable
        pt.paramVariable.varSymbol = tokens.Item(1).tokString
        Set pt.paramVariable.varModule = currentModule
        Set pt.paramVariable.varProc = currentProc
        tokens.Remove 1
        getAsDataType token, tokens, pt.paramVariable, , AnyAllowed
        If pt.paramVariable.varType.dtDataType <> vbVariant Then Err.Raise 1 ' ParamArray must use Variant data type
        If Not CBool(pt.paramVariable.varAttributes Or vbArray) Then Err.Raise 1  ' ParamArray variable must be an array - use ()
        pt.paramVariable.varAttributes = pt.paramVariable.varAttributes Or (VARIABLE_OPTIONAL Or VARIABLE_PARAMARRAY)
        Set getOptionalByAsDataType = pt
    Else
Print #99, "12"
        Set getOptionalByAsDataType = getByAsDataType(tokens, AnyAllowed)
Print #99, "13"
    End If
Print #99, "14"
End Function

Function getByAsDataType(ByVal tokens As Collection, Optional ByVal AnyAllowed As Boolean) As paramTable
Dim byType As Integer
Dim pt As New paramTable
Dim kw As Keywords
Dim token As vbToken
    kw = getKeyword(tokens.Item(1))
    If kw = KW_EMPTY Or kw = KW_BYREF Then
        byType = VT_BYREF
    ElseIf kw = KW_BYVAL Then
        byType = 0
    End If
    If kw <> KW_EMPTY Then tokens.Remove 1
    If tokens.Item(1).tokType <> toksymbol Then Err.Raise 1 ' Expecting parameter name
Print #99, "getByAsDataType 5 ts="; tokens.Item(1).tokString
    Set token = tokens.Item(1)
    Set pt.paramVariable = New vbVariable
    pt.paramVariable.MemberType = vbext_mt_Variable
    pt.paramVariable.varSymbol = tokens.Item(1).tokString
    Set pt.paramVariable.varModule = currentModule
    Set pt.paramVariable.varProc = currentProc
    tokens.Remove 1
    If getSpecialTypes(tokens.Item(1)) = SPECIAL_OP Then
Print #99, "getByAsDataType 6"
        tokens.Remove 1
Print #99, "getByAsDataType 7"
        If getSpecialTypes(tokens.Item(1)) <> SPECIAL_CP Then Err.Raise 1
Print #99, "getByAsDataType 8"
        tokens.Remove 1
Print #99, "getByAsDataType 9 ts="; tokens.Item(1).tokString
        Set pt.paramVariable.varDimensions = New Collection
Print #99, "getByAsDataType 11"
    End If
Print #99, "getByAsDataType 12"
    getAsDataType token, tokens, pt.paramVariable, , AnyAllowed
Print #99, "getByAsDataType 13"
    If pt.paramVariable.varType.dtDataType = VT_VOID And Not CBool(byType And VT_BYREF) Then
        Print #99, "warning: ByVal used with Any"
        byType = byType Or VT_BYREF
    End If
Print #99, "getByAsDataType 14"
'    If Not pt.paramVariable.varDimensions Is Nothing Then pt.paramVariable.varType.dtDataType = pt.paramVariable.varType.dtDataType Or VT_ARRAY
' fixme: don't use VT_BYREF if array - its implied -- or should it always be used for arrays?
'    If pt.paramVariable.varDimensions Is Nothing Then pt.paramVariable.varAttributes = pt.paramVariable.varAttributes Or byType
' Hmmm, arrays are always VT_BYREF
    If Not pt.paramVariable.varDimensions Is Nothing Then byType = byType Or VT_BYREF
    pt.paramVariable.varAttributes = pt.paramVariable.varAttributes Or byType
Print #99, "getByAsDataType 15"
    Set getByAsDataType = pt
Print #99, "getByAsDataType 16"
End Function

Sub getAsDataType(ByVal token As vbToken, ByVal tokens As Collection, ByVal variable As vbVariable, Optional ByVal AsNew As Boolean, Optional ByVal AnyAllowed As Boolean)
Dim dt As vbDataType
Dim kw As Keywords
Dim v As vbVariable
' syntax could be: a() or a() as variant or Function a() as udt()
Print #99, "getAsDataType: 1 s="; token.tokString
If Not IsEOL(tokens) Then
    Print #99, "getAsDataType: 2a ts="; tokens.Item(1).tokString
    If tokens.Item(1).tokString = "(" Then
        Print #99, "getAsDataType: 2b"
        If AsNew Then Err.Raise 1
        tokens.Remove 1
        If IsEOL(tokens) Then Err.Raise 1 ' Missing )
        If tokens.Item(1).tokString <> ")" Then Err.Raise 1 ' Expecting )
        tokens.Remove 1
        Set variable.varDimensions = New Collection
'        dt.dtDataType = dt.dtDataType Or VT_ARRAY
    End If
    Print #99, "getAsDataType: 2c"
    If Not IsEOL(tokens) Then
        Print #99, "getAsDataType: 2d"
        kw = getKeyword(tokens.Item(1))
        If kw = KW_AS Then
            tokens.Remove 1
            kw = getKeyword(tokens.Item(1))
            If kw = KW_NEW Then
                If Not AsNew Then Err.Raise 1 ' New not allowed
                tokens.Remove 1
                variable.varAttributes = variable.varAttributes Or VARIABLE_NEW
            End If
            Set dt = getDataType(tokens, AnyAllowed)
            ' Need routine which compares two varTypes for equality
            If Not variable.varType Is Nothing Then If dt.dtDataType <> variable.varType.dtDataType Then Err.Raise 1 ' Previously declared variable has different type
        End If
    End If
End If
Print #99, "getAsDataType: 3"
If dt Is Nothing Then
    If variable.varType Is Nothing Then
Print #99, "getAsDataType: 4"
        Set dt = New vbDataType
        If token.tokDataType = 0 Then
Print #99, "getAsDataType: 5 s="; token.tokString
            dt.dtDataType = GetDefaultType(token.tokString)
Print #99, "getAsDataType: 6"
        Else
            dt.dtDataType = token.tokDataType
        End If
    Else ' if ReDim variable doesn't specify As, use existing type
        Set dt = variable.varType
        GoTo 10
    End If
End If
Print #99, "getAsDataType: 7 dt="; dt.dtDataType
If variable.varType Is Nothing Then
Print #99, "getAsDataType: 8"
    Set variable.varType = dt
Else
Print #99, "getAsDataType: 9"
Print #99, "getAsDataType: 10"
    If dt.dtDataType = vbObject Then
Print #99, "getAsDataType: 11"
        If Not dt.dtClass Is Nothing Then
            If dt.dtClass.GUID <> variable.varType.dtClass.GUID Then Err.Raise 1 ' Previously declared variable has different type
        ElseIf Not dt.dtClassInfo Is Nothing Then
            If dt.dtClassInfo.GUID <> variable.varType.dtClassInfo.GUID Then Err.Raise 1 ' Previously declared variable has different type
        Else
            Err.Raise 1 ' Missing class info
        End If
Print #99, "getAsDataType: 12"
    ElseIf dt.dtDataType = vbUserDefinedType Then
Print #99, "getAsDataType: 13"
        If Not dt.dtUDT Is variable.varType.dtUDT Then Err.Raise 1 ' Previously declared variable has different type
Print #99, "getAsDataType: 15"
    End If
End If
10
Print #99, "getAsDataType: dt="; dt.dtDataType
End Sub

Function IsCurrentProjectName(ByVal tokens As Collection) As Boolean
Print #99, "IsCurrentProjectName: ts="; tokens.Item(1).tokString
If UCase(tokens.Item(1).tokString) = UCase(currentProject.prjName) Then
    tokens.Remove 1 ' remove project name
    If IsEOL(tokens) Then Err.Raise 1 ' Missing .
    If tokens.Item(1).tokType <> tokMember Then Err.Raise 1 ' Expecting .
    tokens.Remove 1 ' remove period
    If IsEOL(tokens) Then Err.Raise 1 ' Missing class name
    If tokens.Item(1).tokType <> toksymbol Then Err.Raise 1 ' expecting module name
    IsCurrentProjectName = True
End If
Print #99, "IsCurrentProjectName: fnd="; IsCurrentProjectName
End Function

Function GetProjectClass(ByVal tokens As Collection) As vbModule
If IsEOL(tokens) Then
    Print #99, "Expecting project or class name"
    MsgBox "Expecting project or class name"
    Err.Raise 1
End If
Print #99, "GetProjectClass: s=" & tokens.Item(1).tokString & " p=" & currentProject.prjName & " c=" & currentProject.prjModules.count
' Process Project Name - only need to process current project
If IsCurrentProjectName(tokens) Then
    On Error Resume Next
    Set GetProjectClass = currentProject.prjModules.Item(UCase(tokens.Item(1).tokString))
    On Error GoTo 0
    ' fixme: If GetProjectClass.Component.Type <> stdmodule,form,global (appobject) class? Then Err.Raise 1 ' current module name not allowed
Else
' Process Module Name
    On Error Resume Next
    Set GetProjectClass = currentProject.prjModules.Item(UCase(tokens.Item(1).tokString))
    On Error GoTo 0
End If
Print #99, "GetProjectClass: fnd="; Not GetProjectClass Is Nothing
End Function

Function getTLib(ByVal LibName As String) As reference
Print #99, "getTLib: LibName=" & LibName
On Error Resume Next
Set getTLib = currentProject.VBProject.References.Item(LibName)
On Error GoTo 0
Print #99, "getTLib: fnd="; Not getTLib Is Nothing
End Function

Function TypeInfoFromReference(ByVal reference As reference, ByVal typeName As String, ByRef TypeInfo As TypeInfo) As Boolean
Print #99, "TypeInfoFromReference: ref="; reference Is Nothing; " TypeName="; typeName
Dim tlib As TypeLibInfo
Dim s As String
Set TypeInfo = Nothing
If reference Is Nothing Then
    For Each reference In currentProject.VBProject.References
        ' need to optimize TypeLibInfoFromFile stuff
        Print #99, "TypeInfoFromReference: 1 reference="; Not reference Is Nothing
        Print #99, "TypeInfoFromReference: IsBroken="; reference.IsBroken
        Print #99, "TypeInfoFromReference: n="; reference.Name; " guid="; reference.GUID
        s = ""
        On Error Resume Next ' Duwamish Phase 3 ref errored on FullPath until .DLL was recompiled
        s = reference.FullPath
        On Error GoTo 0
        If reference.IsBroken Or s = "" Then
            Print #99, "Invalid Reference filename. Check for Project->Reference->"; reference.Name
            MsgBox "Invalid Reference filename. Check for Project->Reference->" & reference.Name
            Err.Raise 1 ' Typelibrary error
        End If
        Print #99, "TypeInfoFromReference: path="; reference.FullPath
        Set tlib = Nothing
        On Error Resume Next
        Set tlib = TypeLibInfoFromFile(reference.FullPath)
        On Error GoTo 0
        If tlib Is Nothing Then
            Print #99, "Can't open file:"; reference.FullPath; " err="; Err.Description
            MsgBox "Can't open file:" & reference.FullPath & vbCrLf & Err.Description
            Err.Raise 1 ' Typelibrary error
        End If
        Print #99, "tlib="; Not tlib Is Nothing
        Print #99, "tlib.n="; tlib.Name
        Set TypeInfo = tlib.TypeInfos.NamedItem(typeName)
Print #99, "TypeInfoFromReference: ti="; Not TypeInfo Is Nothing
        If Not TypeInfo Is Nothing Then Exit For
Print #99, "TypeInfoFromReference: 3"
    Next
Print #99, "TypeInfoFromReference: 4"
Else
    ' fixme: use above error handling code. Put in procedure.
    On Error Resume Next
    s = reference.Name
    On Error GoTo 0
    If reference.IsBroken Or s = "" Then ' IsBroken doesn't catch all cases - try using s = ""
        Print #99, "Broken typeLib reference: "; reference.Name
        MsgBox "Broken typelib reference: " & reference.Name
        Err.Raise 1
    End If
    Print #99, "TypeInfoFromReference: path="; reference.FullPath
    Set tlib = TypeLibInfoFromFile(reference.FullPath)
    If tlib Is Nothing Then Err.Raise 1 ' Typelibrary error
Print #99, "TypeInfoFromReference: 5"
        Set TypeInfo = tlib.TypeInfos.NamedItem(typeName)
Print #99, "TypeInfoFromReference: 6 "; Not TypeInfo Is Nothing
End If
TypeInfoFromReference = Not TypeInfo Is Nothing
Print #99, "TypeInfoFromReference: fnd="; TypeInfoFromReference
End Function

' get data type (could be defined in either project or TLib)
Function getProjectTLibDataType(ByVal tokens As Collection) As vbDataType
Print #99, "getProjectTLibDataType: s="; tokens.Item(1).tokString
Set getProjectTLibDataType = New vbDataType
Set getProjectTLibDataType.dtClass = GetProjectClass(tokens)
Print #99, "getProjectTLibDataType.dtClass="; Not getProjectTLibDataType.dtClass Is Nothing
If getProjectTLibDataType.dtClass Is Nothing Then
    Dim reference As reference
    Set reference = getTLib(tokens.Item(1).tokString)
    If Not reference Is Nothing Then
        tokens.Remove 1 ' remove TLib reference
        If IsEOL(tokens) Then Err.Raise 1 ' Missing .
        If tokens.Item(1).tokType <> tokMember Then Err.Raise 1 ' Expecting .
        tokens.Remove 1 ' remove .
        If IsEOL(tokens) Then Err.Raise 1 ' Missing class name
        If tokens.Item(1).tokType <> toksymbol Then Err.Raise 1 ' expecting class name
    End If
    ' fixme: Ugh, did dtClassInfo,etc all wrong. Should be dtTypeInfo.
    Dim ti As TypeInfo
    Dim typeName As String
    Dim token As vbToken
    Set token = tokens.Item(1)
    tokens.Remove 1 ' remove class name
    typeName = UCase(token.tokString)
    ' Enhancement opportunity. VB could support interface specification (TLib.TypeName.InterfaceName)
    '    instead VB can only use the default interface. Not sure why.
    If Not IsEOL(tokens) Then If tokens.Item(1).tokType = tokMember Then Err.Raise 1 ' VB doesn't support interface name specification
    If typeName = "OBJECT" Then
        getProjectTLibDataType.dtType = tokIDispatchInterface
        getProjectTLibDataType.dtDataType = vbObject
    ' Hmmm, won't directly return values to getProjectTLibDataType.dtclassinfo and getProjectTLibDataType.dtinterfaceinfo when ByRef, use intermediate variables
    ElseIf TypeInfoFromReference(reference, typeName, ti) Then
        getProjectTLibDataType.dtType = tokReferenceClass
        Set getProjectTLibDataType.dtTypeInfo = ti
'        Do
redo:
            Print #99, "getProjectTLibDataType: ti.n="; ti.Name; " guid="; ti.GUID; " tk="; ti.TypeKind; " am="; Hex(ti.AttributeMask)
            Select Case ti.TypeKind
                Case TKIND_COCLASS
                    Set getProjectTLibDataType.dtClassInfo = ti
                    Set getProjectTLibDataType.dtInterfaceInfo = ti.DefaultInterface
                    ' fixme: other TKIND_COCLASS should determine datatype as follows
                    Print #99, "ii.tk="; getProjectTLibDataType.dtInterfaceInfo.TypeKind; " am="; Hex(getProjectTLibDataType.dtInterfaceInfo.AttributeMask); " ii.c="; getProjectTLibDataType.dtInterfaceInfo.ImpliedInterfaces.count
                    If getProjectTLibDataType.dtInterfaceInfo.TypeKind = TKIND_DISPATCH Then
                        getProjectTLibDataType.dtDataType = vbObject
                    Else
                        Dim ii As InterfaceInfo
                        ' Can't trust AttributeMask flags (VB6 _CheckBox interface) exampl_1\checkbox.vbp
                        ' fixme: use FindInterface instead? Need recursion for implied interfaces?
                        ' fixme: maybe should create IsDispatch(ii)?
                        For Each ii In getProjectTLibDataType.dtInterfaceInfo.ImpliedInterfaces
                            Print #99, "  ii.ii.n="; ii.Name; ii.GUID
                            If ii.GUID = DispatchInterfaceInfo.GUID Then Exit For
                        Next
                        getProjectTLibDataType.dtDataType = IIf(ii Is Nothing, VT_UNKNOWN, vbObject)
                    End If
                Case TKIND_DISPATCH
                    Set getProjectTLibDataType.dtInterfaceInfo = ti
                    getProjectTLibDataType.dtDataType = vbObject
                Case TKIND_INTERFACE
                    Set getProjectTLibDataType.dtInterfaceInfo = ti
                    getProjectTLibDataType.dtDataType = VT_UNKNOWN
                Case TKIND_ENUM
                    Set getProjectTLibDataType.dtConstantInfo = ti
                    getProjectTLibDataType.dtDataType = vbLong
                Case TKIND_RECORD
                    Set getProjectTLibDataType.dtRecordInfo = ti
                    getProjectTLibDataType.dtDataType = vbUserDefinedType
                Case TKIND_ALIAS ' typedef
                    Print #99, "Alias found: "; ti.ResolvedType.varType; " isexternal="; ti.ResolvedType.IsExternalType
                    ' Alias not fully supported
                    If ti.ResolvedType.varType = 0 Then Set ti = ti.ResolvedType.TypeInfo: GoTo redo
                    getProjectTLibDataType.dtDataType = ConvertVarTypeToVBType(ti.ResolvedType.varType)
                Case Else
                    Print #99, "Invalid ClassInfo TypeKind: "; ti.TypeKind
                    MsgBox "Expecting ClassInfo TypeKind: " & ti.TypeKind
                    Err.Raise 1
            End Select
'        While ti.TypeKind = TKIND_ALIAS
    Else
        Set getProjectTLibDataType = Nothing
        ' fixme: if explicit ref is made, then err if not found
        If tokens.count = 0 Then tokens.Add token Else tokens.Add token, , 1
    End If
Else
    Select Case getProjectTLibDataType.dtClass.Component.Type
        Case vbext_ct_StdModule, vbext_ct_ResFile, vbext_ct_RelatedDocument
            Print #99, "Invalid class name: "; getProjectTLibDataType.dtClass.Name
            MsgBox "Expecting class name: " & getProjectTLibDataType.dtClass.Name
            Err.Raise 1
        ' fixme: Create IsForm(ct) method
        Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_PropPage, vbext_ct_UserControl ', vbext_ct_ActiveXDesigner
            ' fixme: Create routine from GetForm to only return inited tokVariable
            Set getProjectTLibDataType = GetForm(getProjectTLibDataType.dtClass).tokVariable.varType
        Case Else
            getProjectTLibDataType.dtType = tokProjectClass
    End Select
    tokens.Remove 1
    getProjectTLibDataType.dtDataType = vbObject
End If
Print #99, "getProjectTLibDataType: fnd="; Not getProjectTLibDataType Is Nothing
End Function

' Is the datatype declared within module? Applies to Enums and UDTs.
Function getDataTypeInModule(ByVal sym As String, ByVal m As vbModule, Optional pa As procattributes = PROC_ATTR_DEFAULT Or PROC_ATTR_PRIVATE Or PROC_ATTR_PUBLIC) As vbDataType
Print #99, "getDataTypeInModule: sym="; sym; " m.n="; m.Name; " types.c="; m.Types.count; " e.c="; m.Enums.count; " pa="; pa
' is symbol a UDT?
' fixme: Public Enums can be in stds, classes or forms.
' but Types (and other things) must be in std or *Public* modules
' so need to revise testing of module types to check for Public modules (ActiveX only)
' and Public members.
' Hmmm. Can this Public searching be handled by TLI instead of Local?
Dim udt As vbType
For Each udt In m.Types
    Print #99, "type="; udt.typeName
    Print #99, "at="; Hex(udt.typeAttributes)
    Print #99, "guid="; udt.typeGUID
    Print #99, "m.c="; udt.typeMembers.count
Next
On Error Resume Next
Set udt = m.Types.Item(UCase(sym))
On Error GoTo 0
' UDTs default to Public
If Not udt Is Nothing Then
    If udt.typeAttributes And pa Then
        Set getDataTypeInModule = New vbDataType
        Set getDataTypeInModule.dtUDT = udt
        getDataTypeInModule.dtDataType = vbUserDefinedType
        getDataTypeInModule.dtType = tokProjectClass
    End If
End If
If getDataTypeInModule Is Nothing Then
    ' Is symbol an Enum?
    Dim en As vbEnum
    On Error Resume Next
    ' fixme: make Enums into vbDataType collection
    Set en = m.Enums.Item(UCase(sym))
    On Error GoTo 0
    If Not en Is Nothing Then
        If en.enumAttributes And pa Then ' Enums default to Public
            Set getDataTypeInModule = New vbDataType
            Set getDataTypeInModule.dtEnum = en
            getDataTypeInModule.dtDataType = vbLong
            getDataTypeInModule.dtType = tokProjectClass
        End If
    End If
End If
Print #99, "getDataTypeInModule: fnd="; Not getDataTypeInModule Is Nothing
End Function

Function getDataTypeInAllModules(ByVal sym As String, ByVal cm As vbModule) As vbDataType
Print #99, "getDataTypeInAllModules: sym="; sym
Dim m As vbModule
For Each m In currentProject.prjModules
    Print #99, "m="; m.Name
    If Not m Is cm Then
        Set getDataTypeInAllModules = getDataTypeInModule(sym, m, PROC_ATTR_DEFAULT Or PROC_ATTR_PUBLIC)
        If Not getDataTypeInAllModules Is Nothing Then Exit For
    End If
Next
Print #99, "getDataTypeInAllModules: fnd="; Not getDataTypeInAllModules Is Nothing
End Function

Function getDataType(ByVal tokens As Collection, Optional ByVal AnyAllowed As Boolean) As vbDataType
Dim dt As TliVarType
If IsEOL(tokens) Then
    Print #99, "Missing data type"
    MsgBox "Missing data type"
    Err.Raise 1
End If
Print #99, "getDataType: s="; tokens.Item(1).tokString
On Error Resume Next
Set getDataType = cDataTypes.Item(UCase(tokens.Item(1).tokString))
On Error GoTo 0
Print #99, "getDataType: 1st class="; Not getDataType Is Nothing
If Not getDataType Is Nothing Then
    ' Need to create VT_ANY type
    If getDataType.dtDataType = VT_VOID And Not AnyAllowed Then Err.Raise 1 ' Any not allowed
    tokens.Remove 1
    If getDataType.dtDataType = VT_BSTR And Not IsEOL(tokens) Then
        If tokens.Item(1).tokString = "*" Then
            ' As String * n - where n is a constant (possibly named), not an expression
            tokens.Remove 1 ' remove *
            Dim OneToken As New Collection
            OneToken.Add tokens.Item(1)
            tokens.Remove 1 ' remove constant
            If PassNumber > 1 Then
                Dim output_stack As New Collection
                oRPN.ConstantRPNize OptimizeConstantExpressions, OneToken, output_stack, vbLong
                Dim g As New vbDataType
                ' fixme: get rid of most vbDataType members. Use classes for each type?
                g.dtLength = output_stack.Item(1).tokValue ' assign String * length
                Set output_stack = Nothing ' remove all
                ' duplicate dt because we need different dtLength
                ' need to create COM datatype class and assign to dtVT, instead of copying
                g.dtAttributes = getDataType.dtAttributes
                g.dtDataName = getDataType.dtDataName
                g.dtDataType = VT_LPWSTR ' getDataType.dtDataType
                Print #99, "dtl="; g.dtLength
                Set getDataType = g
            End If
        End If
    ElseIf getDataType.dtDataType = vbObject Then ' Object
        getDataType.dtType = tokIDispatchInterface
    End If
Else
    Set getDataType = getDataTypeInModule(tokens.Item(1).tokString, currentModule)
' fixme: looks like getDataTypeInModule and InAll need to pass tokens to eliminate Remove 1 test
    If Not getDataType Is Nothing Then tokens.Remove 1
    If getDataType Is Nothing Then
        Set getDataType = getProjectTLibDataType(tokens)
        If getDataType Is Nothing Then
            Set getDataType = getDataTypeInAllModules(UCase(tokens.Item(1).tokString), currentModule)
            If Not getDataType Is Nothing Then tokens.Remove 1
        ElseIf Not IsEOL(tokens) Then
            If tokens.Item(1).tokString = "." Then
                tokens.Remove 1
                If Not getDataType.dtClass Is Nothing Then
                    Set getDataType = getDataTypeInModule(tokens.Item(1).tokString, getDataType.dtClass)
                    If getDataType Is Nothing Then tokens.Remove 1
                ElseIf Not getDataType.dtClassInfo Is Nothing Then
                    Err.Raise 1 ' Type in tlib not implemented
                Else
                    Err.Raise 1
                End If
            Else
                ' fixme: make this If test into method. tests for end-of-expression
                ' fixme: perhaps this test just isn't needed. let subsequent parsing error out.
                ' removed - Typeof ... {Then, Else, :} ... If tokens.Item(1).tokString <> "," And tokens.Item(1).tokString <> "=" And tokens.Item(1).tokString <> "(" And tokens.Item(1).tokString <> ")" And tokens.Item(1).tokString <> ";" Then Err.Raise 1
            End If
        End If
    End If
End If
Print #99, "getDataType: fnd="; Not getDataType Is Nothing
End Function

Function ConvertVarTypeToVBType(ByVal dt As TliVarType) As TliVarType
Print #99, "ConvertVarTypeToVBType: dt="; dt
Select Case dt
Case VT_UI4 ' stdole2.OLE_COLOR is VT_UI4 (VT_UINT?)
    ConvertVarTypeToVBType = vbLong
Case Else
    Dim v As Variant
    On Error GoTo UnsupportedVarType
    ConvertVarTypeToVBType = varType(cVariantTypes.Item(CStr(dt)))
End Select
Exit Function
UnsupportedVarType:
    Print #99, "ConvertVarTypeToVBType: Unsupported VarType:"; dt
    MsgBox "ConvertVarTypeToVBType: Unsupported VarType:" & dt
    Err.Raise 1
End Function

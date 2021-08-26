Attribute VB_Name = "vbtcsubs"
Option Explicit

Private Enum eCOperatorPriority
    COprPriorityLogicalOr
    COprPriorityLogicalAnd
    COprPriorityBitwiseOr
    COprPriorityBitwiseXor
    COprPriorityBitwiseAnd
    COprPriorityEqNe
    COprPriorityLtGtLeGe
    COprPriorityAddSub
    COprPriorityMulDivMod
    COprPriorityUnary
    COprPriorityPrimary
End Enum

'Private nlp As Long ' literal pool counter
Public Const uniq = ""  ' prepend to generated names to make them unique
Public Const IndentString As String = "  "
Public indent As String
Public newindent As String
Public AdditionalTypeLibs As New Collection

Function CollectReferencedTypeLibs(ByVal References As References)
Print #99, "CollectReferencedTypeLibs: r.c="; References.Count
Dim ref As Reference
'On Error GoTo next_ref ' ignore uninspectable typelibs
For Each ref In References
    Print #99, "Processing ref="; ref.Name; " path="; ref.FullPath; " guid="; ref.GUID; " broken="; ref.IsBroken; " major="; ref.Major; " minor="; ref.Minor
    CollectTypeLib ref.FullPath, ref.GUID
'next_ref:
Next
Print #99, "CollectReferencedTypeLibs: atl.c="; AdditionalTypeLibs.Count
End Function

Function CollectTypeLib(ByVal FullPath As String, ByVal GUID As String) As TypeLibInfo ', ByVal major As Long, ByVal minor As Long)
Dim tlib As TypeLibInfo
On Error Resume Next
Set tlib = AdditionalTypeLibs.Item(GUID)
On Error GoTo 0
If tlib Is Nothing Then
' fixme: Why registry method doesn't work (VBA). Is it LCID?
'        Set tlib = TypeLibInfoFromRegistry(ref.GUID, ref.Major, ref.Minor, 0)
    Set tlib = TypeLibInfoFromFile(FullPath)
    If tlib Is Nothing Then Err.Raise 1
    Print #99, "tlib: n="; tlib.Name; " guid="; tlib.GUID
    AdditionalTypeLibs.Add tlib, tlib.GUID
End If
Set CollectTypeLib = tlib
End Function

Function define_guid(ByVal guidType As String, ByVal guidName As String, ByVal GUID As String) As String
Dim i As Integer
If guidType = "CLSID" Then ' bit of a kludge
    define_guid = guidType & " const " & guidType & "_" & guidName
Else
    define_guid = "IID" & " const " & guidType & "_" & guidName
End If
If GUID = "" Then
    define_guid = "EXTERN_C " & define_guid
Else
    define_guid = define_guid & " = { 0x" & Mid(GUID, 2, 8) & ", 0x" & Mid(GUID, 11, 4) & ", 0x" & Mid(GUID, 16, 4) & ", { 0x" & Mid(GUID, 21, 2) & ", 0x" & Mid(GUID, 23, 2)
    For i = 26 To 26 + 10 Step 2
        define_guid = define_guid & ", 0x" & Mid(GUID, i, 2)
    Next
    define_guid = define_guid & " } }"
End If
define_guid = define_guid & "; /* " & GUID & " */"
End Function

Function ResolveTypeKind(ByVal ii As InterfaceInfo) As VarTypeInfo
Print #99, "ResolveTypeKind: name="; ii.Name; " tk="; ii.TypeKind; " tks="; ii.TypeKindString
Set ResolveTypeKind = ii.ResolvedType
Print #99, "ResolveTypeKind: 1"
Do While ResolveTypeKind.TypeInfo.TypeKind = TKIND_ALIAS
Print #99, "ResolveTypeKind: 2"
    Set ResolveTypeKind = ResolveTypeKind.TypeInfo.ResolvedType
Print #99, "ResolveTypeKind: 3"
Loop
Print #99, "ResolveTypeKind: 4"
Print #99, "ResolveTypeKind: resolved="; ResolveTypeKind.TypeInfo.TypeKind
End Function

Function CInterfaceType(ByVal ii As InterfaceInfo) As String
' Pass initial InterfaceInfo, not VTableInterface version
Print #99, "CInterfaceType: "; ii.Name; " iic="; ii.ImpliedInterfaces.Count; " am="; Hex(ii.AttributeMask); " tk="; ii.TypeKind
If ii.ImpliedInterfaces.Count = 0 Then
    ' do nothing
ElseIf ii.TypeKind = TKIND_DISPATCH Then
    CInterfaceType = "IDispatch"
ElseIf ii.TypeKind = TKIND_INTERFACE Then
    CInterfaceType = "IUnknown"
Else
    Print #99, "CInterfaceType: Unknown interface: "; ii.Name; ii.ImpliedInterfaces.Count; Hex(ii.AttributeMask); ii.TypeKind
    MsgBox "CInterfaceType: Unknown interface: " & ii.Name & ii.ImpliedInterfaces.Count & Hex(ii.AttributeMask) & ii.TypeKind
    Err.Raise 1 ' Unknown interface type
End If
End Function

Function ValueToC(ByVal dt As TliVarType, ByVal v As Variant) As String
Print #99, "ValueToC: dt="; dt; " vt="; VarType(v); " n="; TypeName(v)
'If IsEmpty(v) Then
'    ValueToC = CDefaultValueByType(dt And Not VT_BYREF)
If (dt And Not VT_BYREF) = VT_VARIANT And VarType(v) = vbObject Then
    ValueToC = "VarNothing"
Else
    dt = VarType(v) Or (dt And VT_BYREF)
    Dim i As Integer
    Dim s As String
    Dim ss As String
    Select Case dt And Not VT_BYREF
        Case VT_EMPTY, VT_NULL
            ValueToC = CDefaultValueByType(dt And Not VT_BYREF)
        Case VT_R4
            ValueToC = v
            ' Assumes upper case E in exponential numbers (1E+7)
            If InStr(ValueToC, ".") = 0 And InStr(ValueToC, "E") = 0 Then ValueToC = ValueToC & "."
            ValueToC = ValueToC & "f" ' need f suffix to force to float (MSVC)
        Case VT_R8
            ValueToC = v
            If InStr(ValueToC, ".") = 0 And InStr(ValueToC, "E") = 0 Then ValueToC = ValueToC & "."
        Case VT_LPSTR ' VBA.Constants.vbCrLf for example
            ss = StrConv(v, vbUnicode)
'            Print #99, "lenb="; Len(v); " asc="; Asc(Mid(s, i, 1)); " ascb="; AscB(MidB(s, i, 1))
' Need StrToC routine - assumes ascii strings and not wide!!!!
            Dim cat As String
            cat = """ /* catenating strings */ L""" ' Len(cat) must be > 8
            For i = 1 To Len(ss)
                ' MSVS C raises error for strings > 1021 but allows string catenation
                If Len(s) Mod 1024 > 1024 - Len(cat) Then s = s & cat
                Select Case Asc(Mid(ss, i, 1))
                    Case 34 ' "
                        s = s & "\"""
                    Case 92 ' \
                        s = s & "\\"
                    Case 32 To 33, 35 To 91, 93 To 126
                        s = s & Mid(ss, i, 1)
                    Case Else
                        s = s & "\" & Right("00" & Oct(Asc(Mid(ss, i, 1))), 3)
                End Select
            Next
            ValueToC = "LToStr(L""" & s & """)"
        Case VT_BSTR
' Need StrToC routine - assumes ascii strings and not wide!!!!
            ss = v
            cat = """ /* catenating strings */ L""" ' Len(cat) must be > 8
            For i = 1 To Len(ss)
                ' MSVS C raises error for strings > 1021 but allows string catenation
                If Len(s) Mod 1024 > 1024 - Len(cat) Then s = s & cat
                Select Case Asc(Mid(ss, i, 1))
                    Case 34 ' "
                        s = s & "\"""
                    Case 92 ' \
                        s = s & "\\"
                    Case 32 To 33, 35 To 91, 93 To 126
                        s = s & Mid(ss, i, 1)
                    Case Else
                        s = s & "\" & Right("00" & Oct(Asc(Mid(ss, i, 1))), 3)
                End Select
            Next
            ValueToC = "LToStr(L""" & s & """)"
        Case VT_DISPATCH ' occurs when parameter defaultvalue is Nothing
            ValueToC = "ObjNothing"
        Case VT_CY 'fixme: need to output I8 scaled by 10000????
            ValueToC = "DblToCur(" & v & ")"
        Case VT_DATE
            ValueToC = "LToDate(L""" & v & """)"
        Case Else
            ValueToC = v
            ' wow - bug in MSVC 6.0 causes spurious warning for -2147483648 (LONG_MIN)
            ' note: MSVC gives signed/unsigned warning on LngOr(l,0x80000000) but not on LngAnd or others
            If ValueToC = "-2147483648" Then ValueToC = "0x80000000"
        End Select
End If
If dt And VT_BYREF Then ValueToC = AbbrDataTypeRef(dt, ValueToC)
Print #99, "ValueToC: s="; ValueToC
End Function

' Can this be replaced by a TLI function?
Function CDefaultValueByType(ByVal vt As TliVarType) As String
Print #99, "CDefaultValueByType: vt="; vt
If vt And VT_BYREF Then Err.Raise 1
    Select Case vt
        Case VT_EMPTY
            CDefaultValueByType = "_VarEmpty"
        Case VT_NULL
            CDefaultValueByType = "_VarNull"
        Case VT_I2
            CDefaultValueByType = "0"
        Case VT_BOOL
            CDefaultValueByType = "0"
        Case VT_I4
            CDefaultValueByType = "0"
        Case VT_R4
            CDefaultValueByType = "0"
        Case VT_R8
            CDefaultValueByType = "0"
        Case VT_CY
            CDefaultValueByType = "0"
        Case VT_BSTR
            CDefaultValueByType = "LToStr(L"""")"
        Case VT_VARIANT
            CDefaultValueByType = "_VarNull"
        Case VT_UI1
            CDefaultValueByType = "0"
        Case VT_DISPATCH
            CDefaultValueByType = "ObjByRefNothing"
        Case Else
            Print #99, "CDefaultValueByType: Unknown VarType: " & vt
            MsgBox "CDefaultValueByType: Unknown VarType: " & vt
            Err.Raise 1 ' Internal error - Unknown VarType
    End Select
'End If
Print #99, "CDefaultValueByType: done:"; CDefaultValueByType
End Function

Function VarTypeInfoToCType(ByVal vi As VarTypeInfo, Optional ByVal retval As Integer) As String
If vi Is Nothing Then
    VarTypeInfoToCType = "void"
Else
    On Error Resume Next
    Print #99, "VarTypeInfoToCType: vi="; Not vi Is Nothing; " rv="; retval
    Print #99, "VarTypeInfoToCType: vi.typeinfo="; Not vi.TypeInfo Is Nothing
    Print #99, "VarTypeInfoToCType: vi.typeinfo.name="; vi.TypeInfo.Name
    Print #99, "VarTypeInfoToCType: vi.vartype="; vi.VarType
    Print #99, "VarTypeInfoToCType: vi.IsExternalType="; vi.IsExternalType
    Print #99, "VarTypeInfoToCType: vi.PointerLevel="; vi.PointerLevel
    Print #99, "VarTypeInfoToCType: vi.TypeInfoNumber="; vi.TypeInfoNumber
    On Error GoTo 0
    Dim ti As TypeInfo
    If vi.IsExternalType Then
        Set ti = GetTypeInfoFromTLib(vi)
    Else
        Set ti = vi.TypeInfo
    End If
    If Not ti Is Nothing Then
        Print #99, "VarTypeInfoToCType: ti.Name="; ti.Name; " kind="; ti.TypeKind; " tin="; ti.TypeInfoNumber
' needed?
        On Error Resume Next
        If Not ti.ResolvedType Is Nothing Then
            Print #99, "VarTypeInfoToCType: ti.ResolvedType.vartype="; ti.ResolvedType.VarType
            Print #99, "VarTypeInfoToCType: ti.ResolvedType.typeinfo.name="; ti.ResolvedType.TypeInfo.Name
        End If
        On Error GoTo 0
        Print #99, "vi.vt="; vi.VarType; " ti.tk="; ti.TypeKind
        If vi.VarType = 0 Then
' Note that VBA_Err returns a CoClass so use default interface
            Select Case ti.TypeKind
                Case TKIND_COCLASS
                    If vi.PointerLevel <= 0 Then Err.Raise 1 ' TypeLib error
Print #99, "mmmm"
                    VarTypeInfoToCType = TypeInfoToCType(ti)
Print #99, "nnnn"
                    ' Some return values (such as TypeLibInfoFromRegistry) have PointerLevel=1, should be 2?
                    If vi.PointerLevel = retval Then VarTypeInfoToCType = VarTypeInfoToCType & "*"
Print #99, "oooo"
                Case TKIND_INTERFACE, TKIND_DISPATCH
'                    If TypeOf ti Is InterfaceInfo Then
'                        Dim ii As InterfaceInfo
'                        Set ii = ti.VTableInterface
'                        If ii Is Nothing Then Set ii = ti
'                    End If
'                    If vi.PointerLevel < 0 Then Err.Raise 1 ' TypeLib error
                    VarTypeInfoToCType = TypeInfoToCType(ti)
                    ' Some return values have PointerLevel=1, should be 2?
                    If vi.PointerLevel = retval Then VarTypeInfoToCType = VarTypeInfoToCType & "*"
                Case TKIND_ENUM
                    VarTypeInfoToCType = TypeInfoToCType(ti)
                    If vi.PointerLevel <> 0 Then 'Err.Raise 1 ' TypeLib error
                        ' ugh, VBRUN.ParentControlsType has Enum with PointerLevel > 0
                        ' fixme: implement warnings
                        Print #99, "Warning: TypeLib error: Enum "; VarTypeInfoToCType; " has PointerLevel of "; vi.PointerLevel
                    End If
                Case TKIND_RECORD
                    VarTypeInfoToCType = TypeInfoToCType(ti)
                Case TKIND_UNION ' adirec1a\project1.vbp - DirectX 7 VB Lib
' fixme: for non-MSVC compilers, may require prefixing "union _" (where _ is typedef tag prefix)
                    VarTypeInfoToCType = TypeInfoToCType(ti)
                Case Else
                    Print #99, "VarTypeInfoToCType: invalid tk: "; ti.TypeKind
                    Err.Raise 1
            End Select
        Else
'            Print #99, "vttc="; VarTypeToC(VarTypeInfoToVarType(vi)) ' doesnt support UDTs
'            Print #99, "tict="; TypeInfoToCType(ti)
            If vi.VarType And VT_ARRAY Then
                VarTypeInfoToCType = VarTypeToC(vi.VarType) ' = "SAFEARRAY"
            Else
                VarTypeInfoToCType = TypeInfoToCType(ti)
            End If
        End If
    Else
        VarTypeInfoToCType = VarTypeToC(VarTypeInfoToVarType(vi))
    End If
Print #99, "pppp"
    If vi.PointerLevel - retval > 0 Then VarTypeInfoToCType = VarTypeInfoToCType & String(vi.PointerLevel - retval, "*")
    Print #99, "pl="; vi.PointerLevel; " retval="; retval; " vartype="; vi.VarType
End If
Print #99, "VarTypeInfoToCType: done: "; VarTypeInfoToCType
On Error GoTo 0 ' don't need this here
End Function

' Return TypeInfo from newest TypeLib
' This is needed because OfficeBars member of mso97 (v2.0) had error (s/b _OfficeBars)
' This was fixed in newer mso9 (v2.1) fixed.
' Note GUIDs are same between TLibs but filenames maybe different and major version or minor version are higher (newer).
Function GetTypeInfoFromTLib(ByVal vi As VarTypeInfo) As TypeInfo
Print #99, "GetTypeInfoFromTLib: vi="; Not vi Is Nothing
If Not vi.IsExternalType Then Err.Raise 1
Dim eTLib As TypeLibInfo
Set eTLib = vi.TypeLibInfoExternal
On Error Resume Next
Print #99, "vi.vt="; vi.VarType
Print #99, "vi.ti.name="; vi.TypeInfo.Name
Print #99, "vi.Variant="; VarType(vi.TypedVariant)
Print #99, "vi.Variant.tn="; TypeName(vi.TypedVariant)
Print #99, "eTLib.n="; eTLib.Name; " f="; eTLib.ContainingFile; " guid="; eTLib.GUID; " tic"; eTLib.TypeInfos.Count; " major="; eTLib.MajorVersion; " minor="; eTLib.MinorVersion
Print #99, "n="; eTLib.TypeInfos.IndexedItem(vi.TypeInfoNumber).Name
Print #99, "tk="; eTLib.TypeInfos.IndexedItem(vi.TypeInfoNumber).TypeKind
Print #99, "am="; eTLib.TypeInfos.IndexedItem(vi.TypeInfoNumber).AttributeMask
Print #99, "tin="; eTLib.TypeInfos.IndexedItem(vi.TypeInfoNumber).TypeInfoNumber
Print #99, "resolved="; eTLib.TypeInfos.IndexedItem(vi.TypeInfoNumber).ResolvedType.TypeInfo.Name
Print #99, "bet="; eTLib.BestEquivalentType(vi.TypeInfo.Name)
Print #99, eTLib.TypeInfos.Item(vi.TypeInfo.Name).Name
Print #99, eTLib.TypeInfos.Item(vi.TypeInfo.Name).TypeKind
Print #99, eTLib.TypeInfos.Item(vi.TypeInfo.Name).TypeInfoNumber
On Error GoTo 0
If vi.VarType <> 0 Then
    Print #99, "GetTypeInfoFromTLib: unexpected external VarType=" & vi.VarType
    MsgBox "GetTypeInfoFromTLib: unexpected external VarType=" & vi.VarType
    Err.Raise 1
End If
Dim tlib As TypeLibInfo
On Error Resume Next
Set tlib = AdditionalTypeLibs.Item(eTLib.GUID)
On Error GoTo 0
If tlib Is Nothing Then
' STDOLE2.TLB {00020430-0000-0000-C000-000000000046} is implied
    Print #99, "GetTypeInfoFromTLib: can't load typelib: "; eTLib.Name; ". Attempting to resolve: "; vi.TypeInfo.Name
'    MsgBox "GetTypeInfoFromTLib: can't load typelib: " & eTLib.Name & ". Attempting to resolve: " & vi.TypeInfo.Name
'    Err.Raise 1
    Set tlib = CollectTypeLib(vi.TypeLibInfoExternal.ContainingFile, vi.TypeLibInfoExternal.GUID)
 '   Set ti = vi.TypeInfo
End If
    Print #99, "tlib:"; tlib.Name; " guid="; tlib.GUID; " f="; tlib.ContainingFile; " tic"; tlib.TypeInfos.Count
    On Error Resume Next
#If 0 Then
    For Each ti In tlib.TypeInfos
        Print #99, "ti="; ti.Name; ti.TypeKind; ti.TypeInfoNumber
    Next
    For Each ti In tlib.CoClasses
        Print #99, "ci="; ti.Name; ti.TypeKind; ti.TypeInfoNumber
    Next
    Set GetTypeInfoFromTLib = tlib.CoClasses.NamedItem(vi.TypeInfo.Name) ' could be CoClass name too
    If GetTypeInfoFromTLib Is Nothing Then Set GetTypeInfoFromTLib = tlib.Interfaces.NamedItem(vi.TypeInfo.Name) ' could be CoClass name too
#End If
    ' fixme: Hopefully TypeInfos can't have duplicate names.
    Set GetTypeInfoFromTLib = tlib.TypeInfos.NamedItem(vi.TypeInfo.Name) ' could be CoClass name too
    On Error GoTo 0
    If GetTypeInfoFromTLib Is Nothing Then
        Print #99, "GetTypeInfoFromTLib: can't find external interface: "; vi.TypeInfo.Name
        MsgBox "GetTypeInfoFromTLib: can't find external interface: " & vi.TypeInfo.Name
        Err.Raise 1
    End If
'End If
Print #99, "GetTypeInfoFromTLib: n="; GetTypeInfoFromTLib.Name; " guid="; GetTypeInfoFromTLib.GUID
End Function

' must use emitter type priority in case (C) has different operator priority than VB
Function EmitInFix(ByVal proc As procTable, ByVal output_stack As Collection) As String
Dim token As vbToken
Dim operand_stack As New Collection
Dim i As Integer
Dim s As String
newindent = indent
For Each token In output_stack
DoEvents
token.tokOutput = token.tokString
Print #99, "EmitInFix: o="; token.tokOutput; " t="; token.tokType; " dt="; token.tokDataType; " pc="; token.tokPCode; " pcst="; token.tokPCodeSubType; " tc="; token.tokCount; " r="; token.tokRank; " osc="; operand_stack.Count
token.tokPriority = 0
'MsgBox token.tokOutput & " " & token.tokType
' fixme: check that token and operands are same data type
Select Case token.tokType
    Case tokOperator ' not implemented yet
        If token.tokPCode = vbPCodePositive Or token.tokPCode = vbPCodeNegative Or token.tokPCode = vbPCodeNot Then
            token.tokOutput = AbbrDataType(token.tokPCodeSubType) & PCodeToAbbr(token.tokPCode) & "(" & operand_stack.Item(operand_stack.Count).tokOutput & ")"
            operand_stack.Remove operand_stack.Count
        Else
            ' does left expression need parenthesis?
            Print #99, "LHS tp="; operand_stack.Item(operand_stack.Count - 1).tokPriority; " tpc="; COperatorPriority(token.tokPCode)
            If operand_stack.Item(operand_stack.Count - 1).tokPriority <> 0 And operand_stack.Item(operand_stack.Count - 1).tokPriority < COperatorPriority(token.tokPCode) Then
                operand_stack.Item(operand_stack.Count - 1).tokOutput = "(" & operand_stack.Item(operand_stack.Count - 1).tokOutput & ")"
            End If
            ' does right expression need parenthesis?
            Print #99, "RHS tp="; operand_stack.Item(operand_stack.Count).tokPriority; " tpc="; COperatorPriority(token.tokPCode)
            If operand_stack.Item(operand_stack.Count).tokPriority <> 0 And operand_stack.Item(operand_stack.Count).tokPriority < COperatorPriority(token.tokPCode) Then
                operand_stack.Item(operand_stack.Count).tokOutput = "(" & operand_stack.Item(operand_stack.Count).tokOutput & ")"
            End If
            ' new expression uses operator priority
            token.tokPriority = COperatorPriority(token.tokPCode)
    '        operand_stack.Item(operand_stack.Count - 1).tokOutput = CoerseBinaryOperands(operand_stack.Item(operand_stack.Count - 1), token, operand_stack.Item(operand_stack.Count))
            token.tokOutput = AbbrDataType(token.tokPCodeSubType) & PCodeToAbbr(token.tokPCode) & "(" & operand_stack.Item(operand_stack.Count - 1).tokOutput & "," & operand_stack.Item(operand_stack.Count).tokOutput & ")"
            operand_stack.Remove operand_stack.Count
            operand_stack.Remove operand_stack.Count
        End If
        operand_stack.Add token
    Case tokLocalModule, tokGlobalModule ' reference is within module, no object reference is used
Print #99, "tokLocalModule: n="; token.tokLocalFunction.procname
Print #99, "tokLocalModule: ik="; token.tokLocalFunction.InvokeKind
Print #99, "tokLocalModule: mn="; token.tokLocalFunction.procLocalModule.Name
Print #99, "tokLocalModule: ct="; token.tokLocalFunction.procLocalModule.Component.Type
Print #99, "tokLocalModule: pc="; token.tokLocalFunction.procParams.Count
Print #99, "tokLocalModule: pcst="; token.tokPCodeSubType
Print #99, "tokLocalModule: v="; Not token.tokVariable Is Nothing
If token.tokLocalFunction.procParams.Count > 0 Then Print #99, "tokLocalModule: last param="; Hex(token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.Count).paramVariable.varAttributes)

Dim ParamArrayCnt As Long
If Not token.tokVariable Is Nothing Then If Not token.tokVariable.varDimensions Is Nothing Then Print #99, "tokLocalModule: vd="; token.tokVariable.varDimensions.Count
        If token.tokLocalFunction.procParams.Count > 0 Then
            If CBool(token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.Count).paramVariable.varAttributes And VARIABLE_PARAMARRAY) Then
                ParamArrayCnt = token.tokCount - token.tokLocalFunction.procParams.Count + 1
                If ParamArrayCnt < 0 Then Err.Raise 1
            End If
        End If
        Dim interfacename As String
        interfacename = IIf(token.tokPCodeSubType = INVOKE_EVENTFUNC, token.tokLocalFunction.procLocalModule.EventName, token.tokLocalFunction.procLocalModule.Name)
        Select Case token.tokLocalFunction.procLocalModule.Component.Type
            Case vbext_ct_StdModule
'                token.tokOutput = token.tokLocalFunction.procLocalModule.Name & "_" & EmitVariable(token, operand_stack)
' note: tokPCodeSubType is either zero (Func or Get) or actually used InvokeKind, needs to be Ored with available types (variables are Get/Let/Set but not Func).
'   probably wouldn't need to be And'ed if variables weren't Or'ed together.
                token.tokOutput = interfacename & "_" & ProcNameIK(token.tokLocalFunction.procname, IIf(token.tokPCodeSubType, token.tokPCodeSubType, INVOKE_FUNC Or INVOKE_PROPERTYGET) And token.tokLocalFunction.InvokeKind) & EmitVariable(token, operand_stack, , ParamArrayCnt)
            Case Else
'                token.tokOutput = "___" & interfacename & "_" & ProcNameIK(token.tokOutput, IIf(token.tokLocalFunction.procattributes And PROC_ATTR_VARIABLE, token.tokPCodeSubType And (INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF), INVOKE_FUNC))
'                token.tokOutput = "This"
                 ' was vbext_ct_ClassModule)
                ' fixme: perhaps retval token is better
'                If token.tokLocalFunction Is proc Then ' removed - didn't work in class module
'                    token.tokOutput = "(*retval)"
#If 0 Then ' UDT ref
                    If (token.tokDataType And Not VT_ARRAY) = VT_RECORD Then token.tokOutput = "(*" & token.tokOutput & ")"
#End If
'                Else
Print #99, "plm="; token.tokLocalFunction.procLocalModule.Name; token.tokLocalFunction.procLocalModule.interfaceGUID
Print #99, "module="; proc.procLocalModule.Name; proc.procLocalModule.interfaceGUID
                    If token.tokPCodeSubType = INVOKE_EVENTFUNC Then
                        ' used for RaiseEvent
                        token.tokOutput = "_i_" & interfacename & "QI(This)"
                    ElseIf token.tokLocalFunction.procLocalModule.interfaceGUID = proc.procLocalModule.interfaceGUID Then
                        token.tokOutput = "This"
                    Else ' Forms, and globally instantiated classes (appobjects?)
                        token.tokOutput = "_v_" & token.tokLocalFunction.procLocalModule.Name & "_" & interfacename
                    End If
' note: tokPCodeSubType is either zero (Func or Get) or actually used InvokeKind, needs to be Ored with available types (variables are Get/Let/Set but not Func).
'   probably wouldn't need to be And'ed if variables weren't Or'ed together.
                    token.tokOutput = "___" & interfacename & "_" & ProcNameIK(token.tokLocalFunction.procname, IIf(token.tokPCodeSubType, token.tokPCodeSubType, INVOKE_FUNC Or INVOKE_PROPERTYGET) And token.tokLocalFunction.InvokeKind) & EmitVariable(token, operand_stack, token.tokLocalFunction.procLocalModule.Component.Type, ParamArrayCnt)
'                End If
'            Case Else
'                Print #99, "EmitInFix: Unknown ComponentType: " & token.tokLocalFunction.procLocalModule.component.type
'                MsgBox "EmitInFix: Unknown ComponentType: " & token.tokLocalFunction.procLocalModule.component.type
'                Err.Raise 1 ' Module has unknown ComponentType
        End Select
'        If token.tokVariable.varAttributes And VARIABLE_NEW Then token.tokOutput = "New(" & token.tokOutput & ")"
        ProcessSubscripts token, operand_stack
        operand_stack.Add token
#If 0 Then
    Case tokGlobalModule
Print #99, "tokGlobalModule: tlfn="; token.tokLocalFunction.procname
Print #99, "tokGlobalModule: lmn="; token.tokLocalFunction.procLocalModule.Name
Print #99, "tokGlobalModule: ct="; token.tokLocalFunction.procLocalModule.Component.Type
Print #99, "tokGlobalModule: pc="; token.tokLocalFunction.procParams.Count
If token.tokLocalFunction.procParams.Count > 0 Then Print #99, "tokGlobalModule: last param="; Hex(token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.Count).paramVariable.varAttributes)
        token.tokOutput = ""
        If token.tokLocalFunction.procParams.Count > 0 Then
            If CBool(token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.Count).paramVariable.varAttributes And VARIABLE_PARAMARRAY) Then
                ParamArrayCnt = token.tokCount - token.tokLocalFunction.procParams.Count + 1
                If ParamArrayCnt < 0 Then Err.Raise 1
            End If
        End If
        Select Case token.tokLocalFunction.procLocalModule.Component.Type
            Case vbext_ct_StdModule
                token.tokOutput = token.tokLocalFunction.procLocalModule.Name & "_" & ProcNameIK(token.tokLocalFunction.procname, token.tokPCodeSubType) & EmitVariable(token, operand_stack, , ParamArrayCnt)
            Case Else
                token.tokOutput = "___" & token.tokLocalFunction.procLocalModule.Name & "_" & ProcNameIK(token.tokLocalFunction.procname, token.tokPCodeSubType) & EmitVariable(token, operand_stack, token.tokLocalFunction.procLocalModule.Component.Type, ParamArrayCnt)
'                operand_stack.Remove operand_stack.Count
'            Case Else
'                Print #99, "EmitInFix: Unknown ComponentType: " & token.tokLocalFunction.procLocalModule.component.type
'                MsgBox "EmitInFix: Unknown ComponentType: " & token.tokLocalFunction.procLocalModule.component.type
'                Err.Raise 1 ' Module has unknown ComponentType
        End Select
'        If token.tokVariable.varAttributes And VARIABLE_NEW Then token.tokOutput = "New(" & token.tokOutput & ")"
        ProcessSubscripts token, operand_stack
        operand_stack.Add token
#End If
    Case tokQI_Module
        token.tokOutput = "_i_" & token.tokModule.Name & "QI" & IIf(token.tokDataType And VT_BYREF, "Ref", "") & "(" & operand_stack.Item(operand_stack.Count).tokOutput & ")"
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case tokQI_TLibInterface
        token.tokOutput = TypeInfoToCType(token.tokInterfaceInfo) & "QI" & IIf(token.tokDataType And VT_BYREF, "Ref", "") & "(" & operand_stack.Item(operand_stack.Count).tokOutput & ")"
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case tokDeclarationInfo ' Uses DLL interface, no This
        Print #99, "tokDeclarationInfo: mi exists: "; token.tokMemberInfo Is Nothing; " tc="; token.tokCount; " dt="; token.tokDataType; " sym="; PCodeToSymbol(token)
        If token.tokMemberInfo Is Nothing Then Err.Raise 1
        Print #99, "mi.am="; Hex(token.tokMemberInfo.AttributeMask); " pc="; token.tokMemberInfo.Parameters.Count; " ik="; token.tokMemberInfo.InvokeKind
        s = ""
        If token.tokMemberInfo.InvokeKind And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
            s = operand_stack.Item(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
        End If
        Dim ss As String
        ss = ""
        Dim pi As ParameterInfo
        i = operand_stack.Count - token.tokCount + 1
        Dim k As Integer
        k = 0
        Dim sss As String
        Dim dt As TliVarType
        For Each pi In token.tokMemberInfo.Parameters
            k = k + 1
            Print #99, "p1 n="; pi.Name; " dt="; pi.VarTypeInfo.VarType; " f="; Hex(pi.Flags); " i="; i; " pl="; pi.VarTypeInfo.PointerLevel; " k="; k; " tc="; token.tokCount
            If k <= token.tokCount Then
                dt = VarTypeInfoToVarType(pi.VarTypeInfo)
' fixme: create token flag tokParamArray and use it here?
                If IsParamArrayArg(token.tokMemberInfo, k) Then
                    sss = ""
                    While k <= token.tokCount
                        k = k + 1
                        Print #99, "os: ts="; operand_stack.Item(i).tokString; " dt="; operand_stack.Item(i).tokDataType
                        sss = sss & "," & operand_stack.Item(i).tokOutput
                        operand_stack.Remove i
                    Wend
                    sss = "VarArg(" & token.tokCount - token.tokMemberInfo.Parameters.Count + 1 & sss & ")"
                Else
                    Print #99, "os: ts="; operand_stack.Item(i).tokString; " dt="; dt; " os.dt="; operand_stack.Item(i).tokDataType
                    If dt <> (operand_stack.Item(i).tokDataType And Not VT_BYREF) Then Err.Raise 1
                    sss = operand_stack.Item(i).tokOutput
                    operand_stack.Remove i
                End If
            ElseIf pi.Flags And PARAMFLAG_FRETVAL Then
                ' do nothing
            ElseIf pi.Flags And PARAMFLAG_FLCID Then ' LCID
                dt = vbLong
                sss = "0" ' not implemented - just default
            ElseIf pi.Optional Or CBool(pi.Flags And PARAMFLAG_FOPT) Then
                If pi.Optional Xor CBool(pi.Flags And PARAMFLAG_FOPT) Then Print #99, "Funky optional problem"
                If pi.Flags And PARAMFLAG_FHASDEFAULT Then
                    dt = VarType(pi.DefaultValue)
                    sss = CDefaultValueByType(dt)
                Else
                    dt = vbVariant
                    sss = CDefaultValueByType(dt)
                End If
            Else
                Print #99, "Expecting optional parameter"
                Err.Raise 1 ' Expecting optional parameter
            End If
            Dim j As Integer
#If 0 Then
            For j = 1 To pi.VarTypeInfo.PointerLevel - IIf(dt And VT_ARRAY, 1, 0)
                sss = AbbrDataTypeRef(dt, sss)
            Next
#End If
            ss = ss & "," & sss
        Next
Print #99, "s="; s
Print #99, "ss="; ss
        token.tokOutput = PCodeToSymbol(token) & "(" & Mid(ss, 2) & s & ")"
        operand_stack.Add token
    Case tokVariable, tokArrayVariable
Print #99, "tokVariable: v="; Not token.tokVariable Is Nothing; " at="; Hex(token.tokVariable.varAttributes)
Print #99, "tokVariable: p="; Not token.tokVariable.varProc Is Nothing
Print #99, "tokVariable: m="; Not token.tokVariable.varModule Is Nothing
process_variable:
            Print #99, "tokVariable: dt="; token.tokDataType; " tc="; token.tokCount; " mt="; token.tokVariable.MemberType
'            ' v(1,2,3) - where v is a Variant
'            If token.tokCount > 0 Then
'                If token.tokDataType <> vbVariant Then Err.Raise 1
'                GoTo generic_invoke1
'            End If
' done in EmitVariable
'            If token.tokVariable.varProc Is Nothing Then
'                token.tokOutput = "_v_" & token.tokVariable.varModule.Name & "_" & token.tokOutput
'            End If
            token.tokOutput = EmitVariable(token, operand_stack, -1)
#If 0 Then
    If token.tokType = tokVariable Or token.tokType = tokArrayVariable Then
        token.tokOutput = AbbrDataType(token.tokPCodeSubType) & PCodeToAbbr(token.tokPCode) & "(" & operand_stack.Item(operand_stack.Count - 1).tokOutput & "," & operand_stack.Item(operand_stack.Count).tokOutput & ")"
    Else
        If token.tokPCodeSubType And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
            Print #99, "osc="; operand_stack.Count
            token.tokOutput = token.tokOutput & "," & operand_stack.Item(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
        End If
    End If
#End If
            ProcessSubscripts token, operand_stack
'            If token.tokVariable.varAttributes And VARIABLE_NEW Then token.tokOutput = "GetObj(" & token.tokOutput & ")"
            operand_stack.Add token
    Case tokUDT
' fixme: using -1 to force variable processing
        Dim rankOffset As Long
        rankOffset = token.tokRank
        If token.tokPCode = tokReDim Then rankOffset = rankOffset * 2
        If operand_stack.Item(operand_stack.Count - rankOffset).tokDataType And Not VT_ARRAY <> vbUserDefinedType Then
            Print #99, "EmitInFix: Expecting UDT: " & operand_stack.Item(operand_stack.Count - rankOffset).tokDataType
            MsgBox "EmitInFix: Expecting UDT: " & operand_stack.Item(operand_stack.Count - rankOffset).tokDataType
            Err.Raise 1 ' Member must return Object or Variant
        End If
        ' fixme: need to implement PointerLevel here
        ' must use operand_stack.Item(operand_stack.Count - rankOffset).tokOutput, it may contain subscript expression
        token.tokOutput = "(" & operand_stack.Item(operand_stack.Count - rankOffset).tokOutput & ")"
        Print #99, "udt: os.type="; operand_stack.Item(operand_stack.Count - rankOffset).tokType
        Select Case operand_stack.Item(operand_stack.Count - rankOffset).tokType
        Case tokWithValue
            token.tokOutput = token.tokOutput & "->"
        Case Else ' ByRef parameters are already subrefed
            token.tokOutput = token.tokOutput & "."
        End Select
        operand_stack.Remove operand_stack.Count - rankOffset
        token.tokOutput = token.tokOutput & "_v_" & token.tokVariable.varSymbol
        token.tokOutput = EmitVariable(token, operand_stack, -1)
        ProcessSubscripts token, operand_stack
        operand_stack.Add token
    Case tokWithValue
        token.tokOutput = AbbrDataType(token.tokDataType) & "WithValue(" & CStr(token.tokCount) & ")"
        If (token.tokDataType And Not VT_BYREF) = VT_RECORD Then
            Print #99, "WithValue: type="; token.tokVariable.VarType.dtType
            Select Case token.tokVariable.VarType.dtType
                Case tokProjectClass, tokFormClass
                    Print #99, "udt exists="; Not token.tokVariable.VarType.dtUDT Is Nothing
' fixme: create routine which forms a cast, use for SA,UDT,QI - Create macros for each type to do it?
                    token.tokOutput = "(_t_" & token.tokVariable.VarType.dtUDT.typeModule.Name & "_" & token.tokVariable.VarType.dtUDT.TypeName & IIf(token.tokDataType And VT_BYREF, "**)", "*)") & token.tokOutput
                Case tokReferenceClass
                    Print #99, "ri exists="; Not token.tokVariable.VarType.dtRecordInfo Is Nothing
                    token.tokOutput = "(" & TypeInfoToCType(token.tokVariable.VarType.dtRecordInfo) & IIf(token.tokDataType And VT_BYREF, "**)", "*)") & token.tokOutput
                Case Else
                    Err.Raise 1
            End Select
        End If
        operand_stack.Add token
    Case tokProjectClass
projectclass:
        If token.tokLocalFunction Is Nothing Then GoTo process_variable
'        If operand_stack.Item(operand_stack.Count).tokDataType <> vbObject Or operand_stack.Item(operand_stack.Count).tokDataType <> vbVariant Then Err.Raise 1 ' Member must return Object or Variant
Print #99, " var="; Not token.tokVariable Is Nothing
Print #99, " mi="; Not token.tokMemberInfo Is Nothing
Print #99, " lf="; Not token.tokLocalFunction Is Nothing
Print #99, " plm="; Not token.tokLocalFunction.procLocalModule Is Nothing
'        If Not token.tokLocalFunction Is Nothing Then
Print #99, " modulename="; token.tokLocalFunction.procLocalModule.Name
Print #99, " ik="; token.tokLocalFunction.InvokeKind
Print #99, " membername="; token.tokLocalFunction.procname; " dt="; token.tokDataType
Print #99, " ComponentType="; token.tokLocalFunction.procLocalModule.Component.Type
Print #99, " procParams.Count="; token.tokLocalFunction.procParams.Count
        If token.tokLocalFunction.procParams.Count > 0 Then
            If CBool(token.tokLocalFunction.procParams.Item(token.tokLocalFunction.procParams.Count).paramVariable.varAttributes And VARIABLE_PARAMARRAY) Then
                ParamArrayCnt = token.tokCount - token.tokLocalFunction.procParams.Count + 1
                If ParamArrayCnt < 0 Then Err.Raise 1
            End If
        End If
        Dim ik As InvokeKinds
        ik = token.tokLocalFunction.InvokeKind
            Select Case token.tokLocalFunction.procLocalModule.Component.Type
                Case vbext_ct_StdModule
                    token.tokOutput = token.tokLocalFunction.procLocalModule.Name & "_" & ProcNameIK(token.tokLocalFunction.procname, ik) & EmitVariable(token, operand_stack, , ParamArrayCnt)
                Case Else
Print #99, " dt="; token.tokDataType; " pcst="; token.tokPCodeSubType
Print #99, " var="; Not token.tokVariable Is Nothing
' removing cast because it looks suspicious
'                    If token.tokVariable Is Nothing Then operand_stack(operand_stack.Count).tokOutput = "(" & token.tokLocalFunction.procname & " *)" & operand_stack(operand_stack.Count).tokOutput
                    token.tokOutput = operand_stack.Item(operand_stack.Count - token.tokCount).tokOutput ' This
                    operand_stack.Remove operand_stack.Count - token.tokCount
'                    token.tokOutput = "___" & token.tokLocalFunction.procLocalModule.Name & "_" & ProcNameIK(token.tokLocalFunction.procName, IIf(token.tokLocalFunction.procattributes And PROC_ATTR_VARIABLE, token.tokPCodeSubType And (INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF), INVOKE_FUNC)) & EmitVariable(token, operand_stack, token.tokLocalFunction.procLocalModule.Component.Type)
' intending "And Not INVOKE_FUNC" to force output of "Get"
'                    ik = token.tokPCodeSubType And Not IIf(token.tokVariable.MemberType = vbext_mt_Variable, 1, 0)
                    ' fixme: c.aaa = c.aaa - need to switch from PROPERTYPUT to PROPERTYGET. Need to switch according to InvokeKind in vbt.
                    If ik = INVOKE_PROPERTYPUT And token.tokPCodeSubType = INVOKE_PROPERTYGET Then ik = INVOKE_PROPERTYGET
                    If token.tokPCodeSubType = 0 Then ik = ik And (INVOKE_FUNC Or INVOKE_PROPERTYGET) Else ik = ik And token.tokPCodeSubType
                    token.tokOutput = "___" & token.tokLocalFunction.procLocalModule.Name & "_" & ProcNameIK(token.tokLocalFunction.procname, ik) & EmitVariable(token, operand_stack, token.tokLocalFunction.procLocalModule.Component.Type, ParamArrayCnt)
'                    token.tokOutput = "___" & token.tokLocalFunction.procLocalModule.Name & "_" & ProcNameIK(token.tokLocalFunction.procName, INVOKE_PROPERTYGET) & "(" & operand_stack(operand_stack.Count).tokOutput & token.tokOutput & ")"
    ' property1.vbp
'    If ik And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
'        Print #99, "osc="; operand_stack.Count
'        token.tokOutput = token.tokOutput & "," & operand_stack.Item(operand_stack.Count).tokOutput
'        operand_stack.Remove operand_stack.Count
'    End If
'                Case Else
'                    Print #99, "EmitInFix: Unknown ComponentType: " & token.tokLocalFunction.procLocalModule.component.type
'                    MsgBox "EmitInFix: Unknown ComponentType: " & token.tokLocalFunction.procLocalModule.component.type
'                    Err.Raise 1 ' Module has unknown ComponentType
            End Select
    ProcessSubscripts token, operand_stack
    operand_stack.Add token
Case tokReferenceClass, tokFormClass
'        ElseIf Not token.tokInterfaceInfo Is Nothing Then
Print #99, "tokReferenceClass: ii="; Not token.tokInterfaceInfo Is Nothing; " mi="; Not token.tokMemberInfo Is Nothing
            If Not token.tokMemberInfo Is Nothing Then
            Dim ii As InterfaceInfo
            ' This should be the VTableInterface version
            Set ii = token.tokInterfaceInfo
' use this???            set ii = ii.VTableInterface
Print #99, " interface="; ii.Name; " dt="; token.tokDataType
                Dim mi As MemberInfo
'Set mi = token.tokMemberInfo
'Print #99, " memberinfo="; mi.Name; " ik="; mi.InvokeKind; " param.count="; mi.Parameters.Count; " voff="; mi.VTableOffset
'On Error GoTo 11
'For Each mi In token.tokVariable.VarType.dtInterfaceInfo.VTableInterface.Members
'    If mi.Name = token.tokMemberInfo.Name Then
'        Print #99, " memberinfo="; mi.Name; " ik="; mi.InvokeKind; " param.count="; mi.Parameters.Count; " voff="; mi.VTableOffset
'        Exit For
'    End If
'Next
'11
'On Error GoTo 0
                Set mi = token.tokMemberInfo
Print #99, "memberinfo="; mi.Name; " ik="; mi.InvokeKind; " dt="; mi.ReturnType.VarType; " st="; token.tokPCodeSubType; " pc="; mi.Parameters.Count; " tc="; token.tokCount
                Dim mik As InvokeKinds
'                mik = IIf(mi.InvokeKind = 0, token.tokPCodeSubType, mi.InvokeKind) ' dispinterface
'                mik = IIf(token.tokPCodeSubType, token.tokPCodeSubType, mi.InvokeKind) ' dispinterface
                ' need to elimiate pcodesubtype and fully implement invokekind
                If token.tokPCodeSubType = 0 Then
                    mik = mi.InvokeKind
                ElseIf mi.InvokeKind = 0 Then ' won't work for INVOKE_UNKNOWN?
                    mik = token.tokPCodeSubType
                Else
                    mik = mi.InvokeKind And token.tokPCodeSubType
                    ' fixme: bmpinfo.vbp - default changed invokekind from put to get
                    If mik = 0 Then mik = token.tokPCodeSubType
                End If
                Print #99, "mik="; mik
'                If Not mik And token.tokPCodeSubType Then Err.Raise 1
                'Dim ss As String
                ss = ""
                i = operand_stack.Count - token.tokCount + 1
                'Dim k As Integer
                k = 0
                'Dim pi As ParameterInfo
                For Each pi In mi.Parameters
Print #99, "p2 n="; pi.Name; " f="; Hex(pi.Flags); " i="; i; " pl="; pi.VarTypeInfo.PointerLevel; " tc="; token.tokCount
                    k = k + 1
'                    If k = mi.Parameters.Count And (mik And (INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) Then Exit For
                    If k = mi.Parameters.Count And (mik And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) And (mi.ReturnType.VarType = VT_VOID Or mi.ReturnType.VarType = VT_HRESULT) Then Exit For
'                    If k = mi.Parameters.Count And (mik And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) Then Exit For
                    If pi.Flags And PARAMFLAG_FRETVAL Then Exit For ' Function
                    If k <= token.tokCount Then
                        'Dim dt As TliVarType
                        s = operand_stack.Item(i).tokOutput
                        operand_stack.Remove i
                    ElseIf pi.Flags And PARAMFLAG_FLCID Then ' LCID
                        s = "0" ' not implemented - just default
                    ElseIf pi.Optional Or CBool(pi.Flags And PARAMFLAG_FOPT) Then
                        If pi.Optional Xor CBool(pi.Flags And PARAMFLAG_FOPT) Then Print #99, "Funky optional problem"
                        If pi.Flags And PARAMFLAG_FHASDEFAULT Then
                            s = CDefaultValueByType(VarType(pi.DefaultValue))
                        Else
                            s = CDefaultValueByType(vbVariant)
                        End If
                    Else
                        Print #99, "Expecting optional parameter"
                        Err.Raise 1 ' Expecting optional parameter
                    End If
                    ss = ss & "," & s
                Next
'                If mi.InvokeKind = INVOKE_UNKNOWN Then
'                s = ProcNameIK(mi.Name, mik) & "(" & operand_stack.Item(operand_stack.Count).tokOutput ' Me
'                s = ProcNameIK(mi.Name, mi.InvokeKind) & "(" & operand_stack.Item(operand_stack.Count).tokOutput ' Me
'                Print #99, " ii="; token.tokVariable.VarType.dtInterfaceInfo Is Nothing
                Print #99, " ci="; Not token.tokVariable.VarType.dtClassInfo Is Nothing
                If token.tokVariable.VarType.dtClassInfo Is Nothing Then GoTo Class_This
'                Print #99, " ii="; token.tokInterfaceInfo.Name; " am="; token.tokInterfaceInfo.AttributeMask; " ref="; Not token.tokReference Is Nothing
                ' AppObject is an attribute of CoClasses, not Interfaces or Members
                If token.tokVariable.VarType.dtClassInfo.AttributeMask And TYPEFLAG_FAPPOBJECT Then
                    ' AppObjects TLibs are references are explicitly generated "VB_VBGlobal *VB;"
                    ' Compiler could emit an "AppObject *This" token instead
                    ' On second thought, should emit VB_Global_QI(This)
' bmpinfo.vbp
                    s = ProcNameIK(mi.Name, mi.InvokeKind) & "(" & TypeInfoToCType(token.tokInterfaceInfo) & "_AppObject"
                Else
Class_This:
' bmpinfo.vbp
                    s = ProcNameIK(mi.Name, mi.InvokeKind) & "(" & operand_stack.Item(operand_stack.Count).tokOutput ' Me
                    operand_stack.Remove operand_stack.Count
                End If
'                 End If
                Print #99, "mik="; mik
                Print #99, "s="; s
                Print #99, "ss="; ss
                If mik And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
                    Print #99, "osc="; operand_stack.Count
                    ss = ss & "," & operand_stack.Item(operand_stack.Count).tokOutput
                    operand_stack.Remove operand_stack.Count
                End If
'                If mi.InvokeKind <> INVOKE_UNKNOWN Then
'                    s = operand_stack.Item(operand_stack.Count).tokOutput ' Me
'                    operand_stack.Remove operand_stack.Count
'                End If
'                token.tokOutput = "_" & ProcNameIK(mi.name,mi.invokekind)
'                token.tokOutput = TLIRefMemberName(ii, mi, mi.InvokeKind)
Print #99, "ss="; ss
                s = s & ss
            Else
                ' fixme: Form1.SomeControl.Enabled = True - SomeControl doesn't have a MemberInfo, perhaps should, process as tokProjectClass
                If Not token.tokLocalFunction Is Nothing Then GoTo projectclass
' Members not in interface.members collection but interface is extensible (Err.xxxx)
                Print #99, "Member "; token.tokOutput; " is not in interface.members"
                GoTo generic_invoke1
#If 0 Then
Print #99, "cnt="; token.tokCount; " ii="; ii.Name
                token.tokOutput = operand_stack.Item(operand_stack.Count).tokOutput & token.tokOutput
                operand_stack.Remove operand_stack.Count
                s = ProcNameIK(token.tokOutput, token.tokPCodeSubType) & EmitVariable(token, operand_stack) ' Me
#End If
            End If
Print #99, "s="; s
            Print #99, " interface name="; ii.Name
            ' create routine to catenate name (also change ctypename)
            token.tokOutput = "_i_VB_" & ii.Name & "_" & s & ")"
            On Error Resume Next ' Components don't have parent!??
            Print #99, " ii.parent="; ii.Parent Is Nothing
            Print #99, " tlib name="; ii.Parent.Name
            token.tokOutput = TypeInfoToCType(ii) & "_" & s & ")"
            On Error GoTo 0
Print #99, "tokReferenceClass: output="; token.tokOutput
        operand_stack.Add token
    Case tokInvokeDefaultMember
        s = "NULL"
        GoTo generic_invoke2
    Case tokInvoke
    ' fixme: This should do what tokIDispatch does
    Case tokVariantArgs
        If (token.tokDataType And Not VT_BYREF) <> vbVariant Then Err.Raise 1
        If token.tokCount = 0 Then Err.Raise 1
        s = ""
        For i = 1 To token.tokCount
            s = "," & operand_stack.Item(operand_stack.Count).tokOutput & s
            operand_stack.Remove operand_stack.Count
        Next
        s = ",VarArg(" & token.tokCount & "," & Mid(s, 2) & ")"
        Print #99, "s="; s
        ' fixme: Should VariantArgs take a SAFEARRAY argument?
' don't think this is needed here, its used in EmitVariable
'        If token.tokVariable.varProc Is Nothing Then token.tokOutput = token.tokVariable.varModule.Name & "_" & token.tokOutput
        token.tokOutput = EmitVariable(token, operand_stack, -1, -1)
        ProcessSubscripts token, operand_stack
        Print #99, "to="; token.tokOutput
        If token.tokPCodeSubType And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
            s = s & "," & operand_stack.Item(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
        End If
        Print #99, "s="; s
        token.tokOutput = AbbrDataType(token.tokDataType) & ProcNameIK("Args", token.tokPCodeSubType) & "(" & token.tokOutput & s & ")"
        operand_stack.Add token
    Case tokIDispatchInterface
generic_invoke1:
            s = "L""" & token.tokOutput & """"
generic_invoke2:
Print #99, " generic Object: o="; token.tokOutput; " pcst="; token.tokPCodeSubType; " c="; token.tokCount; " r="; token.tokRank; " v="; token.tokValue
            ss = ""
            
' fixme: not right. need to double rank to get correct offset for ReDim variables.
' fixme: ReDim dimensions are being output before variable.
        rankOffset = token.tokRank
        If token.tokPCode = tokReDim Then rankOffset = rankOffset * 2
            j = token.tokCount + rankOffset
            For i = 1 To j
                ss = "," & operand_stack.Item(operand_stack.Count).tokOutput & ss
                operand_stack.Remove operand_stack.Count
            Next
'            ss = EmitVariable(token, operand_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET)
            'Dim sss As String
'            If operand_stack.Item(operand_stack.Count).tokInterfaceInfo.GUID <> tlistdole.GUID Then
'                err.Raise 1
'            End If
            sss = AbbrDataType(operand_stack.Item(operand_stack.Count).tokDataType) & "Invoke(" & operand_stack.Item(operand_stack.Count).tokOutput & ","
            operand_stack.Remove operand_stack.Count
            If token.tokPCodeSubType And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
                ss = ss & "," & operand_stack.Item(operand_stack.Count).tokOutput
                operand_stack.Remove operand_stack.Count
                j = j + 1
            End If
'            ss = Mid(ss, 2, Len(ss) - 2) ' remove parenthesis
Print #99, "s="; s; " ss="; ss; " sss="; sss; " tc="; token.tokCount
            token.tokOutput = sss & s & "," & token.tokPCodeSubType & "," & j & ss & ")"
Print #99, "generic_invoke: output="; token.tokOutput
            operand_stack.Add token
'    Case tokArrayVariable
'        token.tokOutput = IIf(token.tokVariable.varProc Is Nothing, token.tokVariable.varModule.Name & "_", "") & EmitVariable(token, operand_stack, -1)
'        If token.tokVariable.varAttributes And VARIABLE_NEW Then token.tokOutput = "GetObj(" & token.tokOutput & ")"
'        operand_stack.Add token
    Case tokLabelRef
' fixme: outputing labels prefixed by "_" to avoid name conflicts. keep?
#If 0 Then
        If isdigit(Left(token.tokString, 1)) Then token.tokOutput = "_" & token.tokOutput
#Else
        token.tokOutput = "_" & token.tokOutput
#End If
        operand_stack.Add token
    Case tokNothing
        token.tokOutput = AbbrDataType(token.tokDataType) & "Nothing"
        operand_stack.Add token
    Case tokVariant
'        If VarType(token.tokValue) <> vbObject Then token.tokOutput = ValueToC(token.tokDataType, token.tokValue)
        token.tokOutput = ValueToC(token.tokDataType, token.tokValue)
        operand_stack.Add token
    Case tokConst
        token.tokOutput = ValueToC(token.tokDataType, token.tokValue)
'        token.tokOutput = "_c_" & token.tokVariable.varModule.Name & "_" & token.tokVariable.varSymbol
        operand_stack.Add token
    Case tokEnumMember
        token.tokOutput = ValueToC(token.tokDataType, token.tokValue)
'        token.tokOutput = "_e_" & token.tokEnumMember.enumMemberParent.enumModule.Name & "_" & token.tokEnumMember.enumMemberParent.enumName & "_" & token.tokEnumMember.enumMemberName
        operand_stack.Add token
    Case tokConstantInfo
        token.tokOutput = TypeInfoToCType(token.tokVariable.VarType.dtConstantInfo) & "_" & token.tokMemberInfo.Name
        operand_stack.Add token
    Case tokstatement
        EmitStatement token, operand_stack
    Case tokLabelDef
        operand_stack.Add token
#If 0 Then
        If isdigit(Left(token.tokString, 1)) Then token.tokOutput = "_" & token.tokOutput
#Else
        token.tokOutput = "_" & token.tokOutput
#End If
        token.tokOutput = token.tokOutput & ":;" ' ; needed in VC++ before }
    Case tokNewObject
        Select Case token.tokVariable.VarType.dtType
            Case tokProjectClass, tokFormClass
                token.tokOutput = "_i_" & token.tokVariable.VarType.dtClass.Name & "_New"
            Case tokReferenceClass
                token.tokOutput = TypeInfoToCType(token.tokVariable.VarType.dtClassInfo) & "_New"
            Case tokIDispatchInterface
                token.tokOutput = "StdOle_IDispatch_New"
            Case Else
                Err.Raise 1
        End Select
        operand_stack.Add token
    Case tokTypeOf
        Print #99, "tokTypeOf: v="; Not token.tokVariable Is Nothing
        If token.tokVariable Is Nothing Then Err.Raise 1
        Print #99, "dt="; token.tokVariable.VarType.dtDataType
        Select Case token.tokVariable.VarType.dtDataType
            Case vbObject
                Select Case token.tokVariable.VarType.dtType
                    Case tokProjectClass, tokFormClass
                        token.tokOutput = token.tokVariable.VarType.dtClass.interfaceGUID
                    Case tokReferenceClass
                        token.tokOutput = token.tokVariable.VarType.dtInterfaceInfo.GUID
                    Case tokIDispatchInterface
                        token.tokOutput = "" ' fixme: need GUID or NULL?
                    Case Else
                        Err.Raise 1
                End Select
            Case vbUserDefinedType
                Select Case token.tokVariable.VarType.dtType
                    Case tokProjectClass, tokFormClass
                        token.tokOutput = token.tokVariable.VarType.dtUDT.typeGUID
                    Case tokReferenceClass
                        token.tokOutput = token.tokVariable.VarType.dtRecordInfo.GUID
                    Case tokIDispatchInterface
                        token.tokOutput = "" ' fixme: need GUID or NULL?
                    Case Else
                        Err.Raise 1
                End Select
            Case Else
                Err.Raise 1
        End Select
        Print #99, " s="; operand_stack.Item(operand_stack.Count).tokOutput; " dt="; operand_stack.Item(operand_stack.Count).tokDataType
        token.tokOutput = AbbrDataType(operand_stack.Item(operand_stack.Count).tokDataType) & "TypeOf(" & operand_stack.Item(operand_stack.Count).tokOutput & "," & IIf(token.tokOutput = "", "NULL", "L""" & token.tokOutput & """") & ")"
#If 0 Then ' UDT ref change
        If operand_stack.Item(operand_stack.Count).tokDataType = vbUserDefinedType Then
            If Left(operand_stack.Item(operand_stack.Count).tokOutput, 2) <> "(*" Then Err.Raise 1
            operand_stack.Item(operand_stack.Count).tokOutput = Mid(operand_stack.Item(operand_stack.Count).tokOutput, 3, Len(operand_stack.Item(operand_stack.Count).tokOutput) - 3)
        End If
#End If
#If 0 Then
        Print #99, "v="; Not token.tokVariable Is Nothing
        If token.tokVariable Is Nothing Then Err.Raise 1
        Print #99, "type="; token.tokVariable.VarType.dtType
        Select Case token.tokVariable.VarType.dtType
        Case tokIDispatchInterface
            token.tokOutput = AbbrDataType(operand_stack.Item(operand_stack.Count).tokDataType) & "TypeOf(" & operand_stack.Item(operand_stack.Count).tokOutput & ",L""" & "?" & """)"
        Case tokProjectClass, tokFormClass
            Print #99, "class="; Not token.tokVariable.VarType.dtClass Is Nothing
            token.tokOutput = AbbrDataType(operand_stack.Item(operand_stack.Count).tokDataType) & "TypeOf(" & operand_stack.Item(operand_stack.Count).tokOutput & ",L""" & token.tokVariable.VarType.dtClass.interfaceGUID & """)"
        Case tokReferenceClass
            token.tokOutput = AbbrDataType(operand_stack.Item(operand_stack.Count).tokDataType) & "TypeOf(" & operand_stack.Item(operand_stack.Count).tokOutput & ",L""" & token.tokVariable.VarType.dtInterfaceInfo.GUID & """)"
        Case Else
            Err.Raise 1
        End Select
#End If
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case tokCvt ' Internal data type conversions
        If (token.tokPCodeSubType And Not VT_BYREF) <> (operand_stack.Item(operand_stack.Count).tokDataType And Not VT_BYREF) Then Err.Raise 1
        token.tokOutput = PCodeToSymbol(token) & "(" & operand_stack.Item(operand_stack.Count).tokOutput & ")"
'        operand_stack(operand_stack.Count).tokDataType = token.tokDataType
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
'        Set operand_stack.Item(operand_stack.Count) = token ' replace top item
    Case tokAddRef
' fixme: rewrite when PointerLevel is implemented!!
        Print #99, "tt="; operand_stack(operand_stack.Count).tokType
        Select Case operand_stack(operand_stack.Count).tokType
            Case tokVariable, tokArrayVariable
                ' this test is a bit kludgy, ideally we should examine pointer ref value
                If operand_stack(operand_stack.Count).tokVariable.varDimensions Is Nothing Then
                    If operand_stack(operand_stack.Count).tokVariable.varAttributes And VT_BYREF Then
                        If Left(operand_stack(operand_stack.Count).tokOutput, 1) = "*" Then
                            token.tokOutput = Mid(operand_stack(operand_stack.Count).tokOutput, 2)
                        End If
                    Else
                        If Left(operand_stack(operand_stack.Count).tokOutput, 1) = "*" Then Err.Raise 1
                        If operand_stack(operand_stack.Count).tokVariable.varAttributes And VARIABLE_NEW Then
                            token.tokOutput = AbbrDataTypeRef(operand_stack(operand_stack.Count).tokDataType, operand_stack(operand_stack.Count).tokOutput)
                        Else
                            token.tokOutput = "&" & operand_stack(operand_stack.Count).tokOutput
                        End If
                    End If
                ElseIf operand_stack(operand_stack.Count).tokRank = 0 Then
                    token.tokOutput = operand_stack(operand_stack.Count).tokOutput
                Else
                    If Left(operand_stack(operand_stack.Count).tokOutput, 1) = "*" Then
                        token.tokOutput = Mid(operand_stack(operand_stack.Count).tokOutput, 2)
                    Else
                        token.tokOutput = AbbrDataTypeRef(operand_stack(operand_stack.Count).tokDataType, operand_stack(operand_stack.Count).tokOutput)
                    End If
                End If
            Case tokVariant
                token.tokOutput = AbbrDataTypeRef(operand_stack(operand_stack.Count).tokDataType, operand_stack(operand_stack.Count).tokOutput)
            Case Else ' member
                token.tokOutput = AbbrDataTypeRef(operand_stack(operand_stack.Count).tokDataType, operand_stack(operand_stack.Count).tokOutput)
'                Print #99, "EmitInFix: Unknown ByRef type=" & operand_stack(operand_stack.Count).tokType & " dt=" & operand_stack(operand_stack.Count).tokDataType
'                MsgBox "EmitInFix: Unknown ByRef type=" & operand_stack(operand_stack.Count).tokType & " dt=" & operand_stack(operand_stack.Count).tokDataType
'                Err.Raise 1 ' compiler error
        End Select
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case tokSubRef
' fixme: rewrite when PointerLevel is implemented!!
        token.tokOutput = "(*" & operand_stack(operand_stack.Count).tokOutput & ")"
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case tokByVal ' make argument an expression, if not already
        token.tokOutput = "(" & operand_stack(operand_stack.Count).tokOutput & ")"
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case tokme
        Select Case token.tokVariable.VarType.dtClass.Component.Type
        Case vbext_ct_StdModule
            ' do nothing
        Case Else
            token.tokOutput = "This"
            operand_stack.Add token
        End Select
    Case tokCase
        operand_stack(operand_stack.Count).tokOutput = AbbrDataType(operand_stack(operand_stack.Count).tokDataType) & "Case(" & operand_stack(operand_stack.Count).tokOutput & ")"
    Case tokCaseIs
        operand_stack(operand_stack.Count).tokOutput = AbbrDataType(operand_stack(operand_stack.Count).tokDataType) & "CaseIs" & PCodeToAbbr(token.tokPCode) & "(" & operand_stack(operand_stack.Count).tokOutput & ")"
    Case tokCaseTo
        operand_stack(operand_stack.Count - 1).tokOutput = AbbrDataType(operand_stack(operand_stack.Count).tokDataType) & "CaseTo(" & operand_stack(operand_stack.Count).tokOutput & "," & operand_stack(operand_stack.Count - 1).tokOutput & ")"
        operand_stack.Remove operand_stack.Count
    Case tokLBound, tokUBound
        If (operand_stack(operand_stack.Count - 1).tokDataType And Not VT_BYREF) <> vbVariant And Not CBool(operand_stack(operand_stack.Count - 1).tokDataType And (VT_ARRAY Or VT_VECTOR)) Then Err.Raise 1
'        token.tokOutput = IIf((operand_stack(operand_stack.Count - 1).tokDataType And Not VT_BYREF) = vbVariant, "Var", "Sa") & IIf(token.tokType = tokLBound, "L", "U") & "Bound(" & operand_stack(operand_stack.Count - 1).tokOutput & "," & operand_stack(operand_stack.Count).tokOutput & ")"
        token.tokOutput = AbbrDataType(operand_stack(operand_stack.Count - 1).tokDataType) & IIf(token.tokType = tokLBound, "L", "U") & "Bound(" & operand_stack(operand_stack.Count - 1).tokOutput & "," & operand_stack(operand_stack.Count).tokOutput & ")"
        operand_stack.Remove operand_stack.Count
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case tokOperands ' Generic arguments for Spc, Tab, ;, ,
' Hmmm, seems funky
        s = ""
        For i = 1 To token.tokCount
            s = "," & operand_stack.Item(operand_stack.Count).tokOutput & s
            operand_stack.Remove operand_stack.Count
        Next
        token.tokOutput = Mid(s, 2)
        operand_stack.Add token
    Case tokDeclare
        token.tokOutput = ""
'        token.tokOutput = token.tokDeclare.dclModule.Name & "_" & IIf(token.tokDeclare.dclAlias = "", token.tokDeclare.dclName, token.tokDeclare.dclAlias) & EmitVariable(token, operand_stack)
        Print #99, "tokDeclare: pc="; token.tokDeclare.dclParams.Count; " opc="; token.tokDeclare.dclOptionalParams
        If token.tokDeclare.dclOptionalParams = -1 Then ' ParamArray
            token.tokOutput = token.tokDeclare.dclModule.Name & "_" & token.tokDeclare.dclName & EmitVariable(token, operand_stack, , token.tokCount - token.tokDeclare.dclParams.Count + 1)
        Else
            token.tokOutput = token.tokDeclare.dclModule.Name & "_" & token.tokDeclare.dclName & EmitVariable(token, operand_stack)
        End If
        operand_stack.Add token
    Case tokStdProcedure
        If token.tokDataType = vbLong Then
            token.tokOutput = "(long)"
        ElseIf token.tokDataType = vbLong Or VT_BYREF Then
            token.tokOutput = "(long *)"
        Else
            Err.Raise 1 ' Expecting Long
        End If
        ' don't use "_i_" - midiin1a\midiecho.vbp
        token.tokOutput = token.tokOutput & token.tokModule.Name & "_" & token.tokLocalFunction.procname
        operand_stack.Add token
    Case tokAddressOf
        ' don't use "_i_" - midiin1a\midiecho.vbp
        operand_stack.Item(operand_stack.Count).tokOutput = "(long)" & operand_stack.Item(operand_stack.Count).tokModule.Name & "_" & operand_stack.Item(operand_stack.Count).tokOutput
    Case tokMissing
        ' For missing ParamArray, may be beter to output VarArg instead of VarSAMissing
        token.tokOutput = "_" & AbbrDataType(token.tokDataType) & "Missing"
        operand_stack.Add token
    Case Else
        Print #99, "EmitInFix: Unknown pcode: " & token.tokType
        MsgBox "EmitInFix: Unknown pcode: " & token.tokType
        Err.Raise 1
End Select
If operand_stack.Count > 0 Then Print #99, "s="; operand_stack.Item(operand_stack.Count).tokOutput; " dt="; operand_stack.Item(operand_stack.Count).tokDataType
Next
If operand_stack.Count <> 1 Then
    Print #99, "EmitInFix: operand count <> 1: count=" & operand_stack.Count
    MsgBox "EmitInFix: operand count <> 1: count=" & operand_stack.Count
    Err.Raise 1 ' compiler error
End If
EmitInFix = operand_stack.Item(1).tokOutput
If operand_stack.Item(1).tokType <> tokLabelDef Then EmitInFix = indent & EmitInFix
Print #99, "EmitInFix: {"; operand_stack.Item(1).tokOutput; "}"
indent = newindent
End Function

' fixme: klugy - confusing component type and member type
Function EmitVariable(ByVal token As vbToken, ByVal operand_stack As Collection, Optional ByVal ComponentType As vbext_ComponentType = vbext_ct_StdModule, Optional ByVal ParamArrayCnt As Integer = 0) As String
Print #99, "EmitVariable: tv="; Not token.tokVariable Is Nothing; " tc="; token.tokCount; " r="; token.tokRank; " ct="; ComponentType; " pst="; token.tokPCodeSubType; " dt="; token.tokDataType; " osc="; operand_stack.Count; " pac="; ParamArrayCnt
Dim s As String
Dim i As Long
Print #99, 1
If ParamArrayCnt > 0 Then
    For i = 1 To ParamArrayCnt
        s = "," & operand_stack.Item(operand_stack.Count).tokOutput & s
        operand_stack.Remove operand_stack.Count
    Next
    s = ",VarArg(" & ParamArrayCnt & "," & Mid(s, 2) & ")"
End If
Print #99, "2 s="; s
For i = 1 To token.tokCount - IIf(ParamArrayCnt = -1, token.tokCount, ParamArrayCnt)
    s = "," & operand_stack.Item(operand_stack.Count).tokOutput & s
    operand_stack.Remove operand_stack.Count
Next
Print #99, "3 s="; s
Select Case ComponentType
    ' fixme: replace with IsForm() function?
    Case vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_UserControl, vbext_ct_ActiveXDesigner, vbext_ct_PropPage
' duped code
Print #99, "4 s="; s
        If token.tokPCodeSubType = INVOKE_PROPERTYPUT Then
            s = s & "," & operand_stack.Item(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
        ElseIf token.tokPCodeSubType = INVOKE_PROPERTYPUTREF Then
            s = s & "," & operand_stack.Item(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
        End If
        s = "(" & token.tokOutput & s & ")"
Print #99, 5
    Case vbext_ct_StdModule
        If token.tokPCodeSubType = INVOKE_PROPERTYPUT Then
            s = s & "," & operand_stack.Item(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
        ElseIf token.tokPCodeSubType = INVOKE_PROPERTYPUTREF Then
            s = s & "," & operand_stack.Item(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
        End If
        s = "(" & Mid(s, 2) & ")"
    Case Else
Print #99, "6 s="; s
        If token.tokVariable.varAttributes And VARIABLE_PARAMETER Then
'            If CBool(token.tokVariable.varAttributes And VARIABLE_BYREF) And token.tokVariable.varDimensions Is Nothing Then
'        Select Case token.tokVariable.VarType.dtDataType
'            Case VT_BSTR ' do not deref string params (char arrays)
'            Case Else
            If s <> "" Then Err.Raise 1
            s = "_p_" & token.tokOutput
'        End Select
        ElseIf token.tokVariable.varAttributes And VARIABLE_FUNCTION Then
' probably need to do "ElseIf token.toklocalfunction is proc Then"
            If s <> "" Then Err.Raise 1
            If token.tokVariable.varModule.Component.Type = vbext_ct_StdModule Then
                s = "lv._v_" & token.tokOutput
            Else
                s = "(*_p_retval)"
            End If
'        ElseIf Not token.tokVariable.VarType.dtClass Is Nothing Then
'        ElseIf token.tokVariable.varAttributes And VARIABLE_ME Then
'            If s <> "" Then Err.Raise 1
'            s = "This"
'            ElseIf token.tokVariable.varProc Is Nothing Then
''            s = "gv." & s
'            Else
'                s = "lv." & s ' Project class var
'            End If
'        ElseIf Not token.tokVariable.VarType.dtClassInfo Is Nothing Then  ' TLib class var
'            s = "lv." & s
        ElseIf token.tokVariable.varProc Is Nothing Then
            If s <> "" Then Err.Raise 1
            ' module defined variable, except UDT member which has module name omitted.
            If token.tokType = tokUDT Then
                s = token.tokOutput & s
            Else
                s = "_v_" & token.tokVariable.varModule.Name & "_" & token.tokOutput & s
            End If
        Else
            s = "lv._v_" & token.tokOutput & s
        End If
'            If token.tokVariable.varAttributes And VARIABLE_NEW Then s = "New(" & s & ")"
'' fixme: Attempting to specify Currency union member of .int64. Need testing and debuging.
'        If token.tokDataType = vbCurrency Then s = s & ".int64"
#If 0 Then
        If token.tokDataType And VT_BYREF Then
            token.tokDataType = token.tokDataType And Not VT_BYREF
            s = "(*" & s & ")"
        End If
#End If
#If 0 Then ' UDT ref change
        If (token.tokDataType And Not VT_ARRAY) = VT_RECORD Then s = "(*_t_" & s & ")"
#End If
End Select
EmitVariable = s
Print #99, "EmitVariable: ev="; EmitVariable
End Function

Sub ProcessSubscripts(ByVal token As vbToken, ByVal operand_stack As Collection)
Dim ss As String
Dim i As Long
Print #99, "ProcessSubscripts: ts="; token.tokString; " r="; token.tokRank; " dt="; token.tokDataType; " vt="; Not token.tokVariable Is Nothing; " osc="; operand_stack.Count
If token.tokPCode = tokReDim Then ' overloading tokPCode
    token.tokOutput = token.tokOutput & "," & VT_Type(token.tokVariable) & "," & token.tokRank
    For i = (token.tokRank - 1) * 2 To 0 Step -2
        token.tokOutput = token.tokOutput & "," & operand_stack.Item(operand_stack.Count - i - 1).tokOutput & "," & operand_stack.Item(operand_stack.Count - i).tokOutput
        operand_stack.Remove operand_stack.Count - i
        operand_stack.Remove operand_stack.Count - i
    Next
ElseIf token.tokRank > 0 Then
'If token.tokOutput <> "" And Not token.tokVariable Is Nothing Then
'    Print #99, "processSubscripts: tvd="; token.tokVariable.varDimensions Is Nothing
'    If Not token.tokVariable.varDimensions Is Nothing And token.tokRank > 0 Then
'        Print #99, "EmitVariable: tvdc="; token.tokVariable.varDimensions.Count
'        If token.tokVariable.varDimensions.Count = 0 Then
        For i = 1 To token.tokRank
            ss = "," & operand_stack.Item(operand_stack.Count).tokOutput & ss
            operand_stack.Remove operand_stack.Count
        Next
        ' fixme - change when PointerLevel implemented
        Select Case token.tokDataType
        Case VT_RECORD
            ' UDTs must be casted
            Print #99, "processsubscripts: UDT type="; token.tokVariable.VarType.dtType
            Select Case token.tokVariable.VarType.dtType
                Case tokProjectClass, tokFormClass
                    Print #99, "udt exists="; Not token.tokVariable.VarType.dtUDT Is Nothing
'' fixme: generate SA protos and use instead of casts, as is done for RecordInfos
'                    token.tokOutput = "*((_t_" & token.tokVariable.VarType.dtudt.typemodule.Name & "_" & token.tokVariable.VarType.dtUDT.TypeName & "*)" & AbbrDataType(token.tokDataType) & "SA(" & token.tokOutput & "," & token.tokRank & ss & "))"
                    token.tokOutput = "*_t_" & token.tokVariable.VarType.dtUDT.typeModule.Name & "_" & token.tokVariable.VarType.dtUDT.TypeName & "SA(" & token.tokOutput & "," & token.tokRank & ss & ")"
                Case tokReferenceClass
                    Print #99, "ri exists="; Not token.tokVariable.VarType.dtRecordInfo Is Nothing
'                    token.tokOutput = "*((" & TypeInfoToCType(token.tokVariable.VarType.dtRecordInfo) & "*)" & AbbrDataType(token.tokDataType) & "SA(" & token.tokOutput & "," & token.tokRank & ss & "))"
                    token.tokOutput = "*" & TypeInfoToCType(token.tokVariable.VarType.dtRecordInfo) & "SA(" & token.tokOutput & "," & token.tokRank & ss & ")"
                Case Else
                    Err.Raise 1
            End Select
        Case VT_DISPATCH, VT_UNKNOWN
            Print #99, "processsubscripts: Obj type="; token.tokVariable.VarType.dtType
            Select Case token.tokVariable.VarType.dtType
                Case tokProjectClass, tokFormClass
                    token.tokOutput = "*_i_" & token.tokVariable.VarType.dtClass.Name & "SA(" & token.tokOutput & "," & token.tokRank & ss & ")"
                Case tokReferenceClass
                    token.tokOutput = "*" & TypeInfoToCType(token.tokVariable.VarType.dtInterfaceInfo) & "SA(" & token.tokOutput & "," & token.tokRank & ss & ")"
                Case Else
                    Err.Raise 1
            End Select
        Case Else
            token.tokOutput = "*SaTo" & AbbrDataType(token.tokDataType) & "(" & token.tokOutput & "," & token.tokRank & ss & ")"
        End Select
'        Else
'            For i = 1 To token.tokRank
'                ss = "[" & operand_stack.Item(operand_stack.Count).tokOutput & "]" & ss
'                operand_stack.Remove operand_stack.Count
'            Next
'            s = s & ss
'            End If
'    End If
'End If
End If
Print #99, "ProcessSubscripts: osc="; operand_stack.Count
End Sub

Sub EmitStatement(ByVal token As vbToken, ByVal operand_stack As Collection)
Dim dt As TliVarType
Print #99, token.tokOutput; " t="; token.tokType; " dt="; token.tokDataType; " pc="; token.tokPCode; " pcst="; token.tokPCodeSubType; " oc="; operand_stack.Count
Select Case token.tokPCode
    Case vbPCodeLet, vbPCodeLSet, vbPCodeRSet
' fixme - do more of this data type checking for other pcodes
' maybe make a datatype check routine?
' fixme - may be cleaner to use token/remove instead of reassigning to output_stack
' fixme: rewrite PropertyLet processing. Shouldn't be another Case statement.
If operand_stack.Count = 1 Then token.tokPCode = vbPCodePropertyLet: GoTo PropertyLetSet
'If (operand_stack(operand_stack.Count).tokType <> tokVariable And operand_stack(operand_stack.Count).tokType <> tokArrayVariable) Then GoTo PropertyLetSet
'If operand_stack(operand_stack.Count).tokVariable.Component.Type <> vbext_ct_StdModule Then GoTo PropertyLetSet
        ' fixme: change vbt to recognize array assignment, so RHS isn't derefed.
        ' fixme: change vbt to emit correct RHS/LHS/token VT_ARRAY flag
        ' getfil2r\datasmoo.vbp
        If VT_ARRAY And operand_stack.Item(1).tokDataType And operand_stack.Item(2).tokDataType Then
''''            operand_stack.Item(2).tokOutput = AbbrDataType(token.tokDataType) & "Array" & token.tokOutput & "(" & operand_stack.Item(1).tokOutput & "," & operand_stack.Item(2).tokOutput & ")"
            operand_stack.Item(2).tokOutput = AbbrDataType(token.tokDataType) & token.tokOutput & "(" & operand_stack.Item(1).tokOutput & "," & operand_stack.Item(2).tokOutput & ")"
            GoTo 1
        End If
Print #99, "let: t.dt="; token.tokDataType
        dt = operand_stack.Item(2).tokDataType
Print #99, "let: lhs.dt="; dt
        If dt <> token.tokDataType Then
            Print #99, "EmitStatement: Let: lhs.dt <> t.dt"; dt; token.tokDataType
            MsgBox "EmitStatement: Let: lhs.dt <> t.dt " & dt & " " & token.tokDataType
'            Err.Raise 1 ' Internal error - data type conflict
        End If
        dt = dt And Not (VT_ARRAY Or VT_BYREF) ' VT_ARRAY base641a\project1.vbp
Print #99, "let: rhs.dt="; operand_stack.Item(1).tokDataType
        If dt <> operand_stack.Item(1).tokDataType Then
            Print #99, "EmitStatement: Let: lhs.dt <> rhs.dt " & dt & " " & operand_stack.Item(1).tokDataType
            MsgBox "EmitStatement: Let: lhs.dt <> rhs.dt " & dt & " " & operand_stack.Item(1).tokDataType
'            Err.Raise 1 ' Internal error - data type conflict
        End If
Print #99, "2"
#If 0 Then
        If dt = VT_BSTR Then
Print #99, "3"
            If Not operand_stack.Item(2).tokVariable Is Nothing Then
Print #99, "l="; operand_stack.Item(2).tokVariable.VarType.dtLength
                If operand_stack.Item(2).tokVariable.VarType.dtLength <> 0 Then
                    operand_stack.Item(2).tokOutput = PCodeToSymbol(token) & "N(" & operand_stack.Item(1).tokOutput & "," & operand_stack.Item(2).tokOutput & ")"
                    GoTo 1
                End If
            Else
                Print #99, "Let: tokVariable is Nothing"
                MsgBox "Let: tokVariable is Nothing"
                Err.Raise 1
            End If
        End If
#End If
Print #99, "4"
        operand_stack.Item(2).tokOutput = PCodeToSymbol(token) & "(" & operand_stack.Item(1).tokOutput & "," & operand_stack.Item(2).tokOutput & ")"
1
Print #99, "5"
        operand_stack.Remove 1
' fixme: better to form Invoke(...,1,4,result) here using vbPCodeInvoke???
    Case vbPCodePropertyLet, vbPCodePropertySet
PropertyLetSet:
Print #99, "propletset: t.dt="; token.tokDataType
        dt = operand_stack.Item(1).tokDataType
Print #99, "propletset: lhs.dt="; dt
        If dt <> token.tokDataType Then
            Print #99, "EmitStatement: PropLetSet: lhs.dt <> t.dt"; dt; token.tokDataType
            MsgBox "EmitStatement: PropLetSet: lhs.dt <> t.dt " & dt & " " & token.tokDataType
'            Err.Raise 1 ' Internal error - data type conflict
        End If
        operand_stack.Item(1).tokOutput = PCodeToSymbol(token) & "(" & operand_stack.Item(1).tokOutput & ")"
'        operand_stack.Item(2).tokOutput = PCodeToSymbol(token) & "(" & Left(operand_stack.Item(2).tokOutput, Len(operand_stack.Item(2).tokOutput) - 1) & "," & operand_stack.Item(1).tokOutput & "))"
'        operand_stack.Remove 1
    Case vbPCodeSet
If operand_stack.Count = 1 Then token.tokPCode = vbPCodePropertySet: GoTo PropertyLetSet
Print #99, "set: t.dt="; token.tokDataType
        dt = operand_stack.Item(2).tokDataType
Print #99, "set: lhs.dt="; dt
        If dt <> token.tokDataType Then
            Print #99, "EmitStatement: Set: lhs.dt <> t.dt"; dt; token.tokDataType
            MsgBox "EmitStatement: Set: lhs.dt <> t.dt " & dt & " " & token.tokDataType
'            Err.Raise 1 ' Internal error - data type conflict
        End If
        dt = dt And Not VT_BYREF
Print #99, "set: rhs.dt="; operand_stack.Item(1).tokDataType
        If dt <> operand_stack.Item(1).tokDataType Then
            Print #99, "EmitStatement: Set: lhs.dt <> rhs.dt " & dt & " " & operand_stack.Item(1).tokDataType
            MsgBox "EmitStatement: Set: lhs.dt <> rhs.dt " & dt & " " & operand_stack.Item(1).tokDataType
'            Err.Raise 1 ' Internal error - data type conflict
        End If
Print #99, "tokValue="; token.tokValue; " vn="; TypeName(token.tokValue); " vt="; VarType(token.tokValue)
        If IsObj(dt) Then
            If token.tokValue = "" Then
                operand_stack.Item(2).tokOutput = PCodeToSymbol(token) & "(" & operand_stack.Item(1).tokOutput & "," & operand_stack.Item(2).tokOutput & ",NULL)"
            Else
                operand_stack.Item(2).tokOutput = PCodeToSymbol(token) & "(" & operand_stack.Item(1).tokOutput & "," & operand_stack.Item(2).tokOutput & ",L""" & token.tokValue & """)"
            End If
        ElseIf dt = vbVariant Then
            operand_stack.Item(2).tokOutput = PCodeToSymbol(token) & "(" & operand_stack.Item(1).tokOutput & "," & operand_stack.Item(2).tokOutput & ")"
        Else
            Err.Raise 1 ' internal error
        End If
        operand_stack.Remove 1
    Case vbPCodeGet, vbPCodePut
        operand_stack.Item(3).tokOutput = PCodeToSymbol(token) & "(" & operand_stack.Item(1).tokOutput & "," & operand_stack.Item(2).tokOutput & ",0x" & Hex(token.tokDataType) & "," & operand_stack.Item(3).tokOutput & ")"
        operand_stack.Remove 1
        operand_stack.Remove 1
' fixme: same code as Write, create Sub
    Case vbPCodePrint
        Dim s As String
        s = "_PrintStart(" & operand_stack.Item(1).tokOutput & ") "
        operand_stack.Remove 1
        Dim sc As Boolean
        Do While operand_stack.Count > 0
Print #99, "print: c="; operand_stack.Count; " pc="; operand_stack.Item(1).tokPCode
            sc = False
            Select Case operand_stack.Item(1).tokPCode
                Case vbPCodePrintSpc
                    s = s & PCodeToSymbol(operand_stack.Item(1)) & "(" & operand_stack.Item(1).tokOutput & ") "
                Case vbPCodePrintTab
                    s = s & PCodeToSymbol(operand_stack.Item(1)) & operand_stack.Item(1).tokCount & "(" & operand_stack.Item(1).tokOutput & ") "
                Case vbPCodePrintComma
                    s = s & PCodeToSymbol(operand_stack.Item(1)) & " "
                Case vbPCodePrintSemiColon
                    sc = True
                Case Else
                    s = s & "_PrintExpr(" & operand_stack.Item(1).tokOutput & ") "
            End Select
            operand_stack.Remove 1
        Loop
        If Not sc Then s = s & " _PrintNL "
        operand_stack.Add New vbToken
        operand_stack.Item(1).tokOutput = s & " _PrintEnd"
    Case vbPCodeWrite
' fixme: duped code with Print
        s = "_WriteStart(" & operand_stack.Item(1).tokOutput & ") "
        operand_stack.Remove 1
        Do While operand_stack.Count > 0
Print #99, "Write: c="; operand_stack.Count; " pc="; operand_stack.Item(1).tokPCode
            sc = False
            Select Case operand_stack.Item(1).tokPCode
' change vbPCodePrintSpc to vbPCodePrintWriteSpc, etc?
                Case vbPCodePrintSpc
                    s = s & PCodeToSymbol(operand_stack.Item(1)) & "(" & operand_stack.Item(1).tokOutput & ") "
                Case vbPCodePrintTab
                    s = s & PCodeToSymbol(operand_stack.Item(1)) & operand_stack.Item(1).tokCount & "(" & operand_stack.Item(1).tokOutput & ") "
                Case vbPCodePrintComma
                    s = s & PCodeToSymbol(operand_stack.Item(1)) & " "
                Case vbPCodePrintSemiColon
                    sc = True
                Case Else
                    s = s & "_WriteExpr(" & operand_stack.Item(1).tokOutput & ") "
            End Select
            operand_stack.Remove 1
        Loop
        If Not sc Then s = s & " _WriteNL "
        operand_stack.Add New vbToken
        operand_stack.Item(1).tokOutput = s & " _WriteEnd"
    Case vbPCodeDebugPrint
        s = PCodeToSymbol(token) & "() "
        GoTo 999
    Case vbPCodePrintMethod ' do trailing space thing to vbPCodePrint/Write too
' fixme: duped code with Print, Write
        s = PCodeToSymbol(token) & "(" & operand_stack.Item(1).tokOutput & ") "
        operand_stack.Remove 1
999
        Do While operand_stack.Count > 0
Print #99, "printmethod: c="; operand_stack.Count; " pc="; operand_stack.Item(1).tokPCode
            sc = False
            Select Case operand_stack.Item(1).tokPCode
                Case vbPCodePrintSpc
                    s = s & PCodeToSymbol(operand_stack.Item(1)) & "(" & operand_stack.Item(1).tokOutput & ") "
                Case vbPCodePrintTab
                    s = s & PCodeToSymbol(operand_stack.Item(1)) & operand_stack.Item(1).tokCount & "(" & operand_stack.Item(1).tokOutput & ") "
                Case vbPCodePrintComma
                    s = s & PCodeToSymbol(operand_stack.Item(1)) & " "
                Case vbPCodePrintSemiColon
                    sc = True
                Case Else
                    s = s & "_PrintExpr(" & operand_stack.Item(1).tokOutput & ") "
            End Select
            operand_stack.Remove 1
        Loop
        If Not sc Then s = s & "_PrintNL "
        operand_stack.Add New vbToken
        operand_stack.Item(1).tokOutput = s & "_PrintEnd"
    Case vbPCodeInput
        s = "_InputStart(" & operand_stack.Item(1).tokOutput & ")"
        operand_stack.Remove 1
'        Dim sc As Boolean
        Do While operand_stack.Count > 0
            s = s & "_InputVariable(" & operand_stack.Item(1).tokOutput & ")"
            operand_stack.Remove 1
        Loop
        operand_stack.Add New vbToken
        operand_stack.Item(1).tokOutput = s & " _InputEnd"
    Case vbPCodeSetNothing, vbPCodePropertySetNothing
    ' fixme: chaning tokPCode - needs to be rewritten
        If operand_stack.Count = 1 Then token.tokPCode = vbPCodePropertySetNothing Else operand_stack.Remove 1 ' remove Nothing
        GenericOperands token, operand_stack
    Case vbPCodeOnGoSub
        token.tokOutput = AbbrDataType(token.tokDataType) & "OnGoSub(" & operand_stack.Item(1).tokOutput & ")"
        operand_stack.Remove 1
        Dim i As Integer
        Dim j As Integer
        j = operand_stack.Count
        For i = 1 To j
            token.tokOutput = token.tokOutput & " LabelOnGoSub(" & CStr(i) & "," & operand_stack.Item(1).tokOutput & ")"
            operand_stack.Remove 1
        Next
        token.tokOutput = token.tokOutput & " EndOnGoSub"
        operand_stack.Add token
    Case vbPCodeOnGoTo
        token.tokOutput = AbbrDataType(token.tokDataType) & "OnGoTo(" & operand_stack.Item(1).tokOutput & ")"
        operand_stack.Remove 1
        j = operand_stack.Count
        For i = 1 To j
            token.tokOutput = token.tokOutput & " LabelOnGoTo(" & CStr(i) & "," & operand_stack.Item(1).tokOutput & ")"
            operand_stack.Remove 1
        Next
        token.tokOutput = token.tokOutput & " EndOnGoTo"
        operand_stack.Add token
    Case vbPCodeWith
        token.tokOutput = AbbrDataType(token.tokWith.WithValue.tokDataType) & PCodeToSymbol(token)
        If operand_stack.Count <> 1 Then Err.Raise 1
''''        While operand_stack.Count
            token.tokOutput = token.tokOutput & "(WithValue(" & CStr(token.tokWith.WithCount) & ")," & operand_stack.Item(1).tokOutput & ")"
            operand_stack.Remove 1
''''        Wend
        operand_stack.Add token
    Case vbPCodeEndWith
        token.tokOutput = AbbrDataType(token.tokWith.WithValue.tokDataType) & PCodeToSymbol(token) & "(WithValue(" & CStr(token.tokWith.WithCount) & "))"
        operand_stack.Add token
    Case vbPCodeForNext, vbPCodeForNextV, vbPCodeForEachNext, vbPCodeForEachNextV
' obsolete Next, End Select, etc variable by using For class in vbtoken
        token.tokOutput = ""
        For i = 1 To token.tokCount ' Next statment can be "Next", "Next A" or "Next A, B"
            token.tokOutput = token.tokOutput & AbbrDataType(operand_stack.Item(1).tokDataType) & PCodeToSymbol(token) & "/* " & operand_stack.Item(1).tokOutput & " */ "
            operand_stack.Remove 1
        Next
        operand_stack.Add token
    Case vbPCodeSelect
Static select_save As String
select_save = operand_stack.Item(1).tokOutput
        GenericOperands token, operand_stack
    Case vbPCodeCase
        token.tokOutput = ""
'        operand_stack.Add New vbToken, , 1
'        operand_stack.Item(1).tokOutput = select_save
'        operand_stack.Item(1).tokOutput = operand_stack.Count - 1
        Do
            token.tokOutput = token.tokOutput & operand_stack.Item(1).tokOutput
            operand_stack.Remove 1
            If operand_stack.Count > 0 Then token.tokOutput = token.tokOutput & " Or "
        Loop While operand_stack.Count > 0
        token.tokOutput = "Case(" & token.tokOutput & ")"
        operand_stack.Add token
    Case vbPCodeOpen
Print #99, "PCodeOpen: c="; operand_stack.Count
Print #99, "vt="; VarType(operand_stack.Item(2).tokOutput)
Print #99, "o="; operand_stack.Item(2).tokOutput
        operand_stack.Item(2).tokOutput = "0x" & Hex(CLng(operand_stack.Item(2).tokOutput))
Print #99, "1"
        GenericOperands token, operand_stack
Print #99, "2"
    Case vbPCodeCall
Print #99, "3"
        token.tokOutput = "Call"
Print #99, "4"
        GenericOperands token, operand_stack
Print #99, "5"
    Case vbPCodeErase
        ' insert number of arrays as first token
        operand_stack.Add New vbToken, , 1
        operand_stack.Item(1).tokOutput = operand_stack.Count - 1
        GenericOperands token, operand_stack, True
    Case vbPCodeReDim
        ' insert number of variables as first token
        operand_stack.Add New vbToken, , 1
        operand_stack.Item(1).tokOutput = token.tokCount ' Number of variables
        GenericOperands token, operand_stack, True
    Case vbPCodeError
        token.tokOutput = "Error(" & operand_stack.Item(1).tokOutput & ")"
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeCloseFile
        token.tokOutput = "CloseFile" & "((" & token.tokCount
        For i = 1 To token.tokCount
            token.tokOutput = token.tokOutput & "," & operand_stack.Item(1).tokOutput
            operand_stack.Remove 1
        Next
        token.tokOutput = token.tokOutput & "))"
        operand_stack.Add token
    Case vbpcodeCircle, vbpcodeline, vbpcodepset, vbpcodescale
        GenericOperands token, operand_stack
        token.tokOutput = "VoidCall((" & token.tokOutput & "))"
    Case Else ' generic function
        GenericOperands token, operand_stack
End Select
PerformIndentChanges token.tokPCode
If operand_stack.Count = 1 Then
    Print #99, "EmitStatement: {"; operand_stack.Item(1).tokOutput; "}"
Else
    Print #99, "EmitStatement: operand_stack <> 1: " & operand_stack.Count
    MsgBox "EmitStatement: operand_stack <> 1: " & operand_stack.Count
    Err.Raise 1 ' should operand_stack be popped and replaced with token.output?
End If
End Sub

' What about Neg?
Function PCodeToAbbr(ByVal pcode As vbPCodes) As String
Select Case pcode
    Case vbPCodeImp
        PCodeToAbbr = "Imp"
    Case vbPCodeEqv
        PCodeToAbbr = "Eqv"
    Case vbPCodeXor
        PCodeToAbbr = "Xor"
    Case vbPCodeOr
        PCodeToAbbr = "Or"
    Case vbPCodeAnd
        PCodeToAbbr = "And"
    Case vbPCodeNot
        PCodeToAbbr = "Not"
    Case vbPCodeEQ
        PCodeToAbbr = "EQ"
    Case vbPCodeLT
        PCodeToAbbr = "LT"
    Case vbPCodeLE
        PCodeToAbbr = "LE"
    Case vbPCodeNE
        PCodeToAbbr = "NE"
    Case vbPCodeGT
        PCodeToAbbr = "GT"
    Case vbPCodeGE
        PCodeToAbbr = "GE"
    Case vbPCodeIs
        PCodeToAbbr = "Is"
    Case vbPCodeLike
        PCodeToAbbr = "Like"
    Case vbPCodeCat
        PCodeToAbbr = "Cat"
    Case vbPCodeAdd
        PCodeToAbbr = "Add"
    Case vbPCodeSub
        PCodeToAbbr = "Sub"
    Case vbPCodeMod
        PCodeToAbbr = "Mod"
    Case vbPCodeIDiv
        PCodeToAbbr = "IDiv"
    Case vbPCodeMul
        PCodeToAbbr = "Mul"
    Case vbPCodeDiv
        PCodeToAbbr = "Div"
    Case vbPCodePositive
        PCodeToAbbr = "Pos"
    Case vbPCodeNegative
        PCodeToAbbr = "Neg"
    Case vbPCodePow
        PCodeToAbbr = "Pow"
    Case Else
        Print #99, "PCodeToAbbr: Internal error - Unknown PCode: " & pcode
        MsgBox "PCodeToAbbr: Internal error - Unknown PCode: " & pcode
        Err.Raise 1 ' Internal error - Unknown PCode
End Select
End Function

Function COperatorPriority(ByVal pcode As vbPCodes) As Integer
Select Case pcode
    Case vbPCodeImp
        COperatorPriority = COprPriorityPrimary
    Case vbPCodeEqv
        COperatorPriority = COprPriorityPrimary
    Case vbPCodeXor
        COperatorPriority = COprPriorityBitwiseXor
    Case vbPCodeOr
        COperatorPriority = COprPriorityBitwiseOr
    Case vbPCodeAnd
        COperatorPriority = COprPriorityBitwiseAnd
    Case vbPCodeNot
        COperatorPriority = COprPriorityUnary
    Case vbPCodeEQ
        COperatorPriority = COprPriorityEqNe
    Case vbPCodeLT
        COperatorPriority = COprPriorityLtGtLeGe
    Case vbPCodeLE
        COperatorPriority = COprPriorityLtGtLeGe
    Case vbPCodeNE
        COperatorPriority = COprPriorityEqNe
    Case vbPCodeGT
        COperatorPriority = COprPriorityLtGtLeGe
    Case vbPCodeGE
        COperatorPriority = COprPriorityLtGtLeGe
    Case vbPCodeIs
        COperatorPriority = COprPriorityPrimary
    Case vbPCodeLike
        COperatorPriority = COprPriorityPrimary
    Case vbPCodeCat
        COperatorPriority = COprPriorityPrimary
    Case vbPCodeAdd
        COperatorPriority = COprPriorityAddSub
    Case vbPCodeSub
        COperatorPriority = COprPriorityAddSub
    Case vbPCodeMod
        COperatorPriority = COprPriorityMulDivMod
    Case vbPCodeIDiv
        COperatorPriority = COprPriorityMulDivMod
    Case vbPCodeMul
        COperatorPriority = COprPriorityMulDivMod
    Case vbPCodeDiv
        COperatorPriority = COprPriorityMulDivMod
    Case vbPCodePositive
        COperatorPriority = COprPriorityUnary
    Case vbPCodeNegative
        COperatorPriority = COprPriorityUnary
    Case vbPCodePow
        COperatorPriority = COprPriorityPrimary
    Case Else
        Print #99, "COperatorPriority: Internal error - Unknown PCode: " & pcode
        MsgBox "COperatorPriority: Internal error - Unknown PCode: " & pcode
        Err.Raise 1 ' Internal error - Unknown PCode
End Select
End Function

' Use C name, if needed
Function PCodeToSymbol(ByVal token As vbToken) As String
Print #99, "PCodeToSymbol: pc="; token.tokPCode; " dt="; token.tokDataType
Select Case token.tokPCode
    Case vbPCodeLet, vbPCodeLSet, vbPCodeRSet, vbPCodeSet
        PCodeToSymbol = AbbrDataType(token.tokDataType) & token.tokOutput
'        PCodeToSymbol = cTypeName(token.tokVariable) & token.tokOutput
    Case vbPCodePropertyLet
        PCodeToSymbol = "PropertyLet"
    Case vbPCodePropertySet
        PCodeToSymbol = "PropertySet"
    Case vbPCodeCvt ' Fill in other pcodes
        PCodeToSymbol = AbbrDataType(token.tokPCodeSubType) & "To" & AbbrDataType(token.tokDataType)
    Case vbPCodeSetNothing
        PCodeToSymbol = AbbrDataType(token.tokDataType) & "SetNothing"
    Case vbPCodePropertySetNothing
        PCodeToSymbol = AbbrDataType(token.tokDataType) & "PropertySetNothing"
'        PCodeToSymbol = cTypeName(token.tokVariable) & "SetNothing"
    Case vbPCodeDebugAssert
        PCodeToSymbol = "_DebugAssert"
    Case vbPCodeDebugPrint
        PCodeToSymbol = "_DebugPrint"
    Case vbPCodePrint
        PCodeToSymbol = "_Print"
    Case vbPCodePrintMethod
        PCodeToSymbol = "_PrintMethod"
    Case vbPCodePrintSpc
        PCodeToSymbol = "_PrintSpc"
    Case vbPCodePrintTab
        PCodeToSymbol = "_PrintTab" ' PrintTab becomes either PrintTab0 or PrintTab1 depending on arg count
    Case vbPCodePrintComma
        PCodeToSymbol = "_PrintComma"
    Case vbPCodePrintSemiColon
        Print #99, "PCodeToSymbol: Internal error - Unknown PCode: " & token.tokPCode
        MsgBox "PCodeToSymbol: Internal error - Unknown PCode: " & token.tokPCode
        Err.Raise 1 ' Internal Error - Unexpected PCode
' The following should be removed and spaces compressed in token.tokoutput
    Case vbPCodeCaseElse
        PCodeToSymbol = "CaseElse"
    Case vbPCodeEndSelect
        PCodeToSymbol = "EndSelect"
    Case vbPCodeEndWith
        PCodeToSymbol = "EndWith"
    Case vbPCodeDoUntil
        PCodeToSymbol = "DoUntil"
    Case vbPCodeDoWhile
        PCodeToSymbol = "DoWhile"
    Case vbPCodeEndIf
        PCodeToSymbol = "EndIf"
    Case vbPCodeLineInput
        PCodeToSymbol = "LineInput"
    Case vbPCodeLoopInfinite
        PCodeToSymbol = "LoopInfinite"
    Case vbPCodeLoopUntil
        PCodeToSymbol = "LoopUntil"
    Case vbPCodeLoopWhile
        PCodeToSymbol = "LoopWhile"
    Case vbPCodeExitDo
        PCodeToSymbol = "ExitDo"
    Case vbPCodeExitFor
        PCodeToSymbol = "ExitFor"
    Case vbPCodeExitSub
        PCodeToSymbol = "ExitSub"
    Case vbPCodeExitFunction
        PCodeToSymbol = "ExitFunction"
    Case vbPCodeOnError0
        PCodeToSymbol = "OnErrorGoTo0"
    Case vbPCodeOnErrorLabel
        PCodeToSymbol = "OnErrorGoTo"
    Case vbpcodeonerrorresumenext
        PCodeToSymbol = "OnErrorResumeNext"
    Case vbPCodeResume0
        PCodeToSymbol = "Resume0"
    Case vbPCodeResumeNext
        PCodeToSymbol = "ResumeNext"
    Case vbPCodeforeach
        PCodeToSymbol = AbbrDataType(token.tokDataType) & "ForEach"
    Case vbPCodeForNext, vbPCodeForNextV, vbPCodeForEachNext, vbPCodeForEachNextV
        PCodeToSymbol = "Next"
    Case Else
        PCodeToSymbol = token.tokOutput
Print #99, "PCodeToSymbol: ";
Print #99, "tlib ref="; Not token.tokReference Is Nothing;
Print #99, " di="; Not token.tokDeclarationInfo Is Nothing;
Print #99, " ii="; Not token.tokInterfaceInfo Is Nothing;
Print #99, " mi="; Not token.tokMemberInfo Is Nothing
        If Not token.tokReference Is Nothing Then
' fixme: probably should have all tokReference contain valid InterfaceInfo (or DeclarationInfo)
' fixme: need to make DeclarationInfo and InterfaceInfo mutually exclusive but InterfaceInfo sometimes carries result - arg
            If Not token.tokDeclarationInfo Is Nothing Then
                PCodeToSymbol = TypeInfoToCType(token.tokDeclarationInfo) & "_" & ProcNameIK(token.tokMemberInfo.Name, token.tokMemberInfo.InvokeKind)
            ElseIf Not token.tokInterfaceInfo Is Nothing Then
                PCodeToSymbol = TypeInfoToCType(token.tokInterfaceInfo) & "_" & ProcNameIK(token.tokMemberInfo.Name, token.tokMemberInfo.InvokeKind)
            Else
                Err.Raise 1
            End If
        End If
End Select
End Function

Sub GenericOperands(ByVal token As vbToken, ByVal operand_stack As Collection, Optional ByVal DoubleParenthesis As Boolean)
Print #99, "GenericOperands: s="; token.tokOutput; " dt="; token.tokDataType; " osc="; operand_stack.Count; " dp="; DoubleParenthesis
If token.tokDataType <> 0 Then token.tokOutput = AbbrDataType(token.tokDataType) & token.tokOutput
If operand_stack.Count = 0 Then
    token.tokOutput = PCodeToSymbol(token)
Else
    Dim t As vbToken
    Dim s As String
    For Each t In operand_stack
        s = s & ", " & operand_stack.Item(1).tokOutput
        operand_stack.Remove 1
    Next
    token.tokOutput = PCodeToSymbol(token) & IIf(DoubleParenthesis, "((", "(") & Mid(s, 3) & IIf(DoubleParenthesis, "))", ")")
End If
operand_stack.Add token
End Sub

' Perform changes to indent
Sub PerformIndentChanges(ByVal pcode As vbPCodes)
Select Case pcode
    Case vbPCodefor, vbPCodeforeach, vbPCodeIf, vbPCodeDo, vbPCodeDoUntil, vbPCodeDoWhile, vbPCodeSingleIf, vbPCodeWhile, vbPCodeWith
        ' add indent starting with next statement
        newindent = indent & IndentString
    Case vbPCodeSelect
        newindent = indent & IndentString & IndentString
    Case vbPCodeEndIf, vbPCodeEndWith, vbPCodeSingleIfEndIf, vbPCodeSingleIfEndIfElse, vbPCodeForNext, vbPCodeForNextV, vbPCodeForEachNext, vbPCodeForEachNextV, vbPCodeLoop, vbPCodeLoopInfinite, vbPCodeLoopUntil, vbPCodeWend, vbPCodeLoopWhile
        ' remove indent
        indent = Left(indent, Len(indent) - Len(IndentString))
        newindent = indent
    Case vbPCodeCase, vbPCodeCaseElse, vbPCodeElse, vbPCodeSingleIfElse, vbPCodeElseIf
        ' remove indent, add indent for next statement
        indent = Left(indent, Len(indent) - Len(IndentString))
    Case vbPCodeEndSelect
        ' remove two indents
        indent = Left(indent, Len(indent) - Len(IndentString) * 2)
        newindent = indent
    Case Else
        ' no indent changes
End Select
End Sub

' Passing two mi versions due to TLI oddity, first is full version, 2nd is vtable (could be partial)
Sub OutputMemberInfo(ByVal tli As TypeLibInfo, ByVal ii As InterfaceInfo, ByVal mi As MemberInfo, ByVal ik As InvokeKinds, ByVal Helper As Boolean)
Dim i As Integer
Dim s As String
Dim pi As ParameterInfo
Dim retval As ParameterInfo
Print #99, "OutputMemberinfo: "; tli Is Nothing; " "; ii Is Nothing; " "; mi Is Nothing
Print #99, "OutputMemberInfo: tli="; TypeInfoToCType(ii); " ii.tk="; ii.TypeKind; " ii am="; Hex(ii.AttributeMask); " ik="; ik; " mi.ik="; mi.InvokeKind; " mi.am="; Hex(mi.AttributeMask)
Print #99, "OutputMemberInfo: mi.pc="; mi.Parameters.Count
'If mi.AttributeMask And 1 Then Err.Raise 1 ' Help.vbp
On Error Resume Next
Print #99, "mi.vtbl="; mi.VTableOffset
Print #99, "is ex="; mi.ReturnType.IsExternalType
Print #99, "r="; mi.ReturnType.TypeInfo.Name
Print #99, "tli.infos p="; tli.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).VTableInterface.GetMember(mi.Name).Parameters.Count
Print #99, "tli.infos n="; tli.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).VTableInterface.GetMember(mi.Name).ReturnType.TypeInfo.Name
Print #99, "tli.infos di="; tli.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).DefaultInterface.GetMember(0).Name
Print #99, "ex.name="; mi.ReturnType.TypeLibInfoExternal.Name
Print #99, "vt="; VarType(mi.ReturnType.TypedVariant)
Print #99, "tn="; TypeName(mi.ReturnType.TypedVariant)
Print #99, "n="; mi.ReturnType.TypeLibInfoExternal.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).Name
Print #99, "tk="; mi.ReturnType.TypeLibInfoExternal.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).TypeKind
Print #99, "am="; mi.ReturnType.TypeLibInfoExternal.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).AttributeMask
Print #99, "bet="; mi.ReturnType.TypeLibInfoExternal.BestEquivalentType(mi.ReturnType.TypeInfo.Name)
Print #99, "di-ii="; mi.ReturnType.TypeLibInfoExternal.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).DefaultInterface.Name
Print #99, "di-ii-df="; mi.ReturnType.TypeLibInfoExternal.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).DefaultInterface.GetMember(0).Name
Print #99, "di-rtvt="; mi.ReturnType.TypeLibInfoExternal.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).DefaultInterface.ResolvedType.VarType
Print #99, "rtvt="; mi.ReturnType.TypeLibInfoExternal.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).ResolvedType.VarType
Print #99, "rtvt="; mi.ReturnType.TypeLibInfoExternal.TypeInfos.IndexedItem(mi.ReturnType.TypeInfoNumber).ResolvedType.VarType
Print #99, "mi="; mi.Name
On Error GoTo 0
Dim ss As String
On Error GoTo 10
For Each pi In mi.Parameters
    i = i + 1
Print #99, "pi.name="; pi.Name; " i="; i; VarTypeInfoToCType(pi.VarTypeInfo); " "; pi.Name; " ik="; mi.InvokeKind; " f="; Hex(pi.Flags)
'
    If pi.Flags And PARAMFLAG_FRETVAL Then
        If Not retval Is Nothing Then Err.Raise 1 ' TypeLib error - multiple PARAMFLAG_FRETVAL
        Set retval = pi
    Else
        s = s & "," & VarTypeInfoToCType(pi.VarTypeInfo) & " _p_" & pi.Name
        ss = ss & ",_p_" & pi.Name
        ' need to check for last param also?
        ' note: Additem in vb6.idl doesn't contain param names
        If pi.Name = "" Then
            If i = mi.Parameters.Count And (mi.InvokeKind = INVOKE_PROPERTYPUT Or mi.InvokeKind = INVOKE_PROPERTYPUTREF) And (mi.ReturnType.VarType = VT_VOID Or mi.ReturnType.VarType = VT_HRESULT) Then
                s = s & "_p_putval"
                ss = ss & "_p_putval"
            Else
                s = s & "_p_arg" & CStr(i)
                ss = ss & "_p_arg" & CStr(i)
            End If
        End If
    End If
Print #99, i = mi.Parameters.Count; pi Is mi.Parameters.Item(mi.Parameters.Count)
Next

Dim rtdt As TliVarType
Dim rtii As InterfaceInfo
Dim ReturnType As VarTypeInfo
Set ReturnType = GetReturnType(ii, mi, rtdt, rtii)
If Not ReturnType Is Nothing Then
' PictureBox_Picture is INVOKE_PROPERTYPUT with VT_DISPATCH return but doesn't seem to be in idl. Why?
'    If ik = INVOKE_PROPERTYPUT And rtdt = VT_DISPATCH Then Err.Raise 1 ' Exit Sub
    If ik = INVOKE_PROPERTYPUTREF And Not IsVObj(rtdt) Then Err.Raise 1 ' Exit Sub
End If
' this fixes a TypeLib (TLI control) bug(?) for -> CommonDialog1.DialogTitle=""
' This fixes TypeLib bug(?) where propput/ref has putval in ReturnType instead of parameter
If ik And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
'    If mi.Parameters.Count = 0 Then
'    If Not ReturnType Is Nothing Then ' Err.Raise 1
        If mi.ReturnType.VarType <> VT_VOID And mi.ReturnType.VarType <> VT_HRESULT Then
            s = s & "," & VarTypeInfoToCType(ReturnType) & " _p_putval"
            ss = ss & ",_p_putval"
    '    Else
    '        If Not ReturnType Is Nothing Then Err.Raise 1
        End If
'    End If
    Set ReturnType = Nothing
End If


Dim rv As String
Dim sss As String
rv = VarTypeInfoToCType(ReturnType, 1)
sss = rv & " " & TypeInfoToCType(ii) & "_" & ProcNameIK(mi.Name, ik) & "(" & TypeInfoToCType(ii) & "* This" & s & ")"
If Helper Then
    Print #99, "rv="; rv; " rt="; Not mi.ReturnType Is Nothing; " retval="; Not retval Is Nothing
    If Not mi.ReturnType Is Nothing Then Print #99, "rt.ik="; mi.InvokeKind; " rt.vt="; mi.ReturnType.VarType; " rt.pl="; mi.ReturnType.PointerLevel
    If rv = "void" Then
            sss = sss & "{ChkHR(This->lpVtbl->" & MINameVTable(mi) & "(ADJUST_THIS(" & TypeInfoToCType(ii) & ",This," & MINameVTable(mi) & ")" & ss & "),L""" & TypeInfoToCType(ii) & "_" & ProcNameIK(mi.Name, ik) & """); }"
'    ElseIf mi.InvokeKind = INVOKE_FUNC And mi.ReturnType.VarType <> VT_HRESULT And mi.ReturnType.PointerLevel = 0 Then
    ElseIf retval Is Nothing Then
            ' For idl definitions like: long _stdcall IsInTransaction(); COM+ Services
            sss = sss & "{return NoChkHR(This->lpVtbl->" & MINameVTable(mi) & "(ADJUST_THIS(" & TypeInfoToCType(ii) & ",This," & MINameVTable(mi) & ")" & ss & "),L""" & TypeInfoToCType(ii) & "_" & ProcNameIK(mi.Name, ik) & """);}"
    Else
            sss = sss & "{" & rv & " _v_retval" & "; ChkHR(This->lpVtbl->" & MINameVTable(mi) & "(ADJUST_THIS(" & TypeInfoToCType(ii) & ",This," & MINameVTable(mi) & ")" & ss & ",&_v_retval),L""" & TypeInfoToCType(ii) & "_" & ProcNameIK(mi.Name, ik) & """); return _v_retval;}"
    End If
Else
    sss = sss & ";"
End If
Print #99, sss
Print #1, sss
If False Then
10  Print #99, "Bad typelib entry: " & tli.Name & "." & ii.Name & "." & mi.Name
'    MsgBox "Bad typelib entry: " & tli.Name & "." & ii.Name & "." & mi.Name
    Resume 20
End If
20
End Sub

Sub OutputInterfaceInfo(ByVal tli As TypeLibInfo, ByVal ii As InterfaceInfo, ByVal Helper As Boolean)
Dim mi As MemberInfo
Dim tk As TypeKinds
Dim iiv As InterfaceInfo
If Left(ii.Name, 2) <> "_" & Chr(1) Then ' don't output hidden interfaces
    tk = ii.TypeKind
    If tk = TKIND_ALIAS Then
        Print #99, "tk is alias"
    Else
        On Error Resume Next
        Set iiv = ii.VTableInterface
        On Error GoTo 0
        If iiv Is Nothing Then
Print #99, "No vtable interface for "; ii.Name
            Print #1, "/* dispinterface: "; ii.Name; " "; ii.MajorVersion; "."; ii.MinorVersion; " "; ii.GUID; " */"
            For Each mi In ii.Members
#If 1 Then
                Print #1, "/* "; mi.Name; " */"
#Else
                ' first mi is full version, 2nd mi is vtable version (could be partial)
                Dim ReturnType As VarTypeInfo
                
                OutputMemberInfo tli, ii, mi, GetReturnType(ii, mi, dt), INVOKE_PROPERTYGET, Helper
                OutputMemberInfo tli, ii, mi, GetReturnType(ii, mi, dt), INVOKE_PROPERTYPUT, Helper
                OutputMemberInfo tli, ii, mi, GetReturnType(ii, mi, dt), INVOKE_PROPERTYPUTREF, Helper
#End If
            Next
        Else
            For Each mi In iiv.Members
                If mi.AttributeMask And 1 Then
                    Print #1, "/* "; mi.Name; " */" ' Hidden member
                Else
                    OutputMemberInfo tli, iiv, mi, mi.InvokeKind, Helper
                End If
            Next
        End If
    End If ' TKIND_ALIAS
End If ' _ (hidden)
End Sub

Function MINameVTable(ByVal mi As MemberInfo) As String
MINameVTable = "_m_Vtbl" & Right("000" & Hex(mi.VTableOffset), 4)
End Function

Function TypeInfoToCType(ByVal ti As TypeInfo) As String
Print #99, "TypeInfoToCType: "; ti.Parent.Name & "_" & ti.Name; " tk="; ti.TypeKind
Select Case ti.TypeKind
Case TKIND_ALIAS
    TypeInfoToCType = "a"
Case TKIND_COCLASS
    TypeInfoToCType = TypeInfoToCType(ti.DefaultInterface) ' recursive, but just once?
    Exit Function
Case TKIND_DISPATCH ' May or may not have a VTable
    TypeInfoToCType = IIf(ti.VTableInterface Is Nothing, "d", "i")
Case TKIND_ENUM
    TypeInfoToCType = "e"
Case TKIND_INTERFACE
    TypeInfoToCType = "i" ' Has VTable, may or may not be derived from IDispatch
Case TKIND_MAX
    TypeInfoToCType = "x"
Case TKIND_MODULE
    TypeInfoToCType = "m"
Case TKIND_RECORD
    TypeInfoToCType = "r"
Case TKIND_UNION
    TypeInfoToCType = "u"
Case Else
    Err.Raise 1
End Select
TypeInfoToCType = "_" & TypeInfoToCType & "_" & ti.Parent.Name & "_" & ti.Name
' fixme: GIF89LibCtl.Gif89a was a bad interface name. Leave error or convert "." to "_"?
If InStr(TypeInfoToCType, ".") Then
    Print #99, "TypeInfoToCType: Invalid interface name:"; TypeInfoToCType
    ' bad interface name GIF89LibCtl.Gif89a
    MsgBox "Bad interface name:" & TypeInfoToCType
    ' let C compiler catch error, edit .h file by hand
    ' err.raise 1
End If
End Function

Function VarDeclaration(ByVal v As vbVariable, Optional ByVal SupressModuleName As Boolean) As String
Print #99, "VarDeclaration: v="; v.varSymbol; " vt="; v.VarType Is Nothing
Print #99, "dt="; v.VarType.dtDataType; " m.m="; v.varModule Is Nothing
Print #99, "m.n="; v.varModule.Name; " m.ct="; v.varModule.Component.Type
Dim s As String
' fixme: OK for SAFEARRAY declaration to suppresses emitting of length by testing for vt.varDimensions Is Nothing
If v.VarType.dtDataType = VT_LPWSTR Then
'''''If v.varType.dtDataType = VT_BSTR And v.varType.dtLength <> 0 Then
    s = "[" & v.VarType.dtLength & "]"
    If Not v.varDimensions Is Nothing Then s = "/* " & s & " */"
End If
If Not v.varDimensions Is Nothing Then
    s = s & " /* "
    Dim p As vbVarDimension
    For Each p In v.varDimensions
        s = s & "[" & p.varDimensionUBound - p.varDimensionLBound + 1 & "]"
    Next
    s = s & " */"
End If
VarDeclaration = cTypeName(v) & " _v_" & IIf(SupressModuleName, "", v.varModule.Name & "_") & v.varSymbol & s & ";"
End Function

' type strings are used to generate both C and IDL
Function cTypeName(ByVal vt As vbVariable) As String
If vt Is Nothing Then
    cTypeName = "void"
    Exit Function
End If
Print #99, "ctypename: sym="; vt.varSymbol
If vt.VarType Is Nothing Then
    Print #99, "cTypeName: vt.VarType is Nothing"
    MsgBox "cTypeName: vt.VarType is Nothing"
    Err.Raise 1 ' Internal error - vt.VarType is Nothing
End If
Print #99, "ctypename: dt="; vt.VarType.dtDataType
Select Case vt.VarType.dtDataType
Case VT_DISPATCH, VT_UNKNOWN ' IsObj()
    If Not vt.VarType.dtClass Is Nothing Then
        cTypeName = "_i_" & vt.VarType.dtClass.Name & " *"
    ElseIf Not vt.VarType.dtInterfaceInfo Is Nothing Then
' This is really lame, don't know how to get TLib name, so kludge
        cTypeName = "_i_VB_" & vt.VarType.dtInterfaceInfo.Name & " *"
        On Error Resume Next ' _TextBox doesn't seem to have Parent
        cTypeName = TypeInfoToCType(vt.VarType.dtInterfaceInfo) & " *"
        On Error GoTo 0
    Else
        cTypeName = VarTypeToC(vt.VarType.dtDataType)
    End If
Case VT_RECORD
On Error Resume Next
Print #99, "vt.VarType="; Not vt.VarType Is Nothing
Print #99, "vt.VarType.dtUDT="; Not vt.VarType.dtUDT Is Nothing
Print #99, "vt.VarType.dtUDT.TypeName="; vt.VarType.dtUDT.TypeName
Print #99, "vt.VarType.dtUDT.TypeModule="; Not vt.VarType.dtUDT.typeModule Is Nothing
Print #99, "vt.VarType.dtRecordInfo="; Not vt.VarType.dtRecordInfo Is Nothing
On Error GoTo 0
    If Not vt.VarType.dtUDT Is Nothing Then
        cTypeName = "struct __t_" & vt.VarType.dtUDT.typeModule.Name & "_" & vt.VarType.dtUDT.TypeName ' & "*" ' UDT ref change
    ElseIf Not vt.VarType.dtRecordInfo Is Nothing Then
        cTypeName = TypeInfoToCType(vt.VarType.dtRecordInfo) ' & "*" ' UDT ref change
    Else
        Err.Raise 1
    End If
#If 0 Then
Case VT_BSTR
' need for this test suggests that dt should be assigned VT_LPWSTR in vbt
    If vt.VarType.dtLength = 0 Then
        cTypeName = VarTypeToC(vt.VarType.dtDataType)
    Else
        cTypeName = VarTypeToC(VT_LPWSTR)
    End If
#End If
Case Else
Print #99, "ctypename: 3"
    cTypeName = VarTypeToC(vt.VarType.dtDataType)
End Select
If vt.varAttributes And VARIABLE_BYREF Then cTypeName = cTypeName & " * "
' If Not vt.varDimensions Is Nothing Then If vt.varDimensions.Count = 0 Then cTypeName = cTypeName & " * "
If Not vt.varDimensions Is Nothing Then If vt.VarType.dtDataType <> VT_VOID Then cTypeName = "SAFEARRAY * /* " & cTypeName & " */ "
Print #99, "ctypename: s="; cTypeName
End Function


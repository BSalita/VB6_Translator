Attribute VB_Name = "vbtSubs"
Option Explicit

Public Const VT_MAXTYPE As Integer = 32 ' Set to maximum VarType

Public Const VB_OPEN_ACCESS_MASK As Integer = &HF00
Public Const VB_OPEN_ACCESS_NONE As Integer = &H0
Public Const VB_OPEN_ACCESS_READ As Integer = &H100
Public Const VB_OPEN_ACCESS_WRITE As Integer = &H200
Public Const VB_OPEN_ACCESS_READ_WRITE As Integer = &H300

Public Const VB_OPEN_LOCK_MASK As Integer = &HF000
Public Const VB_OPEN_LOCK_READ_WRITE As Integer = &H1000
Public Const VB_OPEN_LOCK_WRITE As Integer = &H2000
Public Const VB_OPEN_LOCK_READ As Integer = &H3000
Public Const VB_OPEN_LOCK_SHARED As Integer = &H4000

Public Const VB_OPEN_MODE_MASK As Integer = &HFF
Public Const VB_OPEN_MODE_APPEND As Integer = &H8
Public Const VB_OPEN_MODE_BINARY As Integer = &H20
Public Const VB_OPEN_MODE_INPUT As Integer = &H1
Public Const VB_OPEN_MODE_OUTPUT As Integer = &H2
Public Const VB_OPEN_MODE_RANDOM As Integer = &H4

Public Enum vbOprPriority
    vbOprPriority0
    vbOprPriorityImp
    vbOprPriorityEqv
    vbOprPriorityXor
    vbOprPriorityOr
    vbOprPriorityAnd
    vbOprPriorityNot
    vbOprPriorityCmp    ' =,<,<=,<>,>,>=,IS,LIKE
    vbOprPriorityCat
    vbOprPriorityAddSub
    vbOprPriorityMod
    vbOprPriorityIDiv
    vbOprPriorityMulDiv
    vbOprPriorityPositiveNegative
    vbOprPriorityPow
End Enum

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As GUID) As Long

Function getGUID() As String
Dim g As GUID
If CoCreateGuid(g) <> 0 Then
    Print #99, "CoCreateGuid failed"
    MsgBox "CoCreateGuid failed"
    Err.Raise 1  ' CoCreateGuid failed
End If
getGUID = "{" & Right("00000000" & Hex(g.Data1), 8) & "-" _
    & Right("0000" & Hex(g.Data2), 4) & "-" _
    & Right("0000" & Hex(g.Data3), 4) & "-" _
    & Right("00" & Hex(g.Data4(0)), 2) _
    & Right("00" & Hex(g.Data4(1)), 2) & "-"
Dim i As Integer
For i = 2 To 7
    getGUID = getGUID & Right("00" & Hex(g.Data4(i)), 2)
Next
getGUID = getGUID & "}"
End Function

Function VT_Type(ByVal vt As vbVariable) As String
' type strings are used to generate both C and IDL
If vt Is Nothing Then
    Print #99, "VT_Type: vt is Nothing"
    MsgBox "VT_Type: vt is Nothing"
    Err.Raise 1
End If
If vt.varType Is Nothing Then
    Print #99, "VT_Type: vt.VarType is Nothing"
    MsgBox "VT_Type: vt.VarType is Nothing"
    Err.Raise 1
End If
If vt.varAttributes And VARIABLE_BYREF Then
    VT_Type = "VT_PTR"
Else
    Select Case vt.varType.dtDataType And Not (VT_ARRAY Or VT_BYREF Or VT_VECTOR)
    Case VT_EMPTY ' 0
        VT_Type = "VT_EMPTY" ' was VT_NULL!?
    Case VT_NULL ' 1
        VT_Type = "VT_NULL"
    Case VT_I2 ' 2
        VT_Type = "VT_I2"
    Case VT_I4 ' 3
        VT_Type = "VT_I4"
    Case VT_R4 '4
        VT_Type = "VT_R4"
    Case VT_R8 ' 5
        VT_Type = "VT_R8"
    Case VT_CY ' 6
        VT_Type = "VT_CY"
    Case VT_DATE ' 7
        VT_Type = "VT_DATE"
    Case VT_BSTR ' 8
        VT_Type = "VT_BSTR"
    Case VT_DISPATCH ' 9
        VT_Type = "VT_DISPATCH"
    Case VT_ERROR   ' 10
        VT_Type = "VT_ERROR"
    Case VT_BOOL ' 11
        VT_Type = "VT_BOOL"
    Case VT_VARIANT ' 12
        VT_Type = "VT_VARIANT"
    Case VT_UNKNOWN ' 13
        VT_Type = "VT_UNKNOWN"
    Case VT_DECIMAL ' 14
        VT_Type = "VT_DECIMAL"
    Case VT_I1 ' 16
        VT_Type = "VT_I1"
    Case VT_UI1 ' 17
        VT_Type = "VT_UI1"
    Case VT_UI2 ' 18
        VT_Type = "VT_UI2"
    Case VT_UI4 ' 19
        VT_Type = "VT_UI4"
    Case VT_I8  ' 20
        VT_Type = "VT_I8"
    Case VT_UI8 ' 21
        VT_Type = "VT_UI8"
    Case VT_INT ' 22
        VT_Type = "VT_INT"
    Case VT_UINT ' 23
        VT_Type = "VT_UINT"
    Case VT_VOID
        VT_Type = "VT_VOID"
    Case VT_HRESULT ' 25 - Form1.Show needs this
        VT_Type = "VT_HRESULT"
    Case VT_LPSTR ' 30
        VT_Type = "VT_LPSTR"
    Case VT_LPWSTR ' 31
        VT_Type = "VT_LPWSTR"
    Case VT_RECORD ' 36
        VT_Type = "VT_RECORD"
    Case Else
        Print #99, "VT_Type: Unknown VarType: "; Hex(vt.varType.dtDataType)
        MsgBox "VT_Type: Unknown VarType: " & Hex(vt.varType.dtDataType)
        Err.Raise 1 ' Internal error - Unknown VarType
    End Select
    If vt.varType.dtDataType And VT_ARRAY Then VT_Type = "(" & VT_Type & "|VT_ARRAY)"
End If
End Function

' check over C data type names, show values
Function VarTypeToC(ByVal vt As TliVarType) As String
If vt And VT_ARRAY Then
'    vt = vt And Not vbArray
    VarTypeToC = "SAFEARRAY"
Else
    Select Case vt And Not VT_VECTOR
    Case VT_EMPTY
        VarTypeToC = "_VarEmpty_t"
    Case VT_NULL
        VarTypeToC = "_VarNull_t"
    Case VT_I2, VT_BOOL
        VarTypeToC = "short" ' int16_t
    Case VT_I4
        VarTypeToC = "long" ' int32_t
    Case VT_R4
        VarTypeToC = "float"
    Case VT_R8
        VarTypeToC = "double"
    Case VT_DATE
        VarTypeToC = "DATE"
    Case VT_CY
        VarTypeToC = "CY"
    Case VT_BSTR
        VarTypeToC = "BSTR"
    Case VT_VARIANT
        VarTypeToC = "VARIANT"
    Case VT_UI1
        VarTypeToC = "unsigned char"
    Case VT_UNKNOWN ' 13
        VarTypeToC = "IUnknown *"
    Case VT_DISPATCH
        VarTypeToC = "IDispatch *"
    Case VT_VOID
        VarTypeToC = "void"
    Case VT_ERROR   ' 10
        VarTypeToC = "ERROR"
    Case VT_DECIMAL ' 14
        VarTypeToC = "DECIMAL"
    Case VT_I1   ' 16
        VarTypeToC = "CHAR"
    Case VT_UI1  ' 17
        VarTypeToC = "BYTE"
    Case VT_UI2  ' 18
        VarTypeToC = "WORD"
    Case VT_UI4  ' 19
        VarTypeToC = "LONG"
    Case VT_I8   ' 20
        VarTypeToC = "LONGLONG"
    Case VT_UI8  ' 21
        VarTypeToC = "ULONGLONG"
    Case VT_INT  ' 22
        VarTypeToC = "INT"
    Case VT_UINT ' 23
        VarTypeToC = "UINT"
    Case VT_HRESULT ' 25
        VarTypeToC = "HRESULT"
    Case VT_LPSTR ' 30
        VarTypeToC = "char"
    Case VT_LPWSTR ' 31
        VarTypeToC = "wchar_t"
'    Case VT_RECORD ' 36 ' not properly implemented, may not be needed, set in TKIND routine, get typedef name in TKIND routine
'        VarTypeToC = "void *"
    Case Else
        VarTypeToC = "Unknown" & Hex(vt)
        Print #99, "VarTypeToC: Unimplemented VarType: "; vt
        MsgBox "VarTypeToC: Unimplemented VarType: " & vt
        Err.Raise 1 ' Internal error - Unknown VarType
    End Select
If vt And VT_VECTOR Then VarTypeToC = VarTypeToC & " *"
End If
End Function

'Function isForiegnAlpha(ByVal c As String) As Boolean
'isForiegnAlpha = InStr(UCase("ãçíõÚ"), UCase(c)) > 0
'End Function

' fixme: isalpha() doesn't handle non-English - Sub SalvaAlterações() in autoba1a\autobak.vbp
Function isalpha(ByVal c As String) As Boolean
isalpha = UCase(c) >= "A" And UCase(c) <= "Z"
End Function

Function isdigit(ByVal c As String) As Boolean
isdigit = c >= "0" And c <= "9"
End Function

Function isalnum(ByVal c As String) As Boolean
isalnum = isalpha(c) Or isdigit(c)
End Function

Function issym(ByVal c As String) As Boolean
issym = isalnum(c) Or c = "_"
End Function

Function AbbrDataTypeRef(ByVal dt As TliVarType, ByVal s As String) As String
AbbrDataTypeRef = AbbrDataType(dt Or VT_BYREF) & "(" & s & ")"
End Function

' fixme: Need to implement abbreviations for all automation types
Function AbbrDataType(ByVal dt As TliVarType) As String
'Select Case Not (dt Imp VT_BYREF)
Select Case dt And Not (VT_ARRAY Or VT_BYREF Or VT_VECTOR)
    Case 0, VT_VOID
        AbbrDataType = "Void"
    Case vbByte
        AbbrDataType = "Byte"
    Case vbBoolean
        AbbrDataType = "Bool"
    Case vbInteger
        AbbrDataType = "Int"
    Case vbLong
        AbbrDataType = "Lng"
    Case vbSingle
        AbbrDataType = "Sng"
    Case vbDate
        AbbrDataType = "Date"
    Case vbDouble
        AbbrDataType = "Dbl"
    Case vbCurrency
        AbbrDataType = "Cur"
    Case vbString
        AbbrDataType = "Str"
    Case vbObject
        AbbrDataType = "Obj"
    Case vbVariant
        AbbrDataType = "Var"
    Case vbDecimal
        AbbrDataType = "Dec"
    Case vbUserDefinedType
        AbbrDataType = "UDT"
    Case VT_HRESULT ' 25 - Form1.Show needs this
        AbbrDataType = "HResult"
    Case VT_UNKNOWN ' 13
        AbbrDataType = "IUnknown"
    Case VT_UI2 ' 18
        AbbrDataType = "UINT16"
    Case VT_UI4 ' 19
        AbbrDataType = "UINT32"
    Case VT_INT ' 22
' fixme: name conflict - Int (vtInteger) vs VT_INT - use INT for now
        AbbrDataType = "INT"
    Case VT_UINT ' 23
        AbbrDataType = "UINT"
    Case VT_LPSTR ' 30
        AbbrDataType = "char"
    Case VT_LPWSTR ' 31
        AbbrDataType = "StrN"
    Case Else
        Print #99, "AbbrDataType: Unimplemented data type: "; dt
        AbbrDataType = "VT" & Right("000" & Hex(dt), 4)
        MsgBox "AbbrDataType: Unimplemented data type: " & dt
        Err.Raise 1 ' Undefined data type
End Select
' new - SA are always Refs so don't output SARef
'AbbrDataType = AbbrDataType & IIf(dt And (VT_ARRAY Or VT_VECTOR), "SA", "") & IIf(dt And VT_BYREF, "Ref", "")
AbbrDataType = AbbrDataType & IIf(dt And (VT_ARRAY Or VT_VECTOR), "SA", IIf(dt And VT_BYREF, "Ref", ""))
End Function

' fixme: Similar to InitVarTypeFromVarTypeInfo in vbt so combine
Function VarTypeInfoToVarType(ByVal vi As VarTypeInfo) As TliVarType
If vi Is Nothing Then
    Print #99, "VarTypeInfoToVarType: vi is Nothing"
    MsgBox "VarTypeInfoToVarType: vi is Nothing"
    Err.Raise 1
End If
Print #99, "VarTypeInfoToVarType: vt="; vi.varType
Select Case vi.varType And Not (VT_ARRAY Or VT_BYREF Or VT_VECTOR)
    Case 0
'        VarTypeInfoToCType = vi.TypeInfo.Name
        If vi.TypeInfo Is Nothing Then
            Print #99, "VarTypeInfoToVarType: vi.TypeInfo is Nothing"
            MsgBox "VarTypeInfoToVarType: vi.TypeInfo is Nothing"
            Err.Raise 1
        End If
        Print #99, "VarTypeInfoToVarType: tk="; vi.TypeInfo.TypeKind
' fixme: create a TKindToCDataType routine, this is oft repeated
        Select Case vi.TypeInfo.TypeKind
            Case TKIND_RECORD ' 1
                VarTypeInfoToVarType = VT_RECORD
            Case TKIND_ENUM
                VarTypeInfoToVarType = VT_I4
            Case TKIND_DISPATCH
            Print #99, "in disp"
                VarTypeInfoToVarType = VT_DISPATCH ' 4
            Print #99, "after disp"
            Case TKIND_INTERFACE
                VarTypeInfoToVarType = VT_UNKNOWN
            Case TKIND_COCLASS ' 5
                VarTypeInfoToVarType = VT_DISPATCH
            Case Else
                Print #99, "VarTypeInfoToVarType: Unimplemented type: 0x"; Hex(vi.TypeInfo.TypeKind)
                MsgBox "VarTypeInfoToVarType: Unimplemented type: 0x" & Hex(vi.TypeInfo.TypeKind)
                Err.Raise 1 ' TypeKind not implemented or invalid
        End Select
    Case Else
        ' removed And Not for IRR(d(),1)
        VarTypeInfoToVarType = vi.varType ' And Not (VT_ARRAY Or VT_BYREF Or VT_VECTOR)
End Select
Print #99, "VarTypeInfoToVarType: dt="; VarTypeInfoToVarType
End Function

Function EmitRPN(ByVal operand_stack As Collection) As String
Dim os As vbToken
For Each os In operand_stack
    If os.tokString = "" Then
        Print #99, "EmitRPN: tokString is empty: " & EmitRPN
        MsgBox "EmitRPN: tokString is empty: " & EmitRPN
        Err.Raise 1
    End If
    EmitRPN = EmitRPN & os.tokString & " "
Next
Print #99, "EmitRPN: {"; EmitRPN; "}"
End Function

Function ProcNameIK(ByVal n As String, ByVal ik As InvokeKinds) As String
Select Case ik
    Case INVOKE_FUNC, INVOKE_FUNC Or INVOKE_PROPERTYGET, INVOKE_UNKNOWN, INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF
        ProcNameIK = n
    Case INVOKE_PROPERTYGET
        ProcNameIK = n & "Get"
    Case INVOKE_PROPERTYPUT
        ProcNameIK = n & "Let"
    Case INVOKE_PROPERTYPUTREF
        ProcNameIK = n & "Set"
    Case INVOKE_EVENTFUNC
        ProcNameIK = n & "Event"
    Case Else
        Print #99, "ProcNameIK: Unimplemented InvokeKind: " & n & ik
        MsgBox "ProcNameIK: Unimplemented InvokeKind: " & n & ik
        Err.Raise 1 ' Unimplemented InvokeKind
End Select
End Function

Function VBTOutputPath(ByVal group As vbGroup, ByVal p As vbPrj, ByVal subdir As String) As String
On Error Resume Next
Print #99, "subdir="; subdir
Print #99, "group.OutputPath="; group.OutputPath
Print #99, "group.Name="; group.Name
VBTOutputPath = group.OutputPath & group.Name & "\"
MkDir VBTOutputPath
If p Is Nothing Then
    VBTOutputPath = VBTOutputPath & group.vbInstance.VBProjects.StartProject.Name & "\"
Else
    Print #99, "BuildFileName="; p.VBProject.BuildFileName
    Print #99, "FileName="; p.VBProject.FileName
    Print #99, "Name="; p.VBProject.Name
    Print #99, "FullName="; p.VBProject.VBE.FullName
    Print #99, "LastUsedPath="; p.VBProject.VBE.LastUsedPath
    Print #99, "StartUpProject.BuildFileName="; group.vbInstance.VBProjects.StartProject.BuildFileName
'    OutputPath = group.vbInstance.VBProjects.StartProject.BuildFileName
'    i = InStr(1, OutputPath, ".")
'    If i > 0 Then OutputPath = Left(OutputPath, i - 1) & "\" & p.prjName & "\c\"
'    OutputPath = p.VBProject.VBE.LastUsedPath & "\c\"
'    On Error Resume Next
'    MkDir OutputPath
'    On Error GoTo 0
    VBTOutputPath = VBTOutputPath & p.prjName & "\"
End If
MkDir VBTOutputPath
VBTOutputPath = VBTOutputPath & subdir & "\"
MkDir VBTOutputPath
On Error GoTo 0
Print #99, "VBTOutputPath="; VBTOutputPath
End Function

Function GetReturnType(ByVal ii As InterfaceInfo, ByVal mi As MemberInfo, ByRef rtdt As TliVarType, ByRef rtii As InterfaceInfo) As VarTypeInfo
Print #99, "GetReturnType: ii="; Not ii Is Nothing; " mi="; Not mi Is Nothing
If Not ii Is Nothing Then
    On Error GoTo badlib
    Print #99, "GetReturnType: ii="; ii.Name
    Print #99, " guid="; ii.GUID
    Print #99, " am="; Hex(ii.AttributeMask)
    Print #99, " tk="; ii.TypeKind
'    Print #99, " tlib="; ii.Parent.ContainingFile
' causes ElitePad to fail
''''    Print #99, " tlib="; ii.Parent.GUID
''''    Print #99, " tin="; ii.TypeInfoNumber ' This caused djxxxx.vbp projects to fail. Projects were deleted because bad component was suspected.
    On Error GoTo 0
    While False
badlib:
        On Error GoTo 0
        Print #99, "Unable to traverse typelib for interface: "; ii.Name
        MsgBox "Unable to traverse typelib for interface: " & ii.Name
        Err.Raise 1
    Wend
End If
Print #99, "GetReturnType: mi="; mi.Name; " am="; Hex(mi.AttributeMask); " ik="; mi.InvokeKind; " pc="; mi.parameters.Count; " dk="; mi.DescKind; " mi.vto="; mi.VTableOffset; " rt="; Not mi.ReturnType Is Nothing
If mi.ReturnType Is Nothing Then Err.Raise 1
''''Print #99, "mi.rt.Ex="; mi.ReturnType.IsExternalType; " pl="; mi.ReturnType.PointerLevel; " ti="; Not mi.ReturnType.TypeInfo Is Nothing
' fixme: this is confused. There should be tokInterfaceInfo and tokReturnTypeInterfaceInfo
Print #99, "2"
Dim pi As ParameterInfo
' could have zero paramaters
For Each pi In mi.parameters
    Print #99, "pi.n="; pi.Name; " pi.f="; Hex(pi.flags)
    If pi.flags And PARAMFLAG_FRETVAL Then
'            If Not pi Is mi.Parameters.Item(mi.Parameters.count) Then ' don't understand why this doesn't work
        If pi.Name <> mi.parameters.Item(mi.parameters.Count).Name Then
            Print #99, "ReturnType is not last parameter: pi.n="; pi.Name; " pi.f="; Hex(pi.flags); " n="; mi.parameters.Item(mi.parameters.Count).Name; " f="; Hex(mi.parameters.Item(mi.parameters.Count).flags)
            MsgBox "ReturnType is not last parameter"
        End If
        If pi.VarTypeInfo Is Nothing Then Err.Raise 1
        Exit For
    End If
Next
Print #99, "3"
If pi Is Nothing Then
    Print #99, "4"
    ' ADODB_Command15_CreateParameter needs INVOKE_FUNC
    If mi.InvokeKind And INVOKE_FUNC Then
        Print #99, "5"
        If (mi.ReturnType.varType <> VT_HRESULT And mi.ReturnType.varType <> VT_VOID) Or mi.ReturnType.PointerLevel <> 0 Then Set GetReturnType = mi.ReturnType
    ElseIf mi.InvokeKind And INVOKE_PROPERTYGET Then
        Print #99, "5a"
        Set GetReturnType = mi.ReturnType
        If GetReturnType Is Nothing Then Err.Raise 1
        If GetReturnType.varType = VT_VOID Then Err.Raise 1
    ElseIf mi.InvokeKind And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
        ' There doesn't seem to be any flags that reliably indicate whether last parameter is a putval
        ' If last parameter is not a putval, then assign putval to ReturnType
        Print #99, "5b"
'            If mi.Parameters.Count = 0 Then
        If mi.ReturnType.varType <> VT_VOID And mi.ReturnType.varType <> VT_HRESULT Then
            Print #99, "5c"
            Set GetReturnType = mi.ReturnType
            If GetReturnType Is Nothing Then Err.Raise 1
            ' "time = 0" is a propput with a ReturnType of VARIANT*. Doesn't seem possible.
'            Dim butnotifthis As Boolean
'            butnotifthis = True
        Else
            If Not CBool(mi.parameters.Item(mi.parameters.Count).flags And PARAMFLAG_FIN) Then Err.Raise 1
            Set GetReturnType = mi.parameters.Item(mi.parameters.Count).VarTypeInfo
        End If
    ElseIf mi.InvokeKind = INVOKE_UNKNOWN Then ' Help.vbp
        Print #99, "5d"
        Set GetReturnType = mi.ReturnType
    Else ' VT_VOID
        Print #99, "5f"
        ' Set GetReturnType = Nothing ' default is Nothing
'            Set GetReturnType = mi.ReturnType ' might happen when INVOKE_UNKNOWN - Text1.Font.Name
    End If
Else
    Print #99, "6a"
    Print #99, "pi="; Not pi Is Nothing
    Print #99, "vti="; Not pi.VarTypeInfo Is Nothing
    Set GetReturnType = pi.VarTypeInfo
    Print #99, "7"
End If
'Set rtii = ii
Print #99, "8"
If GetReturnType Is Nothing Then
Print #99, "9"
    rtdt = VT_VOID
ElseIf GetReturnType.varType = 0 Then
Print #99, "10 extern="; GetReturnType.IsExternalType
    Dim ti As TypeInfo
    If GetReturnType.IsExternalType Then Set ti = GetReturnType.TypeLibInfoExternal.TypeInfos.IndexedItem(GetReturnType.TypeInfoNumber) Else Set ti = GetReturnType.TypeInfo
    If ti Is Nothing Then Err.Raise 1
Print #99, "11"
    Print #99, "GetReturnType: ti: n="; ti.Name; " tk="; ti.TypeKind
    Select Case ti.TypeKind
    Case TKIND_INTERFACE
        Set rtii = ti
        rtdt = VT_UNKNOWN
        Print #99, "rtii="; TypeInfoToVBType(rtii); " am="; Hex(rtii.AttributeMask); " tk="; rtii.TypeKind; " implied interface count="; rtii.ImpliedInterfaces.Count
    Case TKIND_DISPATCH
        Set rtii = ti
        rtdt = vbObject
        Print #99, "rtii="; TypeInfoToVBType(rtii); " am="; Hex(rtii.AttributeMask); " tk="; rtii.TypeKind; " implied interface count="; rtii.ImpliedInterfaces.Count
    Case TKIND_COCLASS
'        Print #99, "ti is a ci"
        Set rtii = ti.DefaultInterface
        rtdt = vbObject
        Print #99, "rtii="; TypeInfoToVBType(rtii); " am="; Hex(rtii.AttributeMask); " tk="; rtii.TypeKind; " implied interface count="; rtii.ImpliedInterfaces.Count
    Case TKIND_ENUM
        Print #99, "ti is a enum"
        Print #99, "mem.count="; ti.Members.Count
        #If 0 Then
        Dim mem As MemberInfo
        For Each mem In ti.Members
            Print #99, "n="; mem.Name; " v="; mem.Value; " vt="; varType(mem.Value)
        Next
        On Error Resume Next
        Print #99, "ResolvedType="; ti.ResolvedType.varType
        On Error GoTo 0
        #End If
'        Set rtii = Nothing
'        Set rtii = ti
        rtdt = vbLong ' hard code - don't know how to obtain
'            Print #99, "custinfo="; ti.CustomDataCollection.count
'            Dim cd As CustomData
'            For Each cd In ti.CustomDataCollection
'                Print #99, "cd="; cd.GUID; cd.Value
'            Next
    Case TKIND_RECORD
'        Set rtii = Nothing
        rtdt = vbUserDefinedType
    Case Else
        Print #99, "GetReturnType: unknown typeinfo tkind: "; ti.Name
        MsgBox "GetReturnType: unknown typeinfo: tkind: " & ti.Name
        Err.Raise 1
    End Select
Else ' VarType <> 0
Print #99, "12 vt="; GetReturnType.varType; " pl="; GetReturnType.PointerLevel
    If Not GetReturnType.TypeInfo Is Nothing Then
        Print #99, "GetReturnType: ti=" & GetReturnType.TypeInfo.Name
        MsgBox "GetReturnType: ti=" & GetReturnType.TypeInfo.Name
        Err.Raise 1
    End If
Print #99, "13"
    rtdt = GetReturnType.varType
    ' howtos1r\graphv~1.vbp
'    If mi.InvokeKind = INVOKE_PROPERTYPUT And GetReturnType.PointerLevel = 1 And Not butnotifthis Then rtdt = rtdt Or VT_BYREF ' Property Let and putval is ByRef
    If mi.InvokeKind = INVOKE_PROPERTYPUT And GetReturnType.PointerLevel = 1 Then rtdt = rtdt Or VT_BYREF ' Property Let and putval is ByRef
'    Set rtii = Nothing
End If
If Not rtii Is Nothing Then Print #99, "guid="; rtii.GUID
Print #99, "GetReturnType: done: rtdt="; rtdt
End Function

' fixme: There must be a better way of checking if value is in list
Function IsAny(ByVal Value As Variant, ParamArray a()) As Integer
Dim v As Variant
For Each v In a
    If v = Value Then IsAny = True: Exit For
Next
End Function

Function IsObj(ByVal vt As TliVarType) As Boolean
vt = vt And Not VT_BYREF
IsObj = vt = VT_DISPATCH Or vt = VT_UNKNOWN
End Function

Function IsVObj(ByVal vt As TliVarType) As Boolean
IsVObj = IsObj(vt) Or vt = VT_VARIANT
End Function

Function IsParamArrayArg(ByVal mi As MemberInfo, ByVal pc As Integer) As Boolean
If mi.parameters.OptionalCount = -1 And pc >= mi.parameters.Count Then
    If mi.parameters.Item(mi.parameters.Count).VarTypeInfo.varType <> (VT_VARIANT Or VT_ARRAY) Then Err.Raise 1
    If mi.parameters.Item(mi.parameters.Count).VarTypeInfo.PointerLevel <> 1 Then Err.Raise 1
    IsParamArrayArg = True
End If
End Function

Sub ValidateReferenceTLibs(ByVal refs As References)
Print #99, "ValidateReferenceTLibs: c="; refs.Count
Dim tli As TypeLibInfo
Dim ref As reference
On Error Resume Next
For Each ref In refs
    Print #99, "CheckReferenceTLibs: ref="
    Dim s As String
    Print #99, ref.Name
    s = ref.Name
    Print #99, " path="
    Print #99, ref.FullPath ' missing FullPath indicates bad registration
    s = ref.FullPath
    Print #99, " guid="
    Print #99, ref.GUID
    Print #99, " broken="
    Print #99, ref.IsBroken
    Print #99, " major="
    Print #99, ref.Major
    Print #99, " minor="
    Print #99, ref.Minor
    Set tli = TypeLibInfoFromRegistry(ref.GUID, ref.Major, ref.Minor, 0)
    If tli Is Nothing Then Set tli = TypeLibInfoFromFile(ref.FullPath)
    If tli Is Nothing Then ' amazin1a\project1.vbp (msflxgrd), axmrquee\axmrquee.vbp (axmrquee.ocx)
        MsgBox "Invalid Project Reference or Component Library. Library may be damaged, missing, or incorrectly registered. Removing libary reference, reinstalling application, library, or Visual Studio may resolve problem. Bad library is: " & s
        On Error GoTo 0
        Err.Raise 1
    Else
        Set tli = Nothing
    End If
Next
Print #99, "ValidateReferenceTLibs: done"
End Sub

Function TypeInfoToVBType(ByVal ti As TypeInfo) As String
Print #99, "TypeInfoToVBType: "; ti.Parent.Name & "." & ti.Name
TypeInfoToVBType = ti.Parent.Name & "." & ti.Name
End Function

Function SymIK(ByVal sym As String, ByVal ik As InvokeKinds) As String
SymIK = UCase(sym) & "." & ik ' merge symbol name and ik, separate with invalid symbol character
End Function

' fixme: replace this logic with Variant change type API call
Function CoerceConstant(ByVal v As Variant, ByVal dt As TliVarType, Optional ByVal l As Long = -1) As Variant
Print #99, "CoerceConstant: v="; varType(v); " dt="; dt; " l="; l
On Error GoTo invalid_conversion ' This should by caught VB but isn't (Bool = "")
Select Case dt
Case vbEmpty
    ' do nothing
Case vbBoolean
    CoerceConstant = CBool(v)
Case vbByte
    CoerceConstant = CByte(v)
Case vbInteger
    CoerceConstant = CInt(v)
Case vbLong ' , VT_VOID Or VT_BYREF ' As Any
    CoerceConstant = CLng(v)
Case vbSingle
    CoerceConstant = CSng(v)
Case vbDouble
    CoerceConstant = CDbl(v)
Case vbDate
    CoerceConstant = CDate(v)
Case vbCurrency
    CoerceConstant = CCur(v)
Case vbString
    CoerceConstant = CStr(v)
Case vbVariant
    CoerceConstant = v
Case VT_LPWSTR
    If l = -1 Then GoTo invalid_conversion
    CoerceConstant = Left(CStr(v), l)
Case VT_UI2, VT_UI4, VT_INT, VT_UINT
    ' fixme: not sure how to handle non-vb types (UI2,UI4). Needs testing.
Case Else
    ' non-vb constants originate in TypeLibs
invalid_conversion:
    On Error GoTo 0
    ' Could be syntax error uncaught by VB, such as attempting to pass a string parameter as a numeric.
    ' MsgBox a$ + 1 becomes MsgBox CStr(CDbl(a$)+1)
        Print #99, "CoerceConstant: Unsupported conversion or parameter coercion. Invalid data type conversion.: " & AbbrDataType(varType(v)) & "To" & AbbrDataType(dt)
        MsgBox "Unsupported conversion or parameter coercion. Invalid data type conversion.: " & AbbrDataType(varType(v)) & "To" & AbbrDataType(dt)
        Err.Raise 1 ' Invalid numeric conversion
End Select
Print #99, "CoerceConstant: vt="; varType(CoerceConstant)
End Function

Attribute VB_Name = "vbtCoerce"
Option Explicit

' Keep cnt as ByRef, its incremented
' Implement optimize flag?
Function CoerceOperand(ByVal OptimizeFlag As OptimizeFlags, ByVal output_stack As Collection, ByRef cnt As Long, ByVal dt As TliVarType, Optional ByVal CoerceInterfaceInfo As InterfaceInfo = Nothing, Optional ByVal CoerceModule As vbModule = Nothing, Optional ByVal ik As InvokeKinds = INVOKE_FUNC Or INVOKE_PROPERTYGET, Optional ByVal NoInsertObjDefault As Boolean) As TliVarType
Dim CvtToken As New vbToken
Print #99, "CoerceOperand: dt=" & dt & " cnt=" & cnt & " s=" & output_stack.Item(cnt).tokString & " operand=" & output_stack.Item(cnt).tokDataType & " toktype=" & output_stack.Item(cnt).tokType; " v="; Not output_stack.Item(cnt).tokVariable Is Nothing; " ii="; Not CoerceInterfaceInfo Is Nothing; " m="; Not CoerceModule Is Nothing; " ik="; ik; " niod="; NoInsertObjDefault
If Not IsObject(output_stack.Item(cnt).tokValue) Then Print #99, "CoerceOperand: val="; output_stack.Item(cnt).tokValue
Print #99, "1"
CoerceOperand = dt
If dt = -1 Then
    Print #99, "1a"
    If Not NoInsertObjDefault And (output_stack.Item(cnt).tokDataType And Not (VT_ARRAY Or VT_BYREF)) = vbObject Then
        Print #99, "1b"
' unneeded, checked in CoerceDefaultMember - If output_stack.Item(cnt).tokType <> tokNothing Then CoerceDefaultMember output_stack, cnt, CoerceOperand, ik
        CoerceDefaultMember output_stack, cnt, CoerceOperand, ik
        Print #99, "1c"
    ElseIf (output_stack.Item(cnt).tokDataType And Not (VT_ARRAY Or VT_BYREF)) = VT_LPWSTR And CBool(ik And INVOKE_PROPERTYPUT) Then
        CoerceOperand = VT_LPWSTR
    End If
    Print #99, "1d"
ElseIf (dt And Not VT_BYREF) = (output_stack.Item(cnt).tokDataType And Not VT_BYREF) Then
    ' Data types are same. Just check if QI and ByRef are needed.
    GoTo check_QI_byrefs
ElseIf dt <> VT_VOID Then
    Print #99, "2"
    If dt = 0 Then Err.Raise 1 ' Invalid data type
    Print #99, "2a"
    If (dt And Not VT_ARRAY) = (VT_VOID Or VT_BYREF) Then
        ' The "As Any" datatype is VT_VOID Or VT_BYREF and maybe VT_ARRAY
        CoerceOperand = dt And Not VT_ARRAY
    ElseIf dt And VT_ARRAY Then
        Print #99, "2b"
 '       If (dt And Not VT_BYREF) <> (output_stack.Item(cnt).tokDataType And Not VT_BYREF) Then
            Select Case output_stack.Item(cnt).tokDataType And Not VT_BYREF
            Case VT_VARIANT
                ' Variants can be coerced to SAFEARRAYs. VarToByteSA() ' multgrid\GridSamp.vbp
                GoTo explicit_conversion
            Case VT_BSTR Or VT_ARRAY
                ' Strings can be coerced to Byte arrays - b() = "string"
                If (dt And Not (VT_ARRAY Or VT_BYREF)) <> VT_UI1 Then Err.Raise 1
                GoTo explicit_conversion
            Case Else
                Err.Raise 1
            End Select
'        End If
        Print #99, "2c"
    ElseIf output_stack.Item(cnt).tokDataType And VT_ARRAY Then
        ' converting array to non-array - must be Variant
        If (dt And Not VT_BYREF) <> VT_VARIANT Then Err.Raise 1
        GoTo explicit_conversion
    Else
        CoerceOperand = dt And Not VT_BYREF
    End If
    Print #99, "3 CoerceOperand="; CoerceOperand; output_stack.Item(cnt).tokDataType And Not (VT_ARRAY Or VT_BYREF)
    ' Not vt.varDimensions Is Nothing forces into SAFEARRAY
    If Not output_stack.Item(cnt).tokVariable Is Nothing Then
        If Not output_stack.Item(cnt).tokVariable.varDimensions Is Nothing Then GoTo 10
        ' Always convert fixed length strings
'        Print #99, "dtlength="; output_stack.Item(cnt).tokVariable.varType.dtLength; " varlength="; output_stack.Item(cnt).tokVariable.varLength
'        If CoerceOperand = vbString And output_stack.Item(cnt).tokVariable.varType.dtLength > 0 Then GoTo convertFixedString
    End If
    If CoerceOperand <> (output_stack.Item(cnt).tokDataType And Not (VT_ARRAY Or VT_BYREF)) Then
10
        Print #99, "4"
        If output_stack.Item(cnt).tokType = tokNothing Then
            ' Picture = Nothing
            output_stack.Item(cnt).tokDataType = CoerceOperand
        ElseIf Not NoInsertObjDefault And (output_stack.Item(cnt).tokDataType And Not (VT_ARRAY Or VT_BYREF)) = vbObject And CoerceOperand <> vbVariant And CoerceOperand <> VT_UNKNOWN Then
            Print #99, "5"
' removed, for case of coercing object to string (has no ii)
'            If Not CoerceInterfaceInfo Is Nothing Then
                If Not CoerceDefaultMember(output_stack, cnt, CoerceOperand, ik) Then
                    Print #99, "Unable to coerce default member"
                    MsgBox "Unable to coerce default member"
                    Err.Raise 1 ' Unable to coerce default member
                End If
'            End If
            Print #99, "6"
        ElseIf (output_stack.Item(cnt).tokDataType And Not (VT_ARRAY Or VT_BYREF)) = VT_LPWSTR And CBool(ik And INVOKE_PROPERTYPUT) Then
            ' c = mid("",1)
            CoerceOperand = VT_LPWSTR
#If 1 Then
'        ElseIf (ik And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) And (output_stack.Item(cnt).tokDataType And Not (VT_ARRAY Or VT_BYREF)) = vbVariant And Not CBool(output_stack.Item(cnt).tokPCodeSubType And ik) Then
        ElseIf CBool(ik And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) And (output_stack.Item(cnt).tokDataType And Not (VT_ARRAY Or VT_BYREF)) = vbVariant And Not CBool(output_stack.Item(cnt).tokPCodeSubType And ik) Then
            ' col(1) = col(2) has only INVOKE_PROPERTYGET, col.item(1) = col.item(2)
'            output_stack.Item(cnt).tokType = tokInvokeDefaultMember
'            CoerceObject output_stack, cnt, vbObject, Nothing, Nothing
'            output_stack(cnt).tokDataType = vbVariant
            output_stack.Item(cnt).tokPCodeSubType = 0 ' clear PROPERTYPUT/REF, default will carry flag
            output_stack.Add DispatchDefaultMember(ik), , , cnt
            cnt = cnt + 1
#End If
        ElseIf OptimizeFlag <> OptimizeNone And CoerceOperand <> VT_LPWSTR And CoerceOperand <> vbVariant And CoerceOperand <> (VT_VOID Or VT_BYREF) And HasValue(output_stack.Item(cnt)) Then
Print #99, "7"
            output_stack.Item(cnt).tokValue = CoerceConstant(output_stack.Item(cnt).tokValue, CoerceOperand)
Print #99, "8"
            If CoerceOperand <> varType(output_stack.Item(cnt).tokValue) Then GoTo explicit_conversion
Print #99, "9"
            ' if token was Const or Enum, make it a Variant
            output_stack.Item(cnt).tokType = tokVariant
            output_stack.Item(cnt).tokDataType = CoerceOperand
Print #99, "10"
'        ElseIf (dt And Not (VT_ARRAY Or VT_BYREF)) <> (output_stack.Item(cnt).tokDataType And Not (VT_ARRAY Or VT_BYREF)) Then ' skip if dt is same and both are arrays
        ElseIf (dt And Not VT_BYREF) <> (output_stack.Item(cnt).tokDataType And Not VT_BYREF) Then ' skip if dt is same and both are arrays
explicit_conversion:
Print #99, "11"
            CvtToken.tokString = AbbrDataType(output_stack.Item(cnt).tokDataType) & "To" & AbbrDataType(CoerceOperand)
            CvtToken.tokType = tokCvt
            CvtToken.tokPCode = vbPCodeCvt
            ' don't want VT_ARRAY anymore - debug.print Array(a)
'            CvtToken.tokDataType = IIf(dt <> vbVariant And CBool(dt And vbArray), dt, CoerceOperand)
            ' Any1.vbp
            ' remove VT_BYREF when originally not a Ref, AddRef below (Var to Any)
            CvtToken.tokDataType = CoerceOperand
            CvtToken.tokPCodeSubType = output_stack.Item(cnt).tokDataType
            output_stack.Add CvtToken, , , cnt
            cnt = cnt + 1 ' increment pointer past inserted Cvt token
Print #99, "12"
        End If
Print #99, "13"
    End If
check_QI_byrefs:
Print #99, "14"
'    If dt = vbObject Then CoerceObject output_stack, cnt, dt, CoerceInterfaceInfo, CoerceModule
    If IsObj(CoerceOperand) Then CoerceObject output_stack, cnt, dt, CoerceInterfaceInfo, CoerceModule
    Print #99, "15"
#If 1 Then
    ' fixme: put in routine (used elsewhere) - implement tokPointerLevel to optimize variables
    If (output_stack.Item(cnt).tokDataType And VT_BYREF) <> 0 And (dt And VT_BYREF) = 0 Then
        OutputSubRef output_stack, cnt
    ElseIf (output_stack.Item(cnt).tokDataType And VT_BYREF) = 0 And (dt And VT_BYREF) <> 0 Then
        OutputAddRef output_stack, cnt
    End If
#End If
    Print #99, "16 co="; CoerceOperand; " os.dt="; output_stack.Item(cnt).tokDataType
    ' VarToByteSA - CoerceOperand uses (VT_ARRAY + VT_BYREF)
    If CoerceOperand <> (VT_VOID Or VT_BYREF) And (CoerceOperand And Not VT_BYREF) <> (output_stack.Item(cnt).tokDataType And Not VT_BYREF) Then Err.Raise 1
End If
Print #99, "CoerceOperand: new dt=" & output_stack.Item(cnt).tokDataType
End Function

' Compare the stacked return value interface (GUID) to newly derived interface (GUID)
Sub CoerceObject(ByVal output_stack As Collection, ByRef cnt As Long, ByVal dt As TliVarType, ByVal CoerceInterfaceInfo As InterfaceInfo, ByVal CoerceModule As vbModule)
Print #99, "CoerceObject: c="; cnt; " dt="; dt; " ii="; Not CoerceInterfaceInfo Is Nothing; " cm="; Not CoerceModule Is Nothing
Print #99, "s="; output_stack.Item(cnt).tokString; " t="; output_stack.Item(cnt).tokType; " os.dt="; output_stack.Item(cnt).tokDataType; " os.pst="; output_stack.Item(cnt).tokPCodeSubType; " v="; Not output_stack.Item(cnt).tokVariable Is Nothing
'If Not IsObj(output_stack.Item(cnt).tokDataType And Not VT_BYREF) Then Err.Raise 1
Dim GUID As String
If Not output_stack.Item(cnt).tokInterfaceInfo Is Nothing Then Print #99, "os.ii="; output_stack.Item(cnt).tokInterfaceInfo.Name
If Not output_stack.Item(cnt).tokMemberInfo Is Nothing Then Print #99, "os.mi="; output_stack.Item(cnt).tokMemberInfo.Name
If output_stack.Item(cnt).tokVariable Is Nothing Then GoTo 10
Print #99, "vt.type="; output_stack.Item(cnt).tokVariable.varType.dtType
' fixme: yuk, want to handle consistently
If Not output_stack.Item(cnt).tokVariable.varType.dtInterfaceInfo Is Nothing Then
    Print #99, "vt.ii="; output_stack.Item(cnt).tokVariable.varType.dtInterfaceInfo.Name
    ' need to FormQI(Form1)
    If output_stack.Item(cnt).tokType <> tokQI_TLibInterface And output_stack.Item(cnt).tokVariable.varType.dtType = tokFormClass Then
' Must generate QI to change Form1 or UserControl1 into Form, UserControl interfaces
' This is done by leaving GUID empty and forcing generation of QI below.
' fixme: This is needed because projects module GUID is matching UserControl interface GUID which is probably wrong.
' Removed following line. Why was it needed?
''''        If output_stack.Item(cnt).tokVariable.varType.dtClass.Component.Type = vbext_ct_ActiveXDesigner Or output_stack.Item(cnt).tokVariable.varType.dtClass.Component.Type = vbext_ct_UserControl Then GoTo iiGUID 'Exit Sub
    Else
''''iiGUID:
        Dim ii As InterfaceInfo
        Dim iiv As InterfaceInfo
        Set ii = output_stack.Item(cnt).tokVariable.varType.dtInterfaceInfo
        Set iiv = Nothing
        On Error Resume Next
        Set iiv = ii.VTableInterface
        On Error GoTo 0
        If iiv Is Nothing Then GUID = ii.GUID Else GUID = iiv.GUID
    End If
Else
'    If output_stack.Item(cnt).tokVariable Is Nothing Then GoTo 10
    If Not output_stack.Item(cnt).tokVariable.varType.dtClass Is Nothing Then
        Print #99, "Class="; output_stack.Item(cnt).tokVariable.varType.dtClass.Name
        GUID = output_stack.Item(cnt).tokVariable.varType.dtClass.interfaceGUID
    ElseIf Not output_stack.Item(cnt).tokVariable.varType.dtClassInfo Is Nothing Then
        Print #99, "InterfaceInfo="; output_stack.Item(cnt).tokVariable.varType.dtInterfaceInfo.Name
        Set ii = output_stack.Item(cnt).tokVariable.varType.dtInterfaceInfo
        Set iiv = Nothing
        On Error Resume Next
        Set iiv = ii.VTableInterface
        On Error GoTo 0
        If iiv Is Nothing Then GUID = ii.GUID Else GUID = iiv.GUID
    Else
10
        If (output_stack.Item(cnt).tokDataType And Not VT_BYREF) = vbObject Then
            Print #99, "IDispatch"
            Set ii = DispatchInterfaceInfo
            GUID = ii.GUID
        ElseIf (output_stack.Item(cnt).tokDataType And Not VT_BYREF) = VT_UNKNOWN Then
            Print #99, "IUnknown"
            Set ii = UnknownInterfaceInfo
            GUID = ii.GUID
        Else
            Err.Raise 1 ' Internal error - unknown object type
        End If
    End If
End If
Print #99, "CoerceObject: GUID="; GUID; " ci="; Not CoerceInterfaceInfo Is Nothing; " cm="; Not CoerceModule Is Nothing
Dim CvtToken As vbToken
' looks like trouble - had to remove If test - SelectedControls.Item(0) where Item is a tlib interface member, not a tokFormClass, called by CoerceOperand
''''If Not CoerceInterfaceInfo Is Nothing And Not CoerceModule Is Nothing Then If output_stack.Item(cnt).tokVariable.varType.dtType <> tokFormClass Then Err.Raise 1
If Not CoerceInterfaceInfo Is Nothing Then
    Print #99, "cii: n="; CoerceInterfaceInfo.Name; " GUID="; CoerceInterfaceInfo.GUID
    Set ii = CoerceInterfaceInfo
    Set iiv = Nothing
    On Error Resume Next
    Set iiv = ii.VTableInterface
    On Error GoTo 0
    Dim ciiGUID As String
    If iiv Is Nothing Then ciiGUID = ii.GUID Else ciiGUID = iiv.GUID
    If Not CoerceInterfaceInfo.VTableInterface Is Nothing Then Print #99, "iivGUID="; CoerceInterfaceInfo.VTableInterface.GUID
    Print #99, "ciiGUID="; ciiGUID
    If GUID <> ciiGUID Then
        Set CvtToken = New vbToken
        CvtToken.tokString = TypeInfoToVBType(ii) & "QI"
        CvtToken.tokType = tokQI_TLibInterface
'                CvtToken.tokPCode = vbPCodeQITLibInterface
        CvtToken.tokDataType = dt And Not VT_BYREF
        Print #99, "CoerceObject: ii.am="; Hex(ii.AttributeMask)
        ' "To" Object data type must be dispatchable - right??
        ' perhaps should check if interface is IDispatch itself
        ' perhaps should if interface has an implied interface of IDispatch
        If (dt And Not VT_BYREF) = VT_DISPATCH Then
            If Not CBool(ii.AttributeMask And (1 Or TYPEFLAG_FRESTRICTED Or TYPEFLAG_FDISPATCHABLE)) Then
                Print #99, "Dispatchable interface expected but lacks attribute: "; ii.Name
                MsgBox "Dispatchable interface expected but lacks attribute: " & ii.Name
            End If
            ' "From" Object data type must be dispatchable - right??
            ' perhaps should check if interface is IDispatch itself
            ' perhaps should if interface has an implied interface of IDispatch
            If (output_stack(cnt).tokDataType And Not VT_BYREF) <> VT_DISPATCH Then
                If Not CBool(output_stack(cnt).tokInterfaceInfo.AttributeMask And (1 Or TYPEFLAG_FRESTRICTED Or TYPEFLAG_FDISPATCHABLE)) Then
                    Print #99, "Interface cannot be made dispatchable: "; output_stack(cnt).tokInterfaceInfo.Name
                    MsgBox "Interface cannot be made dispatchable: " & output_stack(cnt).tokInterfaceInfo.Name
                End If
            End If
        End If
        Set CvtToken.tokInterfaceInfo = ii
        Set CvtToken.tokVariable = output_stack.Item(cnt).tokVariable
        output_stack.Add CvtToken, , , cnt
        cnt = cnt + 1
        On Error Resume Next
        currentProject.CoerceObjects.Add CvtToken, ciiGUID
        On Error GoTo 0
    End If
ElseIf Not CoerceModule Is Nothing Then
    Print #99, "GUID="; GUID; " cm.interfaceGUID="; CoerceModule.interfaceGUID; " cm.Name="; CoerceModule.Name
'    If Not CoerceModule Is output_stack.Item(cnt).tokVariable.varType.dtClass Then
    Dim CoerceModuleGUID As String
    Dim CoerceModuleName As String
    ' fixme: kludge: should be passing InterfaceName/GUID instead of CoerceModule (which is a class)
'    If output_stack.Item(cnt).tokPCodeSubType = INVOKE_EVENTFUNC Then
'        CoerceModuleGUID = CoerceModule.EventGUID
'        CoerceModuleName = CoerceModule.EventName
'    Else
        CoerceModuleGUID = CoerceModule.interfaceGUID
        CoerceModuleName = CoerceModule.interfaceName
    Print #99, " cmg="; CoerceModuleGUID; " cmn="; CoerceModuleName
'    End If
'    If output_stack.Item(cnt).tokVariable Is Nothing Then GoTo 20 ' New, VarToObj has no tokVariable
    If GUID <> CoerceModuleGUID Then
'20
        Set CvtToken = New vbToken
        If CoerceModuleName = "" Then Err.Raise 1
        CvtToken.tokString = "_" & CoerceModuleName & "QI"
        CvtToken.tokType = tokQI_Module
'        CvtToken.tokPCode = vbPCodeQIModule
        CvtToken.tokDataType = dt And Not VT_BYREF
        Set CvtToken.tokModule = CoerceModule
        Set CvtToken.tokVariable = output_stack.Item(cnt).tokVariable
        output_stack.Add CvtToken, , , cnt
        cnt = cnt + 1
        On Error Resume Next
        currentProject.CoerceObjects.Add CvtToken, CoerceModuleGUID
        On Error GoTo 0
    End If
ElseIf (dt And Not VT_BYREF) = VT_UNKNOWN Then
    If GUID <> UnknownInterfaceInfo.GUID Then
        Set CvtToken = New vbToken
        CvtToken.tokString = "_" & UnknownInterfaceInfo.Name & "QI"
        CvtToken.tokType = tokQI_TLibInterface
'                CvtToken.tokPCode = vbPCodeQIDispatch
        CvtToken.tokDataType = dt And Not VT_BYREF
        Set CvtToken.tokInterfaceInfo = UnknownInterfaceInfo
        Set CvtToken.tokVariable = output_stack.Item(cnt).tokVariable
        output_stack.Add CvtToken, , , cnt
        cnt = cnt + 1
        On Error Resume Next
        ' fixme: put "from" interface into QI struct? Might become handy.
        currentProject.CoerceObjects.Add CvtToken, UnknownInterfaceInfo.GUID
        On Error GoTo 0
    End If
ElseIf (dt And Not VT_BYREF) = vbObject Then
    If GUID <> DispatchInterfaceInfo.GUID Then
        Set CvtToken = New vbToken
        CvtToken.tokString = "_" & DispatchInterfaceInfo.Name & "QI"
        CvtToken.tokType = tokQI_TLibInterface
'                CvtToken.tokPCode = vbPCodeQIDispatch
        CvtToken.tokDataType = dt And Not VT_BYREF
        Set CvtToken.tokInterfaceInfo = DispatchInterfaceInfo
        Set CvtToken.tokVariable = output_stack.Item(cnt).tokVariable
        output_stack.Add CvtToken, , , cnt
        cnt = cnt + 1
        On Error Resume Next
        currentProject.CoerceObjects.Add CvtToken, DispatchInterfaceInfo.GUID
        On Error GoTo 0
    End If
Else
    Err.Raise 1 ' Internal error - unknown object type
End If
Print #99, "CoerceObject: done"
End Sub

Function CoerceDefaultMember(ByVal output_stack As Collection, ByRef cnt As Long, ByVal dt As TliVarType, Optional ByVal ik As InvokeKinds = INVOKE_FUNC Or INVOKE_PROPERTYGET) As Boolean
Print #99, "CoerceDefaultMember: dt=" & dt & " cnt=" & cnt & " s=" & output_stack.Item(cnt).tokString & " operand=" & output_stack.Item(cnt).tokDataType & " type=" & output_stack.Item(cnt).tokType; " pst=" & output_stack.Item(cnt).tokPCodeSubType & " ii=" & (Not output_stack.Item(cnt).tokInterfaceInfo Is Nothing) & " mi=" & (Not output_stack.Item(cnt).tokMemberInfo Is Nothing) & " ik="; ik & " v=" & Not output_stack.Item(cnt).tokVariable Is Nothing
If output_stack.Item(cnt).tokType = tokNothing Then CoerceDefaultMember = True: Exit Function
If output_stack.Item(cnt).tokVariable Is Nothing Then Err.Raise 1
If output_stack.Item(cnt).tokVariable.varAttributes And VARIABLE_CONTROLARRAY And output_stack.Item(cnt).tokCount <> 1 Then
    MsgBox "Control array requires one index: " & output_stack.Item(cnt).tokString
    Print #99, "Control array requires one index: " & output_stack.Item(cnt).tokString
    Err.Raise 1
End If
Dim ii As InterfaceInfo
Dim mi As MemberInfo
Dim rtdt As TliVarType
Dim rtii As InterfaceInfo
#If 1 Then
Print #99, "vt="; Not output_stack.Item(cnt).tokVariable.varType Is Nothing
Print #99, "vt.ii="; Not output_stack.Item(cnt).tokVariable.varType.dtInterfaceInfo Is Nothing
Set ii = output_stack.Item(cnt).tokVariable.varType.dtInterfaceInfo
#Else
Set ii = output_stack.Item(cnt).tokInterfaceInfo
Set mi = output_stack.Item(cnt).tokMemberInfo
If Not mi Is Nothing Then
    GetReturnType ii, mi, rtdt, rtii
    Set ii = rtii
End If
#End If
If Not ii Is Nothing Then
    Print #99, "ii="; ii.Name; " am="; Hex(ii.AttributeMask); " tk="; ii.TypeKind
    ' fixme: 1 flag is returned by Text1 AttributeMask - doesn't have FDUAL
'    If Not CBool(ii.AttributeMask And (TYPEFLAG_FDUAL Or 1)) Then Exit Function ' Only get default property if Automation/Dual flags set
'        If ii.AttributeMask And TYPEFLAG_FHIDDEN Then Exit Function
'    Dim ik As InvokeKinds
'    ik = output_stack.Item(cnt).tokPCodeSubType
'    If ik And (INVOKE_FUNC Or INVOKE_PROPERTYGET) Then ik = INVOKE_FUNC Or INVOKE_PROPERTYGET
    Dim token As vbToken
    Set token = InsertDefaultMember(ii, ik, output_stack, New Collection, mi, rtdt, rtii)
    If Not token Is Nothing Then
#If 0 Then
        If ii.TypeKind = TKIND_DISPATCH Then
            ' see ElseIf below
            Print #99, "Hmmm. typekind is dispatch, has default. Should default be inserted?"
            MsgBox "Hmmm. typekind is dispatch, has default. Should default be inserted?"
        End If
#End If
        ' fixme: don't think this is always kosher
        output_stack.Item(cnt).tokPCodeSubType = INVOKE_PROPERTYGET
        If output_stack.Item(cnt).tokDataType And VT_BYREF Then OutputSubRef output_stack, cnt
'        Dim original_cnt As Long
'        original_cnt = cnt
        CoerceObject output_stack, cnt, output_stack(cnt).tokDataType, token.tokInterfaceInfo, Nothing ' force QI
        ' obsolete?
        Print #99, "CoerceDefaultMember: ts="; token.tokString; " os="; output_stack.Item(cnt).tokString; " r="; output_stack.Item(cnt).tokRank; " dt="; output_stack.Item(cnt).tokDataType
        If output_stack.Item(cnt).tokRank > 0 Then
            ' fixme: not sure what to test. vd, va, dt?
            If output_stack.Item(cnt).tokVariable Is Nothing Then GoTo pull_rank
            Print #99, "vd="; Not output_stack.Item(cnt).tokVariable.varDimensions Is Nothing; " va="; Hex(output_stack.Item(cnt).tokVariable.varAttributes)
            If output_stack.Item(cnt).tokVariable.varDimensions Is Nothing Then
pull_rank:
                Print #99, "CoerceDefaultMember: changing rank"
                MsgBox "CoerceDefaultMember: changing rank"
                token.tokRank = output_stack.Item(cnt).tokRank ' is this needed?
                output_stack.Item(cnt).tokRank = 0
            End If
        End If
        Print #99, "13"
        GoTo 30
'    ElseIf ii.TypeKind = TKIND_DISPATCH Then
        ' warning: new code!!!! Perhaps test before default is checked??? (see msgbox above)
        ' Form1.Picture = LoadPicture("") - Dispatch assign - don't want _DEFAULT - FDUAL is only flag set
        ' TreeView1.ImageList = ImageList1 - vbftpjr\ftp2.vbp
'        Print #99, "14"
    ElseIf output_stack.Item(cnt).tokMemberInfo Is Nothing Then
        ' Do this when defaults are required and member is hidden - Control = "my caption"
        ' fixme: don't think this is always kosher
'        output_stack.Item(cnt).tokType = tokInvokeDefaultMember
        ' Making sure that interface is coerced to IDispatch - iconed1a\iconwrks.vbp
'        Print #99, "Coercing to IDispatch"
'        CoerceObject output_stack, cnt, vbObject, Nothing, Nothing
'        output_stack(cnt).tokDataType = vbVariant
        Print #99, "15"
' this should probably assign correct ik to tokPCodeSubType instead of doing it in emitter. note, variables have get/let/set or'ed.
        output_stack.Item(cnt).tokPCodeSubType = 0
        GoTo 20
    Else
        ' Picture = Picture
        Print #99, "CoerceDefaultMember: interface has no default member"
'        GoTo 10
    End If
ElseIf output_stack(cnt).tokPCodeSubType <> INVOKE_FUNC Or Not IsObj(dt) Then
' Isobj for pst = FUNC which returned Object but dt was string - UDL\UDLTest.vbp
10
    Print #99, "16"
' bmpinfo.vbp, addnod1a/Project1.vbp
    'output_stack(cnt).tokPCodeSubType = INVOKE_PROPERTYGET ' replace PROPERTYPUT/REF, if any
' fixme: kludge: relookup memberinfo to switch from Put to Get/Default
' o = Nothing where o is a Property Get returning Object - causes default to be assigned, has no MemberInfo
    If Not output_stack.Item(cnt).tokMemberInfo Is Nothing And (ik And INVOKE_PROPERTYPUT) Then
'    If ik And INVOKE_PROPERTYPUT Then
        Print #99, "os.mi="; output_stack.Item(cnt).tokMemberInfo.Name; " id="; Hex(output_stack.Item(cnt).tokMemberInfo.MemberId); " ii="; Not output_stack.Item(cnt).tokInterfaceInfo Is Nothing; " ik="; ik
        ' if typekind is dispatch, then don't generate a default, even if one exists - TreeView1.ImageList = ImageList1 - vbftpjr\ftp2.vbp
        ' warning: new code!!!!
'        If output_stack.Item(cnt).tokInterfaceInfo.TypeKind = TKIND_DISPATCH Then GoTo 40
        For Each mi In output_stack.Item(cnt).tokInterfaceInfo.Members
            Print #99, "mi="; mi.Name; " id="; Hex(mi.MemberId); " ik="; mi.InvokeKind
            If mi.MemberId = output_stack.Item(cnt).tokMemberInfo.MemberId Then
                If mi.InvokeKind And (INVOKE_FUNC Or INVOKE_PROPERTYGET) Then
                    Set output_stack.Item(cnt).tokMemberInfo = mi
                    output_stack.Item(cnt).tokPCodeSubType = mi.InvokeKind
                    AddTLIInterfaceMember output_stack.Item(cnt).tokInterfaceInfo, mi, Nothing, mi.InvokeKind
                    Exit For
                End If
            End If
        Next
        If mi Is Nothing Then Err.Raise 1 ' Can't switch to get memberinfo
    End If
' o(1) = o(1) caused recursion
    If IsObj(output_stack(cnt).tokDataType) Then CoerceObject output_stack, cnt, vbObject, Nothing, Nothing ' recursive - but only once
'    output_stack.Item(cnt).tokType = tokInvokeDefaultMember
'    CoerceObject output_stack, cnt, vbObject, Nothing, Nothing
'    output_stack(cnt).tokDataType = vbVariant
20
    Set token = DispatchDefaultMember(ik)
30
    output_stack.Add token, , , cnt
    cnt = cnt + 1
    CoerceOperand gOptimizeFlag, output_stack, cnt, dt ' recursive - once for dt, loops for defaults
    CoerceDefaultMember = True
40
End If
Print #99, "CoerceDefaultMember: fnd="; CoerceDefaultMember; " s="; output_stack.Item(cnt).tokString; " t="; output_stack.Item(cnt).tokType; " pst="; output_stack.Item(cnt).tokPCodeSubType
End Function

#If 1 Then
Function DispatchDefaultMember(ByVal ik As InvokeKinds) As vbToken
Print #99, "DispatchDefaultMember: ik="; ik
Set DispatchDefaultMember = New vbToken
DispatchDefaultMember.tokString = "_DEFAULT"
DispatchDefaultMember.tokType = tokInvokeDefaultMember
DispatchDefaultMember.tokDataType = vbVariant
DispatchDefaultMember.tokPCodeSubType = ik
Set DispatchDefaultMember.tokVariable = New vbVariable
Set DispatchDefaultMember.tokVariable.varType = New vbDataType
DispatchDefaultMember.tokVariable.varType.dtType = tokIDispatchInterface
DispatchDefaultMember.tokVariable.varType.dtDataType = vbVariant
End Function
#End If

Function GetDefaultMemberInfo(ByVal ii As InterfaceInfo, ByVal ik As InvokeKinds) As MemberInfo
Print #99, "GetDefaultMemberInfo: ii="; ii.Name; " ik="; ik; " ii.am="; Hex(ii.AttributeMask); " tk="; ii.TypeKind
For Each GetDefaultMemberInfo In ii.Members
    Print #99, "n="; GetDefaultMemberInfo.Name; " mid="; Hex(GetDefaultMemberInfo.MemberId); " ik="; GetDefaultMemberInfo.InvokeKind; " mi.am="; Hex(GetDefaultMemberInfo.AttributeMask)
    If IsDefaultMember(GetDefaultMemberInfo, ik) Then GoTo 10
Next
#If 0 Then ' if used, can't do (.MemberId Mod 65536) because IDispatch default (GetTypeInfoCount=&h60010000) is picked up
For Each ii In ii.ImpliedInterfaces
    Print #99, "GetDefaultMemberInfo: ii="; ii.Name; " ik="; ik
    For Each GetDefaultMemberInfo In ii.Members
        Print #99, "n="; GetDefaultMemberInfo.Name; " mid="; Hex(GetDefaultMemberInfo.MemberId); " ik="; GetDefaultMemberInfo.InvokeKind; " mi.am="; Hex(GetDefaultMemberInfo.AttributeMask)
        If IsDefaultMember(GetDefaultMemberInfo, ik) Then GoTo 10
    Next
Next
#End If
10
Print #99, "GetDefaultMemberInfo: fnd="; Not GetDefaultMemberInfo Is Nothing
End Function

' may need a flag to allow/disallow constant coercions
Function CoerceUnaryOperand(ByVal OptimizeFlag As OptimizeFlags, ByVal output_stack As Collection, ByVal token As vbToken) As TliVarType
Print #99, "CoerceUnaryOperand: osc="; output_stack.count; " o="; Not token.tokOperator Is Nothing
If token.tokOperator Is Nothing Then Err.Raise 1
Print #99, "CoerceUnaryOperand: LHS="; token.tokLHS; " RHS="; token.tokRHS
Dim RHS As Long
RHS = token.tokRHS
If RHS = 0 Then Err.Raise 1
Print #99, "dt="; output_stack.Item(RHS).tokDataType
CoerceUnaryOperand = token.tokOperator.oprGetResultTypeUnary(output_stack.Item(RHS).tokDataType And Not VT_BYREF)
Print #99, "CUO="; CoerceUnaryOperand
If CoerceUnaryOperand = 0 Then
    If (output_stack.Item(RHS).tokDataType And Not VT_BYREF) = vbObject Then
        RHS = token.tokRHS
        If Not CoerceDefaultMember(output_stack, RHS, -1) Then Err.Raise 1
    Else
        Print #99, "Operator "; token.tokOperator.oprPCode; " cannot operate on data type of "; output_stack.Item(RHS).tokDataType
        MsgBox "Operator " & token.tokOperator.oprPCode & " cannot operate on data type of " & output_stack.Item(RHS).tokDataType
        Err.Raise 1
    End If
    CoerceUnaryOperand = token.tokOperator.oprGetResultTypeUnary(output_stack.Item(RHS).tokDataType)
    Print #99, "Operator "; token.tokOperator.oprPCode; " cannot operate on data type of "; output_stack.Item(token.tokRHS).tokDataType; " result type is " & CoerceUnaryOperand
'    MsgBox "Operator " & token.tokOperator.oprPCode & " cannot operate on data type of " & output_stack.Item(token.tokRHS).tokDataType & vbCrLf & " result type is " & CoerceUnaryOperand
    If CoerceUnaryOperand = 0 Then Err.Raise 1
End If
CoerceOperand OptimizeFlag, output_stack, RHS, CoerceUnaryOperand ' might SubRef
token.tokRHS = RHS
Print #99, "CoerceUnaryOperand: new dt="; CoerceUnaryOperand
End Function

Function CoerceBinaryOperands(ByVal OptimizeFlag As OptimizeFlags, ByVal output_stack As Collection, ByVal token As vbToken) As TliVarType
Print #99, "CoerceBinaryOperands: osc="; output_stack.count
' Some binary operators (Like) cannot use this generalized routine.
Dim LHS As Long
Dim RHS As Long
LHS = token.tokLHS
RHS = token.tokRHS
Do While True
    Print #99, "CoerceBinaryOperands: LHS="; LHS; " RHS="; RHS
    If LHS > RHS Then Err.Raise 1 ' internal error
    Print #99, "lhs.dt="; output_stack.Item(LHS).tokDataType; " rhs.dt="; output_stack.Item(RHS).tokDataType
    CoerceBinaryOperands = token.tokOperator.oprGetResultTypeBinary(output_stack.Item(LHS).tokDataType And Not VT_BYREF, output_stack.Item(RHS).tokDataType And Not VT_BYREF)
    If CoerceBinaryOperands Then Exit Do
' fixme: replace oprPCode with oprSymbol when implemented
' fixme: in Sub Main, create collection of operators, loop thru collection initializing
' Invalid syntax not caught by VB compile - "a" is nothing
    If (output_stack.Item(LHS).tokDataType And Not VT_BYREF) = vbObject Then
Print #99, "CBO: 1"
        If (output_stack.Item(RHS).tokDataType And Not VT_BYREF) = vbObject Then If Not CoerceDefaultMember(output_stack, RHS, -1) Then Err.Raise 1
Print #99, "CBO: 2"
        ' do LHS after RHS because LHS can alter RHS position in collection
        If Not CoerceDefaultMember(output_stack, LHS, -1) Then Err.Raise 1
Print #99, "CBO: 3"
        RHS = RHS + (LHS - token.tokLHS)
    ElseIf (output_stack.Item(RHS).tokDataType And Not VT_BYREF) = vbObject Then
        If Not CoerceDefaultMember(output_stack, RHS, -1) Then Err.Raise 1
    ElseIf output_stack.Item(LHS).tokDataType <> vbVariant Or output_stack.Item(RHS).tokDataType <> vbVariant Then
        ' May be non-VB data type (VT_UI4), so force to variant and try again.
        CoerceOperand OptimizeFlag, output_stack, LHS, vbVariant
        RHS = RHS + (LHS - token.tokLHS)
        CoerceOperand OptimizeFlag, output_stack, RHS, vbVariant
    Else
        Print #99, "Operator "; token.tokOperator.oprPCode; " cannot operate on data types of "; output_stack.Item(LHS).tokDataType; " "; output_stack.Item(RHS).tokDataType
        MsgBox "Operator " & token.tokOperator.oprPCode & " cannot operate on data types of " & output_stack.Item(LHS).tokDataType & " " & output_stack.Item(RHS).tokDataType
        Err.Raise 1
    End If
Print #99, "CBO: 4"
    token.tokLHS = LHS
    token.tokRHS = RHS
Loop
Print #99, "CoerceBinaryOperands: 1"
If CoerceBinaryOperands = vbBoolean Then CoerceBinaryOperands = vbOprAdd.vbOpr_oprGetResultTypeBinary(output_stack.Item(LHS).tokDataType And Not VT_BYREF, output_stack.Item(RHS).tokDataType And Not VT_BYREF)
Print #99, "CoerceBinaryOperands: 2"
CoerceBinaryOperandsLR OptimizeFlag, output_stack, token, CoerceBinaryOperands, CoerceBinaryOperands
Print #99, "CoerceBinaryOperands: 3"
Print #99, "CoerceBinaryOperands: l.dt=" & output_stack.Item(token.tokLHS).tokDataType & " r.dt=" & output_stack.Item(token.tokRHS).tokDataType
End Function

Sub CoerceBinaryOperandsLR(ByVal OptimizeFlag As OptimizeFlags, ByVal output_stack As Collection, ByVal token As vbToken, ByVal LHSType As TliVarType, ByVal RHSType As TliVarType)
Print #99, "CBOLR: osc="; output_stack.count; " LHSType="; LHSType; " RHSType="; RHSType
Dim LHS As Long
Dim RHS As Long
LHS = token.tokLHS
RHS = token.tokRHS
Print #99, "LHS="; LHS; " RHS="; RHS
If LHS > RHS Then Err.Raise 1 ' internal error
CoerceOperand OptimizeFlag, output_stack, LHS, LHSType
RHS = RHS + LHS - token.tokLHS
token.tokLHS = LHS
CoerceOperand OptimizeFlag, output_stack, RHS, RHSType
token.tokRHS = RHS
Print #99, "CBOLR: done: osc="; output_stack.count; " LHS="; LHS; " RHS="; RHS
End Sub

Function FindDefaultMemberInInterfaceInfo(ByVal ii As InterfaceInfo, ByVal ik As InvokeKinds, ByRef rtii As InterfaceInfo, ByRef rtmi As MemberInfo) As Boolean
Print #99, "FindDefaultMemberInInterfaceInfo: ii="; ii.Name; " ii.am="; Hex(ii.AttributeMask); " ik="; ik
Dim iiv As InterfaceInfo
Dim defaultCount As Integer
On Error Resume Next
Set iiv = ii.VTableInterface
On Error GoTo 0
If Not iiv Is Nothing Then Print #99, "iiv.am="; Hex(iiv.AttributeMask)
If iiv Is Nothing Then
    ' Help.vbp
    For Each rtmi In ii.Members
        Print #99, "  mi="; rtmi.Name; " id="; Hex(rtmi.MemberId); " ik="; rtmi.InvokeKind; " am="; Hex(rtmi.AttributeMask)
'        If rtmi.MemberId = 0 Then defaultCount = defaultCount + 1: If rtmi.InvokeKind = INVOKE_UNKNOWN Or rtmi.InvokeKind And ik Then Set rtii = ii: Exit For
'        If rtmi.MemberId = 0 Then defaultCount = defaultCount + 1: If rtmi.InvokeKind And ik Then Set rtii = ii: Exit For
        If rtmi.MemberId = 0 Then defaultCount = defaultCount + 1: If (rtmi.InvokeKind = INVOKE_UNKNOWN And (ik And (INVOKE_FUNC Or INVOKE_PROPERTYGET))) Or (rtmi.InvokeKind And ik) Then Set rtii = ii: Exit For
    Next
Else
    Print #99, "iiv="; iiv.Name; " guid="; iiv.GUID
    For Each rtmi In iiv.Members
        Print #99, "  mi="; rtmi.Name; " id="; Hex(rtmi.MemberId); " ik="; rtmi.InvokeKind
        If rtmi.MemberId = 0 Then defaultCount = defaultCount + 1: If rtmi.InvokeKind And ik Then Set rtii = iiv: Exit For
    Next
    If rtmi Is Nothing Then
        Print #99, "default count="; defaultCount; " am="; Hex(iiv.AttributeMask)
        If defaultCount > 0 And Not CBool(iiv.AttributeMask And TYPEFLAG_FDISPATCHABLE) Then ' VB6 is missing this test
            Print #99, "FindDefaultMemberInInterfaceInfo: defaults exist for interface "; ii.Name; " but no InvokeKind match: ik="; ik
' fixme: better??? to return collection of matching members, let calling routing check InvokeKinds
            MsgBox "Default member does not exist for interface: " & ii.Name
            Err.Raise 1
        End If
        Print #99, "iivii="; Not iiv.ImpliedInterfaces Is Nothing
        Print #99, "iiviic="; iiv.ImpliedInterfaces.count
        Dim iivii As InterfaceInfo
        For Each iivii In iiv.ImpliedInterfaces
            If FindDefaultMemberInInterfaceInfo(iivii, ik, rtii, rtmi) Then
                Print #99, "FindDefaultMemberInInterfaceInfo: implied interface: "; iivii.Name; "QI("; iiv.Name; ")"
                Exit For
            End If
        Next
    End If
End If
Print #99, "rtii="; Not rtii Is Nothing; " rtmi="; Not rtmi Is Nothing
FindDefaultMemberInInterfaceInfo = Not rtii Is Nothing And Not rtmi Is Nothing
End Function

Function InsertDefaultProjectClassMember(ByVal p As proctable) As vbToken
    Print #99, "InsertDefaultProjectClassMember: dm="; Not p Is Nothing
    Set InsertDefaultProjectClassMember = New vbToken ' is this line necessary?
    InsertDefaultProjectClassMember.tokType = tokProjectClass
    InsertDefaultProjectClassMember.tokString = p.procName
    Set InsertDefaultProjectClassMember.tokLocalFunction = p
'    InsertDefaultProjectClassMember.tokPCodeSubType = p.InvokeKind
    Print #99, "InsertDefaultProjectClassMember: frt="; Not p.procFunctionResultType Is Nothing
    Dim v As vbVariable
    If p.InvokeKind And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then Set v = p.procParams.Item(p.procParams.count).paramVariable Else Set v = p.procFunctionResultType
    If Not v Is Nothing Then
        Set InsertDefaultProjectClassMember.tokVariable = v
'        Set InsertDefaultProjectClassMember.tokVariable.varType = New vbDataType
'        Set InsertDefaultProjectClassMember.tokVariable.varType.dtClass = v.varType.dtClass
        Print #99, "InsertDefaultProjectClassMember: frt.vt="; Not v.varType Is Nothing
        Set InsertDefaultProjectClassMember.tokInterfaceInfo = v.varType.dtInterfaceInfo
        InsertDefaultProjectClassMember.tokDataType = v.varType.dtDataType
    End If
End Function

Function InsertDefaultMember(ByVal ii As InterfaceInfo, ByVal ik As InvokeKinds, ByVal output_stack As Collection, ByVal arg_stack As Collection, ByRef mi As MemberInfo, ByRef rtdt As TliVarType, ByRef rtii As InterfaceInfo) As vbToken
Print #99, "InsertDefaultMember: ii="; ii.Name; " ik="; ik; " osc="; output_stack.count; " asc="; arg_stack.count
Set mi = Nothing
rtdt = 0
Set rtii = Nothing
Dim iiv As InterfaceInfo
If FindDefaultMemberInInterfaceInfo(ii, ik, iiv, mi) Then
    Print #99, "iiv="; iiv.Name; " guid="; iiv.GUID; " mi.n="; mi.Name; " vto="; mi.VTableOffset
    ' Use actual memberinfo instead of default memberinfo, if exists
    '     its useful for debugging but not otherwise necessary
    '     although it avoids same lookup in emitter instead
'    If mi.VTableOffset = -1 Then Err.Raise 1
    If mi.VTableOffset <> -1 Then
        Dim new_mi As MemberInfo
        For Each new_mi In iiv.Members
            Print #99, "n="; new_mi.Name; " vto="; new_mi.VTableOffset; " id="; Hex(new_mi.MemberId); " ik="; new_mi.InvokeKind; " am="; Hex(new_mi.AttributeMask); " pc="; new_mi.parameters.count
            If mi.VTableOffset = new_mi.VTableOffset Then
                If new_mi.MemberId <> 0 Then
                    Print #99, "new_mi.n="; new_mi.Name
                    Set mi = new_mi
                    Exit For
                End If
            End If
        Next
    End If
'    Dim rt As VarTypeInfo
'    rt = GetReturnType(iiv, mi, rtdt, rtii)
'    If Not m Or IsObj(rtdt) Then
    Set InsertDefaultMember = New vbToken ' is this line necessary?
    Set InsertDefaultMember.tokReturnType = GetReturnType(iiv, mi, rtdt, rtii)
    Set InsertDefaultMember.tokVariable = New vbVariable
    Set InsertDefaultMember.tokVariable.varType = New vbDataType
    InsertDefaultMember.tokType = tokReferenceClass
    InsertDefaultMember.tokString = mi.Name
    Set InsertDefaultMember.tokInterfaceInfo = iiv
    Set InsertDefaultMember.tokMemberInfo = mi
    InsertDefaultMember.tokVariable.varType.dtType = tokReferenceClass
    Set InsertDefaultMember.tokVariable.varType.dtInterfaceInfo = rtii
    Print #99, "rt exists="; Not mi.ReturnType Is Nothing
    If mi.ReturnType Is Nothing Then Err.Raise 1
    InsertDefaultMember.tokDataType = rtdt
'    InsertDefaultMember.tokPCodeSubType = mi.InvokeKind ' won't handle INVOKE_UNKNOWN
    InsertDefaultMember.tokPCodeSubType = ik
    AddTLIInterfaceMember iiv, mi, InsertDefaultMember.tokReturnType, mi.InvokeKind
'        If InsertDefaultMember.tokDataType = vbObject Then Set InsertDefaultMember.tokReturnTypeInterfaceInfo = InsertDefaultMember.tokVariable.varType.dtInterfaceInfo
'    End If
End If
Print #99, "InsertDefaultMember: done: dm exists="; Not InsertDefaultMember Is Nothing; " osc="; output_stack.count; " asc="; arg_stack.count
End Function

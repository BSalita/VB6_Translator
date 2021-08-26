Attribute VB_Name = "vbtVBSubs"
Option Explicit

Public Const IndentString As String = "  "
Public indent As String
Public newindent As String

' must use emitter type priority in case (C) has different operator priority than VB
Function EmitInFix(ByVal m As vbModule, ByVal output_stack As Collection) As String
Dim token As vbToken
Dim operand_stack As New Collection
Dim i As Integer
Dim s As String
newindent = indent
For Each token In output_stack
    token.tokOutput = token.tokString
    token.tokPriority = 0
    Print #99, "EmitInFix: o=" & token.tokOutput & " tt=" & token.tokType & " osc=" & operand_stack.Count & " opc=" & output_stack.Count & " dt=" & token.tokDataType & " tc=" & token.tokCount
    Select Case token.tokType
        Case tokOperator
            If token.tokPCode = vbPCodePositive Or token.tokPCode = vbPCodeNegative Or token.tokPCode = vbPCodeNot Then
    ' May wish to remove space after unary operator but not if operator is alpha (Not)
                token.tokOutput = token.tokOutput & " " & operand_stack.Item(operand_stack.Count).tokOutput
                operand_stack.Remove operand_stack.Count
            Else
                ' does left expression need parenthesis?
                Print #99, "LHS tp="; operand_stack.Item(operand_stack.Count - 1).tokPriority; " tpc="; vbOprPriority(token.tokPCode)
                If operand_stack.Item(operand_stack.Count - 1).tokPriority <> 0 And operand_stack.Item(operand_stack.Count - 1).tokPriority < vbOprPriority(token.tokPCode) Then
                    operand_stack.Item(operand_stack.Count - 1).tokOutput = "(" & operand_stack.Item(operand_stack.Count - 1).tokOutput & ")"
                End If
                ' does right expression need parenthesis?
                Print #99, "RHS tp="; operand_stack.Item(operand_stack.Count).tokPriority; " tpc="; vbOprPriority(token.tokPCode)
                If operand_stack.Item(operand_stack.Count).tokPriority <> 0 And operand_stack.Item(operand_stack.Count).tokPriority < vbOprPriority(token.tokPCode) Then
                    operand_stack.Item(operand_stack.Count).tokOutput = "(" & operand_stack.Item(operand_stack.Count).tokOutput & ")"
                End If
                ' new expression uses operator priority
                token.tokPriority = vbOprPriority(token.tokPCode)
                token.tokOutput = operand_stack.Item(operand_stack.Count - 1).tokOutput & " " & token.tokOutput & " " & operand_stack.Item(operand_stack.Count).tokOutput
                operand_stack.Remove operand_stack.Count
                operand_stack.Remove operand_stack.Count
            End If
            operand_stack.Add token
        Case tokVariable
                ' fixme: use token.tokOutput or token.tokvariable.varsymbol (but Form1 is _Form1)?
                token.tokOutput = token.tokOutput & EmitVariable(token, operand_stack)
                operand_stack.Add token
        Case tokArrayVariable
            s = ""
            If token.tokPCode = tokReDim Then ' overloading tokPCode
                For i = (token.tokRank - 1) * 2 To 0 Step -2
                    s = s & "," & operand_stack.Item(operand_stack.Count - i - 1).tokOutput & " To " & operand_stack.Item(operand_stack.Count - i).tokOutput
                    operand_stack.Remove operand_stack.Count - i
                    operand_stack.Remove operand_stack.Count - i
                Next
                s = "(" & Mid(s, 2) & ")"
            Else
                s = EmitVariable(token, operand_stack)
            End If
            token.tokOutput = token.tokOutput & s
            operand_stack.Add token
#If 0 Then
        Case tokMemberInfo
            Dim paramCount As Integer
            If token.tokDataType = vbString Then token.tokOutput = token.tokOutput & "$"
            If UCase(token.tokOutput) = "SPC" Or UCase(token.tokOutput) = "TAB" Then ' Print statement kludge - revisit this
                token.tokOutput = token.tokOutput & "(" & operand_stack.Item(operand_stack.Count).tokOutput & ")"
                operand_stack.Remove operand_stack.Count
            Else
' fixme: can't use tokMemberInfo.Name because it may be preceded by _B_var_ or _B_str_
'                token.tokOutput = token.tokMemberInfo.Name & EmitVariable(token, operand_stack)
                token.tokOutput = token.tokOutput & EmitVariable(token, operand_stack)
#If 0 Then
' fixme: this code is duped 4+ times
                token.tokOutput = token.tokOutput & "("
                For i = 1 To token.tokCount
    Print #99, " osc=" & operand_stack.Count & " i=" & i & " mic=" & token.tokMemberInfo.Parameters.Count
'                    If i = 1 Then token.tokOutput = token.tokOutput & "("
                    token.tokOutput = token.tokOutput & operand_stack.Item(operand_stack.Count - token.tokCount + i).tokOutput
                    operand_stack.Remove operand_stack.Count - token.tokCount + i
                    If i < token.tokCount Then
                        token.tokOutput = token.tokOutput & ","
'                    Else
'                        token.tokOutput = token.tokOutput & ")"
                    End If
                Next
                token.tokOutput = token.tokOutput & ")"
#End If
            End If
            operand_stack.Add token
#End If
        Case tokVariant
            token.tokOutput = vbOutputVariant(token.tokValue)
            operand_stack.Add token
        Case tokConst
            token.tokOutput = token.tokConst.ConstName
            ' TODO: don't need module qualifier if no conflict exists. Change?
            If Not token.tokConst.ConstModule Is m Then token.tokOutput = token.tokConst.ConstModule.Name & "." & token.tokOutput
            operand_stack.Add token
        Case tokConstantInfo
            ' TODO: don't need tlib qualifier if no conflict exists. Change?
            token.tokOutput = token.tokVariable.VarType.dtConstantInfo.Parent.Name & "." & token.tokVariable.VarType.dtConstantInfo.Name & "." & token.tokMemberInfo.Name
            operand_stack.Add token
        Case tokEnumMember
            token.tokOutput = token.tokEnumMember.enumMemberName
            ' TODO: don't need module qualifier if no conflict exists. Change?
            If Not token.tokEnumMember.enumMemberParent.enumModule Is m Then token.tokOutput = token.tokEnumMember.enumMemberParent.enumModule.Name & "." & token.tokOutput
            operand_stack.Add token
        Case tokStdProcedure, tokLabelRef, tokNothing
            operand_stack.Add token
        Case tokMissing
            token.tokOutput = ""
            operand_stack.Add token
        Case tokAddressOf
            token.tokOutput = token.tokOutput & " " & operand_stack.Item(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
            operand_stack.Add token
        Case tokCvt
' note: tossing away implicit conversion token, let VB do the conversion
'         but must copy previous fields to cvt token, which ever one's are relevant.
'       Perhaps should make showing of implicit conversion a user-configurable option.
#If 0 Then
            If operand_stack.Item(operand_stack.Count).tokType = tokVariant Then
' No!! can't change tokType because it may be scanned again by next emitter
                token.tokType = tokVariant
                token.tokValue = operand_stack.Item(operand_stack.Count).tokValue
            End If
#End If
#If 0 Then ' constant expressions don't appear correctly
            token.tokOutput = operand_stack.Item(operand_stack.Count).tokOutput
            token.tokPriority = operand_stack.Item(operand_stack.Count).tokPriority
            token.tokValue = operand_stack.Item(operand_stack.Count).tokValue
            operand_stack.Remove operand_stack.Count
            operand_stack.Add token
#End If
        Case tokstatement
            EmitStatement token, operand_stack
        Case tokLabelDef
            operand_stack.Add token
            If Not IsNumeric(Left(token.tokOutput, 1)) Then token.tokOutput = token.tokOutput & ":"
        Case tokNewObject
            If Not token.tokVariable.VarType.dtClass Is Nothing Then
                token.tokOutput = token.tokVariable.VarType.dtClass.Name
            ElseIf Not token.tokVariable.VarType.dtClassInfo Is Nothing Then
                token.tokOutput = token.tokVariable.VarType.dtClassInfo.Parent.Name & "." & token.tokVariable.VarType.dtClassInfo.Name
            End If
            token.tokOutput = "New " & token.tokOutput
            operand_stack.Add token
        Case tokProjectClass
    Print #99, "1a osc="; operand_stack.Count
'            If Not token.tokLocalFunction Is Nothing Then
                token.tokOutput = token.tokLocalFunction.procName & EmitVariable(token, operand_stack)
#If 0 Then
    Print #99, "2"
    '            If token.tokDataType = vbString Then token.tokOutput = token.tokOutput & "$"
' fixme: PCodeCall expects () so always put them in. Do we want func() or func for zero args?
                token.tokOutput = token.tokOutput & "("
                For i = 1 To token.tokCount
'                    If i = 1 Then token.tokOutput = token.tokOutput & "("
                    token.tokOutput = token.tokOutput & operand_stack.Item(operand_stack.Count - token.tokCount + i).tokOutput
                    operand_stack.Remove operand_stack.Count - token.tokCount + i
                    If i < token.tokCount Then
                        token.tokOutput = token.tokOutput & ","
'                    Else
'                        token.tokOutput = token.tokOutput & ")"
                    End If
                Next
                token.tokOutput = token.tokOutput & ")"
Print #99, "to1="; token.tokOutput
Print #99, "to2="; token.tokOutput
#End If
    Print #99, "1b osc="; operand_stack.Count
                token.tokOutput = operand_stack.Item(operand_stack.Count).tokOutput & "." & token.tokOutput
    Print #99, "1c osc="; operand_stack.Count
                operand_stack.Remove operand_stack.Count
    Print #99, "1d osc="; operand_stack.Count; " o="; token.tokOutput
                operand_stack.Add token
Case tokDeclarationInfo, tokReferenceClass, tokFormClass
        Dim ss As String
        Dim mi As MemberInfo
        Dim pi As ParameterInfo
        Dim dt As TliVarType
Print #99, "4"
            If Not token.tokMemberInfo Is Nothing Then
Print #99, "5"
                Set mi = token.tokMemberInfo
                ss = ""
                i = operand_stack.Count - token.tokCount + 1
                Dim k As Integer
                k = 0
Print #99, "6 pi="; Not mi.Parameters Is Nothing
Print #99, "6 pi.c="; mi.Parameters.Count
                For Each pi In mi.Parameters
Print #99, "7"
                    k = k + 1
Print #99, "tokMember: k=" & k & " pi.n=" & pi.Name & " vt="; pi.VarTypeInfo.VarType & " mi.c=" & mi.Parameters.Count & " ik=" & Hex(mi.InvokeKind) & " tc=" & token.tokCount & " osc=" & operand_stack.Count
'                    If k = mi.Parameters.Count And (mi.InvokeKind And (INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) Then Exit For
                    If k = mi.Parameters.Count And (mi.InvokeKind And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)) Then Exit For
                    If pi.Flags And PARAMFLAG_FRETVAL Then Exit For ' Function
                    If pi.Flags And PARAMFLAG_FLCID Then ' LCID
                        s = ""
                    ElseIf k <= token.tokCount Then
                        If k = mi.Parameters.Count And mi.Parameters.OptionalCount = -1 Then
                            ' ParamArray
                            s = ""
                            Do While k <= token.tokCount
                                Print #99, "i="; i; " k="; k; " dt="; operand_stack.Item(i).tokDataType
'''' not when tokcvt is discarded If operand_stack.Item(i).tokDataType <> (VT_VARIANT Or VT_BYREF) Then Err.Raise 1  ' Internal error
                                s = s & "," & operand_stack.Item(i).tokOutput
                                operand_stack.Remove i
                                k = k + 1
                            Loop
                        Else
                            s = "," & operand_stack.Item(i).tokOutput
                            operand_stack.Remove i
                        End If
                    ElseIf pi.Optional Or CBool(pi.Flags And PARAMFLAG_FOPT) Then
                        If pi.Optional Xor CBool(pi.Flags And PARAMFLAG_FOPT) Then Print #99, "Funky optional problem"
                        s = ""
                    Else
                        Print #99, "Expecting optional parameter"
                        MsgBox "Expecting optional parameter"
                        Err.Raise 1 ' Expecting optional parameter
                    End If
                    ss = ss & s
                Next
Print #99, "8"
                ' must strip trailing commas because Call MsgBox("",) is invalid
                While Right(ss, 1) = ","
                    ss = Mid(ss, 1, Len(ss) - 1) ' remove trailing ,
                Wend
                If ss <> "" Then ss = "(" & Mid(ss, 2) & ")"
    '                If mi.InvokeKind = INVOKE_UNKNOWN Then
                    If token.tokType = tokReferenceClass Or token.tokType = tokFormClass Then
                        If token.tokVariable.VarType.dtClassInfo Is Nothing Then GoTo 90
                        If token.tokVariable.VarType.dtClassInfo.AttributeMask And TYPEFLAG_FAPPOBJECT Then GoTo 100
90
                        token.tokOutput = NameCheck(operand_stack.Item(operand_stack.Count).tokOutput) & "." & NameCheck(token.tokOutput) & ss
                        operand_stack.Remove operand_stack.Count
                    Else ' tokDeclarationInfo
100
                        token.tokOutput = NameCheck(token.tokOutput) & ss
                    End If
    '                 End If
    '                If token.tokPCodeSubType And (INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF) Then
    '                    ss = ss & "," & operand_stack.Item(operand_stack.Count).tokOutput
    '                    operand_stack.Remove operand_stack.Count
    '                End If
'                    s = s & ss
    Print #99, "9"
                Else
    ' Members not in interface.members collection but interface is extensible (Err.xxxx)
                    Print #99, "Member "; token.tokOutput; " is not in interface.members"
                    GoTo Generic_Invoke
    #If 0 Then
    ' Non-interface members such as vba.VBA__MsgBox
                    s = ii.Name & "(" & operand_stack.Item(operand_stack.Count).tokOutput ' Me
                    operand_stack.Remove operand_stack.Count
    #End If
                End If
'                token.tokOutput = s
'    Print #99, "s="; s
            operand_stack.Add token
Case tokIDispatchInterface
'            Else ' No interface - use Invoke
Generic_Invoke:
' fixme: don't use tokOutput for symbol, get symbol from elsewhere (but where?)
                token.tokOutput = token.tokOutput & EmitVariable(token, operand_stack)
#If 0 Then
    Print #99, "tokmember: invoke: c="; token.tokCount
                ss = ""
                For i = 1 To token.tokCount
                    ss = "," & operand_stack.Item(operand_stack.Count).tokOutput & ss
                    operand_stack.Remove operand_stack.Count
                Next
                If ss <> "" Then ss = "(" & Mid(ss, 2) & ")"
    '            MsgBox "ts=" & token.tokOutput & " os=" & operand_stack.Item(operand_stack.Count).tokOutput & " ss=" & ss
#End If
                token.tokOutput = operand_stack.Item(operand_stack.Count).tokOutput & "." & token.tokOutput
                operand_stack.Remove operand_stack.Count
            operand_stack.Add token
        Case tokDeclare
            Print #99, "tokDeclare: pc="; token.tokDeclare.dclParams.Count; " opc="; token.tokDeclare.dclOptionalParams
            token.tokOutput = ""
            If token.tokDeclare.dclOptionalParams = -1 Then ' ParamArray
                token.tokOutput = token.tokDeclare.dclName & EmitVariable(token, operand_stack, , token.tokCount - token.tokDeclare.dclParams.Count + 1)
            Else
                token.tokOutput = token.tokDeclare.dclName & EmitVariable(token, operand_stack)
            End If
#If 0 Then
            For i = 1 To token.tokDeclare.dclParams.Count
                If i = 1 Then token.tokOutput = token.tokOutput & "("
                token.tokOutput = token.tokOutput & operand_stack.Item(operand_stack.Count - token.tokDeclare.dclParams.Count + i).tokOutput
                operand_stack.Remove operand_stack.Count - token.tokDeclare.dclParams.Count + i
                If i < token.tokDeclare.dclParams.Count Then
                    token.tokOutput = token.tokOutput & ","
                Else
                    token.tokOutput = token.tokOutput & ")"
                End If
            Next
#End If
            operand_stack.Add token
        Case tokWithValue
'            token.tokOutput = "WithValue(" & CStr(token.tokCount) & ")"
            token.tokOutput = ""
            operand_stack.Add token
        Case tokme
            operand_stack.Add token
        Case tokCase
            token.tokOutput = operand_stack(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
            operand_stack.Add token
        Case tokCaseIs
            token.tokOutput = "Is " & PCodeToVBOpr(token.tokPCode) & " " & operand_stack(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
            operand_stack.Add token
        Case tokCaseTo
            token.tokOutput = operand_stack(operand_stack.Count - 1).tokOutput & " To " & operand_stack(operand_stack.Count).tokOutput
            operand_stack.Remove operand_stack.Count
            operand_stack.Remove operand_stack.Count
            operand_stack.Add token
        Case tokOperands ' Generic arguments for Spc, Tab, ;, ,
            s = ""
            For i = 1 To token.tokCount
                s = "," & operand_stack.Item(operand_stack.Count).tokOutput & s
                operand_stack.Remove operand_stack.Count
            Next
            token.tokOutput = Mid(s, 2)
            operand_stack.Add token
        Case tokLocalModule, tokGlobalModule ' Global module isn't implemented - needs to output qualifier
    Print #99, "tokLocalModule: name="; token.tokLocalFunction.procLocalModule.Name
    Print #99, "tokLocalModule: ct="; token.tokLocalFunction.procLocalModule.Component.Type
    Print #99, "tokLocalModule: tc="; token.tokCount; " pc="; token.tokLocalFunction.procParams.Count
            token.tokOutput = token.tokLocalFunction.procName & EmitVariable(token, operand_stack)
#If 0 Then
            Select Case token.tokLocalFunction.procLocalModule.Component.Type
                Case vbext_ct_RelatedDocument
                Case vbext_ct_StdModule
'                    If token.tokCount > 0 Then
                        s = ""
                        For i = 1 To token.tokCount
    '                        If i > 1 Then token.tokOutput = token.tokOutput & ", "
                            s = s & ", " & operand_stack.Item(operand_stack.Count - token.tokCount + i).tokOutput
                            operand_stack.Remove operand_stack.Count - token.tokCount + i
                        Next
                        token.tokOutput = token.tokOutput & "(" & Mid(s, 2) & ")"
'                    End If
                Case Else
'                    If token.tokCount > 0 Then
                        s = ""
                        For i = 1 To token.tokCount
                            s = s & ", " & operand_stack.Item(operand_stack.Count - token.tokCount + i).tokOutput
                            operand_stack.Remove operand_stack.Count - token.tokCount + i
                        Next
                        token.tokOutput = token.tokOutput & "(" & Mid(s, 2) & ")"
'                    End If
    '            Case Else
    '                Print #99, "EmitInFix: Unknown .component.type: " & token.tokLocalFunction.procLocalModule.component.type
    '                MsgBox "EmitInFix: Unknown .component.type: " & token.tokLocalFunction.procLocalModule.component.type
    '                Err.Raise 1 ' Module has unknown .component.type
            End Select
#End If
            operand_stack.Add token
        Case tokUDT
            Dim rankOffset As Long
            rankOffset = token.tokRank
            If token.tokPCode = tokReDim Then rankOffset = rankOffset * 2
            If operand_stack.Item(operand_stack.Count - rankOffset).tokDataType And Not VT_ARRAY <> vbUserDefinedType Then
                Print #99, "EmitInFix: Expecting UDT: " & operand_stack.Item(operand_stack.Count - rankOffset).tokDataType
                MsgBox "EmitInFix: Expecting UDT: " & operand_stack.Item(operand_stack.Count - rankOffset).tokDataType
                Err.Raise 1 ' Member must return Object or Variant
            End If
'            ' fixme: need to implement PointerLevel here
'            token.tokOutput = "(" & operand_stack.Item(operand_stack.Count - rankOffset).tokOutput & ")"
            Print #99, "udt: type="; operand_stack.Item(operand_stack.Count - rankOffset).tokType
            Print #99, "udt: v="; Not operand_stack.Item(operand_stack.Count - rankOffset).tokVariable Is Nothing
#If 0 Then
            Print #99, "udt: v.vt="; Not operand_stack.Item(operand_stack.Count - rankOffset).tokVariable.VarType Is Nothing
            Print #99, "udt: v.vt.type="; operand_stack.Item(operand_stack.Count - rankOffset).tokVariable.VarType.dtType
            If operand_stack.Item(operand_stack.Count - rankOffset).tokVariable.VarType.dtType = tokProjectClass Then
                Print #99, "udt="; operand_stack.Item(operand_stack.Count - rankOffset).tokVariable.VarType.dtUDT Is Nothing
                token.tokOutput = operand_stack.Item(operand_stack.Count - rankOffset).tokVariable.VarType.dtUDT.TypeName
            ElseIf operand_stack.Item(operand_stack.Count - rankOffset).tokVariable.VarType.dtType = tokReferenceClass Then
                Print #99, "ri="; operand_stack.Item(operand_stack.Count - rankOffset).tokVariable.VarType.dtRecordInfo Is Nothing
                token.tokOutput = operand_stack.Item(operand_stack.Count - rankOffset).tokVariable.VarType.dtRecordInfo.Name
            Else
                Err.Raise 1
            End If
            operand_stack.Remove operand_stack.Count - rankOffset
            token.tokOutput = token.tokOutput & "." & token.tokVariable.varSymbol & EmitVariable(token, operand_stack)
#End If
'            token.tokOutput = operand_stack.Item(operand_stack.Count - rankOffset).tokVariable.varSymbol
            ' must use operand_stack.Item(operand_stack.Count - rankOffset).tokOutput, it may contain subscript expression
            token.tokOutput = operand_stack.Item(operand_stack.Count - rankOffset).tokOutput
            operand_stack.Remove operand_stack.Count - rankOffset
            token.tokOutput = token.tokOutput & "." & token.tokVariable.varSymbol & EmitVariable(token, operand_stack)
'            token.tokOutput = token.tokOutput & "." & token.tokVariable.varSymbol
'            token.tokOutput = EmitVariable(token, operand_stack, -1)
'            ProcessSubscripts token, operand_stack
            operand_stack.Add token
        Case tokQI_Module
        Case tokQI_TLibInterface
        Case tokByVal
        Case tokSubRef
'            operand_stack.Item(operand_stack.Count).tokDataType = operand_stack.Item(operand_stack.Count).tokDataType And Not VT_BYREF
        Case tokAddRef
'            operand_stack.Item(operand_stack.Count).tokDataType = operand_stack.Item(operand_stack.Count).tokDataType Or VT_BYREF
        Case tokInvokeDefaultMember
            ' Omit default member. Could be _Default.
            s = ""
            For i = 1 To token.tokCount
                s = s & "," & operand_stack.Item(operand_stack.Count).tokOutput
                operand_stack.Remove operand_stack.Count
            Next
            token.tokOutput = operand_stack.Item(operand_stack.Count).tokOutput & "(" & Mid(s, 2) & ")"
            operand_stack.Remove operand_stack.Count
            operand_stack.Add token
        Case tokVariantArgs
            s = ""
            For i = 1 To token.tokCount
                s = "," & operand_stack.Item(operand_stack.Count).tokOutput & s
                operand_stack.Remove operand_stack.Count
            Next
            token.tokOutput = token.tokOutput & "(" & Mid(s, 2) & ")"
            operand_stack.Add token
        Case tokLBound, tokUBound
            token.tokOutput = IIf(token.tokType = tokLBound, "L", "U") & "Bound(" & operand_stack(operand_stack.Count - 1).tokOutput & "," & operand_stack(operand_stack.Count).tokOutput & ")"
            operand_stack.Remove operand_stack.Count
            operand_stack.Remove operand_stack.Count
            operand_stack.Add token
        Case Else
            Print #99, "EmitInFix: Unknown tokType: tokType=" & token.tokType
            MsgBox "EmitInFix: Unknown tokType: tokType=" & token.tokType
            Err.Raise 1
    End Select
    If operand_stack.Count > 0 Then Print #99, "EmitInFix: o="; operand_stack.Item(operand_stack.Count).tokOutput; " dt="; operand_stack.Item(operand_stack.Count).tokDataType
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

Function EmitVariable(ByVal token As vbToken, ByVal operand_stack As Collection, Optional ByVal ComponentType As vbext_ComponentType = vbext_ct_StdModule, Optional ByVal ParamArrayCnt As Integer = 0) As String
Print #99, "EmitVariable: tv="; Not token.tokVariable Is Nothing; " tc="; token.tokCount; " r="; token.tokRank; " ct="; ComponentType; " pst="; token.tokPCodeSubType; " dt="; token.tokDataType; " osc="; operand_stack.Count; " pac="; ParamArrayCnt
Dim s As String
Dim i As Long
If ParamArrayCnt > 0 Then
    For i = 1 To ParamArrayCnt
        s = "," & operand_stack.Item(operand_stack.Count).tokOutput & s
        operand_stack.Remove operand_stack.Count
    Next
End If
Print #99, "2 s="; s
For i = 1 To token.tokCount - IIf(ParamArrayCnt = -1, token.tokCount, ParamArrayCnt)
    s = "," & operand_stack.Item(operand_stack.Count).tokOutput & s
    operand_stack.Remove operand_stack.Count
Next
If s <> "" Then s = "(" & Mid(s, 2) & ")"
Print #99, "3 s="; s; " pc="; token.tokPCode; " r="; token.tokRank
If token.tokPCode = tokReDim Then ' overloading tokPCode
    Dim ss As String
    For i = (token.tokRank - 1) * 2 To 0 Step -2
        ss = ss & "," & operand_stack.Item(operand_stack.Count - i - 1).tokOutput & " To " & operand_stack.Item(operand_stack.Count - i).tokOutput
        operand_stack.Remove operand_stack.Count - i
        operand_stack.Remove operand_stack.Count - i
    Next
    s = "(" & Mid(ss, 2) & ")" & s
ElseIf token.tokRank > 0 Then
    s = ")" & s
    For i = 1 To token.tokRank
        s = "," & operand_stack.Item(operand_stack.Count).tokOutput & s
        operand_stack.Remove operand_stack.Count
    Next
    s = "(" & Mid(s, 2)
End If
EmitVariable = s
Print #99, "EmitVariable: ev="; EmitVariable
End Function

' remove token.tokoutput - use vbPcodeToSymbol()
Sub EmitStatement(ByVal token As vbToken, operand_stack As Collection)
Dim s As String
Print #99, "EmitStatement: " & token.tokOutput & " osc=" & operand_stack.Count & " pc=" & token.tokPCode & " subpc=" & token.tokPCodeSubType; " dt="; token.tokDataType
Dim t As vbToken
For Each t In operand_stack
Print #99, "EmitStatement: ts="; t.tokOutput; " dt="; t.tokDataType
Next
'Dim t As vbToken
'For Each t In operand_stack
'Print #99, "s="; t.tokOutput; " dt="; t.tokDataType
'Next
Select Case token.tokPCode
    Case vbPCodeLet, vbPCodePropertyLet
' fixme - do more of this data type checking for other pcodes
' maybe make a datatype check routine?
' fixme - may be cleaner to use token/remove instead of reassigning to output_stack
' can't compare datatypes because tokcvts are discarded
'        If (token.tokDataType And Not VT_BYREF) <> operand_stack.Item(1).tokDataType Then Err.Raise 1
'        If token.tokDataType <> operand_stack.Item(2).tokDataType Then Err.Raise 1
        token.tokOutput = operand_stack.Item(2).tokOutput & " = " & operand_stack.Item(1).tokOutput
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeLSet, vbPCodePropertySet, vbPCodeRSet, vbPCodeSet, vbPCodeSetNothing, vbPCodePropertySetNothing
' can't compare datatypes because tokcvts are discarded
'        If token.tokDataType <> operand_stack.Item(1).tokDataType Then Err.Raise 1
'        If token.tokDataType <> operand_stack.Item(2).tokDataType Then Err.Raise 1
        token.tokOutput = PCodeToSymbol(token) & " " & operand_stack.Item(2).tokOutput & " = " & operand_stack.Item(1).tokOutput
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Add token
'    Case vbPCodeSetNothing
'        If token.tokDataType <> operand_stack.Item(1).tokDataType Then Err.Raise 1
'        token.tokOutput = "Set " & operand_stack.Item(1).tokOutput & " = Nothing"
'        operand_stack.Remove 1
'        operand_stack.Add token
    Case vbPCodefor
        token.tokOutput = "For " & operand_stack.Item(1).tokOutput & " = " & operand_stack.Item(2).tokOutput & " To " & operand_stack.Item(3).tokOutput
'        If operand_stack.Item(4).tokType <> tokVariant Or operand_stack.Item(4).tokValue <> 1 Then token.tokOutput = token.tokOutput & " Step " & operand_stack.Item(4).tokOutput
' could be 1 but not tokVariant if tokCvt was processed
        If operand_stack.Item(4).tokOutput <> 1 Then token.tokOutput = token.tokOutput & " Step " & operand_stack.Item(4).tokOutput
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeforeach
        token.tokOutput = "For Each " & operand_stack.Item(1).tokOutput & " In " & operand_stack.Item(2).tokOutput
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeForNext, vbPCodeForEachNext
' obsolete Next, End Select, etc variable by using For class in vbtoken
        token.tokOutput = PCodeToSymbol(token) & " ' " & operand_stack.Item(1).tokOutput
        operand_stack.Remove 1 ' don't need For variable
        operand_stack.Add token
    Case vbPCodeForNextV, vbPCodeForEachNextV
' obsolete Next, End Select, etc variable by using For class in vbtoken
        token.tokOutput = PCodeToSymbol(token) & " " & operand_stack.Item(1).tokOutput
        operand_stack.Remove 1 ' don't need For variable
        operand_stack.Add token
    Case vbPCodeResumeLabel
        If operand_stack.Item(1).tokType = tokVariant Then
            token.tokOutput = token.tokOutput & " " & operand_stack.Item(1).tokOutput
        ElseIf operand_stack.Item(1).tokType = tokLabelRef Then
            token.tokOutput = token.tokOutput & " " & operand_stack.Item(1).tokOutput
        Else
            Print #99, "EmitStatement Unknown Resume tokType: tokType=" & operand_stack.Item(1).tokType
            MsgBox "EmitStatement Unknown Resume tokType: tokType=" & operand_stack.Item(1).tokType
            Err.Raise 1 ' internal error - expecting label
        End If
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeCloseFile
        s = ""
        Dim i As Long
        For i = 1 To token.tokCount
            s = s & ", #" & operand_stack.Item(1).tokOutput
            operand_stack.Remove 1
        Next
        token.tokOutput = PCodeToSymbol(token) & Mid(s, 2)
        operand_stack.Add token
    Case vbPCodeIf, vbPCodeSingleIf
        token.tokOutput = "If " & operand_stack.Item(1).tokOutput & " Then"
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeElseIf
        token.tokOutput = "ElseIf " & operand_stack.Item(1).tokOutput & " Then"
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeOnErrorLabel
        If operand_stack.Item(1).tokType = tokVariant Then
            token.tokOutput = "On Error GoTo " & operand_stack.Item(1).tokOutput
        ElseIf operand_stack.Item(1).tokType = tokLabelRef Then
            token.tokOutput = "On Error GoTo " & operand_stack.Item(1).tokOutput
        Else
            Print #99, "EmitStatement Unknown OnError tokType: tokType=" & operand_stack.Item(1).tokType
            MsgBox "EmitStatement Unknown OnError tokType: tokType=" & operand_stack.Item(1).tokType
            Err.Raise 1 ' internal error - expecting label
        End If
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeLineInput
        operand_stack.Item(1).tokOutput = "Line Input" & " #" & operand_stack.Item(1).tokOutput
        operand_stack.Item(2).tokOutput = operand_stack.Item(1).tokOutput & "," & operand_stack.Item(2).tokOutput
        operand_stack.Remove 1
    Case vbPCodeOpen
        s = "For"
        i = operand_stack.Item(2).tokValue And VB_OPEN_MODE_MASK
        If i = VB_OPEN_MODE_APPEND Then s = s & " Append"
        If i = VB_OPEN_MODE_BINARY Then s = s & " Binary"
        If i = VB_OPEN_MODE_INPUT Then s = s & " Input"
        If i = VB_OPEN_MODE_OUTPUT Then s = s & " Output"
        If i = VB_OPEN_MODE_RANDOM Then s = s & " Random"
        i = operand_stack.Item(2).tokValue And VB_OPEN_ACCESS_MASK
        If i = VB_OPEN_ACCESS_READ Then s = s & " Access Read"
' Hmmm, let's not output default access
'        If i = VB_OPEN_ACCESS_READ_WRITE Then s = s & " Access Read Write"
        If i = VB_OPEN_ACCESS_WRITE Then s = s & " Access Write"
        i = operand_stack.Item(2).tokValue And VB_OPEN_LOCK_MASK
        If i = VB_OPEN_LOCK_READ Then s = s & " Lock Read"
        If i = VB_OPEN_LOCK_READ_WRITE Then s = s & " Lock Read Write"
        If i = VB_OPEN_LOCK_SHARED Then s = s & " Shared"
        If i = VB_OPEN_LOCK_WRITE Then s = s & " Lock Write"
        If operand_stack.Item(4).tokOutput <> "" Then token.tokOutput = " Len = " & operand_stack.Item(4).tokOutput
        token.tokOutput = "Open " & operand_stack.Item(1).tokOutput & " " & s & " As #" & operand_stack.Item(3).tokOutput & operand_stack.Item(4).tokOutput
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeGet, vbPCodePut
        If operand_stack.Item(2).tokOutput = "0" Then
            token.tokOutput = token.tokOutput & " #" & operand_stack.Item(1).tokOutput & ", , " & operand_stack.Item(3).tokOutput
        Else
            token.tokOutput = token.tokOutput & " #" & operand_stack.Item(1).tokOutput & ", " & operand_stack.Item(2).tokOutput & ", " & operand_stack.Item(3).tokOutput
        End If
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeLock, vbPCodeUnlock
        If operand_stack.Item(2).tokOutput = "0" Then
            If operand_stack.Item(3).tokOutput = "0" Then
                token.tokOutput = token.tokOutput & " #" & operand_stack.Item(1).tokOutput
            Else
                token.tokOutput = token.tokOutput & " #" & operand_stack.Item(1).tokOutput & ", To " & operand_stack.Item(3).tokOutput
            End If
        Else
            If operand_stack.Item(3).tokOutput = "0" Then
                token.tokOutput = token.tokOutput & " #" & operand_stack.Item(1).tokOutput & ", " & operand_stack.Item(2).tokOutput
            Else
                token.tokOutput = token.tokOutput & " #" & operand_stack.Item(1).tokOutput & ", " & operand_stack.Item(2).tokOutput & " To " & operand_stack.Item(3).tokOutput
            End If
        End If
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeGoTo, vbPCodeGoSub
        If operand_stack.Item(1).tokType <> tokLabelRef Then
            Print #99, "EmitStatement Unknown tokType: tokType=" & operand_stack.Item(1).tokType
            MsgBox "EmitStatement Unknown GoTo/GoSub tokType: tokType=" & operand_stack.Item(1).tokType
            Err.Raise 1 ' internal error - expecting label
        End If
        token.tokOutput = token.tokOutput & " " & operand_stack.Item(1).tokOutput
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeOnGoTo
        token.tokOutput = "On " & operand_stack.Item(1).tokOutput & " GoTo " & operand_stack.Item(2).tokOutput
        GoTo OnPCode
    Case vbPCodeOnGoSub
        token.tokOutput = "On " & operand_stack.Item(1).tokOutput & " GoSub " & operand_stack.Item(2).tokOutput
OnPCode:
        operand_stack.Remove 1
        operand_stack.Remove 1
        Do While operand_stack.Count > 0
            If operand_stack.Item(1).tokType <> tokLabelRef Then
                Print #99, "EmitStatement Unknown OnGoSub tokType: tokType=" & operand_stack.Item(1).tokType
                MsgBox "EmitStatement Unknown tokType: tokType=" & operand_stack.Item(1).tokType
                Err.Raise 1 ' internal error - expecting label
            End If
            token.tokOutput = token.tokOutput & ", " & operand_stack.Item(1).tokOutput
            operand_stack.Remove 1
        Loop
        operand_stack.Add token
    Case vbPCodeSeek
        token.tokOutput = token.tokOutput & " #" & operand_stack.Item(1).tokOutput & "," & operand_stack.Item(2).tokOutput
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPcodeMid, vbPcodeMidB
        token.tokOutput = token.tokOutput & "(" & operand_stack.Item(1).tokOutput & ", " & operand_stack.Item(2).tokOutput & ", " & operand_stack.Item(3).tokOutput & ") = " & operand_stack.Item(4).tokOutput
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeInput
        token.tokOutput = token.tokOutput & " #" & operand_stack.Item(1).tokOutput
        operand_stack.Remove 1 ' fn
        Do ' Input must have at least one variable
            token.tokOutput = token.tokOutput & ", " & operand_stack.Item(1).tokOutput
            operand_stack.Remove 1
        Loop While operand_stack.Count > 0
        operand_stack.Add token
    Case vbPCodePrint, vbPCodewrite ' Print# or Write #
        s = token.tokOutput & " #" & operand_stack.Item(1).tokOutput & ","
        GoTo 100
    Case vbPCodeDebugPrint, vbPCodePrintMethod ' Print without #
        s = PCodeToSymbol(token)
100
        Do While operand_stack.Count > 0
Print #99, "print: c="; operand_stack.Count; " pc="; operand_stack.Item(1).tokPCode
            Select Case operand_stack.Item(1).tokPCode
                Case vbPCodePrintSpc
                    s = s & " Spc(" & operand_stack.Item(1).tokOutput & ")"
                Case vbPCodePrintTab
                    s = s & " Tab"
                    If operand_stack.Item(1).tokCount = 1 Then s = s & "(" & operand_stack.Item(1).tokOutput & ")"
                Case vbPCodePrintComma
                    s = s & ","
                Case vbPCodePrintSemiColon
                    s = s & ";"
                Case Else
                    s = s & " " & operand_stack.Item(1).tokOutput
            End Select
            operand_stack.Remove 1
        Loop
        token.tokOutput = s
        operand_stack.Add token
    Case vbPCodeCall
#If 0 Then ' fixme: remove Call
        i = InStr(1, operand_stack.Item(1).tokOutput, "(")
        If i > 0 Then
            If Right(operand_stack.Item(1).tokOutput, 1) <> ")" Then
                Print #99, "EmitInFix: Invalid parenthesis: "; operand_stack.Item(1).tokOutput
                MsgBox "EmitInFix: Invalid parenthesis: " & operand_stack.Item(1).tokOutput
                Err.Raise 1
            End If
' having trouble with double " " in call statements - fixed?
            token.tokOutput = Left(operand_stack.Item(1).tokOutput, i - 1) & " " & Mid(operand_stack.Item(1).tokOutput, i + 1, Len(operand_stack.Item(1).tokOutput) - 1 - i)
        End If
#Else
        token.tokOutput = token.tokOutput & " " & operand_stack.Item(1).tokOutput
#End If
        operand_stack.Remove 1
        operand_stack.Add token
    Case vbPCodeErase, vbPCodeCase
        token.tokOutput = token.tokOutput & " " & operand_stack.Item(1).tokOutput
        operand_stack.Remove 1
        Do While operand_stack.Count > 0
            token.tokOutput = token.tokOutput & "," & operand_stack.Item(1).tokOutput
            operand_stack.Remove 1
        Loop
        operand_stack.Add token
    Case vbPCodeReDim
        s = ""
        For i = 1 To token.tokCount
            s = s & "," & operand_stack.Item(1).tokOutput
            ' TODO: create vbTokenToVarName(token) to do below
            Print #99, "v="; Not operand_stack.Item(1).tokVariable Is Nothing; " dt="; operand_stack.Item(1).tokDataType
            If operand_stack.Item(1).tokVariable Is Nothing Then
                s = s & " As " & vbVarName(operand_stack.Item(1).tokDataType)
            Else
            ' TODO: this suggests need for varDataType???
                s = s & vbVariableType(operand_stack.Item(1).tokVariable)
            End If
            operand_stack.Remove 1
        Next
        token.tokOutput = token.tokOutput & " " & Mid(s, 2)
        operand_stack.Add token
    Case vbpcodeCircle
        Dim ss As String
        ss = operand_stack.Item(operand_stack.Count).tokOutput ' Aspect
        operand_stack.Remove operand_stack.Count
        ss = operand_stack.Item(operand_stack.Count).tokOutput & "," & ss ' End
        operand_stack.Remove operand_stack.Count
        ss = operand_stack.Item(operand_stack.Count).tokOutput & "," & ss ' Start
        operand_stack.Remove operand_stack.Count
        ss = operand_stack.Item(operand_stack.Count).tokOutput & "," & ss ' Color
        operand_stack.Remove operand_stack.Count
        ss = operand_stack.Item(operand_stack.Count).tokOutput & "," & ss ' Radius
        operand_stack.Remove operand_stack.Count
        ss = operand_stack.Item(operand_stack.Count).tokOutput & ")," & ss ' Y
        operand_stack.Remove operand_stack.Count
        ss = "(" & operand_stack.Item(operand_stack.Count).tokOutput & "," & ss ' X
        operand_stack.Remove operand_stack.Count
        ss = IIf(operand_stack.Item(operand_stack.Count).tokValue And 1, "Step", "") & ss ' Step
        operand_stack.Remove operand_stack.Count
        token.tokOutput = NameCheck(operand_stack.Item(operand_stack.Count).tokOutput) & "." & NameCheck(token.tokOutput) & " " & ss
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case vbpcodeline
        ss = operand_stack.Item(operand_stack.Count).tokOutput ' Color
        operand_stack.Remove operand_stack.Count
        ss = operand_stack.Item(operand_stack.Count).tokOutput & ")," & ss ' Y2
        operand_stack.Remove operand_stack.Count
        ss = "(" & operand_stack.Item(operand_stack.Count).tokOutput & "," & ss ' X2
        operand_stack.Remove operand_stack.Count
        s = operand_stack.Item(operand_stack.Count).tokOutput & ")-" ' Y1
        operand_stack.Remove operand_stack.Count
        s = "(" & operand_stack.Item(operand_stack.Count).tokOutput & "," & s ' X1
        operand_stack.Remove operand_stack.Count
        ss = ss & IIf(operand_stack.Item(operand_stack.Count).tokValue And 8, ",BF", "") ' Flags - F
        ss = ss & IIf(operand_stack.Item(operand_stack.Count).tokValue And 4, ",B", "") ' Flags - B
        ss = s & IIf(operand_stack.Item(operand_stack.Count).tokValue And 2, "Step", "") & ss ' Flags - 2nd step
        ss = IIf(operand_stack.Item(operand_stack.Count).tokValue And 1, "Step", "") & ss ' Flags - 1st step
        operand_stack.Remove operand_stack.Count
        token.tokOutput = NameCheck(operand_stack.Item(operand_stack.Count).tokOutput) & "." & NameCheck(token.tokOutput) & " " & ss
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case vbpcodepset
        ss = operand_stack.Item(operand_stack.Count).tokOutput ' Color
        operand_stack.Remove operand_stack.Count
        ss = operand_stack.Item(operand_stack.Count).tokOutput & ")," & ss ' Y
        operand_stack.Remove operand_stack.Count
        ss = "(" & operand_stack.Item(operand_stack.Count).tokOutput & "," & ss ' X
        operand_stack.Remove operand_stack.Count
        ss = IIf(operand_stack.Item(operand_stack.Count).tokValue And 1, "Step", "") & ss ' Step
        operand_stack.Remove operand_stack.Count
        token.tokOutput = NameCheck(operand_stack.Item(operand_stack.Count).tokOutput) & "." & NameCheck(token.tokOutput) & " " & ss
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case vbpcodescale
        ss = operand_stack.Item(operand_stack.Count).tokOutput & ")" & ss ' Y2
        operand_stack.Remove operand_stack.Count
        ss = "(" & operand_stack.Item(operand_stack.Count).tokOutput & "," & ss ' X2
        operand_stack.Remove operand_stack.Count
        ss = operand_stack.Item(operand_stack.Count).tokOutput & ")-" & ss ' Y1
        operand_stack.Remove operand_stack.Count
        ss = "(" & operand_stack.Item(operand_stack.Count).tokOutput & "," & ss ' X1
        operand_stack.Remove operand_stack.Count
        operand_stack.Remove operand_stack.Count ' Flags - do nothing
        token.tokOutput = NameCheck(operand_stack.Item(operand_stack.Count).tokOutput) & "." & NameCheck(token.tokOutput) & " " & ss
        operand_stack.Remove operand_stack.Count
        operand_stack.Add token
    Case Else
        GenericOperands token, operand_stack
End Select
PerformIndentChanges token.tokPCode
Print #99, "EmitStatement: {"; operand_stack.Item(1).tokOutput; "}"
If operand_stack.Count <> 1 Then
    Print #99, "EmitStatement: operand count <> 1: count=" & operand_stack.Count
    MsgBox "EmitStatement: operand count <> 1: count=" & operand_stack.Count
    Err.Raise 1 ' compiler error
End If
End Sub

Function vbOutputVariant(ByVal v As Variant) As String
Print #99, "vbOutputVariant: vt="; VarType(v)
Select Case VarType(v)
    Case vbDate
        vbOutputVariant = "#" & v & "#"
    Case vbString
        vbOutputVariant = """" & v & """"
    Case vbNull, vbEmpty, vbObject
        vbOutputVariant = TypeName(v)
    Case Else
        vbOutputVariant = v
End Select
Print #99, "vbOutputVariant: s="; vbOutputVariant
End Function

Function PCodeToVBOpr(ByVal pcode As vbPCodes) As String
Select Case pcode
    Case vbPCodeEQ
        PCodeToVBOpr = "="
    Case vbPCodeLT
        PCodeToVBOpr = "<"
    Case vbPCodeLE
        PCodeToVBOpr = "<="
    Case vbPCodeNE
        PCodeToVBOpr = "<>"
    Case vbPCodeGT
        PCodeToVBOpr = ">"
    Case vbPCodeGE
        PCodeToVBOpr = ">="
    Case Else
        Print #99, "PCodeToVBOpr: Unexpected PCode: PCode=" & pcode
        MsgBox "PCodeToVBOpr: Unexpected PCode: PCode=" & pcode
        Err.Raise 1 ' Internal error - Unknown PCode
End Select
End Function

Function vbOprPriority(ByVal pcode As vbPCodes) As Integer
Select Case pcode
    Case vbPCodeImp
        vbOprPriority = vbOprPriorityImp
    Case vbPCodeEqv
        vbOprPriority = vbOprPriorityEqv
    Case vbPCodeXor
        vbOprPriority = vbOprPriorityXor
    Case vbPCodeOr
        vbOprPriority = vbOprPriorityOr
    Case vbPCodeAnd
        vbOprPriority = vbOprPriorityAnd
    Case vbPCodeNot
        vbOprPriority = vbOprPriorityNot
    Case vbPCodeEQ
        vbOprPriority = vbOprPriorityCmp
    Case vbPCodeLT
        vbOprPriority = vbOprPriorityCmp
    Case vbPCodeLE
        vbOprPriority = vbOprPriorityCmp
    Case vbPCodeNE
        vbOprPriority = vbOprPriorityCmp
    Case vbPCodeGT
        vbOprPriority = vbOprPriorityCmp
    Case vbPCodeGE
        vbOprPriority = vbOprPriorityCmp
    Case vbPCodeIs
        vbOprPriority = vbOprPriorityCmp
    Case vbPCodeLike
        vbOprPriority = vbOprPriorityCmp
    Case vbPCodeCat
        vbOprPriority = vbOprPriorityCat
    Case vbPCodeAdd
        vbOprPriority = vbOprPriorityAddSub
    Case vbPCodeSub
        vbOprPriority = vbOprPriorityAddSub
    Case vbPCodeMod
        vbOprPriority = vbOprPriorityMod
    Case vbPCodeIDiv
        vbOprPriority = vbOprPriorityIDiv
    Case vbPCodeMul
        vbOprPriority = vbOprPriorityMulDiv
    Case vbPCodeDiv
        vbOprPriority = vbOprPriorityMulDiv
    Case vbPCodePositive
        vbOprPriority = vbOprPriorityPositiveNegative
    Case vbPCodeNegative
        vbOprPriority = vbOprPriorityPositiveNegative
    Case vbPCodePow
        vbOprPriority = vbOprPriorityPow
    Case Else
        Print #99, "vbOprPriority: Unknown PCode: PCode=" & pcode
        MsgBox "vbOprPriority: Unknown PCode: PCode=" & pcode
        Err.Raise 1 ' Internal error - Unknown PCode
End Select
End Function

Function PCodeToSymbol(ByVal token As vbToken) As String
Select Case token.tokPCode
    Case vbPCodeCloseFile
        PCodeToSymbol = "Close"
    Case vbPCodeCvt
        Err.Raise 1 ' internal error
    Case vbPCodeDebugAssert
        PCodeToSymbol = "Debug.Assert"
    Case vbPCodeDebugPrint
        PCodeToSymbol = "Debug.Print"
    Case vbPCodeCaseElse
        PCodeToSymbol = "Case Else"
    Case vbPCodeDoUntil
        PCodeToSymbol = "Do Until"
    Case vbPCodeDoWhile
        PCodeToSymbol = "Do While"
    Case vbPCodeEndIf
        PCodeToSymbol = "End If"
    Case vbPCodeEndSelect
        PCodeToSymbol = "End Select"
    Case vbPCodeEndWith
        PCodeToSymbol = "End With"
    Case vbPCodeExitDo
        PCodeToSymbol = "Exit Do"
    Case vbPCodeExitFor
        PCodeToSymbol = "Exit For"
    Case vbPCodeExitFunction
        PCodeToSymbol = "Exit Function"
    Case vbPCodeExitSub
        PCodeToSymbol = "Exit Sub"
    Case vbPCodeforeach
        PCodeToSymbol = "For Each"
    Case vbPCodeLineInput
        PCodeToSymbol = "Line Input"
    Case vbPCodeLoopInfinite
        PCodeToSymbol = "Loop ' Infinite"
    Case vbPCodeLoopUntil
        PCodeToSymbol = "Loop Until"
    Case vbPCodeLoopWhile
        PCodeToSymbol = "Loop While"
    Case vbPCodeOnError0
        PCodeToSymbol = "On Error GoTo 0"
    Case vbPCodeOnErrorLabel
        PCodeToSymbol = "On Error GoTo"
    Case vbpcodeonerrorresumenext
        PCodeToSymbol = "On Error Resume Next"
    Case vbPCodeResume0
        PCodeToSymbol = "Resume 0"
    Case vbPCodeResumeNext
        PCodeToSymbol = "Resume Next"
    Case vbPCodeReturn
        PCodeToSymbol = "Return"
    Case vbPCodeSelect
        PCodeToSymbol = "Select Case"
    Case vbPCodeSet, vbPCodePropertySet, vbPCodeSetNothing, vbPCodePropertySetNothing
        PCodeToSymbol = "Set"
    Case Else
        PCodeToSymbol = token.tokOutput
End Select
End Function

Sub GenericOperands(ByVal token As vbToken, ByVal operand_stack As Collection)
Dim s As String
s = PCodeToSymbol(token)
Dim t As vbToken
For Each t In operand_stack
    s = s & " " & t.tokOutput
    operand_stack.Remove 1
Next
token.tokOutput = s
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
    Case vbPCodeEndIf, vbPCodeEndWith, vbPCodeSingleIfEndIf, vbPCodeForNext, vbPCodeForNextV, vbPCodeForEachNext, vbPCodeForEachNextV, vbPCodeLoop, vbPCodeLoopInfinite, vbPCodeLoopUntil, vbPCodeWend, vbPCodeLoopWhile
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

Function vbTypeName(ByVal vt As vbVariable) As String
Print #99, "vbTypeName: sym="; vt.varSymbol; " varType="; vt.VarType Is Nothing
If Not vt.varDimensions Is Nothing Then
    Dim a As vbVarDimension
    For Each a In vt.varDimensions
        If vbTypeName <> "" Then vbTypeName = vbTypeName & ","
        If a.varDimensionLBound <> 0 Then vbTypeName = vbTypeName & a.varDimensionLBound & " To "
        vbTypeName = vbTypeName & a.varDimensionUBound
    Next
    vbTypeName = "(" & vbTypeName & ")"
End If
vbTypeName = vt.varSymbol & vbTypeName & vbVariableType(vt)
Print #99, "vbTypeName="; vbTypeName
End Function

Function vbVariableType(ByVal vt As vbVariable) As String
#If 0 Then
vbVariableType = vbVariableType & TypeName(vt.varAttributes And Not vbByRef) ' check out why typeName(2) doesn't return integer
#Else
Print #99, "vbVariableType: dt="; vt.VarType.dtDataType; " vd="; Not vt.varDimensions Is Nothing
vbVariableType = " As " & IIf(vt.varAttributes And VARIABLE_NEW, "New ", "") & vbDataType(vt.VarType)
Print #99, "vbVariableType: "; vbVariableType
#End If
End Function

Function vbDataType(ByVal dt As vbDataType) As String
Select Case dt.dtDataType
Case vbObject
    If Not dt.dtClass Is Nothing Then
        vbDataType = vbDataType & dt.dtClass.Name
    ElseIf Not dt.dtClassInfo Is Nothing Then
        vbDataType = vbDataType & dt.dtClassInfo.Name
    Else
        vbDataType = vbDataType & "Object"
    End If
Case VT_UNKNOWN
    If Not dt.dtInterfaceInfo Is Nothing Then
        vbDataType = vbDataType & dt.dtInterfaceInfo.Name
    Else
        Err.Raise 1 ' Internal error - Don't know how to handle IUnknown
    End If
Case vbUserDefinedType
    If Not dt.dtUDT Is Nothing Then
        vbDataType = vbDataType & dt.dtUDT.TypeName
    ElseIf Not dt.dtRecordInfo Is Nothing Then
        ' TODO: Only need Parent qualifier if name conflict. Remove?
        vbDataType = vbDataType & dt.dtRecordInfo.Parent.Name & "." & dt.dtRecordInfo.Name
    Else
        Err.Raise 1 ' Internal error
    End If
Case VT_LPWSTR
    vbDataType = vbDataType & "String * " & dt.dtLength
Case Else
    vbDataType = vbVarName(dt.dtDataType)
End Select
Print #99, "vbDataType: "; vbDataType
End Function

Function vbVarName(vt As TliVarType) As String
Print #99, "vbVarName: vt="; vt
Select Case vt
Case vbNull
    vbVarName = "Null"
Case vbInteger
    vbVarName = "Integer" ' int16_t
Case vbBoolean
    vbVarName = "Boolean" ' int16_t
Case vbLong
    vbVarName = "Long" ' int32_t
Case vbSingle
    vbVarName = "Single"
Case vbDouble
    vbVarName = "Double"
Case vbDate
    vbVarName = "Date"
Case vbCurrency
    vbVarName = "Currency"
Case vbString
    vbVarName = "String"
Case vbVariant
    vbVarName = "Variant"
Case vbByte
    vbVarName = "Byte"
Case VT_VOID ' Using VT_VOID to signify Any
    vbVarName = "Any"
Case Else
    Print #99, "vbVarName: Unknown vt: vbvarname=" & vt
    MsgBox "vbVarName: Unknown vt: vbvarname=" & vt
    Err.Raise 1 ' Internal error - Unknown VarType
End Select
Print #99, "vbVarName: "; vbVarName
End Function

' Rename this function - scope is a misnomer
' fixme: can't use the name "default" because its a C namespace conflict
Function vbScopeAttributes(ByVal pa As procattributes, ByVal deflt As String) As String
If pa And PROC_ATTR_Friend Then vbScopeAttributes = "Friend "
If pa And PROC_ATTR_PRIVATE Then vbScopeAttributes = "Private "
If pa And PROC_ATTR_PUBLIC Then vbScopeAttributes = "Public "
If deflt <> "" And CBool(pa And PROC_ATTR_DEFAULT) Then vbScopeAttributes = deflt & " "
If pa And PROC_ATTR_Static Then vbScopeAttributes = vbScopeAttributes & "Static "
If pa And VARIABLE_WITHEVENTS Then vbScopeAttributes = vbScopeAttributes & "WithEvents "
End Function

Function NameCheck(ByVal v As String) As String
Dim bracket As Boolean
NameCheck = v
Dim i As Long
For i = 1 To Len(v)
    If Not isalnum(Mid(v, i, 1)) Then If Mid(v, i, 1) = "(" Or Mid(v, i, 1) = "." Then Exit For Else If i = 1 Or Mid(v, i, 1) <> "_" Then NameCheck = "[" & NameCheck & "]": Exit For
Next
End Function

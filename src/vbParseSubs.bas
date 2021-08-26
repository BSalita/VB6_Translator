Attribute VB_Name = "vbtParseSubs"
Option Explicit

Function parseDimension(ByVal tokens As Collection, v As vbVariable)
Dim output_stack As New Collection
If Not currentModule Is Nothing Then Print #99, "m="; currentModule.Name
If Not currentProc Is Nothing Then Print #99, "p="; currentProc.procName
If tokens.count > 0 Then
Print #99, "parseDimension: 6 s="; v.varSymbol
    If tokens.Item(1).tokString = "(" Then
        Set v.varDimensions = New Collection
        If tokens.Item(2).tokString = ")" Then
            tokens.Remove 1
        Else
            Do
                tokens.Remove 1
                oRPN.ConstantRPNize OptimizeConstantExpressions, tokens, output_stack, vbLong
                v.varDimensions.Add New vbVarDimension
                Dim vd As vbVarDimension
                Set vd = v.varDimensions.Item(v.varDimensions.count)
                If tokens.Item(1).tokKeyword = KW_TO Then
                    tokens.Remove 1
                    oRPN.ConstantRPNize OptimizeConstantExpressions, tokens, output_stack, vbLong
                    vd.varDimensionLBound = output_stack.Item(1).tokValue ' Expecting a constant expression
                    output_stack.Remove 1
                End If
                vd.varDimensionUBound = output_stack.Item(1).tokValue
                output_stack.Remove 1
Print #99, "parseDimension: lbound="; vd.varDimensionLBound; " ubound="; vd.varDimensionUBound
                If vd.varDimensionUBound > vd.varDimensionUBound Then Err.Raise 1 ' lbound must be lower than ubound
            Loop While tokens.Item(1).tokString = ","
            If tokens.Item(1).tokString <> ")" Then Err.Raise 1
        End If
        tokens.Remove 1
    End If
End If
End Function

' need to pass in attr and assign to variable
Sub parseDim(ByVal tokens As Collection, ByVal pa As procattributes)
Print #99, "parseDim: 1 pass="; PassNumber; " s="; tokens.Item(2).tokString; " pa="; pa; " cm="; Not currentModule Is Nothing
' using PROC_ATTR_PRIVATE test because Private items only get called on first pass (function of .Components)
'If PassNumber < 2 And (pa And PROC_ATTR_PRIVATE) = 0 Then RemoveAll tokens: Exit Sub
If PassNumber < 2 Then RemoveAll tokens: Exit Sub
Dim token As vbToken
Dim attr As VarFlags
Dim variables As Collection
' fixme: can some of this be eliminated using currentvariable?
If currentProc Is Nothing Then
' not sure how Private class variables should be handled - non-ole ok?
Print #99, "ct="; currentModule.Component.Type
'    If currentModule.Component.Type <> vbext_ct_StdModule And (pa And (PROC_ATTR_PRIVATE Or PROC_ATTR_PUBLIC)) Then
    If currentModule.Component.Type <> vbext_ct_StdModule And (pa And PROC_ATTR_PUBLIC) Then
        ' Variable is actually a property get/put method
        ' kludge: using INVOKE_UNKNOWN for public class variables
        ParseFunction tokens, pa Or PROC_ATTR_VARIABLE, vbext_mt_Variable, (INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)
        Set currentProc = Nothing
        Exit Sub
    Else
        Set variables = currentModule.ModuleVars
    End If
Else
    Set variables = currentProc.procLocalVariables
End If
Do
Print #99, "parseDim: 2"
    tokens.Remove 1 ' remove Dim, Public, Static, or Private
    If tokens.count = 0 Then Err.Raise 1 ' Missing variable
    If UCase(tokens.Item(1).tokString) = "WITHEVENTS" Then
        tokens.Remove 1 ' remove WithEvents
        attr = pa Or VARIABLE_WITHEVENTS
        If tokens.count = 0 Then Err.Raise 1 ' Missing variable
' object must follow?
    Else
        attr = pa
    End If
Print #99, "parseDim: 3"
    Set token = tokens.Item(1)
    tokens.Remove 1
Print #99, "parseDim: 4 s="; token.tokString; " vc="; variables.count

' following should be in name check function
    On Error Resume Next
    Dim Stmt As vbStmt
    Set Stmt = Nothing
    Set Stmt = cStatements(token.tokString)
    On Error GoTo 0
    If Not Stmt Is Nothing Then
        Print #99, "Name conflict: "; token.tokString
        MsgBox "Name conflict: " & token.tokString
        Err.Raise 1 ' Name conflict
    End If
Print #99, "parseDim: 5"
Print #99, "parseDim: 6"
    Dim v As vbVariable
    Set v = New vbVariable
    v.varLineNumber = token.tokLineNumber
    Set v.varComponent = token.tokComponent
    v.MemberType = vbext_mt_Variable
    v.varSymbol = token.tokString
    v.varAttributes = attr
    Set v.varModule = currentModule
    Set v.varProc = currentProc
Print #99, "parseDim: 7"
    parseDimension tokens, v
Print #99, "parseDim: 8 vd="; Not v.varDimensions Is Nothing
If Not v.varDimensions Is Nothing Then Print #99, "parseDim: 9 vd.c="; v.varDimensions.count
    getAsNewDataType token, tokens, v
    ' fixme: this code is duped several times
    If tokens.count > 0 Then
        If tokens.Item(1).tokString = "(" Then
            tokens.Remove 1
            If tokens.count = 0 Then Err.Raise 1 ' Missing )
            If tokens.Item(1).tokString <> ")" Then Err.Raise 1 ' Expecting )
            tokens.Remove 1
            Set v.varDimensions = New Collection
    '        v.varType.dtDataType = v.varType.dtDataType Or VT_ARRAY
        End If
    End If
Print #99, "parsedim: 10 tokens.count=" & tokens.count & " variables.count="; variables.count; " variables.dt="; v.varType.dtDataType
    If CBool(v.varAttributes And VARIABLE_WITHEVENTS) And CBool(v.varAttributes And VARIABLE_NEW) Then Err.Raise 1 ' WithEvents can't be used with New
    v.varAttributes = v.varAttributes Or attr
    ' CheckVarDup returns True if Dim statement has been rescaned.
    If Not CheckVarDup(token, variables) Then If variables.count = 0 Then variables.Add v, UCase(token.tokString) Else variables.Add v, UCase(token.tokString), 1
    If tokens.count = 0 Then Exit Do
Loop While tokens.Item(1).tokString = ","
Print #99, "parsedim: end tc="; tokens.count
End Sub

Sub parseReDim(ByVal tokens As Collection)
Dim output_stack As New Collection
Dim token As vbToken
Dim ReDimToken As vbToken
Dim v As vbVariable
Dim existing_variable As Boolean

Print #99, "parseReDim: 1"
Set ReDimToken = tokens.Item(1)
Print #99, "parseReDim: 2"
If getKeyword(tokens.Item(2)) = KW_PRESERVE Then
Print #99, "parseReDim: 3"
    tokens.Remove 2
Print #99, "parseReDim: 4"
    ReDimToken.tokString = "ReDimPreserve"
End If
Print #99, "parseReDim: 5"
Do
Print #99, "parseReDim: 6"
    tokens.Remove 1 ' remove ReDim or ,
Print #99, "parseReDim: 7 ts="; tokens.Item(1).tokString
    Set token = tokens.Item(1)
Print #99, "parseReDim: 8"
    tokens.Remove 1 ' remove variable name
Print #99, "parseReDim: 9"
    ' fixme: make this variable-only search a Func
    Dim varname As String
    varname = UCase(token.tokString)
    Set v = Nothing
    On Error Resume Next
    Set v = currentProc.procLocalVariables.Item(varname)
Print #99, "parseReDim: 9a"; Not v Is Nothing
    If v Is Nothing Then Set v = currentProc.procParams.Item(varname).paramVariable
Print #99, "parseReDim: 9b"; Not v Is Nothing
    If v Is Nothing Then Set v = currentModule.ModuleVars.Item(varname)
Print #99, "parseReDim: 9c"; Not v Is Nothing
    On Error GoTo 0
    If v Is Nothing Then
        ' fixme: make this a Func
        Dim m As vbModule
        For Each m In currentProject.prjModules
            If Not m Is currentModule Then
                If m.Component.Type = vbext_ct_StdModule Then
                    On Error Resume Next
                    Set v = m.ModuleVars.Item(varname)
                    On Error GoTo 0
                    ' fixme: check for missing Exit For in other loops
                    If Not v Is Nothing Then Exit For
                End If
            End If
        Next
    End If
    If v Is Nothing Then
Print #99, "parseReDim: 10"
        If currentProc.procLocalVariables.count = 0 Then currentProc.procLocalVariables.Add New vbVariable, UCase(token.tokString) Else currentProc.procLocalVariables.Add New vbVariable, UCase(token.tokString), 1
        Set v = currentProc.procLocalVariables.Item(1)
        v.MemberType = vbext_mt_Variable
        v.varSymbol = token.tokString
        Set v.varModule = currentModule
        Set v.varProc = currentProc
        existing_variable = False
Print #99, "parseReDim: 11"
    Else
Print #99, "parseReDim: 12 v="; v.varSymbol
        existing_variable = True
    End If
Print #99, "parseReDim: 14"
    token.tokType = tokVariable
    Set token.tokVariable = v
    ReDimToken.tokCount = ReDimToken.tokCount + 1 ' Number of Variables
    Dim rank As Integer
    rank = 0
' fixme: obsolete RankToken? Use tokRank in each variable token?
    Dim RankToken As vbToken
    Set RankToken = New vbToken
    output_stack.Add RankToken
    output_stack.Add token
Print #99, "parseReDim: 15"
    If tokens.count > 0 Then
Print #99, "parseReDim: 16"
        If tokens.Item(1).tokString = "(" Then
Print #99, "parseReDim: 17 exist="; existing_variable
            If v.varDimensions Is Nothing Then
                If existing_variable Then Err.Raise 1 ' Previously declared as non-array
                Set v.varDimensions = New Collection
            Else
                Print #99, "parseReDim: dc="; v.varDimensions.count
                If v.varDimensions.count <> 0 Then Err.Raise 1 ' Previously dimensioned array cannot be redimensioned
            End If
            token.tokType = tokArrayVariable
            Do
                tokens.Remove 1
                Dim count As Integer
                count = output_stack.count
Print #99, "parseReDim: 18 c="; count
                oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbLong
' fixme: should be using tokens.item(1).tokKeyword = KW_... everywhere!!!
                If tokens.Item(1).tokKeyword = KW_TO Then
                    tokens.Remove 1
                    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbLong
                Else
                    output_stack.Add New vbToken, , , count ' insert before TO expr
                    output_stack.Item(count + 1).tokString = "0"
                    output_stack.Item(count + 1).tokType = tokvariant
                    output_stack.Item(count + 1).tokValue = 0&
                    output_stack.Item(count + 1).tokDataType = vbLong
                End If
Print #99, "parseReDim: 19"
                rank = rank + 1
            Loop While tokens.Item(1).tokString = ","
Print #99, "parseReDim: 20"
            If tokens.Item(1).tokString <> ")" Then Err.Raise 1
Print #99, "parseReDim: 21 c="; tokens.count
            tokens.Remove 1
Print #99, "parseReDim: 22 c="; tokens.count
        Else
Print #99, "parseReDim: 23"
            If Not v.varDimensions Is Nothing Then Err.Raise 1 ' Previous declared array is missing array subscript
        End If
    Else
Print #99, "parseReDim: 24"
        If Not v.varDimensions Is Nothing Then Err.Raise 1 ' Previous declared array is missing array subscript
    End If
    getAsDataType token, tokens, v
    ' Perhaps more data type checking should be done
    RankToken.tokString = rank
    RankToken.tokType = tokvariant
    RankToken.tokValue = rank ' remove need for rank token?
    RankToken.tokDataType = vbLong
For Each token In output_stack
Print #99, "o="; token.tokString; " t="; token.tokType; " v="; token.tokValue
Next
    If tokens.count = 0 Then Exit Do
Loop While tokens.Item(1).tokString = ","
Print #99, "parseReDim: 30"
output_stack.Add ReDimToken
currentProc.procStatements.Add output_stack
End Sub

' need to stop using ParseSub for property parsing, leave off membertype param
Function ParseSub(ByVal tokens As Collection, ByVal pa As procattributes, Optional ByVal MemberType As vbext_MemberType = vbext_mt_Method, Optional ByVal InvokeKind As InvokeKinds = INVOKE_FUNC) As proctable
Dim ProcNameIK As String
Dim st As SpecialTypes
Dim pt As proctable
Dim token As vbToken
Set token = tokens.Item(1)
tokens.Remove 1 ' remove Sub, Function, Property, Get, Let, Set
ProcNameIK = UCase(tokens.Item(1).tokString) & IIf(InvokeKind = INVOKE_FUNC Or InvokeKind = INVOKE_PROPERTYGET, INVOKE_FUNC Or INVOKE_PROPERTYGET, InvokeKind)
Print #99, "ParseSub: pnik="; ProcNameIK
On Error Resume Next
Set pt = currentModule.procs.Item(ProcNameIK)
On Error GoTo 0
If pt Is Nothing Then
    Set pt = ProcAdd(tokens.Item(1).tokString, pa Or proc_attr_defined, MemberType, InvokeKind)
Else
    If pt.procattributes And proc_attr_defined Then
        Print #99, "Duplicate definition: "; pt.procName
        MsgBox "Duplicate definition of """ & pt.procName & """. Event and procedure name same?"
        Err.Raise 1 ' duplicate definition
    End If
End If
pt.procNumber = currentModule.procs.count
Set currentProc = pt
tokens.Remove 1
If tokens.count > 0 Then
    st = getSpecialTypes(tokens.Item(1))
    If st = SPECIAL_OP Then
        tokens.Remove 1
        st = getSpecialTypes(tokens.Item(1))
        If st = SPECIAL_CP Then
            tokens.Remove 1
        Else
            Do
                Dim p As paramTable
Print #99, "parsesub: 1"
                Set p = getOptionalByAsDataType(tokens)
Print #99, "parsesub: 2"
                pt.procParams.Add p, UCase(p.paramVariable.varSymbol)
Print #99, "parsesub: 3"
                p.paramVariable.varAttributes = p.paramVariable.varAttributes Or VARIABLE_PARAMETER
Print #99, "parsesub: 4 ts="; tokens.Item(1).tokString
                st = getSpecialTypes(tokens.Item(1))
Print #99, "parsesub: 5 st="; st
                tokens.Remove 1
Print #99, "parsesub: 13"
            Loop While st = special_comma
Print #99, "parsesub: 14 st="; st
            If st <> SPECIAL_CP Then Err.Raise 1
        End If
    End If
End If
Print #99, "parsesub: 15"
For Each p In pt.procParams
    If p.paramVariable.varAttributes And VARIABLE_OPTIONAL Then
        If p.paramVariable.varAttributes And VARIABLE_PARAMARRAY Then
            If pt.procOptionalParams = -1 Then Err.Raise 1 ' Only last parameter can be ParamArray
            If pt.procOptionalParams > 0 Then Err.Raise 1 ' Optional parameters not allowed with ParamArray
            pt.procOptionalParams = -1
        Else
            pt.procOptionalParams = pt.procOptionalParams + 1
        End If
    Else
        If pt.procOptionalParams > 0 Then Err.Raise 1 ' non-optional parameter not allowed after optional parameter
    End If
Next
Set ParseSub = pt
Print #99, "parsesub: 16"
End Function

Sub ParseFunction(ByVal tokens As Collection, ByVal pa As procattributes, Optional ByVal MemberType As vbext_MemberType = vbext_mt_Method, Optional ByVal InvokeKind As InvokeKinds = INVOKE_FUNC)
Print #99, "ParseFunction: 1"
Dim token As vbToken
Dim pt As proctable
Set token = tokens.Item(2) ' Assumed to be Function name token
Print #99, "ParseFunction: 2 ts="; token.tokString
Set pt = ParseSub(tokens, pa Or PROC_ATTR_FUNCTION, MemberType, InvokeKind)
Print #99, "ParseFunction: 4"
Set pt.procFunctionResultType = New vbVariable
pt.procFunctionResultType.MemberType = vbext_mt_Variable
pt.procFunctionResultType.varSymbol = pt.procName
pt.procFunctionResultType.varAttributes = VARIABLE_FUNCTION
If pa And PROC_ATTR_VARIABLE Then
    getAsNewDataType token, tokens, pt.procFunctionResultType
Else
    getAsDataType token, tokens, pt.procFunctionResultType
End If
' fixme: this code is duped several times
If tokens.count > 0 Then
    Print #99, "ParseFunction: 2a ts="; tokens.Item(1).tokString
    If tokens.Item(1).tokString = "(" Then
        Print #99, "ParseFunction: 2b"
        tokens.Remove 1
        If tokens.count = 0 Then Err.Raise 1 ' Missing )
        If tokens.Item(1).tokString <> ")" Then Err.Raise 1 ' Expecting )
        tokens.Remove 1
        Set pt.procFunctionResultType.varDimensions = New Collection
'        dt.dtDataType = dt.dtDataType Or VT_ARRAY
    End If
    Print #99, "ParseFunction: 2c"
End If
Print #99, "ParseFunction: 5 vt="; Not pt.procFunctionResultType.varType Is Nothing
Print #99, "ParseFunction: 6 type="; pt.procFunctionResultType.varType.dtType
'pt.procFunctionResultType.varType.dtType = tokProjectClass
Set pt.procFunctionResultType.varModule = currentModule
Set pt.procFunctionResultType.varProc = currentProc
'Set pt.procFunctionResultType.varType.dtClass = currentModule
Print #99, "ParseFunction: func="; pt.procFunctionResultType.varSymbol
End Sub

Sub ParseProperty(ByVal tokens As Collection, ByVal pa As procattributes)
Dim token As vbToken
Dim pt As proctable
tokens.Remove 1 ' bypass Property
Set token = tokens.Item(1) ' Get, Let, Set
Select Case UCase(token.tokString)
' Make Let/Set a Function -- use Variant data type -- Set propSet = Nothing
    Case "GET"
        ParseFunction tokens, pa, vbext_mt_Property, INVOKE_PROPERTYGET
    Case "LET"
        Set pt = ParseSub(tokens, pa, vbext_mt_Property, INVOKE_PROPERTYPUT)
        pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes = pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes Or VARIABLE_PUTVAL
Print #99, "ParseProperty: Let pa="; Hex(pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes); " type="; pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtType
'        If pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtType = 0 Then Err.Raise 1
' fixme: procFunctionResultType should be same as last parameter???
'        Set pt.procFunctionResultType = pt.procParams.Item(pt.procParams.count).paramVariable
'        pt.procFunctionResultType.varAttributes = pt.procFunctionResultType.varAttributes Or (pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes And VARIABLE_BYREF)
    Case "SET"
        Set pt = ParseSub(tokens, pa, vbext_mt_Property, INVOKE_PROPERTYPUTREF)
        pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes = pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes Or VARIABLE_PUTVAL
'        pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtDataType = pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtDataType Or pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes Or VARIABLE_BYREF
Print #99, "ParseProperty: Set pa="; Hex(pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes); " type="; pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtType
'        If pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtType = 0 Then Err.Raise 1
'        Set pt.procFunctionResultType = pt.procParams.Item(pt.procParams.count).paramVariable
    Case Else
        Err.Raise 1 ' expecting Get, Let or Set keywords
End Select
Print #99, "ParseProprty: done"
End Sub

Sub ParseStatic(ByVal tokens As Collection)
Dim token As vbToken
tokens.Remove 1 ' bypass Static
Set token = tokens.Item(1)
Select Case UCase(token.tokString)
    Case "FUNCTION"
        ParseFunction tokens, PROC_ATTR_Static
    Case "PROPERTY"
        ParseProperty tokens, PROC_ATTR_Static
    Case "SUB"
        ParseSub tokens, PROC_ATTR_Static
    Case Else
        tokens.Add New vbToken, , 1 ' This is stupid, need to rework parseDim?
        parseDim tokens, PROC_ATTR_Static
End Select
End Sub

' need to implement public/private attr
Sub ParseDeclare(ByVal tokens As Collection, ByVal pa As procattributes)
Print #99, "parsedeclare: pass="; PassNumber
If PassNumber <> 2 Then RemoveAll tokens: Exit Sub
Dim dt As New vbDeclare
Dim token As vbToken
Dim dclType As vbToken
' fixme - PassNumber obsoletes ProcessDeclares
If Not ProcessDeclares Then
    Print #99, "Skipping declare processing"
    MsgBox "skipping declare processing - remove this comment"
    RemoveAll tokens
    Exit Sub
End If
Print #99, "parsedeclare: 1"
If Not currentProc Is Nothing Then Err.Raise 1 ' Declare not allowed in procedure
Print #99, "parsedeclare: 2"
' commented out because Public variables are put into procs collection
'If currentModule.procs.count > 0 Then Err.Raise 1 ' Only comments may appear after Function, Property or Sub

Print #99, "parsedeclare: 3"
tokens.Remove 1 ' remove Declare
Set dt.dclModule = currentModule
dt.dclAttributes = pa
If UCase(tokens.Item(1).tokString) = "FUNCTION" Then
    dt.dclAttributes = dt.dclAttributes Or PROC_ATTR_FUNCTION
ElseIf UCase(tokens.Item(1).tokString) <> "SUB" Then
    Err.Raise 1 ' expecting Function or Sub keyword
End If
Print #99, "parsedeclare: 4"
tokens.Remove 1 ' remove Sub or Function
dt.dclName = tokens.Item(1).tokString ' proc Name
Set token = tokens.Item(1)
tokens.Remove 1 ' remove proc name
If UCase(tokens.Item(1).tokString) = "LIB" Then
    tokens.Remove 1 ' remove Lib
    If tokens.Item(1).tokType <> tokvariant Then Err.Raise 1 ' expecting Lib name to be a string constant
    dt.dclLib = tokens.Item(1).tokValue
    tokens.Remove 1 ' remove Lib string constant
End If
If UCase(tokens.Item(1).tokString) = "ALIAS" Then
    tokens.Remove 1 ' remove Alias
    If tokens.Item(1).tokType <> tokvariant Then Err.Raise 1 ' expecting Alias to be a string constant
    dt.dclAlias = tokens.Item(1).tokValue
Print #99, "Alias="; tokens.Item(1).tokValue
    tokens.Remove 1 ' remove Alias string constant
End If
Print #99, "parsedeclare: 5"
Dim st As SpecialTypes
If tokens.count > 0 Then
    If getSpecialTypes(tokens.Item(1)) = SPECIAL_OP Then
        tokens.Remove 1
        If getSpecialTypes(tokens.Item(1)) = SPECIAL_CP Then
            tokens.Remove 1
        Else
            Do
                Dim p As paramTable
                Dim pp As paramTable
    Print #99, "parsedeclare: 6"
                Set p = getOptionalByAsDataType(tokens, True)
                dt.dclParams.Add p, UCase(p.paramVariable.varSymbol)
                Set pp = p
                st = getSpecialTypes(tokens.Item(1))
                tokens.Remove 1
            Loop While st = special_comma
            If st <> SPECIAL_CP Then Err.Raise 1
        End If
    End If
End If
For Each p In dt.dclParams
    If p.paramVariable.varAttributes And VARIABLE_OPTIONAL Then
        If p.paramVariable.varAttributes And VARIABLE_PARAMARRAY Then
            If dt.dclOptionalParams = -1 Then Err.Raise 1 ' Only last parameter can be ParamArray
            If dt.dclOptionalParams > 0 Then Err.Raise 1 ' Optional parameters not allowed with ParamArray
            dt.dclOptionalParams = -1
        Else
            dt.dclOptionalParams = dt.dclOptionalParams + 1
        End If
    Else
        If dt.dclOptionalParams > 0 Then Err.Raise 1 ' non-optional parameter not allowed after optional parameter
    End If
Next
Print #99, "parsedeclare: 7"
If dt.dclAttributes And PROC_ATTR_FUNCTION Then
    Set dt.dclFunctionResultType = New vbVariable
    dt.dclFunctionResultType.MemberType = vbext_mt_Variable
Print #99, "parseDeclare: 8"
    getAsDataType token, tokens, dt.dclFunctionResultType
Print #99, "parseDeclare: 9"
    If dt.dclFunctionResultType.varType Is Nothing Then Err.Raise 1
Print #99, "parseDeclare: 10"
Else
Print #99, "parseDeclare: 11"
    If Not dt.dclFunctionResultType Is Nothing Then Err.Raise 1
Print #99, "parseDeclare: 12"
End If
currentModule.Declares.Add dt, dt.dclName
Print #99, "parseDeclare: 13"
End Sub

Sub ParseConst(ByVal tokens As Collection, ByVal pa As procattributes)
Dim token As vbToken
Dim output_stack As New Collection
Dim variables As Collection

If currentProc Is Nothing Then
' (obsolete comment?) commented out because Public variables are put into procs collection
    If currentModule.procs.count > 0 Then Err.Raise 1 ' Only comments may appear after Function, Property or Sub
    Set variables = currentModule.Consts
Else
    Set variables = currentProc.procConsts
End If
tokens.Remove 1

On Error Resume Next
Set token = variables.Item(UCase(tokens.Item(1).tokString))
On Error GoTo 0
If Not token Is Nothing Then Err.Raise 1
Dim v As vbVariable
Set v = New vbVariable
variables.Add v, UCase(tokens.Item(1).tokString)
v.MemberType = vbext_mt_Const
v.varSymbol = tokens.Item(1).tokString
v.varAttributes = pa
Set v.varModule = currentModule
Set v.varProc = currentProc
Set token = tokens.Item(1)
tokens.Remove 1
getAsDataType token, tokens, v
Print #99, "parseconst: v dt="; v.varType.dtDataType; " s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "=" Then Err.Raise 1
tokens.Remove 1

oRPN.ConstantRPNize OptimizeConstantExpressions, tokens, output_stack, -1
Print #99, "parseconst: rpn dt="; output_stack.Item(1).tokDataType; " v="; output_stack.Item(1).tokValue

If v.varType.dtDataType = vbVariant Then
    v.varType.dtDataType = output_stack.Item(1).tokDataType
Else
    CoerceOperand output_stack, output_stack.count, v.varType.dtDataType
End If

' rename to varValue?
v.varVariant = output_stack.Item(1).tokValue
End Sub

Sub parseUDT(ByVal tokens As Collection, ByVal pa As procattributes)
Dim t As vbType
Dim v As vbVariable

Print #99, "parseUDT: 1"
If Not currentProc Is Nothing Then Err.Raise 1 ' Type not allowed in procedure
Print #99, "parseUDT: 2 m="; currentModule.Name
' commented out because Public variables are put into procs collection
'If currentModule.procs.count > 0 Then Err.Raise 1 ' Only comments may appear after Function, Property or Sub
Print #99, "parseUDT: 3"
tokens.Remove 1 ' remove type
Print #99, "parseUDT: 4a"
If tokens.count = 0 Then Err.Raise 1 ' expecting Type name
Print #99, "parseUDT: 4b"
If tokens.Item(1).tokType <> toksymbol Then Err.Raise 1
Print #99, "parseUDT: 4c s="; tokens.Item(1).tokString
Dim dt As vbDataType
On Error Resume Next
Set dt = currentModule.Types.Item(UCase(tokens.Item(1).tokString))
On Error GoTo 0
Print #99, "parseUDT: 5"
If Not dt Is Nothing Then Err.Raise 1 ' Type previously defined
Print #99, "parseUDT: 6"
Set dt = New vbDataType
Print #99, "parseUDT: 7"
dt.dtType = tokProjectClass
dt.dtDataType = vbUserDefinedType
Set t = New vbType
Set dt.dtUDT = t
Set dt.dtModule = currentModule
currentModule.Types.Add dt, UCase(tokens.Item(1).tokString)
Print #99, "parseUDT: 10"
t.typeName = tokens.Item(1).tokString
Print #99, "parseUDT: 11"
t.typeAttributes = pa
t.typeGUID = getGUID
tokens.Remove 1
If tokens.count <> 0 Then Err.Raise 1 ' execting end of line after Type name
Print #99, "parseUDT: 12"

Do
Print #99, "parseUDT: 13"
    Do
        getTokenizedLine tokens
Print #99, "parseUDT: 14"
    Loop While tokens.count = 0
Print #99, "parseUDT: 15"
' make end a keyword?
    If UCase(tokens.Item(1).tokString) = "END" Then Exit Do
    On Error Resume Next
    Set v = Nothing
    Set v = t.typeMembers.Item(UCase(tokens.Item(1).tokString))
    On Error GoTo 0
Print #99, "parseUDT: 16"
    If Not v Is Nothing Then Err.Raise 1 ' Duplicate definition
Print #99, "parseUDT: 17"
    t.typeMembers.Add New vbVariable, UCase(tokens.Item(1).tokString)
Print #99, "parseUDT: 18"
    Set v = t.typeMembers.Item(t.typeMembers.count)
    v.MemberType = vbext_mt_Variable
Print #99, "parseUDT: 19"
    v.varSymbol = tokens.Item(1).tokString
    Set v.varModule = currentModule
    Set v.varProc = currentProc
    tokens.Remove 1
Print #99, "parseUDT: 19a"
    parseDimension tokens, v
Print #99, "parseUDT: 19d"
    getAsDataType tokens.Item(1), tokens, v
Print #99, "parseUDT: 20 dt="; v.varType.dtDataType
    If tokens.count <> 0 Then Err.Raise 1 ' Expecting EOL
Loop
Print #99, "parseUDT: 21"
tokens.Remove 1

If tokens.count = 0 Then Err.Raise 1 '  missing Type keyword
Print #99, "parseUDT: 22"
If UCase(tokens.Item(1).tokString) <> "TYPE" Then Err.Raise 1 ' expecting Type after End
tokens.Remove 1
Print #99, "parseUDT: 23"
End Sub

Sub ParseEnum(ByVal tokens As Collection, ByVal pa As procattributes)
Dim output_stack As New Collection
Dim e As vbEnum
Dim em As vbEnumMember
Dim emval As Long

If Not currentProc Is Nothing Then Err.Raise 1 ' Enum not allowed in procedure
' commented out because Public variables are put into procs collection
'If currentModule.procs.count > 0 Then Err.Raise 1 ' Only comments may appear after Function, Property or Sub

tokens.Remove 1
On Error Resume Next
If tokens.count = 0 Then Err.Raise 1 ' expecting Enum
Set e = currentModule.Enums.Item(UCase(tokens.Item(1).tokString))
On Error GoTo 0
If Not e Is Nothing Then Err.Raise 1 ' Enum previously defined
currentModule.Enums.Add New vbEnum, UCase(tokens.Item(1).tokString)
Print #99, "ParseEnum: adding "; tokens.Item(1).tokString; " to "; currentModule.Name
Set e = currentModule.Enums.Item(currentModule.Enums.count)
e.enumName = tokens.Item(1).tokString
e.enumAttributes = pa
tokens.Remove 1
If tokens.count <> 0 Then Err.Raise 1 ' execting end of line after enum name

Do
    Do
        getTokenizedLine tokens
    Loop While tokens.count = 0
' make end a keyword?
    If UCase(tokens.Item(1).tokString) = "END" Then Exit Do
    On Error Resume Next
    Set em = Nothing
    Set em = e.enumMembers.Item(UCase(tokens.Item(1).tokString))
    On Error GoTo 0
    If Not em Is Nothing Then Err.Raise 1
    e.enumMembers.Add New vbEnumMember, UCase(tokens.Item(1).tokString)
    Set em = e.enumMembers.Item(e.enumMembers.count)
    em.enumMemberName = tokens.Item(1).tokString
    tokens.Remove 1
    
    If tokens.count <> 0 Then
        If tokens.Item(1).tokString <> "=" Then Err.Raise 1
        tokens.Remove 1

        oRPN.ConstantRPNize OptimizeConstantExpressions, tokens, output_stack, vbLong
        emval = output_stack.Item(output_stack.count).tokValue
    End If
    em.enumMemberValue = emval
    emval = emval + 1

Loop
tokens.Remove 1

If tokens.count = 0 Then Err.Raise 1 '  missing Enum keyword
If UCase(tokens.Item(1).tokString) <> "ENUM" Then Err.Raise 1 ' expecting Enum after End
tokens.Remove 1

End Sub

Sub ParseEvent(ByVal tokens As Collection, ByVal pa As procattributes)
Dim st As SpecialTypes
Dim pt As proctable
Dim token As vbToken

Print #99, "ParseEvent: pass="; PassNumber
If PassNumber <> 2 Then RemoveAll tokens: Exit Sub

If Not currentProc Is Nothing Then Err.Raise 1 ' Event not allowed in procedure
' commented out because Public variables are put into procs collection
'If currentModule.procs.count > 0 Then Err.Raise 1 ' Only comments may appear after Function, Property or Sub

Set token = tokens.Item(1)
tokens.Remove 1 ' remove Event
Set pt = ProcAdd(tokens.Item(1).tokString, pa Or proc_attr_defined, vbext_mt_Event, INVOKE_EVENTFUNC)
pt.procNumber = currentModule.procs.count
tokens.Remove 1
If tokens.count > 0 Then
    st = getSpecialTypes(tokens.Item(1))
    If st = SPECIAL_OP Then
        tokens.Remove 1
        st = getSpecialTypes(tokens.Item(1))
        If st = SPECIAL_CP Then
            tokens.Remove 1
        Else
            Do
                Dim p As paramTable
                Set p = getByAsDataType(tokens)
                pt.procParams.Add p, UCase(p.paramVariable.varSymbol)
                st = getSpecialTypes(tokens.Item(1))
                tokens.Remove 1
            Loop While st = special_comma
            If st <> SPECIAL_CP Then Err.Raise 1
        End If
    End If
End If
'Set pt.procLocalModule = currentModule
'currentModule.events.Add pt
End Sub

Sub parseMidMidB(ByVal tokens As Collection, ByVal pcode As vbPCodes)
Dim token As vbToken
Dim output_stack As New Collection
Dim stmtMidMidB As vbToken
Print #99, "parseMidMidB: 1"
Set stmtMidMidB = tokens.Item(1)
stmtMidMidB.tokPCode = pcode
tokens.Remove 1

Print #99, "parseMidMidB: 2"
If tokens.Item(1).tokString <> "(" Then Err.Raise 1 ' expecting (
tokens.Remove 1

Print #99, "parseMidMidB: 3"
Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_UNKNOWN) ' object expressions not allowed
' is it correct that UDTs are allowed?
If token.tokDataType <> vbString And token.tokDataType <> vbUserDefinedType Then Err.Raise 1 ' Expecting String variable
output_stack.Add token ' Output string variable immediately

If tokens.Item(1).tokString <> "," Then Err.Raise 1 ' expecting ,
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
If tokens.Item(1).tokString = "," Then
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
Else
    output_stack.Add New vbToken
    output_stack.Item(output_stack.count).tokString = "0"
    output_stack.Item(output_stack.count).tokType = tokvariant
    output_stack.Item(output_stack.count).tokDataType = vbInteger
    output_stack.Item(output_stack.count).tokValue = 0
End If
If tokens.Item(1).tokString <> ")" Then Err.Raise 1 ' expecting )
tokens.Remove 1

' use function to get =?
If tokens.count = 0 Then Err.Raise 1 ' expecting =
If tokens.Item(1).tokString <> "=" Then Err.Raise 1
tokens.Remove 1

oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbString

stmtMidMidB.tokDataType = token.tokDataType
output_stack.Add stmtMidMidB ' add Mid/Midb

currentProc.procStatements.Add output_stack
End Sub

Sub parsePrintWriteStmt(ByVal tokens As Collection, ByVal pcode As vbPCodes)
Dim token As vbToken
Dim output_stack As New Collection

Set token = tokens.Item(1)
tokens.Remove 1 ' remove Print or Write
' Print with no # can only occur in Form, and a few others. Must do futher checking.
If getSpecialTypes(tokens.Item(1)) = SPECIAL_NS Then
    token.tokPCode = pcode
    getFileNumber tokens, output_stack
    If getSpecialTypes(tokens.Item(1)) <> special_comma Then Err.Raise 1
    tokens.Remove 1
Else
    If currentModule.Component.Type <> vbext_ct_VBForm Then Err.Raise 1
    output_stack.Add GetForm(currentModule)
    token.tokPCode = vbPCodePrintMethod
End If
parsePrintWriteExpression tokens, output_stack
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub

Function parsePrintWriteExpression(ByVal tokens As Collection, ByVal output_stack As Collection) As Long
Dim special As SpecialTypes
Print #99, "parsePrintWriteExpression: 1"
Do While tokens.count > 0
    If getKeyword(tokens.Item(1)) = KW_ELSE Then Exit Do ' doesn't feel right, use RPNize return code to exit loop?
    Select Case UCase(tokens.Item(1).tokString)
        Case "SPC"
            tokens.Item(1).tokPCode = vbPCodePrintSpc
        Case "TAB"
            tokens.Item(1).tokPCode = vbPCodePrintTab
    End Select
    Select Case tokens.Item(1).tokPCode
        Case vbPCodePrintSpc, vbPCodePrintTab
            Dim print_func As vbToken
            Set print_func = tokens.Item(1)
            print_func.tokType = tokOperand1
            tokens.Remove 1
            If getSpecialTypes(tokens.Item(1)) <> SPECIAL_OP Then Err.Raise 1 ' expecting (
            tokens.Remove 1
            oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
            If getSpecialTypes(tokens.Item(1)) <> SPECIAL_CP Then Err.Raise 1 ' expecting )
            tokens.Remove 1
            output_stack.Add print_func
            parsePrintWriteExpression = parsePrintWriteExpression + 1
        Case Else
            oRPN.RPNize gOptimizeFlag, tokens, output_stack, -1
            CoerceOperand output_stack, output_stack.count, vbVariant ' force default, if needed
            parsePrintWriteExpression = parsePrintWriteExpression + 1
'            output_stack.Add New vbToken
'            output_stack.Item(output_stack.Count).tokType = tokstatement
'            output_stack.Item(output_stack.Count).tokString = "_printExpr"
    End Select
    If tokens.count = 0 Then Exit Do
    special = getSpecialTypes(tokens.Item(1))
    Select Case tokens.Item(1).tokString
    Case ","
        Set print_func = tokens.Item(1)
        print_func.tokType = tokOperand
        print_func.tokPCode = vbPCodePrintComma
        tokens.Remove 1
        output_stack.Add print_func
        parsePrintWriteExpression = parsePrintWriteExpression + 1
    Case ";"
        Set print_func = tokens.Item(1)
        print_func.tokType = tokOperand
        print_func.tokPCode = vbPCodePrintSemiColon
        tokens.Remove 1
        output_stack.Add print_func
        parsePrintWriteExpression = parsePrintWriteExpression + 1
    Case ":" ' check other stmts for need to exit do if ":"
        Exit Do
    Case Else
        ' don't care
    End Select
Print #99, "PWE="; parsePrintWriteExpression
Loop
Print #99, "parsePrintWriteExpression: c="; tokens.count
End Function

' fixme: parseCircleExpression is incomplete - need to output flags
' object.Circle [Step] (x, y), radius, [color, start, end, aspect]
Sub parseCircleExpression(ByVal tokens As Collection, ByVal arg_stack As Collection)
Print #99, "PCE: 1 s="; tokens.Item(1).tokString
Dim output_stack As Collection
' not implemented - output flag value
arg_stack.Add New Collection, "Step"
arg_stack.Add New Collection, "X"
arg_stack.Add New Collection, "Y"
arg_stack.Add New Collection, "Radius"
arg_stack.Add New Collection, "Color"
arg_stack.Add New Collection, "Start"
arg_stack.Add New Collection, "End"
arg_stack.Add New Collection, "Aspect"
Set output_stack = arg_stack.Item(1) ' Step (flags)
output_stack.Add New vbToken
output_stack.Item(1).tokType = tokvariant
output_stack.Item(1).tokString = "0"
output_stack.Item(1).tokValue = 0
output_stack.Item(1).tokDataType = vbInteger
If UCase(tokens.Item(1).tokString) = "STEP" Then
    ' set flag
    tokens.Remove 1
End If
If tokens.Item(1).tokString <> "(" Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(2), vbSingle ' X
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(3), vbSingle ' Y
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(4), vbSingle ' Radius
If tokens.count = 0 Then GoTo 10
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(5), vbLong, , , , , 0 ' Color
If tokens.count = 0 Then GoTo 20
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(6), vbSingle, , , , , 0! ' Start
If tokens.count = 0 Then GoTo 30
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(7), vbSingle, , , , , 0! ' End
If tokens.count = 0 Then GoTo 40
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(8), vbSingle, , , , , 0! ' Aspect
GoTo 50
10
Set output_stack = arg_stack.Item(5) ' Color
output_stack.Add New vbToken
output_stack.Item(1).tokType = tokvariant
output_stack.Item(1).tokString = "0"
output_stack.Item(1).tokValue = 0&
output_stack.Item(1).tokDataType = vbLong
20
Set output_stack = arg_stack.Item(6) ' Start
output_stack.Add New vbToken
output_stack.Item(1).tokType = tokvariant
output_stack.Item(1).tokString = "0"
output_stack.Item(1).tokValue = 0!
output_stack.Item(1).tokDataType = vbSingle
30
Set output_stack = arg_stack.Item(7) ' End
output_stack.Add New vbToken
output_stack.Item(1).tokType = tokvariant
output_stack.Item(1).tokString = "0"
output_stack.Item(1).tokValue = 0!
output_stack.Item(1).tokDataType = vbSingle
40
Set output_stack = arg_stack.Item(8) ' Aspect
output_stack.Add New vbToken
output_stack.Item(1).tokType = tokvariant
output_stack.Item(1).tokString = "0"
output_stack.Item(1).tokValue = 0!
output_stack.Item(1).tokDataType = vbSingle
50
Print #99, "PCE: s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
If tokens.count <> 0 Then Err.Raise 1
Print #99, "PCE: asc="; arg_stack.count; " osc="; output_stack.count
End Sub

' fixme: parseLine is incomplete - need to output flags
' object.Line [Step] (x1, y1) [Step] - (x2, y2), [color], [B][F]
Sub parseLineExpression(ByVal tokens As Collection, ByVal arg_stack As Collection)
' not implemented - output flag value
Dim output_stack As Collection
Print #99, "PLE: 1 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "(" Then Err.Raise 1
tokens.Remove 1
' fixme: probably should create keys from ParameterInfol.Name. However, don't have it available here.
arg_stack.Add New Collection, "Flags"
arg_stack.Add New Collection, "X1"
arg_stack.Add New Collection, "Y1"
arg_stack.Add New Collection, "X2"
arg_stack.Add New Collection, "Y2"
arg_stack.Add New Collection, "Color"
Set output_stack = arg_stack.Item(1) ' flags
output_stack.Add New vbToken
output_stack.Item(1).tokType = tokvariant
output_stack.Item(1).tokString = "0"
output_stack.Item(1).tokValue = 0
output_stack.Item(1).tokDataType = vbInteger
Print #99, "PLE: 2 s="; tokens.Item(1).tokString
If UCase(tokens.Item(1).tokString) = "STEP" Then
    ' set 1st Step flag
    tokens.Remove 1
End If
Print #99, "PLE: 3 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString = "(" Then
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(2), vbSingle ' X1
    If tokens.Item(1).tokString <> "," Then Err.Raise 1
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(3), vbSingle ' Y1
    If tokens.Item(1).tokString <> ")" Then Err.Raise 1
    tokens.Remove 1
Else
    Set output_stack = arg_stack.Item(2) ' X1
    output_stack.Add New vbToken
    output_stack.Item(1).tokType = tokvariant
    output_stack.Item(1).tokString = "0"
    output_stack.Item(1).tokValue = 0!
    output_stack.Item(1).tokDataType = vbSingle
    Set output_stack = arg_stack.Item(3) ' Y1
    output_stack.Add New vbToken
    output_stack.Item(1).tokType = tokvariant
    output_stack.Item(1).tokString = "0"
    output_stack.Item(1).tokValue = 0!
    output_stack.Item(1).tokDataType = vbSingle
End If
Print #99, "PLE: 4 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "-" Then Err.Raise 1
tokens.Remove 1
If UCase(tokens.Item(1).tokString) = "STEP" Then
    ' set 2nd Step flag
    tokens.Remove 1
End If
Print #99, "PLE: 5 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "(" Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(4), vbSingle ' X2
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(5), vbSingle ' Y2
Print #99, "PLE: 6 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
Print #99, "PLE: 7 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString = "," Then
    tokens.Remove 1
    Print #99, "PLE: 8 s="; tokens.Item(1).tokString
    ' Using default color value of 0. May be necessary to set flag indicating missing color value.
    oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(6), vbLong, , , , , 0& ' Color
    If tokens.count > 0 Then
        If tokens.Item(1).tokString <> "," Then Err.Raise 1
        tokens.Remove 1
        Select Case UCase(tokens.Item(1).tokString)
            Case "B"
                ' fixme: set flag
            Case "BF"
                ' fixme: set flag
            Case Else
                ' F or FB are not allowed
                Err.Raise 1
        End Select
        tokens.Remove 1
    End If
Else
    Print #99, "PLE: 9 s="; tokens.Item(1).tokString
    Set output_stack = arg_stack.Item(6) ' Color
    output_stack.Add New vbToken
    output_stack.Item(1).tokType = tokvariant
    output_stack.Item(1).tokString = "0"
    output_stack.Item(1).tokValue = 0&
    output_stack.Item(1).tokDataType = vbLong
End If
Print #99, "PLE: 10 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
If tokens.count <> 0 Then Err.Raise 1
Print #99, "PLE: asc="; arg_stack.count; " osc="; tokens.count
End Sub

' fixme: parsePSet is incomplete - need to output flags
' object.PSet [Step] (x1, y1) [color]
Sub parsePSetExpression(ByVal tokens As Collection, ByVal output_stack As Collection)
Dim token As vbToken
' not implemented - output flag value
Print #99, "PPE: 1 s="; tokens.Item(1).tokString
Set token = tokens.Item(1)
tokens.Remove 1 ' Line
output_stack.Add New vbToken
output_stack.Item(output_stack.count).tokType = tokvariant
output_stack.Item(output_stack.count).tokString = "0" ' step value
output_stack.Item(output_stack.count).tokValue = 0 ' step value
output_stack.Item(output_stack.count).tokDataType = vbInteger
Print #99, "PPE: 2 s="; tokens.Item(1).tokString
If UCase(tokens.Item(1).tokString) = "STEP" Then
    ' set 1st Step flag
    tokens.Remove 1
End If
Print #99, "PPE: 3 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "(" Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbSingle ' X
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbSingle ' Y
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
If tokens.count > 0 Then
    If tokens.Item(1).tokString <> "," Then Err.Raise 1
    tokens.Remove 1
    ' Using default color value of 0. May be necessary to set flag indicating missing color value.
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbLong, , , , , 0&
Else
    output_stack.Add New vbToken
    output_stack.Item(output_stack.count).tokType = tokvariant
    output_stack.Item(output_stack.count).tokString = "0" ' color value
    output_stack.Item(output_stack.count).tokValue = 0& ' color value
    output_stack.Item(output_stack.count).tokDataType = vbLong
End If
token.tokCount = 4 ' this is kludgy - setting param count - assumes things in tlifunctionargs
If tokens.count = 0 Then tokens.Add token Else tokens.Add token, , 1
Print #99, "PPE osc="; output_stack.count
End Sub

Sub parseScaleExpression(ByVal tokens As Collection, ByVal output_stack As Collection)
Print #99, "parseScaleExpression: tc="; tokens.count
Dim token As vbToken
Set token = tokens.Item(1)
tokens.Remove 1 ' Scale
output_stack.Add New vbToken
output_stack.Item(output_stack.count).tokType = tokvariant
output_stack.Item(output_stack.count).tokString = "0" ' flag value
output_stack.Item(output_stack.count).tokValue = 0 ' flag value
output_stack.Item(output_stack.count).tokDataType = vbInteger
If tokens.count = 0 Then
    Dim i As Integer
    For i = 1 To 4
        output_stack.Add New vbToken
        output_stack.Item(output_stack.count).tokType = tokvariant
        output_stack.Item(output_stack.count).tokString = "0" ' X/Y value
        output_stack.Item(output_stack.count).tokValue = 0! ' X/Y value
        output_stack.Item(output_stack.count).tokDataType = vbVariant ' Yes, Variant
    Next
Else
    ' X1,Y1,X2,Y2 are defined as Variants in IDL
    If tokens.Item(1).tokString <> "(" Then Err.Raise 1
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbVariant ' X1
    If tokens.Item(1).tokString <> "," Then Err.Raise 1
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbVariant ' Y1
    If tokens.Item(1).tokString <> ")" Then Err.Raise 1
    tokens.Remove 1
    If tokens.Item(1).tokString <> "-" Then Err.Raise 1
    tokens.Remove 1
    If tokens.Item(1).tokString <> "(" Then Err.Raise 1
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbVariant ' X2
    If tokens.Item(1).tokString <> "," Then Err.Raise 1
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbVariant ' Y2
    If tokens.Item(1).tokString <> ")" Then Err.Raise 1
    tokens.Remove 1 ' add scale
End If
token.tokCount = 5 ' this is kludgy - setting param count - assumes things in tlifunctionargs
If tokens.count = 0 Then tokens.Add token Else tokens.Add token, , 1
Print #99, "parseScaleExpression: done: tc="; tokens.count
End Sub

                                                                                                                                                                                                                                                                                                                                                                          g variable immediately

If tokens.Item(1).tokString <> "," Then Err.Raise 1 ' expecting ,
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
If tokens.Item(1).tokString = "," Then
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
Else
    output_stack.Add New vbToken
    output_stack.Item(output_stack.count).tokString = "0"
    output_stack.Item(output_stack.count).tokType = tokvariant
    output_stack.Item(output_stack.count).tokDataType 
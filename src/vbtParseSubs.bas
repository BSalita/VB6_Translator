Attribute VB_Name = "vbtParseSubs"
Option Explicit

Function parseDimension(ByVal tokens As Collection, v As vbVariable)
Dim output_stack As New Collection
If Not currentModule Is Nothing Then Print #99, "m="; currentModule.Name
If Not currentProc Is Nothing Then Print #99, "p="; currentProc.procName
If Not IsEOL(tokens) Then
Print #99, "parseDimension: 6 s="; v.varSymbol
    If tokens.Item(1).tokString = "(" Then
        Set v.varDimensions = New Collection
        If tokens.Item(2).tokString = ")" Then
            tokens.Remove 1
        Else
            Do
                tokens.Remove 1 ' remove ) or ,
                oRPN.ConstantRPNize OptimizeConstantExpressions, tokens, output_stack, vbLong
Print #99, "parseDimension: 7"
                Dim vd As vbVarDimension
                Set vd = New vbVarDimension
                v.varDimensions.Add vd
Print #99, "parseDimension: 8"
                If tokens.Item(1).tokKeyword = KW_TO Then
Print #99, "parseDimension: 9"
                    tokens.Remove 1
                    oRPN.ConstantRPNize OptimizeConstantExpressions, tokens, output_stack, vbLong
                    vd.varDimensionLBound = output_stack.Item(1).tokValue ' Expecting a constant expression
                    output_stack.Remove 1
                End If
Print #99, "parseDimension: 10 os.c="; output_stack.Count
                vd.varDimensionUBound = output_stack.Item(1).tokValue
Print #99, "parseDimension: 11"
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
'If PassNumber < 3 And (pa And PROC_ATTR_PRIVATE) = 0 Then RemoveAllTokens tokens: Exit Sub
If PassNumber < 4 Then RemoveAllTokens tokens: Exit Sub
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
        Do While True
            If tokens.Count < 2 Then Err.Raise 1 ' Missing variable
            Print #99, "parseDim: Public s="; tokens.Item(2).tokString
            attr = pa
            If UCase(tokens.Item(2).tokString) = "WITHEVENTS" Then
                tokens.Remove 2 ' remove WithEvents
                attr = attr Or VARIABLE_WITHEVENTS
                If IsEOL(tokens) Then Err.Raise 1 ' Missing variable
            End If
            ParseFunction tokens, attr Or PROC_ATTR_VARIABLE, vbext_mt_Variable, (INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)
            Set currentProc = Nothing
            If IsEOL(tokens) Then Exit Do
            If tokens.Item(1).tokString <> "," Then Err.Raise 1
        Loop
        Exit Sub
    Else
        Set variables = currentModule.ModuleVars
    End If
Else
    Set variables = currentProc.procLocalVariables
End If
Do
    If tokens.Count < 2 Then Err.Raise 1 ' Missing variable
    Print #99, "parseDim: non-Public s="; tokens.Item(2).tokString
    tokens.Remove 1 ' remove Dim, Public, Static, or Private
    If IsEOL(tokens) Then Err.Raise 1 ' Missing variable
    attr = pa
    If UCase(tokens.Item(1).tokString) = "WITHEVENTS" Then
        tokens.Remove 1 ' remove WithEvents
        attr = attr Or VARIABLE_WITHEVENTS
        If IsEOL(tokens) Then Err.Raise 1 ' Missing variable
    End If
Print #99, "parseDim: 3"
    Set token = tokens.Item(1)
    tokens.Remove 1
Print #99, "parseDim: 4 s="; token.tokString; " vc="; variables.Count

' following should be in name check function
' fixme: more checking? make some names (Line) invalid in forms modules.
    On Error Resume Next
    Dim Stmt As vbStmt
    Set Stmt = Nothing
    Set Stmt = cStatements(token.tokString)
    On Error GoTo 0
    If Not Stmt Is Nothing Then
bad_name:
        Print #99, "Name conflict: "; token.tokString
        MsgBox "Name conflict: " & token.tokString
        Err.Raise 1 ' Name conflict
    End If
    If IsForm(currentModule.Component.Type) Then
        Select Case UCase(token.tokString)
        Case "LINE"
            GoTo bad_name
        Case Else
            ' do nothing
        End Select
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
If Not v.varDimensions Is Nothing Then Print #99, "parseDim: 9 vd.c="; v.varDimensions.Count
    getAsDataType token, tokens, v, True
    ' fixme: this code is duped several times
    If Not IsEOL(tokens) Then
        If tokens.Item(1).tokString = "(" Then
            tokens.Remove 1
            If IsEOL(tokens) Then Err.Raise 1 ' Missing )
            If tokens.Item(1).tokString <> ")" Then Err.Raise 1 ' Expecting )
            tokens.Remove 1
            Set v.varDimensions = New Collection
    '        v.varType.dtDataType = v.varType.dtDataType Or VT_ARRAY
        End If
    End If
Print #99, "parsedim: 10 tokens.count=" & tokens.Count & " variables.count="; variables.Count; " v.dt="; v.varType.dtDataType; " type="; v.varType.dtType
    If CBool(v.varAttributes And VARIABLE_WITHEVENTS) And CBool(v.varAttributes And VARIABLE_NEW) Then Err.Raise 1 ' WithEvents can't be used with New
    v.varAttributes = v.varAttributes Or attr
    ' CheckVarDup returns True if Dim statement has been rescaned.
    If Not CheckVarDup(token, variables) Then If variables.Count = 0 Then variables.Add v, UCase(token.tokString) Else variables.Add v, UCase(token.tokString), 1
    If IsEOL(tokens) Then Exit Do
Loop While tokens.Item(1).tokString = ","
Print #99, "parsedim: end tc="; tokens.Count
End Sub

Sub parseReDim(ByVal tokens As Collection)
Dim output_stack As New Collection
Dim token As vbToken
Dim ReDimToken As vbToken
Dim v As vbVariable
'Dim existing_variable As Boolean

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
    Set token = Scope.ScopeLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_PROPERTYGET Or INVOKE_FUNC, False, True, True)
Print #99, "parseReDim: 12 t="; token.tokType; " v="; token.tokVariable.varSymbol; " r="; token.tokRank; " c="; token.tokCount; " dt="; token.tokDataType; " osc="; output_stack.Count
'    token.tokType = tokReDim
'    token.tokCount = 0
    token.tokPCode = tokReDim ' overloading tokPCode
    If token.tokType = tokVariantArgs Then
        token.tokType = tokVariable
        token.tokRank = token.tokCount
        token.tokCount = 0
    End If
    output_stack.Add token
'    If token.tokDataType <> VT_VARIANT Then token.tokCount = 0
'    If output_stack.count = 0 Then output_stack.Add token Else output_stack.Add token, , output_stack.count - token.tokRank - token.tokCount + 1
'    Set token = output_stack(output_stack.count)
    Set v = token.tokVariable
'    token.tokType = tokVariable
    ReDimToken.tokCount = ReDimToken.tokCount + 1 ' Number of Variables
'    Dim rank As Integer
'    rank = 0
' fixme: obsolete RankToken? Use tokRank in each variable token?
'    Dim RankToken As vbToken
'    Set RankToken = New vbToken
'    output_stack.Add RankToken
'    token.tokCount = 0 ' remove subscript count
Print #99, "parseReDim: 15"
    ' new - was getAsDataType but As New Class is allowed
    getAsDataType token, tokens, v, True
    ' Perhaps more data type checking should be done
'    RankToken.tokString = token.tokRank
'    RankToken.tokType = tokVariant
'    RankToken.tokValue = token.tokRank ' remove need for rank token?
'    RankToken.tokDataType = vbLong
For Each token In output_stack
Print #99, "o="; token.tokString; " t="; token.tokType; " v="; token.tokValue
Next
    If IsEOL(tokens) Then Exit Do
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
ProcNameIK = SymIK(tokens.Item(1).tokString, IIf(InvokeKind = INVOKE_FUNC Or InvokeKind = INVOKE_PROPERTYGET, INVOKE_FUNC Or INVOKE_PROPERTYGET, InvokeKind))
Print #99, "ParseSub: pnik="; ProcNameIK; " pa="; Hex(pa)
On Error Resume Next
Set pt = currentModule.procs.Item(ProcNameIK)
On Error GoTo 0
If Not pt Is Nothing Then
#If 0 Then
' fixme: Pass3 reparses, so its a valid redefinition. Ignore dup def for now.
    If pt.procattributes And proc_attr_defined Then
        Print #99, "Duplicate definition: "; pt.procName
        MsgBox "Duplicate definition of """ & pt.procName & """. Event and procedure name same?"
        Err.Raise 1 ' duplicate definition
    End If
#Else
    ' fixme: This garbage is due to problem of not being able to remove collection item by key, so must iterate
    ' fixme: At least put this in a procedure until dictionary is implemented
    Dim i As Long
    For i = 1 To currentModule.procs.Count
        Print #99, currentModule.procs.Item(i).procName
        If pt Is currentModule.procs.Item(i) Then Exit For
    Next
    If i > currentModule.procs.Count Then Err.Raise 1
    currentModule.procs.Remove i
#End If
End If
Set pt = ProcAdd(tokens.Item(1).tokString, pa Or proc_attr_defined, MemberType, InvokeKind)
Set currentProc = pt
tokens.Remove 1
If Not IsEOL(tokens) Then
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
' must test ik separately because of variables being ored.
If InvokeKind = INVOKE_PROPERTYPUT Or InvokeKind = INVOKE_PROPERTYPUTREF Then pt.procParams.Item(pt.procParams.Count).paramVariable.varAttributes = pt.procParams.Item(pt.procParams.Count).paramVariable.varAttributes Or VARIABLE_PUTVAL
For Each p In pt.procParams
    If p.paramVariable.varAttributes And VARIABLE_OPTIONAL Then
        If p.paramVariable.varAttributes And VARIABLE_PUTVAL Then Err.Raise 1 ' Let/Set putval can't be optional
        If p.paramVariable.varAttributes And VARIABLE_PARAMARRAY Then
            If pt.procOptionalParams = -1 Then Err.Raise 1 ' Only last parameter can be ParamArray
            If pt.procOptionalParams > 0 Then Err.Raise 1 ' Optional parameters not allowed with ParamArray
            pt.procOptionalParams = -1
        Else
            pt.procOptionalParams = pt.procOptionalParams + 1
        End If
    Else
        ' non-optional parameter not allowed after optional parameter except for last parameter of Property Let/Set
        If pt.procOptionalParams > 0 Then If Not p.paramVariable.varAttributes And VARIABLE_PUTVAL Then Err.Raise 1
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
getAsDataType token, tokens, pt.procFunctionResultType, pa And PROC_ATTR_VARIABLE
' fixme: this code is duped several times
If Not IsEOL(tokens) Then
    Print #99, "ParseFunction: 2a ts="; tokens.Item(1).tokString
    If tokens.Item(1).tokString = "(" Then
        Print #99, "ParseFunction: 2b"
        tokens.Remove 1
        If IsEOL(tokens) Then Err.Raise 1 ' Missing )
        If tokens.Item(1).tokString <> ")" Then Err.Raise 1 ' Expecting )
        tokens.Remove 1
        Set pt.procFunctionResultType.varDimensions = New Collection
        ' no, musn't change dtDataType
        ' fixme: this argues for creating "varDataType" which inherits from dtDataType Oring in VT_BYREF or VT_ARRAY flags.
'        pt.procFunctionResultType.varType.dtDataType = pt.procFunctionResultType.varType.dtDataType Or VT_ARRAY
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
'        pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes = pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes Or VARIABLE_PUTVAL
'Print #99, "ParseProperty: Let pa="; Hex(pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes); " type="; pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtType
'        If pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtType = 0 Then Err.Raise 1
' fixme: procFunctionResultType should be same as last parameter???
'        Set pt.procFunctionResultType = pt.procParams.Item(pt.procParams.count).paramVariable
'        pt.procFunctionResultType.varAttributes = pt.procFunctionResultType.varAttributes Or (pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes And VARIABLE_BYREF)
    Case "SET"
        Set pt = ParseSub(tokens, pa, vbext_mt_Property, INVOKE_PROPERTYPUTREF)
'        pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes = pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes Or VARIABLE_PUTVAL
'        pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtDataType = pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtDataType Or pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes Or VARIABLE_BYREF
'Print #99, "ParseProperty: Set pa="; Hex(pt.procParams.Item(pt.procParams.count).paramVariable.varAttributes); " type="; pt.procParams.Item(pt.procParams.count).paramVariable.varType.dtType
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
If PassNumber <> 4 Then RemoveAllTokens tokens: Exit Sub
Dim dt As New vbDeclare
Dim token As vbToken
Dim dclType As vbToken
' fixme - PassNumber obsoletes ProcessDeclares
If Not ProcessDeclares Then
    Print #99, "Skipping declare processing"
'    MsgBox "skipping declare processing - remove this comment"
    RemoveAllTokens tokens
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
    If tokens.Item(1).tokType <> tokVariant Then Err.Raise 1 ' expecting Lib name to be a string constant
    dt.dclLib = tokens.Item(1).tokValue
    tokens.Remove 1 ' remove Lib string constant
End If
If UCase(tokens.Item(1).tokString) = "ALIAS" Then
    tokens.Remove 1 ' remove Alias
    If tokens.Item(1).tokType <> tokVariant Then Err.Raise 1 ' expecting Alias to be a string constant
    dt.dclAlias = tokens.Item(1).tokValue
Print #99, "Alias="; tokens.Item(1).tokValue
    tokens.Remove 1 ' remove Alias string constant
End If
Print #99, "parsedeclare: 5"
Dim st As SpecialTypes
If Not IsEOL(tokens) Then
    If getSpecialTypes(tokens.Item(1)) = SPECIAL_OP Then
        tokens.Remove 1
        If getSpecialTypes(tokens.Item(1)) = SPECIAL_CP Then
            tokens.Remove 1
        Else
            Do
                Dim p As paramTable
    Print #99, "parsedeclare: 6a"
                Set p = getOptionalByAsDataType(tokens, True)
    Print #99, "parsedeclare: param="; p.paramVariable.varSymbol
                ' VB bug: allows duplicate parameter names - allcrack.vbp
                Dim pp As paramTable
                Set pp = Nothing
                On Error Resume Next
                Set pp = dt.dclParams.Item(UCase(p.paramVariable.varSymbol))
                On Error GoTo 0
    Print #99, "parsedeclare: 6b"; Not pp Is Nothing
                If Not pp Is Nothing Then
                    Print #99, "Declare statement ("; dt.dclName; ") has duplicate parameter names:"; p.paramVariable.varSymbol
                    MsgBox "Declare statement (" & dt.dclName & ") has duplicate parameter names:" & p.paramVariable.varSymbol
                    Err.Raise 1
                End If
    Print #99, "parsedeclare: 6c"
                dt.dclParams.Add p, UCase(p.paramVariable.varSymbol)
    Print #99, "parsedeclare: 6d"
                st = getSpecialTypes(tokens.Item(1))
                tokens.Remove 1
            Loop While st = special_comma
    Print #99, "parsedeclare: 6e"; st
            If st <> SPECIAL_CP Then Err.Raise 1
    Print #99, "parsedeclare: 6f"
        End If
    End If
End If
    Print #99, "parsedeclare: 6g pc="; dt.dclParams.Count
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
    If Not IsEOL(tokens) Then
        Print #99, "parseDeclare: 10a ts="; tokens.Item(1).tokString
        If tokens.Item(1).tokString = "(" Then
            Print #99, "parseDeclare: 10b"
            tokens.Remove 1
            If IsEOL(tokens) Then Err.Raise 1 ' Missing )
            If tokens.Item(1).tokString <> ")" Then Err.Raise 1 ' Expecting )
            tokens.Remove 1
            Set dt.dclFunctionResultType.varDimensions = New Collection
        End If
        Print #99, "parseDeclare: 10c"
    End If
Else
Print #99, "parseDeclare: 11"
    If Not dt.dclFunctionResultType Is Nothing Then Err.Raise 1
Print #99, "parseDeclare: 12"
End If
' ugh, ignore dups - "Declare Sub d Lib "" (): Dim v" will be scanned twice, once d, again for v
' This suggests that VB extensibility can't be relied upon, must parse all source internally.
' Declare must be first keyword of line.
On Error Resume Next
currentModule.Declares.Add dt, dt.dclName
On Error GoTo 0
Print #99, "parseDeclare: 13"
End Sub

Sub ParseConst(ByVal tokens As Collection, ByVal pa As procattributes)
Print #99, "ParseConst: pass="; PassNumber
If PassNumber = 3 Then RemoveAllTokens tokens: Exit Sub
Print #99, "tc="; tokens.Count; " pa="; Hex(pa)
Print #99, "cm="; Not currentModule Is Nothing; " cp="; Not currentProc Is Nothing
'Dim output_stack As New Collection
Dim Consts As Collection

If currentProc Is Nothing Then
' (obsolete comment?) commented out because Public Consts are put into procs collection
' fixme: doesn't work because of multiple passes
'    If currentModule.procs.count > 0 Then Err.Raise 1 ' Only comments may appear after Function, Property or Sub
    Set Consts = currentModule.Consts
Else
    Set Consts = currentProc.procConsts
End If
Do
    tokens.Remove 1 ' remove Const or comma
    Print #99, "ParseConst: sym="; tokens.Item(1).tokString; " v.c="; Consts.Count
    
    Dim c As vbConst
    Set c = Nothing
    On Error Resume Next
    Set c = Consts.Item(UCase(tokens.Item(1).tokString))
    On Error GoTo 0
    If c Is Nothing Then
        Print #99, "ParseConst: adding "; tokens.Item(1).tokString; " to "; currentModule.Name
' could be Pass 5 If PassNumber <> 1 Then Err.Raise 1 ' internal error
        Set c = New vbConst
        c.ConstName = tokens.Item(1).tokString
        c.ConstAttributes = pa
        Set c.ConstModule = currentModule
        Set c.ConstProc = currentProc
        Consts.Add c, UCase(tokens.Item(1).tokString)
    ElseIf PassNumber = 1 Then
        Err.Raise 1 ' duplicate def
    ElseIf PassNumber = 2 Then
        ' Previous RPN stack may have not-yet-defined Consts (posing as tokVariant). Start reeval anew.
        Set c.ConstRPN = New Collection
'    ElseIf PassNumber = 3 Then
'        ' do nothing
    Else
        Err.Raise 1 ' internal error
    End If
#If 0 Then
    v.MemberType = vbext_mt_Const
    v.varSymbol = tokens.Item(1).tokString
    v.varAttributes = pa
    Set v.varModule = currentModule
    Set v.varProc = currentProc
#End If
    Dim token As vbToken
    Set token = tokens.Item(1)
    tokens.Remove 1 ' remove constant name
    Dim kw As Keywords
    kw = getKeyword(tokens.Item(1))
    If kw = KW_AS Then
        tokens.Remove 1
        Dim dt As vbDataType
        ' fixme: only need simple data types for Const
        Set dt = getDataType(tokens, False)
        If dt.dtDataType = VT_RECORD Then Err.Raise 1 ' UDT not allowed
        Set c.ConstDataType = dt
    Else
        Set c.ConstDataType = cDataTypes("VARIANT")
        Set c.ConstModule = currentModule
    End If
    Print #99, "parseconst: dt="; c.ConstDataType.dtDataType
    If tokens.Item(1).tokString <> "=" Then Err.Raise 1
    tokens.Remove 1
#If 0 Then
    If PassNumber = 1 Then
        oRPN.RPNize OptimizeNone, tokens, c.ConstRPN, c.ConstDataType.dtDataType
    Else
        Dim output_stack As New Collection
        oRPN.ConstantRPNize OptimizeConstantExpressions, tokens, output_stack, -1
        Print #99, "parseconst: rpn dt="; output_stack.Item(1).tokDataType; " v="; output_stack.Item(1).tokValue
        If c.ConstDataType.dtDataType = vbVariant Then
wrong -c.ConstDataType.dtDataType = output_stack.Item(1).tokDataType
        Else
            CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, c.ConstDataType.dtDataType
        End If
        ' rename to varValue?
        c.ConstValue = output_stack.Item(1).tokValue
        Set output_stack = Nothing ' clear and reset for New
    End If
#Else
    If PassNumber = 1 Then
        SkipToNextComma tokens
    Else
        oRPN.RPNize OptimizeNone, tokens, c.ConstRPN, -1 ' c.ConstDataType.dtDataType
        If PassNumber = 5 Then
            ' fixme: this code is dupped multiple times, change something.
            c.ConstValue = oRPN.EvalConstRPNStack(c.ConstRPN, c.ConstDataType.dtDataType, c.ConstDataType.dtLength)
            If c.ConstDataType.dtDataType = vbVariant Then
                If Not IsEmpty(c.ConstValue) And Not IsNull(c.ConstValue) Then
                    Set c.ConstDataType = New vbDataType
                    c.ConstDataType.dtDataType = varType(c.ConstValue)
                End If
            End If
        End If
    End If
#End If
    If IsEOL(tokens) Then Exit Do
Loop While tokens.Item(1).tokString = ","
End Sub

' default scope for UDTs is Public.
Sub parseUDT(ByVal tokens As Collection, ByVal pa As procattributes)
Print #99, "parseUDT: pass="; PassNumber; " s="; tokens.Item(2).tokString; " pa="; pa; " cm="; Not currentModule Is Nothing
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
If IsEOL(tokens) Then Err.Raise 1 ' expecting Type name
Print #99, "parseUDT: 4b"
If tokens.Item(1).tokType <> toksymbol Then Err.Raise 1
Print #99, "parseUDT: 4c s="; tokens.Item(1).tokString
On Error Resume Next
Set t = currentModule.types.Item(UCase(tokens.Item(1).tokString))
On Error GoTo 0
Print #99, "parseUDT: 5 t="; Not t Is Nothing
If t Is Nothing Then
    Print #99, "ParseUDT: adding "; tokens.Item(1).tokString; " to "; currentModule.Name
    If PassNumber <> 1 Then Err.Raise 1 ' internal error
    Set t = New vbType
    currentModule.types.Add t, UCase(tokens.Item(1).tokString)
    Print #99, "parseUDT: 6"
    t.typeName = tokens.Item(1).tokString
    t.typeAttributes = pa
    t.typeGUID = getGUID
    Set t.typeModule = currentModule
ElseIf PassNumber = 1 Then
    Err.Raise 1 ' duplicate def
ElseIf PassNumber = 2 Then
    ' do nothing
ElseIf PassNumber = 3 Then
    ' do nothing
Else
    Err.Raise 1
End If
Print #99, "parseUDT: 11"
tokens.Remove 1 ' remove Type name
If Not IsEOL(tokens) Then Err.Raise 1 ' execting end of line after Type name
Print #99, "parseUDT: 12"

If PassNumber < 3 Then
    Print #99, "parseUDT: 13"
    Do
        currentLineNumber = GetNextTokenizedStatement(currentLineNumber, tokens, True)
    Loop Until UCase(tokens.Item(1).tokString) = "END"
Else
    Print #99, "parseUDT: 14"
    Do
        currentLineNumber = GetNextTokenizedStatement(currentLineNumber, tokens)
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
        Set v = t.typeMembers.Item(t.typeMembers.Count)
        v.MemberType = vbext_mt_Variable
    Print #99, "parseUDT: 19"
        v.varSymbol = tokens.Item(1).tokString
        Set v.varModule = currentModule
    '    Set v.varProc = currentProc  - currentProc is Nothing (default)
        tokens.Remove 1
    Print #99, "parseUDT: 19a"
        parseDimension tokens, v
    Print #99, "parseUDT: 19d"
        getAsDataType tokens.Item(1), tokens, v
    Print #99, "parseUDT: 20 dt="; v.varType.dtDataType
        If v.varType.dtDataType = VT_RECORD Then
            If Not v.varType.dtUDT.typeModule Is currentModule Then
                On Error Resume Next
                ' could error if already added
                currentModule.ModuleDependencies.Add v.varType.dtUDT.typeModule, v.varType.dtUDT.typeModule.Name
                On Error GoTo 0
            End If
        End If
        If Not IsEOL(tokens) Then Err.Raise 1 ' Expecting EOL
    Loop
    If t.typeMembers.Count = 0 Then Err.Raise 1 ' Type has zero members
End If
Print #99, "parseUDT: 21"
tokens.Remove 1 ' End
Print #99, "parseUDT: 22"
If UCase(tokens.Item(1).tokString) <> "TYPE" Then Err.Raise 1 ' expecting Type after End
tokens.Remove 1
Print #99, "parseUDT: 23"
End Sub

Sub ParseEnum(ByVal tokens As Collection, ByVal pa As procattributes)
Print #99, "parseEnum: pass="; PassNumber; " s="; tokens.Item(2).tokString; " pa="; pa; " cm="; Not currentModule Is Nothing; ; " cp="; Not currentProc Is Nothing
'Dim output_stack As New Collection
Dim e As vbEnum

If Not currentProc Is Nothing Then Err.Raise 1 ' Enum not allowed in procedure
' commented out because Public variables are put into procs collection
'If currentModule.procs.count > 0 Then Err.Raise 1 ' Only comments may appear after Function, Property or Sub

tokens.Remove 1 ' remove Enum
If IsEOL(tokens) Then Err.Raise 1 ' expecting Enum name
On Error Resume Next
Set e = currentModule.Enums.Item(UCase(tokens.Item(1).tokString))
On Error GoTo 0
If e Is Nothing Then
    Print #99, "ParseEnum: adding "; tokens.Item(1).tokString; " to "; currentModule.Name
    If PassNumber <> 1 Then Err.Raise 1 ' internal error
    Set e = New vbEnum
    e.enumName = tokens.Item(1).tokString
    e.enumAttributes = pa
    Set e.enumModule = currentModule
    currentModule.Enums.Add e, UCase(tokens.Item(1).tokString)
ElseIf PassNumber = 1 Then
    Err.Raise 1 ' duplicate def
Else
    ' do nothing
End If
tokens.Remove 1
If Not IsEOL(tokens) Then Err.Raise 1 ' execting end of line after enum name

If PassNumber = 3 Then
    Print #99, "parseEnum: 14"
    Do
        currentLineNumber = GetNextTokenizedStatement(currentLineNumber, tokens, True)
    Loop Until UCase(tokens.Item(1).tokString) = "END"
Else
    Print #99, "parseEnum: 14"
    Do
        currentLineNumber = GetNextTokenizedStatement(currentLineNumber, tokens)
    Print #99, "parseEnum: 15"
    ' make end a keyword?
        If UCase(tokens.Item(1).tokString) = "END" Then Exit Do
        Print #99, "s="; tokens.Item(1).tokString
        Dim em As vbEnumMember
        On Error Resume Next
        Set em = Nothing
        Set em = e.enumMembers.Item(UCase(tokens.Item(1).tokString))
        On Error GoTo 0
        Print #99, "em="; Not em Is Nothing
        If PassNumber = 1 Then
            If Not em Is Nothing Then Err.Raise 1 ' duplicate def
            Set em = New vbEnumMember
            e.enumMembers.Add em, UCase(tokens.Item(1).tokString)
            em.enumMemberName = tokens.Item(1).tokString
            Set em.enumMemberParent = e
            RemoveAllTokens tokens
        Else
            If em Is Nothing Then Err.Raise 1 ' internal error
            tokens.Remove 1 ' enum member name
            If Not IsEOL(tokens) Then
                Print #99, "=="; tokens.Item(1).tokString
                If tokens.Item(1).tokString <> "=" Then Err.Raise 1
                tokens.Remove 1 ' remove =
        '        oRPN.ConstantRPNize OptimizeConstantExpressions, tokens, output_stack, vbLong
        '        emval = output_stack.Item(output_stack.count).tokValue
                Print #99, "e="; tokens.Item(1).tokString
#If 1 Then
                oRPN.RPNize OptimizeNone, tokens, em.enumMemberRPN, -1 ' vblong
#Else
                If PassNumber = 1 Then
                    oRPN.RPNize OptimizeNone, tokens, em.enumMemberRPN, vbLong
                Else
                    Dim output_stack As New Collection
                    oRPN.ConstantRPNize OptimizeConstantExpressions, tokens, output_stack, vbLong
                    Print #99, "parseenum: rpn dt="; output_stack.Item(1).tokDataType; " v="; output_stack.Item(1).tokValue
                    em.enumMemberValue = output_stack.Item(1).tokValue
                    Set output_stack = Nothing ' clear and reset for New
                End If
#End If
            End If
        End If
    
    Loop
    If e.enumMembers.Count = 0 Then Err.Raise 1 ' Type has zero members
End If
Print #99, "ParseEnum: 21"
tokens.Remove 1 ' End
If UCase(tokens.Item(1).tokString) <> "ENUM" Then Err.Raise 1 ' expecting Enum after End
tokens.Remove 1

End Sub

Sub ParseEvent(ByVal tokens As Collection, ByVal pa As procattributes)
Dim st As SpecialTypes
Dim pt As proctable
Dim token As vbToken

Print #99, "ParseEvent: pass="; PassNumber
If PassNumber <> 4 Then RemoveAllTokens tokens: Exit Sub

If Not currentProc Is Nothing Then Err.Raise 1 ' Event not allowed in procedure
' commented out because Public variables are put into procs collection
'If currentModule.procs.count > 0 Then Err.Raise 1 ' Only comments may appear after Function, Property or Sub

Set token = tokens.Item(1)
tokens.Remove 1 ' remove Event
Set pt = ProcAdd(tokens.Item(1).tokString, pa Or proc_attr_defined, vbext_mt_Event, INVOKE_EVENTFUNC)
tokens.Remove 1
If Not IsEOL(tokens) Then
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
' INVOKE_PROPERTYPUT for fixed string - fstring1.vbp
' don't want to generate _DEFAULT for Variants - using NoInsertObjDefault=True
Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_PROPERTYGET, , , True) ' INVOKE_UNKNOWN) ' object expressions not allowed
' is it correct that UDTs are allowed?
If Not IsAny(token.tokDataType And Not VT_BYREF, vbString, VT_LPWSTR, vbUserDefinedType) Then Err.Raise 1 ' Expecting String or UDT
output_stack.Add token ' Output string variable immediately

If tokens.Item(1).tokString <> "," Then Err.Raise 1 ' expecting ,
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
If tokens.Item(1).tokString = "," Then
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
Else
    output_stack.Add New vbToken
    output_stack.Item(output_stack.Count).tokString = "0"
    output_stack.Item(output_stack.Count).tokType = tokVariant
    output_stack.Item(output_stack.Count).tokDataType = vbInteger
    output_stack.Item(output_stack.Count).tokValue = 0
End If
If tokens.Item(1).tokString <> ")" Then Err.Raise 1 ' expecting )
tokens.Remove 1

' use function to get =?
If IsEOL(tokens) Then Err.Raise 1 ' expecting =
If tokens.Item(1).tokString <> "=" Then Err.Raise 1
tokens.Remove 1

oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbString

stmtMidMidB.tokDataType = token.tokDataType
output_stack.Add stmtMidMidB ' add Mid/Midb

currentProc.procStatements.Add output_stack
End Sub

Sub parsePrintWriteStmt(ByVal tokens As Collection, ByVal pcode As vbPCodes)
Print #99, "parsePrintWriteStmt: tc="; tokens.Count; " pcode="; pcode; " ct="; currentModule.Component.Type
Print #99, "ts="; tokens.Item(1).tokString
Dim token As vbToken
Dim output_stack As New Collection

Set token = tokens.Item(1)
tokens.Remove 1 ' remove Print/Write keyword
' Print with no # can only occur in Form, and a few others. Must do futher checking.
If IsEOL(tokens) Then GoTo 10
If getSpecialTypes(tokens.Item(1)) = SPECIAL_NS Then
    token.tokPCode = pcode
    getFileNumber tokens, output_stack
    If getSpecialTypes(tokens.Item(1)) <> special_comma Then Err.Raise 1
    tokens.Remove 1
Else
10
    Select Case currentModule.Component.Type
    ' fixme: how many of these really support Print statements? Check it out.
    Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_PropPage, vbext_ct_UserControl ', vbext_ct_ActiveXDesigner
        output_stack.Add GetForm(currentModule)
        token.tokPCode = vbPCodePrintMethod
    Case Else
        ' fixme: "Print 1" is allowed in std and class modules but Print = 1 and Sub Print is not valid. So not allowed?
        Print #99, "Print method (without file number) cannot be used outside of form types."
        MsgBox "Print method (without file number) cannot be used outside of form types."
        Err.Raise 1 ' Method not valid without suitable object
    End Select
End If
parsePrintWriteExpression tokens, output_stack
output_stack.Add token
currentProc.procStatements.Add output_stack
End Sub

Function parsePrintWriteExpression(ByVal tokens As Collection, ByVal output_stack As Collection) As Long
Dim special As SpecialTypes
Print #99, "parsePrintWriteExpression: 1"
Do Until IsEOL(tokens)
    Select Case UCase(tokens.Item(1).tokString)
        Case "SPC"
            tokens.Item(1).tokPCode = vbPCodePrintSpc
        Case "TAB"
            tokens.Item(1).tokPCode = vbPCodePrintTab
        Case Else
            ' do nothing
    End Select
    Select Case tokens.Item(1).tokPCode
        Case vbPCodePrintSpc, vbPCodePrintTab
            Dim print_func As vbToken
            Set print_func = tokens.Item(1)
            tokens.Remove 1
            print_func.tokType = tokOperands
            If IsEOL(tokens) Then GoTo 10
            If getSpecialTypes(tokens.Item(1)) = SPECIAL_OP Then
                tokens.Remove 1 ' (
                oRPN.RPNize gOptimizeFlag, tokens, output_stack, vbInteger
                If getSpecialTypes(tokens.Item(1)) <> SPECIAL_CP Then Err.Raise 1 ' expecting )
                tokens.Remove 1 ' )
                print_func.tokCount = 1
            Else
            ' Tab has optional argument, Spc has mandatory argument
10             If print_func.tokPCode = vbPCodePrintSpc Then Err.Raise 1
            End If
            output_stack.Add print_func
            parsePrintWriteExpression = parsePrintWriteExpression + 1
        Case Else
            oRPN.RPNize gOptimizeFlag, tokens, output_stack, -1
            CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, vbVariant ' force default, if needed
            parsePrintWriteExpression = parsePrintWriteExpression + 1
'            output_stack.Add New vbToken
'            output_stack.Item(output_stack.Count).tokType = tokstatement
'            output_stack.Item(output_stack.Count).tokString = "_printExpr"
    End Select
    If IsEOL(tokens) Then Exit Do
    special = getSpecialTypes(tokens.Item(1))
    Select Case tokens.Item(1).tokString
    Case ","
        Set print_func = tokens.Item(1)
        print_func.tokType = tokOperands
        print_func.tokPCode = vbPCodePrintComma
        tokens.Remove 1
        output_stack.Add print_func
        parsePrintWriteExpression = parsePrintWriteExpression + 1
    Case ";"
        Set print_func = tokens.Item(1)
        print_func.tokType = tokOperands
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
Print #99, "parsePrintWriteExpression: c="; tokens.Count; " osc="; output_stack.Count
End Function

' fixme: parseCircleExpression is incomplete - need to output flags
' object.Circle [Step] (x, y), radius, [color, start, end, aspect]
Sub parseCircleExpression(ByVal tokens As Collection, ByVal arg_stack As Collection)
Print #99, "PCE: 1 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "(" Then Err.Raise 1
tokens.Remove 1
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
output_stack.Item(1).tokType = tokVariant
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
Print #99, "00: "; tokens.Count
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(4), vbSingle ' Radius
Print #99, "10: "; tokens.Count
If tokens.Item(1).tokString = ")" Then GoTo 10
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
Print #99, "20: "; tokens.Count
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(5), vbLong, , , , , 0 ' Color
If tokens.Item(1).tokString = ")" Then GoTo 20
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(6), vbSingle, , , , , 0! ' Start
If tokens.Item(1).tokString = ")" Then GoTo 30
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(7), vbSingle, , , , , 0! ' End
If tokens.Item(1).tokString = ")" Then GoTo 40
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(8), vbSingle, , , , , 0! ' Aspect
GoTo 50
10
Print #99, "100: "; tokens.Count
Set output_stack = arg_stack.Item(5) ' Color
output_stack.Add New vbToken
output_stack.Item(1).tokType = tokVariant
output_stack.Item(1).tokString = "0"
output_stack.Item(1).tokValue = 0&
output_stack.Item(1).tokDataType = vbLong
20
Set output_stack = arg_stack.Item(6) ' Start
output_stack.Add New vbToken
output_stack.Item(1).tokType = tokVariant
output_stack.Item(1).tokString = "0"
output_stack.Item(1).tokValue = 0!
output_stack.Item(1).tokDataType = vbSingle
30
Set output_stack = arg_stack.Item(7) ' End
output_stack.Add New vbToken
output_stack.Item(1).tokType = tokVariant
output_stack.Item(1).tokString = "0"
output_stack.Item(1).tokValue = 0!
output_stack.Item(1).tokDataType = vbSingle
40
Set output_stack = arg_stack.Item(8) ' Aspect
output_stack.Add New vbToken
output_stack.Item(1).tokType = tokVariant
output_stack.Item(1).tokString = "0"
output_stack.Item(1).tokValue = 0!
output_stack.Item(1).tokDataType = vbSingle
Print #99, "200: "; tokens.Count
50
Print #99, "300: "; tokens.Count
Print #99, "PCE: s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
If Not IsEOL(tokens) Then Err.Raise 1
Print #99, "PCE: asc="; arg_stack.Count; " osc="; output_stack.Count
End Sub

Sub parseInputExpression(ByVal tokens As Collection, ByVal arg_stack As Collection)
Print #99, "PIE: 1 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "(" Then Err.Raise 1
tokens.Remove 1
' fixme: probably should create keys from ParameterInfol.Name. However, don't have it available here.
arg_stack.Add New Collection, "Number"
arg_stack.Add New Collection, "FileNumber"
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(1), vbLong ' Number
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
getFileNumber tokens, arg_stack.Item(2) ' FileNumber
Print #99, "PIE: 2 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
' if Input(B) is called as function could be , Else, ... if not iseol(tokens) Then Err.Raise 1
Print #99, "PIE: asc="; arg_stack.Count; " osc="; tokens.Count
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
Dim flags As New vbToken
output_stack.Add flags
flags.tokType = tokVariant
flags.tokValue = 0
flags.tokDataType = vbInteger
Print #99, "PLE: 2 s="; tokens.Item(1).tokString
If UCase(tokens.Item(1).tokString) = "STEP" Then
    flags.tokValue = 1
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
    output_stack.Item(1).tokType = tokVariant
    output_stack.Item(1).tokString = "0"
    output_stack.Item(1).tokValue = 0!
    output_stack.Item(1).tokDataType = vbSingle
    Set output_stack = arg_stack.Item(3) ' Y1
    output_stack.Add New vbToken
    output_stack.Item(1).tokType = tokVariant
    output_stack.Item(1).tokString = "0"
    output_stack.Item(1).tokValue = 0!
    output_stack.Item(1).tokDataType = vbSingle
End If
Print #99, "PLE: 4 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "-" Then Err.Raise 1
tokens.Remove 1
If UCase(tokens.Item(1).tokString) = "STEP" Then
    flags.tokValue = flags.tokValue Or 2
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
    If tokens.Item(1).tokString = "," Then
        tokens.Remove 1
        Select Case UCase(tokens.Item(1).tokString)
            Case "B"
                flags.tokValue = flags.tokValue Or 4
            Case "BF"
                flags.tokValue = flags.tokValue Or 8
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
    output_stack.Item(1).tokType = tokVariant
    output_stack.Item(1).tokString = "0"
    output_stack.Item(1).tokValue = 0&
    output_stack.Item(1).tokDataType = vbLong
End If
Print #99, "PLE: 10 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
If Not IsEOL(tokens) Then Err.Raise 1
flags.tokString = flags.tokValue ' flags
Print #99, "PLE: asc="; arg_stack.Count; " osc="; tokens.Count
End Sub

' fixme: parsePSet is incomplete - need to output flags
' object.PSet [Step] (x1, y1) [color]
Sub parsePSetExpression(ByVal tokens As Collection, ByVal arg_stack As Collection)
Dim token As vbToken
Dim output_stack As Collection
' not implemented - output flag value
Print #99, "PPE: 1 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "(" Then Err.Raise 1
tokens.Remove 1
' fixme: probably should create keys from ParameterInfol.Name. However, don't have it available here.
arg_stack.Add New Collection, "Step"
arg_stack.Add New Collection, "X"
arg_stack.Add New Collection, "Y"
arg_stack.Add New Collection, "Color"
Set output_stack = arg_stack.Item(1) ' Step
Dim flags As New vbToken
output_stack.Add flags
flags.tokType = tokVariant
flags.tokValue = 0 ' step value
flags.tokDataType = vbInteger
Print #99, "PPE: 2 s="; tokens.Item(1).tokString
If UCase(tokens.Item(1).tokString) = "STEP" Then
    flags.tokValue = 1
    tokens.Remove 1
End If
Print #99, "PPE: 3 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "(" Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(2), vbSingle ' X
If tokens.Item(1).tokString <> "," Then Err.Raise 1
tokens.Remove 1
oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(3), vbSingle ' Y
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
If tokens.Item(1).tokString = "," Then
    tokens.Remove 1
    ' Using default color value of 0. May be necessary to set flag indicating missing color value.
    oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(4), vbLong, , , , , 0& ' Color
Else
    Set output_stack = arg_stack.Item(4) ' Color
    output_stack.Add New vbToken
    output_stack.Item(1).tokType = tokVariant
    output_stack.Item(1).tokString = "0" ' color value
    output_stack.Item(1).tokValue = 0& ' color value
    output_stack.Item(1).tokDataType = vbLong
End If
Print #99, "PPE: 10 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
flags.tokString = flags.tokValue ' flags
If Not IsEOL(tokens) Then Err.Raise 1
Print #99, "PPE osc="; output_stack.Count; " asc="; arg_stack.Count
End Sub

Sub parseScaleExpression(ByVal tokens As Collection, ByVal arg_stack As Collection)
Print #99, "parseScaleExpression: tc="; tokens.Count
Dim token As vbToken
Dim output_stack As Collection
' not implemented - output flag value
Print #99, "PSE: 1 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> "(" Then Err.Raise 1
tokens.Remove 1
' fixme: probably should create keys from ParameterInfol.Name. However, don't have it available here.
arg_stack.Add New Collection, "Flags"
arg_stack.Add New Collection, "X1"
arg_stack.Add New Collection, "Y1"
arg_stack.Add New Collection, "X2"
arg_stack.Add New Collection, "Y2"
Set output_stack = arg_stack.Item(1) ' Flags
Dim flags As New vbToken
output_stack.Add flags
flags.tokType = tokVariant
flags.tokValue = 0
flags.tokDataType = vbInteger
Print #99, "PSE: 2 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString = "(" Then
    ' X1,Y1,X2,Y2 are defined as Variants in IDL
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(2), vbVariant ' X1
    If tokens.Item(1).tokString <> "," Then Err.Raise 1
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(3), vbVariant ' Y1
    If tokens.Item(1).tokString <> ")" Then Err.Raise 1
    tokens.Remove 1
    If tokens.Item(1).tokString <> "-" Then Err.Raise 1
    tokens.Remove 1
    If tokens.Item(1).tokString <> "(" Then Err.Raise 1
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(4), vbVariant ' X2
    If tokens.Item(1).tokString <> "," Then Err.Raise 1
    tokens.Remove 1
    oRPN.RPNize gOptimizeFlag, tokens, arg_stack.Item(5), vbVariant ' Y2
    If tokens.Item(1).tokString <> ")" Then Err.Raise 1
    tokens.Remove 1
Else
    Dim i As Integer
    For i = 2 To 5
        Set output_stack = arg_stack.Item(i)
        output_stack.Add New vbToken
        output_stack.Item(1).tokType = tokVariant
        output_stack.Item(1).tokString = "0" ' X/Y value
        output_stack.Item(1).tokValue = 0! ' X/Y value
        output_stack.Item(1).tokDataType = vbSingle
        ' Must coerce. No way to create Single Variant constant. Unusual case.
        CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, vbVariant, Nothing, Nothing, INVOKE_FUNC Or INVOKE_PROPERTYGET
    Next
End If
Print #99, "PSE: 10 s="; tokens.Item(1).tokString
If tokens.Item(1).tokString <> ")" Then Err.Raise 1
tokens.Remove 1
If Not IsEOL(tokens) Then Err.Raise 1
flags.tokString = flags.tokValue ' flags
Print #99, "parseScaleExpression: done: tc="; tokens.Count
End Sub


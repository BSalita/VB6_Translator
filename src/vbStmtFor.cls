VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtFor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

' Contrary to documentation, array elements are allowed as counters.

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim tokStmt As vbToken
Dim output_stack As New Collection
Dim f As vbFor

Set tokStmt = tokens.Item(1)
tokens.Remove 1

If getKeyword(tokens.Item(1)) = KW_EACH Then
    tokens.Remove 1
    tokStmt.tokPCode = vbPCodeforeach
    tokStmt.tokString = "ForEach"
' should LetStmt be True here and below?
'    Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_PROPERTYPUTREF, False, True, True) ' use (INVOKE_FUNC Or INVOKE_PROPERTYGET)?
' don't want to generate _DEFAULT for Variants - using NoInsertObjDefault=True
    Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET, False, True, True) ' use (INVOKE_FUNC Or INVOKE_PROPERTYGET)?
' Type variable allowed - need to test
    Print #99, "t.s="; token.tokString; " t.t="; token.tokType; " t.v="; Not token.tokVariable Is Nothing; " t.dt="; token.tokDataType
    If token.tokVariable Is Nothing Then Err.Raise 1 ' must be a variable
    Print #99, "mt="; token.tokVariable.MemberType
    If token.tokVariable.MemberType <> vbext_mt_Variable Then Err.Raise 1
    If (token.tokDataType And Not VT_BYREF) <> vbObject And (token.tokDataType And Not VT_BYREF) <> vbVariant Then Err.Raise 1 ' expecting variable of data type Object or Variant
' fixme: PROC_ATTR_PUBLIC should be set for Public variables, use tokLocalFunction for now
    If Not token.tokLocalFunction Is Nothing Then
        Print #99, "Public For statement counters are not supported"
        MsgBox "Public For statement counters are not supported"
    End If
    output_stack.Add token
    If UCase(tokens.Item(1).tokString) <> "IN" Then Err.Raise 1
    tokens.Remove 1 ' remove IN
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, -1, INVOKE_FUNC Or INVOKE_PROPERTYGET, False, True, True
Print #99, "vbStmtFor: dt="; output_stack(output_stack.Count).tokDataType
    tokStmt.tokDataType = output_stack(output_stack.Count).tokDataType
    If output_stack(output_stack.Count).tokDataType And VT_ARRAY Then
        ' iterate through array, For variable must be a Variant
        If token.tokDataType <> vbVariant Then Err.Raise 1
    ElseIf (output_stack(output_stack.Count).tokDataType And Not VT_BYREF) = vbVariant Then
        CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, vbObject
    ElseIf IsObj(output_stack(output_stack.Count).tokDataType) Then
        ' interate through a collection, For variable must be object or variant
        Print #99, "v="; Not output_stack(output_stack.Count).tokVariable Is Nothing
        If Not output_stack(output_stack.Count).tokVariable Is Nothing Then
            Print #99, "vt.ii="; Not output_stack(output_stack.Count).tokVariable.varType.dtInterfaceInfo Is Nothing
            If Not output_stack(output_stack.Count).tokVariable.varType.dtInterfaceInfo Is Nothing Then
                Print #99, "ii="; output_stack(output_stack.Count).tokVariable.varType.dtInterfaceInfo.Name; "am="; Hex(output_stack(output_stack.Count).tokVariable.varType.dtInterfaceInfo.AttributeMask)
                If output_stack(output_stack.Count).tokVariable.varType.dtInterfaceInfo.AttributeMask And (1 Or TYPEFLAG_FDUAL Or TYPEFLAG_FDISPATCHABLE) Then ' 1 is hidden flag
                    ' do nothing
                Else
                    ' fixme: can object fail this test but have legit dispatchable implied interface?
                    Print #99, "vbStmtFor: Invalid collection interface - object not dispatchable"
                    MsgBox "For Each statement collection object not dispatchable: " & token.tokString
                    Err.Raise 1
                End If
            End If
        End If
    Else
        Err.Raise 1 ' expecting collection of data type Object or Variant
    End If

    On Error Resume Next
    Set f = currentProc.procFors.Item(UCase(token.tokString))
    On Error GoTo 0
    If Not f Is Nothing Then Err.Raise 1 ' For variable in use
    Set f = New vbFor
    If currentProc.procFors.Count = 0 Then currentProc.procFors.Add f Else currentProc.procFors.Add f, UCase(token.tokString), 1
    Set f.stmtTok = tokStmt
    Set f.forVariable = token.tokVariable
' Need to add RepeatLabel and ExitLabel code
    If currentProc.procNests.Count = 0 Then currentProc.procNests.Add f Else currentProc.procNests.Add f, , 1
' Need to add NestLabel
Else
    tokStmt.tokPCode = vbPCodefor
' fixme: using LetStmt parameter to return on = - use for Let, Set, ...?
    Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET, False, True)
' Type variable allowed - need to test
    Print #99, "t.s="; token.tokString; " t.t="; token.tokType; " t.v="; Not token.tokVariable Is Nothing; " t.dt="; token.tokDataType
    If token.tokVariable Is Nothing Then Err.Raise 1 ' must be a variable
    Print #99, "mt="; token.tokVariable.MemberType
    If token.tokVariable.MemberType <> vbext_mt_Variable Then Err.Raise 1
    If token.tokDataType = vbObject Then Err.Raise 1 ' Object ref not allowed
    If token.tokDataType = vbBoolean Then Err.Raise 1 ' Boolean not allowed
' fixme: PROC_ATTR_PUBLIC should be set for Public variables, use tokLocalFunction for now
    If Not token.tokLocalFunction Is Nothing Then
        Print #99, "Public For statement counters are not supported"
        MsgBox "Public For statement counters are not supported"
    End If
    output_stack.Add token
    tokStmt.tokDataType = token.tokDataType
    
    On Error Resume Next
    Set f = currentProc.procFors.Item(UCase(token.tokString))
    On Error GoTo 0
    If Not f Is Nothing Then Err.Raise 1 ' For variable in use
    Set f = New vbFor
    If currentProc.procFors.Count = 0 Then currentProc.procFors.Add f Else currentProc.procFors.Add f, UCase(token.tokString), 1
    Set f.stmtTok = tokStmt
    Set f.forVariable = token.tokVariable
' Need to add RepeatLabel and ExitLabel code
    If currentProc.procNests.Count = 0 Then currentProc.procNests.Add f Else currentProc.procNests.Add f, , 1
' Need to add NestLabel

    If tokens.Item(1).tokString <> "=" Then Err.Raise 1
    tokens.Remove 1

    oRPN.RPNize gOptimizeFlag, tokens, output_stack, tokStmt.tokDataType And Not VT_BYREF
    
    If UCase(tokens.Item(1).tokString) <> "TO" Then Err.Raise 1 ' expecting To keyword
    tokens.Remove 1
    
    oRPN.RPNize gOptimizeFlag, tokens, output_stack, tokStmt.tokDataType And Not VT_BYREF
    
    If IsEOL(tokens) Then GoTo 10
    If UCase(tokens.Item(1).tokString) = "STEP" Then ' could be : or Else
        tokens.Remove 1 ' remove Step
        oRPN.RPNize gOptimizeFlag, tokens, output_stack, tokStmt.tokDataType And Not VT_BYREF
    Else
10
        output_stack.Add New vbToken
        output_stack.Item(output_stack.Count).tokString = "1"
        output_stack.Item(output_stack.Count).tokType = tokVariant
        output_stack.Item(output_stack.Count).tokDataType = vbInteger
        output_stack.Item(output_stack.Count).tokValue = 1 ' should this be same data type as NextVar?
        CoerceOperand gOptimizeFlag, output_stack, output_stack.Count, tokStmt.tokDataType And Not VT_BYREF
    End If
End If

output_stack.Add tokStmt ' add For

currentProc.procStatements.Add output_stack
End Sub



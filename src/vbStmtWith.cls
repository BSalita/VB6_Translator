VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbStmtWith"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements vbStmt

Sub vbStmt_stmtParse(ByVal tokens As Collection)
Dim token As vbToken
Dim stmtTok As vbToken
Dim output_stack As New Collection

Set stmtTok = tokens.Item(1)
stmtTok.tokPCode = vbPCodeWith
tokens.Remove 1

Set token = SymbolLookUp(gOptimizeFlag, tokens, output_stack, INVOKE_FUNC Or INVOKE_PROPERTYGET, True, True, True)
Print #99, "With: dt="; token.tokDataType
Select Case token.tokDataType And Not VT_BYREF
    Case vbUserDefinedType
        If token.tokVariable Is Nothing Then
            Print #99, "With: Bad With expression"
            Err.Raise 1 ' With Nothing should be invalid
        ElseIf Not token.tokVariable.varType.dtRecordInfo Is Nothing Then
            Print #99, "With: n="; token.tokVariable.varType.dtRecordInfo.Name
        ElseIf Not token.tokVariable.varType.dtUDT Is Nothing Then
            Print #99, "With: n="; token.tokVariable.varType.dtUDT.typeName
        Else
            Err.Raise 1
        End If
    Case vbObject, VT_UNKNOWN, vbVariant
        If token.tokVariable Is Nothing Then
            Print #99, "With: Bad With expression"
            Err.Raise 1 ' With Nothing should be invalid
        ElseIf Not token.tokVariable.varType.dtInterfaceInfo Is Nothing Then
            Print #99, "With: ii.n="; token.tokVariable.varType.dtInterfaceInfo.Name; " guid="; token.tokVariable.varType.dtInterfaceInfo.GUID
        ElseIf Not token.tokVariable.varType.dtClass Is Nothing Then
            Print #99, "With: m.n="; token.tokVariable.varType.dtClass.Name; " guid="; token.tokVariable.varType.dtClass.GUID
        Else
            ' do nothing - could be untyped (Object) or Variant
        End If
    Case Else
        Err.Raise 1 ' expecting variable of data type UDT, Object or Variant
End Select
' Init With FIFO's after RPNizing, in case sane With stack is needed in RPN
Dim withTok As vbWith
Set withTok = New vbWith
Set stmtTok.tokWith = withTok
Set withTok.WithValue = token
If currentProc.procWiths.count = 0 Then currentProc.procWiths.Add withTok Else currentProc.procWiths.Add withTok, , 1
withTok.WithCount = currentProc.procWiths.count
If currentProc.procNests.count = 0 Then currentProc.procNests.Add withTok Else currentProc.procNests.Add withTok, , 1
output_stack.Add token
output_stack.Add stmtTok
currentProc.procStatements.Add output_stack
End Sub


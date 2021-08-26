VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbEmiter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements Emiter

Private m_Group As vbGroup
Private TotalLines As Single
Private LineCount As Single

Sub Emiter_Emiter(group As vbGroup)

Dim p As vbPrj
Dim m As vbModule
Dim proc As procTable

Set m_Group = group
Open group.OutputPath & "vbEmitter.lst" For Output As 99
For Each p In group.Projects
    For Each m In p.prjModules
        For Each proc In m.procs
            TotalLines = TotalLines + proc.procStatements.Count
        Next
        For Each proc In m.Events
            TotalLines = TotalLines + proc.procStatements.Count
        Next
    Next
Next
TotalLines = TotalLines / 100
For Each p In group.Projects
    Dim OutputPath As String
    OutputPath = VBTOutputPath(group, p, "bas")
    For Each m In p.prjModules
Print #99, "1 " & m.Name
        m_Group.AddIn.ProgressBarInit "VB Generator: Processing " & m.Name
Print #99, "2 " & m.Component.Type
        Select Case m.Component.Type
            Case vbext_ct_StdModule
Print #99, "std="; m.Name
                Open OutputPath & m.Name & ".bas" For Output As 1
Print #99, p.VBProject.FileName; " - "; OutputPath
                Print #99, "project path=" & p.VBProject.FileName & " - " & OutputPath
                Print #1, "Attribute VB_Name = """; m.Name; """"
            Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_UserControl
                Open OutputPath & m.Name & ".frm" For Output As 1
' fixme: Forms BEGIN/END prologue is not implemented!!!! Using class stuff for now
                Dim prop As Property
                Print #1, "VERSION 5.00" ' don't know how to get version number - hard code for now
                Print #1, "Begin VB.Form "; m.Name
                On Error Resume Next
                For Each prop In m.Component.Properties
Print #99, "prop: " & prop.Name
                    Select Case UCase(prop.Name)
                        Case "NAME"
                        Case Else
                            Select Case VarType(prop.Value)
                                Case vbObject
'fixme: What to do here?    Print #1, TypeName(prop.Value)
                                Case vbString
                                    Print #1, Tab(7); prop.Name; Tab(23); "=   "; """"; prop.Value; """"
                                Case Else
                                    Print #1, Tab(7); prop.Name; Tab(23); "=   "; prop.Value
                            End Select
                    End Select
                Next
                On Error GoTo 0
                Dim f As Object
Print #99, "3"
                Set f = m.Component.Designer
Print #99, "4"
                Dim vbc As VBControl
Print #99, "VBControls: "; f.VBControls.Count
                For Each vbc In f.VBControls
                    Print #99, "ProgId="; vbc.ProgId; " Name="; vbc.Properties.Item("Name").Value ' Do not use vbc.ClassName
                    Print #1, Tab(4); "Begin "; vbc.ProgId; " "; vbc.Properties.Item("Name").Value ' Do not use vbc.ClassName
                    On Error Resume Next
                    For Each prop In vbc.Properties
                        Print #99, "prop.Name="; prop.Name
                        Select Case UCase(prop.Name)
                            Case "NAME"
                            Case "INDEX"
                                If prop.Value <> -1 Then GoTo 20 ' fixme: create Sub to do this
                            Case Else
                                Print #99, "vt="; VarType(prop.Value)
                                Select Case VarType(prop.Value)
                                    Case vbObject
'fixme: What to do here?            Print #1, TypeName(prop.Value)
                                    Case vbString
                                        Print #1, Tab(7); prop.Name; Tab(23); "=   "; """"; prop.Value; """"
                                    Case Else
20                                      Print #1, Tab(7); prop.Name; Tab(23); "=   "; prop.Value
                                End Select
                        End Select
                    Next
                    On Error GoTo 0
                    Print #1, Tab(4); "End"
                Next
                Print #1, "End"
                GoTo 10
            Case vbext_ct_ClassModule
                Open OutputPath & m.Name & ".cls" For Output As 1
                Print #1, "VERSION "; CStr(m.clsVersionMajor); "."; CStr(m.clsVersionMinor); " CLASS"
                Print #1, "BEGIN" ' To do: VB outputs a comment after following values
                Print #1, "  MultiUse = "; CStr(m.clsMultiUse)
                Print #1, "  Persistable = "; CStr(m.clsPersistable)
                Print #1, "  DataBindingBehavior ="; m.clsDataBindingBehavior
                Print #1, "  DataSourceBehavior ="; m.clsDataSourceBehavior
                Print #1, "  MTSTransactionMode ="; m.clsMTSTransactionMode
                Print #1, "END"
10
                Print #1, "Attribute VB_Name = """; m.clsVB_Name; """"
                Print #1, "Attribute VB_GlobalNameSpace = "; m.clsVB_GlobalNameSpace
                Print #1, "Attribute VB_Creatable = "; m.clsVB_Creatable
                Print #1, "Attribute VB_PredeclaredId = "; m.clsVB_PredeclaredID
                Print #1, "Attribute VB_Exposed = "; m.clsVB_Exposed
' fixme: implement VB_Ext_KEY
'                Print #1, "Attribute VB_Ext_KEY = ""SavedWithClassBuilder6"" ,""Yes"""
'                Print #1, "Attribute VB_Ext_KEY = ""Top_Level"" ,""Yes"""
            Case Else
                Print #99, "VBemitter: unknown component type: "; m.Component.Type
                MsgBox "VBemitter: unknown component type: " & m.Component.Type
                Err.Raise 1
        End Select
        Print #1, "' This file is always generated by Softworks VBT"
        OutputModule m
        Close #1
Print #99, "4"
    Next ' module
Print #99, "5"
Next ' project
Print #99, "Emitter: done"
Close #99
End Sub

Sub OutputModule(ByVal m As vbModule)
Print #99, "OutputModule: m="; m.Name
Dim first_time As Boolean
Dim s As String
If m.OptionExplicit Then Print #1, "Option Explicit"
If m.clsImplements.Count > 0 Then Print #1,
Dim dt As vbDataType
For Each dt In m.clsImplements
    If dt.dtType = tokProjectClass Then
        s = "Implements " & dt.dtClass.Name
    ElseIf dt.dtType = tokReferenceClass Then
        s = "Implements " & dt.dtClassInfo.Parent.Name & "." & dt.dtClassInfo.Name
    Else
        Err.Raise 1
    End If
    Print #99, s
    Print #1, s
Next
Print #99, "c.c="; m.Consts.Count
If m.Consts.Count > 0 Then Print #1,
Dim c As vbConst
For Each c In m.Consts
    s = vbScopeAttributes(c.ConstAttributes, "") & "Const " & c.ConstName & IIf(c.ConstDataType.dtDataType = vbVariant, "", " As " & vbDataType(c.ConstDataType)) & " = " & EmitInFix(m, c.ConstRPN)
    If c.ConstRPN.Count > 1 Or c.ConstRPN.Item(1).tokType <> tokVariant Then s = s & " ' " & CStr(vbOutputVariant(c.ConstValue))
    Print #99, s
    Print #1, s
Next
Print #99, "e.c="; m.Enums.Count
Dim e As vbEnum
For Each e In m.Enums
    Print #1,
    s = vbScopeAttributes(e.enumAttributes, "") & "Enum " & e.enumName
    Print #99, s
    Print #1, s
    Dim em As vbEnumMember
    Dim emval As Long
    emval = 0
    For Each em In e.enumMembers
        s = IndentString & em.enumMemberName
        If em.enumMemberValue <> emval Then
            emval = em.enumMemberValue
            s = s & " = " & EmitInFix(m, em.enumMemberRPN) ' CStr(emval);
        End If
        s = s & " ' " & CStr(emval)
        emval = emval + 1
        Print #99, s
        Print #1, s
    Next
    s = "End Enum"
    Print #99, s
    Print #1, s
Next
Dim t As vbType
Print #99, "t.c="; m.types.Count
For Each t In m.types
    Print #1,
    s = vbScopeAttributes(t.typeAttributes, "") & "Type " & t.TypeName
    Print #99, s
    Print #1, s
    Dim v As vbVariable
    For Each v In t.typeMembers
        s = IndentString & vbTypeName(v)
        Print #99, s
        Print #1, s
    Next
    s = "End Type"
    Print #99, s
    Print #1, s
Next
Dim d As vbDeclare
Print #99, "d.c="; m.Declares.Count
If m.Declares.Count > 0 Then Print #1,
For Each d In m.Declares
    s = vbScopeAttributes(d.dclAttributes, "") & "Declare "
    If d.dclAttributes And PROC_ATTR_FUNCTION Then s = s & "Function " Else s = s & "Sub "
    s = s & d.dclName & " Lib """ & d.dclLib & """"
    If d.dclAlias <> "" Then s = s & " Alias """ & d.dclAlias & """"
    s = s & " ("
    Dim param As paramTable
    first_time = True
    For Each param In d.dclParams
        If first_time Then first_time = False Else s = s & ", "
        s = s & vbOptionalByVarType(param.paramVariable)
    Next
    s = s & ") "
    If d.dclAttributes And PROC_ATTR_FUNCTION Then
        s = s & vbVariableType(d.dclFunctionResultType)
        If Not d.dclFunctionResultType.varDimensions Is Nothing Then s = s & "()"
    End If
    Print #99, s
    Print #1, s
Next
Print #99, "v.c="; m.ModuleVars.Count
If m.ModuleVars.Count > 0 Then Print #1,
For Each v In m.ModuleVars
    s = vbScopeAttributes(v.varAttributes, "Dim")
    If s = "" Then Err.Raise 1 ' Internal Error
    s = s & vbTypeName(v)
    Print #99, s
    Print #1, s
Next
' Must output Event declarations in declarations section -- ahead of methods
If m.Events.Count > 0 Then Print #1,
Dim proc As procTable
For Each proc In m.Events
    s = ""
    For Each param In proc.procParams
        s = s & "," & vbOptionalByVarType(param.paramVariable)
    Next
    s = vbScopeAttributes(proc.procattributes, "Event") & proc.procName & "(" & Mid(s, 2) & ")"
    Print #99, s
    Print #1, s
Next
OutputMethods m, m.procs
End Sub

' TODO: eliminate vbModule parameter by adding Module property to procTable
Sub OutputMethods(ByVal m As vbModule, ByVal procs As Collection)
Print #99, "OutputMethods: m="; m.Name; " procs="; procs.Count

Dim proc As procTable
Dim s As String

For Each proc In procs
    If proc.procattributes And PROC_ATTR_FORM_CONTROL Then
        ' do nothing
    ElseIf proc.procattributes And PROC_ATTR_VARIABLE Then
        s = vbScopeAttributes(proc.procattributes, "Dim")
        If s = "" Then Err.Raise 1 ' Internal Error
        s = s & proc.procName & vbVariableType(proc.procFunctionResultType)
        Print #99, s
        Print #1, s
    End If
Next

Dim param As paramTable
Dim first_time As Boolean

For Each proc In procs
    If proc.procattributes And PROC_ATTR_FORM_CONTROL Then
        ' do nothing
    ElseIf proc.procattributes And PROC_ATTR_VARIABLE Then
        ' do nothing
    Else
        Print #1,
        s = vbScopeAttributes(proc.procattributes, "")
        Select Case proc.InvokeKind
            Case INVOKE_PROPERTYGET
                s = s & "Property Get "
            Case INVOKE_PROPERTYPUT
                s = s & "Property Let "
            Case INVOKE_PROPERTYPUTREF
                s = s & "Property Set "
            Case Else
                If proc.procattributes And PROC_ATTR_FUNCTION Then
                    s = s & "Function "
                Else
                    s = s & "Sub "
                End If
        End Select
        s = s & proc.procName & "("
        first_time = True
        For Each param In proc.procParams
            If first_time Then first_time = False Else s = s & ", "
            s = s & vbOptionalByVarType(param.paramVariable)
        Next
        s = s & ")"
        If proc.procattributes And PROC_ATTR_FUNCTION Then
            s = s & vbVariableType(proc.procFunctionResultType)
            If Not proc.procFunctionResultType.varDimensions Is Nothing Then s = s & "()"
        End If
        Print #99, s
        Print #1, s
        Dim v As vbVariable
        For Each v In proc.procLocalVariables
            ' TODO: Declarations of implicitly declared variables is suppressed. Convert to user-selectable option.
            s = vbScopeAttributes(v.varAttributes, "Dim")
            If s <> "" Then
                s = s & vbTypeName(v)
                Print #99, s
                Print #1, s
            End If
        Next
        Dim stmts As Collection
        For Each stmts In proc.procStatements
            m_Group.AddIn.ProgressBarUpdate LineCount / TotalLines
            s = EmitInFix(m, stmts)
            LineCount = LineCount + 1
            Print #99, s
            Print #1, s
        Next
        Print #1, "End ";
        If proc.MemberType = vbext_mt_Property Then
            s = "Property"
        ElseIf proc.procattributes And PROC_ATTR_FUNCTION Then
            s = "Function"
        Else
            s = "Sub"
        End If
        Print #99, s
        Print #1, s
    End If
Next

End Sub

Function vbOptionalByVarType(ByVal variable As vbVariable) As String
vbOptionalByVarType = vbTypeName(variable)
If variable.varAttributes And VARIABLE_PARAMARRAY Then
    vbOptionalByVarType = "ParamArray " & vbOptionalByVarType
    If VarType(variable.varVariant) <> vbEmpty Then vbOptionalByVarType = vbOptionalByVarType & " = " & vbOutputVariant(variable.varVariant)
Else
    If variable.varAttributes And VT_BYREF Then vbOptionalByVarType = "ByRef " & vbOptionalByVarType Else vbOptionalByVarType = "ByVal " & vbOptionalByVarType
    If variable.varAttributes And VARIABLE_OPTIONAL Then
        vbOptionalByVarType = "Optional " & vbOptionalByVarType
        If VarType(variable.varVariant) <> vbEmpty Then vbOptionalByVarType = vbOptionalByVarType & " = " & vbOutputVariant(variable.varVariant)
    End If
End If
End Function


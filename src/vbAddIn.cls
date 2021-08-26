VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_group As vbGroup
Private p As VBProject

Function AddInMain(vbInstance As VBIDE.VBE, clsAddIn As Object) As vbGroup
OutputPath = App.Path & "\"
Open OutputPath & "vbt.lst" For Output As 99
If Not g_main_inited Then
    Print #99, "Main not inited!"
    MsgBox "Main not inited!"
End If
Print #99, "OutputPath=" & OutputPath & " fullname=" & vbInstance.FullName
If Not vbInstance.ActiveVBProject Is Nothing Then Print #99, "active=" & vbInstance.ActiveVBProject.Name
If Not vbInstance.SelectedVBComponent Is Nothing Then Print #99, "selected=" & vbInstance.SelectedVBComponent.Name

Set oMe = Me
Print #99, "1"
Set m_group = group
Set group.vbInstance = vbInstance
group.Name = "groupname" ' fixme: don't know how to get group name from IDE!! -- hardcode for now
Print #99, "2"
Set group.AddIn = clsAddIn
Print #99, "3"
group.OutputPath = OutputPath
Print #99, "4"

'Set msTLI = TypeLibInfoFromFile("MSVBVM60.DLL")
'Set msTLIDeclarations = msTLI.Declarations
'Set msTLIConstants = msTLI.Constants
'gOptimizeFlag = OptimizeNone
gOptimizeFlag = OptimizeConstantExpressions

Dim TotalMembers As Single
Dim MemberCount As Single
Print #99, "5"
' fixme: need to be able to process group if "Make Group" menu option is selected.
' Group processing must occur after all group's projects are compiled.
' For now, only compile active project until group is implemented.
    Set p = vbInstance.ActiveVBProject
'For Each p In group.vbInstance.VBProjects
    Print #99, "p="; p.Name
    Print #99, "VBc="; Not p.VBComponents Is Nothing
    If Not p.VBComponents Is Nothing Then
        For Each currentComponent In p.VBComponents
            Print #99, "Type:"; currentComponent.Type
            Print #99, "IsDirty:"; currentComponent.IsDirty
            If IsModule(currentComponent) Then
                Print #99, "cm="; Not currentComponent.CodeModule Is Nothing
                If Not currentComponent.CodeModule Is Nothing Then If Not currentComponent.CodeModule.Members Is Nothing Then TotalMembers = TotalMembers + currentComponent.CodeModule.Members.Count
            End If
            Print #99, " 5b"
        Next
        Print #99, " 5c"
    End If
    Print #99, " 5d"
'Next
Print #99, "6"
TotalMembers = TotalMembers / 100
Print #99, "7"

'appData.appName = group.vbinstance.Name ' What should this be?
'For Each p In group.vbInstance.VBProjects
    PassNumber = 0
    Dim project As vbPrj
Print #99, "8"
    Set project = New vbPrj
    group.Projects.Add project
    Set currentProject = project
    Set project.VBProject = p
Print #99, "9"
    Dim refVBA As reference
    On Error Resume Next
    Set refVBA = p.References.Item("VBA") ' could do just once?
    On Error GoTo 0
    If refVBA Is Nothing Then
        Print #99, "Can't find VBA TypeLib. Check Project->References for errors."
        MsgBox "Can't find VBA TypeLib. Check Project->References for errors."
        Err.Raise 1
    Else
' TypeLibInfoFromRegistry fails. Is VBA typelib not in registry?
'        Set tlibVBA = TypeLibInfoFromRegistry(refVBA.GUID, refVBA.Major, refVBA.Minor, 0)
        Set tlibVBA = TypeLibInfoFromFile(refVBA.FullPath)
        Set refVBA = Nothing
    End If
    ' StdOle stuff is used only to retrieve IDispatch INTERFACEINFO. Could hard code instead.
    Dim tlibStdOle As TypeLibInfo
    Set tlibStdOle = TypeLibInfoFromRegistry("{00020430-0000-0000-C000-000000000046}", 2, 0, 0)
    Set UnknownInterfaceInfo = tlibStdOle.Interfaces.NamedItem("IUNKNOWN")
    Set DispatchInterfaceInfo = tlibStdOle.Interfaces.NamedItem("IDISPATCH")
    Set tlibStdOle = Nothing
Print #99, "10"
    On Error Resume Next
    Set project.TypeLib = TypeLibInfoFromFile(p.BuildFileName)
    On Error GoTo 0
Print #99, "11"
    If Not project.TypeLib Is Nothing Then Print #99, "12 tlibname="; project.TypeLib.Name
    On Error Resume Next
    Set tliVB = p.References.Item("VB") ' could do just once?
    On Error GoTo 0
    If tliVB Is Nothing Then
        Print #99, "Can't find VB TypeLib. Check Project->References for errors."
        MsgBox "Can't find VB TypeLib. Check Project->References for errors."
        Err.Raise 1
    End If
Print #99, "p.Name="; p.Name
    project.prjName = p.Name
    ValidateReferenceTLibs project.VBProject.References
    ' Process each component module (not looking at code yet).
    For Each currentComponent In p.VBComponents
        Dim m As vbModule
        Print #99, "Pass0: Process Component: "; currentComponent.Name
        clsAddIn.ProgressBarInit "Pass0: Processing " & currentComponent.Name
        If IsModule(currentComponent) Then
            If Not currentComponent.CodeModule Is Nothing Then
                Set m = New vbModule
                Set m.Component = currentComponent
                m.GUID = getGUID ' Class GUID
                m.interfaceGUID = getGUID ' Interface GUID
                m.interfaceName = currentComponent.Name
                m.Name = currentComponent.Name    ' need to obsolete Name
                project.prjModules.Add m, UCase(currentComponent.Name)
                Dim member As member
                For Each member In currentComponent.CodeModule.Members
                    Print #99, currentComponent.Name & "_" & member.Name & " Bindable=" & member.Bindable & " Browsable=" & member.Browsable & " Category=" & member.Category & " DefaultBind=" & member.DefaultBind & " DisplayBind=" & member.DisplayBind & " Scope=" & member.Scope & " StandardMethod=" & Hex(member.StandardMethod) & " Type=" & member.Type & " UIDefault=" & member.UIDefault
                    On Error Resume Next
                    Print #99, "Get: pbl="; currentComponent.CodeModule.ProcBodyLine(member.Name, vbext_pk_Get); " pcl="; currentComponent.CodeModule.ProcCountLines(member.Name, vbext_pk_Get); " psl="; currentComponent.CodeModule.ProcStartLine(member.Name, vbext_pk_Get)
                    If currentComponent.CodeModule.ProcStartLine(member.Name, vbext_pk_Get) > currentComponent.CodeModule.ProcBodyLine(member.Name, vbext_pk_Get) Then Err.Raise 1 ' VB error
                    Print #99, "Let: pbl="; currentComponent.CodeModule.ProcBodyLine(member.Name, vbext_pk_Let); " pcl="; currentComponent.CodeModule.ProcCountLines(member.Name, vbext_pk_Let); " psl="; currentComponent.CodeModule.ProcStartLine(member.Name, vbext_pk_Let)
                    If currentComponent.CodeModule.ProcStartLine(member.Name, vbext_pk_Let) > currentComponent.CodeModule.ProcBodyLine(member.Name, vbext_pk_Let) Then Err.Raise 1 ' VB error
                    Print #99, "Set: pbl="; currentComponent.CodeModule.ProcBodyLine(member.Name, vbext_pk_Set); " pcl="; currentComponent.CodeModule.ProcCountLines(member.Name, vbext_pk_Set); " psl="; currentComponent.CodeModule.ProcStartLine(member.Name, vbext_pk_Set)
                    If currentComponent.CodeModule.ProcStartLine(member.Name, vbext_pk_Set) > currentComponent.CodeModule.ProcBodyLine(member.Name, vbext_pk_Set) Then Err.Raise 1 ' VB error
                    Print #99, "Proc: pbl="; currentComponent.CodeModule.ProcBodyLine(member.Name, vbext_pk_Proc); " pcl="; currentComponent.CodeModule.ProcCountLines(member.Name, vbext_pk_Proc); " psl="; currentComponent.CodeModule.ProcStartLine(member.Name, vbext_pk_Proc)
                    On Error GoTo 0
                Next
                CreateLines m
            End If
        End If
    Next
#If 0 Then ' for debugging info only - won't work because Project.typelib may not be defined
    Dim ccTLI As tli.CoClassInfo
    For Each ccTLI In project.TypeLib.CoClasses
        Print #99, ccTLI.Name
    Next
    For Each currentComponent In p.VBComponents
        Print #99, "Component name="; currentComponent.Name; " type="; currentComponent.Type
        If currentComponent.Type <> vbext_ct_StdModule And IsModule(currentComponent) Then
            Set ccTLI = Nothing
            On Error Resume Next
            Set ccTLI = project.TypeLib.CoClasses.NamedItem(currentComponent.Name)
            On Error GoTo 0
            If Not ccTLI Is Nothing Then
                Print #99, "Component name="; currentComponent.Name
                Dim mi As MemberInfo
                For Each mi In ccTLI.DefaultInterface.Members
                    Print #99, "  mi="; mi.Name; " id="; Hex(mi.MemberId)
                Next
            End If
        End If
    Next
#End If
Dim tokens As Collection
' Process Public declarations: Consts, Enums, Types in all modules - must process here because of "Public v(MyConst)"
PassNumber = 1
#If 0 Then
    For Each currentComponent In p.VBComponents
        Print #99, "Pass1: Process Public Consts: "; currentComponent.Name
        clsAddIn.ProgressBarInit "Pass1: Processing " & currentComponent.Name
        If IsModule(currentComponent) Then
            If Not currentComponent.CodeModule Is Nothing Then
                Print #99, "Pass1: CodeModule: "; currentComponent.Name
                Set m = project.prjModules.Item(UCase(currentComponent.Name))
                Print #99, "Consts: m.name="; m.Name; " CountOfDecl="; currentComponent.CodeModule.CountOfDeclarationLines
                Set currentModule = m
' only need to do once
' fixme: create collection of createlines, put in vbModule
                dcllines = CreateLines(currentComponent.CodeModule, currentComponent.CodeModule.CountOfDeclarationLines, currentComponent.CodeModule.CountOfDeclarationLines)
                ' Process consts in all modules - csdtpckr\csdtpckr.vbp
                Print #99, "Pass1: cm.c="; currentComponent.CodeModule.Members.Count; " dcllines="; m.dcllines; " cdl="; currentComponent.CodeModule.CountOfDeclarationLines; " col="; currentComponent.CodeModule.CountOfLines
                ReDim orderedMembers(currentComponent.CodeModule.CountOfLines) As Collection
                ' fixme: Why aren't Consts in easyvi1a\vbmemcap.prj in source code order?
                ' fixme: Perhaps orderMembers is a good way of cycling through other For Eaches?
                ' Must use collection for each line number, could be multiple members per line.
                ' fixme: Alas, this fails if multiple Publics on same line t
                For Each member In currentComponent.CodeModule.Members
                    Print #99, "m="; member.Name; " cl="; member.CodeLocation
                    If orderedMembers(member.CodeLocation) Is Nothing Then Set orderedMembers(member.CodeLocation) = New Collection
                    orderedMembers(member.CodeLocation).Add member
                Next
                Dim tempVariant As Variant ' Ugh, must use Variant in For Each
                For Each tempVariant In orderedMembers
                    If Not tempVariant Is Nothing Then
                        Print #99, "tempVariant="; tempVariant.Count
                        For Each member In tempVariant
                            Print #99, "Pass1: Process Component Member: " & member.Name & " mt=" & member.Type & " codeloc=" & member.CodeLocation & " sm="; Hex(member.StandardMethod); " scope="; member.Scope
                            clsAddIn.ProgressBarInit "Pass1: Processing " & currentComponent.Name & "." & member.Name
                            If member.Scope = vbext_Friend Or member.Scope = vbext_Public Then
                                Select Case member.Type
                                    Case vbext_mt_Const
'                                        currentLineNumber = member.CodeLocation
'                                        Set tokens = getTokenizedLine(cmlines(member.CodeLocation))
                                        GetNextTokenizedLine member.CodeLocation, tokens
                                        parseLine tokens
                                        If tokens.Count > 0 Then Err.Raise 1 ' Internal error - expecting EOL
'                                        If member.StandardMethod = 0 Then If m.DefaultMember = "" Then m.DefaultMember = member.name Else If m.DefaultMember <> member.name Then Err.Raise 1
                                        Set currentProc = Nothing
                                End Select
                            End If
                        Next
                    End If
                Next
            End If
        End If
    Next
#End If
' Process declaration for symbol definitions: Consts, Enums, Types in all modules
' fixme: probably should reuse OrderedMembers() array above for For Each
    For Each currentComponent In p.VBComponents
        Print #99, "Pass1: Processing Declarations: "; currentComponent.Name
        clsAddIn.ProgressBarInit "Pass1: Processing " & currentComponent.Name
        If IsModule(currentComponent) Then
            If Not currentComponent.CodeModule Is Nothing Then
' fixme: create collection of createlines, put in vbModule
                Print #99, "Pass1: Processing Component Declarations: "; currentComponent.Name
                Set m = project.prjModules.Item(UCase(currentComponent.Name))
                Print #99, "module: m.name="; m.Name; " CountOfDecl="; currentComponent.CodeModule.CountOfDeclarationLines
                Set currentModule = m
                CompileVB ' Process declarations only, not methods or properties
            End If
        End If
    Next
' Re-Process declaration for values: Consts, Enums, Types in all modules
' fixme: probably should reuse OrderedMembers() array above for For Each
    PassNumber = 2
    For Each currentComponent In p.VBComponents
        Print #99, "t="; currentComponent.Type
        If IsModule(currentComponent) Then
            Set m = project.prjModules.Item(UCase(currentComponent.Name))
            Dim c As vbConst
            For Each c In m.Consts
                Print #99, "const="; c.ConstName '; " v="; c.ConstRPN.Item(1).tokValue
            Next
            Dim e As vbEnum
            For Each e In m.Enums
                Print #99, "enum="; e.enumName; " em.c="; e.enumMembers.Count
                Dim em As vbEnumMember
                For Each em In e.enumMembers
                    Print #99, "  em="; em.enumMemberName
                Next
            Next
            Dim udt As vbType
            For Each udt In m.types
                Print #99, "udt="; udt.typeName; " tm.c="; udt.typeMembers.Count
                Dim t As vbVariable
                For Each t In udt.typeMembers
                    Print #99, "  tm="; t.varSymbol; " n="; t.varType.dtDataName
                Next
            Next
        End If
    Next
    For Each currentComponent In p.VBComponents
        Print #99, "Pass2: Processing Component Declarations: "; currentComponent.Name
        clsAddIn.ProgressBarInit "Pass2: Processing " & currentComponent.Name
        If IsModule(currentComponent) Then
            If Not currentComponent.CodeModule Is Nothing Then
' fixme: create collection of createlines, put in vbModule
                Print #99, "Pass2: Processing Component Declarations: "; currentComponent.Name
                Set m = project.prjModules.Item(UCase(currentComponent.Name))
                Print #99, "module: m.name="; m.Name; " CountOfDecl="; currentComponent.CodeModule.CountOfDeclarationLines
                Set currentModule = m
                CompileVB ' Process declarations only, not methods or properties
            End If
        End If
    Next
    For Each currentComponent In p.VBComponents
        If IsModule(currentComponent) Then
            Set m = project.prjModules.Item(UCase(currentComponent.Name))
            ResolveConstantExpressions m
        End If
    Next
    For Each currentComponent In p.VBComponents
        Print #99, "t="; currentComponent.Type
        If IsModule(currentComponent) Then
            Set m = project.prjModules.Item(UCase(currentComponent.Name))
    '        Dim c As vbVariable
            For Each c In m.Consts
                Print #99, "const="; c.ConstName; " v="; c.ConstValue; " dt="; c.ConstDataType.dtDataType; " cv.dt="; varType(c.ConstValue)
            Next
    '        Dim e As vbEnum
            For Each e In m.Enums
                Print #99, "enum="; e.enumName; " em.c="; e.enumMembers.Count
                For Each em In e.enumMembers
                    Print #99, "  em="; em.enumMemberName; " v="; em.enumMemberValue
                Next
            Next
    '        Dim udt As vbType
            For Each udt In m.types
                Print #99, "udt="; udt.typeName; " tm.c="; udt.typeMembers.Count
    '            Dim t As vbVariable
                For Each t In udt.typeMembers
                    Print #99, "  tm="; t.varSymbol; " n="; t.varType.dtDataName
                Next
            Next
        End If
    Next
    ' fixme: optimize by only processing Type statements
    PassNumber = 3
    For Each currentComponent In p.VBComponents
        Print #99, "Pass3: Processing User-defined Types: "; currentComponent.Name
        clsAddIn.ProgressBarInit "Pass3: User-defined Types " & currentComponent.Name
        If IsModule(currentComponent) Then
            If Not currentComponent.CodeModule Is Nothing Then
' fixme: create collection of createlines, put in vbModule
                Print #99, "Pass3: Processing User-defined Types: "; currentComponent.Name
                Set m = project.prjModules.Item(UCase(currentComponent.Name))
                Print #99, "module: m.name="; m.Name; " CountOfDecl="; currentComponent.CodeModule.CountOfDeclarationLines
                Set currentModule = m
                CompileVB ' Process declarations only, not methods or properties
            End If
        End If
    Next
    Dim md As vbModule
    For Each md In m.ModuleDependencies
        Print #99, "  md="; md.Name
    Next
#If 0 Then
    For Each currentComponent In p.VBComponents
        Print #99, "Pass3: Processing Private Component Declarations: "; currentComponent.Name
        clsAddIn.ProgressBarInit "Pass3: Processing " & currentComponent.Name
        If IsModule(currentComponent) Then
            If Not currentComponent.CodeModule Is Nothing Then
' fixme: create collection of createlines, put in vbModule
                Print #99, "Pass3: Processing Component Declarations: "; currentComponent.Name
                Set m = project.prjModules.Item(UCase(currentComponent.Name))
                Print #99, "module: m.name="; m.Name; " CountOfDecl="; currentComponent.CodeModule.CountOfDeclarationLines
                Set currentModule = m
                CompileVB ' Process declarations only, not methods or properties
            End If
        End If
    Next
#End If
    For Each m In currentProject.prjModules
        OrderModuleDependencies m, 0
    Next
    For Each m In currentProject.prjModulesOrdered
        Print #99, "om="; m.Name
    Next
    PassNumber = 4
' Process Methods, Properties
Print #99, "4"
    For Each currentComponent In p.VBComponents
        Print #99, "Pass4: Process Component Methods, Properties: " & currentComponent.Name
        clsAddIn.ProgressBarInit "Pass4: Processing " & currentComponent.Name
        If IsModule(currentComponent) Then
            If Not currentComponent.CodeModule Is Nothing Then
                Print #99, "Pass4: CodeModule: "; UCase(currentComponent.Name)
                Set m = project.prjModules.Item(UCase(currentComponent.Name))
                Print #99, "Components: m.name="; m.Name; " ct="; m.Component.Type; " CountOfDecl="; currentComponent.CodeModule.CountOfDeclarationLines
'                Dim dcllines As Long
                Set currentModule = m
' Alternative, perhaps better way, is to disable below and insert
'   another pass processing only "member.Type = non-method and non-property" (declares only)
#If 0 Then
                currentLineNumber = 1
                col = m.dcllines
                m_eof = False
                CompileVB ' Process declarations only, not methods or properties
#End If
                On Error Resume Next
                Dim prop As Property
                For Each prop In currentComponent.Properties
                    Print #99, "  p.n="; prop.Name;
                    Print #99, " p.vt="; varType(prop.Value);
                    Print #99, " p.ni="; prop.NumIndices;
                    Print #99, " p.tn="; typeName(prop.Value);
                    If Not IsObject(prop.Value) Then Print #99, " v="; prop.Value;
                    If IsNumeric(prop.Value) Then Print #99, "(0x"; Hex(prop.Value); ")";
                    Print #99,
                Next
                On Error GoTo 0
Print #99, "6"
'                m.OptionExplicit = True
                Select Case m.Component.Type
                Case vbext_ct_StdModule
                ' do nothing
                Case vbext_ct_ClassModule, vbext_ct_ActiveXDesigner
Print #99, "7"
                    m.clsMultiUse = -1 'True
                    m.clsPersistable = 0 'NotPersistable
                    m.clsDataBindingBehavior = 0 'vbNone
                    m.clsDataSourceBehavior = 0 'vbNone
                    m.clsMTSTransactionMode = 0 'NotAnMTSObject
                    m.clsVB_Name = m.Name    ' need to obsolete Name
                    m.clsVB_GlobalNameSpace = False
                    m.clsVB_PredeclaredID = False
                    Set prop = Nothing
                    On Error Resume Next ' this is needed
                    Set prop = currentComponent.Properties.Item("Instancing")
                    On Error GoTo 0
Print #99, "7a "; prop Is Nothing
                    If Not prop Is Nothing Then
Print #99, "7b "; varType(prop.Value)
' fixme: Creatable is hardly implemented for classes
                        m.clsVB_Creatable = currentComponent.Properties.Item("Instancing").Value > 2 ' Createable
Print #99, "7c"
' fixme: Exposed is hardly implemented for classes
                        m.clsVB_Exposed = currentComponent.Properties.Item("Instancing").Value = 5 ' MultiUse
Print #99, "7d"
                    End If
Print #99, "7e"
                Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_PropPage, vbext_ct_UserControl ', vbext_ct_ActiveXDesigner
Print #99, "8"
                    m.clsVB_Name = m.Name    ' need to obsolete Name?
Print #99, "8a"
                    m.clsVB_GlobalNameSpace = False
                    m.clsVB_Creatable = False
                    m.clsVB_PredeclaredID = True
                    m.clsVB_Exposed = False
' fixme: assign to specific type, put below in subroutine
'                    Dim f As VBForm ' probably needs different type for each form
                    Dim f As Object
Print #99, "8b"
                    On Error GoTo designer_error
                    If currentComponent.Designer Is Nothing Then
designer_error:
                        Print #99, "AddInMain: Error accessing form designer. Unsupported Designer form or UserControl or possibly corrupt VBT (try restarting VBT)."
                        MsgBox "Error accessing form designer. Unsupported Designer form or UserControl or possibly corrupt VBT (try restarting VBT)."
                        On Error GoTo 0
                        Err.Raise 1
                    End If
                    On Error GoTo 0
                    Set f = currentComponent.Designer
                    Dim vbc As VBControl
                    ' create form control variables (Text1, CommandButton3)
Print #99, "f.c="; Not f.VBControls Is Nothing
                    For Each vbc In f.VBControls
                        Print #99, "prop n="; vbc.Properties.Item("name"); " i="; vbc.Properties.Item("index"); " ti="; varType(vbc.Properties.Item("index")); " cn="; vbc.ClassName; " id="; vbc.ProgId
                        On Error Resume Next
                        For Each prop In vbc.Properties
                            Print #99, "  pn="; prop.Name; " ni="; prop.NumIndices; ; " vt="; varType(prop.Value); " v="; prop.Value;
                            If IsNumeric(prop.Value) Then Print #99, "(0x"; Hex(prop.Value); ")";
                            Print #99,
                        Next
                        Dim ci As CoClassInfo
                        Set ci = Nothing
                        Set ci = ClassInfoFromObject(vbc.ControlObject)
                        If ci Is Nothing Then Print #99, "No CoClassInfo"
                        Print #99, "ci="; TypeInfoToVBType(ci); " p.cf="; ci.Parent.ContainingFile
                        Dim ii As InterfaceInfo
                        Set ii = Nothing
                        Set ii = InterfaceInfoFromObject(vbc.ControlObject)
                        If ii Is Nothing Then Print #99, "No InterfaceInfo"
                        Print #99, "ii="; TypeInfoToVBType(ii); " am="; Hex(ii.AttributeMask)
                        Print #99, " p.cf="; ii.Parent.ContainingFile; " default member="; ii.GetMember(0).Name;
                        Print #99,
                        Dim tlib As TypeLibInfo
                        Set tlib = Nothing
                        If ci Is Nothing Then
                            Set tlib = TypeLibInfoFromFile(tliVB.FullPath)
                        Else
                            Set tlib = TypeLibInfoFromFile(ci.Parent.ContainingFile)
                        End If
                        If tlib Is Nothing Then
                            Print #99, "No TLib"
                        Else
                            Print #99, "tlib="; tlib.Name; " GUID="; tlib.GUID
                            Dim TLibControls As New Collection
                            If TLibControls.Item(tlib.GUID) Is Nothing Then
                                TLibControls.Add tlib, tlib.GUID
                                For Each ci In tlib.CoClasses
                                    Print #99, "  ci.n="; tlib.Name; "_"; ci.Name; "_"; ci.DefaultInterface.Name
                                Next
                                For Each ii In tlib.Interfaces
                                    Print #99, "  ii.n="; tlib.Name; "_"; ii.Name; " am="; Hex(ii.AttributeMask)
                                    Dim mi As MemberInfo
                                    For Each mi In ii.Members
                                        Print #99, "    mi.n="; mi.Name; " ik="; mi.InvokeKind; " mid="; Hex(mi.MemberId); " pc="; mi.parameters.Count; " rt="; mi.ReturnType Is Nothing; " am="; Hex(mi.AttributeMask)
                                        Dim pi As ParameterInfo
                                        For Each pi In mi.parameters
                                            Print #99, "      pi.n="; pi.Name; " vt="; pi.VarTypeInfo.varType; " f="; Hex(pi.flags)
                                        Next
                                    Next
                                    If ii.VTableInterface Is Nothing Then
                                        Print #99, "  No VTable for "; tlib.Name; "_"; ii.Name
                                    Else
                                        Print #99, "  iiv.n="; tlib.Name; "_"; ii.Name; ; " am="; Hex(ii.AttributeMask)
                                        For Each mi In ii.VTableInterface.Members
                                            Print #99, "    mi.n="; mi.Name; " ik="; mi.InvokeKind; " mid="; Hex(mi.MemberId); " pc="; mi.parameters.Count; " rt="; mi.ReturnType Is Nothing; " am="; Hex(mi.AttributeMask)
                                            For Each pi In mi.parameters
                                                Print #99, "      pi.n="; pi.Name; " vt="; pi.VarTypeInfo.varType
                                            Next
                                        Next
                                    End If
                                Next
                            End If
                        End If
                        On Error GoTo 0
                    Next
Print #99, "9"
                    For Each vbc In f.VBControls
Print #99, "10 cn="; vbc.ClassName
                        Dim l As Long
                        l = vbc.Properties.Item("index") ' -1 for non-control array
                        Dim s As String
                        s = vbc.Properties.Item("name")
                        Print #99, "prop name="; s; " l="; l
Print #99, "11"
                        On Error Resume Next
                        Set ci = Nothing
                        Set ci = ClassInfoFromObject(vbc.ControlObject)
Print #99, "12"
                        Set tlib = Nothing
                        If ci Is Nothing Then
                            Set tlib = TypeLibInfoFromFile(tliVB.FullPath)
                            Set ci = tlib.CoClasses.NamedItem(vbc.ClassName)
                        Else
                            Set tlib = TypeLibInfoFromFile(ci.Parent.ContainingFile)
                        End If
                        If tlib Is Nothing Then Err.Raise 1
                        Set ii = Nothing
                        Set ii = InterfaceInfoFromObject(vbc.ControlObject)
                        On Error GoTo 0
Print #99, "13 ci="; ci Is Nothing
                        If ci Is Nothing Then Err.Raise 1
                        If ii Is Nothing Then
                            Print #99, "No InterfaceInfo"
'                            Set ii = ci.DefaultInterface
'Print #99, "14 ci.di.ii.n="; ii.Name; " ii.am="; Hex(ii.AttributeMask); " ii.GUID="; ii.GUID; " ci.di.GUID="; ci.DefaultInterface.GUID
                        Else
' Note: AttributeMask is correct - has FDUAL - but ii.name is hidden (_^A)
'   if ci.DefaultInterface is used, AttributeMask is 1 and ii.name is OK - can't win?
'   (.MemberId Mod 65536) stuff is only needed if this ii isn't over written by ci.DefaultInterface
'   directly effects GetDefaultMemberInfo()
'                            Print #99, "14 ii.n="; ii.Name; " ii.am="; Hex(ii.AttributeMask); " ii.GUID="; ii.GUID; " ci.di.GUID="; ci.DefaultInterface.GUID
'                            If ii.GUID <> ci.DefaultInterface.GUID Then Err.Raise 1
                        End If
                        Set ii = ci.DefaultInterface
Print #99, "14 ii.n="; ii.Name; " ii.am="; Hex(ii.AttributeMask); " ii.GUID="; ii.GUID; " ci.di.GUID="; ci.DefaultInterface.GUID
'Print #99, "ii.p.n="; ii.Parent.Name
'Print #99, "ii.m.c="; ii.Members.count
'Set ii = tlib.Interfaces.NamedItem(ii.Name)
'Print #99, "14 tlib.ii.n="; ii.Name; " ii.am="; Hex(ii.AttributeMask); " ii.GUID="; ii.GUID; " ci.di.GUID="; ci.DefaultInterface.GUID
'Set ci = tlib.CoClasses.NamedItem(ci.Name)
'Set ii = ci.DefaultInterface
'Print #99, "14 tlib.ci.ii.n="; ii.Name; " ii.am="; Hex(ii.AttributeMask); " ii.GUID="; ii.GUID; " ci.di.GUID="; ci.DefaultInterface.GUID
' kludge: need to properly detect property vs. method
'                        If InStr(1, s, "_") <> 0 Then MsgBox "Assuming form property (control): " & s

'                        If 1 Then
                            ' form property
Print #99, "15"
                            Dim pt As proctable
                            On Error Resume Next
                            Set pt = currentModule.procs.Item(SymIK(s, INVOKE_FUNC Or INVOKE_PROPERTYGET))
                            ' fixme: rework "variable" style controls
                            If pt Is Nothing Then Set pt = currentModule.procs.Item(SymIK(s, INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF))
                            On Error GoTo 0
                            Print #99, " pt="; Not pt Is Nothing; " l="; l
                            If pt Is Nothing Then
' kludge: need to properly handle Form controls, not put them in procs collection
' controls, such as label, are similar to variables GET/PUT/PUTREF
                                Set pt = ProcAdd(s, proc_attr_defined Or PROC_ATTR_VARIABLE Or PROC_ATTR_PUBLIC Or PROC_ATTR_FORM_CONTROL, vbext_mt_Variable, INVOKE_PROPERTYGET Or INVOKE_PROPERTYPUT Or INVOKE_PROPERTYPUTREF)
                                Dim v As vbVariable
                                Set v = CreateVariable(s, proc_attr_defined Or PROC_ATTR_PUBLIC)
Print #99, "21"
                                Set pt.procFunctionResultType = v
Print #99, "22"
                                Dim mm As vbModule
                                On Error Resume Next
                                Set mm = Nothing
                                Set mm = currentProject.prjModules.Item(UCase(vbc.ClassName))
                                On Error GoTo 0
                                ' new - Only care about control class match if designer or user control
                                ' ListView2
                                If Not mm Is Nothing Then If mm.Component.Type <> vbext_ct_ActiveXDesigner And mm.Component.Type <> vbext_ct_UserControl Then Set mm = Nothing
                                If mm Is Nothing Then
                                    v.varType.dtType = tokReferenceClass
                                Else
                                    v.varType.dtType = tokFormClass
                                    Set v.varType.dtClass = mm
                                End If
                                ' new - want these for both cases above?
                                Set v.varType.dtClassInfo = ci
                                Set v.varType.dtInterfaceInfo = ii
Print #99, "23"
                                If l <> -1 Then
                                    Print #99, "control array: n="; pt.procName; " b="; l
#If 1 Then
                                    v.varAttributes = v.varAttributes Or VARIABLE_CONTROLARRAY
'#Else
                                    pt.procParams.Add New paramTable
                                    Set pt.procParams.Item(1).paramVariable = New vbVariable
                                    pt.procParams.Item(1).paramVariable.varSymbol = "_Index"
                                    Set pt.procParams.Item(1).paramVariable.varType = New vbDataType
                                    pt.procParams.Item(1).paramVariable.varType.dtDataType = vbLong
#End If
                                End If
                                v.varType.dtDataType = vbObject
Print #99, "25"
                            Else
Print #99, "26"
                                ' fixme: can local variable outscope form label? - check it out
                                If l = -1 Then Err.Raise 1 ' duplicate control name
                            End If
Print #99, "27"
#If 0 Then
                        Else
Print #99, "28"
                            ' form method
                            On Error Resume Next
                            Set pt = currentModule.procs.Item(SymIK(s, INVOKE_FUNC Or INVOKE_PROPERTYGET))
                            On Error GoTo 0
                            Print #99, "prop name="; s; " pt="; pt Is Nothing; " l="; l
                            If pt Is Nothing Then
                                Set pt = ProcAdd(s, proc_attr_defined Or PROC_ATTR_PUBLIC Or PROC_ATTR_FORM_CONTROL, vbext_mt_Method, INVOKE_FUNC)
                            Else
                                ' fixme: can local variable outscope form label - check it out
                                If l = -1 Then Err.Raise 1 ' duplicate control name
                            End If
Print #99, "29"
                        End If
#End If
Print #99, "e"
                        Set pt = Nothing
                    Next
                Case Else
                    Print #99, "Unknown component type: "; m.Component.Type
                    Err.Raise 1
                End Select
                Print #99, "Pass4: Processing Methods and Properties: m="; currentComponent.CodeModule.Members.Count
                Print #99, "m="; m.Name; " m.def="; m.DefaultMember
                ProcessDeclares = True
                On Error Resume Next
                For Each member In currentComponent.CodeModule.Members
                    Print #99, "Name="; member.Name
                    Print #99, "Type="; member.Type
                    Print #99, "CodeLocation="; member.CodeLocation
                    Print #99, "Bindable="; member.Bindable
                    Print #99, "Browsable="; member.Browsable
                    Print #99, "Category="; member.Category
                    Print #99, "Collection="; member.Collection.Count
                    Print #99, "DefaultBind="; member.DefaultBind
                    Print #99, "Description="; member.Description ' generates an error in BTree.vbp Why?
                    Print #99, "DisplayBind="; member.DisplayBind
                    Print #99, "HelpContextID="; member.HelpContextID
                    Print #99, "Hidden="; member.Hidden
                    Print #99, "Scope="; member.Scope
                    Print #99, "StandardMethod="; Hex(member.StandardMethod)
                    Print #99, "Static="; member.Static
                    Print #99, "Type="; member.Type
                    Print #99, "UIDefault="; member.UIDefault
                Next
                On Error GoTo 0
' bug in VB6 causes dup entries in currentComponent.CodeModule.Members collection for Event Sub1 + Sub Sub1
                For Each member In currentComponent.CodeModule.Members
                    ' Process Function or Sub declaration line
                    Print #99, "Pass4: Processing " & member.Name & " m.def="; m.DefaultMember & " mt=" & member.Type & " codeloc=" & member.CodeLocation & " sm="; Hex(member.StandardMethod)
Print #99, "195"
                    clsAddIn.ProgressBarInit "Pass4: Processing " & currentComponent.Name & "." & member.Name
' skip re-processing of Declares and Consts
'                    If member.CodeLocation > dcllines Then
                    ' CodeLocation = 0 can occur when multiple defintions appear on a line
Print #99, "196 cm="; currentModule.Name
                    If member.CodeLocation > 0 Then
Print #99, "197"
                    Select Case member.Type
                        Case vbext_mt_Event, vbext_mt_Method, vbext_mt_Variable
Print #99, "198"
'                            Dim tokens As New Collection
'                            currentLineNumber = member.CodeLocation
'                            Set tokens = getTokenizedLine(member.CodeLocation)
                            GetNextTokenizedLine member.CodeLocation, tokens
                            parseLine tokens
Print #99, "198a"
                            If tokens.Count > 0 Then Err.Raise 1 ' Internal error - expecting EOL
Print #99, "198b"
                            If member.StandardMethod = 0 Then If m.DefaultMember = "" Then m.DefaultMember = member.Name Else If m.DefaultMember <> member.Name Then Err.Raise 1
                            Set currentProc = Nothing
Print #99, "199"
                        Case vbext_mt_Property
Print #99, "200a"
                            Dim i As Integer
                            For i = 1 To 3 ' process all 3 properties types (Get,Put,Set)
                                currentLineNumber = 0
                                ' Need On Error when Get, Let or Set is undeclared. textob\textobj.vbp
                                On Error Resume Next
                                Print #99, "200 i="; i
                                Print #99, "200b: pbl="; currentComponent.CodeModule.ProcBodyLine(member.Name, i); " psl="; currentComponent.CodeModule.ProcStartLine(member.Name, i); " pcl="; currentComponent.CodeModule.ProcCountLines(member.Name, i)
''''                                    ' ProcStartLine is used when enumerating lines, do not use ProcBodyLine (which is used to get declaration)
                                ' No, ProcBodyLine
' MultiColSort\MultiColSort.vbp; projectirc...\projectirc.vbp
                                currentLineNumber = currentComponent.CodeModule.ProcStartLine(member.Name, i)
                                On Error GoTo 0
Print #99, "201"
                                If currentLineNumber > 0 Then
'                                        GetNextTokenizedLine currentLineNumber, tokens
                                    ' ProcStartLine may be off by -1, so try twice. SetDefault1.vbp, EventCtl.vbp
'                                    If tokens.Count = 0 Then Set tokens = getTokenizedLine(currentLineNumber + 1)
                                    Do
                                        Set tokens = getTokenizedLine(currentLineNumber)
                                        currentLineNumber = currentLineNumber + 1
                                    Loop While tokens.Count = 0
                                    parseLine tokens
Print #99, "cp="; Not currentProc Is Nothing
                                    If currentProc Is Nothing Then Err.Raise 1
Print #99, "202"
                                    If tokens.Count > 0 Then Err.Raise 1
Print #99, "203"
                                    Set currentProc = Nothing
                                    If member.StandardMethod = 0 Then If m.DefaultMember = "" Then m.DefaultMember = member.Name Else If m.DefaultMember <> member.Name Then Err.Raise 1
                                End If
Print #99, "204"
                            Next
Print #99, "205"
                    End Select
                    End If
Print #99, "206"
                Next
Print #99, "207"
                ProcessDeclares = False
            End If
        End If
    Next
    PassNumber = 5
    For Each currentComponent In p.VBComponents '.VBE.Windows.Item(1).Collection
        Print #99, "Pass5: Process Component: " & currentComponent.Name & " t=" & currentComponent.Type
        ' note: cddbco2a\project1.vbp is a good test program for here (uses single line method defs)
        clsAddIn.ProgressBarInit "Pass5: Processing " & currentComponent.Name
        If IsModule(currentComponent) Then
            If Not currentComponent.CodeModule Is Nothing Then
                Set m = project.prjModules.Item(UCase(currentComponent.Name))
                Set currentModule = m
'                m.OptionExplicit = True
' createlines should only be done once for each CodeModule - currently done twice - fixme
                Select Case currentComponent.Type
                    Case vbext_ct_StdModule
                    ' fixme: don't know how to handle forms? Do they have GUIDs?
                    Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm ', vbext_ct_ActiveXDesigner , vbext_ct_UserControl
                    Case Else
                        Dim ccTLI As tli.CoClassInfo
                        On Error Resume Next
                        ' fixme: create m.TypeLib - but what about non-public, non-typelib classes?
                        Set ccTLI = project.TypeLib.CoClasses.NamedItem(currentComponent.Name)
                        On Error GoTo 0
                        If Not ccTLI Is Nothing Then
                            m.interfaceName = ccTLI.DefaultInterface.Name
                            m.interfaceGUID = ccTLI.DefaultInterface.GUID
                            m.GUID = ccTLI.GUID
                            If Not ccTLI.DefaultEventInterface Is Nothing Then
                                m.EventName = ccTLI.DefaultEventInterface.Name
                                m.EventGUID = ccTLI.DefaultEventInterface.GUID
                            End If
                            m.clsVersionMajor = ccTLI.MajorVersion
                            m.clsVersionMinor = ccTLI.MinorVersion
Print #99, "  TLI.n="; ccTLI.Name; " di.n="; m.interfaceName; " di.g="; m.interfaceGUID
                            Set ccTLI = Nothing
                        Else ' Is there really an interface and GUID for non-public classes? If not, change emitters.
                            m.interfaceName = m.Name
                            m.interfaceGUID = getGUID
                            m.GUID = getGUID
                            If m.events.Count > 0 Then
                                m.EventName = "_" & m.Name
                                m.EventGUID = getGUID
                            End If
                            m.clsVersionMajor = 1
                            m.clsVersionMinor = 0
Print #99, "  non-TLI: di.n="; m.interfaceName; " di.g="; m.interfaceGUID; " di.en="; m.EventName; " di.eg="; m.EventGUID
                        End If
                        ' init predefined variables
#If 0 Then ' should Me be a variable or special test?
                        Dim v As vbVariable
                        Set v = New vbVariable
                        v.MemberType = vbext_mt_Variable
                        v.varSymbol = "ME"
                        v.varAttributes = VARIABLE_ME
                        Set v.varModule = m
                        Set v.varProc = Nothing
                        Set v.varType = New vbDataType
                        Set v.varType.dtClass = m
                        v.varType.dtDataType = vbObject
''''                        v.varType.dtType = tokProjectClass
                        v.varType.dtType = tokFormClass
                        m.ModuleVars.Add v, "ME"
                        Print #99, "m="; m.Name; m.ModuleVars.Count
#End If
                End Select
        
'                Set currentModule.procs = Nothing ' clear procs, about to reparse
'                Set currentModule.events = Nothing
                For Each member In currentComponent.CodeModule.Members
                    clsAddIn.ProgressBarInit "Pass5: Processing " & currentComponent.Name & "." & member.Name
                    clsAddIn.ProgressBarUpdate MemberCount / TotalMembers
                    MemberCount = MemberCount + 1
                    Print #99, "Pass5: Process Member: " & member.Name & " codeloc=" & member.CodeLocation & " type=" & member.Type
                    ' CodeLocation = 0 can occur when multiple defintions appear on a line
                    If member.CodeLocation > 0 Then
                        Select Case member.Type
                            Case vbext_mt_Event, vbext_mt_Method
    ' skip re-processing of Declares (test file is RESOURCE.VBP)
    ' wouldn't have to skip if Declares were added to procs collection
    Print #99, "150 dcl="; m.dcllines
                                If member.CodeLocation > m.dcllines Then
    Print #99, "151 pnik="; UCase(member.Name) & (INVOKE_FUNC Or INVOKE_PROPERTYGET)
    '                                Set currentProc = currentModule.procs.Item(symik(member.Name,INVOKE_FUNC Or INVOKE_PROPERTYGET))
    '                                pbl = member.CodeLocation
    '                                pbl = currentComponent.CodeModule.ProcBodyLine(member.Name, i)
    '                                Dim col As Long
    '                                col = pbl + currentComponent.CodeModule.ProcCountLines(member.Name, vbext_pk_Proc) - 1
    '                                For currentLineNumber = pbl To col
                                    currentLineNumber = member.CodeLocation
                                    Do
                                    
    Print #99, "152"
    '                                    Set tokens = getTokenizedLine(currentLineNumber)
    '                                    currentLineNumber = currentLineNumber + 1
                                        currentLineNumber = GetNextTokenizedLine(currentLineNumber, tokens)
    Print #99, "153"
                                        parseLine tokens
    Print #99, "153a"
                                        If tokens.Count > 0 Then Err.Raise 1
    Print #99, "153b"
    '                                Next
' fixme: better to loop until currentlinenumber < col -- like in Property below?
                                    Loop Until currentProc Is Nothing
    Print #99, "153c"
    '                                If Not currentProc Is Nothing Then Err.Raise 1
    Print #99, "154"
                                End If
                            Case vbext_mt_Property
    '' skip re-processing of Declares (test file is RESOURCE.VBP)
    '' wouldn't have to skip if Declares were added to procs collection
    '                            If member.CodeLocation > m.dcllines Then
                                For i = 1 To 3
'                                    Dim pbl As Long
                                    currentLineNumber = 0
                                    ' Need On Error when Get, Let or Set is undeclared. textob\textobj.vbp
                                    On Error Resume Next
                                    Print #99, "300a: i="; i
                                    Print #99, "300b: pbl="; currentComponent.CodeModule.ProcBodyLine(member.Name, i); " psl="; currentComponent.CodeModule.ProcStartLine(member.Name, i); " pcl="; currentComponent.CodeModule.ProcCountLines(member.Name, i)
''''                                    ' ProcStartLine is used when enumerating lines, do not use ProcBodyLine (which is used to get declaration)
                                    currentLineNumber = currentComponent.CodeModule.ProcStartLine(member.Name, i)
                                    On Error GoTo 0
                                    If currentLineNumber > 0 Then
    '                                    Set currentProc = currentModule.procs.Item(symik(member.Name, Choose(i, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF, INVOKE_FUNC Or INVOKE_PROPERTYGET))
                                        Dim col As Long
                                        col = currentLineNumber + currentComponent.CodeModule.ProcCountLines(member.Name, i)
Print #99, "300"
                                        Do
                                            Print #99, "prop: cln="; currentLineNumber; " col="; col; " cp="; Not currentProc Is Nothing
Print #99, "301"
'                                            currentLineNumber = GetNextTokenizedLine(currentLineNumber, tokens)
Print #99, "302"
                                            Set tokens = getTokenizedLine(currentLineNumber)
                                            currentLineNumber = currentLineNumber + 1
Print #99, "303"
                                            parseLine tokens
                                            If tokens.Count > 0 Then Err.Raise 1
Print #99, "304"
                                        Loop While currentLineNumber < col
Print #99, "305"
                                        If Not currentProc Is Nothing Then Err.Raise 1
Print #99, "306"
                                    End If
                                Next
    '                            End If
                         End Select
                     End If
Print #99, "155"
                Next
Print #99, "156"
                If Not currentProc Is Nothing Then Err.Raise 1
Print #99, "157"
            End If ' Not currentComponent.CodeModule Is Nothing
        End If ' ismodule(currentcomponent)
Print #99, "158"
    Next
    Set tliVB = Nothing
'Next ' Project

Close #99

Set AddInMain = group

End Function

Sub OrderModuleDependencies(ByVal m As vbModule, ByVal c As Long)
Print #99, "OrderModuleDependencies: m="; m.Name; " c="; c; " md.c="; m.ModuleDependencies.Count
If c > currentProject.prjModules.Count Then Err.Raise 1 ' Circular definitions
Dim mm As vbModule
On Error Resume Next
Set mm = currentProject.prjModulesOrdered.Item(m.Name)
On Error GoTo 0
If mm Is Nothing Then
    For Each mm In m.ModuleDependencies
        OrderModuleDependencies mm, c + 1 ' recursive
    Next
    currentProject.prjModulesOrdered.Add m, m.Name
End If
End Sub

Sub CreateLines(ByVal m As vbModule)
Print #99, "CreateLines: m="; m.Name
Print #99, "cm.n="; m.Component.Name
Dim cm As CodeModule
Set cm = m.Component.CodeModule
Dim CountOfLines As Long
Dim CountOfDeclarationLines As Long
CountOfDeclarationLines = cm.CountOfDeclarationLines
CountOfLines = cm.CountOfLines
Print #99, "col="; CountOfLines; " codl="; CountOfDeclarationLines
currentLineNumber = 0
Dim i As Long
Dim j As Long
Dim s As String
' CountOfLines can be zero
While currentLineNumber < CountOfLines
    i = i + 1
    currentLineNumber = currentLineNumber + 1
    s = cm.lines(currentLineNumber, 1)
    j = currentLineNumber
    While Right(RTrim(s), 2) = " _"
        j = j + 1
        s = Mid(s, 1, Len(RTrim(s)) - 1) & cm.lines(j, 1)
    Wend
    Dim tokens As Collection
    Set tokens = oTok.TokenizeLine(s)
' Process conditional compilation
    If tokens.Count = 0 Then
        ' do nothing
    ElseIf tokens.Item(1).tokString = "#" Then
        Dim CompilerDirective As vbStmt
        tokens.Remove 1 ' remove #
        Set CompilerDirective = Nothing
        On Error Resume Next
        Set CompilerDirective = cCompilerDirectives.Item(UCase(tokens.Item(1).tokString))
        On Error GoTo 0
        If CompilerDirective Is Nothing Then Err.Raise 1 ' Unrecognized compiler directive
        CompilerDirective.stmtParse tokens
        If tokens.Count > 0 Then Err.Raise 1
        If CompilerIfStack.Count <> 0 Then Print #99, "#if="; CompilerIfStack.Item(CompilerIfStack.Count).CompilerIfExpressionValue
    Else
        If CompilerIfStack.Count = 0 Then GoTo 10
        If Not CompilerIfStack.Item(CompilerIfStack.Count).CompilerIfExpressionValue Then
            Set tokens = New Collection
        Else
10          m.ilines.Add i, CStr(currentLineNumber) ' Key must be a string - CStr must be used
            m.lines.Add s, CStr(i)
        End If
    End If
    Print #99, "i="; i; " CLN="; currentLineNumber; " Line = "; s; " t.c="; m.ilines.Count
    currentLineNumber = j
    If currentLineNumber = CountOfDeclarationLines Then m.dcllines = i
Wend
Print #99, "CreateLines: done: dcllines="; m.dcllines; " m.ts.c="; m.ilines.Count
For i = 1 To CountOfLines
    On Error Resume Next
    s = ""
    s = m.lines.Item(CStr(i))
    Print #99, i; "l="; s
    j = 0
    j = m.ilines.Item(CStr(i))
    s = ""
    s = m.lines.Item(j)
    On Error GoTo 0
    Print #99, i; j; "c="; s
Next
End Sub

Function ResolveConstantExpressions(ByVal m As vbModule) As Long
Print #99, "ResolveConstantExpressions: m="; m.Name; " c.c="; m.Consts.Count; " e.n="; m.Enums.Count
Dim c As vbConst
For Each c In m.Consts
    Print #99, "c="; c.ConstName; " e="; IsEmpty(c.ConstValue); " rpn.c="; c.ConstRPN.Count; " dt="; c.ConstDataType.dtDataType
    If IsEmpty(c.ConstValue) Then
        c.ConstValue = oRPN.EvalConstRPNStack(c.ConstRPN, c.ConstDataType.dtDataType, c.ConstDataType.dtLength)
        If c.ConstDataType.dtDataType = vbVariant Then
            If Not IsEmpty(c.ConstValue) And Not IsNull(c.ConstValue) Then
                Set c.ConstDataType = New vbDataType
                c.ConstDataType.dtDataType = varType(c.ConstValue)
            End If
        End If
    End If
Next
Dim e As vbEnum
For Each e In m.Enums
    Dim em As vbEnumMember
    Dim emval As Long
    emval = 0
    For Each em In e.enumMembers
        Print #99, "em="; em.enumMemberName; " e="; IsEmpty(em.enumMemberValue); " rpn.c="; em.enumMemberRPN.Count
        If IsEmpty(em.enumMemberValue) Then
            If em.enumMemberRPN.Count = 0 Then
                em.enumMemberValue = emval
            Else
                em.enumMemberValue = oRPN.EvalConstRPNStack(em.enumMemberRPN, vbLong)
            End If
        End If
        emval = em.enumMemberValue + 1
    Next
Next
Print #99, "ResolveConstantExpressions: done"
End Function

Function IsModule(ByVal c As VBComponent) As Boolean
' using negative
If Not IsAny(c.Type, vbext_ct_RelatedDocument, vbext_ct_ResFile) Then IsModule = True
End Function

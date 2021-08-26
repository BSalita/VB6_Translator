VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbIDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Emiter

Sub Emiter_Emiter(group As vbGroup)
Dim p As vbPrj
Dim m As vbModule
Dim ProjectCount As Single
Dim procNumber As Long

Open group.OutputPath & "vbIDL.lst" For Output As 99
For Each p In group.Projects
    Print #99, "Processing "; p.prjName; " mc="; p.prjModules.Count
    group.AddIn.ProgressBarUpdate ProjectCount / group.Projects.Count * 100
    ProjectCount = ProjectCount + 1
    Dim OutputPath As String
    OutputPath = VBTOutputPath(group, p, "c")
    If p.VBProject.Type = vbext_pt_StandardExe Then
        On Error Resume Next
        Kill OutputPath & p.prjName & ".idl"
        On Error GoTo 0
    Else
        Open OutputPath & p.prjName & ".idl" For Output As 1
        ' insert date/time stamp
        Print #1, "// Generated .IDL file (by VBT)"
        Print #1, "//"
        Print #1, "// typelib filename: " & p.prjName & ".tlb"
        Print #1,
        Print #1, "["
        Print #1, "  uuid(" & Mid(p.TypeLib.GUID, 2, Len(p.TypeLib.GUID) - 2) & "),"
        ' hard coding version
        Print #1, "  version(" & p.TypeLib.MajorVersion & "." & p.TypeLib.MinorVersion & "),"
        ' hard coding GUID
        Print #1, "  custom(50867B00-BB69-11D0-A8FF-00A0C9110059, 8169)"
        Print #1, "]"
        Print #1, "library " & p.prjName
        Print #1, "{"
        ' hard coding GUID
        Print #1, "    // TLib :      // TLib : OLE Automation : {00020430-0000-0000-C000-000000000046}"
        Print #1, "    importlib(""stdole2.tlb"");"
        Print #1,
        Print #1, "    // Forward declare all types defined in this typelib"
        For Each m In p.prjModules
            Print #99, "Pass1: m="; m.Name; " ct="; m.Component.Type
            Select Case m.Component.Type
                Case vbext_ct_StdModule
                Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_UserControl
                Case Else
                    Print #1, "    interface " & m.interfaceName & ";"
            End Select
        Next
        Print #1,
        For Each m In p.prjModules
            Print #99, "Pass2: m="; m.Name; " ct="; m.Component.Type
            Select Case m.Component.Type
                Case vbext_ct_StdModule
                Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_UserControl
                Case Else
                    Dim proc As procTable
                    Dim param As paramTable
                    Dim first_time As Boolean
                    Print #1, "    ["
                    Print #1, "      odl,"
                    Print #1, "      uuid(" & Mid(m.interfaceGUID, 2, Len(m.interfaceGUID) - 2) & "),"
                    ' fixme: hard coding version
                    Print #1, "      version(1.0),"
                    Print #1, "      hidden,"
                    Print #1, "      dual,"
                    Print #1, "      nonextensible,"
                    Print #1, "      oleautomation"
                    Print #1, "    ]"
                    Print #1, "    interface " & m.interfaceName & " : IDispatch {"
                    procNumber = 0
                    For Each proc In m.procs
                        procNumber = procNumber + 1
                        Print #1, "        [id(0x6003" & Right("0000" & Hex(procNumber), 4) & ")]"
                        Print #1, "        HRESULT "; proc.procName; "("
                        first_time = True
                        For Each param In proc.procParams
                            If first_time Then first_time = False Else Print #1, ","
                            Print #1, Space(18);
                            If param.paramVariable.varAttributes And VT_BYREF Then
                                Print #1, "[in, out]";
                            Else
                                Print #1, "[in]";
                            End If
                            Print #1, " " & cTypeName(param.paramVariable) & " " & param.paramVariable.varSymbol;
                        Next
                        Print #1, ");"
                    Next
                    Print #1, "    };"
                    Print #1,
                    Print #1, "    ["
                    Print #1, "      uuid(" & Mid(m.GUID, 2, Len(m.GUID) - 2) & "),"
                    ' fixme: hard coding version
                    Print #1, "      version(1.0)"
                    Print #1, "    ]"
                    Print #1, "    coclass " & m.Name & " {"
                    Print #1, "        [default] interface " & m.interfaceName & ";"
                    If m.Events.Count > 0 Then
                        Print #1, "        [default, source] interface " & m.EventName & ";"
                    End If
                    Print #1, "    };"
                    If m.Events.Count > 0 Then
                        Print #1,
                        Print #1, "    ["
                        Print #1, "      uuid(" & Mid(m.EventGUID, 2, Len(m.EventGUID) - 2) & "),"
                        ' fixme: hard coding version
                        Print #1, "      version(1.0),"
                        Print #1, "      hidden,"
                        Print #1, "      nonextensible"
                        Print #1, "    ]"
                        Print #1, "    dispinterface " & m.EventName & " {"
                        Print #1, "        properties:"
                        ' need to emit properties here
                        Print #1, "        methods:"
                        procNumber = 0
                        For Each proc In m.Events
                            procNumber = procNumber + 1
                            Print #1, "        [id(0x0000" & Right("0000" & Hex(procNumber), 4) & ")]"
                            Print #1, "        HRESULT "; proc.procName; "("
                            first_time = True
                            For Each param In proc.procParams
                                If first_time Then first_time = False Else Print #1, ","
                                Print #1, Space(18);
                                If param.paramVariable.varAttributes And VT_BYREF Then
                                    Print #1, "[in, out]";
                                Else
                                    Print #1, "[in]";
                                End If
                                Print #1, " " & cTypeName(param.paramVariable) & " " & param.paramVariable.varSymbol;
                            Next
                            Print #1, ");"
                        Next
                        Print #1, "    };"
                    End If
            End Select
        Next ' Module
        Print #1, "};"
        Close #1
    End If ' IDL
Next ' Group
Close #99

End Sub


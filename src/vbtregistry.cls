VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Emiter

' Should \\ or / be used as file path component separator?
' Should .exe extension be generated?
' Should \\debug\\ be generated?

Sub Emiter_Emiter(group As vbGroup)
Dim p As vbPrj
Dim m As vbModule
Dim ProjectCount As Single

Open group.OutputPath & "vbtRegistry.lst" For Output As 99
For Each p In group.Projects
    Print #99, "Processing "; p.prjName
    group.AddIn.ProgressBarUpdate ProjectCount / group.Projects.Count * 100
    ProjectCount = ProjectCount + 1
    Dim OutputPath As String
    OutputPath = VBTOutputPath(group, p, "c")
    For Each m In p.prjModules
        If Not p.TypeLib Is Nothing Then
            Dim v As String
            v = p.TypeLib.MajorVersion & "." & p.TypeLib.MinorVersion

'            If m.ComponentType = vbext_ct_StdModule Then
            
            Open OutputPath & p.prjName & ".reg" For Output As 1
            
            Print #1, "REGEDIT4"
            Print #1,
            ' for many entries - use Version numbers from registry
            ' CLSID entries
            Print #1, "[HKEY_CLASSES_ROOT\CLSID\" & m.GUID & "]"
            Print #1, "@=""" & p.prjName & "." & m.Name & """"
            Print #1, """AppID""=""" & m.GUID & """"
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\CLSID\" & m.GUID & "\Implemented Categories]"
            Print #1,
            ' What is this GUID???
            Print #1, "[HKEY_CLASSES_ROOT\CLSID\" & m.GUID & "\Implemented Categories\{40FC6ED5-2438-11CF-A3DB-080036F12502}]"
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\CLSID\" & m.GUID & "\ProgID]"
            Print #1, "@=""" & p.prjName & "." & m.Name & """"
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\CLSID\" & m.GUID & "\Programmable]"
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\CLSID\" & m.GUID & "\TypeLib]"
            Print #1, "@=""" & p.TypeLib.GUID & """"
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\CLSID\" & m.GUID & "\VERSION]"
            Print #1, "@=""" & v & """"
            Print #1,
            'Print #1, "[HKEY_CLASSES_ROOT\CLSID\" & m.GUID & "\_InprocServer32]"
            ' Need to use GetSystemDirectory for path
            'Print #1, "@=""C:\\WINNT\\System32\\autprx32.dll"""
            'Print #1,
            'Print #1, "[HKEY_CLASSES_ROOT\CLSID\" & m.GUID & "\_LocalServer32]"
            Print #1, "[HKEY_CLASSES_ROOT\CLSID\" & m.GUID & "\LocalServer32]"
            ' Need to use GetCurrentDirectory for path
            'Print #1, "@=""" & m.ProjectName & ".exe"""
            Print #1, "@=""" & OutputPath & p.TypeLib.Name & ".exe""" ' .exe can't be correct for UNIX!!
            
            ' Interface entries
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\Interface\" & m.interfaceGUID & "]"
            Print #1, "@=""" & m.Name & """"
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\Interface\" & m.interfaceGUID & "\ProxyStubClsid]"
            ' Need to use ProxyStubClsid from registry
            Print #1, "@=""{00020424-0000-0000-C000-000000000046}"""
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\Interface\" & m.interfaceGUID & "\ProxyStubClsid32]"
            ' Need to use ProxyStubClsid from registry
            Print #1, "@=""{00020424-0000-0000-C000-000000000046}"""
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\Interface\" & m.interfaceGUID & "\TypeLib]"
            Print #1, "@=""" & p.TypeLib.GUID & """"
            Print #1, """Version""=" & v & """"
            
            ' Project entries
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\" & p.prjName & "." & m.Name & "]"
            Print #1, "@=""" & p.prjName & "." & m.Name & """"
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\" & p.prjName & "." & m.Name & "\Clsid]"
            Print #1, "@=""" & m.GUID & """"
            
            ' TypeLib entries
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\TypeLib\" & p.TypeLib.GUID & "]"
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\TypeLib\" & p.TypeLib.GUID & "\" & v & "]"
            Print #1, "@=""" & p.prjName & """"
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\TypeLib\" & p.TypeLib.GUID & "\" & v & "\0]"
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\TypeLib\" & p.TypeLib.GUID & "\" & v & "\0\win32]"
            ' Need to use MS VS path
            Print #1, "@=""" & OutputPath & "debug/" & p.TypeLib.Name & ".tlb""" ' debug/ can't be correct for UNIX!!
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\TypeLib\" & p.TypeLib.GUID & "\" & v & "\FLAGS]"
            Print #1, "@=""0"""
            Print #1,
            Print #1, "[HKEY_CLASSES_ROOT\TypeLib\" & p.TypeLib.GUID & "\" & v & "\HELPDIR]"
            ' Need to use MS VS path
            Print #1, "@=""E:\\Program Files\\Microsoft Visual Studio\\VB98"""
            
            Close #1
        End If
    Next
Next
Close #99

End Sub


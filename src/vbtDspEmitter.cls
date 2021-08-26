VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DspEmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Emit Visual Studio .dsp file

' Todo:
'1. group.OutputPath used to for path to vbt.lib. Probably needs changing.

Option Explicit

Implements Emiter

#If 0 Then
  Private Declare Function FindWindow Lib "user32" _
          Alias "FindWindowA" (ByVal lpClassName As String, _
          ByVal lpWindowName As String) As Long
  
  Private Declare Function SetForegroundWindow Lib "user32" _
          (ByVal hWnd As Long) As Long
#End If

Sub Emiter_Emiter(group As vbGroup)
Dim p As vbPrj
Dim ProjectCount As Single

Open group.OutputPath & "DspEmitter.lst" For Output As 99
For Each p In group.Projects
    Print #99, "Processing "; p.prjName
    group.AddIn.ProgressBarUpdate ProjectCount / group.Projects.Count * 100
    ProjectCount = ProjectCount + 1
    Dim OutputPath As String
    OutputPath = VBTOutputPath(group, p, "c")
    
    ' May need to delete additional MSVS C++ files
    On Error Resume Next
    Kill OutputPath & p.prjName & ".opt" ' delete existing MSVS C++ file
    Kill OutputPath & p.prjName & ".plg" ' delete existing MSVS C++ file
    Kill OutputPath & p.prjName & ".ncb" ' delete existing MSVS C++ file
    Kill OutputPath & p.prjName & ".dsw" ' delete existing MSVS C++ file
    On Error GoTo 0
    Open OutputPath & p.prjName & ".dsp" For Output As 1
    
    Print #1, "# Microsoft Developer Studio Project File - Name="""; p.prjName; """ - Package Owner=<4>"
    Print #1, "# Microsoft Developer Studio Generated Build File, Format Version 6.00"
    Print #1, "# ** DO NOT EDIT **"
    Print #1,
    Print #1, "# TARGTYPE ""Win32 (x86) Console Application"" 0x0103"
    Print #1,
    Print #1, "CFG="; p.prjName; " - Win32 Debug"
    Print #1, "!MESSAGE This is not a valid makefile. To build this project using NMAKE,"
    Print #1, "!MESSAGE use the Export Makefile command and run"
    Print #1, "!MESSAGE"
    Print #1, "!MESSAGE NMAKE /f """; p.prjName; ".mak""."
    Print #1, "!MESSAGE"
    Print #1, "!MESSAGE You can specify a configuration when running NMAKE"
    Print #1, "!MESSAGE by defining the macro CFG on the command line. For example:"
    Print #1, "!MESSAGE"
    Print #1, "!MESSAGE NMAKE /f """; p.prjName; ".mak"" CFG="""; p.prjName; " - Win32 Debug"""
    Print #1, "!MESSAGE"
    Print #1, "!MESSAGE Possible choices for configuration are:"
    Print #1, "!MESSAGE"
    Print #1, "!MESSAGE """; p.prjName; " - Win32 Release"" (based on ""Win32 (x86) Console Application"")"
    Print #1, "!MESSAGE """; p.prjName; " - Win32 Debug"" (based on ""Win32 (x86) Console Application"")"
    Print #1, "!MESSAGE"
    Print #1,
    Print #1, "# Begin Project"
    Print #1, "# PROP AllowPerConfigDependencies 0"
    Print #1, "# PROP Scc_ProjName """""""
    Print #1, "# PROP Scc_LocalPath """""""
    Print #1, "CPP = cl.exe"
    Print #1, "RSC = rc.exe"
    Print #1,
    Print #1, "!IF  ""$(CFG)"" == """; p.prjName; " - Win32 Release"""
    Print #1,
    Print #1, "# PROP BASE Use_MFC 0"
    Print #1, "# PROP BASE Use_Debug_Libraries 0"
    Print #1, "# PROP BASE Output_Dir ""Release"""
    Print #1, "# PROP BASE Intermediate_Dir ""Release; """
    Print #1, "# PROP BASE Target_Dir """""
    Print #1, "# PROP Use_MFC 0"
    Print #1, "# PROP Use_Debug_Libraries 0"
    Print #1, "# PROP Output_Dir ""Release; """
    Print #1, "# PROP Intermediate_Dir ""Release; """
    Print #1, "# PROP Target_Dir """""
    Print #1, "# ADD BASE CPP /nologo /W3 /GX /O2 /D ""WIN32"" /D ""NDEBUG"" /D ""_CONSOLE"" /D ""_MBCS"" /YX /FD /c /I """; Left(group.OutputPath, Len(group.OutputPath) - 1); """"
    Print #1, "# ADD CPP /nologo /W3 /GX /O2 /D ""WIN32"" /D ""NDEBUG"" /D ""_CONSOLE"" /D ""_MBCS"" /YX /FD /c /I """; Left(group.OutputPath, Len(group.OutputPath) - 1); """"
    Print #1, "# ADD BASE RSC /l 0x409 /d ""NDEBUG"""
    Print #1, "# ADD RSC /l 0x409 /d ""NDEBUG"""
    Print #1, "BSC32 = bscmake.exe"
    Print #1, "# ADD BASE BSC32 /nologo"
    Print #1, "# ADD BSC32 /nologo"
    Print #1, "LINK32 = link.exe"
    Print #1, "# ADD BASE LINK32 "; group.OutputPath; "Release\vbt.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib  kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:console /machine:I386"
    Print #1, "# ADD LINK32 "; group.OutputPath; "Release\vbt.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib  kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:console /machine:I386"
    Print #1,
    Print #1, "!ELSEIF  ""$(CFG)"" == """; p.prjName; " - Win32 Debug"""
    Print #1,
    Print #1, "# PROP BASE Use_MFC 0"
    Print #1, "# PROP BASE Use_Debug_Libraries 1"
    Print #1, "# PROP BASE Output_Dir ""Debug"""
    Print #1, "# PROP BASE Intermediate_Dir ""Debug"""
    Print #1, "# PROP BASE Target_Dir """""
    Print #1, "# PROP Use_MFC 0"
    Print #1, "# PROP Use_Debug_Libraries 1"
    Print #1, "# PROP Output_Dir ""Debug"""
    Print #1, "# PROP Intermediate_Dir ""Debug"""
    Print #1, "# PROP Target_Dir """""
    Print #1, "# ADD BASE CPP /nologo /W3 /Gm /GX /ZI /Od /D ""WIN32"" /D ""_DEBUG"" /D ""_CONSOLE"" /D ""_MBCS"" /YX /FD /GZ  /c /I """; Left(group.OutputPath, Len(group.OutputPath) - 1)
    Print #1, "# ADD CPP /nologo /W3 /Gm /GX /ZI /Od /D ""WIN32"" /D ""_DEBUG"" /D ""_CONSOLE"" /D ""_MBCS"" /YX /FD /GZ  /c /I """; Left(group.OutputPath, Len(group.OutputPath) - 1)
    Print #1, "# ADD BASE RSC /l 0x409 /d ""_DEBUG"""
    Print #1, "# ADD RSC /l 0x409 /d ""_DEBUG"""
    Print #1, "BSC32 = bscmake.exe"
    Print #1, "# ADD BASE BSC32 /nologo"
    Print #1, "# ADD BSC32 /nologo"
    Print #1, "LINK32 = link.exe"
    Print #1, "# ADD BASE LINK32 "; group.OutputPath; "Debug\vbt.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib  kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:console /debug /machine:I386 /pdbtype:sept"
    Print #1, "# ADD LINK32 "; group.OutputPath; "Debug\vbt.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib  kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /subsystem:console /debug /machine:I386 /pdbtype:sept"
    Print #1,
    Print #1, "!ENDIF"
    Print #1,
    Print #1, "# Begin Target"
    Print #1,
    Print #1, "# Name """; p.prjName; " - Win32 Release"""
    Print #1, "# Name """; p.prjName; " - Win32 Debug"""
    Print #1, "# Begin Group ""Source Files"""
    Print #1,
    Print #1, "# PROP Default_Filter ""cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"""
    Print #1, "# Begin Source File"
    Print #1,
    Print #1, "SOURCE="; OutputPath; p.prjName; ".c"
    Print #1, "# End Source File"
    Dim m As vbModule
    For Each m In p.prjModules
        Print #1, "# Begin Source File"
        Print #1,
        Print #1, "SOURCE="; OutputPath; m.Name; ".c"
        Print #1, "# End Source File"
    Next
    Dim ref As Reference
    For Each ref In p.VBProject.References
        Print #1, "# Begin Source File"
        Print #1,
        Print #1, "SOURCE="; OutputPath; ref.Name; "_i.c"
        Print #1, "# End Source File"
    Next
    Print #1, "# End Group"
    Print #1, "# Begin Group ""Header Files"""
    Print #1,
    Print #1, "# PROP Default_Filter ""h;hpp;hxx;hm;inl"""
    Print #1, "# End Group"
    Print #1, "# Begin Group ""Resource Files"""
    Print #1,
    Print #1, "# PROP Default_Filter ""ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"""
    Print #1, "# End Group"
    Print #1, "# End Target"
    Print #1, "# End Project"
    Close #1
Next
If group.Projects.Count > 0 Then
    Dim dswOutputPath
    Dim ProjectOutputPath
    OutputPath = VBTOutputPath(group, Nothing, "c")
    dswOutputPath = OutputPath & "group.dsw"
    Print #99, "Creating "; dswOutputPath
    Open dswOutputPath For Output As 1
    Print #1, "Microsoft Developer Studio Workspace File, Format Version 6.00"
    Print #1, "# WARNING: DO NOT EDIT OR DELETE THIS WORKSPACE FILE!"
    Print #1,
    Print #1, "###############################################################################"
    Print #1,
    For Each p In group.Projects
        ProjectOutputPath = VBTOutputPath(group, p, "c") & p.prjName
        Print #1, "Project: """; p.prjName; """="; ProjectOutputPath; ".dsp - Package Owner=<4>"
        Print #1,
        Print #1, "Package=<5>"
        Print #1, "{{{"
        Print #1, "}}}"
        Print #1,
        Print #1, "Package=<4>"
        Print #1, "{{{"
        Print #1, "}}}"
        Print #1,
        Print #1, "###############################################################################"
        Print #1,
    Next
    Print #1, "Global:"
    Print #1,
    Print #1, "Package=<5>"
    Print #1, "{{{"
    Print #1, "}}}"
    Print #1,
    Print #1, "Package=<3>"
    Print #1, "{{{"
    Print #1, "}}}"
    Print #1,
    Print #1, "###############################################################################"
    Print #1,
    Close #1
Else
    dswOutputPath = VBTOutputPath(group, Nothing, "c") & ".dsp"
End If
Print #99, "Shelling out to " & """msdev.exe " & dswOutputPath & """"
On Error Resume Next
Dim taskID As Integer
taskID = Shell("msdev.exe " & dswOutputPath, vbNormalFocus)
#If 0 Then
#If 1 Then
Dim lngHandle As Long
lngHandle = FindWindow("Microsoft Visual C++", App.Title)
DoEvents
If lngHandle <> 0 Then SetForegroundWindow lngHandle
DoEvents
#Else
SetForegroundWindow 0
#End If
#End If
Print #99, "AppActiveate taskID="; taskID
AppActivate taskID, 0 ' don't want to wait
Print #99, "Done"
On Error GoTo 0
Close #99

End Sub

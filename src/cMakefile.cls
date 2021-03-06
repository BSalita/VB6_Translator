VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMakefile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Emit makefile - using UNIX end of line terminator (vbLf)

Option Explicit

Implements Emiter

Sub Emiter_Emiter(group As vbGroup)
Dim p As vbPrj
Dim m As vbModule
Dim ProjectCount As Single

Open group.OutputPath & "cMakefile.lst" For Output As 99
For Each p In group.Projects
    Print #99, "Processing "; p.prjName
    group.AddIn.ProgressBarUpdate ProjectCount / group.Projects.Count * 100
    ProjectCount = ProjectCount + 1
    Dim OutputPath As String
    OutputPath = VBTOutputPath(group, p, "c")
    Open OutputPath & p.prjName & ".mak" For Output As 1
    
    Print #1, "# UNIX make file generated by Softworks vbt"; vbLf;
    Print #1, vbLf;
    Print #1, "# Beware: some make programs are sensitive to special charcters (CR) which editors may render invisible"; vbLf;
    Print #1, vbLf;
    Print #1, "include $(EXXDIR)/$(EXXVERS)/include/makefile.incl"; vbLf;
    Print #1, "include ../../../../vbt1.incl"; vbLf;
    Print #1, vbLf;
'    Print #1, "PRJ_DIR = .."; vbLf; ' not needed
    Print #1, "PRJO = "; p.prjName; ".o";
    For Each m In p.prjModules
        Print #1, " "; m.Name; ".o";
    Next
    Dim ref As Reference
    For Each ref In p.VBProject.References
        Print #1, " "; ref.Name; "_i.o";
    Next
    Print #1, " $(VBTO)"; vbLf;
    Print #1, vbLf;
    Print #1, p.prjName; ": $(PRJO)"; vbLf;
    Print #1, vbTab; "cc $(LFLAGS) $(PRJO) $(LIBS) -o"; p.prjName; vbLf;
    If Not p.TypeLib Is Nothing Then
        ' if ActiveX, register the server by using cmd line switch
        ' fixme: implement regsrvr for DLLs???
        Print #1, vbTab; p.prjName; " -RegServer"; vbLf;
        Print #1, vbTab; "midl $(MIDL_FLAGS) $(MIDL_INCLUDES) $(MIDL_DEFINES) ../"; p.prjName; ".idl"; vbLf;
    End If
    Print #1, vbLf;
    Print #1, "include ../../../../vbt2.incl"; vbLf;
    Print #1, vbLf;
    Print #1, p.prjName; ".o:"; " $(PRJ_DIR)/$(@:.o=.c)"; vbLf;
    Print #1, vbTab; "cc $(DEFS) $(CFLAGSO) $(PRJ_DIR)/"; p.prjName; ".c"; vbLf;
    For Each m In p.prjModules
        Print #1, vbLf;
        Print #1, m.Name; ".o:"; " $(PRJ_DIR)/$(@:.o=.c)"; vbLf;
        Print #1, vbTab; "cc $(DEFS) $(CFLAGSO) $(PRJ_DIR)/"; m.Name; ".c"; vbLf;
    Next
    For Each ref In p.VBProject.References
        Print #1, vbLf;
        Print #1, ref.Name; "_i.o:"; " $(PRJ_DIR)/$(@:.o=.c)"; vbLf;
        Print #1, vbTab; "cc $(DEFS) $(CFLAGSO) $(PRJ_DIR)/"; ref.Name; "_i.c"; vbLf;
    Next
    Close #1
Next
Close #99

End Sub

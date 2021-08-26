VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Emitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Emiter

' ThinkMe: Should group file be created if only 1 project?

Sub Emiter_Emiter(group As vbGroup)

Dim p As vbPrj
Dim m As vbModule
Dim ProjectCount As Single
Dim GroupFile As String
Dim ProjectFile As String

Open group.OutputPath & "vbtVBG.lst" For Output As 99
Dim OutputPath As String
Dim s As String
OutputPath = VBTOutputPath(group, p, "bas")
GroupFile = OutputPath & group.Name & ".vbg"
Print #99, "GroupFile: "; GroupFile; " gpc="; group.Projects.Count
Open GroupFile For Output As 3
Print #3, "VBGROUP 5.0"
'fixme: Insert StartupProject - obtain values from .vbp properties? Print #3, "StartupProject=?"
For Each p In group.Projects
    Print #99, "Processing "; p.prjName; " mc="; p.prjModules.Count
    ProjectFile = OutputPath & p.prjName & ".vbp"
    Print #99, "ProjectFile: "; ProjectFile; " vbfn="; p.VBProject.FileName
    group.AddIn.ProgressBarUpdate 0
    group.AddIn.ProgressBarUpdate ProjectCount / group.Projects.Count * 100
    ProjectCount = ProjectCount + 1
    Dim fso As New FileSystemObject
    If Not fso.FileExists(p.VBProject.FileName) Then
        MsgBox "Project file does not currently exist: " & p.VBProject.FileName
        Print #99, "Project file does not currently exist: " & p.VBProject.FileName
        Err.Raise 1
    End If
    Open p.VBProject.FileName For Input As 1
    Open ProjectFile For Output As 2
    Print #3, "Project="; ProjectFile
    For Each m In p.prjModules
        Select Case m.Component.Type
            Case vbext_ct_ClassModule
                Print #2, "Class="; m.Name; "; "; m.Name; ".cls"
            Case vbext_ct_StdModule
                Print #2, "Module="; m.Name; "; "; m.Name; ".bas"
            Case vbext_ct_MSForm, vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_UserControl
                Print #2, "Form="; m.Name; ".frm"
            Case Else
                Err.Raise 1
        End Select
    Next
    Dim i As Long
    While Not EOF(1)
        Line Input #1, s
        i = InStr(1, s, "=")
        If i > 0 Then
            Select Case UCase(Trim(Mid(s, 1, i - 1)))
                Case "CLASS", "FORM", "MODULE"
                    ' perhaps each filename should be substituted with new filename -- obsoleting above For Each
                    GoTo skipPrint
            End Select
        End If
        Print #2, s
skipPrint:
    Wend
    Close #1
    Close #2
Next
Close #3
Print #99, "Shelling out to """ & group.vbInstance.FullName & " " & GroupFile & """"
On Error Resume Next
Dim taskID As Integer
taskID = Shell(group.vbInstance.FullName & " " & GroupFile, vbNormalFocus)
Print #99, "AppActiveate taskID="; taskID
AppActivate taskID, 0 ' don't want to wait
Print #99, "Done"
On Error GoTo 0
Close #99
End Sub

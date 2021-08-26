VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RPNView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Emiter

Private Sub Class_Terminate()
Unload frmTreeView
End Sub

Sub Emiter_Emiter(group As vbGroup)
Dim g As Node
Dim projects As Node
Dim project As Node
Dim StdModules As Node
Dim StdModule As Node
Dim ClassModules As Node
Dim ClassModule As Node
Dim procs As Node
Dim proc As Node
Dim tok As Node
Dim nod As Node
Dim p As vbPrj
Dim m As vbModule
Dim f As procTable
Dim os As Collection
Dim TotalLines As Single
Dim LineCount As Single

Open group.OutputPath & "RpnView.lst" For Output As 99
group.AddIn.ProgressBarInit "Creating RPN Browsing View"
For Each p In group.projects
    For Each m In p.prjModules
        For Each f In m.procs
            TotalLines = TotalLines + f.procStatements.Count
        Next
    Next
Next
TotalLines = TotalLines / 100
frmTreeView.Caption = "RPN Browsing View"
frmTreeView.Show
Set g = frmTreeView.TreeView1.Nodes.Add(, , , "Group 1")
g.EnsureVisible
Set projects = frmTreeView.TreeView1.Nodes.Add(g.Index, 4, , "Projects")
projects.EnsureVisible
For Each p In group.projects
    Print #99, "Processing "; p.prjName
    Print #99, 15
    Set project = frmTreeView.TreeView1.Nodes.Add(projects.Index, 4, , "Project " & p.prjName)
    Print #99, 16
    project.EnsureVisible
    Print #99, 17
    If Not p.TypeLib Is Nothing Then
        Set nod = frmTreeView.TreeView1.Nodes.Add(project.Index, 4, , "TypeLib.Name " & p.TypeLib.Name)
        Print #99, 18
        nod.EnsureVisible
        Print #99, 19
        Set nod = frmTreeView.TreeView1.Nodes.Add(project.Index, 4, , "TypeLib.GUID " & p.TypeLib.GUID)
        Print #99, 29
        nod.EnsureVisible
        Print #99, 39
        Set nod = frmTreeView.TreeView1.Nodes.Add(project.Index, 4, , "TypeLib.ContainingFile " & p.TypeLib.ContainingFile)
        Print #99, 49
        nod.EnsureVisible
    End If
    Print #99, 59
    For Each m In p.prjModules
    Print #99, 69
        Select Case m.Component.Type
        Case vbext_ct_RelatedDocument
        Case vbext_ct_StdModule
    Print #99, 79
            Set StdModules = frmTreeView.TreeView1.Nodes.Add(project.Index, 4, , "StdModules")
            StdModules.EnsureVisible
            Set StdModule = frmTreeView.TreeView1.Nodes.Add(StdModules.Index, 4, , m.Name)
            StdModule.EnsureVisible
            Set procs = frmTreeView.TreeView1.Nodes.Add(StdModule.Index, 4, , "Procs")
            procs.EnsureVisible
    Print #99, 88
            For Each f In m.procs
    Print #99, 99
                If f.procFunctionResultType Is Nothing Then
                    Set proc = frmTreeView.TreeView1.Nodes.Add(procs.Index, 4, , "Sub " & f.procName)
                Else
                    Set proc = frmTreeView.TreeView1.Nodes.Add(procs.Index, 4, , "Function " & vbTypeName(f.procFunctionResultType))
                End If
    Print #99, 109
                proc.EnsureVisible
                If f.procStatements.Count > 0 Then
    Print #99, 119
                    For Each os In f.procStatements
                        group.AddIn.ProgressBarUpdate LineCount / TotalLines
                        Set tok = frmTreeView.TreeView1.Nodes.Add(proc.Index, 4, , EmitRPN(os))
                        Set tok.Tag = os
                        LineCount = LineCount + 1
                    Next
                    tok.EnsureVisible
                End If
    Print #99, 129
            Next
    Print #99, 139
        Case Else
    ' do only once???
    Print #99, 149
            Set ClassModules = frmTreeView.TreeView1.Nodes.Add(project.Index, 4, , "ClassModules")
            ClassModules.EnsureVisible
    ' end do only once
            Set ClassModule = frmTreeView.TreeView1.Nodes.Add(ClassModules.Index, 4, , m.Name)
            ClassModule.EnsureVisible
    Print #99, 159
            Set nod = frmTreeView.TreeView1.Nodes.Add(ClassModule.Index, 4, , "GUID " & m.GUID)
    Print #99, 169
            nod.EnsureVisible
    Print #99, 179
            Set procs = frmTreeView.TreeView1.Nodes.Add(ClassModule.Index, 4, , "Procs")
    Print #99, 189
            procs.EnsureVisible
            For Each f In m.procs
    Print #99, 199
                If f.procFunctionResultType Is Nothing Then
                    Set proc = frmTreeView.TreeView1.Nodes.Add(procs.Index, 4, , "Sub " & f.procName)
                Else
                    Set proc = frmTreeView.TreeView1.Nodes.Add(procs.Index, 4, , "Function " & vbTypeName(f.procFunctionResultType))
                End If
                proc.EnsureVisible
                If f.procStatements.Count > 0 Then
                    For Each os In f.procStatements
                        group.AddIn.ProgressBarUpdate LineCount / TotalLines
                        Set tok = frmTreeView.TreeView1.Nodes.Add(proc.Index, 4, , EmitRPN(os))
                        Set tok.Tag = os
                        LineCount = LineCount + 1
                    Next
                    tok.EnsureVisible
                End If
            Next
        End Select
    Next
Next
g.Selected = True
'frmTreeView.Timer1.Enabled = True
'frmTreeView.Show
Close #99

End Sub

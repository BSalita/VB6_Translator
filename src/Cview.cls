VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cview"
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

Open group.OutputPath & "CView.lst" For Output As 99
group.AddIn.ProgressBarInit "Creating C Browsing View"
For Each p In group.projects
    For Each m In p.prjModules
        For Each f In m.procs
            TotalLines = TotalLines + f.procStatements.Count
        Next
    Next
Next
TotalLines = TotalLines / 100
frmTreeView.Caption = "C Browsing View"
frmTreeView.Show
Set g = frmTreeView.TreeView1.Nodes.Add(, , , "Group 1")
g.EnsureVisible
Set projects = frmTreeView.TreeView1.Nodes.Add(g.Index, 4, , "Projects")
projects.EnsureVisible
For Each p In group.projects
    Set project = frmTreeView.TreeView1.Nodes.Add(projects.Index, 4, , "Project " & p.prjName)
    project.EnsureVisible
    If Not p.TypeLib Is Nothing Then
        Set nod = frmTreeView.TreeView1.Nodes.Add(project.Index, 4, , "TypeLib.Name " & p.TypeLib.Name)
        nod.EnsureVisible
        Set nod = frmTreeView.TreeView1.Nodes.Add(project.Index, 4, , "TypeLib.GUID " & p.TypeLib.GUID)
        nod.EnsureVisible
        Set nod = frmTreeView.TreeView1.Nodes.Add(project.Index, 4, , "TypeLib.ContainingFile " & p.TypeLib.ContainingFile)
        nod.EnsureVisible
    End If
    For Each m In p.prjModules
        Select Case m.Component.Type
        Case vbext_ct_RelatedDocument
        Case vbext_ct_StdModule
            Set StdModules = frmTreeView.TreeView1.Nodes.Add(project.Index, 4, , "StdModules")
            StdModules.EnsureVisible
            Set StdModule = frmTreeView.TreeView1.Nodes.Add(StdModules.Index, 4, , m.Name)
            StdModule.EnsureVisible
            Set procs = frmTreeView.TreeView1.Nodes.Add(StdModule.Index, 4, , "Procs")
            procs.EnsureVisible
            For Each f In m.procs
                If f.procFunctionResultType Is Nothing Then
                    Set proc = frmTreeView.TreeView1.Nodes.Add(procs.Index, 4, , "Sub " & f.procName)
                Else
                    Set proc = frmTreeView.TreeView1.Nodes.Add(procs.Index, 4, , "Function " & cTypeName(f.procFunctionResultType))
                End If
                proc.EnsureVisible
                If f.procStatements.Count > 0 Then
                    For Each os In f.procStatements
                        group.AddIn.ProgressBarUpdate LineCount / TotalLines
                        Set tok = frmTreeView.TreeView1.Nodes.Add(proc.Index, 4, , EmitInFix(f, os))
                        Set tok.Tag = os
                        LineCount = LineCount + 1
                    Next
                    tok.EnsureVisible
                End If
            Next
        Case Else
    ' do only once???
            Set ClassModules = frmTreeView.TreeView1.Nodes.Add(project.Index, 4, , "ClassModules")
            ClassModules.EnsureVisible
    ' end do only once
            Set ClassModule = frmTreeView.TreeView1.Nodes.Add(ClassModules.Index, 4, , m.Name)
            ClassModule.EnsureVisible
            Set nod = frmTreeView.TreeView1.Nodes.Add(ClassModule.Index, 4, , "GUID " & m.GUID)
            nod.EnsureVisible
            Set procs = frmTreeView.TreeView1.Nodes.Add(ClassModule.Index, 4, , "Procs")
            procs.EnsureVisible
            For Each f In m.procs
                If f.procFunctionResultType Is Nothing Then
                    Set proc = frmTreeView.TreeView1.Nodes.Add(procs.Index, 4, , "Sub " & f.procName)
                Else
                    Set proc = frmTreeView.TreeView1.Nodes.Add(procs.Index, 4, , "Function " & cTypeName(f.procFunctionResultType))
                End If
                proc.EnsureVisible
                If f.procStatements.Count > 0 Then
                    For Each os In f.procStatements
                        group.AddIn.ProgressBarUpdate LineCount / TotalLines
                        Set tok = frmTreeView.TreeView1.Nodes.Add(proc.Index, 4, , EmitInFix(f, os))
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



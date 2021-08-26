VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private frmProgress As frmProgress

Private Sub Class_Initialize()
Set frmProgress = New frmProgress
frmProgress.Show
End Sub

Sub ProgressBarInit(ByVal s1 As Variant, Optional ByVal s2 As Variant, Optional ByVal s3 As Variant)
frmProgress.Label1.Caption = s1
If Not IsMissing(s2) Then frmProgress.Label2.Caption = s2
If Not IsMissing(s3) Then frmProgress.Caption = s3
frmProgress.ZOrder 0 ' Bring window to front - ok?
DoEvents
End Sub

Sub ProgressBarUpdate(ByVal v1 As Variant, Optional ByVal v2 As Variant)
frmProgress.ProgressBar1.Value = v1
If Not IsMissing(v2) Then frmProgress.ProgressBar2.Value = v2
DoEvents
End Sub

Sub ProgressTerminate()
Unload frmProgress
Set frmProgress = Nothing
Exit Sub
frmProgress.Timer1.Enabled = True
DoEvents
End Sub

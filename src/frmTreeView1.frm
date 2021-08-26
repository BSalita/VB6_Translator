VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTreeView 
   Caption         =   "TreeView Browser"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4320
      Top             =   2520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Wait"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8281
      _Version        =   393217
      Indentation     =   353
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub

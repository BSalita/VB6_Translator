VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VBT Make Progress"
   ClientHeight    =   1485
   ClientLeft      =   2640
   ClientTop       =   3120
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   600
      Top             =   480
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7575
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
Unload Me
End Sub

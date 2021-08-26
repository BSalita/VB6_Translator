VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAddIn 
   BackColor       =   &H00000080&
   Caption         =   "Universal Language Translator"
   ClientHeight    =   5460
   ClientLeft      =   765
   ClientTop       =   1050
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5460
   ScaleWidth      =   8880
   Visible         =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Help"
      Height          =   495
      Left            =   5520
      TabIndex        =   16
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3840
      TabIndex        =   15
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      Height          =   495
      Left            =   2160
      TabIndex        =   14
      Top             =   4800
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4065
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7170
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
      TabHeight       =   917
      BackColor       =   128
      TabCaption(0)   =   "Target Language"
      TabPicture(0)   =   "frmAddIn.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Check1(7)"
      Tab(0).Control(1)=   "Check1(6)"
      Tab(0).Control(2)=   "Check1(5)"
      Tab(0).Control(3)=   "Check1(4)"
      Tab(0).Control(4)=   "Check1(8)"
      Tab(0).Control(5)=   "Check1(3)"
      Tab(0).Control(6)=   "Check1(2)"
      Tab(0).Control(7)=   "Check1(1)"
      Tab(0).Control(8)=   "Check1(0)"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Code Generation"
      TabPicture(1)   =   "frmAddIn.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame12"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Support Files"
      TabPicture(2)   =   "frmAddIn.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Activation"
      TabPicture(3)   =   "frmAddIn.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame10"
      Tab(3).Control(1)=   "Frame11"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Treeview"
      TabPicture(4)   =   "frmAddIn.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Check11"
      Tab(4).Control(1)=   "Check6"
      Tab(4).Control(2)=   "Frame7(0)"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Directories"
      TabPicture(5)   =   "frmAddIn.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Frame8"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Frame9"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame10 
         Caption         =   "Add In"
         Height          =   1695
         Left            =   -73320
         TabIndex        =   54
         Top             =   960
         Width           =   1935
         Begin VB.CheckBox Check9 
            Caption         =   "Unload AddIn"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   57
            Top             =   480
            Width           =   1455
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Load on startup"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   56
            Top             =   840
            Width           =   1455
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Command Line"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   55
            Top             =   1200
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "VBT"
         Height          =   1695
         Left            =   -70680
         TabIndex        =   51
         Top             =   960
         Width           =   2535
         Begin VB.CheckBox Check10 
            Caption         =   "Activate on Make Project"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   53
            Top             =   480
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Activate on Make Group"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   840
            Value           =   1  'Checked
            Width           =   2175
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Object Model"
         Height          =   1695
         Left            =   -69360
         TabIndex        =   48
         Top             =   960
         Width           =   2055
         Begin VB.OptionButton Option1 
            Caption         =   "Com/DCom"
            Height          =   495
            Index           =   0
            Left            =   240
            TabIndex        =   50
            Top             =   480
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton Option1 
            Caption         =   "CORBA"
            Height          =   495
            Index           =   1
            Left            =   240
            TabIndex        =   49
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Sync to Code Window"
         Enabled         =   0   'False
         Height          =   255
         Left            =   -74160
         TabIndex        =   45
         Top             =   1500
         Width           =   2415
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Treeview"
         Height          =   255
         Left            =   -74160
         TabIndex        =   44
         Top             =   1140
         Width           =   1575
      End
      Begin VB.Frame Frame9 
         Caption         =   "FTP"
         Height          =   1575
         Left            =   1920
         TabIndex        =   43
         Top             =   1800
         Width           =   4815
         Begin VB.TextBox ftppassword 
            BackColor       =   &H8000000A&
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1200
            PasswordChar    =   "*"
            TabIndex        =   42
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox ftpaddress 
            BackColor       =   &H8000000A&
            Enabled         =   0   'False
            Height          =   285
            HideSelection   =   0   'False
            Left            =   1200
            TabIndex        =   40
            Top             =   360
            Width           =   2535
         End
         Begin VB.TextBox ftpuserid 
            BackColor       =   &H8000000A&
            Enabled         =   0   'False
            Height          =   285
            HideSelection   =   0   'False
            Left            =   1200
            TabIndex        =   41
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox Check8 
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   330
            Width           =   255
         End
         Begin VB.Label passlabel 
            Alignment       =   1  'Right Justify
            Caption         =   "Password :"
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   240
            TabIndex        =   0
            Top             =   1120
            Width           =   855
         End
         Begin VB.Label useridlabel 
            Alignment       =   1  'Right Justify
            Caption         =   "User ID :"
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   240
            TabIndex        =   1
            Top             =   730
            Width           =   855
         End
         Begin VB.Label ftplabel 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
            Caption         =   "Ftp://"
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   600
            TabIndex        =   2
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Output Directory"
         Height          =   855
         Left            =   1920
         TabIndex        =   36
         Top             =   720
         Width           =   4815
         Begin VB.CommandButton Command4 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4320
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox Outputfiletext 
            Height          =   285
            Left            =   120
            TabIndex        =   37
            Text            =   "c:\"
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Expressions"
         Height          =   1215
         Index           =   0
         Left            =   -74400
         TabIndex        =   35
         Top             =   1920
         Width           =   3735
         Begin VB.OptionButton Option4 
            Caption         =   "RPN"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   47
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "INFIX"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Registration Files"
         Height          =   1215
         Left            =   -70560
         TabIndex        =   32
         Top             =   2160
         Width           =   2055
         Begin VB.CheckBox Check5 
            Caption         =   ".VBR"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   34
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox Check5 
            Caption         =   ".REG"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   33
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Type Library"
         Height          =   855
         Left            =   -70560
         TabIndex        =   30
         Top             =   1080
         Width           =   2055
         Begin VB.CheckBox Check4 
            Caption         =   "Intel Byte Order"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "IDL"
         Height          =   1935
         Left            =   -73440
         TabIndex        =   26
         Top             =   1200
         Width           =   2175
         Begin VB.CheckBox Check3 
            Caption         =   "DCE IDL"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   29
            Top             =   1320
            Width           =   1695
         End
         Begin VB.CheckBox Check3 
            Caption         =   "CORBA IDL"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   28
            Top             =   840
            Width           =   1695
         End
         Begin VB.CheckBox Check3 
            Caption         =   "MS IDL"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Alignment"
         Height          =   1095
         Left            =   -74640
         TabIndex        =   23
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton Option3 
            Caption         =   "Native"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Visual Basic"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Byte Order"
         Height          =   1695
         Left            =   -72360
         TabIndex        =   19
         Top             =   960
         Width           =   2535
         Begin VB.OptionButton Option2 
            Caption         =   "Mixed Endian (PDP-11)"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   22
            Top             =   1155
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Big Endian (RISC)"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   765
            Width           =   2055
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Little Endian (INTEL)"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Optimizations"
         Height          =   855
         Left            =   -74640
         TabIndex        =   17
         Top             =   960
         Width           =   2055
         Begin VB.CheckBox Check2 
            Caption         =   "Fold Constants"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "370 Assembler"
         Height          =   255
         Index           =   7
         Left            =   -72000
         TabIndex        =   12
         Top             =   1800
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Caption         =   "68K Assembler"
         Height          =   255
         Index           =   6
         Left            =   -72000
         TabIndex        =   11
         Top             =   1440
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cobol"
         Height          =   255
         Index           =   5
         Left            =   -72000
         TabIndex        =   10
         Top             =   1080
         Width           =   1395
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Visual Basic"
         Height          =   255
         Index           =   4
         Left            =   -74520
         TabIndex        =   9
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "PCode"
         Height          =   255
         Index           =   8
         Left            =   -72000
         TabIndex        =   8
         Top             =   2160
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Caption         =   "PL/I"
         Height          =   255
         Index           =   3
         Left            =   -74520
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Java"
         Height          =   255
         Index           =   2
         Left            =   -74520
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "C++"
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "C"
         Height          =   255
         Index           =   0
         Left            =   -74520
         TabIndex        =   4
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Softworks Ltd."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Connect As Connect

Private Sub CancelButton_Click()
    Connect.Hide
End Sub

Private Sub OKButton_Click()
    MsgBox "AddIn operation on: " & Connect.VBInstance.FullName
End Sub

Private Sub Check6_Click()
 If Check6.Value = 0 Then
    Check11.Enabled = False
    Option4(0).Enabled = False
    Option4(1).Enabled = False
 Else
   Check11.Enabled = True
   Option4(0).Enabled = True
   Option4(1).Enabled = True
    
    
    'Check11.ForeColor = &H80000012
    'useridlabel.ForeColor = &H80000012
  '  passlabel.ForeColor = &H80000012
 '   ftpaddress.BackColor = &HFFFFFF
'    ftpuserid.BackColor = &HFFFFFF
 '   ftppassword.BackColor = &HFFFFFF
 '   ftpaddress.Enabled = True
 '   ftpuserid.Enabled = True
 '   ftppassword.Enabled = True
    
    End If
End Sub

Private Sub Check7_Click(Index As Integer)

End Sub

Private Sub Check8_Click()
    If Check8.Value = 0 Then
    'text=disabled
    ftplabel.ForeColor = &H808080
    useridlabel.ForeColor = &H808080
    passlabel.ForeColor = &H808080
    ftpaddress.BackColor = &H8000000A
    ftpuserid.BackColor = &H8000000A
    ftppassword.BackColor = &H8000000A
    ftpaddress.Enabled = False
    ftpuserid.Enabled = False
    ftppassword.Enabled = False

    Else
    'text=enabled
    ftplabel.ForeColor = &H80000012
    useridlabel.ForeColor = &H80000012
    passlabel.ForeColor = &H80000012
    ftpaddress.BackColor = &HFFFFFF
    ftpuserid.BackColor = &HFFFFFF
    ftppassword.BackColor = &HFFFFFF
    ftpaddress.Enabled = True
    ftpuserid.Enabled = True
    ftppassword.Enabled = True
    
    End If
    
End Sub

Private Sub Command4_Click()
#If 0 Then
Let Opendialog1.InitDir = "c:\"

Opendialog1.ShowOpen
Let Outputfiletext.Text = Opendialog1.FileName
#End If

End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public interfaceName As String
Public interfaceGUID As String

Public Name As String
Public GUID As String
Public procs As New Collection
Public events As New Collection
Public Consts As New Collection
Public Enums As New Collection
Public Types As New Collection
Public ModuleVars As New Collection
Public Declares As New Collection
Public OptionExplicit As Boolean
Public OptionBase As Long
Public OptionPrivateModule As Boolean

Public EventGUID As String
Public EventName As String

Public clsVersionMajor As Long ' VERSION 1.0 CLASS
Public clsVersionMinor As Long ' VERSION 1.0 CLASS
Public clsMultiUse As Integer
Public clsPersistable As Integer
Public clsDataBindingBehavior As Long
Public clsDataSourceBehavior As Long
Public clsMTSTransactionMode As Long
Public clsVB_Creatable As Boolean
Public clsVB_Exposed As Boolean
Public clsVB_Ext_KEY As Boolean
Public clsVB_GlobalNameSpace As Boolean
Public clsVB_Name As String
Public clsVB_PredeclaredID As Boolean

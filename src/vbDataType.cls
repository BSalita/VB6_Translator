VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbDataType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public dtType As tokTypes
Public dtDataName As String
Public dtDataType As TliVarType
Public dtLength As Long
Public dtAttributes As TliVarType
Public dtClass As vbModule ' Local Classes
' fixme: make dtTypeInfo obsolete coclassinfo,interfaceinfo,recordinfo,constantinfo
Public dtTypeInfo As TypeInfo
Public dtClassInfo As CoClassInfo ' TypeLib Classes
Public dtInterfaceInfo As InterfaceInfo
Public dtRecordInfo As RecordInfo
Public dtConstantInfo As ConstantInfo
Public dtUDT As vbType
Public dtEnum As vbEnum
Public dtTLI As TypeLibInfo ' name of defining TLib

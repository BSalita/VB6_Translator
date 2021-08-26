VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "procTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public procName As String
' Public resultType As vbType- just add result as last param (without a name)
Public procParams As New Collection
Public procOptionalParams As Integer
Public procNamedParams As Integer
Public procStatements As New Collection
Public procLocalVariables As New Collection
Public procConsts As New Collection
Public procLabels As New Collection
Public procattributes As procattributes
Public procFunctionResultType As vbVariable
Public procLiteralPool As New Collection
Public procLocalModule As vbModule ' containing module
Public MemberType As vbext_MemberType
Public InvokeKind As InvokeKinds

' Nested structure collections
Public procNests As New Collection
Public procDos As New Collection
Public procFors As New Collection
Public procSingleIfs As New Collection
Public procBlockIfs As New Collection
Public procSelects As New Collection
Public procWhiles As New Collection
Public procWiths As New Collection


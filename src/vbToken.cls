VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public tokType As tokTypes
Public tokString As String
Public tokComponent As VBComponent
Public tokLineNumber As Long
Public tokCol As Long
Public tokDataType As Integer
Public tokOperator As vbOpr
'Public tokInterfaceInfo As InterfaceInfo
Public tokLocalFunction As proctable
Public tokReference As reference ' Type library reference
Public tokDeclarationInfo As DeclarationInfo
Public tokMemberInfo As MemberInfo
Public tokReturnType As VarTypeInfo
Public tokInterfaceInfo As InterfaceInfo
Public tokConst As vbConst
Public tokEnumMember As vbEnumMember
Public tokDeclare As vbDeclare
Public tokVariable As vbVariable
Public tokPriority As Integer   ' expression priority
Public tokPCode As vbPCodes
Public tokPCodeSubType As Integer
Public tokKeyword As Keywords
Public tokOutput As String
Public tokLabel As vbLabel
Public tokValue As Variant
Public tokLHS As Long ' output_stack counter for left-hand side of expression
Public tokRHS As Long ' output_stack counter for right-hand side of expression (always .count?)
Public tokModule As vbModule
Public tokCount As Long
Public tokRank As Long
'Public tokReturnTypeType As tokTypes
'Public tokReturnTypeInterfaceInfo As InterfaceInfo

' might want to implement other control structures - useful for EXIT, etc. statements
' fixme: need to implement these in begin and end control structures (With is done)
Public tokFor As vbFor
Public tokDo As vbDo
Public tokIf As vbIf
Public tokSelect As vbSelect
Public tokWhile As vbWhile
Public tokWith As vbWith


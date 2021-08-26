VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbOprBinaryC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements vbOpr
Implements vbOpr

Public Sub vbOpr_oprOutput(ByVal OptimizeFlag As Boolean, ByVal cOutputStack As Collection, ByVal token As vbToken)

End Sub

Public Function vbOpr_oprPCode() As vbPCodes

End Function

Public Function vbOpr_oprPriority() As Integer

End Function

Public Function vbOpr_oprCoerceOperandsBinary(ByVal cOutputStack As Collection) As tli.TliVarType

End Function

Public Function vbOpr_oprGetResultTypeBinary(ByVal l As tli.TliVarType, ByVal r As tli.TliVarType) As tli.TliVarType

End Function

Public Sub vbOpr_oprLetResultTypeBinary(ByVal l As tli.TliVarType, ByVal r As tli.TliVarType, ByVal vt As tli.TliVarType)

End Sub

Public Function vbOpr_oprOperateBinary(ByVal l As Variant, ByVal r As Variant) As Variant

End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbOpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Sub oprOutput(ByVal OptimizeFlag As Integer, ByVal output_stack As Collection, ByVal token As vbToken)

End Sub

Public Function oprPCode() As vbPCodes

End Function

Public Function oprPriority() As Integer

End Function

Public Function oprGetResultTypeUnary(ByVal o As TliVarType) As TliVarType

End Function

Public Sub oprLetResultTypeUnary(ByVal o As TliVarType, ByVal vt As TliVarType)

End Sub

Public Function oprCoerceOperandUnary(ByVal OptimizeFlag As Long, ByVal output_stack As Collection, ByVal token As vbToken) As TliVarType

End Function

Public Function oprOperateUnary(ByVal o As Variant) As Variant

End Function

Public Function oprGetResultTypeBinary(ByVal l As TliVarType, ByVal r As TliVarType) As TliVarType

End Function

Public Sub oprLetResultTypeBinary(ByVal l As TliVarType, ByVal r As TliVarType, ByVal vt As TliVarType)

End Sub

Public Function oprCoerceOperandsBinary(ByVal OptimizeFlag As Long, ByVal output_stack As Collection, ByVal token As vbToken) As TliVarType

End Function

Public Function oprOperateBinary(ByVal l As Variant, ByVal r As Variant) As Variant

End Function


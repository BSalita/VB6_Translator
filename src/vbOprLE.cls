VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbOprLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements vbOpr

Private m_oprResultType(VT_MAXTYPE, VT_MAXTYPE) As TliVarType

Private Sub Class_Initialize()
BinaryOperatorInit Me
End Sub

Public Function vbOpr_oprGetResultTypeUnary(ByVal o As TliVarType) As TliVarType
Err.Raise 1 ' Internal error - Not a Unary operator
End Function

Public Sub vbOpr_oprLetResultTypeUnary(ByVal o As TliVarType, ByVal vt As TliVarType)
Err.Raise 1 ' Internal error - Not a Unary operator
End Sub

Public Function vbOpr_oprCoerceOperandUnary(ByVal OptimizeFlag As OptimizeFlags, ByVal output_stack As Collection, ByVal token As vbToken) As TliVarType
Err.Raise 1 ' Internal error - Not a Unary operator
End Function

Public Function vbOpr_oprOperateUnary(ByVal o As Variant) As Variant
Err.Raise 1 ' Internal error - Not a Unary operator
End Function

Public Function vbOpr_oprCoerceOperandsBinary(ByVal OptimizeFlag As OptimizeFlags, ByVal output_stack As Collection, ByVal token As vbToken) As TliVarType
vbOpr_oprCoerceOperandsBinary = CoerceBinaryOperands(OptimizeFlag, output_stack, token)
End Function

Public Function vbOpr_oprGetResultTypeBinary(ByVal l As TliVarType, ByVal r As TliVarType) As TliVarType
vbOpr_oprGetResultTypeBinary = m_oprResultType(l, r)
End Function

Public Sub vbOpr_oprLetResultTypeBinary(ByVal l As TliVarType, ByVal r As TliVarType, ByVal vt As TliVarType)
m_oprResultType(l, r) = vt
End Sub

Public Function vbOpr_oprOperateBinary(ByVal l As Variant, ByVal r As Variant) As Variant
vbOpr_oprOperateBinary = l <= r
End Function

Public Sub vbOpr_oprOutput(ByVal OptimizeFlag As Integer, ByVal output_stack As Collection, ByVal token As vbToken)
BinaryOutput OptimizeFlag, output_stack, token, Me
End Sub

Public Function vbOpr_oprPCode() As vbPCodes
vbOpr_oprPCode = vbPCodeLE
End Function

Public Function vbOpr_oprPriority() As Integer
vbOpr_oprPriority = vbOprPriorityCmp
End Function






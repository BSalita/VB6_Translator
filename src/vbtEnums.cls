VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbtEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Keywords
    KW_EMPTY
    KW_ADDRESSOF
    KW_AS
    KW_BYREF
    KW_BYVAL
    KW_DEBUG
    KW_DO ' also in stmt collection
    KW_EACH
    KW_ELSE ' also in stmt collection
    KW_ELSEIF ' also in stmt collection
    KW_FOR ' also in stmt collection
    KW_FUNCTION ' also in stmt collection
    KW_GOTO ' also in stmt collection
    KW_GOSUB ' also in stmt collection
    KW_IF ' also in stmt collection
    KW_IN
    KW_LBOUND
    KW_LOCAL ' obsolete - On Local Error ...
    KW_NEW
    KW_OPTIONAL
    KW_PARAMARRAY
    KW_PRESERVE
    KW_PROPERTY ' also in stmt collection
    KW_SELECT ' also in stmt collection
    KW_STEP
    KW_SUB ' also in stmt collection
    KW_THEN
    KW_TO
    KW_TYPE ' also in stmt collection
    KW_TYPEOF
    KW_UBOUND
    KW_UNTIL
    KW_WHILE ' also in stmt collection
    KW_WITH ' also in stmt collection
    KW_WITHEVENTS
End Enum

Public Enum tokTypes
    tokInvalid ' must be 0
    tokAddRef
    tokAddressOf
    tokArrayVariable
    tokBlanks
    tokByVal
    tokCase
    tokCaseIs
    tokCaseTo
    tokColon
    tokComment
    tokConst
    tokConstantInfo
    tokCvt
    tokDeclarationInfo
    tokDeclare
    tokEnumMember
    tokExpression
    tokFormClass
    tokGlobalModule
    tokIDispatchInterface
    tokInvoke
    tokInvokeDefaultMember
    tokKeyword
    tokLabelDef
    tokLabelRef
    tokLBound
    tokLocalModule
    tokme
    tokMember
    tokMissing
    tokNamedParameter ' symbol :=
    tokNewObject
    tokNothing
    tokOperands
    tokOperator
    tokProjectClass
    tokQI_Module
    tokQI_TLibInterface
    tokReDim
    tokReferenceClass
    tokSpecial
    tokstatement
    tokStdProcedure
    tokSubRef
    toksymbol
    tokTabs
    tokTypeOf
    tokUBound
    tokUDT
    tokVariable
    tokVariant
    tokVariantArgs
    tokWith
    tokWithValue
End Enum

' Needs to be cleaned-up along with variable attributes
Public Enum procattributes
    PROC_ATTR_DEFAULT = &H1
    PROC_ATTR_Friend = &H2
    PROC_ATTR_PRIVATE = &H4
    PROC_ATTR_PUBLIC = &H8
    PROC_ATTR_Static = &H10
'    PROC_ATTR_PROPERTYGET = &H20
'    PROC_ATTR_PROPERTYLET = &H40
'    PROC_ATTR_PROPERTYSET = &H80
    PROC_ATTR_VARIABLE = &H100
    PROC_ATTR_FUNCTION = &H200 ' as opposed to Sub
    proc_attr_defined = &H400
    PROC_ATTR_FORM_CONTROL = &H800
    ' &H1000 to &H8000 are reserved
End Enum

Public Enum vbPCodes
    vbPCode0
' Operators
    vbPCodeImp
    vbPCodeEqv
    vbPCodeXor
    vbPCodeOr
    vbPCodeAnd
    vbPCodeNot
    vbPCodeCmp    ' =,<,<=,<>,>,>=,IS,LIKE
    vbPCodeEQ
    vbPCodeLT
    vbPCodeLE
    vbPCodeNE
    vbPCodeGT
    vbPCodeGE
    vbPCodeIs
    vbPCodeLike
    vbPCodeCat
    vbPCodeAdd
    vbPCodeSub
    vbPCodeMod
    vbPCodeIDiv
    vbPCodeMul
    vbPCodeDiv
    vbPCodePositive
    vbPCodeNegative
    vbPCodePow
' Functions
    vbPCodeCvt
'    vbPCodeMember
    vbPCodeUDT
' Print Statement PCodes
    vbPCodePrintSpc
    vbPCodePrintTab
    vbPCodePrintSemiColon
    vbPCodePrintComma
' Statements
    vbPCodeStmtLet
    vbPCodeCall
    vbPCodeCase
    vbPCodeCaseElse
    vbPCodeCloseFile
'    vbPCodeDate
    vbPCodeDeclareCall
    vbPCodeDebugAssert
    vbPCodeDebugPrint
    vbPCodeDo
    vbPCodeDoUntil
    vbPCodeDoWhile
    vbPCodeElse
    vbPCodeElseIf
    vbPCodeEnd
    vbPCodeEndIf
    vbPCodeEndSelect
    vbPCodeEndWith
    vbPCodeErase
    vbPCodeError
    vbPCodeExitDo
    vbPCodeExitFor
    vbPCodeExitFunction
    vbPCodeExitProperty
    vbPCodeExitSub
    vbPCodefor
    vbPCodeforeach
    vbPCodeGet
    vbPCodeGoSub
    vbPCodeGoTo
    vbPCodeIf
    vbPCodeInput
    vbPCodeLet
    vbPCodeLineInput
    vbPCodeLock
    vbPCodeLoop
    vbPCodeLoopInfinite
    vbPCodeLoopUntil
    vbPCodeLoopWhile
    vbPCodeLSet
    vbPcodeMid
    vbPcodeMidB
    vbPCodeName
    vbPCodeForEachNext
    vbPCodeForEachNextV
    vbPCodeForNext
    vbPCodeForNextV
    vbPCodeOn
    vbPCodeOnError0
    vbPCodeOnErrorLabel
    vbpcodeonerrorresumenext
    vbPCodeOnGoSub
    vbPCodeOnGoTo
    vbPCodeOpen
    vbPCodePrint
    vbPCodePrintMethod
    vbPCodePropertyGet
    vbPCodePropertyLet
    vbPCodePropertySet
    vbPCodePut
    vbPCodeRaiseEvent
    vbPCodeReDim
    vbPCodeResume0
    vbPCodeResumeLabel
    vbPCodeResumeNext
    vbPCodeReturn
    vbPCodeRSet
    vbPCodeSeek
    vbPCodeSelect
    vbPCodeSet
    vbPCodeSetNothing
    vbPCodePropertySetNothing
    vbPCodeSingleIf
    vbPCodeSingleIfElse
    vbPCodeSingleIfEndIf
    vbPCodeSingleIfEndIfElse
    vbPCodeStop
'    vbPCodeTime
    vbPCodeUnlock
    vbPCodeWend
    vbPCodeWhile
    vbPCodeWidth
    vbPCodeWith
    vbPCodeWrite
    vbpcodeCircle
    vbpcodeline
    vbpcodepset
    vbpcodescale
End Enum

Public Enum VariableAttributes
' procAttributes use first 12 bits - need to merge definitions?
    VARIABLE_WITHEVENTS = &H10000
    VARIABLE_NEW = &H20000
    VARIABLE_PRESERVE = &H40000
    VARIABLE_OPTIONAL = &H80000 ' for optional parameters
    VARIABLE_PARAMARRAY = &H100000 ' ParamArray parameter
    VARIABLE_PARAMETER = &H200000 ' variable is a parameter
    VARIABLE_FUNCTION = &H400000 ' variable is a function name
'    VARIABLE_ME = &H800000 ' Me (This) class object
    VARIABLE_PUTVAL = &H1000000 ' Property Let/Set value
    VARIABLE_DEFAULTVALUE = &H2000000 ' Optional parameter has default value
    VARIABLE_CONTROLARRAY = &H4000000 ' Form control array - interface members - Count,Item,LBound,UBound
    VARIABLE_ARRAY = VT_ARRAY ' &H2000
    VARIABLE_BYREF = VT_BYREF ' &H4000
End Enum


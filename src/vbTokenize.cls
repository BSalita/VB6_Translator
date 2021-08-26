VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbTokenize"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Sub TokenizeSymbol(ByRef s As String, ByRef col As Long)
'Print #99, "1 ts="; Mid(s, col); " col="; col
If Mid(s, col, 1) = "[" Then
    ' fixme:? don't like modifying s for later reprocessing.
    s = Left(s, col - 1) & Mid(s, col + 1)
    col = InStr(col, s, "]")
    s = Left(s, col - 1) & Mid(s, col + 1)
'    Print #99, "2 ts="; Mid(s, col); " col="; col
Else
    Do
        col = col + 1
    Loop While issym(Mid(s, col, 1))
    If col <= Len(s) Then
        If asc(Mid(s, col, 1)) < 32 Or asc(Mid(s, col, 1)) > 126 Then
            Print #99, "Invalid character found. Non-English alphabets are not supported: "; s
            MsgBox "Invalid character found. Non-English alphabets are not supported: " & s
            Err.Raise 1
        End If
    End If
End If
End Sub

Sub TokenizeNumeric(ByVal s As String, ByRef col As Long, ByRef v As Variant)
Dim scol As Long
Dim p As Boolean ' contains period?
Dim e As Boolean ' contains exponential notation?
scol = col
' scan initial digits
While isdigit(Mid(s, col, 1))
    col = col + 1
Wend
' stopped on decimal point?
If Mid(s, col, 1) = "." Then
    p = col
    col = col + 1
    ' scan digits after decimal point, if any
    Do
        col = col + 1
    Loop While isdigit(Mid(s, col, 1))
End If
' stopped on exponential notation?
If UCase(Mid(s, col, 1)) = "E" Then
    ' signed integer follows E
    If isdigit(Mid(s, col + 1, 1)) Then
        col = col + 1
    ElseIf (Mid(s, col + 1, 1) = "+" Or Mid(s, col + 1, 1) = "-") And isdigit(Mid(s, col + 2, 1)) Then
        col = col + 2
    Else
        Err.Raise 1 ' invalid exponential notation
    End If
    While isdigit(Mid(s, col, 1))
        col = col + 1
    Wend
End If
' stopped on data type specification?
'Print #99, "Tokenize: s="; s; " scol="; scol; " col="; col
Select Case Mid(s, col, 1)
    Case "#"
        v = CDbl(Mid(s, scol, col - scol))
        col = col + 1
    Case "!"
        v = CSng(Mid(s, scol, col - scol))
        col = col + 1
    Case "&"
        If p Or e Then Err.Raise 1 ' invalid long constant
        v = CLng(Mid(s, scol, col - scol))
        col = col + 1
    Case "%"
        If p Or e Then Err.Raise 1 ' invalid integer constant
        v = CInt(Mid(s, scol, col - scol))
        col = col + 1
    Case "@"
        v = CCur(Mid(s, scol, col - scol))
        col = col + 1
    Case Else
        v = CDbl(Mid(s, scol, col - scol))
        If Not (p Or e) Then
            On Error Resume Next
            v = CLng(v)
            v = CInt(v)
            On Error GoTo 0
        End If
End Select
' check for valid character following numeric - should this be done in TokenizeLine?
'Print #99, "Tokenize: s="; s; " scol="; scol; " col="; col
Select Case Mid(s, col, 1) ' should other tokenize types check for invalid next character?
' don't error if !, its a valid operator
    Case "0" To "9", "a" To "z", "A" To "Z", "#", "&", "%", "@", "$" ' fill out this list. Allow only operators, space, tab, eol?
        Err.Raise 1 ' invalid numeric constant
End Select
End Sub

Sub TokenizeBlanks(ByVal s As String, ByRef col As Long)
Do
    col = col + 1
Loop While Mid(s, col, 1) = " "
End Sub

Sub TokenizeTabs(ByVal s As String, ByRef col As Long)
Do
    col = col + 1
Loop While Mid(s, col, 1) = vbTab
End Sub

Sub TokenizeComment(ByVal s As String, ByRef col As Long)
col = Len(s) + 1
End Sub

Sub TokenizeString(ByVal s As String, ByRef col As Long, ByRef v As Variant)
Dim ecol As Long
v = ""
ecol = Len(s)
Do
    col = col + 1
    While Mid(s, col, 2) = """"""
        v = v & Mid(s, col, 1)
        col = col + 2
    Wend
    If Mid(s, col, 1) = """" Then col = col + 1: Exit Sub
    v = v & Mid(s, col, 1)
Loop While col < ecol
Err.Raise 1 ' missing closing quote
End Sub

Sub TokenizeSpecial(ByVal s As String, ByRef col As Long, ByVal token As vbToken)
Dim v As Variant
Dim scol As Long
col = col + 1
token.tokType = tokSpecial
If Mid(s, col - 1, 1) = "&" Then
    Select Case UCase(Mid(s, col, 1))
        Case "H"
            col = col + 1
            scol = col
            ' need col <= Len(s) because Instr was returning 1 otherwise in VB6
            While col <= Len(s) And InStr(1, "0123456789ABCDEF", UCase(Mid(s, col, 1))) > 0
                col = col + 1
            Wend
            If col = scol Then Err.Raise 1 ' Incomplete hex constant
            v = Val(Mid(s, scol - 2, col - scol + 2)) ' convert to hex
        Case "O"
            col = col + 1
            scol = col
            While col <= Len(s) And InStr(1, "01234567", Mid(s, col, 1)) > 0
                col = col + 1
            Wend
            If col = scol Then Err.Raise 1 ' Incomplete Octal constant
            v = Val(Mid(s, scol - 2, col - scol + 2)) ' convert to octal
        Case "0" To "7"
            scol = col
            While col <= Len(s) And InStr(1, "01234567", Mid(s, col, 1)) > 0
                col = col + 1
            Wend
            If col = scol Then Err.Raise 1 ' Incomplete Octal constant
            v = Val(Mid(s, scol - 1, col - scol + 1)) ' convert to octal
        Case "A" To "Z", "8" To "9"
            Err.Raise 1 ' Invalid syntax
    End Select
    If scol > 0 Then
        If Mid(s, col, 1) = "%" Then
            v = CInt(v)
            col = col + 1
        ElseIf Mid(s, col, 1) = "&" Then
            v = CLng(v)
            col = col + 1
        Else
            v = CLng(v)
            On Error Resume Next
            v = CInt(v)
            On Error GoTo 0
        End If
        token.tokType = tokVariant
        token.tokValue = v
        token.tokDataType = varType(v)
    End If
End If
End Sub

Function TokenizeLine(ByVal s As String) As Collection
Print #99, "TokenizeLine: s="; s
Dim tokens As New Collection
Dim col As Long
Dim c As String * 1
#If 0 Then ' Allow special characters in comments?
For col = 1 To Len(s)
    If asc(Mid(s, col, 1)) < 32 Or asc(Mid(s, col, 1)) > 126 Then
        Print #99, "Invalid character found. Non-English alphabets are not supported: "; s
        MsgBox "Invalid character found. Non-English alphabets are not supported: " & s
        Err.Raise 1
    End If
Next
#End If
Dim v As Variant
col = 1
Dim token As vbToken
Do While col <= Len(s)
10
    c = Mid(s, col, 1)
    Print #99, "c="; c
    Set token = New vbToken
    Set token.tokComponent = currentComponent
    token.tokLineNumber = currentLineNumber
    token.tokCol = col
    Select Case c
    Case "A" To "Z", "a" To "z"
        TokenizeSymbol s, col
        token.tokType = toksymbol
        token.tokString = Mid(s, token.tokCol, col - token.tokCol)
        Select Case Mid(s, col, 1)
        Case "#"
            token.tokDataType = vbDouble
            col = col + 1
        Case "!"
            ' note: ambigous treatment of ! (datatype vs. operator)
            If Not isalpha(Mid(s, col + 1, 1)) And Mid(s, col + 1, 1) <> "[" Then ' skip if ! operator
                token.tokDataType = vbSingle
                col = col + 1
            End If
        Case "&"
            token.tokDataType = vbLong
            col = col + 1
        Case "%"
            token.tokDataType = vbInteger
            col = col + 1
        Case "@"
            token.tokDataType = vbCurrency
            col = col + 1
        Case "$"
            token.tokDataType = vbString
            col = col + 1
        Case Else
            If tokens.count = 0 Then GoTo process_special_symbols
            ' Don't process special symbols if they're members, they are valid member names emanating from C* typelibs, but are invalid for VB compiler.
            If tokens.Item(tokens.count).tokString <> "." Then ' could check against tokMember and tokWith
process_special_symbols:
                Select Case UCase(token.tokString)
                Case "EMPTY"
                    token.tokType = tokVariant
                    token.tokValue = Empty
                    token.tokDataType = vbVariant
                Case "FALSE"
                    token.tokType = tokVariant
                    token.tokValue = False
                    token.tokDataType = varType(token.tokValue)
                Case "NOTHING"
                    token.tokType = tokNothing
                    token.tokDataType = vbObject
                Case "NULL"
                    token.tokType = tokVariant
                    token.tokValue = Null
                    token.tokDataType = vbVariant
                Case "REM"
                    token.tokType = tokComment
                    col = Len(s) + 1
                    token.tokString = Mid(s, token.tokCol)
                Case "TRUE"
                    token.tokType = tokVariant
                    token.tokValue = True
                    token.tokDataType = varType(token.tokValue)
                Case Else
                    ' do nothing
                End Select
            End If
        End Select
        GoTo 100
    Case "."
        ' Distinguish whether . is a numeric constant, member separator or With placeholder
        If isdigit(Mid(s, col + 1, 1)) Then GoTo 20 ' jumps into Case
        token.tokType = tokWith ' don't increment col
        token.tokString = "."
        If tokens.count > 0 Then If tokens.Item(tokens.count).tokType = tokWith Or tokens.Item(tokens.count).tokType = toksymbol Or tokens.Item(tokens.count).tokString = ")" Then token.tokType = tokMember: col = col + 1
        GoTo 100
    Case "0" To "9"
20
        TokenizeNumeric s, col, v
        token.tokType = tokVariant
        token.tokValue = v
        token.tokDataType = varType(v)
    Case " "
        TokenizeBlanks s, col
        token.tokType = tokBlanks
    Case vbTab
        TokenizeTabs s, col
        token.tokType = tokTabs
    Case "!"
        ' not sure how to best handle this, so for now, stuff ("...") in place of !
#If 1 Then
        Print #99, "before ! s="; s
        ' note: ! may appear at the beginning of an expression, its both a ! and With
        If tokens.count = 0 Then GoTo 30
        If tokens.Item(tokens.count).tokType <> tokWith And tokens.Item(tokens.count).tokType <> toksymbol And tokens.Item(tokens.count).tokString <> ")" Then
30          tokens.Add New vbToken
            tokens.Item(tokens.count).tokType = tokWith
            tokens.Item(tokens.count).tokString = "."
        End If
        col = col + 1 ' skip !
        ' allow x![field name]
        TokenizeSymbol s, col
        s = Left(s, token.tokCol - 1) & "(""" & Mid(s, token.tokCol + 1, col - token.tokCol - 1) & """)" & Mid(s, col)
        col = token.tokCol ' reset to ("...")
        Set token = Nothing
        Print #99, "after ! s="; s
        GoTo 10
#Else
        Print #99, "! operator not supported"
        MsgBox "! operator not support"
        Err.Raise 1
#End If
    Case "'"
        TokenizeComment s, col
        token.tokType = tokComment
    Case """"
        TokenizeString s, col, v
        token.tokType = tokVariant
        token.tokValue = v
        token.tokDataType = varType(v)
    Case "["
        token.tokType = toksymbol
        TokenizeSymbol s, col
        token.tokString = Mid(s, token.tokCol, col - token.tokCol)
'        Print #99, "ts="; token.tokString
        GoTo 100
    Case "#"
        Dim i As Long
        i = InStr(col + 1, s, "#")
        If i = 0 Then GoTo isspecial
        token.tokType = tokVariant
        token.tokString = Mid(s, col + 1, i - col - 1)
        If Not IsDate(token.tokString) Then GoTo isspecial
        token.tokValue = CDate(token.tokString)
        token.tokDataType = vbDate
        col = i + 1
    Case ":"
        col = col + 1
        If Mid(s, col, 1) = "=" Then ' fixme: ok to assume ":=" has no embedded spaces?
            col = col + 1
            If tokens.count = 0 Then Err.Raise 1 ' Parameter name missing
            If tokens.Item(tokens.count).tokType <> toksymbol Then Err.Raise 1 ' Parameter name invalid
            tokens.Item(tokens.count).tokType = tokNamedParameter
            GoTo 10
        Else
            token.tokType = tokColon
        End If
    Case Else
isspecial:
        TokenizeSpecial s, col, token
    End Select
    token.tokString = Mid(s, token.tokCol, col - token.tokCol)
100
    tokens.Add token
Loop

' Prepare line to be parsed
For i = tokens.count To 1 Step -1
    If tokens.Item(i).tokType = tokBlanks Or tokens.Item(i).tokType = tokTabs Or tokens.Item(i).tokType = tokComment Then
        tokens.Remove i ' remove blanks, tabs, comments
    Else
        Dim opr As vbOpr
        Set opr = Nothing
        On Error Resume Next
        Set opr = cBinaryOperators.Item(UCase(tokens.Item(i).tokString))
        If opr Is Nothing Then Set opr = cUnaryOperators.Item(UCase(tokens.Item(i).tokString))
        On Error GoTo 0
        If opr Is Nothing Then
            If tokens.Item(i).tokType = toksymbol Then
                Dim kw As Keywords
                kw = getKeyword(tokens.Item(i))
                If kw <> KW_EMPTY Then
                    If i = 1 Then GoTo 1000
                    If tokens.Item(i - 1).tokType <> tokMember Then
1000
                        tokens.Item(i).tokType = tokKeyword
                        tokens.Item(i).tokKeyword = kw
                    End If
                End If
            End If
        Else ' fold double operators into one "< >" --> "<>"
            tokens.Item(i).tokType = tokOperator
            If i < tokens.count Then
                If tokens.Item(i + 1).tokType = tokOperator Then
                    Set opr = Nothing
                    On Error Resume Next
                    Set opr = cBinaryOperators.Item(tokens.Item(i).tokString & tokens.Item(i + 1).tokString) ' don't need UCase
                    On Error GoTo 0
                    If Not opr Is Nothing Then
                        tokens.Item(i).tokString = tokens.Item(i).tokString & tokens.Item(i + 1).tokString
                        tokens.Remove i + 1
                    End If
                End If
            End If
        End If
    End If
Next

Set TokenizeLine = tokens
Print #99, "TokenizeLine: count="; TokenizeLine.count
End Function

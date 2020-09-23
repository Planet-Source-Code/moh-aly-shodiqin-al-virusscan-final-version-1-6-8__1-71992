VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* ********************
'*     Alias Module
'* ********************

Public rootEngine      As cMain

'* Constants
Private Const BEGIN_FUNCTION = "["
Private Const END_FUNCTION = "]"
Private Const BEGIN_SGROUP = "("
Private Const END_SGROUP = ")"
Private Const COMMENT_CHAR = "`"
Private Const COMMENT_ML_CHAR = "#"
Private Const PARAM_DELIM = " "
Private Const BEGIN_VARLOCAL = "$"
Private Const BEGIN_VARGLOBAL = "%"
Private Const ESCAPE_CHAR = "\"


Private currentLine As Integer

'* if stack
Private stack_if()      As typIf
Private stack_ifcount   As Integer
Private Type typIf
    bTrueYet    As Boolean
    bNoEval     As Boolean
End Type

'* while loop stack
Private stack_while()       As typWhileLoop
Private stack_whilecount    As Integer
Private Type typWhileLoop
    bIsFalse    As Boolean
    iReturnLine As Integer
    iLastLine   As Integer
    bInit       As Boolean
End Type

'* loop stack
Private stack_loop()    As typLoop
Private stack_loopcount As Integer
Private Type typLoop
    current As Long
    total   As Long
    linenum As Integer
    varname As String
End Type

'* Args
Private Args()      As String
Private ArgCount    As Integer

'* Var Type
Private variableType As Integer
Private Enum varType
    LOCAL_
    GLOBAL_
End Enum

Public bInComment       As Boolean
Public bGotoNextLine    As Boolean
Public returnValue     As String

'* Alias Information
Private strName         As String
Private AliasType       As enAliasType
Private strExtraParams  As String
Private Enum enAliasType
    at_ALIAS
    at_EVENT
End Enum

'* Local variables
Private variables()     As typVariable
Private varCount        As Integer
Private Enum enumVarType
    xString
    xInteger
    xReal
End Enum
Private Type typVariable
    name    As String
    Value() As String
    type    As enumVarType
End Type

'* Code storage
Private strCode()       As String
Private intCodeLines    As Integer

'* Call stack type
Private Type typAliasCall
    ArgCount    As Integer
    Args()      As Variant
    bQuote      As Boolean
End Type

'* call stack
Private stack_calls()       As typAliasCall
Private stack_callcount     As Integer


'* Execution multiplier
Private em_total    As Integer
Private em_current  As Integer

Public Sub AddCodeLine(strCodeLine As String)
    intCodeLines = intCodeLines + 1
    ReDim Preserve strCode(1 To intCodeLines) As String
    strCode(intCodeLines) = strCodeLine
End Sub


Private Sub AddLocalVar(strVarNameX As String, strValue As String, Optional arrayElement As Integer = 0, Optional tvType As Integer = 0)
    varCount = varCount + 1
    ReDim Preserve variables(1 To varCount) As typVariable
    variables(varCount).name = strVarNameX
    ReDim variables(varCount).Value(arrayElement) As String
    variables(varCount).Value(arrayElement) = strValue
    variables(varCount).type = LOCAL_
End Sub

Private Sub CleanUp()
    On Error Resume Next
    
    ReDim stack_calls(0) As typAliasCall
    stack_callcount = 0
    ReDim stack_calls(0).Args(0)
    stack_calls(0).ArgCount = 0
    
End Sub

Public Sub CopyAlias(ByRef oldAliasClass As cAlias, ByRef newAliasClass As cAlias)
    Set newAliasClass = oldAliasClass
End Sub

Private Sub CopyToArgs(paramlist())
    If UBound(paramlist) = 0 Then
        ReDim Args(0) As String
        ArgCount = 0
        Exit Sub
    End If
    
    ReDim Args(UBound(paramlist)) As String
    ArgCount = UBound(paramlist) + 1
    
    Dim i As Integer
    For i = LBound(paramlist) To UBound(paramlist)
        Args(i) = paramlist(i)
        If isStop = True Then Exit For
    Next i
End Sub

Public Sub dev_evalaliasx()
    Dim i As Integer, strInfo As String, strAliasType, paramlist(), strParamList() As String
    
    i = 1
    If AliasType = at_ALIAS Then strAliasType = "Sub"
    If AliasType = at_EVENT Then strAliasType = "Event"
    strInfo = "Sub Name:   " & strName & vbCrLf & _
              "Sub Type:   " & strAliasType & vbCrLf & _
              "Extra Params: " & strExtraParams & vbCrLf & vbCrLf
    For i = 1 To intCodeLines
        strInfo = strInfo & i & ":  " & strCode(i) & vbCrLf
        If isStop = True Then Exit For
    Next i
    
    'MsgBox strInfo
    strParamList = Split(strExtraParams, " ")
    If UBound(strParamList) > 0 Then
        ReDim paramlist(UBound(strParamList))
        For i = LBound(strParamList) To UBound(strParamList)
            paramlist(i) = strParamList(i)
            If isStop = True Then Exit For
        Next i
    End If
        
    rootEngine.ExecuteAlias strName, paramlist
End Sub

Public Function Execute(paramlist()) As String
    Dim bFinished As Boolean
    currentLine = 1
    bGotoNextLine = True
    bInComment = False
    
    CopyToArgs paramlist
    ReDim stack_while(0) As typWhileLoop
        
    Do
        If isStop = True Then Exit Do
        Call CleanUp
        If currentLine = 0 Then
            Execute = ""
            Exit Function
        End If
        returnValue = ExecuteLine(strCode(currentLine))
                
        If bGotoNextLine Then
            currentLine = currentLine + 1
        End If
    Loop Until currentLine > intCodeLines Or bGotoNextLine = False
    'MsgBox returnValue & "~~"
    Execute = returnValue
End Function

Public Function ExecuteLine(strLine As String)
    '*
    '* This is the heart of SEX
    '*

    'On Error Resume Next

    Dim strBuffer As String, i As Integer, curChar As String, prevChar As String
    Dim inVariable As Boolean, bEscape As Boolean, strLen As String, bWhiteSpace As Boolean
    strLen = Len(strLine)
    
    '* Clean stacks
    CleanUp
    
    i = 1
    
    If stack_whilecount <= UBound(stack_while) Then
        If stack_while(stack_whilecount).bIsFalse And Left(strLine, 3) <> "end" Then
            Exit Function
        End If
    End If
    
    Do Until i > strLen
        If isStop = True Then Exit Do
        curChar = Mid(strLine, i, 1)
        
        If prevChar <> ESCAPE_CHAR Then
            If curChar = COMMENT_ML_CHAR Then
                bWhiteSpace = False
                bInComment = Not bInComment
            End If
            If bInComment Or curChar = COMMENT_ML_CHAR Then GoTo nextchar
        End If
        
        '* Append escape char
        If bEscape Then
            bWhiteSpace = False
            'if invar, end var
            If inVariable Then
                '* change this...
                stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                strBuffer = ""
                inVariable = False
            End If
            Dim strEscapeChar As String
            
            strEscapeChar = GetEscapeChar(curChar)
            
            stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & strEscapeChar
            bEscape = False
            
        '* Comments, end
        ElseIf curChar = COMMENT_CHAR Then
            bWhiteSpace = False
            '* if invar, end var
            If inVariable Then
                '* change this...
                stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & GetLocalVar(strBuffer)

                strBuffer = ""
                inVariable = False
            End If
            
            GoTo finish
        '* Other stuff..parse
        Else
        
        Select Case curChar
            Case ESCAPE_CHAR
                bWhiteSpace = False
                '* if invar, end var
                If inVariable Then
                    '* change this...
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                    strBuffer = ""
                    inVariable = False
                End If
                
                If bEscape Then
                    '* get var, add escape
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & "\"
                
                End If
                
                bEscape = True
            Case BEGIN_FUNCTION
                bWhiteSpace = False
                '* if invar, end var
                If inVariable Then
                    '* change this...
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                    strBuffer = ""
                    inVariable = False
                End If
            
                stack_callcount = stack_callcount + 1
                ReDim Preserve stack_calls(stack_callcount) As typAliasCall
                ReDim stack_calls(stack_callcount).Args(0) As Variant
                stack_calls(stack_callcount).ArgCount = 0
                
            Case END_FUNCTION
                bWhiteSpace = False
                '* if invar, end var
                If inVariable Then
                    '* change this...
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                        
                    strBuffer = ""
                    inVariable = False
                End If
                
                Dim strReturnVal As String
                If stack_callcount = 0 Then GoTo nextchar
                
                If Left(CStr(stack_calls(stack_callcount).Args(0)), 1) = "$" Then
                    If stack_calls(stack_callcount).ArgCount < 1 Then
                        strReturnVal = ""
                        GoTo skipFunction
                    End If
                    
                    strReturnVal = GetLocalVar(Mid(CStr(stack_calls(stack_callcount).Args(0)), 2), CInt(stack_calls(stack_callcount).Args(1)))
                    GoTo skipFunction
                End If
                
                            
                Set rootEngine.cChildAlias = Me
                strReturnVal = rootEngine.ExecuteAlias(CStr(stack_calls(stack_callcount).Args(0)), stack_calls(stack_callcount).Args)
skipFunction:
                stack_calls(stack_callcount).ArgCount = 0
                stack_callcount = stack_callcount - 1
                stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & strReturnVal
                
            Case PARAM_DELIM
                If bWhiteSpace = True Then GoTo nextchar
                
                '* if invar, end var
                If inVariable Then
                    '* change this...
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                    
                    strBuffer = ""
                    inVariable = False
                End If
                
                If stack_calls(stack_callcount).bQuote Then
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & PARAM_DELIM
                Else
                    stack_calls(stack_callcount).ArgCount = stack_calls(stack_callcount).ArgCount + 1
                    ReDim Preserve stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount)
                End If
            Case """"
                bWhiteSpace = False
                
                stack_calls(stack_callcount).bQuote = _
                    Not stack_calls(stack_callcount).bQuote

            Case BEGIN_VARLOCAL
                bWhiteSpace = False
                If prevChar = BEGIN_FUNCTION Then
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & "$"
                    GoTo nextchar
                End If
                
                '* if invar, end var
                If inVariable Then
                    '* change this...
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                    
                    strBuffer = ""
                    inVariable = False
                End If
                variableType = 1
                inVariable = True
                strBuffer = ""
            Case BEGIN_VARGLOBAL
                bWhiteSpace = False
                variableType = 2
                inVariable = True
                strBuffer = ""
            Case Else
                bWhiteSpace = False
                If inVariable Then
                    '* In variable, append to strbuffer
                    Select Case curChar
                        Case "?", ",", ".", "/", "!", "@"
                            '* change this...
                            stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                                stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)
                            
                            strBuffer = ""
                            inVariable = False

                            stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                                stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & curChar
                        
                        Case "}", ">"
                            strBuffer = strBuffer & curChar
                            '* change this...
                            stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                                stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)

                            strBuffer = ""
                            inVariable = False
                        Case Else
                            strBuffer = strBuffer & curChar
                    End Select
                Else
                    '* Else append to other shit..
                    stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = _
                        stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & curChar
                    
                End If
        End Select
                        
        End If
nextchar:
    
        prevChar = curChar
        i = i + 1
    Loop
    
    If strBuffer <> "" Then
        If inVariable Then
            '* change this...
            stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) = stack_calls(stack_callcount).Args(stack_calls(stack_callcount).ArgCount) & GetVar(strBuffer)

            strBuffer = ""
            inVariable = False
        End If
    End If
    
finish:

    Dim strReturn As String
    
    
    If LCase(stack_calls(0).Args(0)) = "return" Then
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
        '*
        '* return statement
        '*
        ExecuteLine = JoinArrayV(stack_calls(0).Args, " ", 2)
        bGotoNextLine = False
        Exit Function
    ElseIf LCase(stack_calls(0).Args(0)) = "set" Then
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
        '*
        '* set a variable
        '*
        If UBound(stack_calls(0).Args) < 2 Then Exit Function
        
        If CStr(stack_calls(0).Args(1)) Like "*:*" Then
            SetLocalVar CStr(Mid(CStr(stack_calls(0).Args(1)), 1, InStr(CStr(stack_calls(0).Args(1)), ":") - 1)), JoinArrayV(stack_calls(0).Args, " ", 3), Mid(CStr(stack_calls(0).Args(1)), InStr(CStr(stack_calls(0).Args(1)), ":") + 1)
        Else
            SetLocalVar CStr(stack_calls(0).Args(1)), JoinArrayV(stack_calls(0).Args, " ", 3)
        End If
    
    '*
    '* WHILE LOOP
    '*
    ElseIf LCase(stack_calls(0).Args(0)) = "while" Then
        If stack_whilecount <= UBound(stack_while) Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
        
        If UBound(stack_calls(0).Args) < 1 Then Exit Function
        
        stack_whilecount = stack_whilecount + 1
        If stack_whilecount > UBound(stack_while) Then ReDim Preserve stack_while(stack_whilecount) As typWhileLoop
        
        If stack_while(stack_whilecount).bInit = False Then
            stack_while(stack_whilecount).iReturnLine = currentLine
            stack_while(stack_whilecount).bInit = True
            stack_while(stack_whilecount).bIsFalse = False
        End If
        
        Dim whileStat
        whileStat = Eval(JoinArrayV(stack_calls(0).Args, " ", 2))
        
        If whileStat <> 0 Then
            stack_while(stack_whilecount).bIsFalse = False
            stack_while(stack_whilecount).iReturnLine = currentLine
        Else
            stack_while(stack_whilecount).bIsFalse = True
            If stack_while(stack_whilecount).iLastLine <> -1 Then
                currentLine = stack_while(stack_whilecount).iLastLine
                Exit Function
            End If
        End If
        
    ElseIf LCase(stack_calls(0).Args(0)) = "loop" Then
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
        
        If UBound(stack_calls(0).Args) < 2 Then Exit Function
        
        stack_loopcount = stack_loopcount + 1
        ReDim Preserve stack_loop(stack_loopcount) As typLoop
    
        stack_loop(stack_loopcount).current = 1
        stack_loop(stack_loopcount).total = stack_calls(0).Args(2)
        stack_loop(stack_loopcount).linenum = currentLine '+ 1
        stack_loop(stack_loopcount).varname = stack_calls(0).Args(1)
        SetLocalVar stack_loop(stack_loopcount).varname, "1"

    ElseIf LCase(stack_calls(0).Args(0)) = "if" Then
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        If UBound(stack_calls(0).Args) < 1 Then Exit Function
        
        stack_ifcount = stack_ifcount + 1
        ReDim Preserve stack_if(stack_ifcount) As typIf
        
        Dim ifStat
        ifStat = Eval(JoinArrayV(stack_calls(0).Args, " ", 2))
       
        If ifStat = 0 Then
            stack_if(stack_ifcount).bTrueYet = False
        Else
            stack_if(stack_ifcount).bTrueYet = True
        End If
    ElseIf LCase(stack_calls(0).Args(0)) = "elseif" Then
        '* if in while loop, damnit!
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        
        If stack_if(stack_ifcount).bNoEval = True Then Exit Function
        If stack_if(stack_ifcount).bTrueYet = True Then
            stack_if(stack_ifcount).bTrueYet = False
            stack_if(stack_ifcount).bNoEval = True
            Exit Function
        End If
        
        If UBound(stack_calls(0).Args) < 1 Then Exit Function
        
        Dim ifStat2
        ifStat2 = Eval(JoinArrayV(stack_calls(0).Args, " ", 2))
        
        If ifStat2 = 0 Then
            stack_if(stack_ifcount).bTrueYet = False
        Else
            stack_if(stack_ifcount).bTrueYet = True
        End If
    ElseIf LCase(stack_calls(0).Args(0)) = "else" Then
        '* if in while loop, damnit!
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        
        If stack_if(stack_ifcount).bNoEval = True Then Exit Function
        If stack_if(stack_ifcount).bTrueYet = True Then
            stack_if(stack_ifcount).bTrueYet = False
            stack_if(stack_ifcount).bNoEval = True
            Exit Function
        End If
        
        stack_if(stack_ifcount).bTrueYet = True
    '*
    '* init variables
    '*
    ElseIf LCase(stack_calls(0).Args(0)) = "init" Then
        Dim iLoop As Integer
        For iLoop = 1 To stack_calls(0).ArgCount - 1
            AddLocalVar CStr(stack_calls(0).Args(iLoop)), "0"
        Next iLoop
    '*
    '* END
    '*
    ElseIf stack_calls(0).Args(0) = "end" Then
    
        If stack_calls(0).ArgCount < 1 Then Exit Function
        
        Select Case LCase(stack_calls(0).Args(1))
            Case "loop"
                If stack_ifcount > 0 Then
                    If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
                End If

                stack_loop(stack_loopcount).current = stack_loop(stack_loopcount).current + 1
                If stack_loop(stack_loopcount).current > stack_loop(stack_loopcount).total Then
                    stack_loopcount = stack_loopcount - 1
                Else
                    currentLine = stack_loop(stack_loopcount).linenum
                End If
                SetLocalVar stack_loop(stack_loopcount).varname, CStr(stack_loop(stack_loopcount).current)

            Case "if"
                stack_ifcount = stack_ifcount - 1
            Case "while"
                If stack_while(stack_whilecount).bIsFalse Then
                    stack_while(stack_whilecount).bInit = False
                    stack_while(stack_whilecount).bIsFalse = False
                    stack_whilecount = stack_whilecount - 1
                    stack_while(stack_whilecount).iLastLine = -1
                Else
                    stack_whilecount = stack_whilecount - 1
                    currentLine = stack_while(stack_whilecount + 1).iReturnLine - 1
                    stack_while(stack_whilecount + 1).iLastLine = currentLine
                End If
        End Select
    Else
        If stack_ifcount > 0 Then
            If stack_if(stack_ifcount).bTrueYet = False Then Exit Function
        End If
        
        '* if in while loop, damnit!
        If UBound(stack_while) >= stack_whilecount Then
            If stack_while(stack_whilecount).bIsFalse Then Exit Function
        End If
        
        '*
        '* nothing special, call the alias
        '*
        strReturn = rootEngine.ExecuteAlias(CStr(stack_calls(0).Args(0)), stack_calls(0).Args)
    End If
    
    ExecuteLine = ""

End Function


Private Function GetEscapeChar(strChar As String) As String
    Select Case strChar
        Case "n"    'new line
            GetEscapeChar = vbCrLf
        Case "t"    'tab
            GetEscapeChar = Chr(9)
        Case "0"    'char 0
            GetEscapeChar = Chr(0)
        Case "1"    'char 1
            GetEscapeChar = Chr(1)
        Case "c"    'color char
            GetEscapeChar = Chr(3)
        Case "r"    'reverse char
            GetEscapeChar = Chr(22)
        Case "b"    'bold char
            GetEscapeChar = Chr(2)
        Case "u"    'underline char
            GetEscapeChar = Chr(31)
        Case "s"    'Quota
            GetEscapeChar = Chr(34)
        Case "p"    'Prosen
            GetEscapeChar = Chr(37)
        Case Else
            GetEscapeChar = strChar
    End Select
End Function

Public Function GetLocalVar(strVarName As String, Optional arrayElement As Integer = 0) As String
    Dim i As Integer
    For i = 1 To varCount
        If isStop = True Then Exit For
        'MsgBox strVarName & "~" & variables(i).Name
        If variables(i).name = strVarName Then
            If arrayElement > UBound(variables(i).Value) Then
                GetLocalVar = ""
            Else
                GetLocalVar = variables(i).Value(arrayElement)
            End If
            Exit Function
        End If
    Next i
    GetLocalVar = ""
End Function

Public Function GetName() As String
    GetName = strName
End Function

Private Function GetVar(strVarName As String) As String
    If (Left(strVarName, 1) = "{" And Right(strVarName, 1) = "}") Or (Left(strVarName, 1) = "<" And Right(strVarName, 1) = ">") Then
        strVarName = Mid(strVarName, 2, Len(strVarName) - 2)
    End If
    
    
    If Right(strVarName, 1) = "-" Then
        If IsNumeric(Left(strVarName, Len(strVarName) - 1)) Then
            GetVar = JoinArray(Args, " ", Left(strVarName, Val(Len(strVarName) - 1)) + 1)
            Exit Function
        End If
    End If
    
    If IsNumeric(strVarName) Then
        If Val(strVarName) > ArgCount Then
            GetVar = ""
        Else
            GetVar = Args(strVarName)
        End If
        Exit Function
    End If
    
    
    If variableType = 1 Then
        GetVar = GetLocalVar(strVarName)
    Else
        '*get global var
        
        ' finish code
        GetVar = rootEngine.GetGlobalVar(strVarName)
    End If
End Function

Public Sub SetInfo(strtheName As String, at As Integer, strtheExtraParams As String)
    strName = strtheName
    AliasType = at
    strExtraParams = strtheExtraParams
End Sub


Public Sub SetLocalVar(strVarName As String, strValue As String, Optional arrayElement As Integer = 0)
    Dim i As Integer
    
    For i = 1 To varCount
        If isStop = True Then Exit For
        If variables(i).name = strVarName Then
            If arrayElement > UBound(variables(i).Value) Then ReDim Preserve variables(i).Value(arrayElement) As String
            variables(i).Value(arrayElement) = strValue
            Exit Sub
        End If
    Next i
    
    AddLocalVar strVarName, strValue, arrayElement
End Sub





VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private scripts()        As cScript
Private script_count     As Integer

'* Global variables
Private variables()     As typVariable
Private varCount        As Integer
Private Enum enumVarType
    xString
    xInteger
    xReal
End Enum

Private Type typVariable
    name    As String
    value   As String
    type    As enumVarType
End Type

Public cChildAlias  As New cAlias
Const vLibName = "Interaction"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Function onProperties(Optional PropertyName As String = "caption") As String
    Select Case LCase(PropertyName)
        Case "name"
            onProperties = vLibName
        Case "caption"
            onProperties = "Interaction Library"
        Case "description"
            onProperties = "Interaction Library 1.0"
        Case "keyword"
            onProperties = " calc clear cmp dec doevents eval input inc instr msgbox mypath printf set sendkeys ticks sleep "
    End Select
End Function

Function GetParamName(nName As String) As String
    On Error GoTo Salah
    Select Case LCase(Trim(nName))
        Case "msgbox"
             GetParamName = "MsgBox {message} [button] [title]" & Chr(0) & "Show message if success"
        Case "input"
             GetParamName = "Input {message} [title] [default value]" & Chr(0) & "Show input dialog if success"
        Case "ticks"
             GetParamName = "Ticks" & Chr(0) & "get current tick time"
        Case "set"
             GetParamName = "Set {variable} {value}" & Chr(0) & "Set new variable"
        Case "instr"
             GetParamName = "Instr {text1} {text2} {startpos}" & Chr(0) & "the position of the first occurrence of one string within another"
        Case "inc"
             GetParamName = "Inc {value1} [value2]" & Chr(0) & "Increase value"
        Case "eval"
             GetParamName = "Eval {value math}" & Chr(0) & "Formula eval (like: 1+2/4-32)"
        Case "doevents"
             GetParamName = "DoEvents" & Chr(0) & "Yields execution so that the operating system can process other events"
        Case "dec"
             GetParamName = "Dec {value1} [value2]" & Chr(0) & "Decrease value"
        Case "calc"
             GetParamName = "Calc {param value}" & Chr(0) & "Calculation formula"
        Case "printf"
             GetParamName = "Printf {param value}" & Chr(0) & "Print to output debugger"
        Case "sleep"
             GetParamName = "Sleep {param value}" & Chr(0) & "Sleep view seconds"
        Case "clear"
             GetParamName = "Clear {param value}" & Chr(0) & "Clear output debugger"
        Case "cmp"
             GetParamName = "Cmp {param value1} {param value2}" & Chr(0) & "Compare Text"
        Case "sendkeys"
             GetParamName = "SendKeys {param value}" & Chr(0) & "Sends keystroke(s) to the active window"
        Case "mypath"
             GetParamName = "MyPath" & Chr(0) & "Specifies the path of the Gateway Machine application file or the path of the executable file when running the application executable file."
    End Select
      
    Exit Function
Salah:
End Function

Function function_printf(paramlist())
    On Error Resume Next
    '
End Function

Function function_clear(paramlist())
    On Error Resume Next
  '
End Function

Function function_sleep(paramlist())
    On Error Resume Next
    'Sleep Val(paramlist(1))
    Sleep paramlist(1)
End Function

Function function_calc(paramlist())
    If UBound(paramlist) < 1 Then function_calc = -1: Exit Function
    
    Dim strEval As String
    strEval = Eval(JoinArrayV(paramlist, " ", 2))
    
    function_calc = strEval
End Function

Function function_dec(paramlist())
    If UBound(paramlist) < 1 Then function_dec = -1: Exit Function
    
    'MsgBox Val(paramlist(1)) & "~" & Val(paramlist(2))
    If UBound(paramlist) >= 2 Then
        function_dec = Val(paramlist(1)) - Val(paramlist(2))
    Else
        function_dec = Val(paramlist(1)) - 1
    End If
End Function

Function function_doevents(paramlist())
    DoEvents
End Function

Function function_eval(paramlist())
    If UBound(paramlist) < 1 Then function_eval = -1: Exit Function
    
    Dim strEval As String
    strEval = Eval(JoinArrayV(paramlist, " ", 2))
    
    function_eval = strEval
End Function

Function function_inc(paramlist())
    If UBound(paramlist) < 1 Then function_inc = -1: Exit Function
    
    If UBound(paramlist) >= 2 Then
        function_inc = Val(paramlist(1)) + Val(paramlist(2))
    Else
        function_inc = Val(paramlist(1)) + 1
    End If
End Function

Function function_instr(paramlist())
    If UBound(paramlist) < 2 Then function_instr = -1: Exit Function
    
    If UBound(paramlist) = 2 Then
        function_instr = InStr(1, paramlist(1), paramlist(2))
    Else
        function_instr = InStr(Val(paramlist(3)), paramlist(1), paramlist(2))
    End If
End Function

Function function_set(paramlist())
    If UBound(paramlist) < 2 Then function_set = -1: Exit Function
    
    'MsgBox paramlist(1) & "~fucker"
    If paramlist(1) Like "*:*" Then
        cChildAlias.SetLocalVar CStr(Mid(paramlist(1), 1, InStr(paramlist(1), ":") - 1)), JoinArrayV(paramlist, " ", 3), Mid(paramlist(1), InStr(paramlist(1), ":") + 1)
    End If
    
    cChildAlias.SetLocalVar CStr(paramlist(1)), JoinArrayV(paramlist, " ", 3)
End Function

Function function_ticks(paramlist())
    function_ticks = Timer
End Function

Public Function GetGlobalVar(strName As String) As String
    Dim i As Integer, GetLocalVar
    For i = 1 To varCount
        If variables(i).name = strName Then
            GetLocalVar = variables(i).value
            Exit Function
        End If
        If isStop = True Then Exit For
    Next i
    GetLocalVar = ""
End Function

Private Sub AddGlobalVar(strName As String, strValue As String, Optional tvType As Integer = 0)
    varCount = varCount + 1
    ReDim Preserve variables(1 To varCount) As typVariable
    variables(varCount).name = strName
    variables(varCount).value = strValue
    variables(varCount).type = tvType
End Sub

Private Sub SetGlobalVar(strName As String, strValue As String)
    Dim i As Integer
    
    For i = 1 To varCount
        If variables(i).name = strName Then
            variables(i).value = strValue
            Exit Sub
        End If
        If isStop = True Then Exit For
    Next i
    AddGlobalVar strName, strValue
End Sub

Public Sub dev_evalscript(script_id As Integer)
    scripts(script_id).dev_evalaliases
End Sub

Public Function ExecuteAlias(strAliasNameX As String, strParamList())
'    If isStop = True Then Exit Function
'
'    On Error GoTo NotBuiltIn
'    Dim returnValue, splFunc() As String
'    splFunc() = Split(strAliasNameX, ".")
'    Select Case UBound(splFunc)
'           Case 0
'                returnValue = CallByName(Me, "function_" & LCase(strAliasNameX), VbMethod, strParamList)
'           Case Else
'                Select Case LCase(splFunc(0))
'                       Case "register"
'                            Dim mysms As New ANVIBI_RegEntry
'                            returnValue = CallByName(mysms, "reg_" & splFunc(1), VbMethod, strParamList)
'                       Case Else
'                            If isPluginAvailable(splFunc(0)) Then
'                               returnValue = CallByName(modplugin(splFunc(0)), "function_" & splFunc(1), VbMethod, strParamList)
'                            Else
'                            End If
'                End Select
'    End Select
'    ExecuteAlias = CStr(returnValue)
'    Exit Function
'
'NotBuiltIn:
'    Dim i As Integer, j As Integer, returnVal As String
'    For i = 1 To script_count
'
'        returnValue = scripts(i).ExecuteAlias(strAliasNameX, strParamList)
'        If scripts(i).bExecuted = False Then
'            '* **
'            '* code to generate error??
'            '* **
'
'            ExecuteAlias = ""
'            Exit Function
'        End If
'        ExecuteAlias = returnValue
'        If isStop = True Then Exit For
'    Next i
End Function

Function function_input(paramlist())
    If UBound(paramlist) < 1 Then function_input = -1: Exit Function
    
    If UBound(paramlist) = 1 Then
        function_input = InputBox(paramlist(1))
    ElseIf UBound(paramlist) = 2 Then
        function_input = InputBox(paramlist(1), paramlist(2))
    Else
        function_input = InputBox(paramlist(1), paramlist(2), paramlist(3))
    End If
End Function

Function function_msgbox(paramlist())
    If UBound(paramlist) < 1 Then function_msgbox = -1: Exit Function
    
    If UBound(paramlist) = 1 Then
        function_msgbox = MsgBox(paramlist(1))
    ElseIf UBound(paramlist) = 2 Then
        function_msgbox = MsgBox(paramlist(1), 0, paramlist(2))
    Else
        function_msgbox = MsgBox(paramlist(1), paramlist(3), paramlist(2))
    End If
End Function

Function function_mypath(paramlist())
    On Error Resume Next
    function_mypath = App.path
End Function

Public Sub LoadScript(script_id As Integer, strFileName As String)
    scripts(script_id).LoadScript strFileName, Me
End Sub

Public Sub NewScript()
    script_count = script_count + 1
    ReDim Preserve scripts(1 To script_count) As cScript
    Set scripts(script_count) = New cScript
    Set scripts(script_count).parent = Me
End Sub

Public Function ScriptCount() As Integer
    ScriptCount = script_count
End Function

Function function_cmp(paramlist())
    On Error GoTo Salah
    If StrComp((paramlist(1)), (paramlist(2)), vbTextCompare) = 0 Then
       function_cmp = -1
    Else
       function_cmp = 0
    End If
    Exit Function
Salah:
    function_cmp = "0"
End Function

Function function_sendkeys(paramlist())
    On Error GoTo Salah
    SendKeys paramlist(1), True
    function_sendkeys = True
    Exit Function
Salah:
    function_sendkeys = False
End Function




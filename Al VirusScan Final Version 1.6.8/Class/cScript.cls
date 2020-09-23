VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private alias()         As cAlias
Private alias_count     As Integer
Public parent   As cMain

Public returnVal As String
Public bExecuted As Boolean
Private Sub ClearScript()
    ReDim Functions(1 To 1) As cAlias
End Sub

Public Sub dev_evalaliases()
    Dim i As Integer
   
    For i = 1 To alias_count
        If isStop = True Then Exit For
        alias(i).dev_evalaliasx
        
    Next i
End Sub

Public Function ExecuteAlias(strName As String, Params()) As String
    Dim i As Integer
    
    For i = 1 To alias_count
        If isStop = True Then Exit For
        If alias(i).GetName = strName Then
            Dim tempAlias As New cAlias
            alias(i).CopyAlias alias(i), tempAlias
            'ExecuteAlias = alias(i).Execute(Params)
            ExecuteAlias = tempAlias.Execute(Params)
            bExecuted = True
            returnVal = ""
            Exit Function
        End If
    Next i
    bExecuted = False
End Function

Public Sub LoadScript(strFileName As String, parent As cMain)
'    '* If file doesnt exist, exit
'    If FileExists(strFileName) = False Then
'        '* ...? echo...
'        Exit Sub
'    End If
'
'    '* Variable declarations
'    Dim FF As Integer, strLine As String, lineCount As Integer
'    Dim inAlias As Boolean, strAlias() As String
'
'    On Error GoTo error_handler
'    FF = FreeFile
'    Open strFileName For Input As #FF
'        Do
'            If isStop = True Then Exit Do
'            Line Input #FF, strLine
'            strLine = (TrimLeft(strLine))
'            If strLine = "" Then GoTo ignoreline
'            '* Now lets check the line for an alias or event tag
'            'strLine = TrimLeft(LCase(strLine))
'            'strLine = LCase(strLine)
'            If LCase(strLine) Like "sub *" Or LCase(strLine) Like "event *" Then
'                If inAlias Then
'                Else
'                    inAlias = True
'                    strAlias = Split(LCase(strLine), " ")
'                    If LCase(strAlias(0)) = "sub" Then
'                        NewAlias strAlias(1), 0, JoinArray(strAlias, " ", 3)
'                    ElseIf LCase(strAlias(0)) = "event" Then
'                        NewAlias strAlias(1), 1, JoinArray(strAlias, " ", 3)
'                    End If
'                    Set alias(alias_count).rootEngine = parent
'                    lineCount = 0
'                End If
'            ElseIf LCase(strLine) = "end sub" Or LCase(strLine) = "end event" Then
'                inAlias = False
'            ElseIf Left(strLine, 1) = "'" Then
'            Else
'                If inAlias Then
'                    alias(alias_count).AddCodeLine strLine
'                End If
'            End If
'ignoreline:
'        Loop Until EOF(1)
'    Close #FF
'    Exit Sub
'error_handler:
'    '* echo an error message
End Sub

Private Sub NewAlias(strName As String, the_type As Integer, Optional strExtraParams As String)
    alias_count = alias_count + 1
    ReDim Preserve alias(1 To alias_count) As cAlias
    Set alias(alias_count) = New cAlias
    
    alias(alias_count).SetInfo strName, the_type, strExtraParams
End Sub

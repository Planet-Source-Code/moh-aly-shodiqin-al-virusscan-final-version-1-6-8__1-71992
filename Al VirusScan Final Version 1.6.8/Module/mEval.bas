Attribute VB_Name = "mEval"

Public Function Eval(expr As String)
    Dim value As Variant, operand As String
    Dim pos As Integer
    
    pos = 1

    Do Until pos > Len(expr)
        If isStop = True Then Exit Do
        Select Case LCase(Mid(expr, pos, 3))
            Case "not", "or ", "and", "xor", "eqv", "imp"
                operand = Mid(expr, pos, 3)
                pos = pos + 3
        End Select


    Select Case LCase(Mid(expr, pos, 1))
        Case " "
            pos = pos + 1
        Case "&", "+", "-", "*", "/", "\", "^"
            operand = Mid(expr, pos, 1)
            pos = pos + 1
        Case ">", "<", "=":

        Select Case LCase(Mid(expr, pos + 1, 1))
            Case "<", ">", "="
                operand = Mid(expr, pos, 2)
                pos = pos + 1
            Case Else
                operand = Mid(expr, pos, 1)
        End Select
    pos = pos + 1
    Case Else
        Select Case LCase(operand)
            Case "": value = Token(expr, pos)
            Case "&":   Eval = Eval & value
                        value = Token(expr, pos)
            Case "+":   Eval = Eval + value
                        value = Token(expr, pos)
            Case "-":   Eval = Eval + value
                        value = -Token(expr, pos)
            Case "^":   value = value ^ Token(expr, pos)
            Case "*":   value = value * Token(expr, pos)
            Case "/":   value = value / Token(expr, pos)
            Case "\":   value = value \ Token(expr, pos)
            Case "not": Eval = Eval + value
                        value = Not Token(expr, pos)
            Case "and": value = value And Token(expr, pos)
            Case "or ": value = value Or Token(expr, pos)
            Case "xor": value = value Xor Token(expr, pos)
            Case "eqv": value = value Eqv Token(expr, pos)
            Case "imp": value = value Imp Token(expr, pos)
            Case "=", "==": value = value = Token(expr, pos)
            Case ">": value = value > Token(expr, pos)
            Case "<": value = value < Token(expr, pos)
            Case ">=", "=>": value = value >= Token(expr, pos)
            Case "<=", "=<": value = value <= Token(expr, pos)
            Case "<>", "!=": value = value <> Token(expr, pos)
        End Select
    End Select
    Loop
    Eval = Eval + value
End Function


Function Token(expr, pos)
    Dim char As String, value As String, fn As String
    Dim es As Integer, pl As Integer
    Const QUOTE As String = """"
    
    On Error Resume Next

    Do Until pos > Len(expr)
        If isStop = True Then Exit Do
        char = Mid(expr, pos, 1)


        Select Case LCase(char)
            Case "&", "+", "-", "/", "\", "*", "^", " ", ">", "<", "=": Exit Do
            Case "("
            pl = 1
            pos = pos + 1
            es = pos


            Do Until pl = 0 Or pos > Len(expr)
                If isStop = True Then Exit Do
                char = Mid(expr, pos, 1)


                Select Case LCase(char)
                    Case "(": pl = pl + 1
                    Case ")": pl = pl - 1
                End Select
            pos = pos + 1
        Loop
        value = Mid(expr, es, pos - es - 1)
        fn = (Token)


        Select Case LCase(fn)
            Case "sin": Token = Sin(Eval(value))
            Case "cos": Token = Cos(Eval(value))
            Case "tan": Token = Tan(Eval(value))
            Case "exp": Token = Exp(Eval(value))
            Case "log": Token = Log(Eval(value))
            Case "atn": Token = Atn(Eval(value))
            Case "abs": Token = Abs(Eval(value))
            Case "sgn": Token = Sgn(Eval(value))
            Case "sqr": Token = Sqr(Eval(value))
            Case "rnd": Token = Rnd(Eval(value))
            Case "int": Token = Int(Eval(value))
            Case "day": Token = Day(Eval(value))
            Case "month": Token = Month(Eval(value))
            Case "year": Token = Year(Eval(value))
            Case "weekday": Token = Weekday(Eval(value))
            Case "hour": Token = Hour(Eval(value))
            Case "minute": Token = Minute(Eval(value))
            Case "second": Token = second(Eval(value))
            Case "date": Token = Date
            Case "date$": Token = Date$
            Case "time": Token = Time
            Case "time$": Token = Time$
            Case "timer": Token = Timer
            Case "now": Token = Now()
            'Case "len": Token = Len(Eval(value))
            'Case "trim": Token = Trim(Eval(value))
            'Case "ltrim": Token = LTrim(Eval(value))
            'Case "rtrim": Token = RTrim(Eval(value))
            'Case "ucase": Token = UCase(Eval(value))
            'Case "lcase": Token = LCase(Eval(value))
            Case "val": Token = Val(Eval(value))
            'Case "chr": Token = Chr(Eval(value))
            'Case "asc": Token = Asc(Eval(value))
            'Case "space": Token = Space(Eval(value))
            Case "hex": Token = Hex(Eval(value))
            Case "oct": Token = Oct(Eval(value))
            Case "environ": Token = Environ$(Eval(value))
            Case "curdir": Token = CurDir$
            Case "apppath": Token = App.path
            'Case "dir": If Len(value) Then Token = Dir(Eval(value)) Else Token = Dir
            Case Else: Token = Eval(value)
        End Select
    Exit Do
    Case QUOTE
    pl = 1
    pos = pos + 1
    es = pos


    Do Until pl = 0 Or pos > Len(expr)
        If isStop = True Then Exit Do
        char = Mid(expr, pos, 1)
        pos = pos + 1


        If char = QUOTE Then


            If Mid(expr, pos, 1) = QUOTE Then
                value = value & QUOTE
                pos = pos + 1
            Else
                Exit Do
            End If
        Else
            value = value & char
        End If
    Loop
    Token = value
    Exit Do
    Case Else
    Token = Token & char
    pos = pos + 1
End Select
Loop



If IsNumeric(Token) Then
Token = Val(Token)
ElseIf IsDate(Token) Then
Token = CDate(Token)
End If
End Function



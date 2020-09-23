Attribute VB_Name = "modSubs"
'*********************************************************************************************
'*              SEAL by M.Mason  2004(c)                                                     *
'*             djhappylobster@hotmail.com                                                    *
'*                                                                                           *
'*  This is not free code.  It is from demonstration purposes                                *
'*  You can use any part of this code for your own PERSONAL use                              *
'*  You may not redistribute this code on your webwsite or other media without permission    *
'*  You must not pass this off as your own                                                   *
'*                                                                                           *
'*  Please contact the author regarding any queries                                          *
'*                                                                                           *
'*  Enjoy try and write something useful with this app!!!!                                   *
'*********************************************************************************************

Option Explicit
Dim RegNumber1 As Integer

Sub ParseOperands(ByVal QuickCheck As Boolean)
       
    
    'Get next symbol
    GetSym False
    
    'look for correct reg or device
    RegNumber1 = -1
    Select Case TempRegDev
    Case 1
        'process command
        If Sym = "ACC" Then
            RegNumber1 = 1
        ElseIf Sym = "INDX" Then
            RegNumber1 = 2
        End If
    Case 2
        'input comand
        If Sym = "KBD" Then
            RegNumber1 = 3
        End If
    Case 3
        'output command
        If Sym = "SCR" Then
            RegNumber1 = 4
        End If
    End Select
    
    'check
    If RegNumber1 = -1 Then
        GetSym False
        If Sym = "COMMA" Then
            If TempRegDev = 1 Then
                ErrorMessages = "Invalid register"
            Else
                ErrorMessages = "Invalid device"
            End If
        Else
            ErrorMessages = "Syntax error"
        End If
        Exit Sub
    End If
    
    'look for comma
    GetSym False
    If Sym = "COMMA" Then
        GetSym False
        ParseAddress QuickCheck
    Else
        'No comma found
        ErrorMessages = "Expected comma"
    End If
    
End Sub
Sub ParseAddress(ByVal QuickCheck As Boolean)
    
    'error handler
    On Error GoTo error_handler
    
    Dim RegNumber2 As Integer
    Dim TempName As String
    Dim TempVal As Variant
    Dim Pos2 As Integer
    Dim TempText As String
    Dim Found As Boolean
    'look for array
    
    TempName = NonSym
    
    If Sym = "ALPHANUMERIC" Then
        
        TempVariable = NonSym
        GetSym False
        If Sym = "OPENBRACKET" Then
            'Found array
            
            'Get subscript valud
            GetSym False
            If Sym = "ACC" Then
                RegNumber2 = 1  'Accumulator
            ElseIf Sym = "INDX" Then
                RegNumber2 = 2  'Index
            ElseIf Sym = "NUMBER" Then
                'Get subscript
                TempVal = Val(NonSym)
                If TempVal >= 0 Then
                    If TempVal <= 32767 Then
                        RegNumber2 = TempVal + 4
                    Else
                        ErrorMessages = "Array index too large"
                        Exit Sub
                    End If
                Else
                    ErrorMessages = "Cannot have negative array indexes"
                    Exit Sub
                End If
'todo: inserted  line to do dry parse
'i.e. do explain line
                ExplainAddress = TempName + "(" + Format(TempVal) + ")"

            End If
            
            If RegNumber2 <> 0 Then
                GetSym False
                If Sym = "CLOSEBRACKET" Then
                    'Look for array name
                    If QuickCheck = False Then
                        If ArrayCount > 0 Then
                            For Pos2 = 0 To ArrayCount - 1
                                If TempName = ArrayName(Pos2) Then
                                    TempOperand1 = RegNumber1 'Set to register
                                    TempOperand2 = RegNumber2 'Set subscript register
                                    TempOperand3 = Pos2       'Set array no to use
                                    Found = True
                                End If
                            Next
                            If Found = False Then
                                ErrorMessages = "Variable not declared"
                            End If
                        Else
                            ErrorMessages = "Array not declared"
                        End If
                    Else
'todo: inserted else and line to do dry parse
'i.e. do explain line
                        If TempOperand2 = 1 Then
                            ExplainAddress = TempName + "(ACC)"
                        ElseIf TempOperand2 = 2 Then
                            ExplainAddress = TempName + "(INDX)"
                        End If
                    End If
                Else
                    ErrorMessages = "Expected closing bracket"
                End If
            Else
                ErrorMessages = "Invalid Subscript"
            End If
        'Check for variable
        ElseIf Sym = "NEWLINE" Or Sym = "END" Then
            If QuickCheck = False Then
                If VariableCount > 0 Then
                    For Pos2 = 0 To VariableCount - 1
                        If TempName = VariableName(Pos2) Then
                            TempOperand1 = RegNumber1   'To register
                            TempOperand2 = 0            'Null
                            TempOperand3 = Pos2         'Variable to use
                            Found = True
                        End If
                    Next
                    
                    If Found = False Then
                        ErrorMessages = "Variable not declared"
                    End If
                Else
                    ErrorMessages = "Variable not declared"
                End If
                
                'Take sym pos to before
                'end of line when not in direct mode
                If CodeState <> 1 Then
                    If Sym <> "END" Then
                        SymPos = SymPos - 2
                    End If
                End If
            Else
'todo: inserted else and line to do dry parse
'i.e. do explain line
                ExplainAddress = TempName
            End If
        Else
           ErrorMessages = "Syntax error"
        End If
    ElseIf Sym = "HASH" Then
        'Make sure we're not copying to immediate address
        If TempOperation <> 11 Then
            GetSym False
            If Sym = "NUMBER" Then
                If QuickCheck = False Then
                    'Range check
                    If IsOutOfRange(Val(NonSym)) = False Then
                        TempOperand1 = RegNumber1
                        TempOperand2 = 3
                        TempOperand3 = Val(NonSym)
                        
                        'Check if the colour value is value
                        Select Case TempOperation
                        Case 26, 28
                            If Val(NonSym) < 0 Or Val(NonSym) > 15 Then
                                ErrorMessages = "Invalid colour value"
                            End If
                        End Select
                    Else
                        ErrorMessages = "Immediate value is out of range"
                    End If
                Else
'todo: inserted else and line to do dry parse
'i.e. do explain line
                    ExplainAddress = "#" & NonSym
                End If
            Else
                'Are we expecting a colour value here?
                If TempOperation = 26 Then
                    ErrorMessages = "Expected valid colour value"
                Else
                    ErrorMessages = "Expected valid address"
                End If
            End If
        Else
            'User trying to assign register
            'to direct address
            ErrorMessages = "Syntax error"
        End If
        
    ElseIf Sym = "RND" Then
        'RND function
        
            
        'Test for command like copy acc,rnd(20)
        If TempOperation = 11 Then
            ErrorMessages = "Syntax error"
            Exit Sub
        End If

        GetSym False
        If Sym = "OPENBRACKET" Then
        
            GetSym False
            If Sym = "ACC" Then
                RegNumber2 = 1  'Accumulator
            ElseIf Sym = "INDX" Then
                RegNumber2 = 2  'Index
            ElseIf Sym = "NUMBER" Then
                'Get subscript
                TempVal = Val(NonSym)
                If TempVal >= 0 Then
                    If TempVal <= 32767 Then
'todo: inserted else and line to do dry parse
'i.e. do explain line
                    ExplainAddress = "RND(" + Format(TempVal) + ")"
                        
                        RegNumber2 = TempVal + 4
                    Else
                        ErrorMessages = "Argument too large"
                        Exit Sub
                    End If
                Else
                    ErrorMessages = "Cannot have negative arguments"
                    Exit Sub
                End If
            End If
        
            If RegNumber2 <> 0 Then
                GetSym False
                If Sym = "CLOSEBRACKET" Then
                    If QuickCheck = False Then
                        TempOperand1 = RegNumber1 'Set to register
                        TempOperand2 = RegNumber2 'Set subscript register
                        TempOperand3 = -1           'Flag that RND is used
                        Found = True
                    Else
'todo: inserted else and line to do dry parse
'i.e. do explain line
                        If TempOperand2 = 1 Then
                            ExplainAddress = "RND(ACC)"
                        ElseIf TempOperand2 = 2 Then
                            ExplainAddress = "RND(INDX)"
                        End If
                    End If
                Else
                    ErrorMessages = "Expected closing bracket"
                End If
            Else
                ErrorMessages = "Invalid argument"
                Exit Sub
            End If
        Else
            ErrorMessages = "Expected opening bracket"
            Exit Sub
        End If
         
    Else
        If TempOperation = 20 Then
            ErrorMessages = "Expected register or variable"
        Else
            ErrorMessages = "Expected valid address"
        End If
    End If
    
    Exit Sub
    
error_handler:
    ErrorMessages = "Compile error: " & Err.Description
End Sub
Sub ParseCommand(ByVal QuickCheck As Boolean)
    Dim RegNumber As Integer
    Dim TempText As String
    'Get keyword
    GetSym False
    
    Select Case Sym
    Case "ADD"
        'Add keyword
        TempOperation = 0
        TempRegDev = 1
        ParseOperands QuickCheck
    Case "SUB"
        TempOperation = 1
        TempRegDev = 1
        ParseOperands QuickCheck
    Case "MPY"
        TempOperation = 2
        TempRegDev = 1
        ParseOperands QuickCheck
    Case "DVD"
        TempOperation = 3
        TempRegDev = 1
        ParseOperands QuickCheck
    Case "MOD"
        TempOperation = 4
        TempRegDev = 1
        ParseOperands QuickCheck
    Case "NEG"
        GetSym False
        If Sym = "ACC" Then
            RegNumber = 1
        ElseIf Sym = "INDX" Then
            RegNumber = 2
        ElseIf Sym = "FLAG" Then
            RegNumber = 3
        Else
            ErrorMessages = "Invalid Register"
        End If
        TempOperation = 5
        TempOperand1 = RegNumber
    Case "CLRZ"
        GetSym False
        If Sym = "ACC" Then
            RegNumber = 1
        ElseIf Sym = "INDX" Then
            RegNumber = 2
        ElseIf Sym = "FLAG" Then
            RegNumber = 3
        Else
            ErrorMessages = "Invalid Register"
        End If
        TempOperation = 6
        TempOperand1 = RegNumber
    Case "INC"
        GetSym False
        If Sym = "ACC" Then
            RegNumber = 1
        ElseIf Sym = "INDX" Then
            RegNumber = 2
        Else
            ErrorMessages = "Invalid Register"
        End If
        
        TempOperation = 7
        TempOperand1 = RegNumber
    Case "DEC"
        GetSym False
        If Sym = "ACC" Then
            RegNumber = 1
        ElseIf Sym = "INDX" Then
            RegNumber = 2
        Else
            ErrorMessages = "Invalid Register"
        End If
        
        TempOperation = 8
        TempOperand1 = RegNumber
    Case "CMPR"
        TempOperation = 9
        TempRegDev = 1
        ParseOperands QuickCheck
    Case "LOAD"
        TempOperation = 10
        TempRegDev = 1
        ParseOperands QuickCheck
    Case "COPY"
        TempOperation = 11
        TempRegDev = 1
        TempRegDev = 1
        ParseOperands QuickCheck
    Case "JUMP"
        TempOperation = 12
        DoLabelCheck QuickCheck
    Case "JEQZ"
        TempOperation = 13
        DoLabelCheck QuickCheck
    Case "JLEZ"
        TempOperation = 14
        DoLabelCheck QuickCheck
    Case "JLTZ"
        TempOperation = 15
        DoLabelCheck QuickCheck
    Case "JGEZ"
        TempOperation = 16
        DoLabelCheck QuickCheck
    Case "JGTZ"
        TempOperation = 17
        DoLabelCheck QuickCheck
    Case "JSUBR"
        TempOperation = 18
        DoLabelCheck QuickCheck
    Case "EXIT"
        TempOperation = 19
    Case "HALT"
        TempOperation = 20
    Case "INPTI"
        TempOperation = 21
        TempRegDev = 2
        ParseOperands QuickCheck
    Case "OUPTI"
        'output command
        TempOperation = 22
        TempRegDev = 3
        ParseOperands QuickCheck
    Case "OUPTS"
        'Output string
        TempOperation = 23
        
        'Look for screen device
        GetSym False
        If Sym <> "SCR" Then
            ErrorMessages = "Invalid device"
            Exit Sub
        End If
        
        'look for comma
        GetSym False
        If Sym = "COMMA" Then
            
            'Look for comma
            GetSym False
            If Sym = "SINGLEQUOTE" Then
                GetSym False
                
                If Sym = "NUMBER" Or Sym = "ALPHANUMERIC" Or Sym = "STRING" Then
                    TempText = NormSym
                    GetSym False
                    If Sym = "SINGLEQUOTE" Then
                        If QuickCheck = False Then
                            TempOperandText = TempText
                        End If
        'todo: Inserted this line for dry explanation
                        ExplainAddress = TempText
                    Else
                        ErrorMessages = "Expected closing literal"
                    End If
                ElseIf Sym = "SINGLEQUOTE" Then 'No string between literals
                Else
                    ErrorMessages = "Expected valid string"
                End If
            Else
                ErrorMessages = "Expected opening literal"
            End If
        Else
           'No comma found
            ErrorMessages = "Expected comma"
        End If
    Case "QUESTION"
        GetSym False
        'look for reg
        TempOperation = 24
        If Sym = "ACC" Then
            TempOperand1 = 0
            TempOperand2 = 3
            TempOperand3 = Acc
        ElseIf Sym = "INDX" Then
            TempOperand1 = 0
            TempOperand2 = 3
            TempOperand3 = Indx
        ElseIf Sym = "FLAG" Then
            TempOperand1 = -1 '-1 so execute command calculates flag reg
            TempOperand2 = 3
            TempOperand3 = LastRegister
        Else
            ParseAddress False
        End If
    Case "LISTVARS"
        If CodeState = 1 Then
            TempOperation = 25
        Else
            TempOperation = -99
        End If
    Case "CLRS"
        TempOperation = 26
        'Look for a colour
        GetSym True
        If Sym <> "NEWLINE" And Sym <> "END" Then
            If Sym = "ACC" Then
                TempOperand1 = 0
                TempOperand2 = 3
                TempOperand3 = Acc
            ElseIf Sym = "INDX" Then
                TempOperand1 = 0
                TempOperand2 = 3
                TempOperand3 = Indx
            Else
                GetSym False
                ParseAddress QuickCheck
            End If
        Else
            'Flag there's no background color to set
            TempOperand1 = -1
        End If
    Case "JNEZ"
        TempOperation = 27
        DoLabelCheck QuickCheck
    Case "OUPTC"
        'output command
        TempOperation = 28
        TempRegDev = 3
        ParseOperands QuickCheck
    Case "NEWLINE"
        TempOperation = -13
        Exit Sub
    Case "END"
        TempOperation = -1
    Case Else
        TempOperation = -99
    End Select
End Sub

Sub GetSym(ByVal PreviewNextSymbol As Boolean)
    Dim Pos As Integer
    Dim Others As Boolean
    Dim Letters As Boolean
    Dim Digits As Boolean
    Dim SymLength As Integer
    Dim EOSpos As Integer
    
    'Get symbol name for group of charcters
    
    'Look for end of symbol char
    If SymPos >= Len(GroupedCode) Then
        Sym = "END"
        Exit Sub
    Else
        EOSpos = InStr(SymPos, GroupedCode, EOSchar)
    End If
    
    Sym = Mid(GroupedCode, SymPos, EOSpos - SymPos)
    SymLength = Len(Sym)
    
    'Do we wish to move the reader position?
    If PreviewNextSymbol = False Then
        SymPos = SymPos + SymLength + 1
    End If
    
    NormSym = Sym
    Sym = UCase(Sym)
    NonSym = Sym
     
    If SymLength > 1 Then
    
        If CodeState = 1 Then
            Select Case Sym
            Case "LISTVARS"
                'List variables
                NonSym = "KEYWORD"
            End Select
            If NonSym = "KEYWORD" Then Exit Sub
         End If
    
        'Look for keywords
        Select Case Sym
        Case "ADD"
            NonSym = "KEYWORD"
        Case "SUB"
            NonSym = "KEYWORD"
        Case "MPY"
            NonSym = "KEYWORD"
        Case "DVD"
            NonSym = "KEYWORD"
        Case "MOD"
            NonSym = "KEYWORD"
        Case "NEG"
            NonSym = "KEYWORD"
        Case "CLRZ"
            NonSym = "KEYWORD"
        Case "INC"
            NonSym = "KEYWORD"
        Case "DEC"
            NonSym = "KEYWORD"
        Case "CMPR"
            NonSym = "KEYWORD"
        Case "LOAD"
            NonSym = "KEYWORD"
        Case "COPY"
            NonSym = "KEYWORD"
        Case "JUMP"
            NonSym = "KEYWORD"
        Case "JEQZ"
            NonSym = "KEYWORD"
        Case "JNEZ"
            NonSym = "KEYWORD"
        Case "JLEZ"
            NonSym = "KEYWORD"
        Case "JLTZ"
            NonSym = "KEYWORD"
        Case "JGEZ"
            NonSym = "KEYWORD"
        Case "JGTZ"
            NonSym = "KEYWORD"
        Case "JSUBR"
            NonSym = "KEYWORD"
        Case "EXIT"
            NonSym = "KEYWORD"
        Case "HALT"
            NonSym = "KEYWORD"
        Case "INPTI"
            NonSym = "KEYWORD"
        Case "OUPTI"
            NonSym = "KEYWORD"
        Case "OUPTS"
            NonSym = "KEYWORD"
        Case "OUPTC"
            NonSym = "KEYWORD"
        Case "CLRS"
            NonSym = "KEYWORD"
        Case "BLOCKI"
        Case "DATAI"
        Case "ACC"
        Case "INDX"
        Case "FLAG"
        Case "KBD"
        Case "SCR"
        Case "RND"
        Case Else
            
            If Left(Sym, 1) = "-" Or Left(Sym, 1) = "+" Then
                'Check for digits and letters beginning with "-"
                For Pos = 2 To SymLength
                    Select Case Asc(Mid(Sym, Pos, 1))
                    '-------------------0-9------A-Z---------_--
                    Case 9, 32, 33 To 47, 58 To 64, 91 To 94, 96 To 126
                        Others = True
                    Case Asc("0") To Asc("9")
                        Digits = True
                    Case Asc("a") To Asc("z"), Asc("A") To Asc("Z"), Asc("_")
                        Letters = True
                    Case Else
                        Sym = "INVALID"
                        Exit Sub
                    End Select
                Next
                If Letters = True Then Others = True
                                
            Else
                'Check for digits and letters not beginning with "-"
                For Pos = 1 To SymLength
                    Select Case Asc(Mid(Sym, Pos, 1))
                    '-------------------0-9------A-Z---------_--
                    Case 9, 32, 33 To 47, 58 To 64, 91 To 94, 96 To 126
                        Others = True
                    Case Asc("0") To Asc("9")
                        Digits = True
                    Case Asc("a") To Asc("z"), Asc("A") To Asc("Z"), Asc("_")
                        Letters = True
                    Case Else
                        Sym = "INVALID"
                        Exit Sub
                    End Select
                Next
            End If

            'Determine what the word is
            If Letters = True Then
                If Others = True Then
                    Sym = "STRING"
                Else
                    Select Case Asc(Left(Sym, 1))
                    Case Asc("A") To Asc("Z")
                        Sym = "ALPHANUMERIC"
                    Case Else
                        Sym = "STRING"
                    End Select
                End If
            Else
                If Digits = True Then
                    If Others = True Then
                        Sym = "STRING"
                    Else
                        Sym = "NUMBER"
                    End If
                Else
                    If Others = True Then
                        Sym = "STRING"
                    Else
                        Sym = "INVALID"
                    End If
                End If
            End If

        End Select
    ElseIf SymLength = 1 Then
        'Single character
        Select Case Asc(Sym)
        Case Asc("(")
            Sym = "OPENBRACKET"
        Case Asc(")")
            Sym = "CLOSEBRACKET"
        Case Asc(",")
            Sym = "COMMA"
        Case Asc("#")
            Sym = "HASH"
        Case Asc(":")
            Sym = "COLON"
        Case Asc("'")
            Sym = "SINGLEQUOTE"
        Case Asc("0") To Asc("9")
            Sym = "NUMBER"
        Case Asc(EOLchar)
            Sym = "NEWLINE"
        Case Asc("?")
            Sym = "QUESTION"
        Case Asc("a") To Asc("z"), Asc("A") To Asc("Z")
            Sym = "ALPHANUMERIC"
        Case 33 To 47, 58 To 64, 91 To 94, 96 To 126
            Sym = "STRING"
        End Select
        
    Else
        'New line
        Sym = "NEWLINE"
    End If
    
End Sub

Sub DoLabelCheck(ByVal QuickCheck As Boolean)
    'Check if label exists
    Dim Pos As Integer
    Dim FoundLabel As Boolean
    
    GetSym False
    TempLabelName2 = Empty
    
    If Sym = "END" Or Sym = "NEWLINE" Then
        ErrorMessages = "Label not found"
    Else
        If QuickCheck = False Then
            For Pos = 0 To LabelCount - 1
                If NonSym = LabelName(Pos) Then
                    TempOperand1 = LabelPos(Pos)
                    FoundLabel = True
                End If
            Next
            If FoundLabel = False Then
                ErrorMessages = "Label not found"
            End If
        End If
        TempLabelName2 = NonSym
        ExplainLabelTo = NonSym
    End If
End Sub





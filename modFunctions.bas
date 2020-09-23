Attribute VB_Name = "modFunctions"
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

Dim intLastLineNumber As Integer

Public Function GetValue()
    'Returns value of opperand or variable
    'of current instruction
    Dim RegisterValue As Integer
   
   
    Select Case TempOperand2
    Case 0
        'Variable
       GetValue = VariableValue(TempOperand3)
       frmAnimate.lstVariables.ListIndex = TempOperand3
    Case 3
        'Direct address
        GetValue = TempOperand3
    Case Else
        'Indexed (array) or RND function
                    
        'check what register to use as subscript
        If TempOperand2 = 1 Then
            'Acc
            RegisterValue = Acc
        ElseIf TempOperand2 = 2 Then
            'Indx
            RegisterValue = Indx
        Else
            'Direct subscript
            RegisterValue = TempOperand2 - 4
        End If
        
        'RND flag set to -1
        If TempOperand3 = -1 Then
            
            'workout RND value
            If CodeState = 2 Then
                GetValue = TempRandomNumber 'Program running
            Else
                'Direct mode
                GetValue = Int(Rnd() * (RegisterValue + 1))
            End If
            Exit Function
            
        End If
        
        'check for out of range
        If RegisterValue > ArrayElements(TempOperand3) Or RegisterValue < 0 Then
            ErrorMessages = "Array index out of range"
        Else
            'range ok
            GetValue = ArrayValue(TempOperand3, RegisterValue)
            frmAnimate.lstVariables.ListIndex = VariableCount - 1 + TempOperand3
        End If
    End Select

End Function

Public Function Inc(ByRef Number)
    'Increment number
    Number = Number + 1
End Function

Public Function Dec(ByRef Number)
    'Decrement number
    Number = Number - 1
End Function

Public Sub SetValue(ByVal RegisterVal As Integer)
    'Sets variable to particular value
    Dim SubscriptVal As Integer
    
    Select Case TempOperand2
    Case 0
        'Variable
        VariableValue(TempOperand3) = RegisterVal
        
        'Highlight line in memory
        frmAnimate.lstVariables.ListIndex = TempOperand3
    Case Else
        'Indexed (array)
        If TempOperand2 = 1 Then
            SubscriptVal = Acc
        ElseIf TempOperand2 = 2 Then
            SubscriptVal = Indx
        Else
            SubscriptVal = TempOperand2 - 4
        End If
        
        'store val
       
        'check for out of range
        If SubscriptVal > ArrayElements(TempOperand3) Or SubscriptVal < 0 Then
            ErrorMessages = "Array index out of range"
        Else
            'range ok
             ArrayValue(TempOperand3, SubscriptVal) = RegisterVal
             
             'Highlight line in memory
             frmAnimate.lstVariables.ListIndex = VariableCount - 1 + TempOperand3
        End If


    End Select

End Sub


Public Function Replace(ByRef InString As String, ByVal FindString As String, ByVal ReplaceString As String)
    'Replaces text
    Dim LeftChunk As String
    Dim RightChunk As String
    Dim FoundPos As Long
    Dim Pos As Long
    
    Pos = 1
    Do
        FoundPos = InStr(Pos, InString, FindString)
        If FoundPos > 0 Then
            LeftChunk = Left(InString, FoundPos - 1)
            RightChunk = Right(InString, Len(InString) - FoundPos)
            InString = LeftChunk + ReplaceString + RightChunk
            Pos = FoundPos - Len(FindString) + Len(ReplaceString)
        Else
            Exit Do
        End If
    Loop
    Replace = InString
End Function

Public Function GroupCode(ByVal InputCode As String)
    Dim CurrentChar As Integer
    Dim Pos As Integer
    Dim StartPos As Integer
    Dim LastSymbolType As String
    Dim EndOfComment As Integer
    Dim LineNo As Integer
    
    'Goes through code and groups valid characters together
    'Ignores comments
    'Keeps track of new lines
    
    'No text input
    If InputCode = Empty Then Exit Function
    
    'InputCode = UCase(InputCode)
    Pos = 1
    Do
        'Get char at current position
        CurrentChar = Asc(Mid(InputCode, Pos, 1))
        
        Select Case CurrentChar
        
        'These symbols are all grouped together ------------------------
        Case Asc("a") To Asc("z"), Asc("A") To Asc("Z"), Asc("-"), Asc("+"), Asc("0") To Asc("9"), Asc("_")
            'Alphanumerics
            If LastSymbolType <> "ALPHANUMERIC" Then
                If LastSymbolType <> "SPACE" Then
                    GroupCode = GroupCode + EOSchar + Chr(CurrentChar)
                Else
                    GroupCode = GroupCode + Chr(CurrentChar)
                End If
            Else
                GroupCode = GroupCode + Chr(CurrentChar)
            End If
            LastSymbolType = "ALPHANUMERIC"
        
        'These symbols are grouped individually ------------------------
        Case Asc(":"), Asc(","), Asc("("), Asc(")"), Asc("#"), Asc("?"), Asc("-")
            'Punctuation symbols
            
            If LastSymbolType <> "SPACE" Then
                GroupCode = GroupCode + EOSchar + Chr(CurrentChar)
            Else
                GroupCode = GroupCode + Chr(CurrentChar)
            End If

            LastSymbolType = "PUNCTUATION"
        
        'Store valid text between literals as they are found ---------------
        Case Asc("'")
            
            'Place eoschar?
            If LastSymbolType <> "SPACE" Then
                GroupCode = GroupCode + EOSchar + Chr(CurrentChar) + EOSchar
            Else
                GroupCode = GroupCode + Chr(CurrentChar) + EOSchar
            End If
            
            StartPos = Pos
            
            'Go through text until end or quote is found
            Do
                If Pos = Len(InputCode) Then Exit Do
                
                Inc Pos
                
                'Get character
                Select Case Asc(Mid(InputCode, Pos, 1))
                Case Asc("'")
                    'Found quote, place EOSchar if required
                    If Pos - StartPos > 1 Then
                        GroupCode = GroupCode + EOSchar + "'"
                    Else
                         GroupCode = GroupCode + "'"
                    End If
                    Exit Do
                    
                'Add valid characters between literals
                Case 32, 9, 33 To 38, 40 To 127
                GroupCode = GroupCode + Mid(InputCode, Pos, 1)
                Case Else
                    ErrorMessages = "Invald Character: " + Mid(InputCode, Pos, 1)
                    Exit Function
                End Select
            Loop
        
        'Strip out comments ---------------------------------------------
        Case Asc(";")
            'Comments
                        
            'Find end of comment
            EndOfComment = InStr(Pos, InputCode, vbCr)
            If EndOfComment = 0 Then
                'store comment in array
                Exit Do
            Else
                'Set pos to just before CR to be picked
                'up next time round
                Pos = EndOfComment - 1
            End If
        Case 32, 9
            'Spaces or tab
            If LastSymbolType <> "SPACE" Then
                GroupCode = GroupCode + EOSchar
            End If
            LastSymbolType = "SPACE"
        Case 13
            'Newlines
            If LastSymbolType <> "SPACE" Then
                GroupCode = GroupCode + EOSchar + EOLchar
            Else
                GroupCode = GroupCode + EOLchar
            End If
            LastSymbolType = "NEWLINE"
            Inc LineNo
            Inc Pos
        Case Else
            'Character not allowed in language
            ErrorMessages = "Invald Character: " + Chr(CurrentChar)
            Exit Function
        End Select
     
        If Pos >= Len(InputCode) Then Exit Do
        Inc Pos
        
    Loop
    
    'Check if last word has terminator
    If Right(GroupCode, 1) <> EOSchar Then
        GroupCode = GroupCode + EOSchar
    End If
    
    'Remove first EOSchar
    GroupCode = Right(GroupCode, Len(GroupCode) - 1)
End Function

Function IsVALValid(ByVal LabelVariableArray As String) As Boolean
    'Check if array, variable or label is a valid name
    
    IsVALValid = True
    
    Select Case UCase(LabelVariableArray)
    Case "LISTVARS"
        IsVALValid = False
    Case "ADD"
        IsVALValid = False
    Case "SUB"
        IsVALValid = False
    Case "MPY"
        IsVALValid = False
    Case "DVD"
        IsVALValid = False
    Case "MOD"
        IsVALValid = False
    Case "CLRZ"
        IsVALValid = False
    Case "CMPR"
        IsVALValid = False
    Case "LOAD"
        IsVALValid = False
    Case "COPY"
        IsVALValid = False
    Case "JUMP"
        IsVALValid = False
    Case "JEQZ"
        IsVALValid = False
    Case "JNEZ"
        IsVALValid = False
    Case "JLEZ"
        IsVALValid = False
    Case "JLTZ"
        IsVALValid = False
    Case "JGEZ"
        IsVALValid = False
    Case "JGTZ"
        IsVALValid = False
    Case "JSUBR"
        IsVALValid = False
    Case "EXIT"
        IsVALValid = False
    Case "HALT"
        IsVALValid = False
    Case "INPTI"
        IsVALValid = False
    Case "OUPTI"
        IsVALValid = False
    Case "OUPTS"
        IsVALValid = False
    Case "OUPTS"
        IsVALValid = False
    Case "BLOCKI"
        IsVALValid = False
    Case "DATAI"
        IsVALValid = False
    Case "ACC"
        IsVALValid = False
    Case "INDX"
        IsVALValid = False
    Case "FLAG"
        IsVALValid = False
    Case "KBD"
        IsVALValid = False
    Case "SCR"
        IsVALValid = False
    Case "RND"
        IsVALValid = False
    Case "CLRS"
        IsVALValid = False
    End Select
    
End Function


Function InterpretCode(ByVal rtbToCheck As RichTextBox, ByVal DisplayMessage As Boolean) As Boolean
    'Test and compile code
    
    'Clear variables
    GroupedCode = Empty
    ArrayCount = 0
    VariableCount = 0
    LabelCount = 0
    ErrorMessages = Empty
    CodeState = 0
    InterpretCode = False
    
    'Set mouse pointer
    Screen.MousePointer = 11
            
    'Groupcode
    GroupedCode = GroupCode(rtbToCheck.Text)
    
    'Get variable declarations
    frmEditor.CheckLabels False
    
    'Display error messages then exit
    If ErrorMessages <> Empty Then
        If DisplayMessage = True Then frmEditor.DisplayError
        InterpretCode = True
        Screen.MousePointer = 0
        Exit Function
    End If
    
    WindowLineCount = WindowLineNumber
    SymPos = 1
    LineNumber = 0
    WindowLineNumber = 0
    
    'parse operations
    Do
        
        'reset temp variables
        TempOperation = 0
        TempOperand1 = 0
        TempOperand2 = 0
        TempOperand3 = 0
        TempOperandText = Empty
        
        'Parse the command
        ParseCommand False
        
        'Error found
        If ErrorMessages <> Empty Then
            If DisplayMessage = True Then frmEditor.DisplayError
            InterpretCode = True
            Screen.MousePointer = 0
            Exit Function
        End If
        
        Select Case TempOperation
        Case -1
            'end of program found
            Exit Do
        Case -13
            'New line
            Inc WindowLineNumber
        Case -99
            'Error found
            ErrorMessages = "Syntax error"
            If DisplayMessage = True Then frmEditor.DisplayError
            InterpretCode = True
            Screen.MousePointer = 0
            Exit Function
        Case Else
                    
            'Keyword found - nexy sym should be new line or end
            Operation(LineNumber) = TempOperation
            Operand1(LineNumber) = TempOperand1
            Operand2(LineNumber) = TempOperand2
            Operand3(LineNumber) = TempOperand3
            OperandText(LineNumber) = TempOperandText
            
            GetSym False
            Inc LineNumber
            
            'Redim the array if needed
            If LineNumber >= UBound(Operation()) Then
                ReDim Preserve Operation(LineNumber)
                ReDim Preserve Operand1(LineNumber)
                ReDim Preserve Operand2(LineNumber)
                ReDim Preserve Operand3(LineNumber)
                ReDim Preserve OperandText(LineNumber)
            End If
            
            If Sym = "NEWLINE" Then
                Inc WindowLineNumber
            ElseIf Sym = "END" Then
                Exit Do
            Else
                'Error no new line or end
                ErrorMessages = "Syntax error"
                Screen.MousePointer = 0
                If DisplayMessage = True Then frmEditor.DisplayError
                InterpretCode = True
                Exit Function
            End If
        End Select
    Loop

    LineCount = LineNumber
    Operation(LineCount) = -1
    
    'Reset the old pointer
    Screen.MousePointer = 0
End Function

Function GetFlagValue() As String
    'Returns the flag value as string
    If LastRegister < 0 Then
        GetFlagValue = "-"
    ElseIf LastRegister > 0 Then
        GetFlagValue = "+"
    Else
        GetFlagValue = "0"
    End If

End Function

Function IsOutOfRange(ByVal lngValue As Long) As Boolean
    'Tests wether a number is out of integer range
    If lngValue < -32768 Or lngValue > 32767 Then
        IsOutOfRange = True
        ErrorMessages = "Overflow"
    Else
        IsOutOfRange = False
    End If
    
End Function

Public Function IsArrayIndexOutOfRange() As Boolean
    'Determines if current array index is out of range
    Dim RegisterValue As Integer
   
    Select Case TempOperand2
    Case 0, 3
        'Direct address or variable
        'No array index to check
        IsArrayIndexOutOfRange = False
    Case Else
        'Indexed (array) or RND function
                    
        'check what register to use as subscript
        If TempOperand2 = 1 Then
            'Acc
            RegisterValue = Acc
        ElseIf TempOperand2 = 2 Then
            'Indx
            RegisterValue = Indx
        Else
            'Direct subscript
            RegisterValue = TempOperand2 - 4
        End If
        
        'RND flag set to -1
        If TempOperand3 = -1 Then
            IsArrayIndexOutOfRange = False
            Exit Function
        End If
        
        'check for out of range
        If RegisterValue > ArrayElements(TempOperand3) Or RegisterValue < 0 Then
            ErrorMessages = "Array index out of range"
            IsArrayIndexOutOfRange = True
        Else
            IsArrayIndexOutOfRange = False
        End If
    End Select

End Function

Function Capitalise(ByVal InText As String)
    'Capitalise a string
    Select Case Len(InText)
    Case 0
        Capitalise = Empty
    Case 1
        Capitalise = UCase(InText)
    Case Else
        Capitalise = UCase(Left(InText, 1)) & LCase(Right(InText, Len(InText) - 1))
    End Select
End Function

'Simply reverse instring
Public Function InStrR(Optional lStart As Long, Optional sTarget As String, Optional sFind As String, Optional iCompare As Integer) As Long
    Dim cFind As Long, i As Long
    cFind = Len(sFind)
    For i = lStart - cFind + 1 To 1 Step -1
        If StrComp(Mid$(sTarget, i, cFind), sFind, iCompare) = 0 Then
            InStrR = i
            Exit Function
        End If
    Next
End Function


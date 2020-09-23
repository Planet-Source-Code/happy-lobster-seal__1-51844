Attribute VB_Name = "modCode"
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

Function ConvertOperation() As String
    Dim a As Integer
    Dim b As Integer
    
    'Convert operation back to code
    Select Case TempOperation
    Case 0
        'Add keyword
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Add Acc,") + GetVariableName
            TempReminder = "Add ACC with " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Add Acc")
        Else
            ConvertOperation = SetCase("Add Indx,") + GetVariableName
            TempReminder = "Add INDX with " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Add Indx")
        End If
    Case 1
        'Sub keyword
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Sub Acc,") + GetVariableName
            TempReminder = "Subtract ACC with " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Sub Acc")
        Else
            ConvertOperation = SetCase("Sub Indx,") + GetVariableName
            TempReminder = "Subtract INDX with " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Sub Indx")
        End If
    Case 2
        'Multiply
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Mpy Acc,") + GetVariableName
            TempReminder = "Mulitply ACC with " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Mpy Acc")
        Else
            ConvertOperation = SetCase("Mpy Indx,") + GetVariableName
            TempReminder = "Multiply INDX with " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Mpy Indx")
        End If
    Case 3
        'Divide
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Dvd Acc,") + GetVariableName
            TempReminder = "Divide ACC with " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Dvd Acc")
        Else
            ConvertOperation = SetCase("Dvd Indx,") + GetVariableName
            TempReminder = "Divide INDX with " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Dvd Indx")
        End If
    Case 4
        'Mod
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Mod Acc,") + GetVariableName
            TempReminder = "Find remainder of ACC when divided by " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Mod Acc")
        Else
            ConvertOperation = SetCase("Mod Indx,") + GetVariableName
            TempReminder = "Find remainder of INDX when divided by " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Mod Indx")
        End If
    Case 5   'NEG
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Neg Acc")
           TempReminder = "Negate ACC"
           TempCodeWithNoAddress = SetCase("Neg Acc")
        ElseIf TempOperand1 = 2 Then
            ConvertOperation = SetCase("Neg Indx")
            TempReminder = "Negate INDX"
            TempCodeWithNoAddress = SetCase("Neg Indx")
        Else
            ConvertOperation = SetCase("Neg Flag")
            TempReminder = "Negate FLAG"
            TempCodeWithNoAddress = SetCase("Neg Flag")
        End If
    Case 6
        'Clear reg
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Clrz Acc")
           TempReminder = "Clear ACC"
           TempCodeWithNoAddress = SetCase("Clrz Acc")
        ElseIf TempOperand1 = 2 Then
            ConvertOperation = SetCase("Clrz Indx")
            TempReminder = "Clear INDX"
            TempCodeWithNoAddress = SetCase("Clrz Indx")
        Else
            ConvertOperation = SetCase("Clrz Flag")
            TempReminder = "Clear FLAG"
            TempCodeWithNoAddress = SetCase("Clrz Flag")
        End If
    Case 7 'inc
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Inc Acc")
           TempReminder = "Increment ACC"
           TempCodeWithNoAddress = SetCase("Inc Acc")
        Else
            ConvertOperation = SetCase("Inc Indx")
            TempReminder = "Increment INDX"
            TempCodeWithNoAddress = SetCase("Inc Indx")
        End If
    Case 8 'dec
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Dec Acc")
           TempReminder = "Decrement ACC"
           TempCodeWithNoAddress = SetCase("Dec Acc")
        Else
            ConvertOperation = SetCase("Dec Indx")
            TempReminder = "Decrement INDX"
            TempCodeWithNoAddress = SetCase("Dec Indx")
        End If
    
    Case 9
        'Compare
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Cmpr Acc,") + GetVariableName
            TempReminder = "Compare ACC with " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Cmpr Acc")
        Else
            ConvertOperation = SetCase("Cmpr indx,") + GetVariableName
            TempReminder = "Compare INDX with " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Cmpr Indx")
        End If
    Case 10
        'Load
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Load Acc,") + GetVariableName
            TempReminder = "Put into ACC value of " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Load Acc")
        Else
            ConvertOperation = SetCase("Load Indx,") + GetVariableName
            TempReminder = "Put into INDX value of " + NoHash(GetVariableName)
            TempCodeWithNoAddress = SetCase("Load Indx")
        End If
    Case 11
        'Copy
        If TempOperand1 = 1 Then
            ConvertOperation = SetCase("Copy Acc,") + GetVariableName
            TempReminder = "Store ACC value into " + GetVariableName
            TempCodeWithNoAddress = SetCase("Copy Acc")
        Else
            ConvertOperation = SetCase("Copy Indx,") + GetVariableName
            TempReminder = "Store INDX value into " + GetVariableName
            TempCodeWithNoAddress = SetCase("Copy Indx")
        End If
    Case 12
        'Jump
        ConvertOperation = SetCase("Jump " + GetLabelName)
        TempReminder = "Jump to " + GetLabelName
        TempCodeWithNoAddress = ConvertOperation
    Case 13
        'Jump =0
         ConvertOperation = SetCase("Jeqz ") + GetLabelName
        TempReminder = "Jump if FLAG=0 to " + GetLabelName
        TempCodeWithNoAddress = ConvertOperation
    Case 14
        'Jump <=0
        ConvertOperation = SetCase("Jlez ") + GetLabelName
        TempReminder = "Jump if FLAG<=0 to " + GetLabelName
        TempCodeWithNoAddress = ConvertOperation
    Case 15
        'Jump <0
        ConvertOperation = SetCase("Jltz ") + GetLabelName
        TempReminder = "Jump if FLAG<0 to " + GetLabelName
        TempCodeWithNoAddress = ConvertOperation
    Case 16
        'Jump >=0
        ConvertOperation = SetCase("Jgez ") + GetLabelName
        TempReminder = "Jump if FLAG>=0 to " + GetLabelName
        TempCodeWithNoAddress = ConvertOperation
    Case 17
        'Jump >0
        ConvertOperation = SetCase("Jgtz ") + GetLabelName
        TempReminder = "Jump if FLAG>0 to " + GetLabelName
        TempCodeWithNoAddress = ConvertOperation
    Case 18
        'Jsubr
        ConvertOperation = SetCase("Jsubr ") + GetLabelName
        TempReminder = "Subroutine jump to " + GetLabelName
        TempCodeWithNoAddress = ConvertOperation
    Case 19
        'Exit sub
        ConvertOperation = SetCase("Exit")
        TempReminder = "Exit subroutine"
        TempCodeWithNoAddress = ConvertOperation
    Case 20
        'Halt
        ConvertOperation = SetCase("Halt")
        TempReminder = "End program"
        TempCodeWithNoAddress = ConvertOperation
    Case 21
        'Input
        ConvertOperation = SetCase("Inpti Kbd,") + GetVariableName
        TempReminder = "Store keyboard input into " + GetVariableName
        TempCodeWithNoAddress = SetCase("Inpti Kbd")
    Case 22
        'Output
        ConvertOperation = SetCase("Oupti Scr,") + GetVariableName
        TempReminder = "Output to screen " + NoHash(GetVariableName) + " to the screen"
        TempCodeWithNoAddress = SetCase("Oupti Scr")
    Case 23
        'Output
        ConvertOperation = SetCase("Oupts Scr,") + "'" + TempOperandText + "'"
        TempReminder = "Output the string '" + TempOperandText + "' to the screen"
        TempCodeWithNoAddress = SetCase("Oupts Scr")
    Case 24 'Query variable
    Case 25 'List vars
    Case 26
        'Clear screen
        ConvertOperation = SetCase("Clrs")
        TempCodeWithNoAddress = SetCase("Clrs")
        TempReminder = "Clear screen"
        If TempOperand1 <> -1 Then
            ConvertOperation = ConvertOperation + " " + GetVariableName
            TempReminder = TempCodeWithNoAddress + " with colour " + NoHash(GetVariableName)
        End If
    Case 27
        'Jump<>0
         ConvertOperation = SetCase("Jnez ") + GetLabelName
         TempReminder = "Jump if FLAG<>0 to " + GetLabelName
    Case 28
        'Output colour
        ConvertOperation = SetCase("Ouptc Scr,") + GetVariableName
        TempReminder = "Set the screen text colour to " + GetVariableName
    End Select
End Function

Function SetCase(InputStr As String)
    'Redundant
    SetCase = InputStr
End Function

Function GetVariableName(Optional ByVal blnGetRegisterValue As Boolean) As String
    
    'Returns value of opperand or variable
    'of current instruction
    Dim RegisterValue As Integer
    Dim ArrayRndStr As String
   
   Debug.Print blnGetRegisterValue
   
    Select Case TempOperand2
    Case 0
        'Variable
       GetVariableName = VariableName(TempOperand3)
    Case 3
        'Direct address
        GetVariableName = "#" & Format(TempOperand3)
    Case Else
        'Indexed (array) or RND function
        
        'Use RND or Array name
        If TempOperand3 = -1 Then
            ArrayRndStr = "RND"
        Else
            ArrayRndStr = ArrayName(TempOperand3)
        End If
        
        'check what register to use as subscript
        If TempOperand2 = 1 Then
            'Acc
            If blnGetRegisterValue = True Then
                GetVariableName = ArrayRndStr + "(" + Format(Acc) + ")"
            Else
                GetVariableName = ArrayRndStr + "(ACC)"
            End If
        ElseIf TempOperand2 = 2 Then
            'Indx
            If blnGetRegisterValue = True Then
                GetVariableName = ArrayRndStr + "(" + Format(Indx) + ")"
            Else
                GetVariableName = ArrayRndStr + "(INDX)"
            End If
        Else
            'Direct subscript
            GetVariableName = ArrayRndStr + "(" + Format(TempOperand2 - 4) + ")"
        End If

    End Select
    GetVariableName = Capitalise(GetVariableName)
    GetVariableName = SetCase(GetVariableName)
    If glbShowCodeView = True Then
        If glbVariableLabelUC = True Then
            GetVariableName = UCase(GetVariableName)
        End If
    End If
    

End Function

Function GetLabelName()
    Dim a As Integer
    Dim TempStr As String
    'Search for variablename
    For a = 0 To LabelCount - 1
        If TempOperand1 = LabelPos(a) Then
             TempStr = LabelName(a)
        End If
    Next
    TempStr = Capitalise(TempStr)
    TempStr = SetCase(TempStr)
    
    'Uppercase variable or label name if required
    If glbShowCodeView = True Then
        If glbVariableLabelUC = True Then
            TempStr = UCase(TempStr)
        End If
    End If
    GetLabelName = TempStr
    
End Function

Function NoHash(ByVal strIn As String) As String
    'Remove the first hash from a string
    If Left(strIn, 1) = "#" Then
        strIn = Right(strIn, Len(strIn) - 1)
    End If
    NoHash = strIn
End Function

Function ExplainOperation() As String
    Dim a As Integer
    Dim b As Integer
    
    'Convert operation back to code
    Select Case TempOperation
    Case 0
        'Add keyword
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Add Acc,") + ExplainAddress
            TempReminder = "Add ACC with " + NoHash(ExplainAddress)
        Else
            ExplainOperation = SetCase("Add Indx,") + ExplainAddress
            TempReminder = "Add INDX with " + NoHash(ExplainAddress)
        End If
    Case 1
        'Sub keyword
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Sub Acc,") + ExplainAddress
            TempReminder = "Subtract ACC with " + NoHash(ExplainAddress)
        Else
            ExplainOperation = SetCase("Sub Indx,") + ExplainAddress
            TempReminder = "Subtract INDX with " + NoHash(ExplainAddress)
        End If
    Case 2
        'Multiply
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Mpy Acc,") + ExplainAddress
            TempReminder = "Mulitply ACC with " + NoHash(ExplainAddress)
        Else
            ExplainOperation = SetCase("Mpy Indx,") + ExplainAddress
            TempReminder = "Multiply INDX with " + NoHash(ExplainAddress)
        End If
    Case 3
        'Divide
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Dvd Acc,") + ExplainAddress
            TempReminder = "Divide ACC with " + NoHash(ExplainAddress)
        Else
            ExplainOperation = SetCase("Dvd Indx,") + ExplainAddress
            TempReminder = "Divide INDX with " + NoHash(ExplainAddress)
        End If
    Case 4
        'Mod
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Mod Acc,") + ExplainAddress
            TempReminder = "Find remainder of ACC when divided by " + NoHash(ExplainAddress)
        Else
            ExplainOperation = SetCase("Mod Indx,") + ExplainAddress
            TempReminder = "Find remainder of INDX when divided by " + NoHash(ExplainAddress)
        End If
    Case 5   'NEG
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Neg Acc")
           TempReminder = "Negate ACC"
        ElseIf TempOperand1 = 2 Then
            ExplainOperation = SetCase("Neg Indx")
            TempReminder = "Negate INDX"
        Else
            ExplainOperation = SetCase("Neg Flag")
            TempReminder = "Negate FLAG"
        End If
    Case 6
        'Clear reg
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Clrz Acc")
           TempReminder = "Clear ACC"
        ElseIf TempOperand1 = 2 Then
            ExplainOperation = SetCase("Clrz Indx")
            TempReminder = "Clear INDX"
        Else
            ExplainOperation = SetCase("Clrz Flag")
            TempReminder = "Clear FLAG"
        End If
    Case 7 'inc
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Inc Acc")
           TempReminder = "Increment ACC"
        Else
            ExplainOperation = SetCase("Inc Indx")
            TempReminder = "Increment INDX"
        End If
    Case 8 'dec
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Dec Acc")
           TempReminder = "Decrement ACC"
        Else
            ExplainOperation = SetCase("Dec Indx")
            TempReminder = "Decrement INDX"
        End If
    
    Case 9
        'Compare
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Cmpr Acc,") + ExplainAddress
            TempReminder = "Compare ACC with " + NoHash(ExplainAddress)
        Else
            ExplainOperation = SetCase("Cmpr indx,") + ExplainAddress
            TempReminder = "Compare INDX with " + NoHash(ExplainAddress)
        End If
    Case 10
        'Load
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Load Acc,") + ExplainAddress
            TempReminder = "Put into ACC value of " + NoHash(ExplainAddress)
        Else
            ExplainOperation = SetCase("Load Indx,") + ExplainAddress
            TempReminder = "Put into INDX value of " + NoHash(ExplainAddress)
        End If
    Case 11
        'Copy
        If TempOperand1 = 1 Then
            ExplainOperation = SetCase("Copy Acc,") + ExplainAddress
            TempReminder = "Store ACC value into " + ExplainAddress
        Else
            ExplainOperation = SetCase("Copy Indx,") + ExplainAddress
            TempReminder = "Store INDX value into " + ExplainAddress
        End If
    Case 12
        'Jump
        ExplainOperation = SetCase("Jump " + ExplainLabelTo)
        TempReminder = "Jump to " + ExplainLabelTo
    Case 13
        'Jump =0
         ExplainOperation = SetCase("Jeqz ") + ExplainLabelTo
         TempReminder = "Jump if FLAG=0 to " + ExplainLabelTo
    Case 14
        'Jump <=0
        ExplainOperation = SetCase("Jlez ") + ExplainLabelTo
        TempReminder = "Jump if FLAG<=0 to " + ExplainLabelTo
    Case 15
        'Jump <0
        ExplainOperation = SetCase("Jltz ") + ExplainLabelTo
        TempReminder = "Jump if FLAG<0 to " + ExplainLabelTo
    Case 16
        'Jump >=0
        ExplainOperation = SetCase("Jgez ") + ExplainLabelTo
        TempReminder = "Jump if FLAG>=0 to " + ExplainLabelTo
    Case 17
        'Jump >0
        ExplainOperation = SetCase("Jgtz ") + ExplainLabelTo
        TempReminder = "Jump if FLAG>0 to " + ExplainLabelTo
    Case 18
        'Jsubr
        ExplainOperation = SetCase("Jsubr ") + ExplainLabelTo
        TempReminder = "Subroutine jump to " + ExplainLabelTo
    Case 19
        'Exit sub
        ExplainOperation = SetCase("Exit")
        TempReminder = "Exit subroutine"
    Case 20
        'Halt
        ExplainOperation = SetCase("Halt")
        TempReminder = "End program"
    Case 21
        'Input
        ExplainOperation = SetCase("Inpti Kbd,") + ExplainAddress
        TempReminder = "Store keyboard input in to " + ExplainAddress
    Case 22
        'Output
        ExplainOperation = SetCase("Oupti Scr,") + ExplainAddress
        TempReminder = "Output to screen " + NoHash(ExplainAddress)
    Case 23
        'Output
        ExplainOperation = SetCase("Oupts Scr,") + "'" + ExplainAddress + "'"
        TempReminder = "Output the string '" + ExplainAddress + "' to the screen"
    Case 24 'Query variable
    Case 25 'List vars
    Case 26
        'Clear screen
        ExplainOperation = SetCase("Clrs")
        TempReminder = "Clear screen"
        If TempOperand1 <> -1 Then
            ExplainOperation = ExplainOperation + " " + ExplainAddress
            TempReminder = TempReminder + " with colour " + ExplainAddress
        End If
    Case 27
        'Jump<>0
         ExplainOperation = SetCase("Jnez ") + ExplainLabelTo
         TempReminder = "Jump if FLAG<>0 to " + ExplainLabelTo
    Case 28
        'Output colour
        ExplainOperation = SetCase("Ouptc Scr,") + ExplainAddress
        TempReminder = "Set the screen text colour to " + NoHash(ExplainAddress)
    End Select
End Function

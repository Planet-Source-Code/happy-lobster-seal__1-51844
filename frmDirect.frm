VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmConsole 
   AutoRedraw      =   -1  'True
   Caption         =   "(F6) Console"
   ClientHeight    =   4110
   ClientLeft      =   7890
   ClientTop       =   6870
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDirect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   274
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   Begin RichTextLib.RichTextBox rtbOutput 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6800
      _Version        =   393217
      ScrollBars      =   3
      RightMargin     =   32767
      OLEDropMode     =   0
      TextRTF         =   $"frmDirect.frx":08D2
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   675
      Top             =   120
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Dim ControlPressed As Boolean
Dim PenColour As Long
Dim UserValue As Integer


Sub RunProgram()
    'Run program
    
    Dim ListItem As String
    
    'User has pressed enter
    If GroupedCode = Empty Then
        Exit Sub
    End If
    
    'reset temp variables
    TempOperation = 0
    TempOperand1 = 0
    TempOperand2 = 0
    TempOperand3 = 0
    TempOperandText = Empty
    ErrorMessages = Empty
    
    'Check for variable additions
    Console = True
    frmEditor.CheckLabels False
    Console = False
    If ErrorMessages <> Empty Then
        'Print output messages
        fPen glbErrorCol
        fPrint ErrorMessages & vbCrLf
        Exit Sub
    End If
    
    SymPos = 1
    
    ParseCommand False
    
    'Check for invalid commands
    Select Case TempOperation
    Case 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 27
        ErrorMessages = "Command invalid in console mode"
    End Select
    
    If ErrorMessages <> Empty Then
        'Print output messages
        fPen glbErrorCol
        fPrint ErrorMessages & vbCrLf
        Exit Sub
    End If
    
    'Check for end of line
    GetSym False
    If Sym <> "END" Then
        ErrorMessages = "Syntax error"
    End If
    
    'Error found
    If ErrorMessages <> Empty Then
    
        'Print output messages
        fPen glbErrorCol
        fPrint ErrorMessages & vbCrLf
        Exit Sub
        
    End If
    
    
    Select Case TempOperation
    Case -1     'end of program found
    Case -13    'New line
    Case -99    'Error found
        
        'Print output messages
        fPen glbErrorCol
        fPrint "Syntax error" & vbCrLf
    Case Else
    
        'Execute command
        ExecuteCommand TempOperation
        
        If ErrorMessages <> Empty Then
        
            'Print output messages
            fPen glbErrorCol
            fPrint ErrorMessages & vbCrLf
        End If
    End Select

End Sub
Sub ExecuteCommand(ByVal CommandNo As Integer)
    Dim V As Integer
    Dim MaxLength
    Dim TempStr As String
    Dim E As Integer
    Dim T As Integer
    
    On Error GoTo error_handler
        
    Select Case CommandNo
    Case 0
        'Add keyword
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc + GetValue) Then GoTo error_handler
            Acc = Acc + GetValue
            LastRegister = Acc
        Else
            If IsOutOfRange(Indx + GetValue) Then GoTo error_handler
            Indx = Indx + GetValue
            LastRegister = Indx
        End If
    Case 1
        'Sub keyword
        
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc - GetValue) Then GoTo error_handler
            Acc = Acc - GetValue
            LastRegister = Acc
        Else
            If IsOutOfRange(Indx - GetValue) Then GoTo error_handler
            Indx = Indx - GetValue
            LastRegister = Indx
        End If
    Case 2
        'Multiply
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc * GetValue) Then GoTo error_handler
            Acc = Acc * GetValue
            LastRegister = Acc
        Else
            If IsOutOfRange(Indx * GetValue) Then GoTo error_handler
            Indx = Indx * GetValue
            LastRegister = Indx
        End If
    Case 3
        'Divide
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc \ GetValue) Then GoTo error_handler
            Acc = Acc \ GetValue
            LastRegister = Acc
        Else
            If IsOutOfRange(Indx \ GetValue) Then GoTo error_handler
            Indx = Indx \ GetValue
            LastRegister = Indx
        End If
    Case 4
        'Mod
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc Mod GetValue) Then GoTo error_handler
            Acc = Acc Mod GetValue
            LastRegister = Acc
        Else
            If IsOutOfRange(Indx Mod GetValue) Then GoTo error_handler
            Indx = Indx Mod GetValue
            LastRegister = Indx
        End If
    Case 5
        'Negate register or flag
        If TempOperand1 = 1 Then
            Acc = Acc * -1
        ElseIf TempOperand1 = 2 Then
            Indx = Indx * -1
        Else
            LastRegister = LastRegister * -1
        End If
    
    Case 6
        'Clear reg
        If TempOperand1 = 1 Then
            Acc = 0
        ElseIf TempOperand1 = 2 Then
            Indx = 0
        Else
            LastRegister = 0
        End If
    Case 7
        'Increase register
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc + 1) Then GoTo error_handler
            Inc Acc
            LastRegister = Acc
        Else
            If IsOutOfRange(Indx + 1) Then GoTo error_handler
            Inc Indx
            LastRegister = Indx
        End If
    Case 8
        'Dec register
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc - 1) Then GoTo error_handler
            Dec Acc
            LastRegister = Acc
        Else
            If IsOutOfRange(Indx - 1) Then GoTo error_handler
            Dec Indx
            LastRegister = Indx
        End If
    
    Case 9
        'Compare - don't actualy change reg values
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc - GetValue) Then GoTo error_handler
            LastRegister = Acc - GetValue
        Else
            If IsOutOfRange(Indx - GetValue) Then GoTo error_handler
            LastRegister = Indx - GetValue
        End If
    Case 10
        'Load
        If TempOperand1 = 1 Then
            Acc = GetValue
            LastRegister = Acc
        Else
            Indx = GetValue
            LastRegister = Indx
        End If
    Case 11
        'Copy
        If TempOperand1 = 1 Then
            SetValue Acc
        Else
            SetValue Indx
        End If
    Case 12
        'Jump
        LineNumber = TempOperand1 - 1

    Case 13
        'Jump =0
        If LastRegister = 0 Then
            LineNumber = TempOperand1 - 1
        End If
    Case 14
        'Jump <=0
        If LastRegister <= 0 Then
            LineNumber = TempOperand1 - 1
        End If
    Case 15
        'Jump <0
        If LastRegister < 0 Then
            LineNumber = TempOperand1 - 1
        End If
    Case 16
        'Jump >=0
        If LastRegister >= 0 Then
            LineNumber = TempOperand1 - 1
        End If
    Case 17
        'Jump >0
        If LastRegister > 0 Then
            LineNumber = TempOperand1 - 1
        End If
    Case 18
        'Jump Sub Routine
        ReturnPos(ReturnNo) = LineNumber + 1
        Inc ReturnNo
        LineNumber = TempOperand1 - 1
    Case 19
        'Exit sub
        If ReturnNo = 0 Then
            ErrorMessages = "No subroutine to return to"
            GoTo error_handler
        Else
            LineNumber = ReturnPos(ReturnNo - 1) - 1
            Dec ReturnNo
        End If
    Case 20
        'Halt
        ErrorMessages = "Program execution sucessful"

    Case 21
        'Input
        If GotUserValue = False Then
            'Get keyboard input
'todo: commented out cos not in run mode or run form
'            rtbOutput.Locked = False
'            frmAnimate.txtKeyboard.Locked = False
            ErrorMessages = "KBDINPUT"
            Exit Sub
        Else
'            rtbOutput.Locked = True
'            frmAnimate.txtKeyboard.Locked = True
            
            SetValue Val(UserValue)
            GotUserValue = False
        End If
    Case 22
        'Output
        TempStr = Format(GetValue)
        fPen glbConsoleTextColour
        fPrint TempStr & vbCrLf
    Case 23
        'Output string
        fPen glbConsoleTextColour
        fPrint TempOperandText & vbCrLf
    Case 24
        '? Query variable or register

        'Tempoperand1 set to -1 to check flag value
        If TempOperand1 = -1 Then
            TempStr = GetFlagValue
        Else
            'Normal check
            TempStr = Format(GetValue)
        End If
        
        'Output
        fPen glbConsoleTextColour
        fPrint TempStr & vbCrLf
                    
    Case 25
        'List variables
       
        'Calculate max length of string names
        MaxLength = 0
    
        For V = 0 To VariableCount - 1
            If Len(VariableName(V)) > MaxLength Then
                MaxLength = Len(VariableName(V))
            End If
        Next
        For V = 0 To ArrayCount - 1
            If Len(ArrayName(V)) > MaxLength Then
                MaxLength = Len(ArrayName(V) + "()")
            End If
        Next
        If MaxLength = 0 Then MaxLength = 4
        
        With rtbOutput
        
            'Print register values
            fPen glbConsoleTextColour
            fPrint "ACC:" + Space(MaxLength + 3 - Len("acc")) & Format(Acc) & vbCrLf
            fPrint "INDX:" + Space(MaxLength + 3 - Len("indx")) & Format(Indx) & vbCrLf
            
            'Print flag value
            fPrint "FLAG:" + Space(MaxLength + 3 - Len("flag")) & GetFlagValue & vbCrLf
            
            'Print variable values
            For V = 0 To VariableCount - 1
                TempStr = VariableName(V) + ":" & _
                Space(MaxLength + 3 - Len(VariableName(V))) & _
                Format(VariableValue(V))
                fPrint TempStr & vbCrLf
            Next
            
            'Print array details
            For V = 0 To ArrayCount - 1
                TempStr = ArrayName(V) & "():" & _
                Space(MaxLength + 3 - Len(ArrayName(V) + "()"))
                
                'Print the values in a line
                For E = 0 To ArrayElements(V)
                    TempStr = TempStr & Format(ArrayValue(V, E)) + " "
                Next
                fPrint TempStr & vbCrLf
            Next
    
        End With
    Case 26
        'Clear screen
        
        'If we have to set the background color
        'check it is in range
        If TempOperand1 <> -1 Then
            T = GetValue
            If T < 0 Or T > 15 Then
                ErrorMessages = "Invalid colour value"
            End If
        End If
    
        If ErrorMessages = Empty Then
             With rtbOutput
             
                'Clear the screen
                .Text = Empty
                .SelLength = 2
                .SelFontName = glbFont
                .SelFontSize = glbFontSize
            
                'Do we have to set the backgroud colour?
                If TempOperand1 <> -1 Then
                    .BackColor = QBColor(T)
                End If
            End With
            fPen glbConsoleTextColour
        End If
        
    Case 27
        'Jump <> 0
        If LastRegister <> 0 Then
            LineNumber = TempOperand1 - 1
        End If
    Case 28
        'Output colour

        'check the colour is in range
        T = GetValue
        If T < 0 Or T > 15 Then
            ErrorMessages = "Invalid colour value"
        End If

        'Set the pen colour, if no error
        If ErrorMessages = Empty Then
            glbConsoleTextColour = QBColor(T)
        End If
        
    End Select
    
    'show variable values
    If GotUserValue = False Then ShowVariableValues
    
    Exit Sub
    
error_handler:
    'Program running thus runtime error
'    If CodeState = 2 Then
'
'        'Error in my code, as error message has not been set
'        If ErrorMessages = Empty Then
'            ErrorMessages = "Run-time error: " & Err.Description
'        Else
'            ErrorMessages = "Run-time error: " & ErrorMessages
'        End If
'    Else
       If ErrorMessages = Empty Then ErrorMessages = Err.Description
'    End If
'    CodeState = 1
'    rtbOutput.Locked = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Hide form
    If CloseAllForms = False Then
        Cancel = True
        HideAWindow
        Me.Hide
    Else
        Unload Me
        Set frmRun = Nothing
    End If
End Sub

Private Sub rtbOutput_KeyDown(KeyCode As Integer, Shift As Integer)
    'Vet keys that the user presses
    On Error GoTo error_handler
    
    Select Case KeyCode
    Case vbKeyDown, vbKeyUp, vbKeyRight, vbKeyLeft
    
        'Disable cursor movements
        KeyCode = 0
        
    Case vbKeyBack
    
        'Disable back if character before is LF or
        'at beginning of text box
        If rtbOutput.SelStart = 0 Then
            Beep
            KeyCode = 0
        ElseIf Mid(rtbOutput.Text, rtbOutput.SelStart, 1) = vbLf Then
            Beep
            KeyCode = 0
        End If
    
    Case vbKeyControl
        
        'Set control flag
        ControlPressed = True
        
    Case ControlPressed And vbKeyC  'Allow Copy
        'Do the copy
        CopyConsoleScreen
        KeyCode = 0
    Case ControlPressed And vbKeyV
    
        'Disable pasting
        KeyCode = 0
        
    Case vbKeyReturn

        'Groupcode if not executing to decode instruction
'        If CodeState <> 2 Then
            
            ErrorMessages = Empty 'Reset flag
            
            'Group the code
            GroupedCode = GroupCode(GetCurrentText)
            
            'Output error message
            If ErrorMessages <> Empty Then
                fPen glbErrorCol
                fPrint vbCrLf
                fPrint ErrorMessages & vbCrLf
                KeyCode = 0
                fPen glbConsoleTextColour
                Exit Sub
            End If
'
'        End If
'        If CodeState = 2 Then
'            'Get uservalue
'            UserValue = Val(GetCurrentText)
'
'            'Set user value flag
'            GotUserValue = True
'        End If
        'Remove return key
        KeyCode = 0
        
        'Manually add the return character
        fPrint vbCrLf
        
        RunProgram
        ErrorMessages = Empty
    
    Case Asc("0") To Asc("9")
        
        'Digits
        SetCursorPosAndColour
        
    Case vbKeySubtract, 189
        
        'Minus key pressed
'todo: commented out because no need to have
'       If CodeState = 2 Then
'            If Len(GetCurrentText) > 0 Then
'                KeyCode = 0
'                Exit Sub
'            End If
''        End If
'        SetCursorPosAndColour
    
    Case Else
        'Any other key
        
        'if prgoram is running then remove keypress
'        If CodeState <> 2 Then
            SetCursorPosAndColour
'        Else
'            KeyCode = 0
'        End If
    End Select
    
    Exit Sub

    'Catch Val conversion errors
error_handler:
    fPen glbErrorCol
    fPrint vbCrLf
    fPrint "Run-time error: " & Err.Description & vbCrLf
    KeyCode = 0
    rtbOutput.Locked = False
    'CodeState = 1
    'NotRunningProgram
End Sub
Sub SetCursorPosAndColour()
        
    'add key presses to end
     rtbOutput.SelStart = Len(rtbOutput.Text)
     
    'Set pen colour to normal if not already
    fPen glbConsoleTextColour

End Sub

Private Sub rtboutput_KeyUp(KeyCode As Integer, Shift As Integer)
    'Reset control flag
    If KeyCode = vbKeyControl Then ControlPressed = False
End Sub

Private Sub rtbOutput_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Show copy menu
    If Button = vbRightButton Then
        glbFormToCopy = intConsole
        PopupMenu frmCopy.mnuCopy
    End If
End Sub

Function GetCurrentText() As String
    Dim Pos As Integer
    
    'Return text on current line
    With rtbOutput
        Pos = Len(.Text)
        Do
            
            If Pos <= 0 Then
                'vbLF not found because on first line
                'Return all the current text
                GetCurrentText = .Text
                Exit Function
            ElseIf Mid(.Text, Pos, 1) = vbLf Then
                'Calculate text from pos of previous vbLF
                GetCurrentText = Right(.Text, Len(.Text) - Pos)
                Exit Function
            End If
            Dec Pos
        Loop

    End With
End Function

Private Sub Form_Load()
    'Initiate form
    tmrInit.Enabled = True
End Sub

Private Sub Form_Resize()
    'Resize rich text box to fit form
    rtbOutput.Height = ScaleHeight
    rtbOutput.Width = ScaleWidth
End Sub


Private Sub rtbOutput_SelChange()
    'Show cut & copy
    CheckSelectionLength
End Sub

Sub CheckSelectionLength()
    'Check length of selection and enable cut, copy
    'delete options
    
    'Dont check if next char is vbCR or if cursor at end
    'of text box
    With frmMain
        If rtbOutput.SelLength > 0 Then
            If rtbOutput.SelStart < Len(rtbOutput.Text) Then
                If Mid(rtbOutput.Text, rtbOutput.SelStart + 1, 1) <> vbCr Then
                    If CutCopyVisible = False Then
                        .mnuEditCut.Enabled = True
                        .mnuEditCopy.Enabled = True
                        .tlbOptions.Buttons(9).Enabled = True
                        .tlbOptions.Buttons(10).Enabled = True
                        CutCopyVisible = True
                    End If
                End If
            End If
            
        Else
            If CutCopyVisible = True Then
                .mnuEditCut.Enabled = False
                .mnuEditCopy.Enabled = False
                .mnuEditDelete.Enabled = False
                .tlbOptions.Buttons(9).Enabled = False
                .tlbOptions.Buttons(10).Enabled = False
                CutCopyVisible = False
            End If
        End If
    End With
End Sub

Sub fPrint(ByVal TextToOutput As String)
    'Output text to text box
    
    With rtbOutput
        .SelStart = Len(.Text)
        .SelColor = PenColour
        .SelText = TextToOutput
    End With
End Sub

Sub fPen(ByVal NewPenColour As Long)
    'Set pen color
    PenColour = NewPenColour
    rtbOutput.SelLength = 1
    rtbOutput.SelStart = Len(rtbOutput.Text)
    rtbOutput.SelColor = NewPenColour
End Sub

Private Sub tmrInit_Timer()
    
    'Disable timer
    tmrInit.Enabled = False
 
    'Print Welcome message
    fPen vbCyan
    fPrint "_________________________________" & vbCrLf
    fPen vbMagenta
    fPrint "        Welcome to SEAL          " & vbCrLf
    fPrint "                                 " & vbCrLf
    fPrint "  Super Easy Assembler Language  " & vbCrLf
    fPrint " *Based on language by N.Fiddian " & vbCrLf
    fPen QBColor(14)
    fPrint " *Program by M.Mason             " & vbCrLf
    fPen vbCyan
    fPrint "_________________________________" & vbCrLf
    fPen vbMagenta
    fPrint "       djhappylobster@hotmail.com" & vbCrLf
    fPen glbConsoleTextColour
End Sub


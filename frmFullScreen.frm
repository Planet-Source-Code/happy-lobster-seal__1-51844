VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmFullScreen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbOutput 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmFullScreen.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmFullScreen"
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
Dim UserValue As String


Sub InitProgram()
    Dim StepStr As String

    WindowLineNumber = 0
    LineNumber = 0
    LastLineNumber = -1
    ReturnNo = 0
    CodeState = 2
    GotUserValue = False
    
    If StepMode = True Then
        StepStr = " (Step Mode)"
    End If
    
    fPen glbErrorCol
    fPrint "Program " & GetFileTitle(frmEditor.Caption) & " initiated" & StepStr & vbCrLf
    fPen glbRunTextColour
    rtbOutput.Locked = True
    frmAnimate.txtKeyboard.Locked = True
    Me.SetFocus
    
    RunProgram
End Sub

Sub RunProgram()
    'Run program
    Dim strNewString As String
    Dim ListItem As String
    
    
    
    'Code running
    If CodeState = 2 Then
            
        'Normal program execution
        Do
            'No code to run so stop program
            If LineCount = 0 Then
                ErrorMessages = "Program execution sucessful"
                Exit Do
            End If
            
            'Allow user to exit
            DoEvents
            If CodeState <> 2 Then
                ErrorMessages = Empty
                Exit Sub
            End If
        
            'Highlight line
            With frmCode.lstCode
                If LineNumber < .ListCount Then
                    'Remove -> symbol from last line executed
                    'unless this was the last line to be executed
                    If LastLineNumber > -1 Then
                        ListItem = .List(LastLineNumber)
                        strNewString = "   " & Right(ListItem, Len(ListItem) - 3)
                        .List(LastLineNumber) = strNewString
                        frmAnimate.lstCode.List(LastLineNumber) = strNewString
                    End If
                
                    'Put -> in line to be executed
                    ListItem = .List(LineNumber)
                    strNewString = "-> " & Right(ListItem, Len(ListItem) - 3)
                    .List(LineNumber) = strNewString
                    frmAnimate.lstCode.List(LineNumber) = strNewString
                    .ListIndex = LineNumber
                    frmAnimate.lstCode.ListIndex = LineNumber
                Else
                    GoTo program_finished
                End If
            End With
            
            'Store last line number
            LastLineNumber = LineNumber
            
            'Set temp op variables
            TempOperation = Operation(LineNumber)
            TempOperand1 = Operand1(LineNumber)
            TempOperand2 = Operand2(LineNumber)
            TempOperand3 = Operand3(LineNumber)
            TempOperandText = OperandText(LineNumber)

            'Clear error messages
            ErrorMessages = Empty
            
            'Execute command
            ExecuteCommand TempOperation

'todo: inserted this to do animation
'frmAnimate.Animate

            'Control error messages
            If ErrorMessages <> Empty Then
                If ErrorMessages = "KBDINPUT" Then
                
                    'Get keyboard input
                    Exit Sub
                Else
                
                    'Finish program
                    CodeState = 1
                    StepMode = False
                    NotRunningProgram
                    rtbOutput.Locked = True
                    frmAnimate.txtKeyboard.Locked = True
                    Exit Do
                End If
            End If
            
            Inc LineNumber
            Inc WindowLineNumber
            If LineNumber > LineCount Then 'todo: should it be (equal to) or (G or E)
                'Program finished
                'direct mode
program_finished:
                ErrorMessages = "Program execution sucessful"
                Exit Do
            End If
        Loop
        
        'Print output messages
        fPen glbErrorCol
        fPrint ErrorMessages & vbCrLf
        fPrint "Press F1 to re-run or Escape to quit" & vbCrLf
        frmRun.rtbOutput.Locked = True
        frmAnimate.txtKeyboard.Locked = True
        ErrorMessages = Empty
        
        CodeState = 1
        StepMode = False
        NotRunningProgram
    End If

End Sub
Sub ExecuteCommand(ByVal CommandNo As Integer)
    Dim V As Integer
    Dim MaxLength
    Dim TempStr As String
    Dim E As Integer
    Dim T As Integer
    Dim RegisterValue As String
    On Error GoTo error_handler
    
    
    'Workout random number ready for use
    If TempOperand3 = -1 Then
        Select Case TempOperand2
        Case 0, 3
        Case Else
            If TempOperand2 = 1 Then
                RegisterValue = Acc
            ElseIf TempOperand2 = 2 Then
                RegisterValue = Indx
            Else
                RegisterValue = TempOperand2 - 4
            End If
            TempRandomNumber = Int(Rnd() * (RegisterValue + 1))
        End Select
    End If
    
    'We have the user value
    If GotUserValue = False Then
        If glbAnimateType <> intNone Then
            'Animate command
            frmAnimate.Animate
        Else
            'Update the labels anyway
            With frmAnimate
                .txtInstruction = ConvertOperation
                .txtExplanation = TempReminder
            End With
        End If

        'Animate form does array index checks
        If ErrorMessages <> Empty Then GoTo error_handler
    End If
  
    
    Select Case CommandNo
    Case 0
        'Add keyword
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc + GetValue) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Acc = Acc + GetValue
                LastRegister = Acc
            End If
        Else
            If IsOutOfRange(Indx + GetValue) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Indx = Indx + GetValue
                LastRegister = Indx
            End If
        End If
    Case 1
        'Sub keyword
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc - GetValue) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Acc = Acc - GetValue
                LastRegister = Acc
            End If
        Else
            If IsOutOfRange(Indx - GetValue) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Indx = Indx - GetValue
                LastRegister = Indx
            End If
        End If
    Case 2
        'Multiply
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc * GetValue) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Acc = Acc * GetValue
                LastRegister = Acc
            End If
        Else
            If IsOutOfRange(Indx * GetValue) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Indx = Indx * GetValue
                LastRegister = Indx
            End If
        End If
    Case 3
        'Divide
        If TempOperand1 = 1 Then
            If GetValue = 0 Then
                ErrorMessages = "Division by zero"
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Acc = Acc \ GetValue
                LastRegister = Acc
            End If
        Else
            If GetValue = 0 Then
                ErrorMessages = "Division by zero"
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Indx = Indx \ GetValue
                LastRegister = Indx
            End If
        End If
    Case 4
        'Mod
        If TempOperand1 = 1 Then
            If GetValue = 0 Then
                ErrorMessages = "Division by zero"
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Acc = Acc Mod GetValue
                LastRegister = Acc
            End If
        Else
            If GetValue = 0 Then
                ErrorMessages = "Division by zero"
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Indx = Indx Mod GetValue
                LastRegister = Indx
            End If
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
        'Incrementregister
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc + 1) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Inc Acc
                LastRegister = Acc
            End If
        Else
            If IsOutOfRange(Indx + 1) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Inc Indx
                LastRegister = Indx
            End If
        End If
    Case 8
        'Dec register
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc - 1) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Dec Acc
                LastRegister = Acc
            End If
        Else
            If IsOutOfRange(Indx - 1) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                Dec Indx
                LastRegister = Indx
            End If
        End If
    
    Case 9
        'Compare - don't actualy change reg values
        If TempOperand1 = 1 Then
            If IsOutOfRange(Acc - GetValue) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                LastRegister = Acc - GetValue
            End If
        Else
            If IsOutOfRange(Indx - GetValue) Then
                frmAnimate.AnimateError intALU, ErrorMessages
                GoTo error_handler
            Else
                LastRegister = Indx - GetValue
            End If
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
            rtbOutput.Locked = False
            frmAnimate.txtKeyboard.Locked = False
            ErrorMessages = "KBDINPUT"
            Exit Sub
        Else
            rtbOutput.Locked = True
            frmAnimate.txtKeyboard.Locked = True
            GotUserValue = False
'todo: should it commented
            'lblControl = ConvertOperation
            If glbAnimateType <> intNone Then
                frmAnimate.Animate_Storage intKeyb, UserValue
            End If
            
            'Range checking is done in animate storage
            If ErrorMessages = Empty Then
                SetValue Val(UserValue)
            End If
        End If
    Case 22
        'Output integer
        TempStr = Format(GetValue)
        fPen glbRunTextColour
        fPrint TempStr & vbCrLf
    Case 23
        'Output string
        fPen glbRunTextColour
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
        fPen glbRunTextColour
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
            fPen glbRunTextColour
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
                frmAnimate.AnimateError intScr, ErrorMessages
                GoTo error_handler
            End If
        End If
    
        If ErrorMessages = Empty Then
            With rtbOutput
                'Clear the screen
                .Text = Empty
                .SelLength = 2
                .SelFontName = glbFont
                .SelFontSize = glbFontSize
            End With
            
            With frmAnimate.rtbScreen
'todo: no font name/size setting as above
                .Text = Empty
                .SelLength = 0
            End With
                
            
            'Do we have to set the backgroud colour?
            If TempOperand1 <> -1 Then
                rtbOutput.BackColor = QBColor(T)
                frmRun.rtbOutput.BackColor = QBColor(T)
                frmAnimate.rtbScreen.BackColor = QBColor(T)
            End If
        End If
        fPen glbRunTextColour
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
                'Animate error
            frmAnimate.AnimateError intScr, ErrorMessages
            GoTo error_handler
        End If

        'Set the pen colour, if no error
        If ErrorMessages = Empty Then
            glbRunTextColour = QBColor(T)
        End If
        
    End Select
    
    'show variable values
    If GotUserValue = False Then ShowVariableValues
    
    Exit Sub
    
error_handler:
    'Program running thus runtime error
    If CodeState = 2 Then
        
        'Error in my code, as error message has not been set
        If ErrorMessages = Empty Then
            ErrorMessages = "Run-time error: " & Err.Description
        Else
            ErrorMessages = "Run-time error: " & ErrorMessages
        End If
    Else
        ErrorMessages = Err.Description
    End If
    CodeState = 1
    frmRun.rtbOutput.Locked = True
    frmAnimate.txtKeyboard.Locked = True
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Hide form
    If CloseAllForms = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
        Set frmRun = Nothing
    End If
End Sub

Sub ReturnToNormal()
    'Show cursor
    ShowCursor True
    
    NotRunningProgram
    
    'Restore resolution
    ChangeRes glbScreenWidth, glbScreenHeight
    AlwaysOnTop frmFullScreen, False
    Me.Hide
    
    frmRun.Show
End Sub


Public Sub rtbOutput_KeyDown(KeyCode As Integer, Shift As Integer)
    'Vet keys that the user presses
    
    On Error GoTo error_handler
    
    'Code finished so return to normal
    Select Case KeyCode
    Case vbKeyEscape, vbKeyF3
    
        'Stop the program!
        If CodeState = 2 Then
            StopProgram
        Else
           ReturnToNormal
           Exit Sub
           KeyCode = 0
        End If
    Case vbKeyF1
        'Run the program again when its finished
        If CodeState <> 2 Then
            frmMain.mnuProgramRun_Click
        End If
    Case vbKeyF4
        'Do a restart
        If CodeState = 2 Then
            StopProgram
            InitProgram
        End If
    Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
    
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
        
        'Do back space on animate form
        If KeyCode <> 0 Then
            frmAnimate.DoBackspace
            frmRun.DoBackspace
        End If
    Case vbKeyControl
        
        'Set control flag
        ControlPressed = True
        
    Case ControlPressed And vbKeyC  'Allow Copy
    Case ControlPressed And vbKeyV
    
        'Disable pasting
        KeyCode = 0
        
    Case vbKeyReturn

        'Got input
        If CodeState = 2 Then
            GotInput
        End If
    
        KeyCode = 0
    Case Asc("0") To Asc("9")
        If Shift = 0 Then
            'Digits
            SetCursorPosAndColour
            frmAnimate.AddCharacter Chr(KeyCode)
            frmRun.AddCharacter Chr(KeyCode)
        Else
            KeyCode = 0
        End If
        
    Case vbKeySubtract, vbKeyAdd, 187, 189
        
        '-/+ key pressed
        If KeyCode = 187 And Shift = 1 Or KeyCode = 189 And Shift = 0 Or KeyCode = vbKeyAdd Or KeyCode = vbKeySubtract Then
            If CodeState = 2 Then
                If Len(GetCurrentText) > 0 Then
                    KeyCode = 0
                    Exit Sub
                End If
            End If
            
            'Convert the keycodes into keyascii values
            SetCursorPosAndColour
            If KeyCode = vbKeySubtract Or KeyCode = 189 Then
                frmAnimate.AddCharacter "-"
                frmRun.AddCharacter "-"
            ElseIf KeyCode = vbKeyAdd Or KeyCode = 187 Then
                frmAnimate.AddCharacter "+"
                frmRun.AddCharacter "+"
            End If
        Else
            KeyCode = 0
        End If
    Case vbKeyShift, vbKeyF1 To vbKeyF12
    Case Else
        'Any other key
        KeyCode = 0
    End Select
    
    Exit Sub

    'Catch Val conversion errors
error_handler:
    fPen glbErrorCol
    fPrint vbCrLf
    fPrint "Run-time error: " & Err.Description & vbCrLf
    fPrint "Press F4 to restart or Escape to quit" & vbCrLf
    KeyCode = 0
    
    'This bit was commented out
    frmRun.rtbOutput.Locked = True
    rtbOutput.Locked = True
    
    CodeState = 1
    NotRunningProgram
End Sub
Private Sub SetCursorPosAndColour()
        
    'add key presses to end
     rtbOutput.SelStart = Len(rtbOutput.Text)
     
    'Set pen colour to normal if not already
    fPen glbRunTextColour

End Sub

Public Sub GotInput()
    'Clear the keyboard text box
    frmAnimate.txtKeyboard = Empty
    
    'Get uservalue
    UserValue = GetCurrentText
    
    'Set user value flag
    GotUserValue = True
    
    frmAnimate.ClearText
    
    fPrint vbCrLf

    RunProgram
End Sub

Private Sub rtboutput_KeyUp(KeyCode As Integer, Shift As Integer)
    'Reset control flag
    If KeyCode = vbKeyControl Then ControlPressed = False
End Sub


Function GetCurrentText() As String
    'Return text on current line
    Dim lngLFposition As Long
    
    With rtbOutput

        lngLFposition = InStrR(Len(.Text), .Text, vbLf)
        If lngLFposition = 0 Then
            GetCurrentText = .Text
        Else
            GetCurrentText = Right(.Text, Len(.Text) - lngLFposition)
        End If
    End With
    
End Function





Private Sub Form_Resize()
    'Resize rich text box to fit form
    rtbOutput.Height = ScaleHeight
    rtbOutput.Width = ScaleWidth
End Sub

Sub fPen(ByVal NewPenColour As Long)
    'Set pen color
    PenColour = NewPenColour
    rtbOutput.SelStart = Len(rtbOutput.Text)
    rtbOutput.SelColor = NewPenColour
    frmAnimate.rtbScreen.SelStart = Len(rtbOutput.Text)
    frmAnimate.rtbScreen.SelColor = NewPenColour
    frmFullScreen.rtbOutput.SelStart = Len(rtbOutput.Text)
    frmFullScreen.rtbOutput.SelColor = NewPenColour
End Sub

Sub fPrint(ByVal TextToOutput As String)
    'Output text to text box
    With rtbOutput
        .SelStart = Len(.Text)
        .SelColor = PenColour
        .SelText = TextToOutput
    End With
    With frmRun.rtbOutput
        .SelStart = Len(.Text)
        .SelColor = PenColour
        .SelText = TextToOutput
    End With
    With frmAnimate.rtbScreen
        .SelStart = Len(.Text)
        .SelColor = PenColour
        .SelText = TextToOutput
    End With
End Sub

Private Sub rtbOutput_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Set sel position to end
    If rtbOutput.Text = Empty Then
        rtbOutput.SelStart = 1
    Else
        rtbOutput.SelStart = Len(rtbOutput.Text) - 1
    End If
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


Sub StopProgram()
    'Stop program from running
    If CodeState = 2 Then
        fPen glbErrorCol
        fPrint vbCrLf & "Program execution stopped" & vbCrLf
        fPrint "Press F4 to restart or Escape to quit" & vbCrLf
        frmRun.rtbOutput.Locked = True
        rtbOutput.Locked = True
        CodeState = 1
        
    End If
End Sub



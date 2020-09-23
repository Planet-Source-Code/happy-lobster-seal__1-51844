VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEditor 
   Caption         =   "Untitled"
   ClientHeight    =   6285
   ClientLeft      =   240
   ClientTop       =   11460
   ClientWidth     =   11100
   Icon            =   "frmTest2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   11100
   Begin RichTextLib.RichTextBox rtbTemp 
      Height          =   5595
      Left            =   5700
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   9869
      _Version        =   393217
      RightMargin     =   10000
      TextRTF         =   $"frmTest2.frx":08D2
   End
   Begin RichTextLib.RichTextBox rtbProgram 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7223
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      RightMargin     =   32767
      OLEDropMode     =   0
      TextRTF         =   $"frmTest2.frx":0954
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
Attribute VB_Name = "frmEditor"
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
'total number of lines in window -1

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Dim ControlPressed As Boolean   'control key pressed
Dim ReturnPressed As Boolean
Dim LastKeyPressed As Integer
Dim NewLineChecked As Boolean
Dim ScrollVal As Long

Sub DisplayError()
    Dim StartPos As Integer
    Dim EndPos As Integer
    'Display error message and highlight line
    
    'inform user
    MsgBox "Compile error: " & ErrorMessages, 48
    
    'Highlight line
    rtbProgram.SetFocus
    GetLinePoints rtbProgram, WindowLineNumber, StartPos, EndPos
    rtbProgram.SelStart = StartPos - 1
    rtbProgram.SelLength = EndPos - StartPos + 1
End Sub

Sub HighlightLine(rtbColour As RichTextBox, ByVal LineNumber As Integer, ByVal LineColour As Long)
    'Highlights line
    Dim StartPos As Integer
    Dim EndPos As Integer

    GetLinePoints rtbColour, LineNumber, StartPos, EndPos
    HighlightSection rtbColour, StartPos - 1, EndPos - 1, LineColour
End Sub

Sub HighlightSection(rtbColour As RichTextBox, ByVal StartPos As Integer, ByVal EndPos As Integer, ByVal LineColour As Long)
    Dim CursorPos As Integer
    Dim CursorLen As Integer
    
    With rtbColour
                        
        'Store cursor details
        CursorPos = .SelStart
        CursorLen = .SelLength
        
        'Highlight section
        LockWindowUpdate frmMain.hwnd
        .SelStart = StartPos
        .SelLength = (EndPos - StartPos) + 1
        .SelColor = LineColour
        LockWindowUpdate 0
    
        'Restore cusor details
        .SelStart = CursorPos
        .SelLength = CursorLen
    
    End With
End Sub


Sub ColourLine(rtbColour As RichTextBox, ByVal StartPos As Integer, ByVal EndPos As Integer)

    'Declare collections of valid words
    Dim Punctuations As New Collection
    Dim Labels As New Collection
    Dim Variables As New Collection
    Dim Commands As New Collection
    Dim Registers As New Collection
    Dim Devices As New Collection
    Dim Numbers As New Collection
    Dim Num As Integer
    Dim StartLit As Integer
    Dim EndLit As Integer
            
    
    'Colour punctuations
    With Punctuations
        .Add "("
        .Add ")"
        .Add ","
        .Add "#"
        .Add ":"
        .Add "'"
        For Num = 1 To Punctuations.Count
            ColourText rtbColour, StartPos, EndPos, Punctuations(Num), glbPunctuationCol, 0
        Next
    End With
        
    'Colour numbers
    With Numbers
        .Add "-"
        .Add "+"
        For Num = 0 To 9
            Numbers.Add Format(Num)
        Next
        For Num = 1 To Numbers.Count
            ColourText rtbColour, StartPos, EndPos, Numbers(Num), glbNumberCol, 0
        Next
    End With
    
    
    'Colour registers
    With Registers
        .Add "ACC"
        .Add "INDX"
        .Add "FLAG"
        For Num = 1 To Registers.Count
            ColourText rtbColour, StartPos, EndPos, Registers(Num), glbRegisterCol, rtfWholeWord
        Next
    End With
    
    'Colour devices
    With Devices
        .Add "KBD"
        .Add "SCR"
        For Num = 1 To Devices.Count
            ColourText rtbColour, StartPos, EndPos, Devices(Num), glbDeviceCol, rtfWholeWord
        Next
    End With
    
    'Colour commands
    With Commands
        .Add "ADD"
        .Add "SUB"
        .Add "MPY"
        .Add "DVD"
        .Add "MOD"
        .Add "NEG"
        .Add "CLRZ"
        .Add "CMPR"
        .Add "INC"
        .Add "DEC"
        .Add "LOAD"
        .Add "COPY"
        .Add "JUMP"
        .Add "JEQZ"
        .Add "JNEZ"
        .Add "JLEZ"
        .Add "JLTZ"
        .Add "JGEZ"
        .Add "JGTZ"
        .Add "JSUBR"
        .Add "EXIT"
        .Add "HALT"
        .Add "INPTI"
        .Add "OUPTI"
        .Add "OUPTC"
        .Add "OUPTS"
        .Add "CLRS"
        .Add "BLOCKI"
        .Add "DATAI"
        For Num = 1 To Commands.Count
            ColourText rtbColour, StartPos, EndPos, Commands(Num), glbCommandCol, rtfWholeWord
        Next
    End With
    
    'Colour variables
    With Variables

        Variables.Add "RND"
        Variables.Add TempVariable
        For Num = 0 To VariableCount - 1
            Variables.Add VariableName(Num)
        Next
        For Num = 0 To ArrayCount - 1
            Variables.Add ArrayName(Num)
        Next
        For Num = 1 To Variables.Count
            ColourText rtbColour, StartPos, EndPos, Variables(Num), glbVariableCol, rtfWholeWord
        Next
    End With

    'Colour labels
    With Labels
        Labels.Add TempLabelName
        Labels.Add TempLabelName2
        For Num = 0 To LabelCount - 1
            Labels.Add LabelName(Num)
        Next
        For Num = 1 To Labels.Count
            ColourText rtbColour, StartPos, EndPos, Labels(Num), glbLabelCol, rtfWholeWord, True
        Next
    End With
    
    'Colour literals
    With rtbColour
        StartLit = .Find("'", StartPos - 1, EndPos - 1)
        If StartLit <> -1 Then
            EndLit = .Find("'", StartLit + 1, EndPos)
            If EndLit - StartLit > 1 Then
                .SelStart = StartLit + 1
                .SelLength = (EndLit - StartLit) - 1
                .SelColor = glbLiteralCol
            End If
        End If
    End With

    'Colour comments
    Dim SemiColPos As Integer
    
    With rtbColour
        'Get semi col pos
        SemiColPos = .Find(";", StartPos - 1, EndPos)
        
        'Semi col found
        If SemiColPos <> -1 Then
            
            'select and colour
            .SelStart = SemiColPos
            .SelLength = EndPos - SemiColPos + 1
            .SelColor = glbCommentCol
        End If

    End With
    

End Sub


Sub ColourText(RTBox As RichTextBox, ByVal StartPos As Integer, _
ByVal EndPos As Integer, ByVal lookString As String, ByVal lookColour As Long, ByVal LookMode As Integer, Optional ByVal blnFindLabel As Boolean)


    
    Dim FindPos As Integer
    Dim OK As Boolean
    
    If lookString = Empty Then Exit Sub


    With RTBox
        Do
            'Get lookString position
            FindPos = InStr(StartPos, UCase(RTBox.Text), lookString)
            OK = False

            If FindPos > 0 And FindPos <= EndPos Then
                If LookMode = rtfWholeWord Then
                    If FindPos > 1 And FindPos + Len(lookString) < Len(RTBox.Text) Then
                        OK = ValidateWholeWord(Mid(RTBox.Text, FindPos - 1, 1), Mid(RTBox.Text, FindPos + Len(lookString), 1))
                    Else
                        OK = True
                    End If
                Else
                    OK = True
                End If
            
                
                If OK = True Then
                    'Highlight string
                    If blnFindLabel = False Then
                        .SelStart = FindPos - 1
                        .SelLength = Len(lookString)
                        .SelColor = lookColour
                    Else
                        'Looking for label so don't search after colon
                        If InStr(StartPos, RTBox.Text, ":") > FindPos Then
                            .SelStart = FindPos - 1
                            .SelLength = Len(lookString)
                            .SelColor = lookColour
                        End If
                    End If
                End If
                
                StartPos = FindPos + Len(lookString)
            
            End If
            
        Loop Until FindPos = 0 Or FindPos > EndPos
    End With
        
End Sub



Function GetLineNumber() As Integer
    'Gets current line number of
    GetLineNumber = rtbProgram.GetLineFromChar(rtbProgram.SelStart)
End Function


Private Sub Form_Load()
    'Set text box background colour
    rtbProgram.BackColor = glbEditorBackColour
    
    'Set width of text box and height
    rtbProgram.Height = ScaleHeight
    rtbProgram.Width = ScaleWidth
End Sub

Private Sub Form_LostFocus()
    'Call rtbprogram lost focus routine
    'tovalidate the line
    rtbProgram_LostFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'User is closing form
    'Check if user wants to save the file
    Dim intResponse As Integer
    Dim strMsg As String
    Dim OldNow As Single
    
    'Is program still running
    If CodeState = 2 Then
        
        'Toggle form bar
        OldNow = Timer
        FlashWindow frmRun.hwnd, 1
        Beep
        
        'Wait a bit
        Do
        Loop Until Timer - OldNow > 0.3
        
        'Toggle back
        FlashWindow frmRun.hwnd, 1
        
        Cancel = True
        Exit Sub
    End If

    'Clear the code lists
    frmCode.lstCode.Clear
    frmAnimate.lstCode.Clear
   
    If CancelOperation = False Then
        
        'Hide the code
        CodeDirty = False
        HideOptions
    Else
        Cancel = True
        Exit Sub
    End If
    
    If CloseAllForms = False Then
        Cancel = True
        glbEditorVisible = False
        Me.Hide
    Else
        Unload Me
        Set frmEditor = Nothing
    End If
    
End Sub

Private Sub Form_Resize()
    'Resize rich text box to fit form
    rtbProgram.Height = ScaleHeight
    rtbProgram.Width = ScaleWidth
End Sub

Private Sub rtbProgram_Change()
    'Set change
    glbFirstTime = True
End Sub

Public Sub rtbProgram_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim CursorPos As Integer
    Dim CursorLength As Integer
    Dim Line As Integer
    Dim GoUpALine As Boolean
    Dim StartPos As Integer
    Dim EndPos As Integer
    Dim PasteLines As Integer
    Dim StartLine As Integer
    Dim CheckBeginning As Boolean
    Dim CheckEnd As Boolean
  
    'The code is running so dont accept keypresses
    If CodeState = 2 Then
        KeyCode = 0
        Exit Sub
    End If
  
  
    Select Case KeyCode
    Case vbKeyUp, vbKeyDown, vbKeyPageUp, vbKeyPageDown
        
        'Validate last line ammended
        If glbColourSyntax = True And glbDisableColour = False Then
            With rtbProgram
                CursorPos = .SelStart
                ParseLine
                .SelStart = CursorPos
            End With
            NewLineChecked = False
        End If
    Case vbKeyControl
        'set control flag
       ControlPressed = True
    Case vbKeyV And ControlPressed
        frmMain.mnuEditPaste_Click
        KeyCode = 0
    Case vbKeyZ And ControlPressed
        'Disable undo keypress
        KeyCode = 0
    Case vbKeyA And ControlPressed
        'Do Edit, Select All
        frmMain.mnuEditSelectAll_Click
    Case vbKeyC And ControlPressed
        frmMain.mnuEditCopy_Click
        
    Case vbKeyLeft, vbKeyRight
    Case vbKeyReturn
    
        'Colour syntaxing on?
        If glbColourSyntax = True And glbDisableColour = False Then
    
            'Return key pressed
             '***********************************************
            If LastKeyPressed <> vbKeyReturn Then
                
                'Check for empty line
                If rtbProgram.SelStart > 0 Then
                    If Mid(rtbProgram.Text, rtbProgram.SelStart, 1) <> vbLf Then
                        
                        'Get line points
                        EditorLine = GetLineNumber
                        GetLinePoints rtbProgram, EditorLine, StartPos, EndPos
                        
                        'Check current line
                        'todo: highlighted out because labels not colouring properlt
                       ' CheckSection rtbProgram, StartPos, rtbProgram.SelStart + 1
                        CheckSection rtbProgram, StartPos, rtbProgram.SelStart
                        
                        'Check next new line when split
                        If rtbProgram.SelStart <= EndPos Then
                            EditorLine = GetLineNumber
                            HighlightSection rtbProgram, rtbProgram.SelStart, EndPos - 1, vbBlack
                        End If
                        
                        LineEditted = True
                        Inc EditorLine
                    End If
                End If
            End If
             '***********************************************
            If LastKeyPressed = vbKeyReturn Then
                If NewLineChecked = False Then
                    If LineEditted = True Then
                    'Check current line
                    CheckLine rtbProgram, GetLineNumber
                    ReturnPressed = 2
                    LineEditted = False
                    NewLineChecked = True
                    End If
                End If
            End If
        End If
    Case vbKeyF1 To vbKeyF12
    
    Case Else
              
        LastKeyPressed = KeyCode
         '***********************************************
        'highlight up line if going up a line
        
        glbLastLineValidated = False


        If KeyCode = vbKeyBack Then
            Line = GetLineNumber
            NewLineChecked = False
            If rtbProgram.SelStart > 0 Then
                If Mid(rtbProgram.Text, rtbProgram.SelStart, 1) = vbLf Then
                    GetLinePoints rtbProgram, Line - 1, StartPos, EndPos
                    If StartPos = EndPos Then
                        If LineEditted = True Then
                            CheckLine rtbProgram, Line
                            LineEditted = False
                            NewLineChecked = True
                        End If
                        Exit Sub
                    Else
                        GoUpALine = True
                        HighlightLine rtbProgram, Line - 1, vbBlack
                    End If
                End If
            End If
        End If
        
        
        If KeyCode = vbKeyDelete Then
            Line = GetLineNumber
            With rtbProgram
                If .SelStart + 1 < Len(.Text) Then 'Set lines to black on delete
                    If Line <> .GetLineFromChar(.SelStart + .SelLength) Then
                        HighlightLine rtbProgram, Line, vbBlack
                        HighlightLine rtbProgram, .GetLineFromChar(.SelStart + .SelLength), vbBlack
                        EditorLine = GetLineNumber
                        LineEditted = True
                        Exit Sub
                    End If
                End If
            End With
        End If
         '***********************************************
         
        'Highlight line in black if not already coloured
       If LineEditted = False Then
            With rtbProgram
                EditorLine = GetLineNumber
                HighlightLine rtbProgram, EditorLine, vbBlack
                LineEditted = True
            End With
        End If
        
        'If going up a line set editorline to next one up
        If GoUpALine = True Then
            Dec EditorLine
        End If
        
    End Select
    
    LastKeyPressed = KeyCode
End Sub

Private Sub rtbProgram_KeyPress(KeyAscii As Integer)
    'Set dirty flag
    CodeDirty = True
End Sub

Private Sub rtbProgram_KeyUp(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
    Case vbKeyControl
        'deactivate control pressed
        ControlPressed = False
    Case vbKeyReturn
        'deactivate return pressed
        ReturnPressed = False
    End Select
End Sub

Sub DoPaste(ByVal strPaste As String)
    Dim CursorPos As Integer
    Dim CursorLength As Integer
    Dim Line As Integer
    Dim GoUpALine As Boolean
    Dim StartPos As Integer
    Dim EndPos As Integer
    Dim PasteLines As Integer
    Dim StartLine As Integer
    Dim CheckBeginning As Boolean
    Dim CheckEnd As Boolean

    If glbColourSyntax = True And glbDisableColour = False Then
        'Check if clipboard is empty

        Me.MousePointer = 11
        
        'Check paste check in background
        SetupTempRTF
        rtbTemp.Text = vbCrLf + strPaste
        
        StartLine = GetLineNumber
        GetLinePoints rtbProgram, StartLine, StartPos, EndPos
        
        PasteLines = GetLineCount(rtbTemp) - 1
        
        'If the last line to be pasted is a vbCrLf Don't colour last line
        If Right(strPaste, 1) = vbLf Then
            Dec PasteLines
        End If
        
        With rtbProgram
            'If more lines than 1 then check last line
            If PasteLines > intMaxLinesPaste Then
                MsgBox "There is too much text in the clipboard - Maximum of " & Format(intMaxLinesPaste) & " lines please", vbExclamation
                MousePointer = 0
                Exit Sub
            End If
            If PasteLines > 1 Then
                
                'Recheck the first line if Sel is not first char on left
                If .SelStart > 0 Then
                    If Mid(.Text, .SelStart, 1) <> vbLf Then
                        CheckBeginning = True
                    End If
                End If
                If .SelStart + 1 < Len(.Text) Then
                    If Mid(.Text, .SelStart + 2) <> vbLf Then
                        CheckEnd = True
                    End If
                End If
            End If
        End With

        If glbColourSyntax = True And glbDisableColour = False Then
            If PasteLines > 1 Then
                'Check all the lines
                CheckLines rtbTemp, 0, PasteLines
            Else
                'Set current line has been editted
                glbLastLineValidated = False
                LineEditted = True
                EditorLine = GetLineNumber
                HighlightLine rtbProgram, EditorLine, vbBlack
            End If
        End If
        
        rtbTemp.SelStart = 2
        rtbTemp.SelLength = Len(rtbTemp.Text) - 1
        rtbProgram.SelRTF = rtbTemp.SelRTF
            
        If CheckBeginning = True Then
            Parse StartLine
        End If
        If CheckEnd = True Then
           Parse StartLine + PasteLines - 1
        End If
        Me.MousePointer = 0
        
    Else
        'Don't colour the paste
        rtbProgram.SelRTF = strPaste
    End If
    
    'Set dirty flag
    CodeDirty = True
    
    If PasteLines <> 1 Then
        glbLastLineValidated = True
    End If

End Sub

Private Sub rtbProgram_LostFocus()

    'Validate last line to be ammended
    If CodeState <> 2 Then
        If glbQuickNoCheck = False Then
            If glbLastLineValidated = False Then ParseLine
        End If
    End If
    
End Sub

Private Sub rtbProgram_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Edit/Enable Copy
    CheckSelectionLength
    
    'Show popup menu
    If Button = vbRightButton Then
        PopupMenu frmMain.mnuEdit
    End If
End Sub

Private Sub rtbProgram_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Edit/Enable Copy
    CheckSelectionLength
    
    'If line has been editted then validate
    If glbColourSyntax = True And glbDisableColour = False Then
        If EditorLine <> -1 Then
            If GetLineNumber() <> EditorLine Then
                ParseLine
            End If
        End If
    End If
End Sub

Sub Parse(ByVal LineNumber As Integer)
    Dim StartPos As Integer
    Dim EndPos As Integer
    Dim CursorPos As Integer
    Dim CursorLen As Integer
    Dim TextSelected As String
    
     '***********************************************
    
    With rtbProgram
        
        'Get the end and start of line
        GetLinePoints rtbProgram, LineNumber, StartPos, EndPos
           
        'Return text selected
        TextSelected = Mid(.Text, StartPos, EndPos - StartPos + 1)
         
         If TextSelected <> Empty Then
         
            'Set temp text box to new line
            SetupTempRTF
            rtbTemp.Text = vbCrLf + TextSelected
            
            'Check line in the backgrounc
            CheckLine rtbTemp, 1
            
            'Store cursor position
            CursorPos = .SelStart
            CursorLen = .SelLength
                
            'Highlightline temp line
            rtbTemp.SelStart = 2
            rtbTemp.SelLength = Len(rtbTemp.Text) - 2
    
            'Repaste line
            LockWindowUpdate frmMain.hwnd
            .SelStart = StartPos - 1
            .SelLength = EndPos - StartPos + 1
            .SelRTF = rtbTemp.SelRTF
            LockWindowUpdate 0
            
            'Restore original cursor settings
            .SelStart = CursorPos
            .SelLength = CursorLen
    
        End If
    End With
 '***********************************************
End Sub

Function ParseLine(Optional blnHideErrorMessage As Boolean)
    Dim blnErrorFound As Boolean
    
    'validate last line ammended
    If LineEditted = True Then
    
        Parse EditorLine
        
        If ErrorMessages <> Empty And glbAutoCheckSyntax = True Then
            If blnHideErrorMessage <> True Then
                MsgBox "Compile Error: " & ErrorMessages, vbInformation
            End If
            blnErrorFound = True
        End If
        
        glbLastLineValidated = True
        LineEditted = False
        NewLineChecked = False
        ErrorMessages = Empty
    End If

    ParseLine = blnErrorFound

End Function

Sub ExplainLine()
    Dim StartPos As Integer
    Dim EndPos As Integer
    Dim CursorPos As Integer
    Dim CursorLen As Integer
    Dim TextSelected As String
    Dim LineNumber As Integer
    Dim ExplainString As String
    Dim ExplainLabel As String
    
     '***********************************************
    
    With rtbProgram
        
        'Find the get line number
        LineNumber = GetLineNumber
        
        'Get the end and start of line
        GetLinePoints rtbProgram, LineNumber, StartPos, EndPos
           
        'Return text selected
        TextSelected = Mid(.Text, StartPos, EndPos - StartPos + 1)
         
         If TextSelected <> Empty Then
         
            'Clear temp variables
            TempReminder = Empty
            TempOperation = -1
            ExplainLabelTo = Empty
            ExplainLabel = Empty
            
            'Set temp text box to new line
            SetupTempRTF
            rtbTemp.Text = vbCrLf + TextSelected
            
            'Check line in the backgrounc
            CheckLine rtbTemp, 1
            
            'Store cursor position
            CursorPos = .SelStart
            CursorLen = .SelLength
                
            'Highlightline temp line
            rtbTemp.SelStart = 2
            rtbTemp.SelLength = Len(rtbTemp.Text) - 2
    
            'Repaste line
            .SelStart = StartPos - 1
            .SelLength = EndPos - StartPos + 1
            .SelRTF = rtbTemp.SelRTF
            
            'Restore original cursor settings
            .SelStart = CursorPos
            .SelLength = CursorLen
            
            'Alert user line couldn't be explained due to error
            If ErrorMessages <> Empty Then
                MsgBox "Compile Error: " & ErrorMessages, 48
                ErrorMessages = Empty
                Exit Sub
            End If

            
            If ExplainLabel <> Empty Then
                ExplainLabel = ExplainLabel + ": "
            End If
            
            'Explain the line!
            If TempReminder = Empty Then
                If TempOperation = -1 Then
                    'No operation to explain
                    MsgBox "There is not an operation on this line to explain", vbExclamation
                    Exit Sub
                Else
                    'Explain data declaration
                    ExplainString = ExplainLabel & ExplainOperation & " - " & TempReminder
                End If
            Else
                'Explain normal command
                ExplainString = ExplainLabel & ExplainCode & " - " & TempReminder
            End If
            MsgBox ExplainString, vbInformation
        Else
            'Empty line
            MsgBox "There is not an operation on this line to explain", vbExclamation
        End If
    End With
 '***********************************************
End Sub

Sub CheckLine(rtbColour As RichTextBox, ByVal LineNumber As Integer)
    Dim StartPos As Integer
    Dim EndPos As Integer
    
    'this proc colours the syntax of a line
    'input the line number to be checked
    'temp rich text box is used
    
    GetLinePoints rtbColour, LineNumber, StartPos, EndPos
    CheckSection rtbColour, StartPos, EndPos
    
End Sub

Sub CheckLines(rtbColour As RichTextBox, ByVal StartLine As Integer, ByVal EndLine As Integer)
    Dim StartPos As Integer
    Dim EndPos As Integer
    Dim StartOfLine As Integer
    Dim EndOfLine As Integer
    Dim Line As Integer

    
    'this proc checklines sequentially
    'quicker for long code
      With rtbColour
        
        'We workout the start and end points of the line
        '******************************************************
        
        'Set initial pointer values
        StartPos = 1
        EndPos = InStr(.Text, vbLf)

        'Do we need to check the first line?
        If StartLine = 0 Then
            'Set end point if no more text
            If EndPos = 0 Then EndPos = Len(.Text) + 2
            
            'Adjust points
            StartOfLine = StartPos
            EndOfLine = EndPos - 2
            
            'Check the section
            CheckSection rtbColour, StartOfLine, EndOfLine
        End If
    
        For Line = 0 To EndLine - 1
            StartPos = EndPos + 1
            EndPos = InStr(StartPos, .Text, vbLf)
        
            If Line >= StartLine Then
                'No CR found because at end of text
                If EndPos = 0 Then EndPos = Len(.Text) + 2
                
                StartOfLine = StartPos
                EndOfLine = EndPos - 2
                
                CheckSection rtbColour, StartOfLine, EndOfLine
                
            End If
        Next
    
    End With
End Sub

Sub GetLinePoints(RichTextObject As RichTextBox, ByVal LineNumber As Long, ByRef StartOfLine As Integer, ByRef EndOfLine As Integer)
    'Get start and end points of a line
    StartOfLine = GetCharFromLine(RichTextObject, LineNumber)
    EndOfLine = StartOfLine + GetLineLength(RichTextObject, GetCharFromLine(RichTextObject, LineNumber))
    Inc StartOfLine
End Sub
Sub CheckSection(rtbColour As RichTextBox, ByVal StartPos As Integer, ByVal EndPos As Integer)
    Dim CursorPos As Integer
    Dim CursorLen As Integer
    
    
    
    With rtbColour
        
        CursorPos = .SelStart
        CursorLen = .SelLength
    
        'validate line if it is not empty
        If EndPos - StartPos >= 0 Then
        
            'Clear errormessages
            ErrorMessages = Empty
            
            'group line code
            GroupedCode = GroupCode(Mid(.Text, StartPos, EndPos - StartPos + 1))
            
            If ErrorMessages <> Empty Then
                HighlightSection rtbColour, StartPos - 1, EndPos - 1, glbErrorCol
            Else
                               
                'check line for labels and remove them
                CheckLabels True
                
                'If error message with line then highlight
                If ErrorMessages <> Empty Then
                    HighlightSection rtbColour, StartPos - 1, EndPos - 1, glbErrorCol
                Else
                    'Reset pointer to first character
                    SymPos = 1
        
                    'Quick parse check of line
                    ParseCommand True
                    
                    'Command not found
                    If TempOperation = -99 Then
                        ErrorMessages = "Syntax Error"
                    End If
                
                    'Highlight error messages
                    If ErrorMessages <> Empty Then
                         HighlightSection rtbColour, StartPos - 1, EndPos - 1, glbErrorCol
                    Else
                        '*****
                        '*****
                        '*****
                        '*****
                        'HighlightLine rtbcolourLineNumber, vbBlack
                        '*****
                        '*****
                        '***** todo: Do this bit when colour syntaxing is turned off
                        
                        GetSym False
                        If Sym = "NEWLINE" Or Sym = "END" Then
                            ColourLine rtbColour, StartPos, EndPos
                        Else
                            HighlightSection rtbColour, StartPos - 1, EndPos - 1, glbErrorCol
                        End If
                    End If
                End If
            End If
        End If
        
        .SelStart = CursorPos
        .SelLength = CursorLen
        
    End With
End Sub


Sub CheckLabels(ByVal QuickCheck As Boolean)
    'checks if labels are present on line
    'and stores the label information
    
    'Quick check input used to determine
    'whether quick syntax check is to be done
    
    Dim TempLabel As String
    Dim StartOfLine As Integer
    Dim LeftChunk As String
    Dim RightChunk As String
    Dim LineLength As Integer
    Dim OldCodeLength As Integer
    Dim Found As Boolean
    Dim Pos2 As Integer
    Dim Pos As Integer
     
    'Remove all labels and variables declarations
    StartOfLine = 1
    SymPos = 1
    LineNumber = 0
    WindowLineNumber = 0
    
    Do
        
        'get symbol
        GetSym False
        TempLabel = NonSym
        TempVariable = NonSym
        TempLabelName = NonSym
        
        If Sym = "ALPHANUMERIC" Then

            GetSym False
            'Look for colon
            If Sym <> "COLON" Then
                ErrorMessages = "Syntax Error"
                Exit Sub
            Else
                GetSym False
                If Sym = "DATAI" Then
                    'Store variable info
                    GetSym False
                    If Sym = "NUMBER" Then
                                                
                        'Look for duplicate variables
                        If QuickCheck = False Then
                            
                            'Check variable name not too long
                            If Len(TempLabel) <= 16 Then
                                If IsVALValid(TempLabel) = True Then
                                    
                                    'Check if variable name has already been used
                                    Found = False
                                    For Pos2 = 0 To VariableCount - 1
                                        If TempLabel = VariableName(Pos2) Then
                                            Found = True
                                            Exit For
                                        End If
                                    Next
                                    
                                    'Name has been used before
                                    If Console = False Then
                                        If Found = True Then
                                            ErrorMessages = "Duplicate variables not allowed"
                                            Exit Sub
                                        End If
                                    End If
                                                                                
                                    'Check initial value
                                    If Val(NonSym) < -32768 Or Val(NonSym) > 32767 Then
                                        ErrorMessages = "Initial value too large or too small"
                                    Else
                                        If Found = True Then
                                            'User wishes to change var value
                                            VariableValue(Pos2) = Val(NonSym)
                                            ShowVariableValues
                                        Else
                                                                            'Check array size
                                            If VariableCount > 255 Then
                                                ErrorMessages = "You have too many variables declared - the maximum is 256"
                                                Exit Sub
                                            Else
                                                VariableName(VariableCount) = TempLabel
                                                VariableValue(VariableCount) = Val(NonSym)
                                                Inc VariableCount
                                                       
                                                'Update variable screen
                                                If Console = True Then
                                                    ShowVariableNames
                                                    ShowVariableValues
                                                End If
                                            End If
                                        End If
                                    End If
                                    TempLabelName = Empty
                                Else
                                    ErrorMessages = "Invalid variable name"
                                    Exit Sub
                                End If
                            Else
                                ErrorMessages = "Variable name too long"
                                Exit Sub
                            End If
                        Else
'todo: Inserted for explanation
                            ExplainCode = TempLabel & ": datai " & NonSym
                            TempReminder = "Declare a variable called " & TempLabel & _
                            " with initial value " & NonSym
                        End If
                        
                        'Check if next symbol is valid
                        GetSym True
                        If Sym <> "END" And Sym <> "NEWLINE" Then
                            ErrorMessages = "Syntax Error"
                            Exit Sub
                        End If
                        
                        
                        TempLabelName = Empty
                        
                        'remove label from line
                        LeftChunk = Left(GroupedCode, StartOfLine - 1)
                        LineLength = Len(GroupedCode) - InStr(StartOfLine, GroupedCode, EOLchar)
                        If LineLength = Len(GroupedCode) Then
                            GroupedCode = LeftChunk
                        Else
                            RightChunk = Right(GroupedCode, LineLength)
                            GroupedCode = LeftChunk + EOLchar + RightChunk
                        End If
                    Else
                        ErrorMessages = "Expected initial value for variable declaration"
                        Exit Sub
                    End If
                ElseIf Sym = "BLOCKI" Then
                    'Store array info
                    GetSym False
                    If Sym = "NUMBER" Then
                        
                        'Look for duplicate arrays
                        If QuickCheck = False Then
                            If Len(TempLabel) <= 16 Then
                                If IsVALValid(TempLabel) = True Then
                                    
                                    'Look for duplicte values
                                    Found = False
                                    For Pos2 = 0 To ArrayCount - 1
                                        If TempLabel = ArrayName(Pos2) Then
                                            Found = True
                                            Exit For
                                        End If
                                    Next
                                    
                                    'Has duplicate been found
                                    If Console = False Then
                                        If Found = True Then
                                            ErrorMessages = "Duplicate arrays not allowed"
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    'Check array size
                                    If Val(NonSym) < 0 Or Val(NonSym) > 256 Then
                                        ErrorMessages = "Array is too big, the maximum is 256"
                                    Else
                                        
                                        If Found = True Then
                                            'User is changing array size
                                            ArrayElements(Pos2) = Val(NonSym)
                                            ShowVariableValues
                                        Else
                                            If ArrayCount > 255 Then
                                                ErrorMessages = "You have too many arrays declared - the maximum is 256"
                                                Exit Sub
                                            Else
                                                ArrayName(ArrayCount) = TempLabel
                                                ArrayElements(ArrayCount) = Val(NonSym)
                                                Inc ArrayCount
                                                
                                                'Update variable screen
                                                If Console = True Then
                                                    ShowVariableNames
                                                    ShowVariableValues
                                                End If
                                            End If
                                        End If
                                    End If
                                    
                                    
                                Else
                                    ErrorMessages = "Invalid array name"
                                    Exit Sub
                                End If
                            Else
                                ErrorMessages = "Array name too long"
                                Exit Sub
                            End If
'todo: Inserted for explanation
                        Else
                            ExplainCode = TempLabel & ": blocki " & NonSym
                            TempReminder = "Declare an array called " & TempLabel & _
                            " with " & NonSym & " elements"
                        End If
                        
                        
                        'Check if next symbol is valid
                        GetSym True
                        If Sym <> "END" And Sym <> "NEWLINE" Then
                            ErrorMessages = "Syntax Error"
                            Exit Sub
                        End If
                        
                        TempLabelName = Empty
                        
                        'remove label from line
                        LeftChunk = Left(GroupedCode, StartOfLine - 1)
                        LineLength = Len(GroupedCode) - InStr(StartOfLine, GroupedCode, EOLchar)
                        If LineLength = Len(GroupedCode) Then
                            GroupedCode = LeftChunk
                        Else
                            RightChunk = Right(GroupedCode, LineLength)
                            GroupedCode = LeftChunk + EOLchar + RightChunk
                        End If
                    Else
                        ErrorMessages = "Expected upperbound for array"
                        Exit Sub
                    End If
                ElseIf Sym = "INVALID" Then
                    ErrorMessages = "Invalid character(s): " + Sym
                    Exit Sub
                ElseIf Sym = "NUMBER" Or Sym = "ALPHANUMERIC" Then
                    ErrorMessages = "Syntax error"
                    Exit Sub
                Else
                    'labels
                    
                    'Has method been called from direct mode?
                    If Console = True Then
                        ErrorMessages = "Invalid in Console"
                        Exit Sub
                    End If
                    
                    'Look for duplicate labels
                    If QuickCheck = False Then
                        If Len(TempLabel) <= 16 Then
                            If IsVALValid(TempLabel) = True Then
                                Found = False
                                For Pos2 = 0 To LabelCount - 1
                                    If TempLabel = LabelName(Pos2) Then Found = True
                                Next
                                If Found = True Then
                                    ErrorMessages = "Duplicate labels not allowed"
                                    Exit Sub
                                End If
                                
                                'Redim the array
                                If LabelCount >= UBound(LabelName()) Then
                                    ReDim LabelName(LabelCount)
                                    ReDim LabelPos(LabelCount)
                                End If
                        
                                'store label info
                                LabelName(LabelCount) = TempLabel
                                LabelPos(LabelCount) = LineNumber
                                Inc LabelCount
                            Else
                                ErrorMessages = "Invalid label name"
                                Exit Sub
                            End If
                        Else
                            ErrorMessages = "Label too long"
                            Exit Sub
                        End If
                        
                        ExplainLabel = TempLabel
                    End If
                    
                    TempVariable = Empty
                    
                    'remove label from line
                    OldCodeLength = Len(GroupedCode)
                    LeftChunk = Left(GroupedCode, StartOfLine - 1)
                    RightChunk = Right(GroupedCode, Len(GroupedCode) - StartOfLine - Len(TempLabel) - 2)
                    GroupedCode = LeftChunk & RightChunk
                    
                    'IF next sym is keyword then
                    'add line number
                    If NonSym = "KEYWORD" Then
                        If QuickCheck = False Then
                            Inc LineNumber
                        End If
                    ElseIf Sym <> "NEWLINE" And Sym <> "END" Then
                        ErrorMessages = "Syntax Error"
                        Exit Sub
                    End If
                    
                End If
            End If
        ElseIf Sym = "INVALID" Then
            ErrorMessages = " invalid character(s): " + Sym
            Exit Sub
        ElseIf Sym = "END" Then
            Exit Do
        ElseIf NonSym = "KEYWORD" Then
            If QuickCheck = False Then
                Inc LineNumber
            End If
            TempVariable = Empty
        End If
        
        Inc WindowLineNumber
        
        'find next new line
        StartOfLine = InStr(StartOfLine, GroupedCode, EOLchar) + 1
        If StartOfLine = 1 Then
            Exit Do
        Else
            Inc StartOfLine
            SymPos = StartOfLine
        End If
        
    Loop
End Sub



Function GetLineCount(rtbLineCount As RichTextBox) As Integer
    'Gets number of lines in
    'program rich text box
    GetLineCount = SendMessage(rtbLineCount.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
End Function

Function GetCharFromLine(txtBox As RichTextBox, LineIndex As Long) As Long
    'Gets the first character on a line
    GetCharFromLine = SendMessage(txtBox.hwnd, EM_LINEINDEX, LineIndex, 0&)
End Function



Public Function GetLineLength(txtBox As RichTextBox, CharPos As Long) As Long
    'Get the length of a line
    GetLineLength = SendMessage(txtBox.hwnd, EM_LINELENGTH, CharPos, 0&)
End Function


Sub CheckSelectionLength()
    'Check length of selection and enable cut, copy
    'delete options
    
    'Dont check if next char is vbCR or if cursor at end
    'of text box
    With frmMain
        If rtbProgram.SelLength > 0 Then
            If rtbProgram.SelStart < Len(rtbProgram.Text) Then
                If Mid(frmEditor.rtbProgram.Text, frmEditor.rtbProgram.SelStart + 1, 1) <> vbCr Then
                    If CutCopyVisible = False Then
                        .mnuEditCut.Enabled = True
                        .mnuEditCopy.Enabled = True
                        .mnuEditDelete.Enabled = True
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

Private Sub rtbProgram_SelChange()

    'If selection is long enough activiate menu items
    CheckSelectionLength

End Sub


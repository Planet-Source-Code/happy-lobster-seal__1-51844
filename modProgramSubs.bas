Attribute VB_Name = "modProgramSubs"
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

Sub Main()
    'Show startup screen
    Load frmSplash
    frmSplash.Show
   ' AlwaysOnTop frmSplash, True
            
    'Set working directory to application directory
    ChDir App.Path
    
    'Set default variables
    OptionsEnabled = False
    CodeDirty = False
    StepMode = False
    
    ReDim Operation(intOperationIndexes)
    ReDim Operand1(intOperationIndexes)
    ReDim Operand2(intOperationIndexes)
    ReDim Operand3(intOperationIndexes)
    ReDim OperandText(intOperationIndexes)
    
    ReDim LabelName(intLabelIndexes)
    ReDim LabelPos(intLabelIndexes)
      
    'User options
    
    'Set all the user options for the first time
    If GetSetting(App.Title, "Options", "runbefore") = Empty Then
        'Fonts
        glbFont = "FixedSys"
        glbFontSize = 9

        'Misc
        glbColourSyntax = True
        glbAutoCheckSyntax = True
        glbClearScreen = True

        'Assign colours
        glbPunctuationCol = vbBlack
        glbLabelCol = vbBlack
        glbVariableCol = vbBlack
        glbCommandCol = vbBlue
        glbRegisterCol = vbGreen
        glbDeviceCol = vbCyan
        glbLiteralCol = vbBlack
        glbEditorBackColour = vbWindowBackground
        glbConsoleBackColour = vbBlack
        glbConsoleTextColour = vbCyan
        glbProgramTextColour = vbGreen
        glbNumberCol = vbBlack
        glbCommentCol = vbMagenta
        glbErrorCol = vbRed
        glbProgramBackColour = vbBlack
        
        'Screens loaded
        glbShowCodeInMemory = True
        glbShowConsole = True
        glbShowLocationTable = True
        glbShowComputerArchitecture = True
        
        'Animation options
        glbAnimateType = intStepped
        glbAnimateSpeed = 60
        glbAnimateStepWait = 500
        
        'Store the fact this is first time
        glbFirstTimeRun = True
        
        'Full screen options
        glbScreenSize = 0
           
        'Setup the form
        Load frmMain
        
        'Save settings
        SaveSettings
        
        'Save window setting
        SaveSetting App.Title, "Window", "WindowState", vbMaximized
        
        'Set the flag
        SaveSetting App.Title, "Options", "runbefore", "true"
    Else
        'Load the settings
        LoadSettings
    End If
    
End Sub



Sub WriteRecentFiles(OPENFILENAME)
    ' This procedure uses the SaveSettings statement to write the names of
    ' recently opened files to the System registry. The SaveSetting
    ' statement requires three parameters. Two of the parameters are
    ' stored as constants and are defined in this module.  The GetAllSettings
    ' function is used in the GetRecentFiles procedure to retrieve the
    ' file names stored in this procedure.
    
    Dim i As Integer
    Dim J As Integer
    Dim strFile As Variant
    Dim Key As String
    Dim FilePos As Integer
    
    'Look for file name position in list
    For i = 1 To 4
        Key = "RecentFile" & i
        strFile = GetSetting(App.Title, "RecentFiles", Key)
        If strFile = OPENFILENAME Then FilePos = i
    Next i
    
    If FilePos = 0 Then
        
        'File not found move files down line
        For i = 3 To 1 Step -1
            Key = "RecentFile" & i
            strFile = GetSetting(App.Title, "RecentFiles", Key)
            If strFile <> "" Then
                Key = "RecentFile" & (i + 1)
                SaveSetting App.Title, "RecentFiles", Key, strFile
            End If
        Next
        
        ' Write the open file to first recent file.
        SaveSetting App.Title, "RecentFiles", "RecentFile1", OPENFILENAME
        
    ElseIf FilePos <> 1 Then
      
      'Sort files
      For i = FilePos - 1 To 1 Step -1
          Key = "RecentFile" & i
          strFile = GetSetting(App.Title, "RecentFiles", Key)
          If strFile <> "" Then
              Key = "RecentFile" & (i + 1)
              SaveSetting App.Title, "RecentFiles", Key, strFile
          End If
      Next
          
      ' Write the open file to first recent file.
      SaveSetting App.Title, "RecentFiles", "RecentFile1", OPENFILENAME
    
    End If
End Sub

Sub GetRecentFiles()
    ' This procedure demonstrates the use of the GetAllSettings function,
    ' which returns an array of values from the Windows registry. In this
    ' case, the registry contains the files most recently opened.  Use the
    ' SaveSetting statement to write the names of the most recent files.
    ' That statement is used in the WriteRecentFiles procedure.
    Dim i As Integer
    Dim J As Integer
    Dim varFiles As Variant ' Varible to store the returned array.
    
    ' Get recent files from the registry using the GetAllSettings statement.
    ' ThisApp and ThisKey are constants defined in this module.
    If GetSetting(App.Title, "RecentFiles", "RecentFile1") = Empty Then Exit Sub
    
    varFiles = GetAllSettings(App.Title, "RecentFiles")
    
    With frmMain
    
        For i = 0 To UBound(varFiles, 1)
            .mnuRecentFile(i + 1).Caption = "&" & Format(i + 1) & " " & varFiles(i, 1)
            .mnuRecentFile(i + 1).Visible = True
       Next i
        
        'Hide no recent files menu
        .mnuNoRecentFiles.Visible = False
    
    End With
End Sub

Sub UpdateFileMenu(Filename)
        ' Check if the open filename is already in the File menu control array.
       ' If OnRecentFilesList(Filename) = False Then
            ' Write open filename to the registry.
            WriteRecentFiles (Filename)
      '  End If
        ' Update the list of the most recently opened files in the File menu control array.
        GetRecentFiles
End Sub

Function OnRecentFilesList(Filename) As Boolean
    Dim i As Integer
    
    'Look for file name in list
    For i = 1 To 4
        If frmMain.mnuRecentFile(i).Caption = Filename Then
            OnRecentFilesList = True
            Exit Function
        End If
    Next i
    OnRecentFilesList = False
End Function

Sub ShowOptions()
    'Shows options available when editting a program
    ToggleOptions True
    NotRunningProgram 'Disable options when not running a program
    OptionsEnabled = True
End Sub

Sub HideOptions()
    'Hides options that are only avaiable when editting
    ToggleOptions False
    OptionsEnabled = False
End Sub

Sub ToggleOptions(ByVal ToggleValue As Boolean)
    With frmMain
        'Enable/disable menu options
        .mnuFileSave.Enabled = ToggleValue
        .mnuFileSaveAs.Enabled = ToggleValue
        
        If ToggleValue = True Then
            EnablePrint 0
        Else
            DisablePrint 0
        End If
        
        .mnuEditPaste.Enabled = ToggleValue
        .tlbOptions.Buttons(11).Enabled = ToggleValue
        .tlbOptions.Buttons(12).Enabled = ToggleValue
        .tlbOptions.Buttons(13).Enabled = ToggleValue
        
        If ToggleValue = False Then
            .mnuEditCut.Enabled = ToggleValue
            .mnuEditCopy.Enabled = ToggleValue
            .mnuEditDelete.Enabled = ToggleValue

            .tlbOptions.Buttons(9).Enabled = ToggleValue
            .tlbOptions.Buttons(10).Enabled = ToggleValue
        End If
        .mnuEditSelectAll.Enabled = ToggleValue
        .mnuEditFind.Enabled = ToggleValue
        .mnuEditFindNext.Enabled = ToggleValue
        .mnuEditReplace.Enabled = ToggleValue
        .mnuEditExplainLine.Enabled = ToggleValue
        .mnuEditInsert.Enabled = ToggleValue
        .mnuViewCode.Enabled = ToggleValue
        .lblCode.Enabled = ToggleValue
        .mnuProgramCodeStatistics.Enabled = ToggleValue
        .mnuProgramTest.Enabled = ToggleValue
        .mnuProgramRun.Enabled = ToggleValue
        .mnuProgramStep.Enabled = ToggleValue
        .mnuProgramStop.Enabled = ToggleValue
        .mnuProgramRestart.Enabled = ToggleValue
        .mnuProgramRunFullScreen.Enabled = ToggleValue

        'Enable menu buttons
        .tlbOptions.Buttons(3).Enabled = ToggleValue
        .tlbOptions.Buttons(5).Enabled = ToggleValue
        .tlbOptions.Buttons(7).Enabled = ToggleValue
        .tlbOptions.Buttons(14).Enabled = ToggleValue
        .tlbOptions.Buttons(15).Enabled = ToggleValue
        .tlbOptions.Buttons(16).Enabled = ToggleValue
        .tlbOptions.Buttons(17).Enabled = ToggleValue
    End With
End Sub

Function CancelOperation() As Boolean
    'Check if user wishes to Save an unsaved file or cancel operation
    
    Dim Msg As String
    Dim intResponse As Integer
    
    If CodeDirty = True Then    'Code needs to be saved
        
        Msg = "Save changes to " & GetFileTitle(frmEditor.Caption)
        intResponse = MsgBox(Msg, vbYesNoCancel + vbExclamation) 'todo: put in title
        
        'See what user has chosen
        Select Case intResponse
            Case vbYes  'User chose yes. Invoke save
            
                If SaveFile() = True Then
                    CancelOperation = False 'User saved file
                Else
                    CancelOperation = True  'User cancelled save
                End If
            Case vbNo   ' User chose No. Continue with operation.
                CancelOperation = False
            Case vbCancel   ' User chose Cancel. Cancel the operation.
                CancelOperation = True
        End Select
    Else
        CancelOperation = False
    End If
End Function


Public Function ValidateWholeWord(PrevLetter As String, NextLetter As String) As Boolean
    'Validate word
   Dim sLetters As String
   ValidateWholeWord = True
   sLetters = "abcdefghijklmnoprqstuvwxyz1234567890"
    
   If InStr(1, sLetters, PrevLetter, vbTextCompare) Or InStr(1, sLetters, NextLetter, vbTextCompare) Then ValidateWholeWord = False
End Function

Function Leftify(ByVal str As String, ByVal intLength) As String
    'Leftifies with padding
    Leftify = str + Space(intLength - Len(str))
End Function


Function ValidateLastLine(Optional blnHideErrorMessage As Boolean) As Boolean
    'Validate line on editor form
    If glbEditorVisible = True Then
        If glbLastLineValidated = False Then
            ValidateLastLine = frmEditor.ParseLine(blnHideErrorMessage)
        End If
    End If
End Function

Function GetVariablesList() As String
    'Returns the variables in memory
    Dim Row As Integer
    Dim V As Integer
    Dim E As Integer
    Dim TempStr As String
    Dim intL As Integer
    
    'Workout length of longest variable
    intL = 4 'for indx
    For V = 0 To VariableCount - 1
        If Len(VariableName(V)) > intL Then
            intL = Len(VariableName(V))
        End If
    Next
    For V = 0 To ArrayCount - 1
        If Len(ArrayName(V)) + 2 > intL Then
            intL = Len(ArrayName(V)) + 2
        End If
    Next
    
    Add intL, 3 'Padding
    
    'Print registers
    AppendString GetVariablesList, Leftify("Acc", intL) + Format(Acc) + vbCrLf
    AppendString GetVariablesList, Leftify("Indx", intL) + Format(Indx) + vbCrLf
    AppendString GetVariablesList, Leftify("Flag", intL) + GetFlagValue + vbCrLf

    'Print variables
    For V = 0 To VariableCount - 1
        AppendString GetVariablesList, Leftify(VariableName(V), intL) + Format(VariableValue(V)) + vbCrLf
    Next
    
    'Print arrays
    For V = 0 To ArrayCount - 1
        TempStr = Empty
        For E = 0 To ArrayElements(V)
            TempStr = TempStr + Format(ArrayValue(V, E)) + "  "
        Next
        AppendString GetVariablesList, Leftify(ArrayName(V), intL) + TempStr
    Next
    
End Function

Sub ShowVariableNames()
    'Show variable names on form
    frmVariables.ShowVariableNames
End Sub

Sub ShowVariableValues()
    'Show variable values on form
    frmVariables.ShowVariableValues
End Sub

Sub SetupTempRTF()
    'Setup temp text box with RTF
     With frmEditor.rtbTemp
        .Text = Empty
        .SelStart = 0
        .SelLength = 2
        .SelFontName = glbFont
        .SelFontSize = glbFontSize
        .SelColor = vbBlack 'todo: not black but variable colour
    End With
End Sub
Sub ShowAWindow(frmWindow As Form)
    'Window is being shown
    
    'If window hiden increase window count
    If frmWindow.Visible = False Then
        Inc NumberOfWindows
    End If
    frmMain.mnuWindow.Visible = True
    
End Sub
Sub HideAWindow()
    'A form has been closed so do we need
    'to check wether to disable the window menu
    Dec NumberOfWindows
    
    If NumberOfWindows = 0 Then
        frmMain.mnuWindow.Visible = False
    End If
    
End Sub
Sub AppendString(ByRef strVariable As String, ByVal strAppend As String)
    'Sub to append strings
    strVariable = strVariable + strAppend
End Sub

Sub Add(ByRef VariableNum, ByVal inValue As Integer)
    'Add
    VariableNum = VariableNum + inValue
End Sub
Sub Multiply(ByRef VariableNum, ByVal inValue As Integer)
    'Mulityply
    VariableNum = VariableNum * inValue
End Sub

Sub DisablePrint(ByVal intWindow As Integer)
    'Disables printer options window
    With frmPrint
        .cmbPrint(intWindow).Enabled = False
        .lblWindow(intWindow).Enabled = False
    End With
End Sub
Sub EnablePrint(ByVal intWindow As Integer)
    'Enable printer options window
    With frmPrint
        .cmbPrint(intWindow).Enabled = True
        .lblWindow(intWindow).Enabled = True
    End With
End Sub


Sub RunningProgram()
    'Enable/disable run options
    With frmMain
        .mnuProgramRun.Enabled = False
        .mnuProgramStop.Enabled = True
        .mnuProgramRestart.Enabled = True
        .mnuProgramTest.Enabled = False
        .mnuProgramRunFullScreen.Enabled = False

        .mnuProgramCodeStatistics.Enabled = False
        .tlbOptions.Buttons(13).Enabled = False
        .tlbOptions.Buttons(15).Enabled = True
        .tlbOptions.Buttons(17).Enabled = False
        .mnuFileNew.Enabled = False
        .mnuFileOpen.Enabled = False
        .mnuFileSave.Enabled = False
        .mnuFileSaveAs.Enabled = False
        .tlbOptions.Buttons(1).Enabled = False
        .tlbOptions.Buttons(2).Enabled = False
        .tlbOptions.Buttons(3).Enabled = False
        .mnuNoRecentFiles.Enabled = False
        .mnuRecentFile(1).Enabled = False
        .mnuRecentFile(2).Enabled = False
        .mnuRecentFile(3).Enabled = False
        .mnuRecentFile(4).Enabled = False
        
        frmEditor.rtbProgram.Locked = True
    End With
End Sub

Sub NotRunningProgram()
    'Enable/disable run options
    With frmMain
        .mnuProgramRun.Enabled = True
        .mnuProgramStop.Enabled = False
        .mnuProgramStep.Enabled = True
        .mnuProgramRestart.Enabled = False
        .mnuProgramTest.Enabled = True
        .mnuProgramRunFullScreen.Enabled = True
        .mnuProgramCodeStatistics.Enabled = True
        .tlbOptions.Buttons(13).Enabled = True
        .tlbOptions.Buttons(15).Enabled = False
        .tlbOptions.Buttons(14).Enabled = True
        .tlbOptions.Buttons(17).Enabled = True
         .mnuFileNew.Enabled = True
        .mnuFileOpen.Enabled = True
        .mnuFileSave.Enabled = True
        .mnuFileSaveAs.Enabled = True
        .tlbOptions.Buttons(1).Enabled = True
        .tlbOptions.Buttons(2).Enabled = True
        .tlbOptions.Buttons(3).Enabled = True
        .mnuNoRecentFiles.Enabled = True
        .mnuRecentFile(1).Enabled = True
        .mnuRecentFile(2).Enabled = True
        .mnuRecentFile(3).Enabled = True
        .mnuRecentFile(4).Enabled = True
        
        frmEditor.rtbProgram.Locked = False
    End With
End Sub

Public Sub SetAllFonts()
    'Sets all the fonts and sizes
    SetFont frmEditor.rtbProgram
    SetFont frmRun.rtbOutput
    SetFont frmConsole.rtbOutput
    SetFont frmFullScreen.rtbOutput
    frmVariables.lstvVariables.Font = glbFont
    frmCode.lstCode.Font = glbFont
    frmCode.lstCode.FontSize = glbFontSize
End Sub


Private Sub SetFont(rtbBox As RichTextBox)
    'Set the font and size of a rtb
    Dim SelPos As Long
    Dim SelLen As Long
    
    With rtbBox
        SelPos = .SelStart
        SelLen = .SelLength
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelFontName = glbFont
        .SelFontSize = glbFontSize
        .SelStart = SelPos
        .SelLength = SelLen
    End With
    
End Sub

Sub CopyAnimateScreen()
    'Copy animate screen
    If frmAnimate.rtbScreen.SelLength = 0 Then
        Clipboard.SetText frmAnimate.rtbScreen.Text
    Else
        Clipboard.SetText frmAnimate.rtbScreen.SelText
    End If
End Sub

Sub CopyKeyboard()
    'Copy animate screen
    If frmAnimate.txtKeyboard.SelLength = 0 Then
        Clipboard.SetText frmAnimate.txtKeyboard.Text
    Else
        Clipboard.SetText frmAnimate.txtKeyboard.SelText
    End If
End Sub

Sub CopyConsoleScreen()
    'Copy console screen
    If frmConsole.rtbOutput.SelLength = 0 Then
        Clipboard.SetText frmConsole.rtbOutput.Text
    Else
        Clipboard.SetText frmConsole.rtbOutput.SelText
    End If
End Sub

Sub CopyRunScreen()
    'Copy run screen
    If frmRun.rtbOutput.SelLength = 0 Then
        Clipboard.SetText frmRun.rtbOutput.Text
    Else
        Clipboard.SetText frmRun.rtbOutput.SelText
    End If
End Sub

Sub CopyCodeInMemory()
    'Copy code in memory
    Dim intLine As Integer
    Dim strTemp As String
    Dim strCR As String
    'Populate
    For intLine = 0 To frmCode.lstCode.ListCount - 1
        If intLine < frmCode.lstCode.ListCount - 1 Then
            AppendString strTemp, frmCode.lstCode.List(intLine) & vbCrLf
        Else
            'Don't add the last CR
            AppendString strTemp, frmCode.lstCode.List(intLine)
        End If
    Next
    
    Clipboard.SetText strTemp
End Sub

Sub CopyVariables()
    'Copy variables window
    Clipboard.SetText GetVariablesList
End Sub

Sub LoadSettings()
    'Load all the programs setting
        
    'Font stuff
    glbFontSize = Val(GetSetting(App.Title, "Font", "FontSize"))
    glbFont = GetSetting(App.Title, "Font", "Font")
    
    'Miscallaenous
    glbColourSyntax = CBool(GetSetting(App.Title, "Misc", "ColourSyntax"))
    glbAutoCheckSyntax = CBool(GetSetting(App.Title, "Misc", "AutoCheckSyntax"))
    glbClearScreen = CBool(GetSetting(App.Title, "Misc", "ClearScreen"))
    
    'Forms loaded
    glbShowCodeInMemory = CBool(GetSetting(App.Title, "Forms", "CodeInMemory"))
    glbShowConsole = CBool(GetSetting(App.Title, "Forms", "Console"))
    glbShowLocationTable = CBool(GetSetting(App.Title, "Forms", "LocationTable"))
    glbShowComputerArchitecture = CBool(GetSetting(App.Title, "Forms", "ComputerArchitecture"))

    'Colours
    glbPunctuationCol = CLng(GetSetting(App.Title, "Colours", "PunctuationCol"))
    glbLabelCol = CLng(GetSetting(App.Title, "Colours", "LabelCol"))
    glbVariableCol = CLng(GetSetting(App.Title, "Colours", "VariableCol"))
    glbCommandCol = CLng(GetSetting(App.Title, "Colours", "CommandCol"))
    glbRegisterCol = CLng(GetSetting(App.Title, "Colours", "RegisterCol"))
    glbDeviceCol = CLng(GetSetting(App.Title, "Colours", "DeviceCol"))
    glbNumberCol = CLng(GetSetting(App.Title, "Colours", "NumberCol"))
    glbCommentCol = CLng(GetSetting(App.Title, "Colours", "CommentCol"))
    glbErrorCol = CLng(GetSetting(App.Title, "Colours", "ErrorCol"))
    glbLiteralCol = CLng(GetSetting(App.Title, "Colours", "LiteralCol"))
    glbEditorBackColour = CLng(GetSetting(App.Title, "Colours", "EditorBackColour"))
    glbConsoleBackColour = CLng(GetSetting(App.Title, "Colours", "ConsoleBackColour"))
    glbConsoleTextColour = CLng(GetSetting(App.Title, "Colours", "ConsoleTextColour"))
    glbProgramBackColour = CLng(GetSetting(App.Title, "Colours", "ProgramBackColour"))
    glbProgramTextColour = CLng(GetSetting(App.Title, "Colours", "ProgramTextColour"))
       
    'Animation
    glbAnimateType = Int(GetSetting(App.Title, "Animation", "AnimateType"))
    glbAnimateSpeed = Int(GetSetting(App.Title, "Animation", "AnimateSpeed"))
    glbAnimateStepWait = Int(GetSetting(App.Title, "Animation", "AnimateStepWait"))
    
    'Screen
    glbScreenSize = Val(GetSetting(App.Title, "Screen", "Size"))


End Sub

Sub SetupChildWindows()
    'Setup child windows
    frmEditor.WindowState = GetSetting(App.Title, "EditorForm", "WindowState")
    If frmEditor.WindowState = vbNormal Then
        frmEditor.Left = Val(GetSetting(App.Title, "EditorForm", "Left"))
        frmEditor.Top = Val(GetSetting(App.Title, "EditorForm", "Top"))
        frmEditor.Width = Val(GetSetting(App.Title, "EditorForm", "Width"))
        frmEditor.Height = Val(GetSetting(App.Title, "EditorForm", "Height"))
    End If
    frmCode.WindowState = GetSetting(App.Title, "CodeInMemoryForm", "WindowState")
    If frmCode.WindowState = vbNormal Then
        frmCode.Left = Val(GetSetting(App.Title, "CodeInMemoryForm", "Left"))
        frmCode.Top = Val(GetSetting(App.Title, "CodeInMemoryForm", "Top"))
        frmCode.Width = Val(GetSetting(App.Title, "CodeInMemoryForm", "Width"))
        frmCode.Height = Val(GetSetting(App.Title, "CodeInMemoryForm", "Height"))
    End If
    frmConsole.WindowState = Val(GetSetting(App.Title, "ConsoleForm", "WindowState"))
    If frmConsole.WindowState = vbNormal Then
        frmConsole.Left = Val(GetSetting(App.Title, "ConsoleForm", "Left"))
        frmConsole.Top = Val(GetSetting(App.Title, "ConsoleForm", "Top"))
        frmConsole.Width = Val(GetSetting(App.Title, "ConsoleForm", "Width"))
        frmConsole.Height = Val(GetSetting(App.Title, "ConsoleForm", "Height"))
    End If
    frmVariables.WindowState = Val(GetSetting(App.Title, "LocationTableForm", "WindowState"))
    If frmVariables.WindowState = vbNormal Then
        frmVariables.Left = Val(GetSetting(App.Title, "LocationTableForm", "Left"))
        frmVariables.Top = Val(GetSetting(App.Title, "LocationTableForm", "Top"))
        frmVariables.Width = Val(GetSetting(App.Title, "LocationTableForm", "Width"))
        frmVariables.Height = Val(GetSetting(App.Title, "LocationTableForm", "Height"))
    End If
    frmAnimate.WindowState = Val(GetSetting(App.Title, "ComputerArchitectureForm", "WindowState"))
    If frmAnimate.WindowState = vbNormal Then
        frmAnimate.Left = Val(GetSetting(App.Title, "ComputerArchitectureForm", "Left"))
        frmAnimate.Top = Val(GetSetting(App.Title, "ComputerArchitectureForm", "Top"))
        frmAnimate.Width = Val(GetSetting(App.Title, "ComputerArchitectureForm", "Width"))
        frmAnimate.Height = Val(GetSetting(App.Title, "ComputerArchitectureForm", "Height"))
    End If
    
    frmRun.Left = Val(GetSetting(App.Title, "ProgramOutputForm", "Left"))
    frmRun.Top = Val(GetSetting(App.Title, "ProgramOutputForm", "Top"))
    frmRun.Width = Val(GetSetting(App.Title, "ProgramOutputForm", "Width"))
    frmRun.Height = Val(GetSetting(App.Title, "ProgramOutputForm", "Height"))
    
End Sub

Sub SaveWindowSettings()
    Dim intTempWindow As Integer
    'Save the current window state
    SaveSetting App.Title, "Window", "WindowState", Format(frmMain.WindowState)
    If frmMain.WindowState = vbNormal Then
        SaveSetting App.Title, "Window", "Left", Format(frmMain.Left)
        SaveSetting App.Title, "Window", "Top", Format(frmMain.Top)
        SaveSetting App.Title, "Window", "Width", Format(frmMain.Width)
        SaveSetting App.Title, "Window", "Height", Format(frmMain.Height)
    End If
End Sub

Sub LoadAndSetWindowSettings()
    Dim intTempWindow As Integer
    'Load and set last window pos
    intTempWindow = Val(GetSetting(App.Title, "Window", "WindowState"))
    If intTempWindow = vbNormal Then
        frmMain.Left = Val(GetSetting(App.Title, "Window", "Left"))
        frmMain.Top = Val(GetSetting(App.Title, "Window", "Top"))
        frmMain.Width = Val(GetSetting(App.Title, "Window", "Width"))
        frmMain.Height = Val(GetSetting(App.Title, "Window", "Height"))
    End If
    frmMain.WindowState = intTempWindow
End Sub

Sub SaveSettings()
    'Save all the programs setting
        
    'Font stuff
    SaveSetting App.Title, "Font", "FontSize", Format(glbFontSize)
    SaveSetting App.Title, "Font", "Font", glbFont
    
    'Miscallaenous
    SaveSetting App.Title, "Misc", "ColourSyntax", glbColourSyntax
    SaveSetting App.Title, "Misc", "AutoCheckSyntax", glbAutoCheckSyntax
    SaveSetting App.Title, "Misc", "ClearScreen", glbClearScreen

    'Save form status
    SaveSetting App.Title, "Forms", "CodeInMemory", frmCode.Visible
    SaveSetting App.Title, "Forms", "Console", frmConsole.Visible
    SaveSetting App.Title, "Forms", "LocationTable", frmVariables.Visible
    SaveSetting App.Title, "Forms", "ComputerArchitecture", frmAnimate.Visible

    'Save form positions
    SaveSetting App.Title, "EditorForm", "WindowState", frmEditor.WindowState
    SaveSetting App.Title, "EditorForm", "Left", frmEditor.Left
    SaveSetting App.Title, "EditorForm", "Top", frmEditor.Top
    SaveSetting App.Title, "EditorForm", "Width", frmEditor.Width
    SaveSetting App.Title, "EditorForm", "Height", frmEditor.Height
    SaveSetting App.Title, "ProgramOutputForm", "Left", frmRun.Left
    SaveSetting App.Title, "ProgramOutputForm", "Top", frmRun.Top
    SaveSetting App.Title, "ProgramOutputForm", "Width", frmRun.Width
    SaveSetting App.Title, "ProgramOutputForm", "Height", frmRun.Height
    SaveSetting App.Title, "CodeInMemoryForm", "WindowState", frmCode.WindowState
    SaveSetting App.Title, "CodeInMemoryForm", "Left", frmCode.Left
    SaveSetting App.Title, "CodeInMemoryForm", "Top", frmCode.Top
    SaveSetting App.Title, "CodeInMemoryForm", "Width", frmCode.Width
    SaveSetting App.Title, "CodeInMemoryForm", "Height", frmCode.Height
    SaveSetting App.Title, "ConsoleForm", "WindowState", frmConsole.WindowState
    SaveSetting App.Title, "ConsoleForm", "Left", frmConsole.Left
    SaveSetting App.Title, "ConsoleForm", "Top", frmConsole.Top
    SaveSetting App.Title, "ConsoleForm", "Width", frmConsole.Width
    SaveSetting App.Title, "ConsoleForm", "Height", frmConsole.Height
    SaveSetting App.Title, "LocationTableForm", "WindowState", frmVariables.WindowState
    SaveSetting App.Title, "LocationTableForm", "Left", frmVariables.Left
    SaveSetting App.Title, "LocationTableForm", "Top", frmVariables.Top
    SaveSetting App.Title, "LocationTableForm", "Width", frmVariables.Width
    SaveSetting App.Title, "LocationTableForm", "Height", frmVariables.Height
    SaveSetting App.Title, "ComputerArchitectureForm", "WindowState", frmAnimate.WindowState
    SaveSetting App.Title, "ComputerArchitectureForm", "Left", frmAnimate.Left
    SaveSetting App.Title, "ComputerArchitectureForm", "Top", frmAnimate.Top
    SaveSetting App.Title, "ComputerArchitectureForm", "Width", frmAnimate.Width
    SaveSetting App.Title, "ComputerArchitectureForm", "Height", frmAnimate.Height
    
    'Colours
    SaveSetting App.Title, "Colours", "PunctuationCol", Format(glbPunctuationCol)
    SaveSetting App.Title, "Colours", "LabelCol", Format(glbLabelCol)
    SaveSetting App.Title, "Colours", "VariableCol", Format(glbVariableCol)
    SaveSetting App.Title, "Colours", "CommandCol", Format(glbCommandCol)
    SaveSetting App.Title, "Colours", "RegisterCol", Format(glbRegisterCol)
    SaveSetting App.Title, "Colours", "DeviceCol", Format(glbDeviceCol)
    SaveSetting App.Title, "Colours", "NumberCol", Format(glbNumberCol)
    SaveSetting App.Title, "Colours", "CommentCol", Format(glbCommentCol)
    SaveSetting App.Title, "Colours", "ErrorCol", Format(glbErrorCol)
    SaveSetting App.Title, "Colours", "LiteralCol", Format(glbLiteralCol)
    SaveSetting App.Title, "Colours", "EditorBackColour", Format(glbEditorBackColour)
    SaveSetting App.Title, "Colours", "ConsoleBackColour", Format(glbConsoleBackColour)
    SaveSetting App.Title, "Colours", "ConsoleTextColour", Format(glbConsoleTextColour)
    SaveSetting App.Title, "Colours", "ProgramBackColour", Format(glbProgramBackColour)
    SaveSetting App.Title, "Colours", "ProgramTextColour", Format(glbProgramTextColour)
    
    'Animation
    SaveSetting App.Title, "Animation", "AnimateType", Format(glbAnimateType)
    SaveSetting App.Title, "Animation", "AnimateSpeed", Format(glbAnimateSpeed)
    SaveSetting App.Title, "Animation", "AnimateStepWait", Format(glbAnimateStepWait)
    
    'Screen
    SaveSetting App.Title, "Screen", "Size", Format(glbScreenSize)
End Sub

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
    'Put a form always on top
    Dim lFlag As Integer
    
    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.hwnd, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Attribute VB_Name = "modFileSubs"
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

Sub OpenFile(Filename)
    Dim fIndex As Integer
    On Error Resume Next
    Dim Line As Integer
    
    'Open the selected file.
    Open Filename For Input As #1
    If Err Then
        MsgBox "File problem - " & Err.Description, vbCritical
        Close #1
        Exit Sub
    End If

    ' Change the mouse pointer to an hourglass.
    Screen.MousePointer = 11

    With frmEditor
    
        ' Change the form's caption and display the new text.
        .Caption = Filename
        
        SetupTempRTF
        
        .rtbTemp.Text = Input(LOF(1), 1)
        Close #1
        
        'Check number of lines, inform user
        If UBound(Split(.rtbTemp.Text, vbCrLf)) > intMaxLinesColour Then
            glbDisableColour = True
            MsgBox "This program is large so colour syntaxing will be disabled", 48
        Else
            glbDisableColour = False
        End If

        'Check each line of code if colour syntaxing is on
        If glbColourSyntax = True And glbDisableColour = False Then
            .CheckLines .rtbTemp, 0, .GetLineCount(.rtbTemp) - 1
        End If
        
        'Set the program text box
        .rtbProgram.TextRTF = .rtbTemp.TextRTF
    End With
    
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
    
    'Set variables
    CodeDirty = False
    
    UpdateFileMenu (Filename)    'Update file list
    ShowOptions 'Show toolbar options

    frmEditor.Show    'Show form
    
    'Set last line validated
    glbLastLineValidated = True
    glbEditorVisible = True
    

End Sub


Sub SaveFileAs(Filename)
    On Error Resume Next

    ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    
    ' Save the document
    frmEditor.rtbProgram.SaveFile Filename, rtfText

    ' Reset the mouse pointer.
    Screen.MousePointer = 0
    
    ' Set the form's caption.
    If Err Then
        MsgBox Error, 48, App.Title
    Else
        frmEditor.Caption = Filename
        
        'Reset the dirty flag
        CodeDirty = False
    End If
End Sub

Function GetFileName(Filename As Variant)
    ' Display a Save As dialog box and return a filename.
    ' If the user chooses Cancel, return an empty string.
    On Error Resume Next
    frmMain.cmDialog1.Filename = Filename
    frmMain.cmDialog1.ShowSave
    If Err <> 32755 Then    ' User chose Cancel.
        GetFileName = frmMain.cmDialog1.Filename
    Else
        GetFileName = ""
    End If
End Function

Function GetFileTitle(ByVal FilePath As String) As String
    'Return project name from given path
    GetFileTitle = Right(FilePath, Len(FilePath) - InStrR(Len(FilePath), FilePath, "\"))
    
End Function



Function SaveFile() As Boolean
    Dim strFilename As String

    If Left(frmEditor.Caption, 8) = "Untitled" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, GetFileName.
        strFilename = GetFileName(strFilename)
    Else
        ' The form's Caption contains the name of the open file.
        strFilename = frmEditor.Caption
    End If
    ' Call the save procedure. If Filename = Empty, then
    ' the user chose Cancel in the Save As dialog box; otherwise,
    ' save the file.
    If strFilename <> "" Then
        SaveFileAs strFilename
        SaveFile = True
    Else
        SaveFile = False
    End If
End Function

Sub FileOpenProc()

    Dim intRetVal As Variant
    Dim strOpenFileName As String
    Dim FileLength As Long
    
    On Error Resume Next
    
    frmMain.cmDialog1.Filename = ""
    frmMain.cmDialog1.ShowOpen
    
    If Err <> 32755 Then    ' User chose Cancel.
        strOpenFileName = frmMain.cmDialog1.Filename
        ' If the file is larger than 65K, it can't
        ' be opened, so cancel the operation.
        
        If FileLen(strOpenFileName) > 65000 Then
            MsgBox "The file is too large to open.", 48
            Exit Sub
        End If
        
        OpenFile (strOpenFileName)  'Open the file
        
    End If
End Sub

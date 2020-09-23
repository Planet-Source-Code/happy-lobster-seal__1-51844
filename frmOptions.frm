VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4515
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmDialog1 
      Left            =   960
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Flags           =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   3180
      TabIndex        =   0
      Top             =   4140
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4035
      Index           =   1
      Left            =   60
      ScaleHeight     =   4035
      ScaleWidth      =   5625
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Width           =   5625
      Begin VB.CheckBox chkAutoCheckSyntax 
         Caption         =   "Alert syntax errors when typing"
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   3660
         Width           =   2715
      End
      Begin VB.CheckBox chkColourSyntaxing 
         Caption         =   "Use colour syntaxing"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   3420
         Width           =   2415
      End
      Begin VB.Frame Frame5 
         Caption         =   "Colours:"
         Height          =   3975
         Left            =   60
         TabIndex        =   16
         Top             =   0
         Width           =   2955
         Begin VB.ListBox lstTextColours 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2400
            ItemData        =   "frmOptions.frx":000C
            Left            =   180
            List            =   "frmOptions.frx":000E
            TabIndex        =   19
            Top             =   300
            Width           =   2595
         End
         Begin VB.CommandButton cmdSetColour 
            Caption         =   "Edit Colour..."
            Height          =   315
            Left            =   1500
            TabIndex        =   18
            Top             =   2940
            Width           =   1155
         End
         Begin VB.CommandButton cmdSetDefaults 
            Caption         =   "Set Defaults"
            Height          =   315
            Left            =   180
            TabIndex        =   17
            Top             =   2940
            Width           =   1155
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Program Output:"
         Height          =   1215
         Left            =   3180
         TabIndex        =   12
         Top             =   2760
         Width           =   2175
         Begin VB.OptionButton optScreen 
            Caption         =   "640 by 480 Full Mode"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   15
            Top             =   900
            Width           =   1875
         End
         Begin VB.OptionButton optScreen 
            Caption         =   "320 by 240 Full Mode"
            Height          =   210
            Index           =   0
            Left            =   180
            TabIndex        =   14
            Top             =   660
            Width           =   1875
         End
         Begin VB.CheckBox chkClearScreen 
            Caption         =   "Always clear screen"
            Height          =   270
            Left            =   180
            TabIndex        =   13
            Top             =   300
            Width           =   1935
         End
      End
      Begin VB.ComboBox cmbFontSize 
         Height          =   330
         ItemData        =   "frmOptions.frx":0010
         Left            =   3180
         List            =   "frmOptions.frx":0029
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   915
         Width           =   1095
      End
      Begin VB.Frame Frame6 
         Caption         =   "Sample:"
         Height          =   1335
         Left            =   3180
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
         Begin VB.PictureBox picSample 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   180
            ScaleHeight     =   735
            ScaleWidth      =   1755
            TabIndex        =   8
            Top             =   360
            Width           =   1755
            Begin VB.Label lblSample 
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "ABCXYZabcxyz"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   120
               TabIndex        =   9
               Top             =   120
               Width           =   1575
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00808080&
               X1              =   16475
               X2              =   0
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00808080&
               X1              =   0
               X2              =   0
               Y1              =   735
               Y2              =   -15
            End
            Begin VB.Line Line3 
               BorderColor     =   &H80000005&
               X1              =   0
               X2              =   1740
               Y1              =   720
               Y2              =   720
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000005&
               X1              =   1740
               X2              =   1740
               Y1              =   720
               Y2              =   0
            End
         End
      End
      Begin VB.ComboBox cmbFonts 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3180
         TabIndex        =   7
         Top             =   675
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Font:"
         Height          =   255
         Left            =   3180
         TabIndex        =   3
         Top             =   0
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmOptions"
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

Dim tmpPunctuationCol As Long
Dim tmpLabelCol As Long
Dim tmpVariableCol As Long
Dim tmpCommandCol As Long
Dim tmpRegisterCol As Long
Dim tmpDeviceCol As Long
Dim tmpNumberCol As Long
Dim tmpCommentCol As Long
Dim tmpErrorCol As Long
Dim tmpEditorBackCol As Long
Dim tmpLiteralCol As Long
Dim tmpConsoleTextColour As Long
Dim tmpConsoleBackColour As Long
Dim tmpProgramTextColour As Long
Dim tmpProgramBackColour As Long

Private Sub cmbFonts_Click()
    'Change sample font
    lblSample.Font = cmbFonts.Text
End Sub

Private Sub cmbFontSize_Click()
    'Change sample font size
    lblSample.FontSize = Val(cmbFontSize.Text)
End Sub

Private Sub cmdCancel_Click()
    'Cancel button
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim blnColourChanged As Boolean
    Dim blnBackChanged As Boolean
    Dim blnFontChanged As Boolean
    'Formatting ---------
    
    Screen.MousePointer = 13
       
    'Set colours
    If glbPunctuationCol <> tmpPunctuationCol Then
        glbPunctuationCol = tmpPunctuationCol
        blnColourChanged = True
    End If
    If glbLabelCol <> tmpLabelCol Then
        glbLabelCol = tmpLabelCol
        blnColourChanged = True
    End If
    If glbVariableCol <> tmpVariableCol Then
        glbVariableCol = tmpVariableCol
        blnColourChanged = True
    End If
    If glbCommandCol <> tmpCommandCol Then
        glbCommandCol = tmpCommandCol
        blnColourChanged = True
    End If
    If glbRegisterCol <> tmpRegisterCol Then
        glbRegisterCol = tmpRegisterCol
        blnColourChanged = True
    End If
    If glbDeviceCol <> tmpDeviceCol Then
        glbDeviceCol = tmpDeviceCol
        blnColourChanged = True
    End If
    If glbNumberCol <> tmpNumberCol Then
        glbNumberCol = tmpNumberCol
        blnColourChanged = True
    End If
    If glbCommentCol <> tmpCommentCol Then
        glbCommentCol = tmpCommentCol
        blnColourChanged = True
    End If
    If glbErrorCol <> tmpErrorCol Then
        glbErrorCol = tmpErrorCol
        blnColourChanged = True
    End If
    If glbEditorBackColour <> tmpEditorBackCol Then
        glbEditorBackColour = tmpEditorBackCol
        blnBackChanged = True
    End If
    If glbLiteralCol <> tmpLiteralCol Then
        glbLiteralCol = tmpLiteralCol
        blnColourChanged = True
    End If
    If cmbFonts.Text <> glbFont Then
        glbFont = cmbFonts.Text
        blnFontChanged = True
    End If
    If Val(cmbFontSize.Text) <> glbFontSize Then
        glbFontSize = Val(cmbFontSize.Text)
        blnFontChanged = True
    End If
    
    'Change console colours
    If tmpConsoleTextColour <> glbConsoleTextColour Then
        glbConsoleTextColour = tmpConsoleTextColour
    End If
    If tmpConsoleBackColour <> glbConsoleBackColour Then
        glbConsoleBackColour = tmpConsoleBackColour
        frmConsole.rtbOutput.BackColor = glbConsoleBackColour
    End If
    
    'Change program output colours
    If tmpProgramTextColour <> glbProgramTextColour Then
        glbProgramTextColour = tmpProgramTextColour
    End If
    If tmpProgramBackColour <> glbProgramBackColour Then
        glbProgramBackColour = tmpProgramBackColour
    End If

      
    'Use colour syntaxing?
    If chkColourSyntaxing.Value = 1 Then
        If glbColourSyntax = False Then
             blnColourChanged = True
        End If
        glbColourSyntax = True
    Else
        If glbColourSyntax = True Then
             blnColourChanged = True
        End If
        glbColourSyntax = False
    End If
    
    'Set clear screen
    glbClearScreen = CBool(chkClearScreen.Value)
    'Set syntax
    glbAutoCheckSyntax = CBool(chkAutoCheckSyntax.Value)
    'Screen size
    If optScreen(0).Value = True Then
        glbScreenSize = 0
    Else
        glbScreenSize = 1
    End If
    
    'Set editor background color
    If blnBackChanged = True Then
        frmEditor.rtbProgram.BackColor = glbEditorBackColour
    End If
    
    'Set the other colours
    If blnColourChanged = True Then
        With frmEditor
            Dim Line As Long
            
            SetupTempRTF
            
            .rtbTemp.Text = .rtbProgram.Text
            
            'Check each line of code
            If glbColourSyntax = True And glbDisableColour = False Then
                For Line = 0 To .GetLineCount(.rtbTemp) - 1
                    .CheckLine .rtbTemp, Line
                Next
            End If
        
            'Set the program text box
            .rtbProgram.TextRTF = .rtbTemp.TextRTF
        End With
    End If
    
    'Do Fonts
    If blnFontChanged = True Then
        SetAllFonts
    End If
    
    'Do background
    frmEditor.rtbProgram.BackColor = glbEditorBackColour
    
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdSetColour_Click()
    'Get new colour
    Dim TempColour As Long
    
    On Error Resume Next
    
    Select Case lstTextColours.ListIndex
    Case 0
        'Punctuation colour
        cmDialog1.Color = tmpPunctuationCol
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpPunctuationCol = cmDialog1.Color
    Case 1
        'Label colour
       cmDialog1.Color = tmpLabelCol
        cmDialog1.ShowColor
       If Err <> 32755 Then tmpLabelCol = cmDialog1.Color
    Case 2
        'variable colour
        cmDialog1.Color = tmpVariableCol
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpVariableCol = cmDialog1.Color
    Case 3
         'Command colour
        cmDialog1.Color = tmpCommandCol
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpCommandCol = cmDialog1.Color
    Case 4
        'Register colour
        cmDialog1.Color = tmpRegisterCol
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpRegisterCol = cmDialog1.Color
    Case 5
        'Device colour
        cmDialog1.Color = tmpDeviceCol
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpDeviceCol = cmDialog1.Color
    Case 6
        'Number colour
        cmDialog1.Color = tmpNumberCol
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpNumberCol = cmDialog1.Color
    Case 7
        'Comment colour
        cmDialog1.Color = tmpCommentCol
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpCommentCol = cmDialog1.Color
    Case 8
        'Error colour
        cmDialog1.Color = tmpErrorCol
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpErrorCol = cmDialog1.Color
    Case 9
        'Literal colour
        cmDialog1.Color = tmpLiteralCol
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpLiteralCol = cmDialog1.Color
    Case 10
        'Editor background colour
        cmDialog1.Color = tmpLiteralCol
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpEditorBackCol = cmDialog1.Color
    Case 11
        'Editor text colour
        cmDialog1.Color = tmpConsoleTextColour
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpConsoleTextColour = cmDialog1.Color
    Case 12
        'Editor background colour
        cmDialog1.Color = tmpConsoleBackColour
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpConsoleBackColour = cmDialog1.Color
    Case 13
        'run text colour
        cmDialog1.Color = tmpProgramTextColour
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpProgramTextColour = cmDialog1.Color
    Case 14
        'Run background colour
        cmDialog1.Color = tmpProgramBackColour
        cmDialog1.ShowColor
        If Err <> 32755 Then tmpProgramBackColour = cmDialog1.Color
    End Select
    
    lstTextColours_Click
    
End Sub

Private Sub cmdSetDefaults_Click()
    'Set the default coolours
    
    tmpPunctuationCol = vbBlack
    tmpLabelCol = vbBlack
    tmpVariableCol = vbBlack
    tmpCommandCol = vbBlue
    tmpRegisterCol = vbGreen
    tmpDeviceCol = vbCyan
    tmpNumberCol = vbBlack
    tmpCommentCol = vbMagenta
    tmpErrorCol = vbRed
    tmpLiteralCol = vbBlack
    tmpEditorBackCol = vbWindowBackground
    tmpConsoleTextColour = vbCyan
    tmpConsoleBackColour = vbBlack
    tmpProgramTextColour = vbGreen
    tmpProgramBackColour = vbBlack
End Sub



Private Sub Form_Load()
    Dim F As Integer

    'Setup Font list
    With Screen
        For F = 0 To .FontCount - 1
            cmbFonts.AddItem .Fonts(F)
        Next
    End With
    
    'todo: wap in available font sizes for a font

    'Setup colour list box
    With lstTextColours
        .AddItem "Punctuation Text"
        .AddItem "Label Text"
        .AddItem "Variables"
        .AddItem "Command Text"
        .AddItem "Register Text"
        .AddItem "Device Text"
        .AddItem "Number Text"
        .AddItem "Comment Text"
        .AddItem "Error Text"
        .AddItem "Literal Text"
        .AddItem "Editor Background"
        .AddItem "Console Text Colour"
        .AddItem "Console Background Colour"
        .AddItem "Program Text Colour"
        .AddItem "Program Background Colour"
    End With
    
    'Set temporary colours
    tmpPunctuationCol = glbPunctuationCol
    tmpLabelCol = glbLabelCol
    tmpVariableCol = glbVariableCol
    tmpCommandCol = glbCommandCol
    tmpRegisterCol = glbRegisterCol
    tmpDeviceCol = glbDeviceCol
    tmpNumberCol = glbNumberCol
    tmpCommentCol = glbCommentCol
    tmpErrorCol = glbErrorCol
    tmpLiteralCol = glbLiteralCol
    tmpEditorBackCol = glbEditorBackColour
    
    tmpConsoleTextColour = glbConsoleTextColour
    tmpConsoleBackColour = glbConsoleBackColour
    tmpProgramTextColour = glbProgramTextColour
    tmpProgramBackColour = glbProgramBackColour
    
    'Set temp font information
    cmbFonts.Text = glbFont
    cmbFontSize.Text = glbFontSize
    

    'Set font
    cmbFonts.Text = glbFont
    
    'Set font size
    cmbFontSize.Text = glbFontSize
    
    'Set colour syntaxing on
    If glbColourSyntax = True Then
        chkColourSyntaxing.Value = 1
    Else
        chkColourSyntaxing.Value = 0
    End If
    
    'Set syntax
    If glbAutoCheckSyntax = True Then
        chkAutoCheckSyntax.Value = 1
    Else
        chkAutoCheckSyntax.Value = 0
    End If
    
    'Set clear screen
    If glbClearScreen = True Then
        chkClearScreen.Value = 1
    Else
        chkClearScreen.Value = 0
    End If
    
    'Set screen szie
    optScreen(glbScreenSize).Value = True

    
End Sub



Private Sub lstTextColours_Click()
    'Set colour of sample text
    'Set background colour
    
    Select Case lstTextColours.ListIndex
    Case 0
        'Punctuation colour
        lblSample.ForeColor = tmpPunctuationCol
        picSample.BackColor = tmpEditorBackCol
    Case 1
        'Label colour
        lblSample.ForeColor = tmpLabelCol
        picSample.BackColor = tmpEditorBackCol
    Case 2
         'Variablecolour
        lblSample.ForeColor = tmpVariableCol
        picSample.BackColor = tmpEditorBackCol
    Case 3
        'Command colour
        lblSample.ForeColor = tmpCommandCol
        picSample.BackColor = tmpEditorBackCol
    Case 4
        'Register colour
        lblSample.ForeColor = tmpRegisterCol
        picSample.BackColor = tmpEditorBackCol
    Case 5
        'Device colour
        lblSample.ForeColor = tmpDeviceCol
        picSample.BackColor = tmpEditorBackCol
    Case 6
        'Number colour
        lblSample.ForeColor = tmpNumberCol
        picSample.BackColor = tmpEditorBackCol
    Case 7
        'Comment colour
        lblSample.ForeColor = tmpCommentCol
        picSample.BackColor = tmpEditorBackCol
    Case 8
        'Error colour
        lblSample.ForeColor = tmpErrorCol
        picSample.BackColor = tmpEditorBackCol
    Case 9
        'Literal colour
        lblSample.ForeColor = tmpLiteralCol
        picSample.BackColor = tmpEditorBackCol
    Case 10
        'Editor background colour
        picSample.BackColor = tmpEditorBackCol
    Case 11, 12
        'Console text colour
        lblSample.ForeColor = tmpConsoleTextColour
        picSample.BackColor = tmpConsoleBackColour
    Case 13, 14
       lblSample.ForeColor = tmpProgramTextColour
        picSample.BackColor = tmpProgramBackColour
    End Select
End Sub



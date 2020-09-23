VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Super Easy Assembly Language"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9480
   Icon            =   "frmMaster.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrActivate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5220
      Top             =   1680
   End
   Begin VB.PictureBox picWindowSwapper 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   9450
      TabIndex        =   1
      Top             =   405
      Width           =   9480
      Begin VB.Label lblComputerArchitecture 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Achitecture"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   5340
         MouseIcon       =   "frmMaster.frx":08D2
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   60
         Width           =   1575
      End
      Begin VB.Label lblCodeInMemory 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code In Memory"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   4020
         MouseIcon       =   "frmMaster.frx":0BDC
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label lblLocationTable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location Table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2760
         MouseIcon       =   "frmMaster.frx":0EE6
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   60
         Width           =   1050
      End
      Begin VB.Label lblProgramOutput 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Program Output"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   1440
         MouseIcon       =   "frmMaster.frx":11F0
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   60
         Width           =   1125
      End
      Begin VB.Label lblConsole 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Console"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   660
         MouseIcon       =   "frmMaster.frx":14FA
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   60
         Width           =   585
      End
      Begin VB.Label lblCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   120
         MouseIcon       =   "frmMaster.frx":1804
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   60
         Width           =   375
      End
   End
   Begin SEAL.epCmDlg dlgPrint 
      Left            =   4560
      Top             =   1800
      _ExtentX        =   661
      _ExtentY        =   635
   End
   Begin MSComDlg.CommonDialog cmDialog1 
      Left            =   3780
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "SEA"
      Filter          =   "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
      FilterIndex     =   557
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3120
      Top             =   1740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":1B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":1FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":249E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":2966
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":2E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":32F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":37BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":3C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":414E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":4616
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":4ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":4FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":546E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":5936
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":5DFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":62C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":678E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMaster.frx":6C56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbOptions 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   714
      ButtonWidth     =   714
      ButtonHeight    =   661
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Find"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Run"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Step"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Test"
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Auto Format"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu s9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoRecentFiles 
         Caption         =   "No Recent Files"
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "File"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "File"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "File"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "File"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "&Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu l3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExplainLine 
         Caption         =   "&Explain Line"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu l231 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditInsert 
         Caption         =   "&Insert"
         Begin VB.Menu mnuInsertIfThen 
            Caption         =   "IF... THEN..."
         End
         Begin VB.Menu mnuInsertIfThenElse 
            Caption         =   "IF... THEN... ELSE..."
         End
         Begin VB.Menu mnuInsertWhileLoop 
            Caption         =   "WHILE... LOOP"
         End
         Begin VB.Menu mnuInsertLoopWhile 
            Caption         =   "LOOP... WHILE"
         End
         Begin VB.Menu mnuInsertForNext 
            Caption         =   "FOR... NEXT"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewCode 
         Caption         =   "&Code"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewConsole 
         Caption         =   "C&onsole"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuViewProgramOutput 
         Caption         =   "&Program Output"
         Enabled         =   0   'False
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuViewLocationTable 
         Caption         =   "&Location Table"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuViewCodeInMemory 
         Caption         =   "Code In &Memory"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuViewComputerArchitecture 
         Caption         =   "Computer &Architecture"
         Shortcut        =   {F12}
      End
      Begin VB.Menu s5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewoptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuProgram 
      Caption         =   "&Program"
      Begin VB.Menu mnuProgramTest 
         Caption         =   "&Test"
         Shortcut        =   ^T
      End
      Begin VB.Menu s8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgramRun 
         Caption         =   "&Run"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuProgramStep 
         Caption         =   "&Step Through"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuProgramStop 
         Caption         =   "S&top"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuProgramRestart 
         Caption         =   "R&estart"
         Shortcut        =   {F4}
      End
      Begin VB.Menu l12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgramRunFullScreen 
         Caption         =   "Run &Full Screen"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu a2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgramCodeStatistics 
         Caption         =   "Code Statistics..."
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWndowTileHorizontally 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuWindowTileVertically 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindowsCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu l984 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuPopCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuPopSelectAll 
         Caption         =   "Select &All"
      End
      Begin VB.Menu s10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopExplainLine 
         Caption         =   "Explain Line"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub lblCode_Click()
    'Show code
    mnuViewCode_Click
End Sub

Private Sub lblCodeInMemory_Click()
    'Show code in memory
    mnuViewCodeInMemory_Click
End Sub

Private Sub lblComputerArchitecture_Click()
    'Show computer architecture
    mnuViewComputerArchitecture_Click
End Sub

Private Sub lblConsole_Click()
    'Show direct mode
    mnuViewConsole_Click
End Sub

Private Sub lblLocationTable_Click()
    'Show location table
    mnuViewLocationTable_Click
End Sub

Private Sub lblProgramOutput_Click()
    'Show program output
    mnuViewProgramOutput_Click
End Sub



Private Sub MDIForm_Load()
    'Set the mouse overs for the window links
    lblCode.MousePointer = 99
    lblCode.MouseIcon = LoadResPicture(1, vbResCursor)
    lblProgramOutput.MousePointer = 99
    lblProgramOutput.MouseIcon = LoadResPicture(1, vbResCursor)
    lblConsole.MousePointer = 99
    lblConsole.MouseIcon = LoadResPicture(1, vbResCursor)
    lblLocationTable.MousePointer = 99
    lblLocationTable.MouseIcon = LoadResPicture(1, vbResCursor)
    lblCodeInMemory.MousePointer = 99
    lblCodeInMemory.MouseIcon = LoadResPicture(1, vbResCursor)
    lblComputerArchitecture.MousePointer = 99
    lblComputerArchitecture.MouseIcon = LoadResPicture(1, vbResCursor)

    'Load all forms windows
    Load frmAbout
    Load frmAnimate
    Load frmCode
    Load frmConsole
    Load frmEditor
    Load frmFind
    Load frmFullScreen
    Load frmOptions
    Load frmPrint
    Load frmRun
    Load frmStats
    Load frmVariables

    tmrActivate.Enabled = True

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
    CloseAllForms = True
    
    'Stop program running
    frmRun.StopProgram
    
    'Save app settings
    SaveSettings

    'Save window setting
    SaveWindowSettings
    
    'Close application
    Unload frmAbout
    Unload frmAnimate
    Unload frmCode
    Unload frmConsole
    Unload frmCopy
    Unload frmEditor
    Unload frmFind
    Unload frmFullScreen
    Unload frmOptions
    Unload frmPrint
    Unload frmRun
    Unload frmStats
    Unload frmEditor
    Unload frmVariables
    
    'End
    Unload Me
    Set frmMain = Nothing
    End
End Sub


Public Sub mnuEditCopy_Click()
    'Edit copy
    If frmEditor.rtbProgram.SelText <> Empty Then
        Clipboard.SetText frmEditor.rtbProgram.SelText
    End If
End Sub

Public Sub mnuEditCut_Click()
    'Cut
    Clipboard.SetText frmEditor.rtbProgram.SelText
    ' Delete the selected text.
    frmEditor.rtbProgram.SelText = ""

    LineEditted = True
    glbLastLineValidated = False
    EditorLine = frmEditor.GetLineNumber
    With frmEditor.rtbProgram
        frmEditor.HighlightLine frmEditor.rtbProgram, .GetLineFromChar(.SelStart + .SelLength), vbBlack
    End With
End Sub

Public Sub mnuEditDelete_Click()
    'Do a delete

    'If the mouse pointer is not at the end of the notepad...
    With frmEditor.rtbProgram
        If .SelStart <> Len(.Text) Then
            ' If nothing is selected, extend the selection by one.
            If .SelLength = 0 Then
                .SelLength = 1
                ' If the mouse pointer is on a blank line, extend the selection by two.
                If Asc(.SelText) = 13 Then
                   .SelLength = 2
                End If
            End If
            
            ' Delete the selected text.
            .SelText = ""
            LineEditted = True
            glbLastLineValidated = False
            EditorLine = frmEditor.GetLineNumber
            frmEditor.HighlightLine frmEditor.rtbProgram, .GetLineFromChar(.SelStart + .SelLength), vbBlack
        End If
    End With
    
End Sub

Private Sub mnuEditExplainLine_Click()
    'Animate currently selected line
    If ValidateLastLine = False Then
        frmEditor.ExplainLine
    End If
End Sub

Public Sub mnuEditFind_Click()
    'Find
    'If its not spanned over a line set the find text
    If InStr(frmEditor.rtbProgram.SelText, vbCr) = 0 Then
        frmFind.cboFind.Text = frmEditor.rtbProgram.SelText
    End If
    glbQuickNoCheck = True
    frmFind.Show , Me
    glbQuickNoCheck = False
End Sub

Private Sub mnuEditFindNext_Click()
    'Find next
    On Error GoTo FindNextError
    Dim lngResult As Integer
    Dim lngPos As Integer
    Dim intOptions As Integer
    
    ValidateLastLine
    
    If frmFind.cboFind.Text = "" Then
        mnuEditFind_Click
    Else
        ' Set search options
        If frmFind.chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
        If frmFind.chkWholeWord.Value = 1 Then intOptions = intOptions + 2
        If frmFind.chkMatchCase.Value = 1 Then intOptions = intOptions + 4
    
        lngPos = frmEditor.rtbProgram.SelStart + frmEditor.rtbProgram.SelLength
        ' Get position of the searched word
        lngResult = frmEditor.rtbProgram.Find(frmFind.cboFind.Text, lngPos, , intOptions)
    
        If lngResult = -1 Then 'Text not found
            MsgBox "Text not found", "ElitePad - FindNext", 1, True 'Show message
            frmFind.cmdFind.Caption = "&Find" 'Set caption
            frmFind.cmdReplace.Enabled = False 'Disable Replace button
            frmFind.cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
 'todo: removed
          ' mnuSearchFindNext.Enabled = False 'Disable Find Next menu
        Else
             frmEditor.rtbProgram.SetFocus 'Set focus
        End If


    End If
FindNextError:
End Sub

Public Sub mnuEditReplace_Click()
    'Replace
    ValidateLastLine
    'Set the highlighted word
    If InStr(frmEditor.rtbProgram.SelText, vbCr) = 0 Then
        frmFind.cboFind.Text = frmEditor.rtbProgram.SelText
    End If
    With frmFind
        .cmdReplace.Top = 150 'Set cmdReplace top
        .cmdReplace.Caption = "&Replace" 'Set caption
        .lblReplace.Visible = True 'Show lblReplace
        .cboReplace.Visible = True 'Show cboReplace
        .cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        .Show , Me
    End With
End Sub

Public Sub mnuEditPaste_Click()
    'Paste
    If Clipboard.GetFormat(vbCFText) Then
        If Clipboard.GetText <> Empty Then
            frmEditor.DoPaste Clipboard.GetText
        Else
            Beep
        End If
    Else
        Beep
    End If
    'frmEditor.rtbProgram.SelText = Clipboard.GetText()End Sub
End Sub

Public Sub mnuEditSelectAll_Click()
    'Select all
    If ValidateLastLine = False Then
        frmEditor.rtbProgram.SelStart = 0
        frmEditor.rtbProgram.SelLength = Len(frmEditor.rtbProgram.Text)
    End If
End Sub



Private Sub mnuFileExit_Click()
    'Remove everything
    End
End Sub

'===============================================
'       File
'===============================================

Private Sub mnuFileOpen_Click()
    'Open file
    If CancelOperation = False Then
        FileOpenProc
    End If
End Sub

Private Sub mnuFileNew_Click()
    'File new
    If CancelOperation = False Then
    
        'Show clear form
        SetupTempRTF
        With frmEditor
            .rtbProgram.TextRTF = .rtbTemp.TextRTF
            .Caption = "Untitled.txt"
            ShowOptions
            glbEditorVisible = True
            glbDisableColour = False
            .Show
        End With
    Else
        ValidateLastLine
    End If
End Sub

Private Sub mnuFilePageSetup_Click()
   'Show page setup dialiog
    glbQuickNoCheck = True
    dlgPrint.ShowPageSetup
    glbQuickNoCheck = False
End Sub

Private Sub mnuFilePrint_Click()
    'Print
    On Error GoTo handler 'Printer access errors
    
    glbQuickNoCheck = True
    frmPrint.lblPrinter = Printer.DeviceName
    frmPrint.Show vbModal
    glbQuickNoCheck = False
    Exit Sub
handler:
    MsgBox "Error accessing printer - Please ensure it is correctly installed", vbExclamation
End Sub

Private Sub mnuFileSave_Click()
    'File save
    ValidateLastLine
    
    Dim strFilename As String

    If Left(frmEditor.Caption, 8) = "Untitled" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, GetFileName.
        
        'Don 't allow null saves
        If frmEditor.rtbProgram.Text = Empty Then
            MsgBox "There is no code to save!", vbExclamation
            Exit Sub
        Else
            strFilename = GetFileName(strFilename)
        End If
    Else
        ' The form's Caption contains the name of the open file.
        strFilename = frmEditor.Caption
    End If
    ' Call the save procedure. If Filename = Empty, then
    ' the user chose Cancel in the Save As dialog box; otherwise,
    ' save the file.
    If strFilename <> "" Then
        SaveFileAs strFilename
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    'File save as
    
    Dim strSaveFileName As String
    Dim strDefaultName As String
    
    'Validate last line, in background
    ValidateLastLine True
    
    ' Assign the form caption to the variable.
    strDefaultName = frmEditor.Caption
    If Left(strDefaultName, 8) = "Untitled" Then
        
        'Don't allow null saves
        If frmEditor.rtbProgram.Text = Empty Then
            MsgBox "There is no code to save!", vbExclamation
        Else
            ' The file hasn't been saved yet.
            ' Get the filename, and then call the save procedure, strSaveFileName.
            strSaveFileName = GetFileName("Untitled.txt")
            If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
            ' Update the list of recently opened files in the File menu control array.
            UpdateFileMenu (strSaveFileName)
        End If
    Else
        ' The form's Caption contains the name of the open file.
        
        strSaveFileName = GetFileName(strDefaultName)
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu (strSaveFileName)
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    'Help about
    glbQuickNoCheck = True
    frmAbout.Show vbModal
    glbQuickNoCheck = False
End Sub



Private Sub mnuHelpContents_Click()
    Dim lRet As Long
    
    lRet = ShellExecute(Me.hwnd, "open", App.Path & "\documentation\html\help.htm", "", App.Path, 9)
End Sub

Private Sub mnuInsertIfThen_Click()
    'Insert If Then
    Dim strToPaste As String
    
    AppendString strToPaste, ";IF... THEN..." & vbCrLf
    AppendString strToPaste, ";" & vbCrLf
    AppendString strToPaste, ";If MyVar=1 Then Action" & vbCrLf
    AppendString strToPaste, ";MyVar: Datai 0 ;Declaration of MyVar (If needed)" & vbCrLf
    AppendString strToPaste, "Start_If:" + Chr(9) + "Load Acc,MyVar" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Jnez End_If" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Jnez End_If" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + Chr(9) + ";Action goes here" & vbCrLf
    AppendString strToPaste, "End_If:" & vbCrLf

    'Paste the string man!
    frmEditor.DoPaste strToPaste
End Sub

Private Sub mnuInsertIfThenElse_Click()
    'Insert If then Else
    Dim strToPaste As String
    
    AppendString strToPaste, ";IF... THEN... ELSE..." & vbCrLf
    AppendString strToPaste, ";" & vbCrLf
    AppendString strToPaste, ";If MyVar=1 Then Action1 Else Action2" & vbCrLf
    AppendString strToPaste, ";MyVar: Datai 0 ;Declaration of MyVar (If needed)" & vbCrLf
    AppendString strToPaste, "Start_If:" + Chr(9) + "Load Acc,MyVar" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Cmpr Acc,#1" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Jnez Else_Part" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + Chr(9) + ";Action1 goes here" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Jump End_If" & vbCrLf
    AppendString strToPaste, "Else_Part:" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + Chr(9) + ";Action2 goes here" & vbCrLf
    AppendString strToPaste, "End_If:" & vbCrLf
    
    'Paste the string man!
    frmEditor.DoPaste strToPaste
End Sub

Private Sub mnuInsertWhileLoop_Click()
    'Insert While loop
    Dim strToPaste As String
    
    CodeDirty = True
    AppendString strToPaste, ";WHILE... LOOP..." & vbCrLf
    AppendString strToPaste, ";" & vbCrLf
    AppendString strToPaste, ";Do While MyVar=1" & vbCrLf
    AppendString strToPaste, ";" + Chr(9) + "Action" & vbCrLf
    AppendString strToPaste, ";Loop" & vbCrLf
    AppendString strToPaste, ";MyVar: Datai 0 ;Declaration of MyVar (If needed)" & vbCrLf
    AppendString strToPaste, "Start_Wh:" + Chr(9) + "Load Acc,MyVar" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Cmpr Acc,#1" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Jnez End_While" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + Chr(9) + ";Action goes here" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Jump Start_Wh" & vbCrLf
    AppendString strToPaste, "End_While:" & vbCrLf
    
    'Paste the string man!
    frmEditor.DoPaste strToPaste
End Sub

Private Sub mnuInsertLoopWhile_Click()
    'Insert Loop While
    Dim strToPaste As String
    
    AppendString strToPaste, ";LOOP... WHILE..." & vbCrLf
    AppendString strToPaste, ";" & vbCrLf
    AppendString strToPaste, ";Do" & vbCrLf
    AppendString strToPaste, ";" + Chr(9) + "Action" & vbCrLf
    AppendString strToPaste, ";Loop While MyVar=1" & vbCrLf
    AppendString strToPaste, ";MyVar: Datai 0 ;Declaration of MyVar (If needed)" & vbCrLf
    AppendString strToPaste, "Start_Rept:" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + ";Action goes here" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Load Acc,MyVar" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Cmpr Acc,#1" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Jeqz Start_Rept" & vbCrLf
    
    'Paste the string man!
    frmEditor.DoPaste strToPaste
End Sub

Private Sub mnuInsertForNext_Click()
    'Inser for next
    Dim strToPaste As String
    
    AppendString strToPaste, ";FOR... NEXT" & vbCrLf
    AppendString strToPaste, ";" & vbCrLf
    AppendString strToPaste, ";For X=1 to 5" & vbCrLf
    AppendString strToPaste, ";MyVar: Datai 0 ;Declaration of MyVar (If needed)" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Load Indx,#1" & vbCrLf
    AppendString strToPaste, "Start_For:" + Chr(9) + "Cmpr Indx,#6" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Jeqz End_For" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + Chr(9) + ";Action goes here" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + Chr(9) + "Inc Indx" & vbCrLf
    AppendString strToPaste, Chr(9) + Chr(9) + "Jump Start_For" & vbCrLf
    AppendString strToPaste, "End_For:" & vbCrLf
    
    'Paste the string man!
    frmEditor.DoPaste strToPaste
End Sub

Private Sub mnuProgramCodeStatistics_Click()
    'Code statistics
    If ValidateLastLine = False Then
    
        'Check for error in code
        If InterpretCode(frmEditor.rtbProgram, True) = True Then Exit Sub
        
        frmStats.ShowInstructions
        glbQuickNoCheck = True
        frmStats.Show vbModal
        glbQuickNoCheck = False
    End If
End Sub
Sub CheckTheFirstLine()
    'Checks the first line

    'Validate the first line
    If frmEditor.GetLineCount(frmEditor.rtbProgram) = 1 Then
        EditorLine = 1
        frmEditor.ParseLine
    End If
End Sub

Private Sub mnuProgramRestart_Click()
    'Restart
    frmRun.StopProgram
        
    'Check for errors
    If InterpretCode(frmEditor.rtbProgram, True) = True Then Exit Sub
    StoreScreenSize
    RunningProgram
    RunProgram
End Sub


Public Sub mnuProgramRun_Click()
    'Run program
    
    'Check for user errors
    If ValidateLastLine = False Then
    
        'Compile error check
        If InterpretCode(frmEditor.rtbProgram, True) = False Then
            'Enable/disable controls
            frmMain.mnuProgramStep.Enabled = False
            frmMain.tlbOptions.Buttons(14).Enabled = False 'step icons
            RunningProgram
            
            'Store screen size
            StoreScreenSize
            
            'Run code
            RunProgram CBool(mnuProgramRunFullScreen.Checked)
        End If
    End If

    
End Sub
Private Sub StoreScreenSize()
    'Stores the current screen size
    glbScreenWidth = Screen.Width / Screen.TwipsPerPixelX
    glbScreenHeight = Screen.Height / Screen.TwipsPerPixelY
End Sub
Private Sub mnuProgramRunFullScreen_Click()
    'Run program full screen
    mnuProgramRunFullScreen.Checked = Not (mnuProgramRunFullScreen.Checked)
End Sub

Private Sub mnuProgramStep_Click()
    'Step thru

    ValidateLastLine
    
    'Is code running already?
    If CodeState <> 2 Then
        'NO
        
        'Check first line
        CheckTheFirstLine

        'Check for errors
        If InterpretCode(frmEditor.rtbProgram, True) = True Then Exit Sub
        RunningProgram
        StepMode = True
        RunProgram
    ElseIf CodeState = 2 Then
        'YES
        DoNextInstruction = True
    End If
    
    If CodeState = 0 Then
        MsgBox "Finito", vbInformation
    End If
End Sub

Public Sub mnuProgramStop_Click()
    'Stop
    frmRun.StopProgram
    NotRunningProgram
End Sub

Sub RunProgram(Optional ByVal blnFullscreen As Boolean)
    'Run
    Dim a As Integer
        
    'Calculate label length
    LabelLength = 0
    For a = 0 To LabelCount - 1
        If Len(LabelName(a)) > LabelLength Then
            LabelLength = Len(LabelName(a))
        End If
    Next
    For a = 0 To VariableCount - 1
        If Len(VariableName(a)) > LabelLength Then
            LabelLength = Len(VariableName(a))
        End If
    Next
    For a = 0 To ArrayCount - 1
        If Len(ArrayName(a)) > LabelLength Then
            LabelLength = Len(ArrayName(a))
        End If
    Next
    
    'Set pen colour
    glbRunTextColour = glbProgramTextColour
    frmRun.rtbOutput.BackColor = glbProgramBackColour
    frmAnimate.rtbScreen.BackColor = glbProgramBackColour
    frmFullScreen.rtbOutput.BackColor = glbProgramBackColour
    
    mnuViewProgramOutput.Enabled = True
    lblProgramOutput.Enabled = True
    EnablePrint 1
    
    'Plus 3 for "*:"
    LabelLength = LabelLength + 2
    
    'Show variable information
    ShowVariableNames
    
    'Show code information
    ShowVariableValues
    
    'Reset animation text boxes
    frmAnimate.txtExplanation = "???"
    frmAnimate.txtInstruction = "???"
    
    'Start program
    frmCode.ShowCode

    'Clear screen?
    If glbClearScreen = True Then
        frmFullScreen.rtbOutput.Text = Empty
        frmRun.rtbOutput.Text = Empty
        frmAnimate.rtbScreen.Text = Empty
    End If

    'Do full screen animation
    If blnFullscreen = False Then
        'Normail animation
        frmRun.Show
        frmRun.InitProgram
    Else
        If StepMode = True Then
            frmRun.Show
            frmRun.InitProgram
        Else
            'Full screen version
            frmFullScreen.rtbOutput.TextRTF = frmRun.rtbOutput.TextRTF
 '           ShowCursor False
            If glbScreenSize = 0 Then
                ChangeRes 320, 240
            Else
                ChangeRes 640, 480
            End If
            frmFullScreen.WindowState = vbMaximized
            frmFullScreen.Show
            'AlwaysOnTop frmFullScreen, True
            frmFullScreen.InitProgram
        End If
    End If
End Sub

Private Sub mnuProgramTest_Click()
    'Test the program for errors
    If ValidateLastLine = False Then
        If InterpretCode(frmEditor.rtbProgram, True) = False Then
            MsgBox "No Errors!", 48 'todo: use app.title?
        End If
    End If
End Sub

Private Sub mnuRecentFile_Click(Index As Integer)
    'Recent files
    Dim FileToOpen As String
    
    If CancelOperation = False Then
    
        ' Call the file open procedure, passing a
        ' reference to the selected file name
        FileToOpen = mnuRecentFile(Index).Caption
        FileToOpen = Right(FileToOpen, Len(FileToOpen) - 3)
        
        OpenFile (FileToOpen)
    Else
        ValidateLastLine
    End If
    
End Sub

'===============================================
'       View
'===============================================

Private Sub mnuViewCode_Click()
    'Show user code
    frmEditor.SetFocus
End Sub

Private Sub mnuViewCodeInMemory_Click()
    'Show code in memory
    ShowAWindow frmCode
    frmCode.Show
    frmCode.SetFocus
End Sub

Private Sub mnuViewConsole_Click()
    'Show direct mode
    ShowAWindow frmConsole
    frmConsole.Show
    frmConsole.SetFocus
End Sub

Private Sub mnuViewProgramOutput_Click()
    'Show location table
    ShowAWindow frmRun
    frmRun.Show
    frmRun.SetFocus
End Sub


Private Sub mnuViewLocationTable_Click()

    'Show location table
    ShowAWindow frmVariables
    frmVariables.Show
    frmVariables.SetFocus
End Sub

Private Sub mnuViewComputerArchitecture_Click()
    'Show code in memory
    ShowAWindow frmAnimate
    frmAnimate.Show
    frmAnimate.SetFocus
End Sub

Private Sub mnuViewoptions_Click()
    'View options
    glbQuickNoCheck = True
    frmOptions.Show vbModal
    glbQuickNoCheck = False
End Sub


'===============================================
'       Window Options
'===============================================

Private Sub mnuWindowArrangeIcons_Click()
    'Window arrange icons
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowsCascade_Click()
    'Cascade windows
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowTileVertically_Click()
    'Tile windows vertically
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWndowTileHorizontally_Click()
    'Tile windows horizontally
    Me.Arrange vbTileHorizontal
End Sub


Private Sub tlbOptions_ButtonClick(ByVal Button As MSComctlLib.Button)
    'Toolbar button has been pressed
    'Execute
    
    Select Case Button.Index
    Case 1  'New
        mnuFileNew_Click
    Case 2 'Open
        mnuFileOpen_Click
    Case 3 'Save
        mnuFileSave_Click
    Case 5 'Print
        mnuFilePrint_Click
    Case 7 'Find
        mnuEditFind_Click
    Case 9 'Cut
        mnuEditCut_Click
    Case 10 'Copy
        mnuEditCopy_Click
    Case 11 'Paste
        mnuEditPaste_Click
    Case 13 'Run
        mnuProgramRun_Click
    Case 14 'Step
        mnuProgramStep_Click
    Case 15 'Stop
        mnuProgramStop_Click
    Case 17 'Test
        mnuProgramTest_Click
    End Select
End Sub

Private Sub tmrActivate_Timer()
    'Set all the fonts and backgroud colours
    frmRun.rtbOutput.BackColor = glbProgramBackColour
    frmEditor.rtbProgram.BackColor = glbEditorBackColour
    frmAnimate.rtbScreen.BackColor = glbProgramBackColour
    frmConsole.rtbOutput.BackColor = glbConsoleBackColour
    frmFullScreen.rtbOutput.BackColor = glbProgramBackColour
    SetAllFonts
    
    'Disable/hide options
    HideOptions
   
    'Update files list
    GetRecentFiles

    'Setup form positions
    SetupChildWindows
    
    'Show forms
    If glbShowConsole Then
        frmConsole.Show
        Inc NumberOfWindows
    End If
    If glbShowLocationTable Then
        frmVariables.Show
        Inc NumberOfWindows
    End If
    If glbShowCodeInMemory Then
        frmCode.Show
        Inc NumberOfWindows
    End If
    If glbShowComputerArchitecture Then
        frmAnimate.Show
        Inc NumberOfWindows
    End If
    
    'If first run display example code
    If glbFirstTimeRun Then
        OpenFile App.Path & "\examples\snail.txt"
        mnuWindowTileVertically_Click 'Sort out windows too
    End If
    
    
    tmrActivate.Enabled = False
End Sub

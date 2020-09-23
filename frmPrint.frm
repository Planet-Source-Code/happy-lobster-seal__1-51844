VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4275
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4275
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbTemp 
      Height          =   795
      Left            =   1860
      TabIndex        =   5
      Top             =   900
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1402
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmPrint.frx":08D2
   End
   Begin VB.ComboBox cmbPrint 
      Height          =   330
      Index           =   4
      ItemData        =   "frmPrint.frx":0949
      Left            =   1620
      List            =   "frmPrint.frx":0953
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1620
      Width           =   2535
   End
   Begin VB.ComboBox cmbPrint 
      Height          =   330
      Index           =   3
      ItemData        =   "frmPrint.frx":097B
      Left            =   1620
      List            =   "frmPrint.frx":0985
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ComboBox cmbPrint 
      Height          =   330
      Index           =   2
      ItemData        =   "frmPrint.frx":09AD
      Left            =   1620
      List            =   "frmPrint.frx":09BA
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1200
      Width           =   2535
   End
   Begin VB.ComboBox cmbPrint 
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      ItemData        =   "frmPrint.frx":09F7
      Left            =   1620
      List            =   "frmPrint.frx":0A04
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   780
      Width           =   2535
   End
   Begin VB.ComboBox cmbPrint 
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      ItemData        =   "frmPrint.frx":0A41
      Left            =   1620
      List            =   "frmPrint.frx":0A4E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   2535
   End
   Begin SEAL.epCmDlg dlgPrint 
      Left            =   1980
      Top             =   0
      _ExtentX        =   661
      _ExtentY        =   635
   End
   Begin VB.CommandButton cmdSetup 
      Caption         =   "&Setup..."
      Height          =   315
      Left            =   3060
      TabIndex        =   4
      Top             =   2460
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   2460
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   60
      X2              =   4260
      Y1              =   300
      Y2              =   300
   End
   Begin VB.Label lblWindow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code In Memory"
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   2100
      Width           =   1155
   End
   Begin VB.Label lblWindow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location Table"
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1050
   End
   Begin VB.Label lblWindow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Console"
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1260
      Width           =   585
   End
   Begin VB.Label lblWindow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program Output"
      Enabled         =   0   'False
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label lblWindow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code Window"
      Enabled         =   0   'False
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   420
      Width           =   1020
   End
   Begin VB.Label lblPrinter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "None"
      Height          =   210
      Left            =   720
      TabIndex        =   1
      Top             =   60
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Printer:   "
      Height          =   210
      Left            =   105
      TabIndex        =   0
      Top             =   60
      Width           =   645
   End
End
Attribute VB_Name = "frmPrint"
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

'Longest variable found

Private Sub cmdCancel_Click()
    'Close
    Me.Hide
End Sub

Private Sub SetupTempBox(ByVal strHeading As String, ByVal intEntireDoc As Integer, _
    Optional ByVal rtbCopyTextBox As RichTextBox)
    'Setup temp text box with RTF
     With rtbTemp
        'Setup the text box
        .Text = Empty
        .SelStart = 0
        .SelLength = 2
        .SelFontName = glbFont
        .SelFontSize = glbFontSize
        .SelColor = vbBlack
        
        'Copy document to temp document
        If intEntireDoc > 0 Then
            If intEntireDoc = 2 Then
                .SelRTF = rtbCopyTextBox.TextRTF
            Else
                .SelRTF = rtbCopyTextBox.SelRTF
            End If
        End If
        
        'Insert heading
        .SelStart = 0
        .SelLength = 0
        .SelText = strHeading & vbCrLf & vbCrLf
    End With
End Sub

Private Sub cmdOK_Click()
    'Print everything!
    
    MousePointer = vbHourglass
    
    'Code
    If cmbPrint(0).Enabled = True Then
        If cmbPrint(0).ListIndex > 0 Then
            SetupTempBox "Code: " & frmEditor.Caption, cmbPrint(0).ListIndex, frmEditor.rtbProgram
            PrintRTF rtbTemp, 720, 720, 720, 720
        End If
    End If
    
    'Program output
    If cmbPrint(1).Enabled = True Then
        If cmbPrint(1).ListIndex > 0 Then
            SetupTempBox "Program Output of: " & frmEditor.Caption, cmbPrint(1).ListIndex, frmRun.rtbOutput
            PrintRTF rtbTemp, 720, 720, 720, 720
        End If
    End If
    
    'Console
    If cmbPrint(2).ListIndex > 0 Then
        SetupTempBox "Console Output", cmbPrint(2).ListIndex, frmConsole.rtbOutput
        PrintRTF rtbTemp, 720, 720, 720, 720
    End If
   
    'Location table
    If cmbPrint(3).ListIndex = 1 = True Then
        SetupTempBox "Output of Registers and Variables", 0
        PrintVariables
    End If
    
    'Code in memory
    If cmbPrint(4).ListIndex = 1 = True Then
        SetupTempBox "Output of Code In Memory", 0
        PrintCodeInMemory
    End If
    
    'Reset mouse
    MousePointer = 0
    
    'bye
    Me.Hide
    
End Sub

Private Sub cmdSetup_Click()
    'Show printer dialogue box
    dlgPrint.ShowPrinter
    lblPrinter = Printer.DeviceName
End Sub


Private Sub Form_Load()
    'Setup the combo boxes
    Dim intBox As Integer
    
    For intBox = 0 To 4
        cmbPrint(intBox).ListIndex = 0
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Hide form
    If CloseAllForms = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
        Set frmPrint = Nothing
    End If
End Sub


Private Sub PrintVariables()
    'Set the text
    rtbTemp.SelText = GetVariablesList()

    'Print the damn thing
    PrintRTF rtbTemp, 720, 720, 720, 720

End Sub

Private Sub PrintCodeInMemory()
    'Print code in memory
    Dim intLine As Integer
        
    'Populate
    For intLine = 0 To frmCode.lstCode.ListCount - 1
        If intLine < frmCode.lstCode.ListCount - 1 Then
            fPrint frmCode.lstCode.List(intLine) & vbCrLf
        Else
            'Don't add the last CR
            fPrint frmCode.lstCode.List(intLine)
        End If
    Next
    
    'Print it man!
    PrintRTF rtbTemp, 720, 720, 720, 720
End Sub

Private Sub fPrint(str As String)
    'Prints in the rich text box
    rtbTemp.Text = rtbTemp.Text + str
End Sub




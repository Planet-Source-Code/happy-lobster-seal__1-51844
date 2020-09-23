VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBar 
      BorderStyle     =   0  'None
      Height          =   1290
      Left            =   75
      ScaleHeight     =   1290
      ScaleWidth      =   5340
      TabIndex        =   6
      Top             =   900
      Width           =   5340
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Replace &All"
         Height          =   315
         Left            =   4275
         TabIndex        =   13
         Top             =   525
         Width           =   990
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "&Help"
         Height          =   315
         Left            =   4275
         TabIndex        =   12
         Top             =   900
         Width           =   990
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "&Replace..."
         Height          =   315
         Left            =   4275
         TabIndex        =   11
         Top             =   150
         Width           =   990
      End
      Begin VB.Frame Frame1 
         Caption         =   "Search Options"
         Height          =   1215
         Left            =   60
         TabIndex        =   7
         Top             =   0
         Width           =   4065
         Begin VB.CheckBox chkWholeWord 
            Caption         =   "Find Whole Word &Only"
            Height          =   240
            Left            =   150
            TabIndex        =   10
            Top             =   300
            Width           =   1965
         End
         Begin VB.CheckBox chkMatchCase 
            Caption         =   "Match Ca&se"
            Height          =   240
            Left            =   150
            TabIndex        =   9
            Top             =   600
            Width           =   1965
         End
         Begin VB.CheckBox chkNoHighlight 
            Caption         =   "No &Highlight"
            Height          =   240
            Left            =   180
            TabIndex        =   8
            Top             =   900
            Width           =   1965
         End
      End
   End
   Begin VB.ComboBox cboReplace 
      Height          =   330
      Left            =   1200
      TabIndex        =   5
      Top             =   450
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4350
      TabIndex        =   3
      Top             =   450
      Width           =   990
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   315
      Left            =   4350
      TabIndex        =   1
      Top             =   75
      Width           =   990
   End
   Begin VB.ComboBox cboFind 
      Height          =   330
      Left            =   1200
      TabIndex        =   0
      Top             =   75
      Width           =   3015
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace &With:"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   525
      Width           =   1065
   End
   Begin VB.Label lblFind 
      Caption         =   "Fin&d What:"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   840
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
    
    '***************************************************************'
    '                         ELITEPAD                              '
    '                        Written by                             '
    '                       Andrea Batina                           '
    '                                                               '
    '  You are free to use the source code in your private,         '
    '  non-commercial, projects without permission. If you want     '
    '  to use this code in commercial projects EXPLICIT permission  '
    '  from the author is required.                                 '
    '                                                               '
    '                                                               '
    '               Copyright © Andrea Batina 1999-2000             '
    '***************************************************************'

Option Explicit

Private Sub cmdFind_Click()
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    ValidateLastLine True

    If cmdFind.Caption = "&Find" Then 'If first time
        'Add to box
        AddToCombo cboFind
    
        ' Get position of the searched word
        lngResult = frmEditor.rtbProgram.Find(cboFind.Text, 0, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgBox "Text not found", 48 'Show message
            cmdFind.Caption = "&Find" 'Set caption
'todo            frmMDI.mnuSearchFindNext.Enabled = False 'Disable Find Next menu
        Else 'Text found
            frmEditor.rtbProgram.SetFocus 'Set focus to text box
            cmdReplace.Enabled = True 'Enable Replace button
            cmdReplaceAll.Enabled = True 'Enable ReplaceAll button
            cmdFind.Caption = "&Find Next" 'Set caption
'todo            frmMDI.mnuSearchFindNext.Enabled = True 'Enable Find Next menu
        End If
    Else 'Find Next
    
        'Add to box
        AddToCombo cboFind
    
        lngPos = frmEditor.rtbProgram.SelStart + frmEditor.rtbProgram.SelLength
        lngResult = frmEditor.rtbProgram.Find(cboFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgBox "Text not found", 48 'Show message
            cmdFind.Caption = "&Find" 'Set caption
            cmdReplace.Enabled = False 'Disable Replace button
            cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
'todo            frmMDI.mnuSearchFindNext.Enabled = False 'Disable Find Next menu
        Else 'Text found
            frmEditor.rtbProgram.SetFocus 'Set focus to text box
'todo            frmMDI.mnuSearchFindNext.Enabled = True 'Enable Find Next menu
        End If
    End If
End Sub


Private Sub AddToCombo(ByRef cmbBox As ComboBox)
    'Looks for text in a combo box
    Dim intE As Integer
    Dim blnFound As Boolean
    Dim strText As String
    
    'Look for word already in list
    With cmbBox
        
        strText = .Text
        
        'Empty box so add
        If .ListCount = 0 Then
            .AddItem strText
            .Text = strText
        End If
    
        For intE = 0 To .ListCount - 1
            If .List(intE) = .Text Then
                blnFound = True
            End If
        Next
        
        'Not in list so add to it
        If blnFound = False Then
            .AddItem " "
            'Move items down the list
            For intE = .ListCount - 1 To 1 Step -1
                .List(intE) = .List(intE - 1)
            Next
            .List(0) = strText
        End If
        
        .Text = strText
    End With
End Sub

Private Sub cmdReplace_Click()
    Dim lngResult As Long
    Dim lngPos As Long
    Dim intOptions As Integer
    
    'Add to box
    AddToCombo cboFind
    AddToCombo cboReplace
    
    If cmdReplace.Caption = "&Replace..." Then 'Show replace
        cmdReplace.Top = 150 'Set cmdReplace top
        cmdReplace.Caption = "&Replace" 'Set caption
        lblReplace.Visible = True 'Show lblReplace
        cboReplace.Visible = True 'Show cboReplace
        cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        Exit Sub
    End If

    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4
    
    With frmEditor.rtbProgram
        .SelText = cboReplace.Text 'Replace text
        ' Find next
        lngPos = .SelStart + .SelLength
        ' Get position of the searched word
        lngResult = .Find(cboFind.Text, lngPos, , intOptions)

        If lngResult = -1 Then 'Text not found
            MsgBox "Text not found", 48 'Show message
            cmdFind.Caption = "&Find" 'Set caption
            cmdReplace.Enabled = False 'Disable Replace button
            cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
        Else 'Text found
            .SetFocus 'Set focus
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdReplaceAll_Click()
    Dim intCount As Integer
    Dim lngPos As Long
    Dim intOptions As Integer
    ' Set search options
    If chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If chkMatchCase.Value = 1 Then intOptions = intOptions + 4
    
    ValidateLastLine False
    
    'Add to box
    AddToCombo cboFind
    AddToCombo cboReplace
    
    intCount = 0
    lngPos = 0
    With frmEditor.rtbProgram
        Do
            If .Find(cboFind.Text, lngPos, , intOptions) = -1 Then 'Text not fount
                If intCount > 0 Then 'Show how many replacments have been made
                    MsgBox "The specified region has been searched. " & vbCrLf & _
                    intCount & " replacements have been made", 48
                End If
                cmdFind.Caption = "&Find" 'Set caption
                cmdReplace.Enabled = False 'Disable Replace button
                cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
                Exit Do
            Else 'Text found
                lngPos = .SelStart + .SelLength
                intCount = intCount + 1 'Increase counter by 1
                .SelText = cboReplace.Text 'Replace text
            End If
        Loop
    End With
End Sub

Private Sub Form_Load()
    'Check user hasn't selected many lines
    With frmEditor.rtbProgram
        If .SelLength > 0 Then
            If InStr(.SelText, vbCr) + InStr(.SelText, vbLf) = 0 Then
                cboFind.AddItem frmEditor.rtbProgram.SelText 'Add selected text to find combobox
                cboFind.Text = frmEditor.rtbProgram.SelText  'Set text in cbo
            End If
        End If
    End With
    
    cmdReplace.Top = 525 'Set cmdReplace top
    lblReplace.Visible = False 'Hide lblReplace
    cboReplace.Visible = False 'Hide cboReplace
    cmdReplaceAll.Visible = False 'Hide cmdReplaceAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Hide form
    If CloseAllForms = False Then
        Cancel = True
        Me.Hide
    Else
        Unload Me
        Set frmFind = Nothing
    End If
End Sub

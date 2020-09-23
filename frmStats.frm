VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Analysis"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   2295
      ScaleWidth      =   5415
      TabIndex        =   2
      Top             =   600
      Width           =   5415
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   5400
         X2              =   5400
         Y1              =   2280
         Y2              =   120
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   26
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   25
         Top             =   1920
         Width           =   495
      End
      Begin VB.Shape shpExternal 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         Top             =   1200
         Width           =   135
      End
      Begin VB.Shape shpControl 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         Top             =   1560
         Width           =   135
      End
      Begin VB.Shape shpInternal 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         Top             =   840
         Width           =   135
      End
      Begin VB.Shape shpArithmetic 
         BackColor       =   &H008080FF&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         Top             =   480
         Width           =   135
      End
      Begin VB.Line lneHoriz 
         X1              =   3000
         X2              =   5040
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line lneVert 
         X1              =   3000
         X2              =   3000
         Y1              =   1920
         Y2              =   360
      End
      Begin VB.Label lblControl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2520
         TabIndex        =   24
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label lblExternal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2520
         TabIndex        =   23
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblInternal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2520
         TabIndex        =   22
         Top             =   840
         Width           =   240
      End
      Begin VB.Label lblArithmetic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2520
         TabIndex        =   21
         Top             =   480
         Width           =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000005&
         X1              =   2280
         X2              =   5400
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000003&
         X1              =   2280
         X2              =   2280
         Y1              =   120
         Y2              =   2280
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000003&
         X1              =   2280
         X2              =   5400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line8 
         BorderColor     =   &H8000000E&
         X1              =   0
         X2              =   2040
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000011&
         BorderWidth     =   2
         X1              =   0
         X2              =   2040
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Instruction Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Control:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   16
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "External Data Transfer (I/O)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   15
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Internal Data Transfer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arithmetic:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   765
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   0
      Left            =   300
      ScaleHeight     =   2415
      ScaleWidth      =   5535
      TabIndex        =   1
      Top             =   600
      Width           =   5535
      Begin VB.Label lblControlPerc 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   31
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblExternalPerc 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblInternalPerc 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   29
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblArithmeticPerc 
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   28
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   0
         Width           =   1335
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000E&
         X1              =   0
         X2              =   4440
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000011&
         BorderWidth     =   2
         X1              =   0
         X2              =   4440
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lblTotal 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   19
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "Label14"
         Height          =   15
         Left            =   0
         TabIndex        =   18
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Instruction Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Control:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "External Data Transfer (I/O)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   10
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Internal Data Transfer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Arithmetic:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   0
         TabIndex        =   8
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label lblArithmeticNo 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblInternalNo 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblExternalNo 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblControlNo 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   1560
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   0
      Top             =   3180
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   3015
      Left            =   120
      TabIndex        =   32
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Statistics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Graph"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmStats"
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

Private Sub cmdOK_Click()
    Unload Me
End Sub

Public Sub ShowInstructions()
    'Show instructions statistics
    
    Dim ArithmeticCount As Double
    Dim ExternalCount As Double
    Dim InternalCount As Double
    Dim ControlCount As Double
    Dim InstructionNo As Integer
    Dim ScaleVal As Double
    Dim MaxVal As Double
    
    'Count the different types of commands
    For InstructionNo = 0 To LineCount - 1
        
        Select Case Operation(InstructionNo)
        Case 0 To 9
            'Arithmetic
            Inc ArithmeticCount
        Case 10, 11
            'Internal data transfer
            Inc InternalCount
        Case 12 To 20
            'Jumps, halts, subs
            Inc ControlCount
        Case 21 To 23
            'External data transfer
            Inc ExternalCount
        End Select
    Next
    
    'Display numbers of instructions
    lblArithmeticNo = Format(ArithmeticCount)
    lblControlNo = Format(ControlCount)
    lblInternalNo = Format(InternalCount)
    lblExternalNo = Format(ExternalCount)
    
    'Display total number of instructions
    lblTotal = Format(LineCount)
    
    'Calculate the percentage values
    ArithmeticCount = ArithmeticCount / LineCount
    InternalCount = InternalCount / LineCount
    ControlCount = ControlCount / LineCount
    ExternalCount = ExternalCount / LineCount
    
    MaxVal = ArithmeticCount
    If InternalCount > MaxVal Then
        MaxVal = InternalCount
    End If
    If ControlCount > MaxVal Then
        MaxVal = ControlCount
    End If
    If ExternalCount > MaxVal Then
        MaxVal = ExternalCount
    End If
    
    ScaleVal = Abs(lneHoriz.X1 - lneHoriz.X2)
    
    'Draw the little graph
    shpArithmetic.Width = (ArithmeticCount / MaxVal) * ScaleVal
    shpInternal.Width = (InternalCount / MaxVal) * ScaleVal
    shpExternal.Width = (ExternalCount / MaxVal) * ScaleVal
    shpControl.Width = (ControlCount / MaxVal) * ScaleVal
    
    'Labels on stats tab
    lblArithmeticPerc = Format(ArithmeticCount * 100, "0") & "%"
    lblControlPerc = Format(ControlCount * 100, "0") & "%"
    lblInternalPerc = Format(InternalCount * 100, "0") & "%"
    lblExternalPerc = Format(ExternalCount * 100, "0") & "%"
    
    'Labels on graph tab
    lblArithmetic = lblArithmeticPerc
    lblControl = lblControlPerc
    lblInternal = lblInternalPerc
    lblExternal = lblExternalPerc
    
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = tbsOptions.SelectedItem.Index
        If i = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(i + 1)
        End If
    End If
End Sub

Private Sub tbsOptions_Click()
    
    Dim i As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsOptions.Tabs.Count - 1
        If i = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(i).Left = 360
            picOptions(i).Enabled = True
        Else
            picOptions(i).Left = -20000
            picOptions(i).Enabled = False
        End If
    Next
    
End Sub


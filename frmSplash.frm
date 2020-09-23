VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3750
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6015
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   401
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4860
      Top             =   2820
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "djhappylobster@hotmail.com"
      Height          =   210
      Left            =   3420
      TabIndex        =   5
      Top             =   2340
      Width           =   2070
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5400
      Picture         =   "frmSplash.frx":4942E
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by Marcus Mason"
      Height          =   270
      Left            =   3420
      TabIndex        =   4
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 0.5.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3420
      TabIndex        =   3
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2002 Supercrab.com.  All Rights Reserved."
      ForeColor       =   &H0080FF80&
      Height          =   210
      Left            =   2760
      TabIndex        =   2
      Top             =   3480
      Width           =   3120
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Super Easy Assembly Language"
      Height          =   210
      Left            =   3420
      TabIndex        =   1
      Top             =   1020
      Width           =   2370
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEAL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      Left            =   3360
      TabIndex        =   0
      Top             =   180
      Width           =   2160
   End
   Begin VB.Line Line1 
      X1              =   400
      X2              =   0
      Y1              =   249
      Y2              =   249
   End
   Begin VB.Shape Shape1 
      Height          =   3735
      Left            =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmSplash"
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
Private Sub Form_Load()
    'Set version
    lblProductName = App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    tmrShow.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Remove form
    Set frmSplash = Nothing
End Sub

Private Sub tmrShow_Timer()
    Dim OldNow As Single
    'Do the splash
        
    'Wait a bit
    OldNow = Timer
    Do
        DoEvents
    Loop Until Timer - OldNow > 1
    
    'Setup the window
    LoadAndSetWindowSettings
    frmMain.Show

    'Wait another bit
    Do
    Loop Until Timer - OldNow > 3
    
    
    'Close window
    Unload frmSplash
End Sub

VERSION 5.00
Begin VB.Form frmEgg 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9915
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   661
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Label lblMonster 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Beer Monster!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1125
      Left            =   -1.50000e5
      TabIndex        =   0
      Top             =   2760
      Width           =   6465
   End
   Begin VB.Label lblMonster2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Beer Monster!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1125
      Left            =   -1.50000e5
      TabIndex        =   1
      Top             =   2760
      Width           =   6465
   End
   Begin VB.Image imgEgg 
      Appearance      =   0  'Flat
      Height          =   5190
      Left            =   0
      Picture         =   "frmEgg.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7050
   End
End
Attribute VB_Name = "frmEgg"
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

Private Sub Form_Activate()
    'Set the img properties
    
    imgEgg.Width = Me.ScaleWidth
    imgEgg.Height = Me.ScaleHeight

    With lblMonster
        .Left = (Me.ScaleWidth - .Width) / 2
        .Top = Me.ScaleHeight - .Height
    End With
    With lblMonster2
        .Left = ((Me.ScaleWidth - .Width) / 2) - 4
        .Top = Me.ScaleHeight - .Height
    End With

End Sub

Private Sub imgEgg_Click()
    'Hide me
    Me.Hide
End Sub

Private Sub Timer1_Timer()

End Sub

VERSION 5.00
Begin VB.Form frmCopy 
   Caption         =   "Form2"
   ClientHeight    =   3180
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuCopy 
      Caption         =   "Copy"
      Begin VB.Menu mnuDoCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmCopy"
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

Private Sub mnuDoCopy_Click()
    'Copy the selected form
    Select Case glbFormToCopy
    Case intAnimate
        'Copy animate screen
        CopyAnimateScreen
    Case intKeyboard
        'Copy keyboard screen
        CopyKeyboard
    Case intConsole
        'Copy console screen
        CopyConsoleScreen
    Case intRun
        'Copy run screen
        CopyRunScreen
    Case intVariables
        'Copy variables screen
        CopyVariables
    Case intCodeInMemory
        'Copy code in memory
        CopyCodeInMemory
    End Select
    
    
End Sub


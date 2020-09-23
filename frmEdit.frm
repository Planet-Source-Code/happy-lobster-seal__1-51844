VERSION 5.00
Begin VB.Form frmEdit 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4680
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
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
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu a1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelecAtll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu a2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExplainLine 
         Caption         =   "&Explain Line..."
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuEditAnimateLine_Click()
    'Animate current cursor lines
    frmTest.AnimateLine
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Close form
    Unload Me
    Set frmEdit = Nothing
End Sub

Private Sub mnuEditCut_Click()
    'Do cut
    frmMain.mnuEditCut_Click
End Sub

Private Sub mnuEditDelete_Click()
    'Do delete
    frmMain.mnuEditDelete_Click
End Sub

Private Sub mnuEditExplainLine_Click()
    'Explain current cursor line
    frmTest.ExplainLine
End Sub

Private Sub mnuEditPaste_Click()
    'Do paste
    frmMain.mnuEditPaste_Click
End Sub

Private Sub mnuEditSelecAtll_Click()
    'Do select all
    frmMain.mnuEditSelectAll
End Sub

VERSION 5.00
Begin VB.Form WorkPage 
   Appearance      =   0  'Flat
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Valentine"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   3855
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Master 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   210
      Picture         =   "WorkPage.frx":0000
      ScaleHeight     =   155
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   226
      TabIndex        =   0
      Top             =   135
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.PictureBox WorkScr 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1080
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2970
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.PictureBox CleanScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   285
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2985
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "WorkPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



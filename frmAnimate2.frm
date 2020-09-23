VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAnimate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "(F12) Computer Architecture"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAnimate2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   494
   Begin VB.ComboBox cmbAnimate 
      Height          =   330
      ItemData        =   "frmAnimate2.frx":08D2
      Left            =   5400
      List            =   "frmAnimate2.frx":08DF
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox txtIndx 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1860
      Width           =   675
   End
   Begin VB.TextBox txtFlag 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1140
      Width           =   255
   End
   Begin VB.TextBox txtAcc 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   1500
      Width           =   675
   End
   Begin VB.TextBox txtExplanation 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   2340
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1500
      Width           =   2175
   End
   Begin VB.TextBox txtInstruction 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1140
      Width           =   2175
   End
   Begin VB.TextBox txtKeyboard 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   5580
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rtbScreen 
      Height          =   1515
      Left            =   5340
      TabIndex        =   9
      Top             =   540
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2672
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      RightMargin     =   32767
      OLEDropMode     =   0
      TextRTF         =   $"frmAnimate2.frx":0900
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstVariables 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   420
      TabIndex        =   7
      Top             =   4800
      Width           =   4215
   End
   Begin VB.ListBox lstCode 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   870
      Left            =   420
      TabIndex        =   3
      Top             =   3780
      Width           =   4215
   End
   Begin MSComctlLib.Slider sldSpeed 
      Height          =   375
      Left            =   5280
      TabIndex        =   24
      Top             =   5400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      MousePointer    =   9
      LargeChange     =   1
      Min             =   1
      Max             =   101
      SelStart        =   1
      TickStyle       =   3
      Value           =   1
   End
   Begin VB.Label lblControl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "oupti scr,myvariablethats"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   5
      Top             =   -1500
      Width           =   1560
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Some Data Is Here"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6840
      TabIndex        =   6
      Top             =   -1500
      Width           =   1155
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Animation:"
      Height          =   195
      Left            =   5400
      TabIndex        =   27
      Top             =   4440
      Width           =   1155
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Animation Speed:"
      Height          =   255
      Left            =   5400
      TabIndex        =   26
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fast                     Slow"
      Height          =   210
      Left            =   5400
      TabIndex        =   25
      Top             =   5760
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Arithmetic and Logic Unit"
      Height          =   435
      Left            =   420
      TabIndex        =   22
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Indx:"
      Height          =   195
      Left            =   540
      TabIndex        =   18
      Top             =   1860
      Width           =   495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Flag:"
      Height          =   195
      Left            =   540
      TabIndex        =   17
      Top             =   1140
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Acc:"
      Height          =   195
      Left            =   540
      TabIndex        =   16
      Top             =   1500
      Width           =   435
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   2
      X1              =   156
      X2              =   156
      Y1              =   184
      Y2              =   236
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   7
      X1              =   404
      X2              =   404
      Y1              =   184
      Y2              =   232
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   5
      X1              =   224
      X2              =   404
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   1
      X1              =   64
      X2              =   156
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   0
      X1              =   84
      X2              =   84
      Y1              =   148
      Y2              =   204
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   0
      X1              =   64
      X2              =   64
      Y1              =   148
      Y2              =   184
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   4
      X1              =   244
      X2              =   244
      Y1              =   148
      Y2              =   204
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   4
      X1              =   224
      X2              =   224
      Y1              =   148
      Y2              =   184
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Instruction:"
      Height          =   195
      Left            =   2340
      TabIndex        =   15
      Top             =   960
      Width           =   1455
   End
   Begin VB.Shape shpControlUnit 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   2220
      Top             =   900
      Width           =   2415
   End
   Begin VB.Label lblControlUnit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Control Unit"
      Height          =   255
      Left            =   2220
      TabIndex        =   8
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5580
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Shape shpKeyboard 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   675
      Left            =   5400
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblScreen 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Screen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5340
      TabIndex        =   2
      Top             =   300
      Width           =   1815
   End
   Begin VB.Shape shpScreen 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   5160
      Top             =   300
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Central Processing Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   420
      TabIndex        =   10
      Top             =   300
      Width           =   4155
   End
   Begin VB.Shape shpALU 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   420
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   2175
      Left            =   240
      Top             =   240
      Width           =   4575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Variables:"
      Height          =   195
      Left            =   3720
      TabIndex        =   14
      Top             =   4620
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Code:"
      Height          =   195
      Left            =   3720
      TabIndex        =   13
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lblMemory 
      BackStyle       =   0  'Transparent
      Caption         =   "Memory"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   420
      TabIndex        =   0
      Top             =   3540
      Width           =   3615
   End
   Begin VB.Shape shpMemory 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   240
      Top             =   3540
      Width           =   4575
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   7
      X1              =   424
      X2              =   424
      Y1              =   204
      Y2              =   232
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   5
      X1              =   244
      X2              =   424
      Y1              =   204
      Y2              =   204
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   6
      X1              =   404
      X2              =   404
      Y1              =   148
      Y2              =   184
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   3
      X1              =   156
      X2              =   224
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   3
      X1              =   176
      X2              =   244
      Y1              =   204
      Y2              =   204
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   2
      X1              =   176
      X2              =   176
      Y1              =   204
      Y2              =   236
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   6
      X1              =   424
      X2              =   424
      Y1              =   148
      Y2              =   204
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   1
      X1              =   84
      X2              =   176
      Y1              =   204
      Y2              =   204
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   5955
      Left            =   60
      Top             =   60
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      Height          =   1635
      Left            =   5220
      Top             =   4380
      Width           =   2055
   End
End
Attribute VB_Name = "frmAnimate"
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


Private Const intLeftFromLine = 2
Private Const intRightFromLine = 2
Private Const intUpFromLine = 2
Private Const intDownFromLine = 2
Private Const intUp = 0
Private Const intDown = 1
Private Const intLeft = 0
Private Const intRight = 1
Private Const intControlLine = 0
Private Const intDataLine = 1
Private Const intHorizontal = 0
Private Const intVertical = 1
Private Const lngLineColour = &H80C0FF  'Entire path colour
Private Const lngLineSectionColour = &HFF 'Section of line colour
Private Const lngComponentColour = &HFF

Private WithEvents tmrAnimate As ccrpTimer
Attribute tmrAnimate.VB_VarHelpID = -1
Private WithEvents tmrWait As ccrpCountdown
Attribute tmrWait.VB_VarHelpID = -1

'Used for animation purposes
Private intAnimateOrientation As Integer
Private intAnimateDirection As Integer
Private intAnimateStartPos As Integer
Private intAnimateEndPos As Integer
Private intAnimateStep As Integer
Private intAnimateLine As Integer
Private blnFinishedAnimation As Boolean

Private ControlPressed As Boolean
Private ControlPressed2 As Boolean
Private blnhidekey As Boolean


Sub Animate()
    
    'Update the CU labels
    If TempOperation <> -1 Then
        Access_Memory intCU, "Get Next Instruction", ConvertOperation, False
        txtInstruction = ConvertOperation
        txtExplanation = TempReminder
    End If
    
    'Animate the command
    Select Case TempOperation
    Case -1 'End of program
        'Don't clear the text boxes
    Case 0 To 11 'ALU commands
        Animate_ALU_Command
    Case 12, 18, 19, 20 'Program commands
        'Animate nothing
        'jump, jubsr, exit, halt
    Case 13 To 17, 27 'Conditional jumps
        Animate_Get_Flag_Value
    Case 21 'Input commands
        'Animate_Input
        ColourComponent intKeyb
    Case 22, 23, 26, 28 'Output commands
        Animate_Output
    End Select

End Sub

Sub Animate_ALU_Command()
    Dim intTemp As Integer
    Dim strVal
    Select Case TempOperation
    Case 5 To 8
        'Neg,Inc,Dec,Clrz
        lblControl = TempCodeWithNoAddress
        Control_CU_To_ALU
        If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
    Case 0 To 4, 9 To 10
        'Arithmetic
        If TempOperand2 = 3 Then
            'Get immediate value
            lblControl = TempCodeWithNoAddress
            Control_CU_To_ALU
            If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
            strVal = GetVariableName
            lblData = Right(strVal, Len(strVal) - 1)
            Data_CU_To_Alu
            If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
        Else
            If TempOperand3 = -1 Then
                'Rnd function
                lblControl = TempCodeWithNoAddress
                Control_CU_To_ALU
                If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
                lblData = Format(TempRandomNumber)
                Data_CU_To_Alu
                If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
            Else
                'Get variable/array from memory
                lblControl = ConvertOperation
                Control_CU_To_ALU
                If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
                
                'Test array index
                If IsArrayIndexOutOfRange = True Then
                    Access_Memory intALU, "Get " + GetVariableName(True), ErrorMessages, True
                Else
                    Access_Memory intALU, "Get " + GetVariableName(True), Format(GetValue), False
                End If
            End If
        End If
    Case 11
        'Copy
        lblControl = ConvertOperation
        Animate_Storage intALU, Format(LastRegister)
    End Select
    
    ClearComponents
End Sub

Sub Animate_Storage(ByVal intSource As Integer, ByVal strStoreValue As String)
    
        
    'Test array index
    If IsArrayIndexOutOfRange = True Then
        Store_To_Memory intSource, "Store " + GetVariableName(True), ErrorMessages, True
    Else
        'Attempt to store value
        Store_To_Memory intSource, "Store " + GetVariableName(True), strStoreValue, False
        
        If Len(strStoreValue) > 6 Then
            'String is too big to attempt converting
            ErrorMessages = "Overflow - number is too big to fit in memory"
            AnimateError intMemory, ErrorMessages
        ElseIf IsOutOfRange(Val(strStoreValue)) = True Then
            'Attempt to convert string
            ErrorMessages = "Overflow - number is too big to fit in memory"
            AnimateError intMemory, ErrorMessages
        End If
    End If
End Sub

Sub AnimateError(intSource As Integer, strMessage As String)
    'Animate error message from a component
    If glbAnimateType = intNone Then Exit Sub
    
    lblControl = strMessage
    Select Case intSource
    Case intALU
        Control_ALU_TO_CU
    Case intMemory
        Control_Memory_To_CU
    Case intScr
        Control_Scr_To_CU
    End Select
End Sub

Sub Animate_Input()
    'Animate input
    lblControl = ConvertOperation
    Control_CU_To_Keyb
End Sub

Sub Animate_Output()
    'Animate output to screen
    If TempOperation = 28 Then
        lblControl = "Ouptc"
        Control_CU_To_Scr
        If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
    ElseIf TempOperation = 26 Then
        lblControl = "Clrs"
        Control_CU_To_Scr
        If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
    End If
            
    'Output number to screen
    If TempOperand2 = 3 Then
         'Get immediate value
        lblData = GetVariableName
        Data_CU_To_Scr
    ElseIf TempOperandText <> Empty Then
        lblData = TempOperandText
        Data_CU_To_Scr
    Else
        'Animate getting variable from memory to scr
        'Test array index
        If IsArrayIndexOutOfRange = True Then
            Access_Memory intCU, "Get " + GetVariableName(True), ErrorMessages, True
        Else
            Access_Memory intCU, "Get " + GetVariableName(True), Format(GetValue), False, intScr
        End If
        
    End If

End Sub

Sub Animate_Get_Flag_Value()
    'Get Flag Value
    lblControl = "Get Flag Value"
    Control_CU_To_ALU
    If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
    lblData = txtFlag
    Data_ALU_To_CU
End Sub


Sub ColourComponent(ByVal intComponent As Integer)
    If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
    
    'Colour a component
    Select Case intComponent
    Case intALU
        shpALU.BorderWidth = 2
        shpALU.BorderColor = lngComponentColour
    Case intCU
        shpControlUnit.BorderWidth = 2
        shpControlUnit.BorderColor = lngComponentColour
    Case intKeyb
        shpKeyboard.BorderWidth = 2
        shpKeyboard.BorderColor = lngComponentColour
    Case intScr
        shpScreen.BorderWidth = 2
        shpScreen.BorderColor = lngComponentColour
    Case intMemory
        shpMemory.BorderWidth = 2
        shpMemory.BorderColor = lngComponentColour
    End Select
End Sub

Sub ClearComponents()
    'Clear highlighted components
    shpALU.BorderWidth = 1
    shpALU.BorderColor = 0
    shpControlUnit.BorderWidth = 1
    shpControlUnit.BorderColor = 0
    shpKeyboard.BorderWidth = 1
    shpKeyboard.BorderColor = 0
    shpScreen.BorderWidth = 1
    shpScreen.BorderColor = 0
    shpMemory.BorderWidth = 1
    shpMemory.BorderColor = 0
End Sub



Sub Access_Memory(ByVal intSource As Integer, ByVal strSend As String, ByVal strGet As String, ByVal blnError As Boolean, Optional ByVal intDestination As Integer)
    lblControl = strSend
    lblData = strGet
    
    
    'Animate fetch/get memory
    If intSource = intALU Then
        Control_ALU_To_Memory
        If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
        If blnError = False Then
            'Animate data returning
            Data_Memory_To_ALU
            If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
        Else
            'Animate error
            lblControl = strGet
            Control_Memory_To_CU
            If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
        End If
    ElseIf intSource = intCU Then
        Control_CU_To_Memory
        If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
        If blnError = False Then
            'Animate data returning
            If intDestination = intScr Then
                Data_Memory_To_Scr 'for output command
            Else
                Data_Memory_To_CU
            End If
        Else
            'Animate error
            lblControl = strGet
            Control_Memory_To_CU
        End If
    End If
End Sub

Sub Store_To_Memory(ByVal intSource As Integer, ByVal strSend As String, ByVal strGet As String, ByVal blnError As Boolean)


    lblControl = strSend
    lblData = strGet
    
    'Animate fetch/get memory
    If intSource = intALU Then
        Control_ALU_To_Memory
        If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
        If blnError = False Then
            'Animate data returning
            Data_ALU_To_Memory
            If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
        Else
            'Animate error
            lblControl = strGet
            Control_Memory_To_CU
        End If
    ElseIf intSource = intKeyb Then
        Control_Keyb_To_Memory
        If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
        If blnError = False Then
            'Animate data returning
            Data_Keyb_To_Memory
        Else
            'Animate error
            lblControl = strGet
            Control_Memory_To_CU
        End If
    End If

End Sub

Sub Control_CU_To_ALU()
    ColourComponent intCU
    lneControl(4).BorderColor = lngLineColour
    lneControl(3).BorderColor = lngLineColour
    lneControl(1).BorderColor = lngLineColour
    lneControl(0).BorderColor = lngLineColour
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(3), intLeft, False, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(1), intLeft, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(0), intUp, False, True
    ColourComponent intALU
    DoSteppedWait
    ResetControlLines
    ClearComponents
End Sub

Sub Data_CU_To_Alu()
    ColourComponent intCU
    lneData(4).BorderColor = lngLineColour
    lneData(3).BorderColor = lngLineColour
    lneData(1).BorderColor = lngLineColour
    lneData(0).BorderColor = lngLineColour
    Animate_Line intDataLine, intVertical, lblData, lneData(4), intDown, True, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(3), intLeft, False, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(1), intLeft, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(0), intUp, False, True
    ColourComponent intALU
    DoSteppedWait
    ResetDataLines
    ClearComponents
End Sub

Sub Control_ALU_To_Memory()
    ColourComponent intALU
    lneControl(0).BorderColor = lngLineColour
    lneControl(1).BorderColor = lngLineColour
    lneControl(2).BorderColor = lngLineColour
    Animate_Line intControlLine, intVertical, lblControl, lneControl(0), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(1), intRight, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(2), intDown, True, True
    ColourComponent intMemory
    DoSteppedWait
    ResetControlLines
    ClearComponents
End Sub

Sub Control_CU_To_Memory()
    ColourComponent intCU
    lneControl(4).BorderColor = lngLineColour
    lneControl(3).BorderColor = lngLineColour
    lneControl(2).BorderColor = lngLineColour
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(3), intLeft, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(2), intDown, True, True
    ColourComponent intMemory
    DoSteppedWait
    ResetControlLines
    ClearComponents
End Sub

Sub Data_Memory_To_ALU()
    ColourComponent intMemory
    lneData(2).BorderColor = lngLineColour
    lneData(1).BorderColor = lngLineColour
    lneData(0).BorderColor = lngLineColour
    Animate_Line intDataLine, intVertical, lblData, lneData(2), intUp, True, True
    Animate_Line intDataLine, intHorizontal, lblData, lneData(1), intLeft, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(0), intUp, False, True
    ColourComponent intALU
    DoSteppedWait
    ResetDataLines
    ClearComponents
End Sub

Sub Data_Memory_To_CU()
    ColourComponent intMemory
    lneData(2).BorderColor = lngLineColour
    lneData(3).BorderColor = lngLineColour
    lneData(4).BorderColor = lngLineColour
    Animate_Line intDataLine, intVertical, lblData, lneData(2), intUp, True, True
    Animate_Line intDataLine, intHorizontal, lblData, lneData(3), intRight, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(4), intUp, False, True
    ColourComponent intCU
    DoSteppedWait
    ResetDataLines
    ClearComponents
End Sub

Sub Data_Memory_To_Scr()
    ColourComponent intMemory
    lneData(2).BorderColor = lngLineColour
    lneData(3).BorderColor = lngLineColour
    lneData(5).BorderColor = lngLineColour
    lneData(6).BorderColor = lngLineColour
    Animate_Line intDataLine, intVertical, lblData, lneData(2), intUp, True, True
    Animate_Line intDataLine, intHorizontal, lblData, lneData(3), intRight, False, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(5), intRight, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(6), intUp, False, True
    ColourComponent intScr
    DoSteppedWait
    ResetDataLines
    ClearComponents
End Sub

Sub Data_ALU_To_CU()
    ColourComponent intALU
    lneData(0).BorderColor = lngLineColour
    lneData(1).BorderColor = lngLineColour
    lneData(3).BorderColor = lngLineColour
    lneData(4).BorderColor = lngLineColour
    Animate_Line intDataLine, intVertical, lblData, lneData(0), intDown, True, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(1), intRight, False, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(3), intRight, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(4), intUp, False, True
    ColourComponent intCU
    DoSteppedWait
    ResetDataLines
    ClearComponents
End Sub

'These calls are used for animating error messages
Sub Control_ALU_TO_CU()
    ColourComponent intALU
    lneControl(0).BorderColor = lngLineColour
    lneControl(1).BorderColor = lngLineColour
    lneControl(3).BorderColor = lngLineColour
    lneControl(4).BorderColor = lngLineColour
    Animate_Line intControlLine, intVertical, lblControl, lneControl(0), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(1), intRight, False, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(3), intRight, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intUp, False, True
    ColourComponent intCU
    DoSteppedWait
    ResetDataLines
    ClearComponents
End Sub

Sub Control_Memory_To_CU()
    ColourComponent intMemory
    lneControl(2).BorderColor = lngLineColour
    lneControl(3).BorderColor = lngLineColour
    lneControl(4).BorderColor = lngLineColour
    Animate_Line intControlLine, intVertical, lblControl, lneControl(2), intUp, True, True
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(3), intRight, True, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intUp, False, True
    ColourComponent intCU
    DoSteppedWait
    ResetControlLines
    ClearComponents
End Sub

Sub Control_Scr_To_CU()
    ColourComponent intScr
    lneControl(6).BorderColor = lngLineColour
    lneControl(5).BorderColor = lngLineColour
    lneControl(4).BorderColor = lngLineColour
    Animate_Line intControlLine, intVertical, lblControl, lneControl(6), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(5), intLeft, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intUp, False, True
    ColourComponent intCU
    DoSteppedWait
    ResetControlLines
    ClearComponents
End Sub

'This is for input
Sub Control_CU_To_Keyb()
    ColourComponent intCU
    lneControl(4).BorderColor = lngLineColour
    lneControl(5).BorderColor = lngLineColour
    lneControl(7).BorderColor = lngLineColour
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(5), intRight, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(7), intDown, True, True
    ColourComponent intKeyb
    DoSteppedWait
    ResetControlLines
    ClearComponents
End Sub

Sub Data_Keyb_To_Memory()
    ColourComponent intKeyb
    lneData(7).BorderColor = lngLineColour
    lneData(5).BorderColor = lngLineColour
    lneData(3).BorderColor = lngLineColour
    lneData(2).BorderColor = lngLineColour
    Animate_Line intDataLine, intVertical, lblData, lneData(7), intUp, True, True
    Animate_Line intDataLine, intHorizontal, lblData, lneData(5), intLeft, False, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(3), intLeft, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(2), intDown, True, True
    ColourComponent intMemory
    DoSteppedWait
    ResetDataLines
    ClearComponents
End Sub

Sub Control_Keyb_To_Memory()
    ColourComponent intKeyb
    lneControl(7).BorderColor = lngLineColour
    lneControl(5).BorderColor = lngLineColour
    lneControl(3).BorderColor = lngLineColour
    lneControl(2).BorderColor = lngLineColour
    Animate_Line intControlLine, intVertical, lblControl, lneControl(7), intUp, True, True
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(5), intLeft, False, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(3), intLeft, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(2), intDown, True, True
    ColourComponent intMemory
    DoSteppedWait
    ResetControlLines
    ClearComponents
End Sub


'This is for copy
Sub Data_ALU_To_Memory()
    ColourComponent intALU
    lneData(0).BorderColor = lngLineColour
    lneData(1).BorderColor = lngLineColour
    lneData(2).BorderColor = lngLineColour
    Animate_Line intDataLine, intVertical, lblData, lneData(0), intDown, True, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(1), intRight, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(2), intDown, True, True
    ColourComponent intMemory
    DoSteppedWait
    ResetDataLines
    ClearComponents
End Sub

'This is for output
Sub Control_CU_To_Scr()
    ColourComponent intCU
    lneControl(4).BorderColor = lngLineColour
    lneControl(5).BorderColor = lngLineColour
    lneControl(6).BorderColor = lngLineColour
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(5), intRight, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(6), intUp, False, True
    ColourComponent intScr
    DoSteppedWait
    ResetControlLines
    ClearComponents
End Sub
Sub Data_CU_To_Scr()
    ColourComponent intCU
    lneData(4).BorderColor = lngLineColour
    lneData(5).BorderColor = lngLineColour
    lneData(6).BorderColor = lngLineColour
    Animate_Line intDataLine, intVertical, lblData, lneData(4), intDown, True, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(5), intRight, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(6), intUp, False, True
    ColourComponent intScr
    DoSteppedWait
    ResetDataLines
    ClearComponents
End Sub


Private Sub DoSteppedWait()
    If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
    
    'Do wait if in stepped mode
    If glbAnimateType = intStepped Then
        tmrWait.Enabled = True
        Do
            DoEvents
            
        Loop While tmrWait.Enabled = True
    End If
End Sub

'=================================================
' Bottom level Animation
'=================================================

'bolted on ---------------------------------------

Sub Animate_Line(ByVal intControlOrDataLine As Integer, ByVal intOrientation As Integer, ByRef lblLabel As Label, ByRef lneLine As Line, ByVal intDirection As Integer, ByVal blnStartInComponent, ByVal blnEndInComponent As Boolean)
    Dim intXDisplace As Integer
    Dim intYDisplace As Integer
    Dim intStartPos As Integer
    Dim intEndPos As Integer
    Dim intStep As Integer
    Dim intPos As Integer
    Dim intHalfLabelWidth As Integer
    
    'If code has been stopped, don't animate
    If CodeState <> 2 Or glbAnimateType = intNone Then Exit Sub
    lneLine.BorderColor = lngLineSectionColour
        
    'Init variables for label animat sub
    intAnimateOrientation = intOrientation
    intAnimateDirection = intDirection
    intAnimateLine = intControlOrDataLine
    
    'Calculate X Displacement
    intHalfLabelWidth = lblLabel.Width / 2
    
    
    If intOrientation = intHorizontal Then
        'Work out start positions for horizontal loop
        lblLabel.Top = lneLine.Y1 - lblLabel.Height - intUpFromLine
        
        If glbAnimateType = intFull Then
            'Full animation
            If intDirection = intLeft Then
                intAnimateStartPos = lneLine.X2 - intHalfLabelWidth
                intAnimateEndPos = lneLine.X1 - intHalfLabelWidth
                intAnimateStep = -1
            ElseIf intDirection = intRight Then
                intAnimateStartPos = lneLine.X1 - intHalfLabelWidth
                intAnimateEndPos = lneLine.X2 - intHalfLabelWidth
                intAnimateStep = 1
            End If
            lblLabel.Left = intAnimateStartPos
        Else
            'Stepped animation
            lblLabel.Left = (lneLine.X1 + ((lneLine.X2 - lneLine.X1) / 2)) - intHalfLabelWidth
            tmrWait.Enabled = True
            Do
                DoEvents
                
            Loop While tmrWait.Enabled = True
            lblLabel.Top = -1000
            Exit Sub
        End If
    ElseIf intOrientation = intVertical Then
        
        'Work out start positions for vertical loop
        lblLabel.Left = lneLine.X1 - intHalfLabelWidth
        
        If glbAnimateType = intFull Then
            'Full animation
            If intDirection = intUp Then
                If blnStartInComponent = True Then
                    intAnimateStartPos = lneLine.Y2 + intDownFromLine
                Else
                    intAnimateStartPos = lneLine.Y2 - lblLabel.Height - intUpFromLine
                End If
                If blnEndInComponent Then
                    intAnimateEndPos = lneLine.Y1 - lblLabel.Height - intUpFromLine
                Else
                    intAnimateEndPos = lneLine.Y1 + intDownFromLine
                End If
                intAnimateStep = -1
                
            ElseIf intDirection = intDown Then
                
                If blnStartInComponent = True Then
                    intAnimateStartPos = lneLine.Y1 - lblLabel.Height - intUpFromLine
                Else
                    intAnimateStartPos = lneLine.Y1 + intDownFromLine
                End If
                If blnEndInComponent = True Then
                    intAnimateEndPos = lneLine.Y2 + intDownFromLine
                Else
                    intAnimateEndPos = lneLine.Y2 - lblLabel.Height - intUpFromLine
                End If
                intAnimateStep = 1
            End If
            lblLabel.Top = intAnimateStartPos
        Else
            'Stepped animation
            lblLabel.Top = lneLine.Y1 + ((lneLine.Y2 - lneLine.Y1) / 2) - (lblLabel.Height / 2)
            tmrWait.Enabled = True
            Do
                DoEvents
                
            Loop While tmrWait.Enabled = True
            lblLabel.Top = -1000
            Exit Sub
        End If
    End If
    
    Multiply intAnimateStep, 2
    
    tmrAnimate.Enabled = True
    Do While tmrAnimate.Enabled = True
         DoEvents
         
    Loop


End Sub

'========options stuff
Private Sub cmbAnimate_Click()
    Select Case cmbAnimate.ListIndex
    Case intNone
        'No animation
        glbAnimateType = intNone
        sldSpeed.Enabled = False
        
        'End if code is running so disable timers
        If CodeState = 2 Then
            tmrAnimate.Enabled = False
            tmrWait.Enabled = False
            lblData.Top = -1000
            lblControl.Top = -1000
        End If
    Case intStepped
        'Stepped animation
        glbAnimateType = intStepped
        sldSpeed.Enabled = True
        sldSpeed.Min = 100
        sldSpeed.Max = 1500
        sldSpeed.Value = glbAnimateStepWait
    Case intFull
        'Full animation
        glbAnimateType = intFull
        sldSpeed.Enabled = True
        
        sldSpeed.Min = 10
        sldSpeed.Max = 110
        sldSpeed.Value = glbAnimateSpeed
    End Select
     
End Sub

Private Sub rtbScreen_KeyDown(KeyCode As Integer, Shift As Integer)
    'Copy the window
    Select Case KeyCode
    Case vbKeyControl
        ControlPressed2 = True
    Case ControlPressed2 And vbKeyC
        'Copy window
        CopyAnimateScreen
        KeyCode = 0
    Case Else
        ControlPressed2 = False
    End Select
End Sub

Private Sub rtbScreen_KeyUp(KeyCode As Integer, Shift As Integer)
    'Remove control flag
    If KeyCode = vbKeyControl Then
        ControlPressed2 = False
    End If
End Sub

Private Sub rtbScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Show copy menu
    If Button = vbRightButton Then
        glbFormToCopy = intAnimate
        PopupMenu frmCopy.mnuCopy
    End If
End Sub

Private Sub sldSpeed_Change()
    'Update speeds
    sldSpeed_Click
End Sub

Private Sub sldSpeed_Click()
    'Update the timers
    With sldSpeed
        If glbAnimateType = intFull Then
            tmrAnimate.Interval = .Value
        ElseIf glbAnimateType = intStepped Then
            tmrWait.Duration = .Value
        End If
    End With
End Sub



'Wait timer
Private Sub tmrWait_Tick(ByVal TimeRemaining As Long)
    'Do events
    DoEvents
End Sub

Private Sub tmrWait_Timer()
    'Nout
End Sub

Private Sub tmrAnimate_Timer(ByVal Milliseconds As Long)
    'Slow animation timer
    If intAnimateLine = intControlLine Then
        AnimateLabel lblControl
    ElseIf intAnimateLine = intDataLine Then
        AnimateLabel lblData
    End If
End Sub

Sub AnimateLabel(ByVal lblLabel As Label)
    'Move the label
    If intAnimateOrientation = intVertical Then
        
        If intAnimateDirection = intUp Then
                
            'Ensure label doesnt go too far up
            If lblLabel.Top + intAnimateStep < intAnimateEndPos Then
                lblLabel.Top = intAnimateEndPos
                tmrAnimate.Enabled = False
            Else
               lblLabel.Top = lblLabel.Top + intAnimateStep
            End If
        ElseIf intAnimateDirection = intDown Then
        
            'Ensure label doesnt go too far down
            If lblLabel.Top + intAnimateStep > intAnimateEndPos Then
                lblLabel.Top = intAnimateEndPos
                tmrAnimate.Enabled = False
            Else
               lblLabel.Top = lblLabel.Top + intAnimateStep
            End If
        End If
    ElseIf intAnimateOrientation = intHorizontal Then
    
    
        If intAnimateDirection = intLeft Then
        
            'Ensure label doesn't go too far left
            If lblLabel.Left + intAnimateStep < intAnimateEndPos Then
                lblLabel.Left = intAnimateEndPos
                tmrAnimate.Enabled = False
            Else
                lblLabel.Left = lblLabel.Left + intAnimateStep
            End If
            
        ElseIf intAnimateDirection = intRight Then
        
            'Ensure label doesn't go too far right
            If lblLabel.Left + intAnimateStep > intAnimateEndPos Then
                lblLabel.Left = intAnimateEndPos
                tmrAnimate.Enabled = False
            Else
                lblLabel.Left = lblLabel.Left + intAnimateStep
            End If
        End If
    End If
End Sub

'=================================================
'Line Highlights
'=================================================


Private Sub ResetAllLines()
    'Reset both control and data lines
    ResetControlLines
    ResetDataLines
End Sub

Private Sub ResetControlLines()
    'Reset Control Line colours
    lneControl(0).BorderColor = &HFF00&
    lneControl(1).BorderColor = &HFF00&
    lneControl(2).BorderColor = &HFF00&
    lneControl(3).BorderColor = &HFF00&
    lneControl(4).BorderColor = &HFF00&
    lneControl(5).BorderColor = &HFF00&
    lneControl(6).BorderColor = &HFF00&
    lneControl(7).BorderColor = &HFF00&
End Sub

Private Sub ResetDataLines()
    'Reset data Line colours
    lneData(0).BorderColor = &O0
    lneData(1).BorderColor = &O0
    lneData(2).BorderColor = &O0
    lneData(3).BorderColor = &O0
    lneData(4).BorderColor = &O0
    lneData(5).BorderColor = &O0
    lneData(6).BorderColor = &O0
    lneData(7).BorderColor = &O0
End Sub


Private Sub Form_Load()

    'Setup the full animation timer
    Set tmrAnimate = New ccrpTimer
    With tmrAnimate
        .EventType = TimerPeriodic
        .Interval = glbAnimateSpeed
        .Enabled = False
    End With

    'Setup the waitanimate timer
    Set tmrWait = New ccrpCountdown
    With tmrWait
        .Duration = glbAnimateStepWait
        .Interval = 20
        .Enabled = False
    End With
    
    'Set the animation options speed
    cmbAnimate.ListIndex = glbAnimateType
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Hide form
    
    'stop any animation
    StopAnimation
    
    If CloseAllForms = False Then
        Cancel = True
        HideAWindow
        Me.Hide
    Else
        'Kill timer
        KillTimers
        Unload Me
        Set frmAnimate = Nothing
    End If
End Sub

Sub KillTimers()
    'kill timer
    tmrAnimate.Enabled = False
    Set tmrAnimate = Nothing
    tmrWait.Enabled = False
    Set tmrWait = Nothing
End Sub

Sub StopAnimation()
    'Stop animation
    frmAnimate.txtKeyboard = Empty
    frmAnimate.txtKeyboard.Locked = True
    lblControl.Top = -50
    lblData.Top = -50
    StopTimers
    ClearComponents
End Sub

Sub StopTimers()
    'Stop timer
    tmrAnimate.Enabled = False
    tmrWait.Enabled = False
End Sub
Private Sub SetCursorPos()
        
    'add key presses to end
    txtKeyboard.SelStart = Len(txtKeyboard.Text)

End Sub

Private Sub txtKeyboard_KeyDown(KeyCode As Integer, Shift As Integer)
    'Vet keys that the user presses
    
    'If locked do nothing
    If txtKeyboard.Locked = True Then
         KeyCode = 0
         blnhidekey = True
         Exit Sub
    End If
    
    blnhidekey = False
    
    Select Case KeyCode
    Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight
    
        'Disable cursor movements
        KeyCode = 0
        
    Case vbKeyBack
        'Enable keyback
        If Len(txtKeyboard) > 0 Then
            'Do backspace on run form
            frmRun.DoBackspace
            
            SetCursorPos
            
            'Backspace on screen
            ScreenDoBackspace
        Else
            Beep
        End If
    Case vbKeyControl
        
        'Set control flag
        ControlPressed = True
        
    Case ControlPressed And vbKeyC  'Allow Copy
    Case ControlPressed And vbKeyV
    
        'Disable pasting
        blnhidekey = True
        
    Case vbKeyReturn
        'Do new line on run form
        KeyCode = 0
        blnhidekey = True
        frmRun.GotInput
    Case Asc("0") To Asc("9")
        If Shift = 0 Then
            'Digits ok
            SetCursorPos
            ScreenAddCharacter Chr(KeyCode)
        Else
            blnhidekey = True
        End If
    Case vbKeySubtract, 189, vbKeyAdd, 187
        
        '-/+ key pressed
        If KeyCode = 187 And Shift = 1 Or KeyCode = 189 And Shift = 0 Or KeyCode = vbKeyAdd Or KeyCode = vbKeySubtract Then
            If Len(txtKeyboard) > 0 Then
                blnhidekey = True
                Exit Sub
            Else
                'Convert keycodes into characters
                SetCursorPos
                If KeyCode = vbKeySubtract Or KeyCode = 189 Then
                    ScreenAddCharacter "-"
                ElseIf KeyCode = vbKeyAdd Or KeyCode = 187 Then
                    ScreenAddCharacter "+"
                End If
            End If
        Else
            blnhidekey = True
        End If
    Case vbKeyShift
    Case Else
        'Any other key
        blnhidekey = True
        KeyCode = 0
    End Select

End Sub

Public Sub AddCharacter(ByVal strChar As String)
    'Add a character to the keyboard text box
    If frmRun.rtbOutput.Locked = False Then
        txtKeyboard = txtKeyboard + strChar
        
        'Add character to screen
        ScreenAddCharacter strChar
    End If
End Sub

Public Sub ClearText()
    'Clear the keyboard text box
    txtKeyboard = Empty
End Sub

Public Sub DoBackspace()
    If frmRun.rtbOutput.Locked = False Then
        'Remove charcter from keyboard text box
        txtKeyboard = Left(txtKeyboard, Len(txtKeyboard) - 1)
        
        'Remove character from screen
        ScreenDoBackspace
    End If
End Sub

Private Sub ScreenDoBackspace()
    'Remove character from screen text box
    rtbScreen.SelStart = Len(rtbScreen.Text) - 1
    rtbScreen.SelLength = 1
    rtbScreen.SelText = Empty
End Sub

Private Sub ScreenAddCharacter(ByVal strChar As String)
    'Add character to the screen text box
    rtbScreen.SelStart = Len(rtbScreen.Text)
    rtbScreen.SelText = strChar
End Sub

Private Sub txtKeyboard_KeyPress(KeyAscii As Integer)
    'Remove certain key strokes vetted by keydown method
    If blnhidekey = True Then
        KeyAscii = 0
    Else
        'Don't add the back char to run form
        If KeyAscii <> vbKeyBack Then
            frmRun.AddCharacter Chr(KeyAscii)
        End If
    End If
    blnhidekey = False
End Sub

Private Sub txtKeyboard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Show copy menu
    If Button = vbRightButton Then
        glbFormToCopy = intRun
        PopupMenu frmCopy.mnuCopy
    End If
End Sub

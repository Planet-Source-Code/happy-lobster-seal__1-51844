VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAnimate 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   443
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   Begin VB.Timer tmrAnimate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6600
      Top             =   2640
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   1395
      Left            =   6480
      TabIndex        =   18
      Top             =   4500
      Width           =   195
   End
   Begin VB.TextBox txtExplanation 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1680
      Width           =   2115
   End
   Begin VB.TextBox txtFlag 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "+"
      Top             =   1440
      Width           =   315
   End
   Begin VB.TextBox txtIndx 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "+32888"
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtCurrentInstruction 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Load Acc,#23"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtAcc 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "+32666"
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   195
      Left            =   6540
      TabIndex        =   12
      Top             =   5520
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox rtbScreen 
      Height          =   1515
      Left            =   6720
      TabIndex        =   10
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2672
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmAnimate.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
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
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   3120
      TabIndex        =   8
      Top             =   4680
      Width           =   2775
   End
   Begin VB.TextBox txtKeyboard 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "+32888"
      Top             =   4680
      Width           =   1395
   End
   Begin VB.ListBox lstCode 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1380
      Left            =   360
      TabIndex        =   4
      Top             =   4680
      Width           =   2775
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
      Left            =   960
      TabIndex        =   11
      Top             =   300
      Width           =   4335
   End
   Begin VB.Label lblControlUnit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Control Unit"
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   660
      Width           =   2415
   End
   Begin VB.Shape shpControlUnit 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   2880
      Top             =   900
      Width           =   2415
   End
   Begin VB.Shape shpFlag 
      Height          =   315
      Left            =   1380
      Top             =   1440
      Width           =   315
   End
   Begin VB.Shape shpIndx 
      Height          =   315
      Left            =   1140
      Top             =   1800
      Width           =   735
   End
   Begin VB.Shape shpAccumulator 
      Height          =   315
      Left            =   1140
      Top             =   1080
      Width           =   735
   End
   Begin VB.Shape shpALU 
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   1275
      Left            =   960
      Top             =   960
      Width           =   1095
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   7
      X1              =   544
      X2              =   544
      Y1              =   244
      Y2              =   296
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   5
      X1              =   284
      X2              =   544
      Y1              =   248
      Y2              =   248
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
      Left            =   9600
      TabIndex        =   7
      Top             =   480
      Width           =   1155
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
      Left            =   4320
      TabIndex        =   6
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   4
      X1              =   260
      X2              =   260
      Y1              =   148
      Y2              =   224
   End
   Begin VB.Label lblScreen 
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
      Left            =   6720
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblKeyboard 
      Alignment       =   1  'Right Justify
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
      Left            =   6900
      TabIndex        =   2
      Top             =   4440
      Width           =   795
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
      Left            =   360
      TabIndex        =   1
      Top             =   4440
      Width           =   5475
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Arithmetic and Logic Unit"
      Height          =   435
      Left            =   960
      TabIndex        =   0
      Top             =   540
      Width           =   1095
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   5
      X1              =   260
      X2              =   520
      Y1              =   224
      Y2              =   224
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   6
      X1              =   520
      X2              =   520
      Y1              =   152
      Y2              =   224
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   3
      X1              =   200
      X2              =   256
      Y1              =   224
      Y2              =   224
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   3
      X1              =   224
      X2              =   280
      Y1              =   248
      Y2              =   248
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   2
      X1              =   196
      X2              =   196
      Y1              =   224
      Y2              =   296
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   7
      X1              =   520
      X2              =   520
      Y1              =   224
      Y2              =   296
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   1
      X1              =   96
      X2              =   192
      Y1              =   224
      Y2              =   224
   End
   Begin VB.Line lneControl 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      Index           =   0
      X1              =   96
      X2              =   96
      Y1              =   152
      Y2              =   224
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   2
      X1              =   220
      X2              =   220
      Y1              =   248
      Y2              =   296
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   6
      X1              =   544
      X2              =   544
      Y1              =   152
      Y2              =   248
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   220
      Y1              =   248
      Y2              =   248
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   4
      X1              =   284
      X2              =   284
      Y1              =   152
      Y2              =   248
   End
   Begin VB.Line lneData 
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   120
      Y1              =   152
      Y2              =   248
   End
   Begin VB.Shape shpScreen 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1935
      Left            =   6540
      Top             =   240
      Width           =   2175
   End
   Begin VB.Shape shpKeyboard 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   6720
      Top             =   4440
      Width           =   1755
   End
   Begin VB.Shape shpMemory 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   1815
      Left            =   180
      Top             =   4440
      Width           =   5895
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   2235
      Left            =   780
      Top             =   180
      Width           =   4695
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   6495
      Left            =   0
      Top             =   -120
      Width           =   6255
   End
End
Attribute VB_Name = "frmAnimate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const intLeftFromLine = 4
Private Const intRightFromLine = 4
Private Const intUpFromLine = 4
Private Const intDownFromLine = 4
Private Const intSleepVal = 10
Private Const intUp = 0
Private Const intDown = 1
Private Const intLeft = 0
Private Const intRight = 1
Private Const intControlLine = 0
Private Const intDataLine = 1
Private Const intHorizontal = 0
Private Const intVertical = 1

'Used for animation purposes
Private intAnimateOrientation As Integer
Private intAnimateDirection As Integer
Private intAnimateStartPos As Integer
Private intAnimateEndPos As Integer
Private intAnimateStep As Integer
Private intAnimateLine As Integer
Private blnFinishedAnimation As Boolean

Sub Animate()
    txtExplanation = ConvertOperation
    Select Case TempOperation
    Case 0 To 11
        Animate_ALU_Command
    Case 12 To 17, 27
        Animate_Get_Flag_Value
    Case 21
        Animate_Input
    Case 22, 23, 26, 28
        Animate_Output
    End Select

End Sub

Sub Animate_ALU_Command()
    Dim intTemp As Integer
    
    Select Case TempOperation
    Case 5 To 8
        'Neg,Inc,Dec,Clrz
        lblControl = TempCodeWithNoAddress
        Control_CU_To_ALU
    Case 0 To 4, 9 To 10
        'Arithmetic
        If TempOperand2 = 3 Then
            'Get immediate value
            lblControl = TempCodeWithNoAddress
            Control_CU_To_ALU
            lblData = GetVariableName
            Data_CU_To_Alu
        Else
            'Get variable/array from memory
            lblControl = ConvertOperation
            Control_CU_To_ALU
            
            'Test array index
            If IsArrayIndexOutOfRange = True Then
                Access_Memory intALU, "Get " + GetVariableName(True), ErrorMessages, True
            Else
                Access_Memory intALU, "Get " + GetVariableName(True), Format(GetValue), False
            End If
        End If
    Case 11
        'Copy
        lblControl = ConvertOperation
        Animate_Storage intALU, Format(LastRegister)
    End Select
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
            ErrorMessages = "Overflow"
            AnimateError intMemory, ErrorMessages
        ElseIf IsOutOfRange(Val(strStoreValue)) = True Then
            'Attempt to convert string
            ErrorMessages = "Overflow"
            AnimateError intMemory, ErrorMessages
        End If
    End If
End Sub

Sub AnimateError(intSource As Integer, strMessage As String)
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
    ElseIf TempOperation = 26 Then
        lblControl = "Clrs"
        Control_CU_To_Scr
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

'            lblData = Format(GetValue)
'            Data_CU_To_Scr
            
        End If

End Sub

Sub Animate_Get_Flag_Value()
    'Get Flag Value
    lblControl = "Get Flag Value"
    Control_CU_To_ALU
    lblData = txtFlag
    Data_ALU_To_CU
End Sub



Sub Access_Memory(ByVal intSource As Integer, ByVal strSend As String, ByVal strGet As String, ByVal blnError As Boolean, Optional ByVal intDestination As Integer)
    lblControl = strSend
    lblData = strGet
    
    'Animate fetch/get memory
    If intSource = intALU Then
        Control_ALU_To_Memory
        If blnError = False Then
            'Animate data returning
            Data_Memory_To_ALU
        Else
            'Animate error
            lblControl = strGet
            Control_Memory_To_CU
        End If
    ElseIf intSource = intCU Then
        Control_CU_To_Memory
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
        If blnError = False Then
            'Animate data returning
            Data_ALU_To_Memory
        Else
            'Animate error
            lblControl = strGet
            Control_Memory_To_CU
        End If
    ElseIf intSource = intKeyb Then
        Control_Keyb_To_Memory
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
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(3), intLeft, False, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(1), intLeft, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(0), intUp, False, True

End Sub

Sub Data_CU_To_Alu()
    Animate_Line intDataLine, intVertical, lblData, lneData(4), intDown, True, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(3), intLeft, False, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(1), intLeft, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(0), intUp, False, True
End Sub

Sub Control_ALU_To_Memory()
    Animate_Line intControlLine, intVertical, lblControl, lneControl(0), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(1), intRight, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(2), intDown, True, True
End Sub

Sub Control_CU_To_Memory()
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(3), intLeft, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(2), intDown, False, True
End Sub

Sub Data_Memory_To_ALU()
    Animate_Line intDataLine, intVertical, lblData, lneData(2), intUp, True, True
    Animate_Line intDataLine, intHorizontal, lblData, lneData(1), intLeft, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(0), intUp, False, True
End Sub

Sub Data_Memory_To_CU()
    Animate_Line intDataLine, intVertical, lblData, lneData(2), intUp, True, True
    Animate_Line intDataLine, intHorizontal, lblData, lneData(3), intRight, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(4), intUp, False, True
End Sub

Sub Data_Memory_To_Scr()
    Animate_Line intDataLine, intVertical, lblData, lneData(2), intUp, True, True
    Animate_Line intDataLine, intHorizontal, lblData, lneData(3), intRight, False, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(5), intRight, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(6), intUp, False, True
End Sub

Sub Data_ALU_To_CU()
    Animate_Line intDataLine, intVertical, lblData, lneData(0), intDown, True, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(1), intRight, False, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(3), intRight, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(4), intUp, False, True
End Sub

'These calls are used for animating error messages
Sub Control_ALU_TO_CU()
    Animate_Line intControlLine, intVertical, lblControl, lneControl(0), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(1), intRight, False, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(3), intRight, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intUp, False, True
End Sub

Sub Control_Memory_To_CU()
    Animate_Line intControlLine, intVertical, lblControl, lneControl(2), intUp, True, True
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(3), intRight, True, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intUp, False, True
End Sub

Sub Control_Scr_To_CU()
    Animate_Line intControlLine, intVertical, lblControl, lneControl(6), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(5), intLeft, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intUp, False, True
End Sub

'This is for input
Sub Control_CU_To_Keyb()
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(5), intRight, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(7), intDown, True, True
End Sub

Sub Data_Keyb_To_Memory()
    Animate_Line intDataLine, intVertical, lblData, lneData(7), intUp, True, True
    Animate_Line intDataLine, intHorizontal, lblData, lneData(5), intLeft, False, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(3), intLeft, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(2), intDown, True, True
End Sub

Sub Control_Keyb_To_Memory()
    Animate_Line intControlLine, intVertical, lblControl, lneControl(7), intUp, True, True
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(5), intLeft, False, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(3), intLeft, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(2), intDown, True, True
End Sub


'This is for copy
Sub Data_ALU_To_Memory()
    Animate_Line intDataLine, intVertical, lblData, lneData(0), intDown, True, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(1), intRight, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(2), intDown, True, True
End Sub

'This is for output
Sub Control_CU_To_Scr()
    Animate_Line intControlLine, intVertical, lblControl, lneControl(4), intDown, True, False
    Animate_Line intControlLine, intHorizontal, lblControl, lneControl(5), intRight, False, False
    Animate_Line intControlLine, intVertical, lblControl, lneControl(6), intUp, False, True
End Sub
Sub Data_CU_To_Scr()
    Animate_Line intDataLine, intVertical, lblData, lneData(4), intDown, True, False
    Animate_Line intDataLine, intHorizontal, lblData, lneData(5), intRight, False, False
    Animate_Line intDataLine, intVertical, lblData, lneData(6), intUp, False, True
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
    
    
    intAnimateOrientation = intOrientation
    intAnimateDirection = intDirection
    intAnimateLine = intControlOrDataLine
    
    'Calculate X Displacement
    intHalfLabelWidth = lblLabel.Width / 2

    lneLine.BorderColor = &HFF&
    
    
    If intOrientation = intHorizontal Then
        'Work out start positions for horizontal loop
        lblLabel.Top = lneLine.Y1 - lblLabel.Height - intUpFromLine
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
    ElseIf intOrientation = intVertical Then
        
    
        'Work out start positions for vertical loop
        lblLabel.Left = lneLine.X1 - intHalfLabelWidth
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
    End If
    
    tmrAnimate.Enabled = True
    Do While tmrAnimate.Enabled = True
         DoEvents
    Loop


End Sub

Private Sub tmrAnimate_Timer()
    
    If intAnimateLine = intControlLine Then
        AnimateLabel lblControl
    ElseIf intAnimateLine = intDataLine Then
        AnimateLabel lblData
    End If
    

End Sub

Sub AnimateLabel(ByVal lblLabel As Label)
    If intAnimateOrientation = intVertical Then
        lblLabel.Top = lblLabel.Top + intAnimateStep
        If lblLabel.Top = intAnimateEndPos Then
            tmrAnimate.Enabled = False
        End If
    ElseIf intAnimateOrientation = intHorizontal Then
        lblLabel.Left = lblLabel.Left + intAnimateStep
        If lblLabel.Left = intAnimateEndPos Then
            tmrAnimate.Enabled = False
        End If
    End If
End Sub

'=================================================
' Control Line Highlights
'=================================================


'bolted on---------------------------------
Sub HighlightALU_Mem_ControlLine()
    'Highlight ALU/Mem control line
    lneControl(0).BorderColor = &HFF&
    lneControl(1).BorderColor = &HFF&
    lneControl(3).BorderColor = &HFF&
    lneControl(4).BorderColor = &HFF&
End Sub
'------------------------------------------

Sub HighlightCU_ALU_ControlLine()
    'Highlight CU/ALU control line
    lneControl(0).BorderColor = &HFF&
    lneControl(1).BorderColor = &HFF&
    lneControl(2).BorderColor = &HFF&
End Sub

Sub HighlightCU_Mem_ControlLine()
    'Highlight CU/Mem control line
    lneControl(2).BorderColor = &HFF&
    lneControl(3).BorderColor = &HFF&
    lneControl(4).BorderColor = &HFF&
End Sub

Sub HighlightCU_Keyb_ControlLine()
    'Highlight CU/Keyb control line
    lneControl(2).BorderColor = &HFF&
    lneControl(3).BorderColor = &HFF&
    lneControl(5).BorderColor = &HFF&
    lneControl(7).BorderColor = &HFF&
End Sub


Sub HighlightCU_Scr_ControlLine()
    'HighlightCU/Scr Control line
    lneControl(2).BorderColor = &HFF&
    lneControl(3).BorderColor = &HFF&
    lneControl(5).BorderColor = &HFF&
    lneControl(6).BorderColor = &HFF&
End Sub
'=================================================
'Data Line Highlights
'=================================================

'bolted on---------------------------------
Sub HighlightALU_Mem_DataLine()
    'Highlight ALU/Mem control line
    lneData(0).BorderColor = &HFF&
    lneData(1).BorderColor = &HFF&
    lneData(3).BorderColor = &HFF&
    lneData(4).BorderColor = &HFF&
End Sub
'------------------------------------------

Sub HighlightCU_ALU_DataLine()
    'Highlight CU/ALU data line
    lneData(0).BorderColor = &HFF&
    lneData(1).BorderColor = &HFF&
    lneData(2).BorderColor = &HFF&
End Sub

Sub HighlightCU_Mem_DataLine()
    'Highlight CU/Mem data line
    lneData(2).BorderColor = &HFF&
    lneData(3).BorderColor = &HFF&
    lneData(4).BorderColor = &HFF&
End Sub

Sub HighlightCU_Keyb_DataLine()
    'Highlight CU/Keyb data line
    lneData(2).BorderColor = &HFF&
    lneData(3).BorderColor = &HFF&
    lneData(5).BorderColor = &HFF&
    lneData(7).BorderColor = &HFF&
End Sub

Sub HighlightCU_Scr_DataLine()
    'HighlightCU/Scr data line
    lneData(2).BorderColor = &HFF&
    lneData(3).BorderColor = &HFF&
    lneData(5).BorderColor = &HFF&
    lneData(6).BorderColor = &HFF&
End Sub

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
    Dim strT As String
    Dim a As Integer
    Dim b As Integer
    
    For a = 1 To 18
        strT = "012345678901234567869"
        strT = Empty
        For b = 1 To Int(Rnd() * 8) + 4
            strT = strT + Chr(Int(Rnd() * (Asc("Z") - Asc("A"))) + Asc("A"))
        Next
        lstCode.AddItem strT + " 1212"
    Next
        
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Hide form
    If CloseAllForms = False Then
        Cancel = True
        HideAWindow
        Me.Hide
    Else
        Unload Me
        Set frmAnimate = Nothing
    End If
End Sub



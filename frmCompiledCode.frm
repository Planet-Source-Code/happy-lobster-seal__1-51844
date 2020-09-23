VERSION 5.00
Begin VB.Form frmCode 
   Caption         =   "(F11) Code In Memory"
   ClientHeight    =   4605
   ClientLeft      =   405
   ClientTop       =   5985
   ClientWidth     =   6660
   Icon            =   "frmCompiledCode.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4605
   ScaleWidth      =   6660
   Begin VB.ListBox lstCode 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      ItemData        =   "frmCompiledCode.frx":08D2
      Left            =   0
      List            =   "frmCompiledCode.frx":08D4
      TabIndex        =   0
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmCode"
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

Dim OldIndex As Integer
Dim ControlPressed As Boolean

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Hide form
    If CloseAllForms = False Then
    
        Cancel = True
        HideAWindow
        Me.Hide
    Else
        Unload Me
        Set frmCode = Nothing
    End If
End Sub

Sub ShowCode()
    Dim a As Integer
    Dim LabelStr As String
    Dim KeywordText As String
    Dim ListString As String
    
    'Add compiled code to listbox
    LineNumber = 0
    lstCode.Clear
    frmAnimate.lstCode.Clear
    glbShowCodeView = True
    Do
        'Set temp op variables
        TempOperation = Operation(LineNumber)
        TempOperand1 = Operand1(LineNumber)
        TempOperand2 = Operand2(LineNumber)
        TempOperand3 = Operand3(LineNumber)
        TempOperandText = OperandText(LineNumber)
        
        'Search for label positions
        LabelStr = Empty
        For a = 0 To LabelCount - 1
            If LineNumber = LabelPos(a) Then
                If LabelStr = Empty Then
                    LabelStr = Capitalise(LabelName(a)) + ":"
                ElseIf InStr(LabelStr, "*") = 0 Then
                    LabelStr = "*" + LabelStr
                End If
            End If
        Next
        
        'Set the label str case
        LabelStr = SetCase(LabelStr)
        If glbVariableLabelUC = True Then
            LabelStr = UCase(LabelStr)
        End If
        
        '3 spaces inserted for arrow to point to current line running
        
        'add command to list with padded spaces
        If LineNumber < LineCount Then

            ListString = "   " & LabelStr + Space(LabelLength - Len(LabelStr)) + ConvertOperation
            lstCode.AddItem ListString
            frmAnimate.lstCode.AddItem ListString
        Else
            'if no more code after label
            'no need to print operations
            If LabelStr <> Empty Then
                ListString = "   " & LabelStr
            
                lstCode.AddItem ListString
                frmAnimate.lstCode.AddItem ListString
            End If
        End If
        
        Inc LineNumber
        'are we at end of program?
        If LineNumber = LineCount + 1 Then
            'add 1 to linecount to include
            'lablels with no code after them
            Exit Do
        End If
    Loop
    glbShowCodeView = False
End Sub



Private Sub Form_Resize()
    'Resize list box
    With lstCode
        .Width = ScaleWidth
        If WindowState = vbNormal Then
            If Me.Height - 420 > 0 Then
                .Height = Me.Height - 420
                Me.Height = .Height + 420
            End If
        ElseIf WindowState = vbMaximized Then
            .Height = ScaleHeight
        End If
    End With
End Sub


Private Sub lstCode_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Show copy menu
    If Button = vbRightButton Then
        glbFormToCopy = intCodeInMemory
        PopupMenu frmCopy.mnuCopy
    End If
End Sub

Private Sub lstCode_KeyDown(KeyCode As Integer, Shift As Integer)
    'Copy the window
    Select Case KeyCode
    Case vbKeyControl
        ControlPressed = True
    Case ControlPressed And vbKeyC
        'Copy window
        CopyCodeInMemory
        KeyCode = 0
    Case Else
        ControlPressed = False
    End Select
End Sub

Private Sub lstCode_KeyUp(KeyCode As Integer, Shift As Integer)
    'Remove control flag
    If KeyCode = vbKeyControl Then
        ControlPressed = False
    End If
End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVariables 
   Caption         =   "(F8) Location Table"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3030
   Icon            =   "frmVariables.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   3030
   Begin MSComctlLib.ListView lstvVariables 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   5106
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmVariables"
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

Dim ControlPressed As Boolean

Sub ShowVariableNames()
    Dim Row As Integer
    Dim V As Integer
    
    'Add variable names to form
    With frmVariables.lstvVariables
        
        'clear all
        .ListItems.Clear
        frmAnimate.lstVariables.Clear
        
        'Add acc
        .ListItems.Add
        .ListItems(1).Text = "ACC"
        
        'Add indx
        .ListItems.Add
        .ListItems(2).Text = "INDX"
        
        'Add flag
        .ListItems.Add
        .ListItems(3).Text = "FLAG"
        
        Row = 4
        For V = 0 To VariableCount - 1
            .ListItems.Add
            .ListItems(Row).Text = VariableName(V)
            frmAnimate.lstVariables.AddItem VariableName(V)
            Row = Row + 1
        Next
        
        'Add array names
        For V = 0 To ArrayCount - 1
            .ListItems.Add
            .ListItems(Row).Text = ArrayName(V) + "()"
            frmAnimate.lstVariables.AddItem ArrayName(V) + "()"
            Row = Row + 1
        Next
    End With
    frmVariables.Refresh
End Sub

Sub ShowVariableValues()
    Dim Row As Integer
    Dim V As Integer
    Dim E As Integer
    Dim TempStr As String
    Dim FlagVal As String
    
    
    'Add variable names to form
    With frmVariables.lstvVariables
                
        'Add register values
        .ListItems(1).SubItems(1) = Format(Acc)
        .ListItems(2).SubItems(1) = Format(Indx)
        .ListItems(3).SubItems(1) = GetFlagValue
        
        'Update animation form
        With frmAnimate
            .txtAcc = Format(Acc)
            .txtIndx = Format(Indx)
            .txtFlag = GetFlagValue
        End With
        
        'add variable values
        Row = 4
        For V = 0 To VariableCount - 1
            .ListItems(Row).SubItems(1) = VariableValue(V)
            frmAnimate.lstVariables.List(Row - 4) = VariableName(V) + " " + Format(VariableValue(V))
            Row = Row + 1
        Next
        
        'Add array values
        For V = 0 To ArrayCount - 1
            TempStr = Empty
            For E = 0 To ArrayElements(V)
                TempStr = TempStr + Format(ArrayValue(V, E)) + " "
            Next
            .ListItems(Row).SubItems(1) = TempStr
            frmAnimate.lstVariables.List(Row - 4) = ArrayName(V) + " " + TempStr
            Row = Row + 1
        Next
    End With
    
    'Pause program
    If StepMode = True Then
        If CodeState = 2 Then
            Me.Caption = ConvertOperation
            DoNextInstruction = False
            'frmRun.Enabled = False
            
            If TempOperation <> 21 Then
                Do
                    DoEvents
                Loop While DoNextInstruction = False And CodeState = 2
            End If
            'frmRun.Enabled = True
            frmRun.SetFocus
        End If
        frmVariables.Refresh
    End If
End Sub

Private Sub Form_Load()
    ShowVariableNames
    ShowVariableValues
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Hide form
    If CloseAllForms = False Then
        Cancel = True
        HideAWindow
        Me.Hide
    Else
        Unload Me
        Set frmVariables = Nothing
    End If
End Sub

Private Sub Form_Resize()
    Dim NewWidth As Integer
    
    'Resize list box
    With lstvVariables
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

Private Sub lstvVariables_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Show copy menu
    If Button = vbRightButton Then
        glbFormToCopy = intVariables
        PopupMenu frmCopy.mnuCopy
    End If
End Sub

Private Sub lstvVariables_KeyDown(KeyCode As Integer, Shift As Integer)
    'Copy the window
    Select Case KeyCode
    Case vbKeyControl
        ControlPressed = True
    Case ControlPressed And vbKeyC
        'Copy window
        CopyVariables
        KeyCode = 0
    Case Else
        ControlPressed = False
    End Select
End Sub

Private Sub lstvVariables_KeyUp(KeyCode As Integer, Shift As Integer)
    'Remove control flag
    If KeyCode = vbKeyControl Then
        ControlPressed = False
    End If
End Sub

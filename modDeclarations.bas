Attribute VB_Name = "modDeclarations"
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

Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Public Const EM_LINESCROLL = &HB6

Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOACTIVATE = &H10

Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const FLOAT = 1
Global Const SINK = 0

Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long


Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETFIRSTVISIBLELINE = &HCE

'Old variables used in old find/replace box
Public glbFirstTime As Boolean
Public glbEditorVisible As Boolean

'Variables for animation
Public glbAnimateType As Integer    'Animation type
Public glbAnimateStepWait As Integer 'Stepped animation wait time
Public glbAnimateSpeed As Integer   'Full animation speed

'Animation constants
Global Const intNone = 0
Global Const intStepped = 1
Global Const intFull = 2

'Constants for animation
Global Const intALU = 1
Global Const intMemory = 2
Global Const intCU = 3
Global Const intScr = 4
Global Const intKeyb = 5
Global Const intAcc = 6
Global Const intIndx = 7

'Original screen width and height
Public intScreenWidth As Single
Public intScreenHeight As Single

'Special Characters
Public Const EOSchar = "Ä"              'End of symbol character
Public Const EOLchar = "Å"              'End of line character


Public CodeDirty As Boolean             'Flag if code has been adjusted
Public CloseAllForms As Boolean         'Are we shutting down the system

Public LastLineNumber As Integer        'Last line to executed
Public LineNumber As Integer            'Line number of code in memory
Public LineCount As Integer             'Total number of lines

Public glbFirstTimeRun As Boolean       'Is this the first time we are running this program

'Has the line been validated?
Public glbLastLineValidated As Boolean

Public DoNextInstruction As Boolean     'Do next instruction?

Public LabelName() As String            'Label names
Public LabelCount As Integer            'Total number of labels
Public LabelPos() As Integer            'Label positions

Public ArrayName(255) As String         'Array names
Public ArrayCount As Integer            'Number of arrays
Public ArrayValue(255, 255) As Integer  'Array values
Public ArrayElements(255) As Integer    'Number of elements in array

Public VariableName(255) As String      'Variable names
Public VariableCount As Integer         'Number of variables
Public VariableValue(255) As Integer

Public Const intMaxLinesPaste = 128     'Max lines allowed for pasting
Public Const intMaxLinesColour = 128    'Max lines allowed to colour in file
Public Const intOperationIndexes = 128  'Presize of all code
Public Const intLabelIndexes = 128

Public Operation() As Integer           'Operation to execute
Public Operand1() As Integer            'Operand 1
Public Operand2() As Integer            'Operand 2
Public Operand3() As Integer            'Operand 3
Public OperandText() As String          'Output text

Public StepMode As Boolean              'Is program to be run in step mode
Public WindowLineCount As Integer       'Number of window lines a program takes

Public CodeState As Integer             'State of operation to be run
Public Acc As Long                      'Acc value
Public Indx As Long                     'Indx value
Public LastRegister As Integer          'Last register value or FLAG

Public GotUserValue As Boolean          'Flag if user input has been o
Public OptionsEnabled As Boolean        'Flag if user options are enabled

Public ReturnPos(255) As Integer        'Return positions
Public ReturnNo As Integer              'Return counter

Public EditorLine As Integer            'current line cursor
Public LineEditted As Boolean           'has current line been editted?


Public SyntaxCode As String             'Temp syntax code
Public Sym As String                    'Current symbol
Public NormSym As String                'Upper case sym
Public NonSym As String                 'Actual symbol
Public SymPos As Integer                'Symbol position
Public ErrorMessages As String          'Error messages
Public WindowLineNumber As Integer      'Line number relative window
Public GroupedCode As String            'Grouped code string
Public CutCopyVisible As Boolean        'Is Cut&Copy available?
Public Console As Boolean               'Are we calling frmEditor.checklabels from console?

Public CodeExplainAddress As String
Public ExplainAddress As String         'Address of command explained
Public ExplainLabelTo As String         'Jump to Label of command explained
Public ExplainDeclaration As String     'Declaration explained
Public ExplainCode As String            'Actual code of explained line
Public ExplainLabel As String           'Label on a line

Public TempOperation As Integer         'Temp operation
Public TempOperand1 As Integer          'Temp operand 1
Public TempOperand2 As Integer          'Temp operand 2
Public TempOperand3 As Integer          'Temp operand 3
Public TempOperandText As String        'Temp literal text
Public TempRegDev As Integer            'Temp register or device
Public TempReminder As String           'Temp reminder
Public TempLabelName As String          'Temp label name
Public TempLabelName2 As String         'Temp label name 2
Public TempVariable As String           'Temp variable name
Public TempCodeWithNoAddress As String  'Temp line of code without address
Public TempRandomNumber As Integer      'Tem random number(for animation)

Public LabelLength As Integer           'Max length of label
Public VariableLength As Integer        'Max length of variable

Public NumberOfWindows As Integer       'Number of windows visible

'User settings variables
'-----------------------------------------
Public glbFontSize As Integer           'System font size
Public glbFont As String                'Symtem font
Public glbVariableLabelUC As Boolean    'Uppercase variable
Public glbShowCodeView As Boolean
Public glbColourSyntax As Boolean       'Are we gonna colour syntax
Public glbAutoCheckSyntax As Boolean    'Automatically check syntax?
Public glbQuickNoCheck As Boolean

Public glbScreenWidth As Single         'Original screen width
Public glbScreenHeight As Single        'Original screen height

Public glbScreenSize As Byte
Public glbClearScreen As Boolean        'Should we clear screen on each new run

Public glbFormToCopy As Integer         'The form we wish to copy
Public Const intAnimate = 0
Public Const intRun = 1
Public Const intConsole = 2
Public Const intFullScreen = 3
Public Const intKeyboard = 4
Public Const intVariables = 5
Public Const intCodeInMemory = 6

'Declaraton of colour variables
Public glbPunctuationCol As Long
Public glbLabelCol As Long
Public glbVariableCol As Long
Public glbCommandCol As Long
Public glbRegisterCol As Long
Public glbDeviceCol As Long
Public glbNumberCol As Long
Public glbCommentCol As Long
Public glbErrorCol As Long
Public glbLiteralCol As Long
Public glbProgramTextColour As Long     'Program text colour
Public glbProgramBackColour As Long     'Program background colour

Public glbEditorBackColour As Long      'Editor back colour
Public glbRunTextColour As Long         'Text colour when running
Public glbConsoleTextColour As Long     'Console text colour
Public glbConsoleBackColour As Long     'Console back colour

Public glbShowCodeInMemory As Boolean
Public glbShowLocationTable As Boolean
Public glbShowComputerArchitecture As Boolean
Public glbShowConsole As Boolean

Public glbDisableColour As Boolean      'Show colour syntaxing be disabled

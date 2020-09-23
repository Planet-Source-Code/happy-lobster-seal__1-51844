Attribute VB_Name = "modBouncingBalls"
Option Explicit
'********************************************************************************
'* Flickerless, Smooth Animation in VB with just pure VB Code!!!!
'* By Douglas J. Puckett
'* 12/16/2000
'* Okay, so I'm giving away the secret of my animation technique that I use in my
'* screensavers.  Yes, you can do very smooth, flickerless animation
'* in VB if you do it the right way.  The trick is to use hidden picture boxes
'* as a screen buffer.  One picture box is a work area, the other contains a
'* copy the the background to refresh with.
'* When you run this demo, you will see that there is no flicker at all.
'* Here's the real trick to get the speed... Do not restore the whole screen
'* when erasing a sprite or copying to the presentation screen... Restore only
'* the area that has changed.  You will be amazed at what you can do with this
'* technique.  You can make it as fast as you need without ASM, DLLs or other
'* controls.  Just pure VB Code (okay, a couple api calls here and there).
'* This method is also versatile.  You can animate a sprite with frames or just
'* bounce one around the screen.  To create the graphics, you need to know
'* how to use a paint program of some type (or have someone do it for you).
'* You can see how I made the simple graphics for this demo by looking at the
'* Workpage Form.
'* The source here is commented as much as I can.  Most experienced VB Programmers
'* can probably pick up on what I'm doing here.  This module is where most of the
'* Code lies.  The only other code is in the frmabout and only contains mouse
'* and keyboard events to shut the program down.
'* If there is more interest in this code, I will write a complete write up on
'* how it is done and include even more examples but for now, you'll have to look
'* at the comments.
'* If you use any of this code, let me know.  I'd like to see what you come up
'* with.  Just drop me an email at dpuckett@thelittleman.com
'* Enjoy!
'********************************************************************************

'** Declaration Section - Types
Type Critter                           'Define Sprite User Defined Object
    FrIndex As Integer                 'Current Frame Index
    Xp As Integer                      'X Position on Display Screen
    Yp As Integer                      'Y Postion on Display Screen
    Width As Integer                   'Width of Critter in pixels
    Height As Integer                  'Height of Critter in pixels
    Xmove As Integer                   'Amount to Move Horizontally
    Ymove As Integer                   'amount to move vertically
    Frames As Integer                  'Amount of Frames in Sprite Set
    Show As Boolean                    'Display or not to display (true=display)
    'Increase array element number for frames amount
    ImageSrcX(3) As Integer            'X Position in Source File Main Image
    ImageSrcY(3) As Integer            'Y Position in Source File Main Graphic
    ImageMaskX(3) As Integer           'X Position in Source File Main Image (Mask)
    ImageMaskY(3) As Integer           'Y Position in Source File Main Graphic (Mask)
End Type

'** Set the number of balls
Private Const intRedBalls = 3
Private Const IntGreenBalls = 2
Private Const intBlueBalls = 1


'** Create # of Critter Sprite Objects
Private Sprites(30) As Critter         'Change Number to amount of Sprites needed

Private SpCnt As Integer               'Counter used for Cylcing through Sprites
Public RightX As Integer              'Right Screen Limit
Public BottomY As Integer
Private SpriteCount As Integer         'Critter Set Count
Public ASpeed As Integer              'Animation Speed

'** API declarations
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDC& Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName$, ByVal lpDeviceName$, ByVal lpOutput$, ByVal lpInitData&)
Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Private Declare Function StretchBlt& Lib "gdi32" (ByVal hDestDC&, ByVal x&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal XSrc&, ByVal YSrc&, ByVal nSrcWidth&, ByVal nSrcHeight&, ByVal dwRop&)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const NOTSRCCOPY = &H330008

'** BitBlt Variables
Private Ind As Long
Private Xo As Long
Private Yo As Long
Private Xs As Long
Private Ys As Long
Private XSrc As Long
Private YSrc As Long
Private DDC As Long
Private SDC As Long
Private res As Long

Sub ShowAboutBox()
      
    ASpeed = 15 '**** this will change the speed of the animation (< faster, > slower)
    glbStopBalls = False
    SpriteCount = intRedBalls + intBlueBalls + IntGreenBalls 'the amount of sprites to show on the screen
    'In this demo, I'm using 10 red balls, 10 blue balls and 10 green balls
    
    '** copying screen Variables
    Dim DestDC As Long               'Destination DC
    Dim XPixels As Long              'Transfer Picture Width
    Dim YPixels As Long              'Trandfer Picture Height
    Dim destX As Long                'Destination X Position
    Dim destY As Long                'Destination Y Position
    Dim srcDC As Long                'Source DC
    Dim SrcX As Long                 'Source X Position
    Dim SrcY As Long                 'Source Y Position
    Dim RasterOp As Long             'Raster Operation to Perform (Copy, And, Or)
    
    frmAbout.Show                      'Bring up the main display form
    BottomY = frmAbout.ScaleHeight     'Set Bottom Screen Limit
    RightX = frmAbout.ScaleWidth       'Set Right Screen Limit
        
    '** Make CleanScreen and WorkSpace Screens are the same size as frmabout Screen
    WorkPage.CleanScreen.Width = frmAbout.Width
    WorkPage.CleanScreen.Height = frmAbout.Height
    WorkPage.WorkScr.Width = frmAbout.Width
    WorkPage.WorkScr.Height = frmAbout.Height
    frmAbout.Refresh  'Make sure frmabout is current
    
    DoEvents
    
    '** Copy frmabout Page to Clean Screen
    '** this saves us a clean background to refresh from when restoring the background undeneath a sprite
    DestDC = WorkPage.CleanScreen.hdc
    destX = 0
    destY = 0
    XPixels = frmAbout.ScaleWidth
    YPixels = frmAbout.ScaleHeight
    srcDC = frmAbout.hdc
    SrcX = 0
    SrcY = 0
    RasterOp& = SRCCOPY
    BitBlt DestDC, destX, destY, XPixels, YPixels, srcDC, SrcX, SrcY, RasterOp
    WorkPage.CleanScreen.Refresh
    
    '** Copy frmabout to WorkSpace Screen
    '** copy the same background picture to the work screen
    DestDC = WorkPage.WorkScr.hdc
    XPixels = frmAbout.ScaleWidth
    YPixels = frmAbout.ScaleHeight
    srcDC = frmAbout.hdc
    SrcX = 0
    SrcY = 0
    RasterOp& = SRCCOPY
    BitBlt DestDC, 0, 0, XPixels, YPixels, srcDC, SrcX, SrcY, RasterOp
    WorkPage.WorkScr.Refresh
    
    'clear these out so they don't take up any space anymore
    DestDC = 0
    XPixels = 0
    YPixels = 0
    srcDC = 0
    SrcX = 0
    SrcY = 0
    RasterOp = 0
    
    Randomize (Timer) 'initialize the random number generator
    
    'the following code initializes my sprites
    Dim z As Long ' generic to count through the sprites
    Dim r As Integer 'Random number from 0 to 5.???
    
    'Setup the initial positions for the sprites (randomly)
    For z = 0 To SpriteCount - 1  'loop through my sprites and setup ramdom start positions
        With Sprites(z)
            .Xp = Int(Rnd * (RightX - 100)) 'set sprites initial position (horizontal)
            .Yp = BottomY - 75 'Set sprites initial position (vertical)
        
            r = Rnd * 6
            If r > 3 Then 'this is used to setup a random left or right movement to start with
                .Xmove = 5
            Else
                .Xmove = -5
            End If
        
            .Ymove = (Rnd * 10) + 4 'amount to move vertically
            .Width = 75      'the width of the sprites used
            .Height = 75     'the height of the sprites used
            .Show = True     'enable the show flag
            .Frames = 1      'not really used in this demo but tells the program how many frames a sprite has
        End With
    Next
            
    'Connect the sprites to an actual picture
    
    'set blue balls
    For z = 0 To intBlueBalls - 1
        With Sprites(z)
            .ImageSrcX(0) = 0 'xy position of the first frame (from the Workpage Master Picturebox)
            .ImageSrcY(0) = 0
            .ImageMaskX(0) = 0   'xy of first frame mask
            .ImageMaskY(0) = 75
        End With
    Next
    
    'set red balls
    For z = intBlueBalls To intBlueBalls + intRedBalls - 1
        With Sprites(z)
            .ImageSrcX(0) = 75 'xy position of the first frame (from the Workpage Master Picturebox)
            .ImageSrcY(0) = 0
            .ImageMaskX(0) = 75   'xy of first frame mask
            .ImageMaskY(0) = 75
        End With
    Next
    
    'Set red balls
    For z = intBlueBalls + intRedBalls To SpriteCount
        With Sprites(z)
            .ImageSrcX(0) = 150 'xy position of the first frame (from the Workpage Master Picturebox)
            .ImageSrcY(0) = 0
            .ImageMaskX(0) = 150   'xy of first frame mask
            .ImageMaskY(0) = 75
        End With
    Next
            
    'end of sprite initialization
    
    Controller  ' Startup the animation (controller routine below)

End Sub
Sub Controller()
    'This is the heart of the system
    'The code below will hit the animation sub, then pause for a determined time,
    'then loop and do it over again until someone shuts it down by clicking a
    'mouse or hitting a key.
    'The animation sub does all of the animation work.  It checks eachs sprites position,
    'advances it and will do just about anything else you want it to but it must release
    'control back to this sub so other functions of the pc can happen.
    'You may also notice that we're not using VB's timer control for the main animation
    'since it is a bit too slow for this purpose.
    Do
        ' controls the speed (normally would be a setting in registry)
        DoAnimation 'do it to it
        DoEvents 'allows system to check and see if key was pressed or mouse was clicked
        If ASpeed > 0 Then Sleep (ASpeed) 'delay to slow down the animation
        
    Loop While glbStopBalls = False

End Sub

Sub DoAnimation()
    
    'Reset Background
    DDC = WorkPage.WorkScr.hdc
    SDC = WorkPage.CleanScreen.hdc
    
    For SpCnt = 0 To SpriteCount ' Erase all Sprites from screen
    With Sprites(SpCnt)
        Xo = .Xp
        Yo = .Yp
        Xs = .Width
        Ys = .Height
        res = BitBlt(DDC, Xo, Yo, Xs, Ys, SDC, Xo, Yo, SRCCOPY)
    End With
    Next
    
    For SpCnt = 0 To SpriteCount 'loop through and move each sprite
    With Sprites(SpCnt)
        .FrIndex = 0 'default standing position
        If .Yp >= BottomY - 50 Then
            .Yp = BottomY - 50
            .Ymove = (Rnd * (BottomY / 10)) + 4
            If .Ymove > 50 Then .Ymove = 50
            .Xp = .Xp - .Xmove
        End If
        If .Xp >= RightX - .Width + 24 Then .Xmove = -((.Xmove * Rnd(5)) + 2)
        If .Xp < -24 Then .Xmove = (.Xmove * Rnd(5) + 2)
        
        .Xp = .Xp + .Xmove
    
        .Yp = .Yp + (.Ymove * -0.5) '-10 if gravity
        .Ymove = .Ymove - 1
    
        'And mask to WorkScr
        Ind = 0 ' .FrIndex
        Xo = .Xp
        Yo = .Yp
        Xs = .Width
        Ys = .Height
        XSrc = .ImageMaskX(Ind)
        YSrc = .ImageMaskY(Ind)
        DDC = WorkPage.WorkScr.hdc
        SDC = WorkPage.Master.hdc
        res = BitBlt(DDC, Xo, Yo, Xs, Ys, SDC, XSrc, YSrc, SRCAND)
    
        'Or image to WorkScr
        XSrc = .ImageSrcX(Ind)
        YSrc = .ImageSrcY(Ind)
        res = BitBlt(DDC, Xo, Yo, Xs, Ys, SDC, XSrc, YSrc, SRCPAINT)
    End With
    Next
    
    'copy all sprites to screen
    For SpCnt = 0 To SpriteCount
    With Sprites(SpCnt)
        Xo = .Xp
        Yo = .Yp
        Xs = .Width
        Ys = .Height
        DDC = frmAbout.hdc
        SDC = WorkPage.WorkScr.hdc
        res = BitBlt(DDC, Xo, Yo, Xs, Ys, SDC, Xo, Yo, SRCCOPY)
    End With
    Next

End Sub


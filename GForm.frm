VERSION 5.00
Begin VB.Form GForm 
   BorderStyle     =   0  'None
   Caption         =   "VB DOOM - by Simon Price - www.VBgames.co.uk"
   ClientHeight    =   5772
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7692
   ForeColor       =   &H0000FF00&
   Icon            =   "GForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "GForm.frx":030A
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame CollisionF 
      Caption         =   "Collision Detection"
      Height          =   972
      Left            =   1440
      TabIndex        =   14
      Top             =   2640
      Width           =   4812
      Begin VB.CheckBox CollisionO 
         Caption         =   "Collision Detection"
         Height          =   192
         Left            =   1560
         TabIndex        =   16
         Top             =   720
         Width           =   1692
      End
      Begin VB.Label Label3 
         Caption         =   "The collision detection at the moment is a bit dodgy, so you can turn this off if you want, allowing you to move more easily."
         Height          =   372
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4572
      End
   End
   Begin VB.Frame LevelF 
      Caption         =   "Level Selection"
      Height          =   1572
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   4812
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 10 (biggest)"
         Height          =   252
         Index           =   10
         Left            =   2640
         TabIndex        =   13
         Top             =   1200
         Value           =   -1  'True
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 9"
         Height          =   252
         Index           =   9
         Left            =   2640
         TabIndex        =   12
         Top             =   960
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 8"
         Height          =   252
         Index           =   8
         Left            =   2640
         TabIndex        =   11
         Top             =   720
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 7"
         Height          =   252
         Index           =   7
         Left            =   2640
         TabIndex        =   10
         Top             =   480
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 6"
         Height          =   252
         Index           =   6
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 5"
         Height          =   252
         Index           =   5
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 4"
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 3"
         Height          =   252
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 2"
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1692
      End
      Begin VB.OptionButton LevelO 
         Caption         =   "Level 1 (smallest)"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1692
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   492
      Left            =   3240
      TabIndex        =   2
      Top             =   3840
      Width           =   1212
   End
   Begin VB.PictureBox LevelPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   240
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox PB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2880
      Left            =   4320
      Picture         =   "GForm.frx":2126A
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   3840
   End
End
Attribute VB_Name = "GForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrameCount As Integer 'I used this to test the fps
Dim StopFlashing As Boolean 'true tells the start message to stop flashing
Dim DoorPos As tCoOrd 'exit co-ords
Dim LevelNo As Byte 'level chosen
Dim Collisions As Boolean 'if collision detection is on or not
Dim Color(1 To 8) As Long

Private Sub cmdStart_Click()
Dim Path As String
'load level map
LevelPic = LoadPicture(App.Path & "\Levels\" & LevelNo & ".bmp")
'make options dissapear
LevelF.Visible = False
CollisionF.Visible = False
cmdStart.Visible = False
'set collsion detection
If CollisionO.Value Then Collisions = True
'load the the level
LoadLevel
'enter main loop
MainLoop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyEscape
  Hide 'exit the game
  MsgBox "Thankyou for playing this early version of VB Doom by Simon Price. There are two other far better perspective texture mapped versions of this game available, from Planet Source Code or my website www.VBgames.co.uk (I have lots of VB games there!).", vbInformation, "Thanks 4 Playing!"
  MsgBox "The game will now attempt to revert back to your original screen resolution, however, it is not always successful so you may have to do this manually.", vbInformation, "Reverting to original screen res"
Select Case iWidth 'put screen res back to normal
Case 800
ChangeScreenSettings 800, 600, 16
Case 1024
ChangeScreenSettings 1024, 768, 16
End Select
    End
  Case vbKeyC 'capture the screen
    PB.Picture = PB.Image
    SavePicture PB.Picture, App.Path & "\Pic.bmp"
End Select
End Sub

Private Sub Form_Load()
Dim Col As tRGBcolor
Dim lpPoint As POINTAPI

CreateStars

'remember some colors
Color(1) = vbCyan
Color(2) = vbYellow
Color(3) = vbBlue
Color(4) = QBColor(6)
Color(5) = QBColor(7)
Color(6) = vbMagenta
Color(7) = vbWhite
Color(8) = vbGreen

Randomize Timer
'remember the screen res b4 messing with it
RememberScreenRes
'change to low-res, 16 bit color
ChangeScreenSettings 640, 480, 16

Show
'build TONS of look-up tables
RememberStuff
'sort out the forms layout
Move 0, 0, RAYSby2 * Screen.TwipsPerPixelX, RAYSby1andHALF * Screen.TwipsPerPixelY
PB.Move 0, 0, RAYSby2, RAYSby1andHALF
LevelNo = 10
End Sub

Public Sub MainLoop()
On Error Resume Next

Dim x As Long
Dim y As Long
Dim Temp As POINTAPI
Dim RayAngle As Single
Dim ScrX As Integer
Dim StepX As Integer
Dim StepY As Integer
Dim Length As Integer
Dim Hit As Long
Dim Height As Integer
Dim LongColor As Long
Dim Shade As tRGBcolor
Dim Zshade As Single

LoadLevel

Do
DoEvents

FrameCount = FrameCount + 1
Caption = FrameCount

PB.Cls
'walk forward
If (GetKeyState(vbKeyUp) And KEY_DOWN) Then
      Man.x = Man.x + Cosine(Man.Angle + ADD90DEGREES) / 10
      Man.y = Man.y - Sine(Man.Angle + ADD90DEGREES) / 10
      If Collisions Then 'check for walls
      If Level.Tile(Man.x, Man.y) <> NOWT Then
        Man.x = Man.x - Cosine(Man.Angle + ADD90DEGREES) / 10
        Man.y = Man.y + Sine(Man.Angle + ADD90DEGREES) / 10
      End If
      End If
End If
'walk backwards
If (GetKeyState(vbKeyDown) And KEY_DOWN) Then
      Man.x = Man.x - Cosine(Man.Angle + ADD90DEGREES) / 10
      Man.y = Man.y + Sine(Man.Angle + ADD90DEGREES) / 10
      If Collisions Then 'check for walls
      If Level.Tile(Man.x, Man.y - 0.5) <> NOWT Then
        Man.x = Man.x + Cosine(Man.Angle + ADD90DEGREES) / 10
        Man.y = Man.y - Sine(Man.Angle + ADD90DEGREES) / 10
      End If
      End If
End If
Const STARSPEED = 15
'turn left
If (GetKeyState(vbKeyLeft) And KEY_DOWN) Then
    If Man.Angle < 0 Then
      Man.Angle = BACKHALFVIEW
    Else
      Man.Angle = Man.Angle - TURNANGLE
    End If
    'rotate stars
    For i = 1 To STARS
      Star(i).x = Star(i).x + STARSPEED
      If Star(i).x > 320 Then
        Star(i).x = Star(i).x - 320
        Star(i).y = Int(Rnd * 120)
      End If
    Next
End If
'turn right
If (GetKeyState(vbKeyRight) And KEY_DOWN) Then
    If Man.Angle > BACKHALFVIEW Then
      Man.Angle = 0
    Else
      Man.Angle = Man.Angle + TURNANGLE
    End If
    'rotate stars
    For i = 1 To STARS
      Star(i).x = Star(i).x - STARSPEED
      If Star(i).x < 0 Then
        Star(i).x = Star(i).x + 320
        Star(i).y = Int(Rnd * 120)
      End If
    Next
End If

'this draws all the stars
PB.ForeColor = RGB(Int(Rnd * 25) + 230, Int(Rnd * 25) + 230, Int(Rnd * 25) + 230)
PB.FillColor = PB.ForeColor
For i = 1 To STARS
  i2 = Int(Rnd * 2) + 1
  PB.Circle (Star(i).x, Star(i).y), i2
  i2 = i2 + i2
  PB.Line (Star(i).x - i2, Star(i).y)-(Star(i).x + i2, Star(i).y)
  PB.Line (Star(i).x, Star(i).y - i2)-(Star(i).x, Star(i).y + i2)
Next

'this sets the first ray 30 degrees to the left of view
RayAngle = Man.Angle - HALFVIEWRAYS

'loop through all 320 rays, drawing a slither of screen each time
For ScrX = 0 To RAYS
  
  x = Man.x * 1200000 'multiply up so that the fixed-point maths is
  y = Man.y * 1200000 'accurate enough
  StepX = Sine(RayAngle) / RAYDETAIL * 1200000 'i.e. 1/10th of a unit
  StepY = Cosine(RayAngle) / RAYDETAIL * 1200000
  Length = 0 'length of ray is reset
  
  Do
    x = x - StepX
    y = y - StepY 'move ray along a bit
    Length = Length + 1 'increment length
    Hit = Level.Tile(x \ 1200000, y \ 1200000) 'GetPixel(LevelPic.hdc, x \ 1200000, y \ 1200000) 'see wot's hit
  Loop Until Hit 'keep the ray going until a hit is detected
  
  Height = Dist2Height(Length) 'and how tall it looks based on ray length
        
    LongColor = Color(Hit)
    Zshade = Length / LIGHT
    Shade.r = LongColor And 255
    Shade.g = (LongColor And 65280) \ 256&
    Shade.b = (LongColor And 16711680) \ 65535
    Shade.b = Shade.b / Zshade
    Shade.g = Shade.g / Zshade
    Shade.r = Shade.r / Zshade
    PB.ForeColor = RGB(Shade.r, Shade.g, Shade.b)
      
    MoveToEx PB.hdc, ScrX, 120 + Height, Temp
    LineTo PB.hdc, ScrX, 120 - Height
      
    'fire next ray 1 pixel further along
    RayAngle = RayAngle + 1
Next
'copy from backbuffer
StretchBlt hdc, 0, 0, RAYSby2, RAYSby1andHALF, PB.hdc, 0, 0, RAYS, RAYSby3div4, vbSrcCopy

Loop

End Sub

Public Sub LoadLevel()
Dim x As Byte
Dim y As Byte

Man.Angle = 0
'loads a level by transferring bitmap into memory
Level.Size = LevelPic.Width - 1
ReDim Level.Tile(0 To Level.Size, 0 To Level.Size)

For x = 0 To Level.Size
For y = 0 To Level.Size
  Select Case GetPixel(GForm.LevelPic.hdc, x, y)
    Case vbBlack
      Level.Tile(x, y) = NOWT
    Case vbCyan
      Level.Tile(x, y) = 1
    Case vbYellow
      Level.Tile(x, y) = 2
    Case vbBlue
      Level.Tile(x, y) = 3
    Case QBColor(6)
      Level.Tile(x, y) = 4
    Case QBColor(7)
      Level.Tile(x, y) = 5
    Case vbMagenta
      Level.Tile(x, y) = 6
    Case vbWhite
      Level.Tile(x, y) = 7
    Case vbGreen
      Level.Tile(x, y) = NOWT
      Man.x = x
      Man.y = y
    Case vbRed
      Level.Tile(x, y) = 8
      DoorPos.x = x
      DoorPos.y = y
  End Select
Next

Next

End Sub


Private Sub Form_Unload(Cancel As Integer)
'change screen res back 2 norm
Select Case iWidth
Case 800
ChangeScreenSettings 800, 600, 16
Case 1024
ChangeScreenSettings 1024, 768, 16
End Select
End Sub

Private Sub LevelO_Click(Index As Integer)
LevelNo = Index
End Sub

Sub CreateStars()
For i = 1 To STARS
  Star(i).x = Int(Rnd * 320)
  Star(i).y = Int(Rnd * 100)
Next
End Sub

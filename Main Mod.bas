Attribute VB_Name = "VB_DOOM_MOD"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Type POINTAPI
  x As Integer
  y As Integer
End Type

Public Type tCoOrd
  x As Byte
  y As Byte
End Type

Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function SetWindowPos Lib "user32" _
  (ByVal hwnd As Long, _
  ByVal hWndInsertAfter As Long, _
  ByVal x As Long, ByVal y As Long, _
  ByVal cx As Long, ByVal cy As Long, _
  ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public lpPoint As POINTAPI

Public Type tMan
  x As Double
  y As Double
  Angle As Integer
End Type

Public Man As tMan

Public Type tLevel
  Tile() As Byte
  Size As Byte
End Type

Public Level As tLevel

Public Type tRGBcolor
  r As Byte
  g As Byte
  b As Byte
End Type

Public i As Integer
Public i2 As Integer
Public i3 As Integer

'try increasing this constant - double it
'and graphics will double in quality!!!
'-ofcourse, the game will run alot slower
'doh! - you can't have your cake AND eat it
Public Const GRAPHICS = 1.85
'how much light there is, try messing around with it
Public Const LIGHT = 50
'how many rays my raycaster casts
Public Const RAYS = 320
Public Const RAYSby2 = RAYS * 2
Public Const RAYSby1andHALF = RAYS * 1.5
'how fast u turn
Public Const TURNANGLE = RAYS \ 12
Public Const PI = 3.1415
'used for converting radians/degrees - well annoying
Public Const PIdiv180 = PI / 180
Public Const RAY_INC = 60 / RAYS
Public Const NUMRAYS = 360 / RAY_INC
Public Const RAYSdiv2 = RAYS \ 2
Public Const HALFVIEWRAYS = NUMRAYS \ 12
Public Const BACKHALFVIEW = NUMRAYS - HALFVIEWRAYS
Public Const RAYSby3div4 = RAYS * 3 \ 4
Public Const RAYSby3div8 = RAYS * 3 \ 8
'a 90 degree turn
Public Const ADD90DEGREES = NUMRAYS \ 4
'height of walls - you could try meddling wiv this 1
Public Const WALLHEIGHT = RAYS * 7 * GRAPHICS
'this affect graphical detail - could be meddled wiv
Public Const RAYDETAIL = 20 * GRAPHICS
'this is the furthest distance the eye can see
'-so don't go off drawing mazes bigger than that
'or the game will crap up
Public Const MAXDIST = 1000 * GRAPHICS

'FAT amounts of look-up tables
'sine, cosine, wall height, shadows, the lot
Public Sine(-HALFVIEWRAYS To NUMRAYS + ADD90DEGREES) As Double
Public Cosine(-HALFVIEWRAYS To NUMRAYS + ADD90DEGREES) As Double
Public Dist2Height(1 To MAXDIST) As Integer

Public Const KEY_TOGGLED As Integer = &H1
Public Const KEY_PRESSED As Integer = &H1000
Public Const KEY_DOWN As Integer = &H1000

Public Const NOWT = 0
'original screen res
Public iHeight As Integer
Public iWidth As Integer

Public Type tStar
  x As Integer
  y As Byte
End Type

Public Const STARS = 30
Public Star(1 To STARS) As tStar

Public Sub RememberStuff()
'this takes a huge amount of memory
'but increases speed so we'll allow it

'work out all sin and cos stuff
Dim Angle As Double

For i = -HALFVIEWRAYS To NUMRAYS + ADD90DEGREES
  Angle = RAY_INC * i * PIdiv180
  Sine(i) = Sin(Angle)
  Cosine(i) = Cos(Angle)
Next

'remember loadsa distances
For i2 = 1 To 5
For i = (i2 - 1) * 100 + 1 To MAXDIST / 5 * i2
  Dist2Height(i) = WALLHEIGHT / i
Next
Next

End Sub



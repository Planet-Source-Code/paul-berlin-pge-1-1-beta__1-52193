Attribute VB_Name = "modPgeGlobal"
'Pab Game Engine Global Varaibles & Functions
'--------------------------------------------
'Created by Paul Berlin 2002-2003
'
'Some of these functions & subs are not being used in this project, but
'they can be useful when making games! =)
'
'Intersect - Use this to check if two sprites are in same place (collision)
'
'ScrollX, ScrollY can be used to scroll the screen.
'
'FrameSkip should be set to true for best performance.
'
Option Explicit

Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'DirectX globals
Public DirectX As New DirectX8
Public Direct3D As Direct3D8
Public Direct3DDevice As Direct3DDevice8
Public Direct3DX As New D3DX8
Public Sprites As D3DXSprite
Public DirectInput As DirectInput8

Public pTex As pgeTexture

Public Target As RECT 'this is set to the size of the render area

Type tWorld
  ScrollX As Single
  ScrollY As Single
  sFadeRed As Single 'These color value goes from 0-100, they will affect
  sFadeGreen As Single 'color values of all sprites & text-fields.
  sFadeBlue As Single 'For example, an value of 50 on red will half all
  sFadeAlpha As Single 'red color values
End Type

Public World As tWorld

Public Const FrameSkip As Boolean = True

Public Const PI = 3.1415926

Public Function tob2(ByVal Val As Single) As Byte
  If Val < 0 Then Val = 0
  If Val > 255 Then Val = 255
  tob2 = CByte(Val)
End Function

Public Function D2R(ByVal degrees As Double) As Double
  'converts degrees to radians
  D2R = degrees * PI / 180
End Function

Public Function R2D(ByVal radians As Double) As Double
  'converts radians to degrees
  R2D = radians * 180 / PI
End Function

Public Function RGBA(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer, ByVal a As Integer) As Long
  'creates an long RGBA color value
  RGBA = D3DColorRGBA(r, g, b, a)
End Function

Public Function RGBA2(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As Long
  'creates an long RGBA color value without alpha value
  RGBA2 = D3DColorRGBA(r, g, b, 255)
End Function

'Function vec3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
'  vec3.x = x
'  vec3.y = y
'  vec3.z = z
'End Function

Public Function vec2(ByVal x As Single, ByVal y As Single) As D3DVECTOR2
  vec2.x = x
  vec2.y = y
End Function

Public Function Sine(ByVal Degrees_Arg As Single) As Single
  Sine = Sin(Degrees_Arg * Atn(1) / 45)
End Function

Public Function Cosine(ByVal Degrees_Arg As Single) As Single
  Cosine = Cos(Degrees_Arg * Atn(1) / 45)
End Function

Public Function ReturnFont(sFont As String, Optional lSize As Integer = 8, Optional bBold As Boolean, Optional bItalic As Boolean, Optional bUnderline As Boolean, Optional bStrikethrough As Boolean) As StdFont
  Set ReturnFont = New StdFont
  With ReturnFont
    .name = sFont
    .Size = lSize
    .Bold = bBold
    .Italic = bItalic
    .Underline = bUnderline
    .Strikethrough = bStrikethrough
  End With
End Function

Public Function ReturnRECT(ByVal x As Long, ByVal y As Long, ByVal x2 As Long, ByVal y2 As Long) As RECT
  With ReturnRECT
    .Left = x
    .Top = y
    .Right = x2
    .bottom = y2
  End With
End Function

Public Function Intersect(Sprite1 As pgeSprite, Sprite2 As pgeSprite) As Long
    Dim tmpRect As RECT
    Intersect = IntersectRect(tmpRect, Sprite1.GetDestRect, Sprite2.GetDestRect)
End Function

Public Function IntersectR(Rect1 As RECT, Rect2 As RECT) As Long
    Dim tmpRect As RECT
    IntersectR = IntersectRect(tmpRect, Rect1, Rect2)
End Function

Public Function GetDist(ByVal X1 As Single, ByVal Y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Single
  'Returns distance between two 2d points
  GetDist = Sqr((X1 - x2) * (X1 - x2) + (Y1 - y2) * (Y1 - y2))
End Function

Function GetAngle(ByVal X1 As Single, ByVal Y1 As Single, ByVal x2 As Single, ByVal y2 As Single) As Single
  'Returns angle between two 2d points
  On Error Resume Next
  Dim Cislo1 As Single
  Dim Cislo2 As Single
  Dim Uhol As Double
  Dim Poloha As Single
  
  If X1 = x2 And Y1 < y2 Then
   Cislo2 = 0
   Poloha = 180
  
  
  ElseIf X1 = x2 And Y1 > y2 Then
   Cislo2 = 0
   Poloha = 0
  ElseIf X1 < x2 And Y1 = y2 Then
   Cislo2 = 0
   Poloha = 90
  ElseIf X1 > x2 And Y1 = y2 Then
   Cislo2 = 0
   Poloha = 270
  ElseIf X1 < x2 And Y1 > y2 Then
   Cislo1 = Abs(x2 - X1)
   Cislo2 = Abs(y2 - Y1)
   Poloha = 0
  ElseIf X1 < x2 And Y1 < y2 Then
   Cislo1 = Abs(Y1 - y2)
   Cislo2 = Abs(x2 - X1)
   Poloha = 90
  ElseIf X1 > x2 And Y1 < y2 Then
   Cislo1 = Abs(X1 - x2)
   Cislo2 = Abs(Y1 - y2)
   Poloha = 180
  ElseIf X1 > x2 And Y1 > y2 Then
   Cislo1 = Abs(y2 - Y1)
   Cislo2 = Abs(X1 - x2)
   Poloha = 270
  End If
  
On Error GoTo Chyba
  Uhol = Atn(Cislo1 / Cislo2) * 57
Chyba:

  GetAngle = Uhol + Poloha
End Function

Public Function RotatePixel(ByVal rot As Single, ByVal speed As Single) As D3DVECTOR2
  RotatePixel.x = speed * Sine(rot)
  RotatePixel.y = speed * Cosine(rot)
End Function

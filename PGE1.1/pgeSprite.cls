VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeSprite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Sprite Class 1.1
'--------------------------------
'Created by Paul Berlin 2002-2003
'
'Used to animate & draw sprites
'
'OVERVIEW:
'
'CHANGELOG:
'1.1 06-06-2003
'* Rewrote most of everything to handle tilesets. It is not so easy to create
'  animated sprites without an tileset now, sadly. Non-animated sprites are easy
'  still though.
Option Explicit

Private Type tRGBA
  bR As Single
  bG As Single
  bB As Single
  bA As Single
End Type

' SPRITE CHANGEABLE PARAMETERS
Private SprCol As tRGBA
Private SprPos As D3DVECTOR2 'Current Position of sprite
Private SprRot As Single 'Current rotation of sprite
Private SprRotC As D3DVECTOR2 'The center of rotation on sprite
Private SprScale As D3DVECTOR2 'Current scale of sprite
Private lSubFrames As Long 'number of subframes

Public active As Boolean 'Is sprite active? Only active sprites are drawn

Public CurrentFrame As Long 'Current frame number
Public CurrentSubFrame As Long 'Current subframe number

'ANIMATION FLAGS
Public bLoop As Boolean
Public bAnimStop As Boolean
Public bCenterScale As Boolean
Public bReverse As Boolean 'Reverse animation
Public bAnimated As Boolean
Public bYMirror As Boolean
Public bXMirror As Boolean
Public bNoGlobalColor As Boolean 'diables use of global color changes on this sprite

'AUTOMATED MOVEMENTS
Private ARotAngle As Single 'Angle to rotate sprite each interval
Private ARotInterval As Long 'Time in ms between each rotation

Private AScale As D3DVECTOR2 'Scale goal
Private AScaleDelta As D3DVECTOR2 'scale change per 10 ms

Private AMove As D3DVECTOR2 'Delta each interval
Private AMoveInterval As Long 'Time in ms between each movement

Private APathGoal As D3DVECTOR2 'Destination coordinates
Private APathDelta As D3DVECTOR2 'Delta each interval

Private AColDelta As tRGBA 'Color delta each interval
Private AColGoal As tRGBA 'Goal color
Private AColDisable As Boolean 'If true, disable sprite when reached fade stop

'FRAME INFO
Private Type tImg
  lTex As Long
  offset As D3DVECTOR2
  Pos As RECT
  vSize As RECT
End Type

Private Type tFrame
  img() As tImg
  lDelay As Long 'Frame delay in milliseconds
End Type
  
Private Frame() As tFrame

'RENDER VALUES
Private DestRect As RECT

Public Sub CreateStatic(ByVal sTexID As String)
  'This creates an sprite that isn't animated, using the whole texture as
  'the sprite. Use CreateStaticEx for more options.
  'sTexID = ID of texture to use
  ReDim Frame(1 To 1)
  With Frame(1)
    ReDim .img(1 To 1)
    With .img(1)
      .lTex = pTex.GetNumFromID(sTexID)
      .Pos = pTex.GetDimension(sTexID)
    End With
  End With
  
  'Setup/reset variables
  Reset
  lSubFrames = 1
  bAnimated = False
End Sub

Public Sub CreateStaticEx(ByVal sTexID As String, ByRef rPos As RECT, ByRef rVSize As RECT)
  'This creates an sprite that isn't animated, using rPos as location on texture
  'and rVSize as virtual size.
  'sTexID = ID of texture to use
  'rPos = position on texture that is the sprite
  'rVSize = virtual size
  
  lSubFrames = 1
  bAnimated = False
  ReDim Frame(1 To 1)
  With Frame(1)
    ReDim .img(1 To 1)
    With .img(1)
      .lTex = pTex.GetNumFromID(sTexID)
      .Pos = rPos
      .vSize = rVSize
    End With
  End With
  
  'Setup/reset variables
  Reset
  lSubFrames = 1
  bAnimated = False
End Sub

Public Sub CreateAnimation(Optional lNumSubframes As Long = 1)
  'This inits values for an animated sprite, or an sprite with more than 1 subframe.
  'Use AddFrame to add frames to the animation and SetSubFrame to setup the
  'subframes of a frame.
  'lNumSubframes = Number of subframes each frame has
  Reset
  If lNumSubframes < 1 Then lNumSubframes = 1
  lSubFrames = lNumSubframes
  bAnimated = True
End Sub

Public Function AddFrame(ByVal lDelay As Long) As Long
  'This adds a frame with the specified properties.
  'lDelay = Delay at this frame in seconds
  'returns frame number, for use with SetSubFrame

  ReDim Preserve Frame(UBound(Frame) + 1)
  With Frame(UBound(Frame))
    ReDim .img(1 To lSubFrames)
    .lDelay = lDelay
  End With
  AddFrame = UBound(Frame)
  
End Function

Public Sub SetSubFrame(ByVal sTexID As String, ByRef rPos As RECT, ByRef rVSize As RECT, ByRef xyOffset As D3DVECTOR2, ByVal lNumSubframe As Long, Optional lNumFrame As Long = 0)
  'This sets up an frame with the specified properties.
  'sTexID = Id of texture to use
  'rPos = Position on texture
  'rVSize = virtual size
  'xyOffset = Offset in x & y angle
  'lNumSubframe = subframe to change
  'lNumFrame = frame to change, 0 will change the last created frame
  
  If lNumFrame <= 0 Then lNumFrame = UBound(Frame)
  If lNumSubframe <= 0 Then lNumSubframe = 1
  With Frame(lNumFrame).img(lNumSubframe)
    .lTex = pTex.GetNumFromID(sTexID)
    .Pos = rPos
    .vSize = rVSize
    .offset = xyOffset
  End With
  
End Sub

Public Sub Render()
  
  If active Then
        
    Dim sVec As D3DVECTOR2
    Dim rCent As D3DVECTOR2
    Dim offX As Single, offY As Single
    Static fTime As Long 'frame Timer
    Static mTime As Long 'move Timer
    Static rTime As Long 'rotate Timer
    Static cTime As Long 'color Timer
    Static sTime As Long 'scale Timer
    Static pTime As Long 'path Timer
    Dim tSng As Single, tLng As Long
    Dim cRGBA As tRGBA
    Dim r As RECT
    
    'make sure timers are on
    If mTime = 0 Then mTime = timeGetTime
    If fTime = 0 Then fTime = timeGetTime
    If rTime = 0 Then rTime = timeGetTime
        
    '#### AUTOMOVE
    If timeGetTime - mTime >= AMoveInterval And AMoveInterval > 0 Then
      If FrameSkip Then 'Get missed movements & add them
        tSng = (timeGetTime - mTime) / AMoveInterval
      Else
        tSng = 1
      End If
      mTime = timeGetTime
      
      With SprPos
        .x = .x + (AMove.x * tSng)
        .y = .y + (AMove.y * tSng)
      End With
    Else
      If AMoveInterval = 0 Then mTime = timeGetTime
    End If
    
    '#### AUTOPATH
    If (APathDelta.x <> 0 Or APathDelta.y <> 0) Then
      If FrameSkip And pTime <> 0 Then 'Get missed moves & add them
        tSng = (timeGetTime - pTime) / 10
      Else
        tSng = 1
      End If
      pTime = timeGetTime
      
      If APathDelta.x <> 0 Then
        With SprPos
          .x = .x + (APathDelta.x * tSng)
          If APathDelta.x < 0 Then
            If .x <= APathGoal.x Then
              .x = APathGoal.x
              APathDelta.x = 0
            End If
          Else
            If .x >= APathGoal.x Then
              .x = APathGoal.x
              APathDelta.x = 0
            End If
          End If
        End With
      End If
      If APathDelta.y <> 0 Then
        With SprPos
          .y = .y + (APathDelta.y * tSng)
          If APathDelta.y < 0 Then
            If .y <= APathGoal.y Then
              .y = APathGoal.y
              APathDelta.y = 0
            End If
          Else
            If .y >= APathGoal.y Then
              .y = APathGoal.y
              APathDelta.y = 0
            End If
          End If
        End With
      End If
      
    Else
      pTime = 0
    End If
    '####
    
    'DESTINIATION RECTANGLE, USED FOR INTERSECTING SPRITES
    With Frame(CurrentFrame).img(CurrentSubFrame)
      If bCenterScale Then
        tLng = ((.Pos.Right - .Pos.Left) / 2)
        DestRect.Left = SprPos.x - (tLng - .vSize.Left) * SprScale.x
        DestRect.Right = SprPos.x + (tLng + .vSize.Right) * SprScale.x
        tLng = ((.Pos.bottom - .Pos.Top) / 2)
        DestRect.Top = SprPos.y - (tLng - .vSize.Top) * SprScale.y
        DestRect.bottom = SprPos.y + (tLng + .vSize.bottom) * SprScale.y
      Else
        DestRect.Left = SprPos.x + .vSize.Left * SprScale.x
        DestRect.Top = SprPos.y + .vSize.Top * SprScale.y
        tLng = (.Pos.Right - .Pos.Left)
        DestRect.Right = SprPos.x + (tLng + .vSize.Right) * SprScale.x
        tLng = (.Pos.bottom - .Pos.Top)
        DestRect.bottom = SprPos.y + (tLng + .vSize.bottom) * SprScale.y
      End If
    End With
        
    '#### AUTOSCALING
    If (AScaleDelta.x <> 0 Or AScaleDelta.y <> 0) Then
      If FrameSkip And sTime <> 0 Then 'Get missed scales & add them
        tSng = (timeGetTime - sTime) / 10
      Else
        tSng = 1
      End If
      sTime = timeGetTime
          
      If AScaleDelta.x <> 0 Then
        With SprScale
          .x = .x + (AScaleDelta.x * tSng)
          If AScaleDelta.x < 0 Then
            If .x <= AScale.x Then
              .x = AScale.x
              AScaleDelta.x = 0
            End If
          Else
            If .x >= AScale.x Then
              .x = AScale.x
              AScaleDelta.x = 0
            End If
          End If
        End With
      End If
      If AScaleDelta.y <> 0 Then
        With SprScale
          .y = .y + (AScaleDelta.y * tSng)
          If AScaleDelta.y < 0 Then
            If .y <= AScale.y Then
              .y = AScale.y
              AScaleDelta.y = 0
            End If
          Else
            If .y >= AScale.y Then
              .y = AScale.y
              AScaleDelta.y = 0
            End If
          End If
        End With
      End If
      
    Else
      sTime = 0
    End If
    '####
    
    '#### AUTOROTATE
    If timeGetTime - rTime >= ARotInterval And ARotInterval > 0 Then
      If FrameSkip Then 'Get missed rotations & add them
        tSng = (timeGetTime - rTime) / ARotInterval
      Else
        tSng = 1
      End If
      rTime = timeGetTime
        
      SprRot = SprRot + (ARotAngle * tSng)
    Else
      If ARotInterval = 0 Then rTime = timeGetTime
    End If
    
    'increase, decrease frame if specified
    If timeGetTime - fTime >= Frame(CurrentFrame).lDelay And Frame(CurrentFrame).lDelay > 0 And bAnimated Then
              
      If FrameSkip Then 'Skip frame if double time or more has elapsed
        tLng = (timeGetTime - fTime) / Frame(CurrentFrame).lDelay
      Else
        tLng = 1
      End If
      fTime = timeGetTime
              
      If bReverse Then
        CurrentFrame = CurrentFrame - (1 * tLng)
      Else
        CurrentFrame = CurrentFrame + (1 * tLng)
      End If
              
      'bounds of animation:
      If CurrentFrame > UBound(Frame) Then
        If bAnimStop Then
          active = False
          Exit Sub
        End If
        If bLoop Then
          CurrentFrame = 1
        Else
          CurrentFrame = UBound(Frame)
        End If
      End If
      'it can go reverse too...
      If CurrentFrame < 1 Then
        If bLoop Then
          CurrentFrame = UBound(Frame)
        Else
          CurrentFrame = 1
        End If
      End If
      
    End If

    '#### MIRROR SPRITE
    sVec = SprScale
    With Frame(CurrentFrame).img(CurrentSubFrame)
      offX = .offset.x
      offY = .offset.y
      If bXMirror Then
        sVec.x = -sVec.x
        offX = offX + (.Pos.Right - .Pos.Left)
      End If
      If bYMirror Then
        sVec.y = -sVec.y
        offY = offY + (.Pos.bottom - .Pos.Top)
      End If
    End With
  
    '#### AUTOFADE
    If (AColDelta.bR <> 0 Or AColDelta.bG <> 0 Or AColDelta.bB <> 0 Or AColDelta.bA <> 0) Then
      If FrameSkip And cTime <> 0 Then  'Get missed color changes & add them
        tSng = (timeGetTime - cTime) / 10
      Else
        tSng = 1
      End If
      cTime = timeGetTime
      
      If AColDelta.bR <> 0 Then
        SprCol.bR = SprCol.bR + (AColDelta.bR * tSng)
        If AColDelta.bR < 0 Then
          If SprCol.bR <= AColGoal.bR Then
            SprCol.bR = AColGoal.bR
            AColDelta.bR = 0
          End If
        Else
          If SprCol.bR >= AColGoal.bR Then
            SprCol.bR = AColGoal.bR
            AColDelta.bR = 0
          End If
        End If
      End If
      If AColDelta.bG <> 0 Then
        SprCol.bG = SprCol.bG + (AColDelta.bG * tSng)
        If AColDelta.bG < 0 Then
          If SprCol.bG <= AColGoal.bG Then
            SprCol.bG = AColGoal.bG
            AColDelta.bG = 0
          End If
        Else
          If SprCol.bG >= AColGoal.bG Then
            SprCol.bG = AColGoal.bG
            AColDelta.bG = 0
          End If
        End If
      End If
      If AColDelta.bB <> 0 Then
        SprCol.bB = SprCol.bB + (AColDelta.bB * tSng)
        If AColDelta.bB < 0 Then
          If SprCol.bB <= AColGoal.bB Then
            SprCol.bB = AColGoal.bB
            AColDelta.bB = 0
          End If
        Else
          If SprCol.bB >= AColGoal.bB Then
            SprCol.bB = AColGoal.bB
            AColDelta.bB = 0
          End If
        End If
      End If
      If AColDelta.bA <> 0 Then
        SprCol.bA = SprCol.bA + (AColDelta.bA * tSng)
        If AColDelta.bA < 0 Then
          If SprCol.bA <= AColGoal.bA Then
            SprCol.bA = AColGoal.bA
            AColDelta.bA = 0
          End If
        Else
          If SprCol.bA >= AColGoal.bA Then
            SprCol.bA = AColGoal.bA
            AColDelta.bA = 0
          End If
        End If
      End If
      
      If SprCol.bA = AColGoal.bA And SprCol.bR = AColGoal.bR And SprCol.bG = AColGoal.bG And SprCol.bB = AColGoal.bB And AColDisable Then
        AColDisable = False
        active = False
      End If
    Else
      cTime = 0
    End If

    'Calculate scaled rotation center
    If bCenterScale Then
      With Frame(CurrentFrame).img(CurrentSubFrame)
        rCent.x = (((.Pos.Right - .Pos.Left) / 2) + SprRotC.x) * SprScale.x
        rCent.y = (((.Pos.bottom - .Pos.Top) / 2) + SprRotC.y) * SprScale.y
      End With
    Else
      rCent.x = SprRotC.x * SprScale.x
      rCent.y = SprRotC.y * SprScale.y
    End If
    
    cRGBA = SprCol
    If Not bNoGlobalColor Then
      cRGBA.bR = cRGBA.bR * (World.sFadeRed / 100)
      cRGBA.bG = cRGBA.bG * (World.sFadeGreen / 100)
      cRGBA.bB = cRGBA.bB * (World.sFadeBlue / 100)
      cRGBA.bA = cRGBA.bA * (World.sFadeAlpha / 100)
    End If
    
    With Frame(CurrentFrame).img(CurrentSubFrame)
      If bCenterScale Then
        Sprites.Draw pTex.GetTexture(.lTex), _
        .Pos, sVec, rCent, SprRot, _
        vec2(SprPos.x + offX + World.ScrollX - ((.Pos.Right - .Pos.Left) * SprScale.x) / 2, _
        SprPos.y + offY + World.ScrollY - ((.Pos.bottom - .Pos.Top) * SprScale.y) / 2), _
        RGBA(cRGBA.bR, cRGBA.bG, cRGBA.bB, cRGBA.bA)
      Else
        Sprites.Draw pTex.GetTexture(.lTex), _
        .Pos, sVec, rCent, SprRot, _
        vec2(SprPos.x + offX + World.ScrollX, _
        SprPos.y + offY + World.ScrollY), _
        RGBA(cRGBA.bR, cRGBA.bG, cRGBA.bB, cRGBA.bA)
      End If
    End With
  End If
    
End Sub

Private Sub Class_Initialize()
  ReDim Frame(0)
End Sub

Public Sub SetAutoMovement(ByVal mX As Single, ByVal mY As Single, ByVal IntervalMs As Long)
  'This will move sprite mX & mY pixels each interval
  'mX, mY = change to sprite coordinates each interval
  'IntervalMs = Interval in milliseconds
  AMove.x = mX
  AMove.y = mY
  AMoveInterval = IntervalMs
End Sub

Public Function GetAutoMovement() As D3DVECTOR2
  GetAutoMovement.x = AMove.x
  GetAutoMovement.y = AMove.y
End Function

Public Sub SetAutoRotation(ByVal RotDegree As Single, ByVal IntervalMs As Long)
  'This will rotate the sprite by RotDegree degrees each interval.
  'RotDegree = change to sprite rotation each interval
  'IntervalMs = Interval in milliseconds
  ARotAngle = D2R(RotDegree)
  ARotInterval = IntervalMs
End Sub

Public Function GetAutoRotation() As Single
  GetAutoRotation = ARotAngle
End Function

Public Sub SetAutoFade(ldestR As Integer, ldestG As Integer, ldestB As Integer, ldestA As Integer, TotalMs As Long, Optional bDisable As Boolean = False)
  'This will fade the color of the sprite from it's current value to the destination
  'values within the specified time. If bDisable = True then the sprite will be
  'disabled (sets Active=False) when it reaches it's fade goal.
  'If any color value is to be remained unchanged, enter an unvalid value (<0, >255)
  'ldestR-A = Color value
  'TotalMs = Time it takes for the fade to reach it's goal
  'bDisable = disables sprite when done if true
  If ldestR >= 0 And ldestR <= 255 Then
    AColGoal.bR = ldestR
    AColDelta.bR = (ldestR - SprCol.bR) / (TotalMs / 10)
  Else
    AColGoal.bR = SprCol.bR
  End If
  If ldestG >= 0 And ldestG <= 255 Then
    AColGoal.bG = ldestG
    AColDelta.bG = (ldestG - SprCol.bG) / (TotalMs / 10)
  Else
    AColGoal.bG = SprCol.bG
  End If
  If ldestB >= 0 And ldestB <= 255 Then
    AColGoal.bB = ldestB
    AColDelta.bB = (ldestB - SprCol.bB) / (TotalMs / 10)
  Else
    AColGoal.bB = SprCol.bB
  End If
  If ldestA >= 0 And ldestA <= 255 Then
    AColGoal.bA = ldestA
    AColDelta.bA = (ldestA - SprCol.bA) / (TotalMs / 10)
  Else
    AColGoal.bA = SprCol.bA
  End If
  AColDisable = bDisable
End Sub

Public Sub SetAutoScale(destX As Single, destY As Single, TotalMs As Long)
  'This will scale the sprite by to the specified scale within the specified time.
  'destX-Y = X & Y scale (1.0=normal,2.0=double size, etc.)
  'TotalMs = Time it takes for the scale to reach it's goal
  AScale.x = destX
  AScale.y = destY
  AScaleDelta.x = (AScale.x - SprScale.x) / (TotalMs / 10)
  AScaleDelta.y = (AScale.y - SprScale.y) / (TotalMs / 10)
End Sub

Public Sub SetScale(x As Single, y As Single)
  'Sets the sprite scale
  'X, Y = X & Y scale (1.0=normal,2.0=double size, etc.)
  SprScale.x = x
  SprScale.y = y
End Sub

Public Function GetScale() As D3DVECTOR2
  GetScale.x = SprScale.x
  GetScale.y = SprScale.y
End Function

Public Sub SetCenterRot(x As Single, y As Single)
  'Sets the center of rotation on the sprite. This is the pixel of where the
  'roation axel is located.
  SprRotC.x = x: SprRotC.y = y
End Sub

Public Function GetCenterRot() As D3DVECTOR2
  GetCenterRot.x = SprRotC.x
  GetCenterRot.y = SprRotC.y
End Function

Public Sub SetColor(r As Byte, g As Byte, b As Byte, a As Byte)
  'Sets the sprite color
  'a = alpha
  SprCol.bA = a
  SprCol.bR = r
  SprCol.bG = g
  SprCol.bB = b
End Sub

Public Sub SetRotation(RotDegree As Single)
  'Sets the sprite rotation
  SprRot = D2R(RotDegree)
End Sub

Public Function GetRotationD() As Single
  GetRotationD = R2D(SprRot)
End Function

Public Sub SetPosition(x As Single, y As Single)
  'Sets the sprite position on screen
  SprPos.x = x
  SprPos.y = y
End Sub

Public Function GetPosition() As D3DVECTOR2
  GetPosition = SprPos
End Function

Public Function GetDestRect() As RECT
  GetDestRect = DestRect
End Function

Public Sub SetAutoPathTime(destX As Single, destY As Single, TotalMs As Long)
  'Sets an destination to move the sprite to within the specified time.
  APathGoal.x = destX
  APathGoal.y = destY
  APathDelta.x = (APathGoal.x - SprPos.x) / (TotalMs / 10)
  APathDelta.y = (APathGoal.y - SprPos.y) / (TotalMs / 10)
End Sub

'Public Sub SetAutoPath(destX As Single, destY As Single, Speed As Single)
'  'Sets an destination to move the sprite to within the specified time.
'  APathGoal.x = destX
'  APathGoal.y = destY
'  APathDelta.x = (APathGoal.x - SprPos.x) / ((APathGoal.x - SprPos.x) * (Speed / 10))
'  APathDelta.y = (APathGoal.y - SprPos.y) / ((APathGoal.y - SprPos.y) * (Speed / 10))
'End Sub

Private Sub Reset()
  'This just resets some value to it's defaults
  AScaleDelta.x = 0: AScaleDelta.y = 0
  APathDelta.x = 0: APathDelta.y = 0
  AColDelta.bR = 0: AColDelta.bG = 0: AColDelta.bB = 0: AColDelta.bA = 0
  SprPos.x = 0: SprPos.y = 0
  SprRotC.x = 0: SprRotC.y = 0
  AMoveInterval = 0
  ARotInterval = 0
  active = True
  bCenterScale = True
  bLoop = True
  
  CurrentFrame = 1
  CurrentSubFrame = 1
  
  SprRot = 0
  SprCol.bR = 255: SprCol.bG = 255: SprCol.bB = 255: SprCol.bA = 255
  SprScale.x = 1: SprScale.y = 1
End Sub

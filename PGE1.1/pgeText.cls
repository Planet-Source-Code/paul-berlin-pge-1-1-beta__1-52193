VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Text Class 1.0
'------------------------------
'Created by Paul Berlin 2002-2003
'
'Used to Animate & Draw Text
'
'
'OVERVIEW:
'* Use Create(font) to initialize pgeText with an already created pgeFont.
'* Use SetColor, SetSize, SetPosition, sCaption to set the text properties.
'* Use SetAutoFade(ldestR, ldestG, ldestB, ldestA, TotalMs, bDisable) to fade the
'  text from it's current color to a new one.
'* Use SetAutoMovement(mX, mY, IntervalMs) to move the text the specified pixels
'  each interval.
'* Use SetAutoPath(destX, destY, TotalMs) to move the text from its current position
'  to an other screen position during the specified time.
'* Use SetAutoSize(destW, destH, TotalMs) to change the size of the text-box from
'  it's current size to the specified size during the specified time.
'
'CHANGELOG:
'
Option Explicit

'Used font
Private pFont As pgeFont

Private Type tRGBA
  bR As Single
  bG As Single
  bB As Single
  bA As Single
End Type

'Public settings
Public bNoGlobalColor As Boolean 'Disables global color changes on this text
Public bCenterPos As Boolean    'When true will center text at it's coordinates.
Public active As Boolean        'If this is false, the text is not rendered.
Public sCaption As String       'Text caption
Public Format As CONST_DTFLAGS  'Text format

'Current values
Private TxtCol As tRGBA 'Current color
Private TxtPos As D3DVECTOR2
Private TxtSize As D3DVECTOR2 'x=width,y=height

'Automated Color Fade
'--------------------
Private AColDelta As tRGBA 'Color delta each interval
Private AColGoal As tRGBA 'Color goal
Private AColDisable As Boolean 'If true, disable sprite when reached fade stop

'Path Movement
'-------------
Private APathGoal As D3DVECTOR2 'Destination coordinates
Private APathDelta As D3DVECTOR2 'Delta each interval

'Automated Movement
'------------------
Private AMove As D3DVECTOR2 'Delta each interval
Private AMoveInterval As Long 'Time in ms between each movement

'Size Movement
'-------------
Private ASizeGoal As D3DVECTOR2 'Destination coordinates
Private ASizeDelta As D3DVECTOR2 'Delta each interval

Public Sub Create(ByRef font As pgeFont)
  Set pFont = font
End Sub

Public Sub SetPosition(x As Single, y As Single)
  TxtPos.x = x
  TxtPos.y = y
End Sub

Public Sub SetSize(w As Long, h As Long)
  TxtSize.x = w
  TxtSize.y = h
End Sub

Public Sub SetColor(r As Byte, g As Byte, b As Byte, A As Byte)
  TxtCol.bR = r
  TxtCol.bG = g
  TxtCol.bB = b
  TxtCol.bA = A
End Sub

Private Sub Class_Initialize()
  active = True
  TxtSize.x = 100
  TxtSize.y = 25
  TxtCol.bA = 255
  TxtCol.bR = 255
  TxtCol.bG = 255
  TxtCol.bB = 255
  Format = DT_LEFT
End Sub

Public Sub Render()
  If active Then
    Dim tSng As Single, cRGBA As tRGBA
    Static cTime As Long 'color timer
    Static mTime As Long 'automove timer
    Static pTime As Long 'autopath timer
    Static sTime As Long 'autosize timer
    
    'make sure timers are on
    If mTime = 0 Then mTime = timeGetTime
    
    '#### AUTOMOVE
    If timeGetTime - mTime >= AMoveInterval And AMoveInterval > 0 Then
      If FrameSkip Then 'Get missed movements & add them
        tSng = (timeGetTime - mTime) / AMoveInterval
      Else
        tSng = 1
      End If
      mTime = timeGetTime
      
      With TxtPos
        .x = .x + (AMove.x * tSng)
        .y = .y + (AMove.y * tSng)
      End With
    Else
      If AMoveInterval = 0 Then mTime = timeGetTime
    End If
  
    '#### AUTOPATH
    If (APathDelta.x <> 0 Or APathDelta.y <> 0) Then
      If FrameSkip And pTime <> 0 Then 'Get missed scales & add them
        tSng = (timeGetTime - pTime) / 10
      Else
        tSng = 1
      End If
      pTime = timeGetTime
      
      If APathDelta.x <> 0 Then
        With TxtPos
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
        With TxtPos
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
    
    '#### AUTOSIZE
    If (ASizeDelta.x <> 0 Or ASizeDelta.y <> 0) Then
      If FrameSkip And sTime <> 0 Then 'Get missed scales & add them
        tSng = (timeGetTime - sTime) / 10
      Else
        tSng = 1
      End If
      sTime = timeGetTime
      
      If ASizeDelta.x <> 0 Then
        With TxtSize
          .x = .x + (ASizeDelta.x * tSng)
          If ASizeDelta.x < 0 Then
            If .x <= ASizeGoal.x Then
              .x = ASizeGoal.x
              ASizeDelta.x = 0
            End If
          Else
            If .x >= ASizeGoal.x Then
              .x = ASizeGoal.x
              ASizeDelta.x = 0
            End If
          End If
        End With
      End If
      If ASizeDelta.y <> 0 Then
        With TxtSize
          .y = .y + (ASizeDelta.y * tSng)
          If ASizeDelta.y < 0 Then
            If .y <= ASizeGoal.y Then
              .y = ASizeGoal.y
              ASizeDelta.y = 0
            End If
          Else
            If .y >= ASizeGoal.y Then
              .y = ASizeGoal.y
              ASizeDelta.y = 0
            End If
          End If
        End With
      End If
      
    Else
      sTime = 0
    End If
  
    '#### AUTOFADE
    If (AColDelta.bR <> 0 Or AColDelta.bG <> 0 Or AColDelta.bB <> 0 Or AColDelta.bA <> 0) Then
      If FrameSkip And cTime <> 0 Then  'Get missed color changes & add them
        tSng = (timeGetTime - cTime) / 10
      Else
        tSng = 1
      End If
      cTime = timeGetTime
          
      If AColDelta.bR <> 0 Then
        TxtCol.bR = TxtCol.bR + (AColDelta.bR * tSng)
        If AColDelta.bR < 0 Then
          If TxtCol.bR <= AColGoal.bR Then
            TxtCol.bR = AColGoal.bR
            AColDelta.bR = 0
          End If
        Else
          If TxtCol.bR >= AColGoal.bR Then
            TxtCol.bR = AColGoal.bR
            AColDelta.bR = 0
          End If
        End If
      End If
      If AColDelta.bG <> 0 Then
        TxtCol.bG = TxtCol.bG + (AColDelta.bG * tSng)
        If AColDelta.bG < 0 Then
          If TxtCol.bG <= AColGoal.bG Then
            TxtCol.bG = AColGoal.bG
            AColDelta.bG = 0
          End If
        Else
          If TxtCol.bG >= AColGoal.bG Then
            TxtCol.bG = AColGoal.bG
            AColDelta.bG = 0
          End If
        End If
      End If
      If AColDelta.bB <> 0 Then
        TxtCol.bB = TxtCol.bB + (AColDelta.bB * tSng)
        If AColDelta.bB < 0 Then
          If TxtCol.bB <= AColGoal.bB Then
            TxtCol.bB = AColGoal.bB
            AColDelta.bB = 0
          End If
        Else
          If TxtCol.bB >= AColGoal.bB Then
            TxtCol.bB = AColGoal.bB
            AColDelta.bB = 0
          End If
        End If
      End If
      If AColDelta.bA <> 0 Then
        TxtCol.bA = TxtCol.bA + (AColDelta.bA * tSng)
        If AColDelta.bA < 0 Then
          If TxtCol.bA <= AColGoal.bA Then
            TxtCol.bA = AColGoal.bA
            AColDelta.bA = 0
          End If
        Else
          If TxtCol.bA >= AColGoal.bA Then
            TxtCol.bA = AColGoal.bA
            AColDelta.bA = 0
          End If
        End If
      End If
      
      If TxtCol.bA = AColGoal.bA And TxtCol.bR = AColGoal.bR And TxtCol.bG = AColGoal.bG And TxtCol.bB = AColGoal.bB And AColDisable Then
        AColDisable = False
        active = False
      End If
    Else
      cTime = 0
    End If
    
    cRGBA = TxtCol
    If Not bNoGlobalColor Then
      cRGBA.bR = cRGBA.bR * (World.sFadeRed / 100)
      cRGBA.bG = cRGBA.bG * (World.sFadeGreen / 100)
      cRGBA.bB = cRGBA.bB * (World.sFadeBlue / 100)
      cRGBA.bA = cRGBA.bA * (World.sFadeAlpha / 100)
    End If
    
    If bCenterPos Then
      pFont.DrawText sCaption, ReturnRECT(TxtPos.x - (TxtSize.x / 2), TxtPos.y - (TxtSize.y / 2), TxtPos.x + (TxtSize.x / 2), TxtPos.y + (TxtSize.y / 2)), RGBA(cRGBA.bR, cRGBA.bG, cRGBA.bB, cRGBA.bA), Format
    Else
      pFont.DrawText sCaption, ReturnRECT(TxtPos.x, TxtPos.y, TxtPos.x + TxtSize.x, TxtPos.y + TxtSize.y), RGBA(cRGBA.bR, cRGBA.bG, cRGBA.bB, cRGBA.bA), Format
    End If
  End If
End Sub

Public Function GetPosition() As D3DVECTOR2
  GetPosition.x = TxtPos.x
  GetPosition.y = TxtPos.y
End Function

Public Function GetSize() As D3DVECTOR2
  GetSize.x = TxtSize.x
  GetSize.y = TxtSize.y
End Function

Public Sub SetAutoFade(ldestR As Integer, ldestG As Integer, ldestB As Integer, ldestA As Integer, TotalMs As Long, Optional bDisable As Boolean = False)
  'If any color value is to be remained unchanged, enter an unvalid value (<0, >255)
  If ldestR >= 0 And ldestR <= 255 Then
    AColGoal.bR = ldestR
    AColDelta.bR = (ldestR - TxtCol.bR) / (TotalMs / 10)
  Else
    AColGoal.bR = TxtCol.bR
  End If
  If ldestG >= 0 And ldestG <= 255 Then
    AColGoal.bG = ldestG
    AColDelta.bG = (ldestG - TxtCol.bG) / (TotalMs / 10)
  Else
    AColGoal.bG = TxtCol.bG
  End If
  If ldestB >= 0 And ldestB <= 255 Then
    AColGoal.bB = ldestB
    AColDelta.bB = (ldestB - TxtCol.bB) / (TotalMs / 10)
  Else
    AColGoal.bB = TxtCol.bB
  End If
  If ldestA >= 0 And ldestA <= 255 Then
    AColGoal.bA = ldestA
    AColDelta.bA = (ldestA - TxtCol.bA) / (TotalMs / 10)
  Else
    AColGoal.bA = TxtCol.bA
  End If
  AColDisable = bDisable
End Sub

Public Sub SetAutoMovement(ByVal mX As Single, ByVal mY As Single, ByVal IntervalMs As Long)
  AMove.x = mX
  AMove.y = mY
  AMoveInterval = IntervalMs
End Sub

Public Sub SetAutoPath(destX As Single, destY As Single, TotalMs As Long)
  APathGoal.x = destX
  APathGoal.y = destY
  APathDelta.x = (APathGoal.x - TxtPos.x) / (TotalMs / 10)
  APathDelta.y = (APathGoal.y - TxtPos.y) / (TotalMs / 10)
End Sub

Public Sub SetAutoSize(destW As Single, destH As Single, TotalMs As Long)
  ASizeGoal.x = destW
  ASizeGoal.y = destH
  ASizeDelta.x = (ASizeGoal.x - TxtSize.x) / (TotalMs / 10)
  ASizeDelta.y = (ASizeGoal.y - TxtSize.y) / (TotalMs / 10)
End Sub

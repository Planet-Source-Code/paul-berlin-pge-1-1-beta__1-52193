VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeTileset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Tileset Loader Class 1.0
'----------------------------------------
'Created by Paul Berlin 2003
'
'Used to load tilesets and create sprites from the tileset.
'
'OVERVIEW:
'* Use LoadTiles(sFile) to load your compiled tileset (*.pbt).
'* Use Create(spr, sID) to set sprite (spr) to loaded animation (sID).
'
'CHANGELOG:
'
Option Explicit

Private Const ComVer As Byte = 1

Private Enum enumOptions
  VirtualSize_1 = 1             'Virtual frame size is included
  SubImages_2 = 2               'Animation have subimages
  Loop_4 = 4                    'Default animation to loop on
  Center_8 = 8                  'Default animation to center oriented on
End Enum

Private Type tResDest
  lRes As Long                'Pointer to resource containing frame data
  Pos As RECT                 'Location on resource image of frame data
  VirtualSize As RECT         'Size of frame to be used in intersecting
  offset As D3DVECTOR2
End Type

Private Type tFrameControlInfo
  lDelay As Long              'The delay for this frame
End Type

Private Type tFrame
  img() As tResDest           'The image data for each of this frame's subimages
  Ctrl As tFrameControlInfo   'This frame's control info
End Type

Private Type tAnim
  sID As String               'ID of this animation
  eOptions As enumOptions     'Options for this animation
  lSubimages As Long          'Number of subimages this animation has
  Frame() As tFrame           'the frames of this animation
End Type

Private Ani() As tAnim

Private lPosition     As Long   'File Position
Private Free          As Integer 'File number

Public Function LoadTiles(ByVal sFile As String) As Boolean
  'This reads an tileset into memory
  'sFile = location of your tileset file
  On Error GoTo errh
  Dim x As Long, y As Long, z As Long, b() As Byte
  Dim F As New pgeFile
  LoadTiles = True
  
  F.filename = sFile
  
  sFile = F.ReadStrFixed(5)
  If Left(sFile, 4) <> "PGEC" Or Asc(Right(sFile, 1)) <> ComVer Then GoTo errh
  F.SkipField 2 'skip id & descrption (it's just there to show you if you h4xx0r the fajl)
  For x = 1 To F.ReadNumber
    F.ReadFileMem b()
    pTextures.LoadMemTexture b(), "tileres" & x, F.ReadNumber
  Next
  ReDim b(0)
  For x = 1 To F.ReadNumber
    ReDim Preserve Ani(UBound(Ani) + 1)
    With Ani(UBound(Ani))
      .sID = F.ReadStr
      .eOptions = F.ReadNumber
      If (.eOptions And SubImages_2) = SubImages_2 Then
        .lSubimages = F.ReadNumber
      Else
        .lSubimages = 1
      End If
      ReDim .Frame(F.ReadNumber)
      For y = 1 To UBound(.Frame)
        With .Frame(y)
          ReDim .img(Ani(x).lSubimages)
          .Ctrl.lDelay = F.ReadNumber
          For z = 1 To UBound(.img)
            With .img(z)
              .lRes = F.ReadNumber
              .offset.x = F.ReadNumber
              .offset.y = F.ReadNumber
              .Pos.Left = F.ReadNumber
              .Pos.Right = F.ReadNumber
              .Pos.Top = F.ReadNumber
              .Pos.bottom = F.ReadNumber
              If (Ani(x).eOptions And VirtualSize_1) = VirtualSize_1 Then
                .VirtualSize.Left = F.ReadNumber
                .VirtualSize.Right = F.ReadNumber
                .VirtualSize.Top = F.ReadNumber
                .VirtualSize.bottom = F.ReadNumber
              End If
            End With
          Next
        End With
      Next
    End With
  Next
  
  'Debug.Print "Loaded tileset:"; UBound(Ani); "total tiles"

  Exit Function
errh:
  LoadTiles = False
End Function

Private Sub Class_Initialize()
  ReDim Ani(0)
End Sub

Public Sub Create(ByRef spr As pgeSprite, ByVal sID As String)
  'sets up an sprite according to loaded sprite.
  'spr = pgeSprite class
  'sID = ID of sprite/animation/tile to set.
  Dim num As Long, x As Long, y As Long
  
  num = GetNumFromID(sID)
  If num > 0 Then
    With Ani(num)
      If UBound(.Frame) > 1 Or .lSubimages > 1 Then 'animated
        spr.CreateAnimation .lSubimages
        For x = 1 To UBound(.Frame)
          spr.AddFrame .Frame(x).Ctrl.lDelay
          For y = 1 To .lSubimages
            With .Frame(x).img(y)
              spr.SetSubFrame "tileres" & .lRes, .Pos, .VirtualSize, .offset, y
            End With
          Next
        Next
        If Not (.eOptions And Center_8) = Center_8 Then spr.bCenterScale = False
        If Not (.eOptions And Loop_4) = Loop_4 Then spr.bLoop = False
      Else 'static
        With .Frame(1).img(1)
          spr.CreateStatic "tileres" & .lRes
          spr.SetSubFrame "tileres" & .lRes, .Pos, .VirtualSize, .offset, 1, 1
        End With
        If Not (.eOptions And Center_8) = Center_8 Then spr.bCenterScale = False
      End If
    End With
  End If
End Sub

Private Function GetNumFromID(ByVal sID As String) As Long
  Dim x As Long
  For x = 1 To UBound(Ani)
    If Ani(x).sID = sID Then
      GetNumFromID = x
      Exit Function
    End If
  Next
End Function


'##############################################################
'The below functions are used when reading the compiled tileset
'##############################################################

Private Function ReadStr() As String
  Dim vSize As Byte   'The size of the short string
  Dim vByte(3) As Byte 'The size of the long string

  Get Free, lPosition, vSize: lPosition = lPosition + 1
  If vSize = 103 Then 'It is an short string
  
    Get Free, lPosition, vSize: lPosition = lPosition + 1 'Get size
    ReadStr = Space(vSize) 'Setup the size of the string to read into
    Get Free, lPosition, ReadStr: lPosition = lPosition + vSize

  ElseIf vSize = 104 Then 'It is an long string
  
    Get Free, lPosition, vByte: lPosition = lPosition + 4 'get size
    ReadStr = Space(Byte2Long(vByte)) 'Setup the size of the string to read into
    Get Free, lPosition, ReadStr: lPosition = lPosition + Byte2Long(vByte)
    
  ElseIf vSize = 200 Then 'The string is 0 length, skip
    ReadStr = ""
  
  Else 'This aint no string
    Err.Raise 566
  End If

End Function

Private Function ReadStrFixed(ByVal vSize As Long) As String
  ReadStrFixed = Space(vSize) 'Setup the size of the string to read into
  Get Free, lPosition, ReadStrFixed: lPosition = lPosition + vSize
End Function

Private Function Long2ByteStr(ByVal Val As Long) As String
  Dim x As Integer

  ReDim byteArray(3) As Byte
  For x = 0 To 3
    byteArray(x) = Val \ (2 ^ ((3 - x) * 7)) And 127
  Next x
  
  Long2ByteStr = Chr(byteArray(0)) & Chr(byteArray(1)) & Chr(byteArray(2)) & Chr(byteArray(3))
End Function

Private Function Byte2Long(ByRef byteArray() As Byte) As Long
  Dim x As Integer

  Byte2Long = 0
  For x = 0 To 3
    Byte2Long = Byte2Long + (byteArray(x) * (2 ^ ((3 - x) * 7)))
  Next x

End Function

Private Sub ReadFile(ByRef b() As Byte)
  Dim lLen As Long
  Dim vByte(3) As Byte

  Get Free, lPosition, vByte(0): lPosition = lPosition + 1 'The ID
  If vByte(0) <> 106 Then
    Err.Raise 566
    Exit Sub
  End If

  Get Free, lPosition, vByte: lPosition = lPosition + 4
  lLen = Byte2Long(vByte())
  ReDim b(lLen)
  Get Free, lPosition, b
  lPosition = lPosition + lLen
End Sub

Private Function ReadNumber() As Long
  Dim vType As Byte   'The type of number to read
  Dim vByte(3) As Byte '4 bytes to read longs into

  Get Free, lPosition, vType: lPosition = lPosition + 1
  If vType = 102 Then 'It is an byte
  
    Get Free, lPosition, vType: lPosition = lPosition + 1 'Get size
    ReadNumber = vType
    
  ElseIf vType = 101 Then 'It is an negative long
  
    Get Free, lPosition, vByte: lPosition = lPosition + 4
    ReadNumber = Byte2Long(vByte()) * -1

  ElseIf vType = 100 Then 'It is an positive long
  
    Get Free, lPosition, vByte: lPosition = lPosition + 4
    ReadNumber = Byte2Long(vByte())
    
  ElseIf vType = 200 Then 'The number is 0
    
    ReadNumber = 0
    
  Else 'This aint no number
    Err.Raise 566
  End If
End Function


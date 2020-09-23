VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Texture Class 1.1
'---------------------------------
'Created by Paul Berlin 2002-2003
'
'Using this class you load & manage textures.
'Texture sizes must be a power of 2! (2,4,8,16,32,64,128,256,512... etc.)
'
'OVERVIEW:
'* Use LoadFileTexture(sFilename, sID, lColorKey) to load a texture.
'* Use LoadMemTexture(SrcData(), sID, lColorKey) to load an texture file
'  from an array.
'
'CHANGELOG:
'1.1 06-06-2003
'* Rewrote entire texture system, from using two classes, one of them the texture pool
'  and the other representing each texture, to this small class.
Option Explicit

Private Type tTextureData
  sID As String
  oTex As Direct3DTexture8
  lWidth As Long
  lHeight As Long
End Type

Private cTex() As tTextureData

Public Sub LoadFileTexture(ByVal sFilename As String, ByVal sID As String, Optional lColorkey As Long = 0)
  Dim mipinfo As D3DSURFACE_DESC
  
  ReDim Preserve cTex(UBound(cTex) + 1)
  With cTex(UBound(cTex))
    .sID = sID
    Set .oTex = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, sFilename, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, lColorkey, ByVal 0, ByVal 0)
    .oTex.GetLevelDesc 0, mipinfo
    .lHeight = mipinfo.Height
    .lWidth = mipinfo.Width
  End With
  
End Sub

Public Sub LoadMemTexture(ByRef srcdata() As Byte, ByVal sID As String, Optional lColorkey As Long = 0)
  Dim mipinfo As D3DSURFACE_DESC
  
  ReDim Preserve cTex(UBound(cTex) + 1)
  With cTex(UBound(cTex))
    .sID = sID
    Set .oTex = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3DDevice, srcdata(0), UBound(srcdata), -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, lColorkey, ByVal 0, ByVal 0)
    .oTex.GetLevelDesc 0, mipinfo
    .lHeight = mipinfo.Height
    .lWidth = mipinfo.Width
  End With
  
End Sub

Public Function GetTexture(ByVal lNum As Long) As Direct3DTexture8
  Set GetTexture = cTex(lNum).oTex
End Function

Private Sub Class_Initialize()
  ReDim cTex(0)
End Sub

Public Function GetNumFromID(ByVal sID As String) As Long
  Dim x As Long
  For x = 1 To UBound(cTex)
    If cTex(x).sID = sID Then
      GetNumFromID = x
      Exit Function
    End If
  Next
End Function

Public Function GetDimension(ByVal sID As String) As RECT
  Dim x As Long
  x = GetNumFromID(sID)
  With GetDimension
    .Top = 0
    .Left = 0
    .Right = cTex(x).lWidth
    .bottom = cTex(x).lHeight
  End With
End Function

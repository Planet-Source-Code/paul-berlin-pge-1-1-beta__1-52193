Attribute VB_Name = "modMain"
Option Explicit

'Type RECT
'  Left As Long
'  Top As Long
'  Right As Long
'  Bottom As Long
'End Type

Public Const PrjVer As Byte = 1
Public Const ComVer As Byte = 1

Type tPrgVars
  bChanged As Boolean
  bConfirm(3) As Boolean
  bSnap As Boolean
  bSnapAuto As Boolean
  bShowAuto As Boolean
  bShowInView As Boolean
  lFrmBgCol As Long
  lFW As Long
  lFH As Long
  lAutoX(3) As Long
  lAutoY(3) As Long
  bAutoFrame As Boolean
  bAniAutochange As Boolean
End Type

Enum enumOptions
  VirtualSize_1 = 1             'Virtual frame size is included
  SubImages_2 = 2               'Animation have subimages
  Loop_4 = 4                    'Default animation to loop on
  Center_8 = 8                  'Default animation to center oriented on
End Enum

Type tResDest
  lRes As Long                'Pointer to resource containing frame data
  Pos As RECT                 'Location on resource image of frame data
  VirtualSize As RECT         'Size of frame to be used in intersecting
  Offset As D3DVECTOR2
End Type

Type tFrameControlInfo
  lDelay As Long              'The delay for this frame
End Type

Type tFrame
  img() As tResDest           'The image data for each of this frame's subimages
  Ctrl As tFrameControlInfo   'This frame's control info
End Type

Type tAnim
  sID As String               'ID of this animation
  sDescr As String            'Animation Description
  eOptions As enumOptions     'Options for this animation
  lSubimages As Long          'Number of subimages this animation has
  Frame() As tFrame           'the frames of this animation
End Type

Type tAnimCollection
  sID As String               'ID of this collection
  sDescr As String            'Collection Description
  Anim() As tAnim             'All animations in this collection
End Type

Type tResources
  sFilename As String         'The filename of resource (PNG-image)
  lTranscolor As Long         'If image has no alpha channel, this is the color that is transparent, else -1
End Type

Type tProject
  sID As String               'ID/Name of file
  sDescr As String            'Description/Notes/Copyrights
  sFilename As String         'Filename used when saving project
  sFilenameCompile As String  'Filename used when compiling
  eDefOptions As enumOptions  'Default animation options used when creating one
  Res() As tResources         'collection of resources (PNG-images)
  Coll() As tAnimCollection   'Collection of animations
End Type

Public Prj As tProject

Public Prg As tPrgVars

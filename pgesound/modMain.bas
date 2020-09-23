Attribute VB_Name = "modMain"
Option Explicit

Public Const PrjVer As Byte = 1
Public Const ComVer As Byte = 1

Enum enumMus
  mus1_Loop = 1
  mus2_reverb = 2
End Enum

Enum enumSfx
  Sfx1_Loop = 1
End Enum

Type tSfx
  sID As String
  sDescr As String
  sFile As String
  eFlags As enumSfx
  bVol As Byte
  bPlaybacks As Byte
  bPriority As Byte
End Type

Type tMus
  sID As String
  sDescr As String
  sFile As String
  eFlags As enumMus
  bVol As Byte
End Type

Type tPrj
  sID As String
  sDescr As String
  sFilename As String         'Filename used when saving project
  sFilenameCompile As String  'Filename used when compiling
  eSfxFlags As enumSfx
  eMusFlags As enumMus
  Sfx() As tSfx
  Mus() As tMus
End Type

Type tPrg
  bChanged As Boolean
  bConfirm As Boolean
End Type

Public Prj As tPrj
Public Prg As tPrg

Public ManStopped As Boolean

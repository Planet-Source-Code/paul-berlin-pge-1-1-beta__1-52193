VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Font Class 1.0
'------------------------------
'Created by Paul Berlin 2002-2003
'
'This class just holds the font and lets you print text with it.
'Use PgeText for more advanced options.
'
'OVERVIEW:
'---------
'* Use Create(fnt) to create an font. Use the included ReturnFont function to do
'  it easier.
'* Use DrawText(sText, rt, lColor, Format) to draw text onto the screen.
'* Use GetInto() to get some info about the font.
'
'CHANGELOG:
'----------
'
Option Explicit

Private oFont As D3DXFont
Private Info As LOGFONT

Public Function GetInfo() As LOGFONT
  GetInfo = Info
End Function

Public Sub Create(ByRef fnt As IFont)
  Set oFont = Direct3DX.CreateFont(Direct3DDevice, fnt.hFont)
  oFont.GetLogFont Info
End Sub

Public Sub DrawText(ByVal sText As String, rt As RECT, lColor As Long, Optional eFormat As CONST_DTFLAGS = 0)
  Direct3DX.DrawText oFont, lColor, sText, rt, eFormat
End Sub


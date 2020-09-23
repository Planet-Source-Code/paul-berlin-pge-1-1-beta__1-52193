VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pab Game Engine Demonstration - By Paul Berlin 2003"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mr As RECT
Dim pS As New pgeSprite
Dim bEnd As Byte 'This variable controls the flow of the program
Dim md As Boolean

Private Sub Form_Load()
  Me.Show 'This must be done before init for mouse & keyboard to work with directX
  DoEvents

  If InitEngine Then
    DoDemo
  End If
  
  bEnd = 1
  Unload Me
End Sub

Public Function InitEngine() As Boolean
  On Error GoTo errh
  InitEngine = True
  
  Randomize Timer
  
  pMain.Init Me.hwnd, True, , , False
  pKeyboard.Create Me.hwnd 'init keyboard
  pSound.Init , 64, 100
  Set pTex = pTextures 'This must be set when using sprites, so they can access textures
  pFontArial8.Create ReturnFont("Arial", 8)

  pTileset.LoadTiles App.Path & "\data\test.pbt"
  pTextures.LoadFileTexture App.Path & "\data\devil.png", "devil"
  
  pSound.SfxLoad "jingle", App.Path & "\data\jingle.mp3"
  
  Exit Function
errh:
  InitEngine = False
  MsgBox "Pge Engine Could not init.", vbCritical, "Init Error"
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  md = True
  Form_MouseMove Button, Shift, x, y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  mr.Left = x
  mr.Top = y
  mr.Right = x + 1
  mr.bottom = y + 1
  If md Then
    pS.SetAutoPathTime x, y, 1000
    On Error Resume Next
    pS.CurrentSubFrame = Int(GetAngle(pS.GetPosition.x, pS.GetPosition.y, x, y) / 22.5) + 1
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  md = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'This will set bEnd to 1 if it is not and cancel the unload,
  'so that the program will end correctly by exiting all its loops.
  If bEnd = 0 Then
    bEnd = 1
    Cancel = 1
  End If
End Sub

Public Sub DoDemo()
  Dim pt As New pgeText
  pt.Create pFontArial8
  pt.SetSize 100, 15
  pt.SetColor 255, 0, 0, 255
  
  pTileset.Create pS, "Arrow"
  pS.SetPosition 320, 240
  pMain.lClearColor = RGBA(255, 255, 255, 255)
  
  pKeyboard.SetTimerEx DIK_O, 500
  Dim lp As Long, lp2 As Long
  pSound.SfxPlay "jingle"
  'lp = pSound.SfxPlay("snd")
  'pSound.MusicPlay "mod"
  
  Do
    DoEvents
    pMain.Clear
    
      pS.Render
      pt.sCaption = pMain.lFPS
      pt.Render
      
    pMain.Render
    
    'pSound.SfxSetPlayingPanVol lp, vec2(mr.Left, mr.Top), pS.GetPosition
    'pSound.Update
    'If pKeyboard.KeyDown(DIK_P) Then pSound.MusicSetVolumeFade "mp3", 255, 3000
    If pKeyboard.KeyDown(DIK_ESCAPE) Then bEnd = 1
'    If pKeyboard.KeyDown(DIK_O) Then pSound.SfxSetPlayingPanningFade lp, 0, 2000, False
'    If pKeyboard.KeyDown(DIK_P) Then pSound.SfxSetPlayingPanningFade lp, 255, 2000, False
'    If pKeyboard.KeyDown(DIK_L) Then
'      lp2 = pSound.SfxPlay("crow")
'      pSound.SfxSetPlayingVolumeFade lp2, 0, 4000, True
'    End If
    If pKeyboard.KeyDown(DIK_A) Then pS.SetAutoScale 2, 2, 1000
    If pKeyboard.KeyDown(DIK_Z) Then pS.SetAutoScale 1, 1, 1000
    If pKeyboard.KeyDown(DIK_S) Then pS.SetAutoFade 255, 255, 255, 100, 1000
    If pKeyboard.KeyDown(DIK_X) Then pS.SetAutoFade 255, 255, 255, 255, 1000
    If pKeyboard.KeyDown(DIK_D) Then pS.SetAutoPathTime 500, 100, 1000
    If pKeyboard.KeyDown(DIK_C) Then pS.SetAutoPathTime 100, 400, 1000
    If pKeyboard.KeyDown(DIK_F) Then pS.SetAutoRotation 1, 10
    If pKeyboard.KeyDown(DIK_V) Then pS.SetAutoRotation 0, 10: pS.SetRotation 0
  Loop Until bEnd = 1
  
End Sub

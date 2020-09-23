VERSION 5.00
Begin VB.Form frmPlayback 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   1080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox p 
      Height          =   3015
      Index           =   1
      Left            =   600
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.PictureBox p 
      Height          =   3015
      Index           =   0
      Left            =   120
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Timer tmr 
      Interval        =   50
      Left            =   240
      Top             =   3360
   End
   Begin VB.CommandButton cmd 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmPlayback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SOK As Boolean
Dim ptrStream As Long
Dim ptrMusic As Long
Dim chnStream As Long

Private Sub cmd_Click(index As Integer)
  On Error Resume Next
  If SOK Then
    Dim x As Long, y As Long
    y = frmMain.Tabs.SelectedItem.index - 1
    x = frmMain.lst(y).ListIndex + 1
    Select Case index
      Case 0
        If y = 0 Then
          With Prj.Sfx(x)
            If ptrStream <> 0 Then
              FSOUND_StopSound FSOUND_ALL
              FSOUND_Sample_Free ptrStream
              ptrStream = 0
              chnStream = 0
            End If
            If ptrMusic <> 0 Then
              FMUSIC_StopSong ptrMusic
              FMUSIC_FreeSong ptrMusic
              ptrMusic = 0
            End If
            ptrStream = FSOUND_Sample_Load(FSOUND_FREE, .sFile, IIf((.eFlags And Sfx1_Loop) = Sfx1_Loop, 2, 1), 0)
            chnStream = FSOUND_PlaySound(FSOUND_FREE, ptrStream)
            FSOUND_SetVolume chnStream, .bVol
          End With
        Else
          With Prj.Mus(x)
            If ptrMusic <> 0 Then
              FMUSIC_StopSong ptrMusic
              FMUSIC_FreeSong ptrMusic
              ptrMusic = 0
            End If
            If ptrStream <> 0 Then
              FSOUND_StopSound FSOUND_ALL
              FSOUND_Sample_Free ptrStream
              ptrStream = 0
              chnStream = 0
            End If
            Select Case LCase(sFilename(.sFile, efpFileExt))
              Case "mp3", "wav", "ogg"
                ptrStream = FSOUND_Sample_Load(FSOUND_FREE, .sFile, IIf((.eFlags And Mus1_Loop) = Mus1_Loop, 2, 1), 0)
                chnStream = FSOUND_PlaySound(FSOUND_FREE, ptrStream)
                FSOUND_SetVolume chnStream, .bVol
              Case "mid", "midi", "xm", "it", "s3m", "mod"
                ptrMusic = FMUSIC_LoadSong(.sFile)
                FMUSIC_SetLooping ptrMusic, CBool((.eFlags And Mus1_Loop) = Mus1_Loop)
                FMUSIC_SetReverb CBool((.eFlags And Mus2_Reverb) = Mus2_Reverb)
                FMUSIC_SetMasterVolume ptrMusic, .bVol
                FMUSIC_PlaySong ptrMusic
            End Select
          End With
        End If
      Case 1
        If ptrStream <> 0 Then
            FSOUND_StopSound FSOUND_ALL
            FSOUND_Sample_Free ptrStream
            ptrStream = 0
            chnStream = 0
          End If
        If y = 1 Then
          If ptrMusic <> 0 Then
            FMUSIC_StopSong ptrMusic
            FMUSIC_FreeSong ptrMusic
            ptrMusic = 0
          End If
        End If
    End Select
  End If
End Sub

Private Sub Form_Load()
  On Error GoTo errh
  
  If Not SOK Then
    SOK = True
    
    FSOUND_SetBufferSize 100
    FSOUND_SetMixer FSOUND_MIXER_QUALITY_AUTODETECT
    FSOUND_SetOutput -1
    FSOUND_Init 44100, 64, FSOUND_INIT_GLOBALFOCUS Or FSOUND_INIT_ACCURATEVULEVELS
  
  End If
  
  Exit Sub
errh:
  SOK = False
  MsgBox "Could not init FMOD... fix your soundcard!", vbExclamation, "FMOD Error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If ptrStream <> 0 Then
    FSOUND_StopSound FSOUND_ALL
    FSOUND_Sample_Free ptrStream
    ptrStream = 0
    chnStream = 0
  End If
  If ptrMusic <> 0 Then
    FMUSIC_StopSong ptrMusic
    FMUSIC_FreeSong ptrMusic
  End If
  FSOUND_Close
End Sub

Private Sub tmr_Timer()
  On Error Resume Next
  If ptrStream <> 0 Then
    Dim l As Single, r As Single
    FSOUND_GetCurrentLevels chnStream, l, r
    p(0).Cls
    p(0).Line (0, p(0).ScaleHeight - l * p(0).ScaleHeight)-(p(0).ScaleWidth, p(0).ScaleHeight), vbRed, BF
    p(1).Cls
    p(1).Line (0, p(1).ScaleHeight - r * p(1).ScaleHeight)-(p(1).ScaleWidth, p(1).ScaleHeight), vbRed, BF
  End If
End Sub

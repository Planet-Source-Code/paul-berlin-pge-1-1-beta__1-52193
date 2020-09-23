VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Sound Class 1.0
'-------------------------------
'Created by Paul Berlin 2002-2003
'
'Used to load & play sound effect & music.
'Handles sound effects & music using the fmod sound system.
'
'Filetypes supported for music:
'Type Stream - MP3, MP2, OGG, WMA, ASF, MID, MIDI, RMI, WAV, SGM
'Type Music - MOD, S3M, XM, IT
'WMA & ASF will probably not work very well or not at all, stick to MP3 or OGG instead
'Some MOD, S3M, XM & IT crashed fmod (very rare) at an certain point
'Filetypes supported for soundeffects:
'MP3, MP2, OGG, WAV
'
'OVERVIEW:
'
'
'CHANGELOG:
'

Private Enum enumMus
  Mus1_Loop = 1
  Mus2_Reverb = 2
End Enum

Private Enum enumSfx
  Sfx1_Loop = 1
End Enum

Public Enum enumPgePlayMode
  LOOP_OFF = 1
  LOOP_NORMAL = 2
  LOOP_BIDI = 4
End Enum

Private Enum tSoundType
  st1_Sfx = 1         'sound effect or streamed music in memory
  st2_Mod = 2         'module/midi both from file and memory
  st4_Stream = 4      'streamed music
End Enum

Private Type tSound
  sID As String       'Id of sound
  lPointer As Long    'Pointer to sound in memory
End Type

Private Type tChange
  sCur As Single
  sDest As Single
  sChange As Single
  lTime As Long
  bAutoStop As Boolean
End Type

Private Type tPlaying
  lPointer As Long
  Volume As tChange
  pan As tChange
  freq As tChange
End Type

Private Type tMusic
  sID As String       'Id of sound
  lPointer As Long    'Pointer to sound in memory
  lChannel As Long    'Channel used to play
  eType As tSoundType 'type of sound
  Volume As tChange
  bLoop As Boolean
End Type

Public MasterSfxVolume As Byte
Public MasterMusicVolume As Byte

Private Snd() As tSound
Private Mus() As tMusic
Private Playing() As tPlaying

Private lMusicPlaying As Long 'Pointer to music pool of currently playing music
Private lMusicChannel As Long 'Pointer to channel stream music is playing on

Private Const ComVer As Byte = 1

Option Explicit

Public Function Init(Optional ByVal lFreq As Long = 44100, Optional lChannels As Long = 32, Optional ByVal lBufferMs As Long = 200, Optional ByVal eOutput As FSOUND_OUTPUTTYPES = FSOUND_OUTPUT_DSOUND, Optional ByVal eMixer As FSOUND_MIXERTYPES = FSOUND_MIXER_QUALITY_AUTODETECT, Optional ByVal lDriver As Long = 0) As Boolean
  'This function must be called before any other
  'It initializes the fmod system
  On Error GoTo errh
  
  Call FSOUND_SetBufferSize(lBufferMs)
  Call FSOUND_SetOutput(eOutput)
  Call FSOUND_SetMixer(eMixer)
  Call FSOUND_SetDriver(lDriver)
  Call FSOUND_Init(lFreq, lChannels, 0)
  
  Init = True
  Exit Function
errh:
  Init = False
End Function

Private Sub Class_Initialize()
  ReDim Mus(0)
  ReDim Snd(0)
  ReDim Playing(0)
  MasterSfxVolume = 100
  MasterMusicVolume = 100
End Sub

Private Sub Class_Terminate()
  Dim x As Long
  FSOUND_StopSound FSOUND_ALL
  For x = 1 To UBound(Mus)
    With Mus(x)
      If .eType = st4_Stream Then
        FSOUND_Stream_Close .lPointer
      ElseIf .eType = st2_Mod Then
        FMUSIC_FreeSong .lPointer
      End If
    End With
  Next
  FSOUND_Close
End Sub

Public Function LoadSoundset(ByVal sFile As String) As Boolean
'  'On Error GoTo errh
'  LoadSoundset = True
'
'  Dim F As New pgeFile, x As Long, y As enumSfx, b() As Byte
'  F.filename = sFile
'  sFile = F.ReadStrFixed(6)
'  If Left(sFile, 5) <> "PGECS" Or Asc(Right(sFile, 1)) <> ComVer Then GoTo errh
'  F.SkipField 2 'skip id & description
'  For x = 1 To F.ReadNumber
'    ReDim Preserve Snd(UBound(Snd) + 1)
'    With Snd(UBound(Snd))
'      F.ReadFileMem b()
'      Debug.Print VarPtrArray(b()); UBound(b); Time
'      .lPointer = FSOUND_Sample_Load(FSOUND_FREE, b(0), FSOUND_LOADMEMORY, UBound(b))
'      .sID = F.ReadStr
'      FSOUND_Sample_SetDefaults .lPointer, -1, F.ReadNumber, -1, F.ReadNumber
'      FSOUND_Sample_SetMode .lPointer, IIf((F.ReadNumber And Sfx1_Loop) = Sfx1_Loop, 2, 1)
'      FSOUND_Sample_SetMaxPlaybacks .lPointer, F.ReadNumber
'    End With
'  Next
'  For x = 1 To F.ReadNumber
'    ReDim Preserve Mus(UBound(Mus) + 1)
'    With Mus(UBound(Mus))
'      Select Case F.ReadNumber
'        Case 1
'          F.ReadFileMem b()
'          .lPointer = FSOUND_Sample_Load2(FSOUND_FREE, VarPtrArray(b), FSOUND_LOADMEMORY, UBound(b))
'        Case 2
'          F.ReadFileMem b()
'          Debug.Print UBound(b); Time
'          .lPointer = FMUSIC_LoadSongEx(VarPtrArray(b), UBound(b), FSOUND_LOADMEMORY, 0, 0)
'
'      End Select
'      .sID = F.ReadStr
'      .bLoop = IIf((F.ReadNumber And Sfx1_Loop) = Sfx1_Loop, 2, 1)
'      .Volume.sCur = F.ReadNumber
'    End With
'  Next
'
'
'  Exit Function
'errh:
'  LoadSoundset = False
End Function

Public Sub GetDrivers(ByVal eOutput As FSOUND_OUTPUTTYPES, ByRef sDrivers() As String)
  'This returns the available drivers for the specified output type
  'Could be useful if you want the user to select an sound device to use
  'this must be called BEFORE init
  FSOUND_SetOutput eOutput
  Dim x As Long
  ReDim sDrivers(0)
  For x = 0 To FSOUND_GetNumDrivers - 1
    ReDim Preserve sDrivers(x)
    sDrivers(x) = GetStringFromPointer(FSOUND_GetDriverName(x))
  Next
End Sub

Private Function GetNum(ByVal sID As String) As Long
  Dim x As Long
  For x = 1 To UBound(Snd)
    If Snd(x).sID = sID Then
      GetNum = x
      Exit Function
    End If
  Next
End Function

Private Function GetNumMus(ByVal sID As String) As Long
  Dim x As Long
  For x = 1 To UBound(Mus)
    If Mus(x).sID = sID Then
      GetNumMus = x
      Exit Function
    End If
  Next
End Function

Private Function NewPlayingNum() As Long
  Dim x As Long
  For x = 1 To UBound(Playing)
    If Not SfxIsPlaying(Playing(x).lPointer) Then
      NewPlayingNum = x
      Exit Function
    End If
  Next
  ReDim Preserve Playing(UBound(Playing) + 1)
  NewPlayingNum = UBound(Playing)
End Function

Private Function GetPlayingNum(ByVal lPointer As Long) As Long
  Dim x As Long
  For x = 1 To UBound(Playing)
    If Playing(x).lPointer = lPointer Then
      GetPlayingNum = x
      Exit Function
    End If
  Next
End Function

Public Sub MusicLoad(ByVal sID As String, ByVal sFile As String, ByVal bModMidi As Boolean)

  ReDim Preserve Mus(UBound(Mus) + 1)
  With Mus(UBound(Mus))
    .sID = sID
    .Volume.sCur = 255
    .bLoop = True
    If bModMidi Then
      .eType = st2_Mod
      .lPointer = FMUSIC_LoadSong(sFile)
    Else
      .eType = st4_Stream
      .lPointer = FSOUND_Stream_OpenFile(sFile, 0, 0)
    End If
  End With

End Sub

Public Sub MusicLoadEx(ByVal sID As String, ByVal sFile As String, ByVal bModMidi As Boolean, Optional bLoop As Boolean = True, Optional bVolume As Byte = 255)

  ReDim Preserve Mus(UBound(Mus) + 1)
  With Mus(UBound(Mus))
    .sID = sID
    .Volume.sCur = bVolume
    .bLoop = bLoop
    If bModMidi Then
      .eType = st2_Mod
      .lPointer = FMUSIC_LoadSong(sFile)
    Else
      .eType = st4_Stream
      .lPointer = FSOUND_Stream_OpenFile(sFile, 0, 0)
    End If
  End With

End Sub

Public Sub MusicPlay(ByVal sID As String)

  With Mus(GetNumMus(sID))
    If .eType = st2_Mod Then
      If FMUSIC_IsPlaying(.lPointer) Then FMUSIC_StopSong .lPointer
      FMUSIC_SetLooping .lPointer, .bLoop
      FMUSIC_PlaySong .lPointer
      FMUSIC_SetMasterVolume .lPointer, CByte(.Volume.sCur * (MasterMusicVolume / 100))
    ElseIf .eType = st4_Stream Then
      If .lChannel <> 0 Then
        FSOUND_StopSound .lChannel
        FSOUND_Stream_SetTime .lPointer, 0
      End If
      FSOUND_Stream_SetMode .lPointer, IIf(.bLoop, FSOUND_LOOP_NORMAL, FSOUND_LOOP_OFF)
      .lChannel = FSOUND_Stream_PlayEx(FSOUND_FREE, .lPointer, 0, True)
      FSOUND_SetVolume .lChannel, CByte(.Volume.sCur * (MasterMusicVolume / 100))
      FSOUND_SetPaused .lChannel, False
    ElseIf .eType = st1_Sfx Then
      If .lChannel <> 0 Then FSOUND_StopSound .lChannel
      .lChannel = FSOUND_PlaySoundEx(FSOUND_FREE, .lPointer, 0, True)
      FSOUND_Sample_SetMode .lPointer, IIf(.bLoop, FSOUND_LOOP_NORMAL, FSOUND_LOOP_OFF)
      FSOUND_SetVolume .lChannel, CByte(.Volume.sCur * (MasterMusicVolume / 100))
      FSOUND_SetPaused .lChannel, False
    End If
  End With

End Sub

Public Sub MusicStop(ByVal sID As String)

  With Mus(GetNumMus(sID))
    If .eType = st2_Mod Then
      FMUSIC_StopSong .lPointer
    ElseIf .eType = st4_Stream Then
      If .lChannel <> 0 Then
        FSOUND_StopSound .lChannel
        FSOUND_Stream_SetTime .lPointer, 0
        .lChannel = 0
      End If
    ElseIf .eType = st1_Sfx Then
      If .lChannel <> 0 Then FSOUND_StopSound .lChannel
      .lChannel = 0
    End If
  End With

End Sub

Public Sub MusicSetup(ByVal sID As String, Optional ByVal lVol As Long = -1, Optional ByVal bLoop As Boolean = True, Optional ByVal Speed As Single = -1)

  With Mus(GetNumMus(sID))
    If lVol > -1 Then .Volume.sCur = lVol
    .bLoop = bLoop
    If .eType = st2_Mod Then
      If Speed > -1 Then FMUSIC_SetMasterSpeed Mus(GetNumMus(sID)).lPointer, Speed
      FMUSIC_SetLooping .lPointer, .bLoop
      FMUSIC_SetMasterVolume .lPointer, CByte(.Volume.sCur * (MasterMusicVolume / 100))
    ElseIf .eType = st4_Stream Then
      If Speed > -1 Then FSOUND_SetFrequency .lChannel, Int(Speed)
      FSOUND_Stream_SetMode .lPointer, IIf(.bLoop, FSOUND_LOOP_NORMAL, FSOUND_LOOP_OFF)
      FSOUND_SetVolume .lChannel, CByte(.Volume.sCur * (MasterMusicVolume / 100))
    ElseIf .eType = st1_Sfx Then
      If Speed > -1 Then FSOUND_SetFrequency .lChannel, Int(Speed)
      FSOUND_Sample_SetMode .lPointer, IIf(.bLoop, FSOUND_LOOP_NORMAL, FSOUND_LOOP_OFF)
      FSOUND_SetVolume .lChannel, CByte(.Volume.sCur * (MasterMusicVolume / 100))
    End If
  End With

End Sub

Public Sub MusicSetVolumeFade(ByVal sID As String, Optional ByVal bDestVol As Byte = 0, Optional ByVal lTimeMs As Long = 1000, Optional ByVal bAutoStop As Boolean = False)
  With Mus(GetNumMus(sID)).Volume
    .sDest = bDestVol
    .sChange = (.sDest - .sCur) / (lTimeMs / 10)
    .bAutoStop = bAutoStop
  End With
End Sub

Public Sub SfxLoad(ByVal sID As String, ByVal sFile As String)
  
  ReDim Preserve Snd(UBound(Snd) + 1)
  With Snd(UBound(Snd))
    .sID = sID
    .lPointer = FSOUND_Sample_Load(FSOUND_FREE, sFile, 0, 0)
  End With
  
End Sub

Public Sub SfxLoadEx(ByVal sID As String, ByVal sFile As String, Optional ByVal lPriority As Long = -1, Optional ByVal lPlaybacks As Long = -1, Optional ByVal eLoop As enumPgePlayMode = -1)
  'The three last values can all be skipped by entering -1
  
  ReDim Preserve Snd(UBound(Snd) + 1)
  With Snd(UBound(Snd))
    .sID = sID
    .lPointer = FSOUND_Sample_Load(FSOUND_FREE, sFile, 0, 0)
    FSOUND_Sample_SetDefaults .lPointer, -1, -1, -1, lPriority
    If lPlaybacks <> -1 Then FSOUND_Sample_SetMaxPlaybacks .lPointer, lPlaybacks
    If eLoop <> -1 Then FSOUND_Sample_SetMode .lPointer, eLoop
  End With
  
End Sub

Public Sub SfxSetup(ByVal sID As String, Optional ByVal lFreq As Long = -1, Optional ByVal lVol As Long = -1, Optional ByVal lPan As Long = -1)
  'All values can be skipped by entering -1
  
  With Snd(GetNum(sID))
    FSOUND_Sample_SetDefaults .lPointer, lFreq, lVol, lPan, -1
  End With

End Sub

Public Sub SfxSetupEx(ByVal sID As String, Optional ByVal lFreq As Long = -1, Optional ByVal lVol As Long = -1, Optional ByVal lPan As Long = -1, Optional ByVal lPriority As Long = -1, Optional ByVal lPlaybacks As Long = -1, Optional ByVal eLoop As enumPgePlayMode = -1)
  'All values can be skipped by entering -1
  
  With Snd(GetNum(sID))
    FSOUND_Sample_SetDefaults .lPointer, lFreq, lVol, lPan, lPriority
    If lPlaybacks > -1 Then FSOUND_Sample_SetMaxPlaybacks .lPointer, lPlaybacks
    If eLoop > -1 Then FSOUND_Sample_SetMode .lPointer, eLoop
  End With

End Sub

Public Function SfxPlay(ByVal sID As String) As Long
  Dim v(2) As Long
  
  With Snd(GetNum(sID))
    SfxPlay = FSOUND_PlaySoundEx(FSOUND_FREE, .lPointer, 0, True)
    FSOUND_Sample_GetDefaults .lPointer, v(0), v(1), v(2), 0
  End With
  
  FSOUND_SetFrequency SfxPlay, v(0)
  FSOUND_SetPan SfxPlay, v(2)
  FSOUND_SetVolume SfxPlay, CByte(v(1) * (MasterSfxVolume / 100))
  
  With Playing(NewPlayingNum)
    .lPointer = SfxPlay
    .freq.sCur = v(0)
    .freq.sChange = 0
    .pan.sCur = v(2)
    .pan.sChange = 0
    .Volume.sCur = v(1)
    .Volume.sChange = 0
  End With
  
  FSOUND_SetPaused SfxPlay, False
    
End Function

Public Function SfxPlayEx(ByVal sID As String, Optional ByVal lVol As Long = -1, Optional ByVal lPan As Long = -1) As Long
  Dim v(2) As Long
  
  With Snd(GetNum(sID))
    SfxPlayEx = FSOUND_PlaySoundEx(FSOUND_FREE, .lPointer, 0, True)
    FSOUND_Sample_SetDefaults .lPointer, -1, lVol, lPan, -1
    FSOUND_Sample_GetDefaults .lPointer, v(0), v(1), v(2), 0
  End With
  
  FSOUND_SetFrequency SfxPlayEx, v(0)
  FSOUND_SetPan SfxPlayEx, v(2)
  FSOUND_SetVolume SfxPlayEx, CByte(v(1) * (MasterSfxVolume / 100))
  
  With Playing(NewPlayingNum)
    .lPointer = SfxPlayEx
    .freq.sCur = v(0)
    .freq.sChange = 0
    .pan.sCur = v(2)
    .pan.sChange = 0
    .Volume.sCur = v(1)
    .Volume.sChange = 0
  End With
  
  FSOUND_SetPaused SfxPlayEx, False
  
End Function

Public Function SfxIsPlaying(ByVal lPointer As Long) As Boolean
  'This will return true if specified sound is playing
  SfxIsPlaying = FSOUND_IsPlaying(lPointer)
End Function

Public Sub SfxStop(ByVal lPointer As Long)
  Call FSOUND_StopSound(lPointer)
End Sub

Public Sub SfxEndLoop(ByVal lPointer As Long)
  'This will end the loop of the sound so that it wont loop any more
  'the sound will eventually stop, when it has played
  FSOUND_SetLoopMode lPointer, False
End Sub

Public Sub SfxSetupPlaying(ByVal lPointer As Long, Optional ByVal lFreq As Long = -1, Optional ByVal lVol As Long = -1, Optional ByVal lPan As Long = -1)
  'This will only affect the playing sound
  If lFreq > -1 Then FSOUND_SetFrequency lPointer, lFreq
  If lPan > -1 Then FSOUND_SetPan lPointer, lPan
  If lVol > -1 Then FSOUND_SetVolume lPointer, CByte(lVol * (MasterSfxVolume / 100))
End Sub

Public Sub SfxSetPlayingPanVol(ByVal lPointer As Long, ByRef EarPos As D3DVECTOR2, ByRef SoundPos As D3DVECTOR2, Optional ByVal sVolFalloffPerPixel As Single = 0.5, Optional ByVal bPanFalloffPerPixel As Single = 1)
  Dim x As Long
  x = 255 - (GetDist(EarPos.x, EarPos.y, SoundPos.x, SoundPos.y) * sVolFalloffPerPixel)
  FSOUND_SetVolume lPointer, CByte(tob2(x) * (MasterSfxVolume / 100))
  If EarPos.x < SoundPos.x Then
    x = 128 + (SoundPos.x - EarPos.x) * bPanFalloffPerPixel
  Else
    x = 128 - (EarPos.x - SoundPos.x) * bPanFalloffPerPixel
  End If
  FSOUND_SetPan lPointer, tob2(x)
End Sub

Public Sub Update()
  Dim x As Long, tSng As Single, b As Boolean
  
  For x = 1 To UBound(Playing)
    b = False
    '###### VOLUME AUTOFADE
    With Playing(x).Volume
      If .sChange <> 0 Then
        b = True
        If FrameSkip And .lTime <> 0 Then
          tSng = (timeGetTime - .lTime) / 10
        Else
          tSng = 1
        End If
        .lTime = timeGetTime
        
          .sCur = .sCur + (.sChange * tSng)
          If .sChange < 0 Then
            If .sCur <= .sDest Then
              .sCur = .sDest
              .sChange = 0
            End If
          Else
            If .sCur >= .sDest Then
              .sCur = .sDest
              .sChange = 0
            End If
          End If
        
        If .sCur = .sDest And .bAutoStop Then
          SfxStop Playing(x).lPointer
        End If
      Else
        .lTime = 0
      End If
    End With
    
    '###### PANNING AUTOFADE
    With Playing(x).pan
      If .sChange <> 0 Then
        b = True
        If FrameSkip And .lTime <> 0 Then
          tSng = (timeGetTime - .lTime) / 10
        Else
          tSng = 1
        End If
        .lTime = timeGetTime
        
          .sCur = .sCur + (.sChange * tSng)
          If .sChange < 0 Then
            If .sCur <= .sDest Then
              .sCur = .sDest
              .sChange = 0
            End If
          Else
            If .sCur >= .sDest Then
              .sCur = .sDest
              .sChange = 0
            End If
          End If
        
        If .sCur = .sDest And .bAutoStop Then
          SfxStop Playing(x).lPointer
        End If
      Else
        .lTime = 0
      End If
    End With
    
    '###### FREQUENCY AUTOFADE
    With Playing(x).freq
      If .sChange <> 0 Then
        b = True
        If FrameSkip And .lTime <> 0 Then
          tSng = (timeGetTime - .lTime) / 10
        Else
          tSng = 1
        End If
        .lTime = timeGetTime
        
          .sCur = .sCur + (.sChange * tSng)
          If .sChange < 0 Then
            If .sCur <= .sDest Then
              .sCur = .sDest
              .sChange = 0
            End If
          Else
            If .sCur >= .sDest Then
              .sCur = .sDest
              .sChange = 0
            End If
          End If
        
        If .sCur = .sDest And .bAutoStop Then
          SfxStop Playing(x).lPointer
        End If
      Else
        .lTime = 0
      End If
    End With
    
    If b Then
      With Playing(x)
        FSOUND_SetPan .lPointer, .pan.sCur
        FSOUND_SetVolume .lPointer, CByte(.Volume.sCur * (MasterSfxVolume / 100))
        FSOUND_SetFrequency .lPointer, .freq.sCur
      End With
    End If
  
  Next
  
  For x = 1 To UBound(Mus)
    b = False
    With Mus(x).Volume
      If .sChange <> 0 Then
        b = True
        If FrameSkip And .lTime <> 0 Then
          tSng = (timeGetTime - .lTime) / 10
        Else
          tSng = 1
        End If
        .lTime = timeGetTime
        
          .sCur = .sCur + (.sChange * tSng)
          If .sChange < 0 Then
            If .sCur <= .sDest Then
              .sCur = .sDest
              .sChange = 0
            End If
          Else
            If .sCur >= .sDest Then
              .sCur = .sDest
              .sChange = 0
            End If
          End If
        
        If .sCur = .sDest And .bAutoStop Then
          MusicStop Mus(x).sID
        End If
      Else
        .lTime = 0
      End If
    End With
    
    If b Then
      With Mus(x)
        If .eType = st2_Mod Then
          FMUSIC_SetMasterVolume .lPointer, CByte(.Volume.sCur * (MasterMusicVolume / 100))
        Else
          If .lChannel <> 0 Then
            FSOUND_SetVolume .lChannel, CByte(.Volume.sCur * (MasterMusicVolume / 100))
          End If
        End If
      End With
    End If
  Next
End Sub

Public Sub SfxSetPlayingVolumeFade(ByVal lPointer As Long, Optional ByVal bDestVol As Byte = 0, Optional ByVal lTimeMs As Long = 1000, Optional ByVal bAutoStop As Boolean = False)
  With Playing(GetPlayingNum(lPointer)).Volume
    .sDest = bDestVol
    .sChange = (.sDest - .sCur) / (lTimeMs / 10)
    .bAutoStop = bAutoStop
  End With
End Sub

Public Sub SfxSetPlayingPanningFade(ByVal lPointer As Long, Optional ByVal lDestPan As Long = 127, Optional ByVal lTimeMs As Long = 1000, Optional ByVal bAutoStop As Boolean = False)
  With Playing(GetPlayingNum(lPointer)).pan
    .sDest = lDestPan
    .sChange = (.sDest - .sCur) / (lTimeMs / 10)
    .bAutoStop = bAutoStop
  End With
End Sub

Public Sub SfxSetPlayingFrequencyFade(ByVal lPointer As Long, ByVal lDestFreq As Long, Optional ByVal lTimeMs As Long = 1000, Optional ByVal bAutoStop As Boolean = False)
  With Playing(GetPlayingNum(lPointer)).freq
    .sDest = lDestFreq
    .sChange = (.sDest - .sCur) / (lTimeMs / 10)
    .bAutoStop = bAutoStop
  End With
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFmod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'By Paul Berlin 2002
'Version 1.2
'This version was stripped & modified for use in PgeSound

Public Enum enumPgePlayMode
  LOOP_OFF = 1
  LOOP_NORMAL = 2
  LOOP_BIDI = 4
End Enum

Dim stream As Long 'Pointers to loaded streams
Dim channel As Long 'Pointers to used channels
Dim module As Long 'pointer to loaded module

Private Sub Class_Terminate()
  FSOUND_StopSound FSOUND_ALL
  If stream <> 0 Then FSOUND_Sample_Free stream
  If module <> 0 Then
    FMUSIC_StopSong module
    FMUSIC_FreeSong module
  End If
  FSOUND_Close
End Sub

Public Function Init(ByVal lMixrate As Long, ByVal lChannels As Long) As Boolean
  'This inits FMOD with selected settings
  On Error GoTo errh
  
  If FSOUND_Init(lMixrate, lChannels, FSOUND_INIT_GLOBALFOCUS Or FSOUND_INIT_ACCURATEVULEVELS Or FSOUND_INIT_ENABLEOUTPUTFX) = False Then GoTo errh
  
  Init = True
  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "Init - FMOD Error"
  Init = False
End Function

Public Function InitOutput(ByVal OutType As enumInitOutput) As Boolean
  'Sets output type
  'MUST BE CALLED BEFORE Init
  'use ReturnDrivers to see which drives support this output type
  On Error GoTo errh
  InitOutput = True
  
  If OutType = Out_Autodetect Then
    If FSOUND_SetOutput(-1) = False Then GoTo errh
  ElseIf OutType = Out_DirectSound Then
    If FSOUND_SetOutput(FSOUND_OUTPUT_DSOUND) = False Then GoTo errh
  ElseIf OutType = Out_WinMM Then
    If FSOUND_SetOutput(FSOUND_OUTPUT_WINMM) = False Then GoTo errh
  End If
  
  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "InitOutput - FMOD Error"
  InitOutput = False
End Function

Public Function InitMixer(ByVal mixer As enumInitMixer) As Boolean
  'Sets mixer to use
  'MUST BE CALLED BEFORE Init
  'autodetect, CPU/FPU 32-bit, MMX P5 16-bit, MMX P6 16-bit
  On Error GoTo errh
  InitMixer = True
  
  Select Case mixer
    Case Mix_Autodetect
      If FSOUND_SetMixer(FSOUND_MIXER_QUALITY_AUTODETECT) = False Then GoTo errh
    Case Mix_CPU_FPU_32
      If FSOUND_SetMixer(FSOUND_MIXER_QUALITY_FPU) = False Then GoTo errh
    Case Mix_MMX_P5_16
      If FSOUND_SetMixer(FSOUND_MIXER_QUALITY_MMXP5) = False Then GoTo errh
    Case Mix_MMX_P6_16
      If FSOUND_SetMixer(FSOUND_MIXER_QUALITY_MMXP6) = False Then GoTo errh
  End Select
  
  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "InitMixer - FMOD Error"
  InitMixer = False
End Function

Public Function InitBuffer(ByVal ms As Long) As Boolean
  'Sets buffer size in ms
  'MUST BE CALLED BEFORE Init
  InitBuffer = True
  If FSOUND_SetBufferSize(ms) = False Then GoTo errh

  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "InitBuffer - FMOD Error"
  InitBuffer = False
End Function

Public Function InitDriver(ByVal DriveNum As Byte) As Boolean
  'Must be called before init
  'drive num 0 will set to default driver
  On Error GoTo errh
  InitDriver = True
  If FSOUND_SetDriver(DriveNum) = False Then GoTo errh
  
  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "InitDriver - FMOD Error"
  InitDriver = False
End Function

Public Function SfxPlay(ByVal sFilename As String, ByVal bVolume As Byte, ByVal e As enumPgePlayMode) As Boolean
  'Will play the filename stream
  On Error GoTo errh
  SfxPlay = True

  FSOUND_StopSound FSOUND_ALL
  If stream <> 0 Then FSOUND_Sample_Free stream
  
  stream = FSOUND_Sample_Load(FSOUND_FREE, sFilename, e, 0)
  
  If stream = 0 Then
    MsgBox "Could not recognize & play " & Chr(34) & sFilename & Chr(34) & ".", vbExclamation, "FMOD Error"
    SfxPlay = False
    Exit Function
  End If
  
  channel = FSOUND_PlaySound(FSOUND_FREE, stream)
  FSOUND_SetVolume channel, bVolume
  
  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "SfxPlay - FMOD Error"
  SfxPlay = False
End Function

Public Function StreamPausePlay() As Boolean
  'this will pause or play the current stream
  On Error GoTo errh
  
  If stream <> 0 Then
    StreamPausePlay = True

    If FSOUND_GetPaused(channel) Then
      FSOUND_SetPaused channel, False
    Else
      FSOUND_SetPaused channel, True
    End If
    
  End If
  
  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "StremPausePlay - FMOD Error"
  StreamPausePlay = False
End Function

Public Function StreamStop() As Boolean
  'this will stop and resume stream
  On Error GoTo errh
  StreamStop = True

  If stream <> 0 Then
    StreamStop = True
    If channel <> 0 Then
      FSOUND_StopSound channel
      channel = 0
    Else
      channel = FSOUND_Stream_Play(FSOUND_FREE, stream)
    End If
  End If

  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "StreamStop - FMOD Error"
  StreamStop = False
End Function

Public Function StreamUnload() As Boolean
  'this will stop and unload stream
  On Error GoTo errh
  StreamUnload = True

  FSOUND_StopSound FSOUND_ALL
  If stream <> 0 Then FSOUND_Stream_Close stream
  stream = 0
  channel = 0
    

  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "StreamUnload - FMOD Error"
  StreamUnload = False
End Function

Public Property Get StreamFrequency() As Long
  'NOTE: FSOUND_GetFrequency(Channel) does not seem to work with wma, asf
  StreamFrequency = FSOUND_GetFrequency(channel)
End Property

Public Property Let StreamFrequency(ByVal lFreq As Long)
  FSOUND_SetFrequency channel, lFreq
End Property

Public Function StreamMute() As Boolean
  'this will mute/unmute stream
  On Error GoTo errh
  
  If stream <> 0 Then
    StreamMute = True

    FSOUND_SetMute channel, Not FSOUND_GetMute(channel)
    
  End If
  
  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "StreamMute - FMOD Error"
  StreamMute = False
End Function

Public Property Get StreamVolume() As Byte
  StreamVolume = FSOUND_GetVolume(channel)
End Property

Public Property Let StreamVolume(ByVal bVolume As Byte)
  FSOUND_SetVolume channel, bVolume
End Property

Public Property Get StreamIsLoaded() As Boolean
  If stream <> 0 Then StreamIsLoaded = True
End Property

Public Property Get StreamIsPlaying() As Boolean
  If channel <> 0 Then
    If FSOUND_GetPaused(channel) = 0 Then StreamIsPlaying = True
  End If
End Property

Public Property Get StreamSongLen() As Long
  If stream <> 0 Then
    StreamSongLen = FSOUND_Stream_GetLengthMs(stream)
  End If
End Property

Public Property Get StreamSongPos() As Long
  If stream <> 0 And channel <> 0 Then
    StreamSongPos = FSOUND_Stream_GetTime(stream)
  End If
End Property

Public Property Let StreamSongPos(ByVal lPos As Long)
  If stream <> 0 And channel <> 0 Then
    FSOUND_Stream_SetTime stream, lPos
  End If
End Property

Public Function GetStreamLength(ByVal filename As String, Optional ByVal VBR As Boolean = False) As Long
  'This opens an stream, gets the length, and closes it
  On Error GoTo errh
  Dim tmpStream As Long
  
  'if the mp3 is vbr, read whole file into memory
  If VBR Then
    tmpStream = FSOUND_Stream_OpenFile(filename, FSOUND_NORMAL Or FSOUND_MPEGACCURATE, 0)
  Else
    tmpStream = FSOUND_Stream_OpenFile(filename, FSOUND_NORMAL, 0)
  End If
  GetStreamLength = FSOUND_Stream_GetLengthMs(tmpStream) / 1000
  FSOUND_Stream_Close tmpStream
  
  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "GetStreamLength - FMOD Error"
End Function

Public Sub GetMusicData(ByVal filename As String, ByRef length As Long, ByRef name As String)
  'This opens an music file, gets the length, and closes it
  'This will calculate the songs length in ms
  'This is mostly accurate, but can be highly inaccurate
  On Error Resume Next
  Dim Kbps As Double, st As Double, i As Long, t As Long
  Dim tmpMusic As Long
  
  tmpMusic = FMUSIC_LoadSong(filename)  'load song
  If FMUSIC_GetType(tmpMusic) = 0 Then
    FMUSIC_FreeSong tmpMusic
    Exit Sub
  End If
  Call FMUSIC_SetMasterVolume(tmpMusic, 0)  'set volume to 0
  Call FMUSIC_PlaySong(tmpMusic)  'play song, must be played to get BPM
  'set order to in the middle, because this will increase the accuracy of the time
  Call FMUSIC_SetOrder(tmpMusic, Int(FMUSIC_GetNumOrders(tmpMusic) / 2))
  name = Trim(GetStringFromPointer(FMUSIC_GetName(tmpMusic))) 'get name
  
  'calc time
  Kbps = ((FMUSIC_GetBPM(tmpMusic) * 2) / 5)
  st = ((1 / Kbps) * 1000) * FMUSIC_GetSpeed(tmpMusic)
  For i = 0 To (FMUSIC_GetNumOrders(tmpMusic) - 1)
    t = t + FMUSIC_GetPatternLength(tmpMusic, i)
  Next i
  length = (t * st) / 1000
  If length > 3600 Then length = 0
  If length < 0 Then length = 0
  
  'stop and unload song
  FMUSIC_StopSong tmpMusic
  FMUSIC_FreeSong tmpMusic
  
End Sub

Public Function GetMusicLength(ByVal filename As String) As Long
  'This opens an music file, gets the length, and closes it
  'This will calculate the songs length in ms
  'This is mostly accurate, but can be highly inaccurate
  On Error Resume Next
  Dim Kbps As Double, st As Double, i As Long, t As Long
  Dim tmpMusic As Long
  
  tmpMusic = FMUSIC_LoadSong(filename)
  If FMUSIC_GetType(tmpMusic) = 0 Then
    FMUSIC_FreeSong tmpMusic
    Exit Function
  End If
  
  Kbps = ((FMUSIC_GetBPM(tmpMusic) * 2) / 5)
  st = ((1 / Kbps) * 1000) * FMUSIC_GetSpeed(tmpMusic)

  For i = 0 To (FMUSIC_GetNumOrders(tmpMusic) - 1)
    t = t + FMUSIC_GetPatternLength(tmpMusic, i)
  Next i

  GetMusicLength = t * st
  
  FMUSIC_FreeSong tmpMusic
  
End Function

Public Function GetMusicOK(ByVal filename As String) As Boolean
  'This opens an music file and checks if fmod recognises it
  Dim tmpMusic As Long
  On Error GoTo errh
  
  tmpMusic = FMUSIC_LoadSong(filename)
  If FMUSIC_GetType(tmpMusic) <> 0 Then GetMusicOK = True
  FMUSIC_FreeSong tmpMusic
  
errh:
End Function

Public Function GetMusicName(ByVal filename As String) As String
  'This opens an music file, gets it's name, and closes it
  Dim tmpMusic As Long
  On Error GoTo errh
  
  tmpMusic = FMUSIC_LoadSong(filename)
  If FMUSIC_GetType(tmpMusic) = 0 Then
    FMUSIC_FreeSong tmpMusic
    Exit Function
  End If
  GetMusicName = GetStringFromPointer(FMUSIC_GetName(tmpMusic))
  FMUSIC_FreeSong tmpMusic
  
  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "GetMusicName - FMOD Error"
End Function

Public Function MusicPlay(ByVal sFilename As String, ByVal loopit As Boolean, Optional bVolume As Byte = 255) As Boolean
  'Will play the filename music
  On Error GoTo errh
  MusicPlay = True

  If module <> 0 Then
    FMUSIC_StopSong module
    FMUSIC_FreeSong module
  End If
  module = FMUSIC_LoadSong(sFilename)
  
  If module = 0 Then
    MsgBox "Could not recognize & play " & Chr(34) & sFilename & Chr(34) & ".", vbExclamation, "FMOD Error"
    MusicPlay = False
    Exit Function
  End If
  
  If Not loopit Then Call FMUSIC_SetLooping(module, False)
  Call FMUSIC_PlaySong(module)
  FMUSIC_SetMasterVolume module, bVolume

  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "MusicPlay - FMOD Error"
  MusicPlay = False
End Function

Public Function MusicStop() As Boolean
  'this will stop and resume music
  On Error GoTo errh

  If module <> 0 Then
    MusicStop = True
    If FMUSIC_IsPlaying(module) Then
      FMUSIC_StopSong module
    Else
      FMUSIC_PlaySong module
    End If
  End If
    

  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "MusicStop - FMOD Error"
  MusicStop = False
End Function

Public Function MusicUnload() As Boolean
  'this will stop and unload music
  On Error GoTo errh

  If module <> 0 Then
    MusicUnload = True
    FMUSIC_StopSong module
    FMUSIC_FreeSong module
    module = 0
  End If
    

  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "MusicUnload - FMOD Error"
  MusicUnload = False
End Function

Public Function MusicPausePlay() As Boolean
  'this will stop and unload music
  On Error GoTo errh

  If module <> 0 Then
    MusicPausePlay = True

    FMUSIC_SetPaused module, Not FMUSIC_GetPaused(module)
  End If
    

  Exit Function
errh:
  MsgBox FSOUND_GetErrorString(FSOUND_GetError), vbCritical, "MusicPausePlay - FMOD Error"
  MusicPausePlay = False
End Function

Public Property Get MusicVolume() As Byte
  MusicVolume = FMUSIC_GetMasterVolume(module)
End Property

Public Property Let MusicVolume(ByVal bVolume As Byte)
  FMUSIC_SetMasterVolume module, bVolume
End Property

Public Property Get MusicName() As String
  MusicName = GetStringFromPointer(FMUSIC_GetName(module))
End Property

Public Property Get MusicNumChannels() As Long
  MusicNumChannels = FMUSIC_GetNumChannels(module)
End Property

Public Property Get MusicNumIntruments() As Long
  MusicNumIntruments = FMUSIC_GetNumInstruments(module)
End Property

Public Property Get MusicNumOrders() As Long
  MusicNumOrders = FMUSIC_GetNumOrders(module)
End Property

Public Property Get MusicNumPatterns() As Long
  MusicNumPatterns = FMUSIC_GetNumPatterns(module)
End Property

Public Property Get MusicNumSamples() As Long
  MusicNumSamples = FMUSIC_GetNumSamples(module)
End Property

Public Property Get MusicBPM() As Long
  MusicBPM = FMUSIC_GetBPM(module)
End Property

Public Property Get MusicIsLoaded() As Boolean
  If module <> 0 Then MusicIsLoaded = True
End Property

Public Property Get MusicOrder() As Long
  MusicOrder = FMUSIC_GetOrder(module)
End Property

Public Property Let MusicOrder(ByVal lNewOrder As Long)
  FMUSIC_SetOrder module, lNewOrder
End Property

Public Property Get MusicPattern() As Long
  MusicPattern = FMUSIC_GetPattern(module)
End Property

Public Property Get MusicRow() As Long
  MusicRow = FMUSIC_GetRow(module)
End Property

Public Property Get MusicSpeed() As Long
  MusicSpeed = FMUSIC_GetSpeed(module)
End Property

Public Property Get MusicTime() As Long
  MusicTime = FMUSIC_GetTime(module)
End Property

Public Property Get MusicIsPlaying() As Boolean
  If FMUSIC_GetPaused(module) = False Then
    MusicIsPlaying = FMUSIC_IsPlaying(module)
  End If
End Property

Public Property Get MusicIsFinished() As Boolean
  MusicIsFinished = FMUSIC_IsFinished(module)
End Property

Public Property Get MusicNumRows() As Long
  MusicNumRows = FMUSIC_GetPatternLength(module, FMUSIC_GetOrder(module))
End Property

Public Function MusicPanSeperation(ByVal sPanSep As Single)
  MusicPanSeperation = FMUSIC_SetPanSeperation(module, sPanSep)
End Function

Public Property Get CPUUsage() As Single
  CPUUsage = FSOUND_GetCPUUsage
End Property

Public Property Get FMODVersion() As Single
  FMODVersion = FSOUND_GetVersion
End Property

Public Property Get StreamKbps() As Long
  'This calculates Kbps by dividing length of stream in bits with length of stream in seconds.
  On Error Resume Next
  StreamKbps = (FSOUND_Stream_GetLength(stream) / (FSOUND_Stream_GetLengthMs(stream) / 1000)) * 8 / 1000
End Property

Public Property Get StreamPanning() As Byte
  StreamPanning = FSOUND_GetPan(channel)
End Property

Public Property Let StreamPanning(ByVal bPanning As Byte)
  If bPanning = 128 Then
    FSOUND_SetPan FSOUND_ALL, FSOUND_STEREOPAN
  Else
    FSOUND_SetPan FSOUND_ALL, bPanning
  End If
End Property

Public Property Get StreamSurround() As Boolean
  StreamSurround = FSOUND_GetSurround(channel)
End Property

Public Property Let StreamSurround(ByVal surround As Boolean)
  Call FSOUND_SetSurround(FSOUND_ALL, surround)
End Property

Public Property Get MusicSongLen() As Long
  'This will calculate the songs length in ms
  'This is mostly accurate, but can be highly inaccurate
  On Error Resume Next
  Dim Kbps As Double, st As Double, i As Long, t As Long
  
  Kbps = ((FMUSIC_GetBPM(module) * 2) / 5)
  st = ((1 / Kbps) * 1000) * FMUSIC_GetSpeed(module)

  For i = 0 To (FMUSIC_GetNumOrders(module) - 1)
    t = t + FMUSIC_GetPatternLength(module, i)
  Next i

  MusicSongLen = t * st
End Property

Public Property Get MusicSongPos() As Long
  'This will calculate the current song time in ms
  'This is mostly accurate, but can be highly inaccurate
  On Error Resume Next
  Dim Kbps As Double, st As Double, i As Long, t As Long
  
  Kbps = ((FMUSIC_GetBPM(module) * 2) / 5)
  st = ((1 / Kbps) * 1000) * FMUSIC_GetSpeed(module)

  For i = 0 To (FMUSIC_GetOrder(module))
    t = t + FMUSIC_GetPatternLength(module, i)
  Next i
  t = t - (FMUSIC_GetPatternLength(module, FMUSIC_GetOrder(module)) - FMUSIC_GetRow(module))
  
  MusicSongPos = t * st
End Property

Public Function MusicGetRows(ByVal order As Long) As Long
  MusicGetRows = FMUSIC_GetPatternLength(module, order)
End Function

Public Property Get ChannelsPlaying() As Long
  ChannelsPlaying = FSOUND_GetChannelsPlaying
End Property

Public Property Let SpeakerSetup(ByVal Spk As FSOUND_SPEAKERMODES)
  FSOUND_SetSpeakerMode Spk
End Property

Public Property Let MusicPanSep(ByVal pan As Single)
  FMUSIC_SetPanSeperation module, pan
End Property

Public Sub GetVU(ByRef Left As Single, ByRef Right As Single)
  Call FSOUND_GetCurrentLevels(channel, Left, Right)
End Sub

Public Sub GetSpecificVU(ByVal lChannel, ByRef Left As Single, ByRef Right As Single)
  Call FSOUND_GetCurrentLevels(lChannel, Left, Right)
End Sub

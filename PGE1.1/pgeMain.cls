VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Main Class 1.1
'------------------------------
'PGE GLOBAL VERSION: 1.1
'Created by Paul Berlin 2002-2003
'berlin_paul@hotmail.com
'http://pab.dyndns.org
'
'Using the FMOD sound system.
'http://www.fmod.org/
'FMOD is only free to use for non-commercial products!
'
'Parts of this code is based on the Boom2D DirectX engine,
'written by Bart van de Sande, also available on PlanetSourceCode.com:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=38321&lngWId=1
'
'Feel free to change/add/improve this code, but if you do
'I would appreciate if you send me the updated code =).
'berlin_paul@hotmail.com
'
'If you use this engine in any of your own projects, please
'mention me in your credits.
'
'OVERVIEW:
'---------
'* Use Init(hWnd, bWindowed, w, h, bVSync) to init the engine at the specified
'  window.
'* Use Clear() to clear the rendering area before rendering. This must be called
'  before being able do render anything.
'* Use Render(hWnd) to render the backbuffer to the window.
'* Use lFPS to get the current fps. This is unaccurate at high fps's
'  (around 1000fps), but more accurate at lower speeds.
'* Use lClearColor to set the color to clear the screen to
'
'CHANGELOG:
'----------
'1.1 06-06-2003
'* Fixed some minor things.
'* Updated fps meter to an better one.
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private mStart As Long, mTime(101) As Long 'fps calc

Public lFPS As Long 'The number of calculated frames per second
Public lClearColor As Long 'The color to clear the backbuffer with

Public Sub Init(ByVal hwnd As Long, ByVal bWindowed As Boolean, Optional w As Long = -1, Optional h As Long = -1, Optional bVSync As Boolean = True)
  'This intializes the engine.
  'NOTES:
  '* The Default adapter is used.
  '* D3DDEVTYPE_HAL is used as device type
  '-
  '* w & h must be entered when using fullscreen
  '
  'hWnd = Window Handle
  'bWindowed = Fullscreen or window
  'w = fullscreen width
  'h = fullscreen height
  'bVSync = vertical sync on/off
  
  World.sFadeAlpha = 100
  World.sFadeRed = 100
  World.sFadeGreen = 100
  World.sFadeBlue = 100
  
  Dim params As D3DPRESENT_PARAMETERS
  Dim dp As D3DDISPLAYMODE
  
  'Create Direct3D
  Set Direct3D = DirectX.Direct3DCreate
  
  'retrieve the display mode
  Direct3D.GetAdapterDisplayMode 0, dp

  If bWindowed Then 'If windowed
    'Gets render area size
    GetWindowRect hwnd, Target
    With Target
      w = .Right - .Left
      h = .bottom - .Top
      .Left = 0
      .Right = w
      .Top = 0
      .bottom = h
    End With
    
    'Sets up render parameters
    With params
      .BackBufferFormat = dp.Format
      .EnableAutoDepthStencil = 0
      .Windowed = 1
      .SwapEffect = IIf(bVSync, D3DSWAPEFFECT_COPY_VSYNC, D3DSWAPEFFECT_COPY)
    End With
  Else 'If fullscreen
    'Setup screen size
    With Target
      Target.Left = 0
      Target.Right = w
      Target.Top = 0
      Target.bottom = h
    End With
    
    'Sets up render parameters
    With params
      .BackBufferFormat = dp.Format
      .EnableAutoDepthStencil = 0
      .BackBufferWidth = w
      .BackBufferHeight = h
      .Windowed = 0
      .SwapEffect = D3DSWAPEFFECT_FLIP
    End With
  End If
  
  'Creates device
  Set Direct3DDevice = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, params)
  
  'Inits sprites
  Set Sprites = Direct3DX.CreateSprite(Direct3DDevice)
  
  'Init Input
  Set DirectInput = DirectX.DirectInputCreate
  
  'Turn off culling
  Direct3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
  'Turn off D3D lighting, since we are providing our own vertex colors
  Direct3DDevice.SetRenderState D3DRS_LIGHTING, 0
End Sub

Public Sub Clear()
  'This clears the backbuffer
  Direct3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, lClearColor, 1, 0
  Direct3DDevice.BeginScene
  Sprites.Begin
End Sub

Public Sub Render(Optional hwnd As Long)
  'This renders the backbuffer to the render area
  'hWnd = Window handle
  On Error Resume Next
  
  'Calculate fps
  Dim ns As Long, vs As Long
  For ns = 0 To 100
    mTime(ns) = mTime(ns + 1)
    vs = vs + mTime(ns)
  Next ns
  mTime(ns) = timeGetTime - mStart
  mStart = timeGetTime
  lFPS = 100000 / (vs + mTime(ns))
  
  'presents the rendered image
  Sprites.End
  Direct3DDevice.EndScene
  Direct3DDevice.Present ByVal 0, ByVal 0, hwnd, ByVal 0
End Sub

Private Sub Class_Terminate()
  On Error Resume Next
  
  Set Sprites = Nothing
  Set Direct3DDevice = Nothing
  Set Direct3D = Nothing
  Set DirectInput = Nothing
End Sub



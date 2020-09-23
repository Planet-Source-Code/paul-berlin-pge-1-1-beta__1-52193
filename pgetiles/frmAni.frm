VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAni 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Animation"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   591
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3735
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
      Begin VB.Frame Frame6 
         Caption         =   "Subframe"
         Height          =   1215
         Left            =   0
         TabIndex        =   20
         Top             =   840
         Width           =   2295
         Begin VB.TextBox txtO 
            Height          =   315
            Index           =   1
            Left            =   1440
            TabIndex        =   25
            Text            =   "0"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtO 
            Height          =   315
            Index           =   0
            Left            =   480
            TabIndex        =   24
            Text            =   "0"
            Top             =   720
            Width           =   495
         End
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   1440
            TabIndex        =   21
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "imlToolbarIcons(1)"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "prev"
                  Object.ToolTipText     =   "Previous subframe"
                  ImageKey        =   "prev"
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "next"
                  Object.ToolTipText     =   "Next subframe"
                  ImageKey        =   "next"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Offset:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   465
         End
         Begin VB.Label lblP 
            AutoSize        =   -1  'True
            Caption         =   "Y:"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   27
            Top             =   795
            Width           =   150
         End
         Begin VB.Label lblP 
            AutoSize        =   -1  'True
            Caption         =   "X:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   795
            Width           =   150
         End
      End
      Begin VB.CheckBox chkDelay 
         Caption         =   "Apply this delay to all frames"
         Height          =   255
         Left            =   0
         TabIndex        =   19
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtD 
         Height          =   315
         Left            =   720
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
         Caption         =   "ms"
         Height          =   195
         Index           =   7
         Left            =   1680
         TabIndex        =   10
         Top             =   240
         Width           =   195
      End
      Begin VB.Label lblP 
         AutoSize        =   -1  'True
         Caption         =   "Delay:"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   165
         Width           =   450
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3735
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
      Visible         =   0   'False
      Begin VB.CheckBox chkSub 
         Caption         =   "Change subframe each frame."
         Height          =   495
         Left            =   0
         TabIndex        =   23
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Frame Frame5 
         Caption         =   "View"
         Height          =   1215
         Left            =   0
         TabIndex        =   15
         Top             =   1200
         Width           =   2295
         Begin MSComctlLib.Slider sldZoom 
            Height          =   615
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   393216
            Min             =   1
            Max             =   6
            SelStart        =   1
            TickStyle       =   2
            Value           =   1
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Zoom:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   450
         End
         Begin VB.Label lblZoom 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1x"
            Height          =   195
            Left            =   2010
            TabIndex        =   17
            Top             =   240
            Width           =   165
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Slowdown"
         Height          =   1095
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   2295
         Begin MSComctlLib.Slider sldSlow 
            Height          =   495
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   873
            _Version        =   393216
            Min             =   10
            Max             =   30
            SelStart        =   10
            TickFrequency   =   5
            Value           =   10
         End
         Begin VB.Label lblSlow 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "x1 slowdown"
            Height          =   195
            Left            =   1260
            TabIndex        =   14
            Top             =   240
            Width           =   915
         End
      End
   End
   Begin MSComctlLib.TabStrip tabs 
      Height          =   4215
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Control"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Misc"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Playback"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2535
      Begin VB.Label lblSub 
         AutoSize        =   -1  'True
         Caption         =   "Subframe: 1/1"
         Height          =   195
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Offsets: 0, 0"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delay: 100ms"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   960
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame: 1/1"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.PictureBox pDX 
      DrawMode        =   7  'Invert
      Height          =   5535
      Left            =   2760
      MousePointer    =   15  'Size All
      ScaleHeight     =   365
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin MSComctlLib.Toolbar tbrPlayback 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlToolbarIcons(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pbegin"
            Object.ToolTipText     =   "Jump to first frame"
            ImageKey        =   "pbegin"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pprev"
            Object.ToolTipText     =   "Step back one frame"
            ImageKey        =   "pprev"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ppause"
            Object.ToolTipText     =   "Pause"
            ImageKey        =   "ppause"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pplay"
            Object.ToolTipText     =   "Play"
            ImageKey        =   "pplay"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pnext"
            Object.ToolTipText     =   "Step forward one frame"
            ImageKey        =   "pnext"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pend"
            Object.ToolTipText     =   "Jump to last frame"
            ImageKey        =   "pend"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            ImageKey        =   "help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   0
      Left            =   120
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8421376
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAni.frx":0000
            Key             =   "pbegin"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAni.frx":0112
            Key             =   "pprev"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAni.frx":0224
            Key             =   "ppause"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAni.frx":0336
            Key             =   "pplay"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAni.frx":0448
            Key             =   "pnext"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAni.frx":055A
            Key             =   "pend"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAni.frx":066C
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Timer cptimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2040
      Top             =   1680
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   1
      Left            =   3825
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8421376
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAni.frx":077E
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAni.frx":0890
            Key             =   "next"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Type tTex
  img As Direct3DTexture8
  name As String
End Type

'DirectX
Private DirectX As New DirectX8
Private Direct3D As Direct3D8
Private Direct3DDevice As Direct3DDevice8
Private Direct3DX As New D3DX8
Private Sprites As D3DXSprite
Private tex() As tTex
Private texBG As Direct3DTexture8
Private texLoaded As Boolean 'textures are loaded
Private DXOK As Boolean 'DirectX ok

'other
Private bMd As Boolean
Private vecMove As D3DVECTOR2
Private texVec As D3DVECTOR2 'image offset
Private lCurFrame As Long
Private lCurSubFrame As Long
Private lFrmCol As Long
Private lFrmAni As Long
Private bPlay As Boolean
Private bEnd As Boolean

Private Sub chkSub_Click()
  Prg.bAniAutochange = CBool(chkSub.Value)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
  On Error Resume Next
  Select Case Button.Key
    Case "prev"
      lCurSubFrame = lCurSubFrame - 1
      If lCurSubFrame < 1 Then lCurSubFrame = 1
    Case "next"
      lCurSubFrame = lCurSubFrame + 1
      If lCurSubFrame > Prj.Coll(lFrmCol).Anim(lFrmAni).lSubimages Then lCurSubFrame = Prj.Coll(lFrmCol).Anim(lFrmAni).lSubimages
  End Select
  RefreshInfo
End Sub

Private Sub cptimer_Timer()
  'On Error Resume Next
  InitDX
  DrawIt 'main loooop
  Set Sprites = Nothing
  Set Direct3DDevice = Nothing
  Set Direct3D = Nothing
  Set texBG = Nothing
  Dim x As Long
  For x = 1 To UBound(tex)
    Set tex(x).img = Nothing
  Next
  Unload Me
End Sub

Private Sub Form_Load()
  ReDim tex(0)
  bEnd = False
  texVec.x = 75: texVec.Y = 75
  lFrmCol = frmMain.lSelCol
  lFrmAni = frmMain.lSelAni
  chkSub.Value = Abs(Prg.bAniAutochange)
  With Prj.Coll(lFrmCol).Anim(lFrmAni)
    Me.Caption = "Edit Frames [" & .sID & "]"
    If UBound(.Frame) < 1 Then
      MsgBox "This animation has no frames to set up. Please go to the 'Edit Frames' window and change this before ever coming back again! bastard.", vbExclamation, "Missing frames"
    Else
      lCurFrame = 1
      lCurSubFrame = 1
    End If
  End With
  RefreshInfo
  cptimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not bEnd Then
    Cancel = 1
    bEnd = True
  End If
End Sub

Private Sub sldSlow_Scroll()
  lblSlow = "x" & sldSlow.Value / 10 & " slowdown"
End Sub

Private Sub tabs_Click()
  Dim x As Long
  For x = 0 To 1
    If x <> tabs.SelectedItem.Index - 1 Then
      frm(x).Visible = False
    Else
      frm(x).Visible = True
    End If
  Next
End Sub

Private Sub pDX_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  bMd = True
  vecMove.x = x - texVec.x
  vecMove.Y = Y - texVec.Y
End Sub

Private Sub pDX_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If bMd And texLoaded Then
    texVec.x = x - vecMove.x
    texVec.Y = Y - vecMove.Y
  End If
End Sub

Private Sub pDX_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  bMd = False
End Sub

Private Sub tbrPlayback_ButtonClick(ByVal Button As MSComCtlLib.Button)
  On Error Resume Next
  If UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame) = 0 And Button.Key <> "help" Then Exit Sub
  Select Case Button.Key
    Case "pbegin"
      bPlay = False
      lCurFrame = 1
    Case "pprev"
      bPlay = False
      lCurFrame = lCurFrame - 1
      If lCurFrame = 0 Then lCurFrame = 1
    Case "ppause"
      bPlay = False
    Case "pplay"
      bPlay = True
    Case "pnext"
      bPlay = False
      lCurFrame = lCurFrame + 1
      If lCurFrame > UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame) Then lCurFrame = UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame)
    Case "pend"
      bPlay = False
      lCurFrame = UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame)
    Case "help"
      frmHelp.Show , Me
      frmHelp.web.Navigate App.Path & "\doc\ani.html"
  End Select
  RefreshInfo
End Sub

Public Sub InitDX()
  On Error GoTo errh
  Dim params As D3DPRESENT_PARAMETERS
  Dim dp As D3DDISPLAYMODE
  
  Set Direct3D = DirectX.Direct3DCreate
  Direct3D.GetAdapterDisplayMode 0, dp

  With params
    .BackBufferFormat = dp.Format
    .EnableAutoDepthStencil = 0
    .Windowed = 1
    .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
  End With
  
  Set Direct3DDevice = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, pDX.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, params)
  Set Sprites = Direct3DX.CreateSprite(Direct3DDevice)
  
  Set texBG = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, App.Path & "\back.png", -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, 0, ByVal 0, ByVal 0)
  
  Direct3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
  Direct3DDevice.SetRenderState D3DRS_LIGHTING, 0
  With Prj.Coll(lFrmCol).Anim(lFrmAni)
    If lCurFrame > 0 Then
      texLoaded = True
      Dim x As Long, Y As Long
      For x = 1 To UBound(.Frame)
        For Y = 1 To .lSubimages
          addTex Prj.Res(.Frame(x).img(Y).lRes).sFilename, Prj.Res(.Frame(x).img(Y).lRes).lTranscolor
        Next
      Next
    End If
  End With
  
  DXOK = True
  'pDX_Paint
  Exit Sub
errh:
  MsgBox "It seems that DirectX 8 3D H&L support is nowhere to be found... DirectX init failed!", vbCritical, "DX Error"
End Sub

Public Sub addTex(ByVal sF As String, ByVal t As Long)
  Dim x As Long
  For x = 1 To UBound(tex)
    If tex(x).name = LCase(sF) Then Exit Sub
  Next
  ReDim Preserve tex(UBound(tex) + 1)
  With tex(UBound(tex))
    .name = LCase(sF)
    If t <> -1 Then
      Set .img = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, .name, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, D3DColorRGBA(t Mod &H100, (t \ &H100) Mod &H100, (t \ &H10000) Mod &H100, 255), ByVal 0, ByVal 0)
    Else
      Set .img = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, .name, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, 0, ByVal 0, ByVal 0)
    End If
  End With
End Sub

Public Function getTex(ByVal sF As String) As Long
  Dim x As Long
  For x = 1 To UBound(tex)
    If tex(x).name = LCase(sF) Then
      getTex = x
      Exit Function
    End If
  Next
End Function

Public Sub DrawIt()
  Dim t As Long
  
  Do
    DoEvents
    If DXOK Then
      Direct3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0
      Direct3DDevice.BeginScene
      Sprites.Begin
      
      Dim x As Long, Y As Long
      For x = 0 To pDX.Width Step 128
        For Y = 0 To pDX.Height Step 128
          Sprites.Draw texBG, ByVal 0, vec2(1, 1), vec2(0, 0), 0, vec2(x, Y), &HFFFFFFFF
        Next
      Next
      
      If texLoaded Then
        With Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame)
          x = getTex(Prj.Res(.img(lCurSubFrame).lRes).sFilename)
          If (Prj.Coll(lFrmCol).Anim(lFrmAni).eOptions And Center_8) = Center_8 Then
            Sprites.Draw tex(x).img, .img(lCurSubFrame).Pos, vec2(sldZoom.Value, sldZoom.Value), vec2(0, 0), 0, vec2(texVec.x + .img(lCurSubFrame).Offset.x - ((.img(lCurSubFrame).Pos.Right - .img(lCurSubFrame).Pos.Left) / 2) * sldZoom.Value, texVec.Y + .img(lCurSubFrame).Offset.Y - ((.img(lCurSubFrame).Pos.bottom - .img(lCurSubFrame).Pos.Top) / 2) * sldZoom.Value), D3DColorRGBA(255, 255, 255, 255)
          Else
            Sprites.Draw tex(x).img, .img(lCurSubFrame).Pos, vec2(sldZoom.Value, sldZoom.Value), vec2(0, 0), 0, vec2(texVec.x + .img(lCurSubFrame).Offset.x, texVec.Y + .img(lCurSubFrame).Offset.Y), D3DColorRGBA(255, 255, 255, 255)
          End If
        End With
      End If
      
      If bPlay Then
        If timeGetTime - t > Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).Ctrl.lDelay * sldSlow.Value / 10 Then
          lCurFrame = lCurFrame + 1
          If lCurFrame > UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame) Then
            If (Prj.Coll(lFrmCol).Anim(lFrmAni).eOptions And Loop_4) = Loop_4 Then
              lCurFrame = 1
            Else
              lCurFrame = UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame)
              bPlay = False
            End If
          End If
          t = timeGetTime
          If chkSub.Value = vbChecked Then
            lCurSubFrame = lCurSubFrame + 1
            If lCurSubFrame > Prj.Coll(lFrmCol).Anim(lFrmAni).lSubimages Then lCurSubFrame = 1
          End If
        End If
        RefreshInfo
      Else
        t = timeGetTime
      End If
  
      
      Sprites.End
      Direct3DDevice.EndScene
      Direct3DDevice.Present ByVal 0, ByVal 0, pDX.hwnd, ByVal 0
      
    End If
  Loop Until bEnd
  
End Sub

Public Sub RefreshInfo()
  With Prj.Coll(lFrmCol).Anim(lFrmAni)
    Dim x As Long
    lblInfo(0) = "Frame: " & lCurFrame & "/" & UBound(.Frame)
    lblInfo(1) = "Delay: " & .Frame(lCurFrame).Ctrl.lDelay & " ms"
    lblInfo(2) = "Offsets: " & .Frame(lCurFrame).img(lCurSubFrame).Offset.x & ", " & .Frame(lCurFrame).img(lCurSubFrame).Offset.Y
    lblSub = "Subframe: " & lCurSubFrame & "/" & .lSubimages
    
    If .lSubimages <= 1 Then
      Toolbar1.Buttons(1).Enabled = False
      Toolbar1.Buttons(2).Enabled = False
      chkSub.Enabled = False
    Else
      Toolbar1.Buttons(1).Enabled = True
      Toolbar1.Buttons(2).Enabled = True
      chkSub.Enabled = True
    End If
    
    
    If bPlay Then
      For x = 0 To 1
        lblP(x).Enabled = False
        txtO(x).Enabled = False
      Next
      txtD.Enabled = False
      chkDelay.Enabled = False
    Else
      For x = 0 To 1
        lblP(x).Enabled = True
        txtO(x).Enabled = True
      Next
      txtD.Enabled = True
      txtD.Text = .Frame(lCurFrame).Ctrl.lDelay
      txtO(0).Text = .Frame(lCurFrame).img(lCurSubFrame).Offset.x
      txtO(1).Text = .Frame(lCurFrame).img(lCurSubFrame).Offset.Y
      chkDelay.Enabled = True
    End If
  End With
End Sub

Private Sub txtD_KeyPress(KeyAscii As Integer)
  KeyAscii = ctlKeyPress(KeyAscii, NumbersOnly)
End Sub

Private Sub txtD_KeyUp(KeyCode As Integer, Shift As Integer)
  If chkDelay Then
    Dim x As Long
    For x = 1 To UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame)
      Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(x).Ctrl.lDelay = Val(txtD)
    Next
  Else
    Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).Ctrl.lDelay = Val(txtD)
  End If
  Prg.bChanged = True
  RefreshInfo
End Sub

Private Sub txtO_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = ctlKeyPress(KeyAscii, NumbersOnly Or AllowNegative)
End Sub

Private Sub txtO_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).Offset.x = Val(txtO(0))
  Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).Offset.Y = Val(txtO(1))
  Prg.bChanged = True
  RefreshInfo
End Sub

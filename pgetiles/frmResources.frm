VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResources 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resources"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pDX 
      Height          =   4695
      Left            =   2400
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   6
      Top             =   120
      Width           =   5535
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   582
      ButtonWidth     =   609
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Object.ToolTipText     =   "Add"
            ImageKey        =   "addres"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "remres"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "help"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reset Scroll"
            Object.ToolTipText     =   "Reset Scroll"
            ImageKey        =   "reset"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transparency"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   2175
      Begin VB.OptionButton optTrans 
         Caption         =   "2. Use Color Value:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton optTrans 
         Caption         =   "1. Use Alpha Channel"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 0 0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   750
         Width           =   1695
      End
   End
   Begin VB.ListBox lstRes 
      Height          =   3180
      ItemData        =   "frmResources.frx":0000
      Left            =   120
      List            =   "frmResources.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8421376
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":0004
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":0116
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":0228
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":033A
            Key             =   "Tab Left"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":044C
            Key             =   "TAB-LEFT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":055E
            Key             =   "addres"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":0670
            Key             =   "remres"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":0782
            Key             =   "help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmResources.frx":0894
            Key             =   "reset"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmResources"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private DirectX As New DirectX8
Private Direct3D As Direct3D8
Private Direct3DDevice As Direct3DDevice8
Private Direct3DX As New D3DX8
Private Sprites As D3DXSprite
Private texBG As Direct3DTexture8
Private texImg As Direct3DTexture8
Private texVec As D3DVECTOR2
Private texLoaded As Boolean
Private texInfo As D3DXIMAGE_INFO
Private bMd As Boolean
Private vecMove As D3DVECTOR2
Private DXOK As Boolean

Private Sub Form_Unload(Cancel As Integer)

  Set Sprites = Nothing
  Set texBG = Nothing
  Set Direct3DDevice = Nothing
  Set Direct3D = Nothing
  Set texImg = Nothing
End Sub

Private Sub optTrans_Click(Index As Integer)
  With Prj.Res(lstRes.ListIndex + 1)
    Select Case Index
      Case 0
        .lTranscolor = -1
        lblColor.Enabled = False
      Case 1
        lblColor.Enabled = True
        .lTranscolor = lblColor.BackColor
        lblColor.Caption = .lTranscolor Mod &H100 & " " & (.lTranscolor \ &H100) Mod &H100 & " " & (.lTranscolor \ &H10000) Mod &H100
        lblColor.ForeColor = RGB(255 - (.lTranscolor Mod &H100), 255 - ((.lTranscolor \ &H100) Mod &H100), 255 - ((.lTranscolor \ &H10000) Mod &H100))
    End Select
    If .lTranscolor = -1 Then
      Set texImg = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, .sFilename, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, 0, texInfo, ByVal 0)
    Else
      Set texImg = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, .sFilename, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, D3DColorRGBA(.lTranscolor Mod &H100, (.lTranscolor \ &H100) Mod &H100, (.lTranscolor \ &H10000) Mod &H100, 255), texInfo, ByVal 0)
    End If
  End With
  Prg.bChanged = True
  bMd = False
  pDX_Paint
End Sub

Private Sub pDX_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bMd = True
  vecMove.X = X - texVec.X
  vecMove.Y = Y - texVec.Y
End Sub

Private Sub pDX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If bMd And texLoaded Then
    texVec.X = X - vecMove.X
    texVec.Y = Y - vecMove.Y
    pDX_Paint
  End If
End Sub

Private Sub pDX_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  bMd = False
End Sub

Private Sub pDX_Paint()
  If DXOK Then
    Direct3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1, 0
    Direct3DDevice.BeginScene
    Sprites.Begin
    
    Dim X As Long, Y As Long
    For X = 0 To pDX.Width Step 128
      For Y = 0 To pDX.Height Step 128
        Sprites.Draw texBG, ByVal 0, vec2(1, 1), vec2(0, 0), 0, vec2(X, Y), &HFFFFFFFF
      Next
    Next
    If texLoaded Then
      Sprites.Draw texImg, ByVal 0, vec2(1, 1), vec2(0, 0), 0, texVec, &HFFFFFFFF
    End If
    
    Sprites.End
    Direct3DDevice.EndScene
    Direct3DDevice.Present ByVal 0, ByVal 0, pDX.hwnd, ByVal 0
  End If
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComCtlLib.Button)
  Dim X As Long
  Select Case Button.Key
    Case "Reset Scroll"
      texVec.X = 0
      texVec.Y = 0
      pDX_Paint
    Case "Help"
      frmHelp.Show , Me
      frmHelp.web.Navigate App.Path & "\doc\res.html"
    Case "Add"
      On Error GoTo errh
      With frmMain.cdgMain
        .FileName = ""
        .DialogTitle = "Load Resource"
        .Filter = "PNG-Images|*.png|Other supported formats|*.gif;*.bmp;*.tif;*.tiff;*.tga;*.jpg;*.jpeg;*.jpe|All files|*.*"
        .flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNPathMustExist Or cdlOFNHideReadOnly Or cdlOFNAllowMultiselect
        .ShowOpen
        Dim fls() As String
        
        If InStr(1, .FileName, Chr(0)) Then
          fls = Split(.FileName, Chr(0))
          fls(0) = fls(0) & "\"
        Else
          ReDim fls(1)
          fls(0) = Left(.FileName, Len(.FileName) - InStrRev(.FileName, "\"))
          fls(1) = Right(.FileName, Len(.FileName) - Len(fls(0)))
        End If
        For X = 1 To UBound(fls)
          ReDim Preserve Prj.Res(UBound(Prj.Res) + 1)
          Prj.Res(UBound(Prj.Res)).sFilename = fls(0) & fls(X)
          Prj.Res(UBound(Prj.Res)).lTranscolor = -1
          
          lstRes.AddItem UBound(Prj.Res) & ". " & Right(fls(0) & fls(X), Len(fls(0) & fls(X)) - InStrRev(fls(0) & fls(X), "\"))
        
        Next
        
        lstRes.ListIndex = UBound(Prj.Res) - 1
        Prg.bChanged = True
        
      End With
      On Error GoTo 0
    Case "Delete"
      If Prg.bConfirm(3) Then
        X = MsgBox("Are you sure you want to remove this resource from the tileset? Doing so will invalidate all animations that use this resource!", vbExclamation Or vbYesNo, "Confirmation")
      Else
        X = vbYes
      End If
      If X = vbYes Then
        Dim Y As Long, z As Long, w As Long
        
        For X = lstRes.ListIndex + 1 To UBound(Prj.Res) - 1
          With Prj.Res(X)
            .lTranscolor = Prj.Res(X + 1).lTranscolor
            .sFilename = Prj.Res(X + 1).sFilename
          End With
        Next
    
        For X = 1 To UBound(Prj.Coll)
          For Y = 1 To UBound(Prj.Coll(X).Anim)
            For z = 1 To UBound(Prj.Coll(X).Anim(Y).Frame)
              For w = 1 To Prj.Coll(X).Anim(Y).lSubimages
                If Prj.Coll(X).Anim(Y).Frame(z).img(w).lRes = lstRes.ListIndex + 1 Then Prj.Coll(X).Anim(Y).Frame(z).img(w).lRes = 0
              Next
            Next
          Next
        Next
        
        ReDim Preserve Prj.Res(UBound(Prj.Res) - 1)
        
        Prg.bChanged = True
        lstRes.Clear
        For X = 1 To UBound(Prj.Res)
          lstRes.AddItem X & ". " & Right(Prj.Res(X).sFilename, Len(Prj.Res(X).sFilename) - InStrRev(Prj.Res(X).sFilename, "\"))
        Next
        lstRes_Click
      End If
  End Select
  
errh:
End Sub

Private Sub Form_Load()
  InitDX
  If UBound(Prj.Res) > 0 Then
    Dim X As Long
    For X = 1 To UBound(Prj.Res)
      lstRes.AddItem X & ". " & Right(Prj.Res(X).sFilename, Len(Prj.Res(X).sFilename) - InStrRev(Prj.Res(X).sFilename, "\"))
    Next
    lstRes.ListIndex = 0
  Else
    lstRes_Click
  End If
End Sub

Private Sub lblColor_Click()
  With Prj.Res(lstRes.ListIndex + 1)
    Load frmColor
    frmColor.Tag = .lTranscolor
    frmColor.Show vbModal, Me
    .lTranscolor = frmColor.Tag
    lblColor.BackColor = .lTranscolor
    lblColor.ForeColor = RGB(255 - (.lTranscolor Mod &H100), 255 - ((.lTranscolor \ &H100) Mod &H100), 255 - ((.lTranscolor \ &H10000) Mod &H100))
    lblColor.Caption = .lTranscolor Mod &H100 & " " & (.lTranscolor \ &H100) Mod &H100 & " " & (.lTranscolor \ &H10000) Mod &H100
    frmColor.Tag = "Y"
    Unload frmColor
    Prg.bChanged = True
    If .lTranscolor = -1 Then
      Set texImg = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, .sFilename, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, 0, texInfo, ByVal 0)
    Else
      Set texImg = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, .sFilename, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, D3DColorRGBA(.lTranscolor Mod &H100, (.lTranscolor \ &H100) Mod &H100, (.lTranscolor \ &H10000) Mod &H100, 255), texInfo, ByVal 0)
    End If
    bMd = False
  End With
  pDX_Paint
End Sub

Private Sub lstRes_Click()
  If lstRes.ListIndex <> -1 Then
    optTrans(0).Enabled = True
    optTrans(1).Enabled = True
    With Prj.Res(lstRes.ListIndex + 1)
      If .lTranscolor = -1 Then
        optTrans(0).Value = True
        lblColor.Enabled = False
      Else
        optTrans(1).Value = True
        lblColor.Enabled = True
        lblColor.BackColor = .lTranscolor
        lblColor.Caption = .lTranscolor Mod &H100 & " " & (.lTranscolor \ &H100) Mod &H100 & " " & (.lTranscolor \ &H10000) Mod &H100
        lblColor.ForeColor = RGB(255 - (.lTranscolor Mod &H100), 255 - ((.lTranscolor \ &H100) Mod &H100), 255 - ((.lTranscolor \ &H10000) Mod &H100))
      End If
      If DXOK Then
        If .lTranscolor = -1 Then
          Set texImg = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, .sFilename, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, 0, texInfo, ByVal 0)
        Else
          Set texImg = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, .sFilename, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, D3DColorRGBA(.lTranscolor Mod &H100, (.lTranscolor \ &H100) Mod &H100, (.lTranscolor \ &H10000) Mod &H100, 255), texInfo, ByVal 0)
        End If
        Me.Caption = "Resources (" & texInfo.Width & "x" & texInfo.Height & ")"
        texVec.X = 0
        texVec.Y = 0
        pDX.MousePointer = vbSizeAll
        tbrMain.Buttons(6).Enabled = True
        texLoaded = True
        bMd = False
      End If
    End With
  Else
    Me.Caption = "Resources"
    tbrMain.Buttons(6).Enabled = False
    pDX.MousePointer = vbDefault
    texLoaded = False
    Set texImg = Nothing
    optTrans(0).Enabled = False
    optTrans(1).Enabled = False
    lblColor.Enabled = False
  End If
  pDX_Paint
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
  
  Direct3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
  Direct3DDevice.SetRenderState D3DRS_LIGHTING, 0
  
  'loads bg
  Set texBG = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, App.Path & "\back.png", -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, 0, ByVal 0, ByVal 0)
  
  texLoaded = False
  DXOK = True
  pDX_Paint
  Exit Sub
errh:
  MsgBox "It seems that DirectX 8 3D H&L support is nowhere to be found... DirectX init failed!", vbCritical, "DX Error"
End Sub

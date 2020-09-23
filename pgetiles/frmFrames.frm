VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFrames 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Edit Frames"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
      Begin VB.CheckBox chkShowInView 
         Caption         =   "Show in view"
         Height          =   255
         Left            =   0
         TabIndex        =   69
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Virtual Size"
         Height          =   1575
         Left            =   0
         TabIndex        =   13
         Top             =   2040
         Width           =   2415
         Begin VB.CommandButton cmdVSet 
            Caption         =   "Set from Selection"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtV 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   30
            Top             =   675
            Width           =   495
         End
         Begin VB.TextBox txtV 
            Height          =   315
            Index           =   2
            Left            =   360
            TabIndex        =   28
            Top             =   675
            Width           =   495
         End
         Begin VB.TextBox txtV 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   26
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtV 
            Height          =   315
            Index           =   0
            Left            =   360
            TabIndex        =   24
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblV 
            AutoSize        =   -1  'True
            Caption         =   "H:"
            Height          =   195
            Index           =   3
            Left            =   1080
            TabIndex        =   29
            Top             =   750
            Width           =   165
         End
         Begin VB.Label lblV 
            AutoSize        =   -1  'True
            Caption         =   "W:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   750
            Width           =   210
         End
         Begin VB.Label lblV 
            AutoSize        =   -1  'True
            Caption         =   "Y:"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   25
            Top             =   315
            Width           =   150
         End
         Begin VB.Label lblV 
            AutoSize        =   -1  'True
            Caption         =   "X:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   315
            Width           =   150
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Position && Dimension"
         Height          =   1575
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   2415
         Begin VB.CommandButton cmdPSet 
            Caption         =   "Set from Selection"
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtP 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   21
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtP 
            Height          =   315
            Index           =   2
            Left            =   360
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtP 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   17
            Top             =   280
            Width           =   495
         End
         Begin VB.TextBox txtP 
            Height          =   315
            Index           =   0
            Left            =   360
            TabIndex        =   15
            Top             =   280
            Width           =   495
         End
         Begin VB.Label lblP 
            AutoSize        =   -1  'True
            Caption         =   "H:"
            Height          =   195
            Index           =   3
            Left            =   1080
            TabIndex        =   20
            Top             =   795
            Width           =   165
         End
         Begin VB.Label lblP 
            AutoSize        =   -1  'True
            Caption         =   "W:"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   795
            Width           =   210
         End
         Begin VB.Label lblP 
            AutoSize        =   -1  'True
            Caption         =   "Y:"
            Height          =   195
            Index           =   1
            Left            =   1080
            TabIndex        =   16
            Top             =   360
            Width           =   150
         End
         Begin VB.Label lblP 
            AutoSize        =   -1  'True
            Caption         =   "X:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   150
         End
      End
      Begin MSComctlLib.Toolbar tbrFrame 
         Height          =   330
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
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
      Begin VB.Label lblsubframe 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0/0"
         Height          =   195
         Left            =   2160
         TabIndex        =   10
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Index           =   2
      Left            =   240
      TabIndex        =   41
      Top             =   1080
      Width           =   2415
      Begin VB.CommandButton cmdDo 
         Caption         =   "Do it"
         Height          =   375
         Left            =   0
         TabIndex        =   67
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Frame Frame5 
         Caption         =   "Settings"
         Height          =   3495
         Left            =   0
         TabIndex        =   44
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton optOri 
            Caption         =   "Horizontal"
            Height          =   255
            Index           =   1
            Left            =   1200
            TabIndex        =   71
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optOri 
            Caption         =   "Vertical"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.CheckBox chkAutoSnap 
            Caption         =   "Snap selections"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   3120
            Width           =   1575
         End
         Begin VB.CheckBox chkAutoShow 
            Caption         =   "Show in view"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   2880
            Width           =   1575
         End
         Begin VB.TextBox txtASX 
            Height          =   315
            Index           =   3
            Left            =   360
            TabIndex        =   62
            Text            =   "2"
            Top             =   2520
            Width           =   495
         End
         Begin VB.TextBox txtASY 
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   61
            Text            =   "6"
            Top             =   2520
            Width           =   495
         End
         Begin VB.TextBox txtASX 
            Height          =   315
            Index           =   2
            Left            =   360
            TabIndex        =   57
            Text            =   "0"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtASY 
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   56
            Text            =   "0"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtASX 
            Height          =   315
            Index           =   0
            Left            =   360
            TabIndex        =   48
            Text            =   "0"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtASY 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   47
            Text            =   "0"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtASX 
            Height          =   315
            Index           =   1
            Left            =   360
            TabIndex        =   46
            Text            =   "64"
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox txtASY 
            Height          =   315
            Index           =   1
            Left            =   1320
            TabIndex        =   45
            Text            =   "64"
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lbl 
            Caption         =   "X:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   64
            Top             =   2595
            Width           =   255
         End
         Begin VB.Label Label19 
            Caption         =   "Y:"
            Height          =   255
            Index           =   2
            Left            =   1080
            TabIndex        =   63
            Top             =   2595
            Width           =   255
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Number of frames:"
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   2280
            Width           =   1290
         End
         Begin VB.Label lbl 
            Caption         =   "X:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Top             =   1995
            Width           =   255
         End
         Begin VB.Label Label19 
            Caption         =   "Y:"
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   58
            Top             =   1995
            Width           =   255
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Spacing:"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Start Location:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label Label15 
            Caption         =   "X:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   795
            Width           =   255
         End
         Begin VB.Label Label16 
            Caption         =   "Y:"
            Height          =   255
            Left            =   1080
            TabIndex        =   52
            Top             =   795
            Width           =   255
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Frame size:"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   795
         End
         Begin VB.Label lbl 
            Caption         =   "W:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   1395
            Width           =   255
         End
         Begin VB.Label Label19 
            Caption         =   "H:"
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   49
            Top             =   1395
            Width           =   255
         End
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Subframes"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   43
         Top             =   0
         Width           =   1095
      End
      Begin VB.OptionButton optAuto 
         Caption         =   "Frames"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
      Begin VB.ListBox lstRes 
         Height          =   3570
         ItemData        =   "frmFrames.frx":0000
         Left            =   0
         List            =   "frmFrames.frx":0002
         TabIndex        =   7
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblResInfo 
         AutoSize        =   -1  'True
         Caption         =   "128x128"
         Height          =   195
         Left            =   0
         TabIndex        =   9
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Resource for this frame:"
         Height          =   195
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1680
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4215
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
      Begin VB.Frame Frame4 
         Caption         =   "Selection"
         Height          =   1095
         Left            =   0
         TabIndex        =   37
         Top             =   1800
         Width           =   2415
         Begin VB.CheckBox chkSnap 
            Caption         =   "Snap Selection"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Background color:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label lblColor 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0 0 0"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   2175
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "View"
         Height          =   1695
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   2415
         Begin VB.CommandButton cmdReset 
            Caption         =   "Reset Scroll"
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   2175
         End
         Begin MSComctlLib.Slider sldZoom 
            Height          =   615
            Left            =   120
            TabIndex        =   33
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1085
            _Version        =   393216
            Min             =   1
            Max             =   6
            SelStart        =   1
            TickStyle       =   2
            Value           =   1
         End
         Begin VB.Label lblZoom 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1x"
            Height          =   195
            Left            =   2130
            TabIndex        =   35
            Top             =   240
            Width           =   165
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Zoom:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   450
         End
      End
   End
   Begin MSComctlLib.TabStrip tabs 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8281
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Resource"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Frame"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Auto"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Misc"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pDX 
      DrawMode        =   7  'Invert
      Height          =   5535
      Left            =   2880
      MousePointer    =   2  'Cross
      ScaleHeight     =   365
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlToolbarIcons(0)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "addfrm"
            Object.ToolTipText     =   "Add Frame"
            ImageKey        =   "addfrm"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remfrm"
            Object.ToolTipText     =   "Remove Frame"
            ImageKey        =   "remfrm"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "prev"
            Object.ToolTipText     =   "Previous Frame"
            ImageKey        =   "prev"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "next"
            Object.ToolTipText     =   "Next Frame"
            ImageKey        =   "next"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "select"
            Object.ToolTipText     =   "Select Tool"
            ImageKey        =   "select"
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "move"
            Object.ToolTipText     =   "Move Tool"
            ImageKey        =   "move"
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Quickhelp for current tab"
            ImageKey        =   "help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   0
      Left            =   1560
      Top             =   4080
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
            Picture         =   "frmFrames.frx":0004
            Key             =   "addfrm"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrames.frx":0116
            Key             =   "remfrm"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrames.frx":0228
            Key             =   "help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrames.frx":033A
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrames.frx":044C
            Key             =   "next"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrames.frx":055E
            Key             =   "select"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrames.frx":0670
            Key             =   "move"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   1
      Left            =   3900
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8421376
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrames.frx":0782
            Key             =   "prev"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrames.frx":0894
            Key             =   "next"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFrames.frx":09A6
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCord 
      Alignment       =   2  'Center
      Caption         =   "0, 0"
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   5460
      Width           =   2655
   End
   Begin VB.Label lblSelframe 
      AutoSize        =   -1  'True
      Caption         =   "Frame 0/0"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long) As Long

'DirectX
Private DirectX As New DirectX8
Private Direct3D As Direct3D8
Private Direct3DDevice As Direct3DDevice8
Private Direct3DX As New D3DX8
Private Sprites As D3DXSprite
Private texImg As Direct3DTexture8 'texture
Private texLoaded As Boolean 'texture is loaded
Private DXOK As Boolean 'DirectX ok

'Mouse
Private bNoSel As Boolean 'dont draw selection in pDX_Paint
Private rSel As RECT 'selection rectangle
Private vecMove As D3DVECTOR2 'used when moving image
Private texVec As D3DVECTOR2 'image offset
Private bMd As Boolean 'mouse button down
Private bMove As Boolean 'else select

'other
Private lCurFrame As Long
Private lCurSubFrame As Long
Private lFrmCol As Long
Private lFrmAni As Long

Private Sub chkAutoShow_Click()
  Prg.bShowAuto = CBool(chkAutoShow.Value)
  pDX_Paint
End Sub

Private Sub chkAutoSnap_Click()
  Prg.bSnapAuto = CBool(chkAutoSnap.Value)
End Sub

Private Sub chkShowInView_Click()
  Prg.bShowInView = CBool(chkShowInView.Value)
  pDX_Paint
End Sub

Private Sub chkSnap_Click()
  Prg.bSnap = chkSnap.Value
End Sub

Private Sub cmdDo_Click()
  Dim x As Long, Y As Long, r As RECT
  With Prj.Coll(lFrmCol).Anim(lFrmAni)
  
    If optAuto(0) Then
      If UBound(.Frame) > 0 Then
        x = MsgBox("This will remove your existing frames. Are you sure you want to continue?", vbQuestion Or vbYesNo, "Confirm frame autocreation")
      Else
        x = vbYes
      End If
      If x = vbYes Then
        ReDim .Frame(0)
        If optOri(0) Then
          For x = 1 To Val(txtASX(3))
            For Y = 1 To Val(txtASY(3))
              r.Left = Val(txtASX(0)) + ((x - 1) * Val(txtASX(1))) + ((x - 1) * Val(txtASX(2)))
              r.Right = Val(txtASX(0)) + (x * Val(txtASX(1))) + ((x - 1) * Val(txtASX(2)))
              r.Top = Val(txtASY(0)) + ((Y - 1) * Val(txtASY(1))) + ((Y - 1) * Val(txtASY(2)))
              r.bottom = Val(txtASY(0)) + (Y * Val(txtASY(1))) + ((Y - 1) * Val(txtASY(2)))
              If Prg.bSnapAuto Then SnapSelection r
              
              ReDim Preserve .Frame(UBound(.Frame) + 1)
              ReDim .Frame(UBound(.Frame)).img(.lSubimages)
              .Frame(UBound(.Frame)).img(1).lRes = lstRes.ListIndex + 1
              .Frame(UBound(.Frame)).img(1).Pos = r
              .Frame(UBound(.Frame)).img(1).VirtualSize = r
              
              .Frame(UBound(.Frame)).Ctrl.lDelay = 100
              .Frame(UBound(.Frame)).img(lCurSubFrame).lRes = lstRes.ListIndex + 1
              
            Next
          Next
        Else
          For Y = 1 To Val(txtASY(3))
            For x = 1 To Val(txtASX(3))
              r.Left = Val(txtASX(0)) + ((x - 1) * Val(txtASX(1))) + ((x - 1) * Val(txtASX(2)))
              r.Right = Val(txtASX(0)) + (x * Val(txtASX(1))) + ((x - 1) * Val(txtASX(2)))
              r.Top = Val(txtASY(0)) + ((Y - 1) * Val(txtASY(1))) + ((Y - 1) * Val(txtASY(2)))
              r.bottom = Val(txtASY(0)) + (Y * Val(txtASY(1))) + ((Y - 1) * Val(txtASY(2)))
              If Prg.bSnapAuto Then SnapSelection r
              
              ReDim Preserve .Frame(UBound(.Frame) + 1)
              ReDim .Frame(UBound(.Frame)).img(.lSubimages)
              .Frame(UBound(.Frame)).img(1).lRes = lstRes.ListIndex + 1
              .Frame(UBound(.Frame)).img(1).Pos = r
              .Frame(UBound(.Frame)).img(1).VirtualSize = r
              
              .Frame(UBound(.Frame)).Ctrl.lDelay = 100
              .Frame(UBound(.Frame)).img(lCurSubFrame).lRes = lstRes.ListIndex + 1
              
            Next
          Next
        End If
        lCurFrame = UBound(.Frame)
        lCurSubFrame = 1
        RefreshControls
        Prg.bChanged = True
      End If
    
    Else
      If Val(txtASY(3)) * Val(txtASX(3)) > .lSubimages Then
        MsgBox "These settings vill ammount to more subframes than this frame has. The autocreation cannot proceed.", vbExclamation
      ElseIf Val(txtASY(3)) * Val(txtASX(3)) < .lSubimages Then
        MsgBox "These settings vill ammount to less subframes than this frame has. The autocreation cannot proceed.", vbExclamation
      Else
        Dim z As Long
        z = 1
        If optOri(0) Then
          For x = 1 To Val(txtASX(3))
            For Y = 1 To Val(txtASY(3))
              r.Left = Val(txtASX(0)) + ((x - 1) * Val(txtASX(1))) + ((x - 1) * Val(txtASX(2)))
              r.Right = Val(txtASX(0)) + (x * Val(txtASX(1))) + ((x - 1) * Val(txtASX(2)))
              r.Top = Val(txtASY(0)) + ((Y - 1) * Val(txtASY(1))) + ((Y - 1) * Val(txtASY(2)))
              r.bottom = Val(txtASY(0)) + (Y * Val(txtASY(1))) + ((Y - 1) * Val(txtASY(2)))
              If Prg.bSnapAuto Then SnapSelection r
              
              .Frame(lCurFrame).img(z).lRes = lstRes.ListIndex + 1
              .Frame(lCurFrame).img(z).Pos = r
              .Frame(lCurFrame).img(z).VirtualSize = r
              z = z + 1
            Next
          Next
        Else
          For Y = 1 To Val(txtASY(3))
            For x = 1 To Val(txtASX(3))
              r.Left = Val(txtASX(0)) + ((x - 1) * Val(txtASX(1))) + ((x - 1) * Val(txtASX(2)))
              r.Right = Val(txtASX(0)) + (x * Val(txtASX(1))) + ((x - 1) * Val(txtASX(2)))
              r.Top = Val(txtASY(0)) + ((Y - 1) * Val(txtASY(1))) + ((Y - 1) * Val(txtASY(2)))
              r.bottom = Val(txtASY(0)) + (Y * Val(txtASY(1))) + ((Y - 1) * Val(txtASY(2)))
              If Prg.bSnapAuto Then SnapSelection r
              
              .Frame(lCurFrame).img(z).lRes = lstRes.ListIndex + 1
              .Frame(lCurFrame).img(z).Pos = r
              .Frame(lCurFrame).img(z).VirtualSize = r
              z = z + 1
              
            Next
          Next
        End If
        RefreshControls
        Prg.bChanged = True
      End If
    End If
  End With
End Sub

Private Sub cmdPSet_Click()
  txtP(0) = rSel.Left
  txtP(1) = rSel.Top
  txtP(2) = rSel.Right - rSel.Left
  txtP(3) = rSel.bottom - rSel.Top
  txtP_KeyUp 0, 0, 0
  Prg.bChanged = True
  pDX_Paint
End Sub

Private Sub cmdReset_Click()
  texVec.x = 0
  texVec.Y = 0
  pDX_Paint
End Sub

Private Sub cmdVSet_Click()
  txtV(0) = rSel.Left
  txtV(1) = rSel.Top
  txtV(2) = rSel.Right - rSel.Left
  txtV(3) = rSel.bottom - rSel.Top
  txtV_KeyUp 0, 0, 0
  Prg.bChanged = True
  pDX_Paint
End Sub

Private Sub Form_Load()
  On Error Resume Next
  Dim x As Long
  Me.Width = Prg.lFW
  Me.Height = Prg.lFH
  lFrmCol = frmMain.lSelCol
  lFrmAni = frmMain.lSelAni
  Me.Caption = "Edit Frames [" & Prj.Coll(lFrmCol).Anim(lFrmAni).sID & "]"
  chkShowInView.Value = Abs(Prg.bShowInView)
  chkAutoShow.Value = Abs(Prg.bShowAuto)
  chkAutoSnap.Value = Abs(Prg.bSnapAuto)
  chkSnap.Value = Abs(Prg.bSnap)
  For x = 0 To 3
    txtASX(x) = Prg.lAutoX(x)
    txtASY(x) = Prg.lAutoY(x)
  Next
  If Prg.bAutoFrame Then optOri(0).Value = True Else optOri(1).Value = True
  lblColor.BackColor = Prg.lFrmBgCol
  lblColor.ForeColor = RGB(255 - (Prg.lFrmBgCol Mod &H100), 255 - ((Prg.lFrmBgCol \ &H100) Mod &H100), 255 - ((Prg.lFrmBgCol \ &H10000) Mod &H100))
  lblColor.Caption = Prg.lFrmBgCol Mod &H100 & " " & (Prg.lFrmBgCol \ &H100) Mod &H100 & " " & (Prg.lFrmBgCol \ &H10000) Mod &H100
  tabs_Click
  InitDX
  If UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame) > 0 Then
    lCurFrame = 1
    lCurSubFrame = 1
  Else
    lCurFrame = 0
    lCurSubFrame = 0
  End If
  If UBound(Prj.Res) > 0 Then
    For x = 1 To UBound(Prj.Res)
      lstRes.AddItem x & ". " & Right(Prj.Res(x).sFilename, Len(Prj.Res(x).sFilename) - InStrRev(Prj.Res(x).sFilename, "\"))
    Next
    If UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame) > 0 Then
      lstRes.ListIndex = Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).lRes - 1
    Else
      lstRes.ListIndex = 0
    End If
  Else
    lstRes_Click
    MsgBox "You have not loaded any resources yet. Put simply, resources are images that contain the frames for your animations. To load an image as an resource, exit this window and select 'Resources...' from the 'Tools' dropdown menu.", vbInformation, "No resources loaded!"
  End If
  RefreshControls
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  If Me.Width < 7000 Then Me.Width = 7000
  If Me.Height < 6130 Then Me.Height = 6130
  pDX.Width = Me.ScaleWidth - pDX.Left - 6
  pDX.Height = Me.ScaleHeight - pDX.Top - 6
  lblCord.Top = Me.ScaleHeight - lblCord.Height - 2
  tabs.Height = Me.ScaleHeight - tabs.Top - lblCord.Height - 5
  frm(0).Height = tabs.Height - 32
  lblResInfo.Top = (frm(0).Height * Screen.TwipsPerPixelY) - lblResInfo.Height
  lstRes.Height = (frm(0).Height * Screen.TwipsPerPixelY) - lblResInfo.Height - 240
  InitDX
  lstRes_Click
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Prg.lFW = Me.Width
  Prg.lFH = Me.Height
  Dim x As Long
  For x = 0 To 3
    Prg.lAutoX(x) = Val(txtASX(x))
    Prg.lAutoY(x) = Val(txtASY(x))
  Next
  Prg.bAutoFrame = optOri(0).Value
  Set Sprites = Nothing
  Set Direct3DDevice = Nothing
  Set Direct3D = Nothing
  Set texImg = Nothing
End Sub

Private Sub lblColor_Click()
  Load frmColor
  frmColor.Tag = Prg.lFrmBgCol
  frmColor.Show vbModal, Me
  Prg.lFrmBgCol = frmColor.Tag
  lblColor.BackColor = Prg.lFrmBgCol
  lblColor.ForeColor = RGB(255 - (Prg.lFrmBgCol Mod &H100), 255 - ((Prg.lFrmBgCol \ &H100) Mod &H100), 255 - ((Prg.lFrmBgCol \ &H10000) Mod &H100))
  lblColor.Caption = Prg.lFrmBgCol Mod &H100 & " " & (Prg.lFrmBgCol \ &H100) Mod &H100 & " " & (Prg.lFrmBgCol \ &H10000) Mod &H100
  frmColor.Tag = "Y"
  Unload frmColor
  
  pDX_Paint
End Sub

Private Sub lstRes_Click()
  On Error Resume Next
  Dim texInfo As D3DXIMAGE_INFO
  If lstRes.ListIndex <> -1 Then
    With Prj.Res(lstRes.ListIndex + 1)
      If DXOK Then
        If .lTranscolor = -1 Then
          Set texImg = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, .sFilename, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, 0, texInfo, ByVal 0)
        Else
          Set texImg = Direct3DX.CreateTextureFromFileEx(Direct3DDevice, .sFilename, -1, -1, 1, 0, 0, D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_DEFAULT, D3DColorRGBA(.lTranscolor Mod &H100, (.lTranscolor \ &H100) Mod &H100, (.lTranscolor \ &H10000) Mod &H100, 255), texInfo, ByVal 0)
        End If
        lblResInfo = "Size: " & texInfo.Width & "x" & texInfo.Height
        texLoaded = True
      End If
    End With
    If lCurFrame > 0 And lCurSubFrame > 0 Then
      Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).lRes = lstRes.ListIndex + 1
      Prg.bChanged = True
    End If
  Else
    lblResInfo = "Size: "
    texLoaded = False
    Set texImg = Nothing
  End If
  pDX_Paint
End Sub

Private Sub pDX_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  bMd = True
  If Not bMove Then
    rSel.Left = x - texVec.x
    rSel.Top = Y - texVec.Y
    rSel.Right = rSel.Left
    rSel.bottom = rSel.Top
  Else
    vecMove.x = x - texVec.x
    vecMove.Y = Y - texVec.Y
  End If
End Sub

Private Sub pDX_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If texLoaded And bMd Then
    If Not bMove Then
      rSel.Right = x - texVec.x
      rSel.bottom = Y - texVec.Y
      If rSel.Right < rSel.Left Then rSel.Right = rSel.Left
      If rSel.bottom < rSel.Top Then rSel.bottom = rSel.Top
      lblCord = "X: " & rSel.Left & ", Y: " & rSel.Top & ", W: " & rSel.Right - rSel.Left & ", H: " & rSel.bottom - rSel.Top
    Else
      texVec.x = x - vecMove.x
      texVec.Y = Y - vecMove.Y
      lblCord = "OffsetX: " & texVec.x & ", OffsetY: " & texVec.Y
    End If
    pDX_Paint
  Else
    Dim c As Long
    c = GetPixel(pDX.hdc, x, Y)
    lblCord = "X: " & x - texVec.x & ", Y: " & Y - texVec.Y & ", RGB: " & c Mod &H100 & " " & (c \ &H100) Mod &H100 & " " & (c \ &H10000) Mod &H100
  End If
End Sub

Private Sub pDX_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  bMd = False
  If Not bMove And Prg.bSnap Then SnapSelection rSel
  pDX_Paint
End Sub

Private Sub pDX_Paint()
  If DXOK Then
    Direct3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(Prg.lFrmBgCol Mod &H100, (Prg.lFrmBgCol \ &H100) Mod &H100, (Prg.lFrmBgCol \ &H10000) Mod &H100, 255), 1, 0
    Direct3DDevice.BeginScene
    Sprites.Begin
    
    If texLoaded Then
      Sprites.Draw texImg, ByVal 0, vec2(sldZoom.Value, sldZoom.Value), vec2(0, 0), 0, texVec, &HFFFFFFFF
    End If
    
    Sprites.End
    Direct3DDevice.EndScene
    Direct3DDevice.Present ByVal 0, ByVal 0, pDX.hwnd, ByVal 0
    
    If rSel.Left <> rSel.Right And Not bNoSel Then
      pDX.Line (rSel.Left + texVec.x, rSel.Top + texVec.Y)-(rSel.Right + texVec.x, rSel.bottom + texVec.Y), vbWhite, B
    End If
    
    If tabs.SelectedItem.Index = 2 And Prg.bShowInView And Not bNoSel Then
      If lCurFrame > 0 And lCurSubFrame > 0 Then
        With Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame)
          pDX.DrawMode = 13
          If .Pos.Left <> .Pos.Right Then
            pDX.Line (.Pos.Left + texVec.x, .Pos.Top + texVec.Y)-(.Pos.Right + texVec.x, .Pos.bottom + texVec.Y), &H808080, B
          End If
          If (Prj.Coll(lFrmCol).Anim(lFrmAni).eOptions And VirtualSize_1) = VirtualSize_1 Then
            pDX.Line (.VirtualSize.Left + texVec.x, .VirtualSize.Top + texVec.Y)-(.VirtualSize.Right + texVec.x, .VirtualSize.bottom + texVec.Y), &HC0C0C0, B
          End If
          pDX.DrawMode = 7
        End With
      End If
    End If
    
    If tabs.SelectedItem.Index = 3 And Prg.bShowAuto And Not bNoSel Then
      pDX.DrawMode = 13
      Dim x As Long, Y As Long, r As RECT
      For x = 1 To Val(txtASX(3))
        For Y = 1 To Val(txtASY(3))
          r.Left = Val(txtASX(0)) + ((x - 1) * Val(txtASX(1))) + ((x - 1) * Val(txtASX(2)))
          r.Right = Val(txtASX(0)) + (x * Val(txtASX(1))) + ((x - 1) * Val(txtASX(2)))
          r.Top = Val(txtASY(0)) + ((Y - 1) * Val(txtASY(1))) + ((Y - 1) * Val(txtASY(2)))
          r.bottom = Val(txtASY(0)) + (Y * Val(txtASY(1))) + ((Y - 1) * Val(txtASY(2)))
          pDX.Line (r.Left + texVec.x, r.Top + texVec.Y)-(r.Right + texVec.x, r.bottom + texVec.Y), vbRed, B
        Next
      Next
      pDX.DrawMode = 7
    End If
    
  End If
End Sub

Private Sub sldZoom_Scroll()
  lblZoom = sldZoom.Value & "x"
  pDX_Paint
End Sub

Private Sub tabs_Click()
  Dim x As Long
  For x = 0 To 3
    If x <> tabs.SelectedItem.Index - 1 Then
      frm(x).Visible = False
    Else
      frm(x).Visible = True
    End If
  Next
  sldZoom.Value = 1
  pDX_Paint
End Sub

Private Sub tbrFrame_ButtonClick(ByVal Button As MSComCtlLib.Button)
  On Error Resume Next
  Select Case Button.Key
    Case "prev"
      lCurSubFrame = lCurSubFrame - 1
      If lCurSubFrame < 1 Then lCurSubFrame = 1
    Case "next"
      lCurSubFrame = lCurSubFrame + 1
      If lCurSubFrame > Prj.Coll(lFrmCol).Anim(lFrmAni).lSubimages Then lCurSubFrame = Prj.Coll(lFrmCol).Anim(lFrmAni).lSubimages
      If Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).lRes = 0 Then Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).lRes = lstRes.ListIndex + 1
  End Select
  RefreshControls
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComCtlLib.Button)
  On Error Resume Next
  Select Case Button.Key
    Case "move"
      bMove = True
      tbrMain.Buttons("select").Value = tbrUnpressed
      tbrMain.Buttons("move").Value = tbrPressed
      pDX.MousePointer = vbSizeAll
    Case "select"
      bMove = False
      tbrMain.Buttons("move").Value = tbrUnpressed
      tbrMain.Buttons("select").Value = tbrPressed
      pDX.MousePointer = 2
    Case "next"
      lCurFrame = lCurFrame + 1
      If lCurFrame > UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame) Then lCurFrame = UBound(Prj.Coll(lFrmCol).Anim(lFrmAni).Frame)
      RefreshControls
    Case "prev"
      lCurFrame = lCurFrame - 1
      If lCurFrame < 1 Then lCurFrame = 1
      RefreshControls
    Case "addfrm"
      Prg.bChanged = True
      With Prj.Coll(lFrmCol).Anim(lFrmAni)
        ReDim Preserve .Frame(UBound(.Frame) + 1)
        lCurFrame = UBound(.Frame)
        ReDim .Frame(lCurFrame).img(.lSubimages)
        lCurSubFrame = 1
        .Frame(lCurFrame).Ctrl.lDelay = 100
        .Frame(lCurFrame).img(lCurSubFrame).lRes = lstRes.ListIndex + 1
        
        RefreshControls
        
      End With
    Case "remfrm"
      Dim x As Long, Y As Long
      If Prg.bConfirm(2) Then
        x = MsgBox("Are you sure you want to remove this frame from the animation?", vbExclamation Or vbYesNo, "Confirmation")
      Else
        x = vbYes
      End If
      If x = vbYes Then
        Prg.bChanged = True
        With Prj.Coll(lFrmCol).Anim(lFrmAni)
          For x = lCurFrame To UBound(.Frame) - 1
            .Frame(x).Ctrl.lDelay = .Frame(x + 1).Ctrl.lDelay
            For Y = 1 To .lSubimages
              .Frame(x).img(Y).Offset = .Frame(x + 1).img(Y).Offset
              .Frame(x).img(Y).lRes = .Frame(x + 1).img(Y).lRes
              .Frame(x).img(Y).Pos = .Frame(x + 1).img(Y).Pos
              .Frame(x).img(Y).VirtualSize = .Frame(x + 1).img(Y).VirtualSize
            Next
          Next x
          ReDim Preserve .Frame(UBound(.Frame) - 1)
          If UBound(.Frame) > 0 Then
            lCurSubFrame = 1
            If lCurFrame > UBound(.Frame) Then lCurFrame = UBound(.Frame)
          Else
            lCurFrame = 0
            lCurSubFrame = 0
          End If
          RefreshControls
        End With
      End If
      
    Case "help"
      frmHelp.Show , Me
      frmHelp.web.Navigate App.Path & "\doc\frm.html"
  End Select
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
  
  texLoaded = False
  DXOK = True
  pDX_Paint
  Exit Sub
errh:
  MsgBox "It seems that DirectX 8 3D H&L support is nowhere to be found... DirectX init failed!", vbCritical, "DX Error"
End Sub

Private Sub SnapSelection(ByRef rs As RECT)
  Dim x As Long, c As Long, b As Boolean
  If rs.Right + texVec.x < 0 Or rs.Right + texVec.x > pDX.Width Or rs.bottom + texVec.Y < 0 Or rs.bottom + texVec.Y > pDX.Height Then
    MsgBox "The selection must be fully visible in the view area for snap to work.", vbInformation
    Exit Sub
  End If
  
  bNoSel = True
  pDX_Paint
  bNoSel = False
  
  Me.MousePointer = vbArrowHourglass
  DoEvents
  
  'upper edge
  c = rs.Top + texVec.Y - 1
  Do
    c = c + 1
    For x = rs.Left + texVec.x To rs.Right + texVec.x
      If GetPixel(pDX.hdc, x, c) <> Prg.lFrmBgCol Then
        b = True
        Exit For
      End If
    Next
  Loop Until b Or c >= rs.bottom + texVec.Y
  rs.Top = c - texVec.Y
  
  'lower edge
  b = False
  c = rs.bottom + texVec.Y '+ 1
  Do
    c = c - 1
    For x = rs.Left + texVec.x To rs.Right + texVec.x
      If GetPixel(pDX.hdc, x, c) <> Prg.lFrmBgCol Then
        b = True
        Exit For
      End If
    Next
  Loop Until b Or c <= rs.Top + texVec.Y
  rs.bottom = c - texVec.Y + 1

  'left edge
  b = False
  c = rs.Left + texVec.x - 1
  Do
    c = c + 1
    For x = rs.Top + texVec.Y To rs.bottom + texVec.Y
      If GetPixel(pDX.hdc, c, x) <> Prg.lFrmBgCol Then
        b = True
        Exit For
      End If
    Next
  Loop Until b Or c >= rs.Right + texVec.x
  rs.Left = c - texVec.x

  'right edge
  b = False
  c = rs.Right + texVec.x '+ 1
  Do
    c = c - 1
    For x = rs.Top + texVec.Y To rs.bottom + texVec.Y
      If GetPixel(pDX.hdc, c, x) <> Prg.lFrmBgCol Then
        b = True
        Exit For
      End If
    Next
  Loop Until b Or c <= rs.Left + texVec.x
  rs.Right = c - texVec.x + 1

  Me.MousePointer = vbDefault

End Sub

Private Sub txtASX_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = ctlKeyPress(KeyAscii, NumbersOnly)
End Sub

Private Sub txtASX_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  pDX_Paint
End Sub

Private Sub txtASY_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = ctlKeyPress(KeyAscii, NumbersOnly)
End Sub

Private Sub txtASY_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  pDX_Paint
End Sub

Private Sub txtP_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = ctlKeyPress(KeyAscii, NumbersOnly)
End Sub

Private Sub txtP_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  With Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).Pos
    .Left = Val(txtP(0))
    .Top = Val(txtP(1))
    .Right = .Left + Val(txtP(2))
    .bottom = .Top + Val(txtP(3))
    Prg.bChanged = True
  End With
End Sub

Private Sub txtV_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = ctlKeyPress(KeyAscii, NumbersOnly)
End Sub

Public Sub RefreshControls()
  With Prj.Coll(lFrmCol).Anim(lFrmAni)
    tbrMain.Buttons(2).Enabled = (lCurFrame <> 0)
    tbrMain.Buttons(3).Enabled = Not (lCurFrame <= 1)
    tbrMain.Buttons(4).Enabled = Not (lCurFrame >= UBound(.Frame))
    tbrFrame.Buttons(1).Enabled = Not (lCurSubFrame <= 1)
    tbrFrame.Buttons(2).Enabled = Not (lCurSubFrame >= .lSubimages Or lCurSubFrame = 0)
    lblSelframe = "Frame: " & lCurFrame & "/" & UBound(.Frame)
    lblsubframe = "Subframe: " & lCurSubFrame & "/" & .lSubimages
    lblsubframe.Enabled = (UBound(.Frame) > 0)
    
    Dim x As Long
    For x = 0 To 3
      txtP(x).Enabled = (lCurFrame > 0)
      lblP(x).Enabled = (lCurFrame > 0)
      If (.eOptions And VirtualSize_1) = VirtualSize_1 Then
        txtV(x).Enabled = (lCurFrame > 0)
        lblV(x).Enabled = (lCurFrame > 0)
      Else
        txtV(x).Enabled = False
        lblV(x).Enabled = False
      End If
    Next
    cmdPSet.Enabled = (lCurFrame > 0)
    If (.eOptions And VirtualSize_1) = VirtualSize_1 Then
      cmdVSet.Enabled = (lCurFrame > 0)
    Else
      cmdVSet.Enabled = False
    End If

    If lCurFrame > 0 And lCurSubFrame > 0 Then
      txtP(0) = .Frame(lCurFrame).img(lCurSubFrame).Pos.Left
      txtP(1) = .Frame(lCurFrame).img(lCurSubFrame).Pos.Top
      txtP(2) = .Frame(lCurFrame).img(lCurSubFrame).Pos.Right - .Frame(lCurFrame).img(lCurSubFrame).Pos.Left
      txtP(3) = .Frame(lCurFrame).img(lCurSubFrame).Pos.bottom - .Frame(lCurFrame).img(lCurSubFrame).Pos.Top
      If (.eOptions And VirtualSize_1) = VirtualSize_1 Then
        txtV(0) = .Frame(lCurFrame).img(lCurSubFrame).VirtualSize.Left
        txtV(1) = .Frame(lCurFrame).img(lCurSubFrame).VirtualSize.Top
        txtV(2) = .Frame(lCurFrame).img(lCurSubFrame).VirtualSize.Right - .Frame(lCurFrame).img(lCurSubFrame).VirtualSize.Left
        txtV(3) = .Frame(lCurFrame).img(lCurSubFrame).VirtualSize.bottom - .Frame(lCurFrame).img(lCurSubFrame).VirtualSize.Top
      End If
      
      If Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).lRes > 0 And lstRes.ListIndex + 1 <> Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).lRes Then
        lstRes.ListIndex = Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).lRes - 1
        'lstRes_Click
      End If
    End If

  End With
  pDX_Paint
End Sub

Private Sub txtV_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  With Prj.Coll(lFrmCol).Anim(lFrmAni).Frame(lCurFrame).img(lCurSubFrame).VirtualSize
    .Left = Val(txtV(0))
    .Top = Val(txtV(1))
    .Right = .Left + Val(txtV(2))
    .bottom = .Top + Val(txtV(3))
    Prg.bChanged = True
  End With
End Sub

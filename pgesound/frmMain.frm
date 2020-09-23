VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PgeSound"
   ClientHeight    =   5295
   ClientLeft      =   150
   ClientTop       =   465
   ClientWidth     =   7335
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   6855
      Visible         =   0   'False
      Begin VB.Frame Frame1 
         Caption         =   "Music Settings"
         Height          =   4455
         Left            =   2160
         TabIndex        =   5
         Top             =   120
         Width           =   4695
         Begin VB.TextBox txtMFile 
            ForeColor       =   &H80000011&
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   240
            Width           =   3975
         End
         Begin VB.CommandButton cmdMBrowse 
            Caption         =   "..."
            Height          =   315
            Left            =   4200
            TabIndex        =   40
            Top             =   240
            Width           =   375
         End
         Begin VB.Frame Frame2 
            Caption         =   "Music Defaults"
            Height          =   1095
            Left            =   120
            TabIndex        =   10
            Top             =   2280
            Width           =   4455
            Begin VB.CheckBox chkMReverb 
               Caption         =   "&Reverb (MIDI only)"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1080
               TabIndex        =   14
               Top             =   720
               Width           =   1695
            End
            Begin MSComctlLib.Slider sldMVol 
               Height          =   255
               Left            =   840
               TabIndex        =   13
               Top             =   360
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   51
               Max             =   255
               SelStart        =   255
               TickFrequency   =   20
               Value           =   255
            End
            Begin VB.CheckBox chkMLoop 
               Caption         =   "&Loop"
               Height          =   255
               Left            =   120
               TabIndex        =   11
               Top             =   720
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Volume:"
               Height          =   195
               Left            =   120
               TabIndex        =   12
               Top             =   360
               Width           =   570
            End
         End
         Begin VB.TextBox txtMDescr 
            Height          =   795
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   1440
            Width           =   4455
         End
         Begin VB.TextBox txtMID 
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   4455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Music Description:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   1305
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Music ID:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   675
         End
      End
      Begin MSComctlLib.Toolbar tbrMus 
         Height          =   330
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "addmus"
               Object.ToolTipText     =   "Add Music"
               ImageKey        =   "addmus"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "remmus"
               Object.ToolTipText     =   "Remove Music"
               ImageKey        =   "remmus"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cfg"
               Object.ToolTipText     =   "General Settings"
               ImageKey        =   "cfg"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "playback"
               Object.ToolTipText     =   "Toggle Playback Window"
               ImageKey        =   "playback"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "help"
               Object.ToolTipText     =   "Help"
               ImageKey        =   "help"
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lst 
         Height          =   4155
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame frm 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4575
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6855
      Begin MSComctlLib.Toolbar tbrSfx 
         Height          =   330
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "addsnd"
               Object.ToolTipText     =   "Add Sound Effect"
               ImageKey        =   "addsnd"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "remsnd"
               Object.ToolTipText     =   "Remove Sound Effect"
               ImageKey        =   "remsnd"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cfg"
               Object.ToolTipText     =   "General Settings"
               ImageKey        =   "cfg"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "playback"
               Object.ToolTipText     =   "Toggle Playback Window"
               ImageKey        =   "playback"
               Style           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "help"
               ImageKey        =   "help"
            EndProperty
         EndProperty
      End
      Begin VB.ListBox lst 
         Height          =   4155
         Index           =   0
         Left            =   0
         TabIndex        =   24
         Top             =   360
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sound Effect Settings"
         Height          =   4455
         Left            =   2160
         TabIndex        =   15
         Top             =   120
         Width           =   4695
         Begin VB.CommandButton cmdSBrowse 
            Caption         =   "..."
            Height          =   315
            Left            =   4200
            TabIndex        =   39
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtSFile 
            ForeColor       =   &H80000011&
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   3975
         End
         Begin VB.TextBox txtSID 
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   4455
         End
         Begin VB.TextBox txtSDescr 
            Height          =   795
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   1440
            Width           =   4455
         End
         Begin VB.Frame Frame4 
            Caption         =   "Sound Effect Defaults"
            Height          =   1815
            Left            =   120
            TabIndex        =   16
            Top             =   2280
            Width           =   4455
            Begin VB.TextBox txtSPriority 
               Height          =   315
               Left            =   2880
               TabIndex        =   29
               Text            =   "255"
               Top             =   1080
               Width           =   615
            End
            Begin VB.TextBox txtSPback 
               Height          =   315
               Left            =   1440
               TabIndex        =   27
               Text            =   "10"
               Top             =   1080
               Width           =   615
            End
            Begin VB.CheckBox chkSLoop 
               Caption         =   "&Loop"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   720
               Width           =   735
            End
            Begin MSComctlLib.Slider sldSVol 
               Height          =   255
               Left            =   840
               TabIndex        =   17
               Top             =   360
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   51
               Max             =   255
               SelStart        =   255
               TickFrequency   =   20
               Value           =   255
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Priority:"
               Height          =   195
               Left            =   2280
               TabIndex        =   28
               Top             =   1120
               Width           =   510
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Max playbacks:"
               Height          =   195
               Left            =   120
               TabIndex        =   26
               Top             =   1120
               Width           =   1110
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Volume:"
               Height          =   195
               Left            =   120
               TabIndex        =   19
               Top             =   360
               Width           =   570
            End
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sound Effect ID:"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Sound Effect Description:"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1200
            Width           =   1815
         End
      End
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.TabStrip tabs 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sound Effects"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Music"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   3060
      Top             =   2400
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
            Picture         =   "frmMain.frx":030A
            Key             =   "addmus"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041C
            Key             =   "remmus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052E
            Key             =   "addsnd"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0642
            Key             =   "remsnd"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0756
            Key             =   "help"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":086A
            Key             =   "cfg"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":097E
            Key             =   "playback"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmStart 
      Caption         =   "PgeSound"
      Height          =   4455
      Left            =   2400
      TabIndex        =   30
      Top             =   600
      Width           =   4695
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "PgeSound"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   840
         TabIndex        =   37
         Top             =   360
         Width           =   2235
      End
      Begin VB.Label lblVer 
         AutoSize        =   -1  'True
         Caption         =   "1.0.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   36
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "By Paul Berlin 2003 - For used with Pab Game Engine."
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   960
         Width           =   3840
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "berlin_paul@hotmail.com"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         MousePointer    =   10  'Up Arrow
         TabIndex        =   34
         Top             =   1200
         Width           =   1770
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Pge is available from "
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1560
         Width           =   1500
      End
      Begin VB.Label lblPsc 
         AutoSize        =   -1  'True
         Caption         =   "http://www.planetsourcecode.com"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1755
         MousePointer    =   10  'Up Arrow
         TabIndex        =   32
         Top             =   1560
         Width           =   2490
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   300
         Picture         =   "frmMain.frx":0A92
         ToolTipText     =   "Hoppla"
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Freeware!"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   1920
         Width           =   705
      End
   End
   Begin VB.Menu menFile 
      Caption         =   "&File"
      Begin VB.Menu menNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu menLine1 
         Caption         =   "-"
      End
      Begin VB.Menu menOpen 
         Caption         =   "&Open Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu menLine2 
         Caption         =   "-"
      End
      Begin VB.Menu menSave 
         Caption         =   "&Save Project"
         Shortcut        =   ^S
      End
      Begin VB.Menu menSaveAs 
         Caption         =   "Save Project As..."
      End
      Begin VB.Menu menLine3 
         Caption         =   "-"
      End
      Begin VB.Menu menCompile 
         Caption         =   "&Compile Project..."
      End
      Begin VB.Menu menLine4 
         Caption         =   "-"
      End
      Begin VB.Menu menExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menTools 
      Caption         =   "&Tools"
      Begin VB.Menu menGeneralSettings 
         Caption         =   "General Settings..."
      End
      Begin VB.Menu menPlayback 
         Caption         =   "Playback Control"
      End
   End
   Begin VB.Menu menOptions 
      Caption         =   "&Options"
      Begin VB.Menu menConfirm 
         Caption         =   "&Confirm Removing"
      End
   End
   Begin VB.Menu menHelp 
      Caption         =   "&Help"
      Begin VB.Menu menContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu menLine5 
         Caption         =   "-"
      End
      Begin VB.Menu menAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMLoop_Click()
  SaveMusFlags
End Sub

Private Sub chkMReverb_Click()
  SaveMusFlags
End Sub

Private Sub chkSLoop_Click()
  SaveSfxFlags
End Sub

Private Sub cmdMBrowse_Click()
  On Error GoTo errh
  With cdg
    .DialogTitle = "Change Music"
    .filename = Prj.Mus(lst(1).ListIndex + 1).sFile
    .Flags = cdlOFNExplorer Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
    .Filter = "All supported filetypes|*.mp3;*.ogg;*.wav;*.mid;*.midi;*.mod;*.s3m;*.it;*.xm|MPEG Layer 3|*.mp3|OGG Vorbis|*.ogg|Wave-files|*.wav|Midi|*.mid;*.midi|Modules|*.mod;*.it;*.xm;*.s3m"
    .ShowOpen
    Prj.Mus(lst(1).ListIndex + 1).sFile = .filename
    lst_Click lst(1).ListIndex
    Prg.bChanged = True
  End With
errh:
End Sub

Private Sub cmdSBrowse_Click()
  On Error GoTo errh
  With cdg
    .DialogTitle = "Change Sound Effect"
    .filename = Prj.Sfx(lst(0).ListIndex + 1).sFile
    .Flags = cdlOFNExplorer Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
    .Filter = "All supported filetypes|*.mp3;*.ogg;*.wav|MPEG Layer 3|*.mp3|OGG Vorbis|*.ogg|Wave-files|*.wav"
    .ShowOpen
    Prj.Sfx(lst(0).ListIndex + 1).sFile = .filename
    lst_Click lst(0).ListIndex
    Prg.bChanged = True
  End With
errh:
End Sub

Private Sub Form_Load()
  lblVer = App.Major & "." & App.Minor & "." & App.Revision
  LoadSettings
  If Not modAss.CheckAssociation("spj", "PgeSound.spj") Then
    modAss.CreateFileType "PgeSound.spj", "PgeSound Project", App.Path & "\pgesound.exe,0"
    modAss.CreateFileTypeAction "PgeSound.spj", "Open", Chr(34) & App.Path & "\pgesound.exe" & Chr(34) & " -o %1"
    modAss.CreateFileTypeAction "PgeSound.spj", "Compile", Chr(34) & App.Path & "\pgesound.exe" & Chr(34) & " -c %1"
    modAss.CreateAssociation ".spj", "PgeSound.backup", "PgeSound.spj"
  End If

  Dim s As String
  s = Trim(Command)
  If Len(s) > 0 Then
    Dim cmd As String
    cmd = Left(s, 2)
    Select Case cmd
      Case "-o"
        Me.Show
        DoEvents
        OpenProject Trim(Right(s, Len(s) - 2))
      Case "-c"
        OpenProject Trim(Right(s, Len(s) - 2))
        If Len(Prj.sFilenameCompile) > 0 Then
          CompileProject Prj.sFilenameCompile, True
        Else
          menCompile_Click
        End If
        Unload Me
      Case Else
        MsgBox "Command not recognised.", vbExclamation, "Error"
        menNew_Click
    End Select
  Else
    menNew_Click
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  Unload frmPlayback
  If Prg.bChanged Then
    Dim x As Long
    x = MsgBox("You have made changes to the open project without saving. Do you want to save before exiting?", vbInformation Or vbYesNoCancel, "Discard changes?")
    If x = vbYes Then
      menSave_Click
    ElseIf x = vbCancel Then
      Cancel = 1
      If menPlayback.Checked Then
        frmPlayback.Move Me.Left - frmPlayback.Width - 100, Me.Top
        frmPlayback.Show , Me
      End If
    End If
  End If
  SaveSettings
  
End Sub

Private Sub frmStart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblEmail.FontUnderline = False
  lblPsc.FontUnderline = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblEmail.FontUnderline = False
  lblPsc.FontUnderline = False
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblEmail.FontUnderline = False
  lblPsc.FontUnderline = False
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblEmail.FontUnderline = False
  lblPsc.FontUnderline = False
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblEmail.FontUnderline = False
  lblPsc.FontUnderline = False
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblEmail.FontUnderline = False
  lblPsc.FontUnderline = False
End Sub

Private Sub lblEmail_Click()
  ShellExecute Me.hwnd, "open", "mailto:berlin_paul@hotmail.com?subject=PgeTiles", vbNullString, vbNullString, 1
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not lblEmail.FontUnderline Then
    lblEmail.FontUnderline = True
  End If
End Sub

Private Sub lblPsc_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Not lblPsc.FontUnderline Then
    lblPsc.FontUnderline = True
  End If
End Sub

Private Sub lblPsc_Click()
  ShellExecute Me.hwnd, "open", "http://www.planetsourcecode.com", vbNullString, vbNullString, 1
End Sub

Private Sub lblVer_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  lblEmail.FontUnderline = False
  lblPsc.FontUnderline = False
End Sub

Private Sub lst_Click(index As Integer)
  If lst(index).ListIndex > -1 Then
    frmStart.Visible = False
    Select Case index
      Case 0
        tbrSfx.Buttons(2).Enabled = True
        With Prj.Sfx(lst(0).ListIndex + 1)
          txtSID = .sID
          txtSDescr = .sDescr
          sldSVol.Value = .bVol
          txtSPback = .bPlaybacks
          txtSPriority = .bPriority
          txtSFile = .sFile
          txtSFile.SelStart = Len(txtSFile)
          chkSLoop.Value = IIf((.eFlags And Sfx1_Loop) = Sfx1_Loop, 1, 0)
        End With
      Case 1
        tbrMus.Buttons(2).Enabled = True
        With Prj.Mus(lst(1).ListIndex + 1)
          txtMID = .sID
          txtMDescr = .sDescr
          sldMVol.Value = .bVol
          txtMFile = .sFile
          txtMFile.SelStart = Len(txtMFile)
          chkMLoop.Value = IIf((.eFlags And Mus1_Loop) = Mus1_Loop, 1, 0)
          chkMReverb.Value = IIf((.eFlags And Mus2_Reverb) = Mus2_Reverb, 1, 0)
        End With
    End Select
  Else
    frmStart.Visible = True
    Select Case index
      Case 0
        tbrSfx.Buttons(2).Enabled = False
      Case 1
        tbrMus.Buttons(2).Enabled = False
    End Select
  End If
End Sub

Private Sub menAbout_Click()
  frmStart.Visible = True
End Sub

Private Sub menCompile_Click()
  On Error GoTo errh
  With cdg
    .DialogTitle = "Compile Project"
    .filename = Prj.sFilenameCompile
    .Flags = cdlOFNExplorer Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    .Filter = "Compiled soundset (*.pbs)|*.pbs"
    .ShowSave
    
    Prj.sFilenameCompile = .filename
    CompileProject .filename
  End With
errh:
End Sub

Private Sub menConfirm_Click()
  menConfirm.Checked = Not menConfirm.Checked
  Prg.bConfirm = menConfirm.Checked
End Sub

Private Sub menContents_Click()
  frmHelp.Show , Me
  frmHelp.web.Navigate App.Path & "\doc\index.html"
End Sub

Private Sub menExit_Click()
  Unload Me
End Sub

Private Sub menGeneralSettings_Click()
  frmGenSet.Show vbModal, Me
  Me.Caption = "PgeSound [" & Prj.sID & "]"
End Sub

Private Sub menNew_Click()
  If Prg.bChanged Then
    If MsgBox("You have made changes to the open project without saving. Are you sure you want to discard them?", vbOKCancel Or vbExclamation, "Discard changes?") = vbCancel Then Exit Sub
  End If
  
  ReDim Prj.Mus(0)
  ReDim Prj.Sfx(0)
  Prj.eMusFlags = Mus1_Loop
  Prj.eSfxFlags = 0
  Prj.sFilename = ""
  Prj.sFilenameCompile = ""
  Prj.sDescr = "Untitled Soundset" & vbCrLf & "-----------------" & vbCrLf & "Copyright © Creator." & vbCrLf & "Created " & Now & " with PgeSound v" & App.Major & "." & App.Minor & "." & App.Revision & "., written by Paul Berlin."
  Prj.sID = "Untitled Soundset"
  
  Me.Caption = "PgeSound [" & Prj.sID & "]"
  tbrMus.Buttons(2).Enabled = False
  tbrSfx.Buttons(2).Enabled = False
End Sub

Private Sub menOpen_Click()
  If Prg.bChanged Then
    If MsgBox("You have made changes to the open project without saving. Are you sure you want to discard them?", vbOKCancel Or vbExclamation, "Discard changes?") = vbCancel Then Exit Sub
  End If
  On Error GoTo errh
  With cdg
    .DialogTitle = "Open Project"
    .filename = ""
    .Flags = cdlOFNExplorer Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
    .Filter = "PgeSound Project file (*.spj)|*.spj"
    .ShowOpen
    
    OpenProject .filename
    Prg.bChanged = False
  End With
errh:
End Sub

Private Sub menPlayback_Click()
  menPlayback.Checked = Not menPlayback.Checked
  tbrSfx.Buttons(5).Value = Abs(menPlayback.Checked)
  tbrMus.Buttons(5).Value = Abs(menPlayback.Checked)
  If menPlayback.Checked Then
    frmPlayback.Move Me.Left - frmPlayback.Width - 100, Me.Top
    frmPlayback.Show , Me
  Else
    frmPlayback.Hide
  End If
End Sub

Private Sub menSave_Click()
  If Len(Prj.sFilename) > 0 Then
    SaveProject Prj.sFilename
    Prg.bChanged = False
  Else
    menSaveAs_Click
  End If
End Sub

Private Sub menSaveAs_Click()
  On Error GoTo errh
  With cdg
    .DialogTitle = "Save Project"
    .filename = Prj.sFilename
    .Flags = cdlOFNExplorer Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    .Filter = "PgeSound Project file (*.spj)|*.spj"
    .ShowSave
    
    Prj.sFilename = .filename
    SaveProject .filename
    Prg.bChanged = False
  End With
errh:
End Sub

Private Sub sldMVol_Scroll()
  Prj.Mus(lst(1).ListIndex + 1).bVol = sldMVol.Value
  Prg.bChanged = True
End Sub

Private Sub sldSVol_Scroll()
  Prj.Sfx(lst(0).ListIndex + 1).bVol = sldSVol.Value
  Prg.bChanged = True
End Sub

Private Sub tabs_Click()
  Dim x As Long
  For x = 0 To 1
    If x <> tabs.SelectedItem.index - 1 Then
      frm(x).Visible = False
    Else
      frm(x).Visible = True
    End If
  Next
  lst_Click tabs.SelectedItem.index - 1
End Sub

Private Sub tbrsfx_ButtonClick(ByVal Button As MSComCtlLib.Button)
  On Error Resume Next
  Dim x As Long
  Select Case Button.Key
    Case "cfg"
      menGeneralSettings_Click
    Case "playback"
      menPlayback_Click
    Case "addsnd"
      On Error GoTo errh
      With cdg
        .DialogTitle = "Add Sound Effect"
        .filename = ""
        .Flags = cdlOFNExplorer Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .Filter = "All supported filetypes|*.mp3;*.ogg;*.wav|MPEG Layer 3|*.mp3|OGG Vorbis|*.ogg|Wave-files|*.wav"
        .ShowOpen
      End With
      On Error Resume Next
      With Prj
        ReDim Preserve .Sfx(UBound(.Sfx) + 1)
        With .Sfx(UBound(.Sfx))
          .sFile = cdg.filename
          .sID = LCase(sFilename(.sFile, efpFileNameAndExt))
          .eFlags = Prj.eSfxFlags
          .bPlaybacks = 10
          .bPriority = 255
          .bVol = 255
          lst(0).AddItem .sID
          lst(0).ListIndex = lst(0).ListCount - 1
        End With
      End With
      frmStart.Visible = False
      Prg.bChanged = True
    Case "remsnd"
      If Prg.bConfirm Then
        x = MsgBox("You are about to remove the current selected sound effect, '" & Prj.Sfx(lst(0).ListIndex + 1).sID & "'. This cannot be undone. Are you sure?", vbExclamation Or vbYesNo, "Confirm remove")
      Else
        x = vbYes
      End If
      If x = vbYes Then
        
        For x = lst(0).ListIndex + 1 To UBound(Prj.Sfx) - 1
          With Prj.Sfx(x)
            .bPlaybacks = Prj.Sfx(x + 1).bPlaybacks
            .bPriority = Prj.Sfx(x + 1).bPriority
            .bVol = Prj.Sfx(x + 1).bVol
            .eFlags = Prj.Sfx(x + 1).eFlags
            .sDescr = Prj.Sfx(x + 1).sDescr
            .sFile = Prj.Sfx(x + 1).sFile
            .sID = Prj.Sfx(x + 1).sID
          End With
        Next
        
        ReDim Preserve Prj.Sfx(UBound(Prj.Sfx) - 1)
        
        x = lst(0).ListIndex
        lst(0).RemoveItem lst(0).ListIndex
        If x > lst(0).ListCount - 1 Then x = x - 1
        lst(0).ListIndex = x
        
        lst_Click 0
        Prg.bChanged = True
      End If
      
    Case "help"
      frmHelp.Show , Me
      frmHelp.web.Navigate App.Path & "\doc\sfx.html"
  End Select
errh:
End Sub

Private Sub tbrmus_ButtonClick(ByVal Button As MSComCtlLib.Button)
  On Error Resume Next
  Select Case Button.Key
    Case "cfg"
      menGeneralSettings_Click
    Case "playback"
      menPlayback_Click
    Case "addmus"
      On Error GoTo errh
      With cdg
        .DialogTitle = "Add Music"
        .filename = ""
        .Flags = cdlOFNExplorer Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .Filter = "All supported filetypes|*.mp3;*.ogg;*.wav;*.mid;*.midi;*.mod;*.s3m;*.it;*.xm|MPEG Layer 3|*.mp3|OGG Vorbis|*.ogg|Wave-files|*.wav|Midi|*.mid;*.midi|Modules|*.mod;*.it;*.xm;*.s3m"
        .ShowOpen
      End With
      On Error Resume Next
      With Prj
        ReDim Preserve .Mus(UBound(.Mus) + 1)
        With .Mus(UBound(.Mus))
          .sFile = cdg.filename
          .sID = LCase(sFilename(.sFile, efpFileNameAndExt))
          .eFlags = Prj.eMusFlags
          .bVol = 255
          lst(1).AddItem .sID
          lst(1).ListIndex = lst(1).ListCount - 1
        End With
      End With
      frmStart.Visible = False
      Prg.bChanged = True
    Case "remmus"
      Dim x As Long
      If Prg.bConfirm Then
        x = MsgBox("You are about to remove the current selected music, '" & Prj.Mus(lst(1).ListIndex + 1).sID & "'. This cannot be undone. Are you sure?", vbExclamation Or vbYesNo, "Confirm remove")
      Else
        x = vbYes
      End If
      If x = vbYes Then
        
        For x = lst(1).ListIndex + 1 To UBound(Prj.Mus) - 1
          With Prj.Mus(x)
            .bVol = Prj.Mus(x + 1).bVol
            .eFlags = Prj.Mus(x + 1).eFlags
            .sDescr = Prj.Mus(x + 1).sDescr
            .sFile = Prj.Mus(x + 1).sFile
            .sID = Prj.Mus(x + 1).sID
          End With
        Next
        
        ReDim Preserve Prj.Mus(UBound(Prj.Mus) - 1)
        
        x = lst(1).ListIndex
        lst(1).RemoveItem lst(1).ListIndex
        If x > lst(1).ListCount - 1 Then x = x - 1
        lst(1).ListIndex = x
        
        lst_Click 1
        
      End If
      Prg.bChanged = True
    Case "help"
      frmHelp.Show , Me
      frmHelp.web.Navigate App.Path & "\doc\mus.html"
  End Select
errh:
End Sub

Public Sub LoadSettings()
  Prg.bConfirm = GetSetting(App.ProductName, "settings", "confirm", 1)
  menConfirm.Checked = Prg.bConfirm
End Sub

Public Sub SaveSettings()
  SaveSetting App.ProductName, "settings", "confirm", Prg.bConfirm
End Sub

Private Sub txtMDescr_KeyUp(KeyCode As Integer, Shift As Integer)
  Prj.Mus(lst(1).ListIndex + 1).sDescr = txtMDescr
  Prg.bChanged = True
End Sub

Private Sub txtMID_KeyUp(KeyCode As Integer, Shift As Integer)
  Prj.Mus(lst(1).ListIndex + 1).sID = txtMID
  lst(1).List(lst(1).ListIndex) = txtMID
  Prg.bChanged = True
End Sub

Private Sub txtSDescr_KeyUp(KeyCode As Integer, Shift As Integer)
  Prj.Sfx(lst(0).ListIndex + 1).sDescr = txtSDescr
  Prg.bChanged = True
End Sub

Private Sub txtSID_KeyUp(KeyCode As Integer, Shift As Integer)
  Prj.Sfx(lst(0).ListIndex + 1).sID = txtSID
  lst(0).List(lst(0).ListIndex) = txtSID
  Prg.bChanged = True
End Sub

Private Sub txtSPback_KeyPress(KeyAscii As Integer)
  KeyAscii = ctlKeyPress(KeyAscii, NumbersOnly)
End Sub

Private Sub txtSPback_KeyUp(KeyCode As Integer, Shift As Integer)
  If Val(txtSPback) < 0 Then txtSPback = 0
  If Val(txtSPback) > 255 Then txtSPback = 255
  Prj.Sfx(lst(0).ListIndex + 1).bPlaybacks = Val(txtSPback)
  Prg.bChanged = True
End Sub

Private Sub txtSPriority_KeyPress(KeyAscii As Integer)
  KeyAscii = ctlKeyPress(KeyAscii, NumbersOnly)
End Sub

Public Sub SaveSfxFlags()
  Prg.bChanged = True
  With Prj.Sfx(lst(0).ListIndex + 1)
    .eFlags = 0
    If chkSLoop Then .eFlags = .eFlags Or Sfx1_Loop
  End With
End Sub

Public Sub SaveMusFlags()
  Prg.bChanged = True
  With Prj.Mus(lst(1).ListIndex + 1)
    .eFlags = 0
    If chkMLoop Then .eFlags = .eFlags Or Mus1_Loop
    'If chkMReverb Then .eFlags = .eFlags Or Mus2_Reverb
  End With
End Sub

Private Sub txtSPriority_KeyUp(KeyCode As Integer, Shift As Integer)
  If Val(txtSPriority) < 0 Then txtSPriority = 0
  If Val(txtSPriority) > 255 Then txtSPriority = 255
  Prj.Sfx(lst(0).ListIndex + 1).bPriority = Val(txtSPriority)
  Prg.bChanged = True
End Sub

Public Sub SaveProject(ByVal sF As String)
  On Error GoTo errh
  
  If FileExist(sF) Then
    If FileExist(sF & ".bak") Then Kill sF & ".bak"
    Name sF As sF & ".bak"
  End If
  
  Dim cF As New clsDatafile
  Dim x As Long
  
  cF.filename = sF
  cF.WriteStrFixed "PGESP" & Chr(PrjVer)
  
  sF = Left(sF, InStrRev(sF, "\"))
  
  With Prj
    cF.WriteStr .sID
    cF.WriteStr .sDescr
    cF.WriteStr .sFilenameCompile
    cF.WriteNumber .eSfxFlags
    cF.WriteNumber .eMusFlags
    cF.WriteNumber UBound(.Sfx)
  End With
  For x = 1 To UBound(Prj.Sfx)
    With Prj.Sfx(x)
      cF.WriteStr .sID
      cF.WriteStr TruncFilename(sF, .sFile)
      cF.WriteStr .sDescr
      cF.WriteNumber .eFlags
      cF.WriteNumber .bVol
      cF.WriteNumber .bPriority
      cF.WriteNumber .bPlaybacks
    End With
  Next
  cF.WriteNumber UBound(Prj.Mus)
  For x = 1 To UBound(Prj.Mus)
    With Prj.Mus(x)
      cF.WriteStr .sID
      cF.WriteStr TruncFilename(sF, .sFile)
      cF.WriteStr .sDescr
      cF.WriteNumber .eFlags
      cF.WriteNumber .bVol
    End With
  Next
  
  
  Exit Sub
errh:
  MsgBox "There was an error while saving to """ & cF.filename & """. Perhaps you should try again with an other filename or on an other drive.", vbExclamation, "Save Error"
End Sub

Public Function TruncFilename(ByVal loc As String, ByVal file As String) As String
  If InStr(1, LCase(file), LCase(loc)) Then
    TruncFilename = Right(file, Len(file) - Len(loc))
  Else
    TruncFilename = file
  End If
End Function

Public Sub OpenProject(ByVal sF As String)
  On Error GoTo errh
  
  Dim cF As New clsDatafile
  Dim x As Long, s As String
  
  cF.filename = sF
  s = cF.ReadStrFixed(6)
  If Left(s, 5) <> "PGESP" Then
    MsgBox "The selected file does not seem to be an PgeSound project file.", vbExclamation, "Error"
    Exit Sub
  End If
  If Asc(Right(s, 1)) <> PrjVer Then
    MsgBox "The selected project was saved with an earlier version of PgeSound. It cannot be opened.", vbExclamation, "Error"
    Exit Sub
  End If
  
  Prj.sFilename = sF
  
  sF = Left(sF, InStrRev(sF, "\"))
  
  With Prj
    .sID = cF.ReadStr
    .sDescr = cF.ReadStr
    .sFilenameCompile = cF.ReadStr
    .eSfxFlags = cF.ReadNumber
    .eMusFlags = cF.ReadNumber
    ReDim .Sfx(cF.ReadNumber)
  End With
  For x = 1 To UBound(Prj.Sfx)
    With Prj.Sfx(x)
      .sID = cF.ReadStr
      lst(0).AddItem .sID
      s = cF.ReadStr
      If FileExist(sF & s) Then
        .sFile = sF & s
      Else
        .sFile = s
      End If
      .sDescr = cF.ReadStr
      .eFlags = cF.ReadNumber
      .bVol = cF.ReadNumber
      .bPriority = cF.ReadNumber
      .bPlaybacks = cF.ReadNumber
    End With
  Next
  ReDim Prj.Mus(cF.ReadNumber)
  For x = 1 To UBound(Prj.Mus)
    With Prj.Mus(x)
      .sID = cF.ReadStr
      lst(1).AddItem .sID
      s = cF.ReadStr
      If FileExist(sF & s) Then
        .sFile = sF & s
      Else
        .sFile = s
      End If
      .sDescr = cF.ReadStr
      .eFlags = cF.ReadNumber
      .bVol = cF.ReadNumber
    End With
  Next
  
  On Error GoTo 0
  Me.Caption = "PgeSound [" & Prj.sID & "]"
  lst(0).ListIndex = lst(0).ListCount - 1
  lst(1).ListIndex = lst(1).ListCount - 1
  
  Exit Sub
errh:
  MsgBox "There was an error while opening """ & cF.filename & """.", vbExclamation, "Read Error"
End Sub

Public Sub CompileProject(ByVal sF As String, Optional bOk As Boolean = False)
  On Error GoTo errh
  
  Me.MousePointer = vbHourglass
  
  If FileExist(sF) Then
    If FileExist(sF & ".bak") Then Kill sF & ".bak"
    Name sF As sF & ".bak"
  End If
  
  Dim cF As New clsDatafile
  Dim x As Long
  
  cF.filename = sF
  cF.WriteStrFixed "PGECS" & Chr(ComVer)
    
  With Prj
    cF.WriteStr .sID
    cF.WriteStr .sDescr
    cF.WriteNumber UBound(.Sfx)
  End With
  For x = 1 To UBound(Prj.Sfx)
    With Prj.Sfx(x)
      cF.WriteFile .sFile
      cF.WriteStr .sID
      cF.WriteNumber .bVol
      cF.WriteNumber .bPriority
      cF.WriteNumber .eFlags
      cF.WriteNumber .bPlaybacks
    End With
  Next
  cF.WriteNumber UBound(Prj.Mus)
  For x = 1 To UBound(Prj.Mus)
    With Prj.Mus(x)
      Select Case LCase(sFilename(.sFile, efpFileExt))
        Case "mp3", "mp2", "ogg"
          cF.WriteNumber 1
        Case "mod", "mid", "midi", "xm", "it", "s3m"
          cF.WriteNumber 2
      End Select
      cF.WriteFile .sFile
      cF.WriteStr .sID
      cF.WriteNumber .eFlags
      cF.WriteNumber .bVol
    End With
  Next
  
  
  Me.MousePointer = vbDefault
  If bOk Then
    MsgBox "Compilation to """ & cF.filename & """ successful.", vbInformation, "Compiled"
    Prg.bChanged = False
  Else
    Prg.bChanged = True
  End If
  
  Exit Sub
errh:
  MsgBox "There was an error while saving to """ & cF.filename & """. Perhaps you should try again with an other filename or on an other drive.", vbExclamation, "Save Error"
  Me.MousePointer = vbDefault
End Sub


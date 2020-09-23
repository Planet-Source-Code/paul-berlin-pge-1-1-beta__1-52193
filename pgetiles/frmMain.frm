VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PgeTiles"
   ClientHeight    =   5430
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmStart 
      Caption         =   "PgeTiles"
      Height          =   5175
      Left            =   2640
      TabIndex        =   12
      Top             =   120
      Width           =   4815
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Freeware!"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1920
         Width           =   705
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmMain.frx":030A
         Top             =   430
         Width           =   480
      End
      Begin VB.Label lblPsc 
         AutoSize        =   -1  'True
         Caption         =   "http://pab.dyndns.org"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   1755
         MousePointer    =   10  'Up Arrow
         TabIndex        =   19
         Top             =   1560
         Width           =   1560
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Pge is available from "
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1560
         Width           =   1500
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "berlin_paul@hotmail.com"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   240
         MousePointer    =   10  'Up Arrow
         TabIndex        =   17
         Top             =   1200
         Width           =   1770
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "By Paul Berlin 2003 - For used with Pab Game Engine."
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   3840
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
         Left            =   2760
         TabIndex        =   15
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "PgeTiles"
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
         TabIndex        =   14
         Top             =   360
         Width           =   1845
      End
   End
   Begin VB.Frame frm 
      Caption         =   "Animation Settings"
      Height          =   5175
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   4815
      Visible         =   0   'False
      Begin VB.CommandButton cmdEditAnim 
         Caption         =   "Edit Animation"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3840
         Width           =   4575
      End
      Begin VB.CommandButton cmdAniFrames 
         Caption         =   "Edit Frames"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3360
         Width           =   4575
      End
      Begin VB.ComboBox cmbAniSub 
         Height          =   315
         ItemData        =   "frmMain.frx":0614
         Left            =   2640
         List            =   "frmMain.frx":0621
         TabIndex        =   21
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton cmdAniFlags 
         Caption         =   "Animation Flags (0)"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   4575
      End
      Begin VB.TextBox txtAniID 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   4575
      End
      Begin VB.TextBox txtAniDescr 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label lblsub 
         AutoSize        =   -1  'True
         Caption         =   "Number of subframes (enter to set):"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2925
         Width           =   2475
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Animation ID:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Animation Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.Frame frm 
      Caption         =   "Collection Settings"
      Height          =   5175
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Visible         =   0   'False
      Begin VB.TextBox txtColDescr 
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox txtColID 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Collection Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Collection ID:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   945
      End
   End
   Begin MSComctlLib.ImageList imlTreeIcons 
      Left            =   1200
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   8421376
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":062F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":072B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0827
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0923
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "addcol"
            Object.ToolTipText     =   "Add Collection"
            ImageKey        =   "addcol"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remcol"
            Object.ToolTipText     =   "Remove Collection"
            ImageKey        =   "remcol"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "addani"
            Object.ToolTipText     =   "Add Animation"
            ImageKey        =   "addani"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remani"
            Object.ToolTipText     =   "Remove Animation"
            ImageKey        =   "remani"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "addfrm"
            Object.ToolTipText     =   "Change General Settings"
            ImageKey        =   "cfg"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remfrm"
            Object.ToolTipText     =   "Edit Resources"
            ImageKey        =   "res"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdgMain 
      Left            =   1440
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.TreeView tvwMain 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   8493
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "imlTreeIcons"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1560
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8421376
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A1F
            Key             =   "addcol"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B31
            Key             =   "remcol"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C43
            Key             =   "addani"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D55
            Key             =   "remani"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E67
            Key             =   "addfrm"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F79
            Key             =   "remfrm"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":108B
            Key             =   "cfg"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":119D
            Key             =   "res"
         EndProperty
      EndProperty
   End
   Begin VB.Menu menFile 
      Caption         =   "&File"
      Begin VB.Menu menNew 
         Caption         =   "&New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu menLine0 
         Caption         =   "-"
      End
      Begin VB.Menu menOpen 
         Caption         =   "&Open Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu menLine1 
         Caption         =   "-"
      End
      Begin VB.Menu menSave 
         Caption         =   "&Save Project"
         Shortcut        =   ^S
      End
      Begin VB.Menu menSaveAs 
         Caption         =   "Save Project &as..."
      End
      Begin VB.Menu menLine2 
         Caption         =   "-"
      End
      Begin VB.Menu menCompile 
         Caption         =   "&Compile Project..."
      End
      Begin VB.Menu menLine3 
         Caption         =   "-"
      End
      Begin VB.Menu menExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menTools 
      Caption         =   "&Tools"
      Begin VB.Menu menGenSettings 
         Caption         =   "&General Settings..."
      End
      Begin VB.Menu menResources 
         Caption         =   "&Resources..."
      End
   End
   Begin VB.Menu menOptions 
      Caption         =   "&Options"
      Begin VB.Menu menConfirm 
         Caption         =   "&Confirm"
         Begin VB.Menu menCon 
            Caption         =   "Remove &Collection"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu menCon 
            Caption         =   "Remove &Animation"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu menCon 
            Caption         =   "Remove &Frame"
            Checked         =   -1  'True
            Index           =   2
         End
         Begin VB.Menu menCon 
            Caption         =   "Remove &Resource"
            Checked         =   -1  'True
            Index           =   3
         End
      End
   End
   Begin VB.Menu menHelp 
      Caption         =   "&Help"
      Begin VB.Menu menIndex 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu menline4 
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
Public lSelCol As Long
Public lSelAni As Long

Private Sub cmbAniSub_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn And Val(cmbAniSub.Text) > 0 Then
    With Prj.Coll(lSelCol).Anim(lSelAni)
      If Val(cmbAniSub.Text) < .lSubimages And UBound(.Frame) > 0 Then
        If MsgBox("Warning! Changing the ammount of subimages to an lower value than before will remove the higher subframes! Are you sure you want to do this?", vbExclamation Or vbYesNo, "Confirm subframe change") = vbNo Then
          cmbAniSub.Text = .lSubimages
          Exit Sub
        End If
      End If
      Dim x As Long
      .lSubimages = Val(cmbAniSub.Text)
      For x = 1 To UBound(.Frame)
        ReDim Preserve .Frame(x).img(.lSubimages)
      Next
      Prg.bChanged = True
    End With
  End If
End Sub

Private Sub cmbAniSub_LostFocus()
  cmbAniSub.Text = Prj.Coll(lSelCol).Anim(lSelAni).lSubimages
End Sub

Private Sub cmdAniFlags_Click()
  With Prj.Coll(lSelCol).Anim(lSelAni)
    Load frmFlags
    frmFlags.Tag = .eOptions
    frmFlags.Show vbModal, Me
    .eOptions = Val(frmFlags.Tag)
    frmFlags.Tag = "Y"
    Unload frmFlags
    cmdAniFlags.Caption = "Animation Flags (" & .eOptions & ")"
    If (.eOptions And SubImages_2) = SubImages_2 Then
      lblsub.Enabled = True
      cmbAniSub.Enabled = True
      cmbAniSub.Text = .lSubimages
    Else
      lblsub.Enabled = False
      cmbAniSub.Enabled = False
      cmbAniSub.Text = 1
    End If
    Prg.bChanged = True
  End With
End Sub

Private Sub cmdAniFrames_Click()
  frmFrames.Show vbModal, Me
  tvwMain.SelectedItem.Text = Prj.Coll(lSelCol).Anim(lSelAni).sID & " (" & UBound(Prj.Coll(lSelCol).Anim(lSelAni).Frame) & ")"
End Sub

Private Sub cmdEditAnim_Click()
  'On Error Resume Next
  frmAni.Show vbModal, Me
End Sub

Private Sub Form_Load()
  lblVer = App.Major & "." & App.Minor & "." & App.Revision
  LoadSettings
  If Not modAss.CheckAssociation("tpj", "PgeTiles.tpj") Then
    modAss.CreateFileType "PgeTiles.tpj", "PgeTiles Project", App.Path & "\pgetiles.exe,0"
    modAss.CreateFileTypeAction "PgeTiles.tpj", "Open", Chr(34) & App.Path & "\pgetiles.exe" & Chr(34) & " -o %1"
    modAss.CreateFileTypeAction "PgeTiles.tpj", "Compile", Chr(34) & App.Path & "\pgetiles.exe" & Chr(34) & " -c %1"
    modAss.CreateAssociation ".tpj", "PgeTiles.backup", "PgeTiles.tpj"
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
  If Prg.bChanged Then
    Dim x As Long
    x = MsgBox("You have made changes to the open project without saving. Do you want to save before exiting?", vbInformation Or vbYesNoCancel, "Discard changes?")
    If x = vbYes Then
      menSave_Click
    ElseIf x = vbCancel Then
      Cancel = 1
    End If
  End If
  SaveSettings
End Sub

Private Sub frmStart_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  lblEmail.FontUnderline = False
  lblPsc.FontUnderline = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  frmStart_MouseMove Button, Shift, x, Y
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  frmStart_MouseMove Button, Shift, x, Y
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  frmStart_MouseMove Button, Shift, x, Y
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  frmStart_MouseMove Button, Shift, x, Y
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  lblEmail.FontUnderline = False
  lblPsc.FontUnderline = False
End Sub

Private Sub lblEmail_Click()
  ShellExecute Me.hwnd, "open", "mailto:berlin_paul@hotmail.com?subject=PgeTiles", vbNullString, vbNullString, 1
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Not lblEmail.FontUnderline Then
    lblEmail.FontUnderline = True
  End If
End Sub

Private Sub lblPsc_Click()
  ShellExecute Me.hwnd, "open", "http://pab.dyndns.org", vbNullString, vbNullString, 1
End Sub

Private Sub lblPsc_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Not lblPsc.FontUnderline Then
    lblPsc.FontUnderline = True
  End If
End Sub

Private Sub lblVer_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  frmStart_MouseMove Button, Shift, x, Y
End Sub

Private Sub menAbout_Click()
  frm(0).Visible = False
  frm(1).Visible = False
End Sub

Private Sub menCompile_Click()
  On Error GoTo errh
  With cdgMain
    .DialogTitle = "Compile Project"
    .FileName = Prj.sFilenameCompile
    .flags = cdlOFNExplorer Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    .Filter = "Compiled tilset (*.pbt)|*.pbt"
    .ShowSave
    
    Prj.sFilenameCompile = .FileName
    CompileProject .FileName
  End With
errh:
End Sub

Private Sub menCon_Click(Index As Integer)
  menCon(Index).Checked = Not menCon(Index).Checked
  Prg.bConfirm(Index) = menCon(Index).Checked
End Sub

Private Sub menExit_Click()
  Unload Me
End Sub

Private Sub menIndex_Click()
  frmHelp.Show , Me
  frmHelp.web.Navigate App.Path & "\doc\index.html"
End Sub

Private Sub menOpen_Click()
  If Prg.bChanged Then
    If MsgBox("You have made changes to the open project without saving. Are you sure you want to discard them?", vbOKCancel Or vbExclamation, "Discard changes?") = vbCancel Then Exit Sub
  End If
  On Error GoTo errh
  With cdgMain
    .DialogTitle = "Open Project"
    .FileName = ""
    .flags = cdlOFNExplorer Or cdlOFNNoReadOnlyReturn Or cdlOFNPathMustExist Or cdlOFNFileMustExist Or cdlOFNHideReadOnly
    .Filter = "PgeTiles Project file (*.tpj)|*.tpj"
    .ShowOpen
    
    OpenProject .FileName
  End With
errh:
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
  With cdgMain
    .DialogTitle = "Save Project"
    .FileName = Prj.sFilename
    .flags = cdlOFNExplorer Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    .Filter = "PgeTiles Project file (*.tpj)|*.tpj"
    .ShowSave
    
    Prj.sFilename = .FileName
    SaveProject .FileName
    Prg.bChanged = False
  End With
errh:
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComCtlLib.Button)
  On Error Resume Next
  Dim NewNode As Node, x As Long, Y As Long, z As Long, w As Long
  Select Case Button.Key
    Case "addcol"
      ReDim Preserve Prj.Coll(UBound(Prj.Coll) + 1)
      With Prj.Coll(UBound(Prj.Coll))
        .sID = "Collection " & UBound(Prj.Coll)
        ReDim .Anim(0)
        Set NewNode = tvwMain.Nodes.Add(, , "c" & UBound(Prj.Coll), .sID & " (0)", 1, 2)
        NewNode.Selected = True
        tvwMain_NodeClick NewNode
      End With
      Prg.bChanged = True
    Case "remcol"
      If Prg.bConfirm(0) Then
        x = MsgBox("You are about to remove the current selected collection, '" & Prj.Coll(lSelCol).sID & "'. That will remove all it's animations and cannot be undone. Are you sure?", vbExclamation Or vbYesNo, "Confirm remove")
      Else
        x = vbYes
      End If
      If x = vbYes Then
        
        Prg.bChanged = True
        'Overwrite the removed collection with data from following collections,
        'and remove the last item.
        For x = lSelCol To UBound(Prj.Coll) - 1
          With Prj.Coll(x)
            .sDescr = Prj.Coll(x + 1).sDescr
            .sID = Prj.Coll(x + 1).sID
            ReDim .Anim(UBound(Prj.Coll(x + 1).Anim))
            For Y = 1 To UBound(.Anim)
              .Anim(Y).eOptions = Prj.Coll(x + 1).Anim(Y).eOptions
              .Anim(Y).lSubimages = Prj.Coll(x + 1).Anim(Y).lSubimages
              .Anim(Y).sDescr = Prj.Coll(x + 1).Anim(Y).sDescr
              .Anim(Y).sID = Prj.Coll(x + 1).Anim(Y).sID
              ReDim .Anim(Y).Frame(UBound(Prj.Coll(x + 1).Anim(Y).Frame))
              For z = 1 To UBound(.Anim(Y).Frame)
                .Anim(Y).Frame(z).Ctrl = Prj.Coll(x + 1).Anim(Y).Frame(z).Ctrl
                ReDim .Anim(Y).Frame(z).img(.Anim(Y).lSubimages)
                For w = 1 To .Anim(Y).lSubimages
                  .Anim(Y).Frame(z).img(w).lRes = Prj.Coll(x + 1).Anim(Y).Frame(z).img(w).lRes
                  .Anim(Y).Frame(z).img(w).Pos = Prj.Coll(x + 1).Anim(Y).Frame(z).img(w).Pos
                  .Anim(Y).Frame(z).img(w).VirtualSize = Prj.Coll(x + 1).Anim(Y).Frame(z).img(w).VirtualSize
                Next
              Next
            Next
          End With
        Next
        
        ReDim Preserve Prj.Coll(UBound(Prj.Coll) - 1)
        'Find and remove the collection node
        'SelectedItem cannot be used, as animation nodes ca also be selected when removing collections
        For x = 1 To tvwMain.Nodes.Count
          If InStr(1, tvwMain.Nodes(x).Key, "a") = 0 Then
            Y = Val(Right(tvwMain.Nodes(x).Key, Len(tvwMain.Nodes(x).Key) - 1))
            If Y = lSelCol Then tvwMain.Nodes.Remove tvwMain.Nodes(x).Index
          End If
        Next
        
        'go through each node and change it's key
        For x = 1 To tvwMain.Nodes.Count
          With tvwMain.Nodes(x)
            If InStr(1, .Key, "a") Then
              z = Val(Right(.Key, Len(.Key) - InStrRev(.Key, "a")))
              Y = Val(Right(.Parent.Key, Len(.Parent.Key) - 1))
              If Y > lSelCol Then .Key = "c" & Y - 1 & "a" & z
            Else
              Y = Val(Right(.Key, Len(.Key) - 1))
              If Y > lSelCol Then .Key = "c" & Y - 1
            End If
          End With
        Next
        
        
        If tvwMain.Nodes.Count > 0 Then
          tvwMain_NodeClick tvwMain.SelectedItem
        Else
          lSelCol = 0
          frm(0).Visible = False
          frm(1).Visible = False
          tbrMain.Buttons(2).Enabled = False
          tbrMain.Buttons(4).Enabled = False
          tbrMain.Buttons(5).Enabled = False
        End If
        
        
      End If
    Case "addani"
      With Prj.Coll(lSelCol)
        ReDim Preserve .Anim(UBound(.Anim) + 1)
        .Anim(UBound(.Anim)).eOptions = Prj.eDefOptions
        .Anim(UBound(.Anim)).sID = "Animation " & UBound(.Anim)
        .Anim(UBound(.Anim)).lSubimages = 1
        ReDim .Anim(UBound(.Anim)).Frame(0)
        Set NewNode = tvwMain.Nodes.Add("c" & lSelCol, tvwChild, "c" & lSelCol & "a" & UBound(.Anim), .Anim(UBound(.Anim)).sID & " (0)", 3, 4)
        NewNode.Parent.Text = Prj.Coll(lSelCol).sID & " (" & UBound(.Anim) & ")"
        NewNode.Selected = True
        tvwMain_NodeClick NewNode
      End With
      Prg.bChanged = True
    Case "remani"
      If Prg.bConfirm(1) Then
        x = MsgBox("You are about to remove the current selected animation, '" & Prj.Coll(lSelCol).Anim(lSelAni).sID & "'. That will remove all it's frames and cannot be undone. Are you sure?", vbExclamation Or vbYesNo, "Confirm remove")
      Else
        x = vbYes
      End If
      If x = vbYes Then
        
        Prg.bChanged = True
        'Overwrite the removed animation with data from following animations,
        'and remove the last item.
        For x = lSelAni To UBound(Prj.Coll(lSelCol).Anim) - 1
          With Prj.Coll(lSelCol).Anim(x)
            .eOptions = Prj.Coll(lSelCol).Anim(x + 1).eOptions
            .lSubimages = Prj.Coll(lSelCol).Anim(x + 1).lSubimages
            .sDescr = Prj.Coll(lSelCol).Anim(x + 1).sDescr
            .sID = Prj.Coll(lSelCol).Anim(x + 1).sID
            ReDim .Frame(UBound(Prj.Coll(lSelCol).Anim(x + 1).Frame))
            For Y = 1 To UBound(.Frame)
              ReDim .Frame(Y).img(.lSubimages)
              .Frame(Y).Ctrl = .Frame(Y + 1).Ctrl
              For z = 1 To .lSubimages
                .Frame(Y).img(z).lRes = Prj.Coll(lSelCol).Anim(x + 1).Frame(Y).img(z).lRes
                .Frame(Y).img(z).Pos = Prj.Coll(lSelCol).Anim(x + 1).Frame(Y).img(z).Pos
                .Frame(Y).img(z).VirtualSize = Prj.Coll(lSelCol).Anim(x + 1).Frame(Y).img(z).VirtualSize
              Next
            Next
          End With
        Next
        
        ReDim Preserve Prj.Coll(lSelCol).Anim(UBound(Prj.Coll(lSelCol).Anim) - 1)
        'Find and remove the animation node
        tvwMain.Nodes.Remove tvwMain.SelectedItem.Index
        
        'go through each node and change it's key
        For x = 1 To tvwMain.Nodes.Count
          With tvwMain.Nodes(x)
            If InStr(1, .Key, "a") Then
              z = Val(Right(.Key, Len(.Key) - InStrRev(.Key, "a")))
              Y = Val(Right(.Parent.Key, Len(.Parent.Key) - 1))
              If Y = lSelCol And z > lSelAni Then .Key = "c" & Y & "a" & z - 1
            End If
          End With
        Next
        
        
        If tvwMain.Nodes.Count > 0 Then
          tvwMain_NodeClick tvwMain.SelectedItem
        Else
          lSelCol = 0
          lSelAni = 0
          frm(0).Visible = False
          frm(1).Visible = False
          tbrMain.Buttons(2).Enabled = False
          tbrMain.Buttons(4).Enabled = False
          tbrMain.Buttons(5).Enabled = False
        End If
        
        
      End If
    Case "addfrm" 'general settings
      menGenSettings_Click
    Case "remfrm" 'resources
      menResources_Click
  End Select
End Sub

Private Sub menGenSettings_Click()
  frmGenSet.Show vbModal, Me
  Me.Caption = "PgeTiles [" & Prj.sID & "]"
End Sub

Private Sub menNew_Click()
  If Prg.bChanged Then
    If MsgBox("You have made changes to the open project without saving. Are you sure you want to discard them?", vbOKCancel Or vbExclamation, "Discard changes?") = vbCancel Then Exit Sub
  End If
  
  ReDim Prj.Res(0)
  ReDim Prj.Coll(0)
  Prj.eDefOptions = 0 Or Loop_4 Or Center_8
  Prj.sFilename = ""
  Prj.sDescr = "Untitled Tileset" & vbCrLf & "----------------" & vbCrLf & "Copyright © Creator." & vbCrLf & "Created " & Now & " with PgeTiles v" & App.Major & "." & App.Minor & "." & App.Revision & "., written by Paul Berlin."
  Prj.sFilenameCompile = ""
  Prj.sID = "Untitled Tileset"
   
  Me.Caption = "PgeTiles [" & Prj.sID & "]"
  tbrMain.Buttons(2).Enabled = False
  tbrMain.Buttons(4).Enabled = False
  tbrMain.Buttons(5).Enabled = False
  frm(0).Visible = False
  frm(1).Visible = False
  RefreshTree
End Sub

Private Sub menResources_Click()
  frmResources.Show vbModal, Me
End Sub

Public Sub RefreshTree()
  Dim NewNode As Node, x As Long, Y As Long
  tvwMain.Nodes.Clear
  For x = 1 To UBound(Prj.Coll)
    Set NewNode = tvwMain.Nodes.Add(, , "c" & x, Prj.Coll(x).sID, 1, 2)
    For Y = 1 To UBound(Prj.Coll(x).Anim)
      With Prj.Coll(x)
        Set NewNode = tvwMain.Nodes.Add("c" & x, tvwChild, "c" & x & "a" & Y, .Anim(Y).sID & " (" & UBound(.Anim(Y).Frame) & ")", 3, 4)
      End With
    Next
  Next
  
  tbrMain.Buttons(2).Enabled = False
  tbrMain.Buttons(4).Enabled = False
  tbrMain.Buttons(5).Enabled = False
End Sub

Private Sub tvwMain_NodeClick(ByVal Node As MSComCtlLib.Node)
  Dim xC As Long, xA As Long
  If InStr(1, Node.Key, "a") Then 'animation selected
    lSelCol = Val(Right(Node.Parent.Key, Len(Node.Parent.Key) - 1))
    lSelAni = Val(Right(Node.Key, Len(Node.Key) - InStrRev(Node.Key, "a")))
    frm(0).Visible = False
    frm(1).Visible = True
    With Prj.Coll(lSelCol).Anim(lSelAni)
      txtAniID = .sID
      txtAniDescr = .sDescr
      cmdAniFlags.Caption = "Animation Flags (" & .eOptions & ")"
      If (.eOptions And SubImages_2) = SubImages_2 Then
        lblsub.Enabled = True
        cmbAniSub.Enabled = True
        cmbAniSub.Text = .lSubimages
      Else
        lblsub.Enabled = False
        cmbAniSub.Enabled = False
        cmbAniSub.Text = 1
      End If
    End With
    tbrMain.Buttons(2).Enabled = True
    tbrMain.Buttons(4).Enabled = True
    tbrMain.Buttons(5).Enabled = True
    
  Else 'collection selected
    lSelCol = Val(Right(Node.Key, Len(Node.Key) - 1))
    frm(0).Visible = True
    frm(1).Visible = False
    txtColID = Prj.Coll(lSelCol).sID
    txtColDescr = Prj.Coll(lSelCol).sDescr
    tbrMain.Buttons(2).Enabled = True
    tbrMain.Buttons(4).Enabled = True
    tbrMain.Buttons(5).Enabled = False
  End If
End Sub

Private Sub txtAniDescr_KeyUp(KeyCode As Integer, Shift As Integer)
  Prj.Coll(lSelCol).Anim(lSelAni).sDescr = txtAniDescr
  Prg.bChanged = True
End Sub

Private Sub txtAniID_KeyUp(KeyCode As Integer, Shift As Integer)
  Prj.Coll(lSelCol).Anim(lSelAni).sID = txtAniID
  tvwMain.SelectedItem.Text = txtAniID & " (" & UBound(Prj.Coll(lSelCol).Anim(lSelAni).Frame) & ")"
  tvwMain.Sorted = False
  tvwMain.Sorted = True
  Prg.bChanged = True
End Sub

Private Sub txtColDescr_KeyUp(KeyCode As Integer, Shift As Integer)
  Prj.Coll(lSelCol).sDescr = txtColDescr
  Prg.bChanged = True
End Sub

Private Sub txtColID_KeyUp(KeyCode As Integer, Shift As Integer)
  Prj.Coll(lSelCol).sID = txtColID
  tvwMain.SelectedItem.Text = txtColID & " (" & UBound(Prj.Coll(lSelCol).Anim) & ")"
  tvwMain.Sorted = False
  tvwMain.Sorted = True
  Prg.bChanged = True
End Sub

Public Sub LoadSettings()
  Prg.bConfirm(0) = GetSetting(App.ProductName, "settings", "confirm0", 1)
  Prg.bConfirm(1) = GetSetting(App.ProductName, "settings", "confirm1", 1)
  Prg.bConfirm(2) = GetSetting(App.ProductName, "settings", "confirm2", 1)
  Prg.bConfirm(3) = GetSetting(App.ProductName, "settings", "confirm3", 1)
  Prg.bShowAuto = GetSetting(App.ProductName, "settings", "showauto", 1)
  Prg.bSnap = GetSetting(App.ProductName, "settings", "snap", 1)
  Prg.bSnapAuto = GetSetting(App.ProductName, "settings", "snapauto", 1)
  Prg.bShowInView = GetSetting(App.ProductName, "settings", "showinview", 1)
  Prg.lFrmBgCol = GetSetting(App.ProductName, "settings", "frmbgcol", vbWhite)
  Prg.lFH = GetSetting(App.ProductName, "settings", "fh", 7000)
  Prg.lFW = GetSetting(App.ProductName, "settings", "fw", 10000)
  Prg.bAniAutochange = GetSetting(App.ProductName, "settings", "aniautochange", 0)
  Prg.lAutoX(0) = GetSetting(App.ProductName, "settings", "x0", 0)
  Prg.lAutoX(1) = GetSetting(App.ProductName, "settings", "x1", 64)
  Prg.lAutoX(2) = GetSetting(App.ProductName, "settings", "x2", 0)
  Prg.lAutoX(3) = GetSetting(App.ProductName, "settings", "x3", 2)
  Prg.lAutoY(0) = GetSetting(App.ProductName, "settings", "y0", 0)
  Prg.lAutoY(1) = GetSetting(App.ProductName, "settings", "y1", 64)
  Prg.lAutoY(2) = GetSetting(App.ProductName, "settings", "y2", 0)
  Prg.lAutoY(3) = GetSetting(App.ProductName, "settings", "y3", 2)
  Prg.bAutoFrame = GetSetting(App.ProductName, "settings", "autoframe", 0)
  
  Dim x As Long
  For x = 0 To 3
    menCon(x).Checked = Prg.bConfirm(x)
  Next
End Sub

Public Sub SaveSettings()
  SaveSetting App.ProductName, "settings", "confirm0", Prg.bConfirm(0)
  SaveSetting App.ProductName, "settings", "confirm1", Prg.bConfirm(1)
  SaveSetting App.ProductName, "settings", "confirm2", Prg.bConfirm(2)
  SaveSetting App.ProductName, "settings", "confirm3", Prg.bConfirm(3)
  SaveSetting App.ProductName, "settings", "showauto", Prg.bShowAuto
  SaveSetting App.ProductName, "settings", "snap", Prg.bSnap
  SaveSetting App.ProductName, "settings", "snapauto", Prg.bSnapAuto
  SaveSetting App.ProductName, "settings", "showinview", Prg.bShowInView
  SaveSetting App.ProductName, "settings", "frmbgcol", Prg.lFrmBgCol
  SaveSetting App.ProductName, "settings", "fh", Prg.lFH
  SaveSetting App.ProductName, "settings", "fw", Prg.lFW
  SaveSetting App.ProductName, "settings", "aniautochange", Prg.bAniAutochange
  SaveSetting App.ProductName, "settings", "x0", Prg.lAutoX(0)
  SaveSetting App.ProductName, "settings", "x1", Prg.lAutoX(1)
  SaveSetting App.ProductName, "settings", "x2", Prg.lAutoX(2)
  SaveSetting App.ProductName, "settings", "x3", Prg.lAutoX(3)
  SaveSetting App.ProductName, "settings", "y0", Prg.lAutoY(0)
  SaveSetting App.ProductName, "settings", "y1", Prg.lAutoY(1)
  SaveSetting App.ProductName, "settings", "y2", Prg.lAutoY(2)
  SaveSetting App.ProductName, "settings", "y3", Prg.lAutoY(3)
  SaveSetting App.ProductName, "settings", "autoframe", Prg.bAutoFrame
End Sub

Public Sub SaveProject(ByVal sF As String)
  On Error GoTo errh
  
  If FileExist(sF) Then
    If FileExist(sF & ".bak") Then Kill sF & ".bak"
    Name sF As sF & ".bak"
  End If
  
  Dim cF As New clsDatafile
  Dim x As Long, Y As Long, z As Long, w As Long
  
  cF.FileName = sF
  cF.WriteStrFixed "PGEP" & Chr(PrjVer)
  
  sF = Left(sF, InStrRev(sF, "\"))
  
  With Prj
    cF.WriteStr .sID
    cF.WriteStr .sDescr
    cF.WriteStr .sFilenameCompile
    cF.WriteNumber .eDefOptions
    cF.WriteNumber UBound(.Res)
  End With
  For x = 1 To UBound(Prj.Res)
    With Prj.Res(x)
      cF.WriteStr TruncFilename(sF, .sFilename)
      cF.WriteNumber .lTranscolor
    End With
  Next
  cF.WriteNumber UBound(Prj.Coll)
  For x = 1 To UBound(Prj.Coll)
    With Prj.Coll(x)
      cF.WriteStr .sID
      cF.WriteStr .sDescr
      cF.WriteNumber UBound(.Anim)
      For Y = 1 To UBound(.Anim)
        With .Anim(Y)
          cF.WriteStr .sID
          cF.WriteStr .sDescr
          cF.WriteNumber .eOptions
          If (.eOptions And SubImages_2) = SubImages_2 Then
            cF.WriteNumber .lSubimages
          End If
          cF.WriteNumber UBound(.Frame)
          For z = 1 To UBound(.Frame)
            With .Frame(z)
              cF.WriteNumber .Ctrl.lDelay
              For w = 1 To UBound(.img)
                With .img(w)
                  cF.WriteNumber .lRes
                  cF.WriteNumber CInt(.Offset.x)
                  cF.WriteNumber CInt(.Offset.Y)
                  cF.WriteNumber .Pos.Left
                  cF.WriteNumber .Pos.Right
                  cF.WriteNumber .Pos.Top
                  cF.WriteNumber .Pos.bottom
                  If (Prj.Coll(x).Anim(Y).eOptions And VirtualSize_1) = VirtualSize_1 Then
                    cF.WriteNumber .VirtualSize.Left
                    cF.WriteNumber .VirtualSize.Right
                    cF.WriteNumber .VirtualSize.Top
                    cF.WriteNumber .VirtualSize.bottom
                  End If
                End With
              Next
            End With
          Next z
        End With
      Next
    End With
  Next
  
  
  Exit Sub
errh:
  MsgBox "There was an error while saving to """ & cF.FileName & """. Perhaps you should try again with an other filename or on an other drive.", vbExclamation, "Save Error"
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
  Dim x As Long, Y As Long, z As Long, w As Long
  Dim s As String
  
  cF.FileName = sF
  s = cF.ReadStrFixed(5)
  If Left(s, 4) <> "PGEP" Then
    MsgBox "The selected file does not seem to be an PgeTiles project file.", vbExclamation, "Error"
    Exit Sub
  End If
  If Asc(Right(s, 1)) <> PrjVer Then
    MsgBox "The selected project was saved with an earlier version of PgeTiles. It cannot be opened.", vbExclamation, "Error"
    Exit Sub
  End If
  
  Prj.sFilename = sF
  Prg.bChanged = False
  
  sF = Left(sF, InStrRev(sF, "\"))
  
  With Prj
    .sID = cF.ReadStr
    .sDescr = cF.ReadStr
    .sFilenameCompile = cF.ReadStr
    .eDefOptions = cF.ReadNumber
    ReDim .Res(cF.ReadNumber)
  End With
  For x = 1 To UBound(Prj.Res)
    With Prj.Res(x)
      s = cF.ReadStr
      If FileExist(sF & s) Then
        .sFilename = sF & s
      Else
        .sFilename = s
      End If
      .lTranscolor = cF.ReadNumber
    End With
  Next
  ReDim Prj.Coll(cF.ReadNumber)
  For x = 1 To UBound(Prj.Coll)
    With Prj.Coll(x)
      .sID = cF.ReadStr
      .sDescr = cF.ReadStr
      ReDim .Anim(cF.ReadNumber)
      For Y = 1 To UBound(.Anim)
        With .Anim(Y)
          .sID = cF.ReadStr
          .sDescr = cF.ReadStr
          .eOptions = cF.ReadNumber
          If (.eOptions And SubImages_2) = SubImages_2 Then
            .lSubimages = cF.ReadNumber
          Else
            .lSubimages = 1
          End If
          ReDim .Frame(cF.ReadNumber)
          For z = 1 To UBound(.Frame)
            With .Frame(z)
              .Ctrl.lDelay = cF.ReadNumber
              ReDim .img(Prj.Coll(x).Anim(Y).lSubimages)
              For w = 1 To UBound(.img)
                With .img(w)
                  .lRes = cF.ReadNumber
                  .Offset.x = cF.ReadNumber
                  .Offset.Y = cF.ReadNumber
                  .Pos.Left = cF.ReadNumber
                  .Pos.Right = cF.ReadNumber
                  .Pos.Top = cF.ReadNumber
                  .Pos.bottom = cF.ReadNumber
                  If (Prj.Coll(x).Anim(Y).eOptions And VirtualSize_1) = VirtualSize_1 Then
                    .VirtualSize.Left = cF.ReadNumber
                    .VirtualSize.Right = cF.ReadNumber
                    .VirtualSize.Top = cF.ReadNumber
                    .VirtualSize.bottom = cF.ReadNumber
                  Else
                    .VirtualSize.Left = 0
                    .VirtualSize.Right = 0
                    .VirtualSize.Top = 0
                    .VirtualSize.bottom = 0
                  End If
                End With
              Next
            End With
          Next z
        End With
      Next
    End With
  Next

  On Error GoTo 0
  Me.Caption = "PgeTiles [" & Prj.sID & "]"
  RefreshTree
  'tvwMain_NodeClick tvwMain.SelectedItem
  
  
  Exit Sub
errh:
  MsgBox "There was an error while opening """ & cF.FileName & """.", vbExclamation, "Read Error"
End Sub

Public Sub CompileProject(ByVal sF As String, Optional bOk As Boolean = False)
  On Error GoTo errh
  
  Me.MousePointer = vbHourglass
  
  If FileExist(sF) Then
    If FileExist(sF & ".bak") Then Kill sF & ".bak"
    Name sF As sF & ".bak"
  End If
  
  Dim cF As New clsDatafile
  Dim x As Long, Y As Long, z As Long, w As Long
  
  cF.FileName = sF
  cF.WriteStrFixed "PGEC" & Chr(ComVer)
    
  With Prj
    cF.WriteStr .sID
    cF.WriteStr .sDescr
    cF.WriteNumber UBound(.Res)
  End With
  For x = 1 To UBound(Prj.Res)
    With Prj.Res(x)
      cF.WriteFile .sFilename
      cF.WriteNumber D3DColorRGBA(.lTranscolor Mod &H100, (.lTranscolor \ &H100) Mod &H100, (.lTranscolor \ &H10000) Mod &H100, 255)
    End With
  Next
  x = 0
  For x = 1 To UBound(Prj.Coll)
    x = x + UBound(Prj.Coll(x).Anim)
  Next
  cF.WriteNumber x
  For x = 1 To UBound(Prj.Coll)
    With Prj.Coll(x)
      For Y = 1 To UBound(.Anim)
        With .Anim(Y)
          cF.WriteStr .sID
          cF.WriteNumber .eOptions
          If (.eOptions And SubImages_2) = SubImages_2 Then
            cF.WriteNumber .lSubimages
          End If
          cF.WriteNumber UBound(.Frame)
          For z = 1 To UBound(.Frame)
            With .Frame(z)
              cF.WriteNumber .Ctrl.lDelay
              For w = 1 To UBound(.img)
                With .img(w)
                  cF.WriteNumber .lRes
                  cF.WriteNumber CInt(.Offset.x)
                  cF.WriteNumber CInt(.Offset.Y)
                  cF.WriteNumber .Pos.Left
                  cF.WriteNumber .Pos.Right
                  cF.WriteNumber .Pos.Top
                  cF.WriteNumber .Pos.bottom
                  If (Prj.Coll(x).Anim(Y).eOptions And VirtualSize_1) = VirtualSize_1 Then
                    cF.WriteNumber .VirtualSize.Left - .Pos.Left
                    cF.WriteNumber .VirtualSize.Right - .Pos.Right
                    cF.WriteNumber .VirtualSize.Top - .Pos.Top
                    cF.WriteNumber .VirtualSize.bottom - .Pos.bottom
                  End If
                End With
              Next
            End With
          Next z
        End With
      Next
    End With
  Next
  
  Me.MousePointer = vbDefault
  If bOk Then
    MsgBox "Compilation to """ & cF.FileName & """ successful.", vbInformation, "Compiled"
  End If
  
  Exit Sub
errh:
  MsgBox "There was an error while saving to """ & cF.FileName & """. Perhaps you should try again with an other filename or on an other drive.", vbExclamation, "Save Error"
  Me.MousePointer = vbDefault
End Sub

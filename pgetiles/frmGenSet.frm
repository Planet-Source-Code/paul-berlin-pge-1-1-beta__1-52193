VERSION 5.00
Begin VB.Form frmGenSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "General Settings"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDescr 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton cmdFlags 
      Caption         =   "Default Animation Flags (0)"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   4455
   End
   Begin VB.TextBox txtTitle 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Untitled"
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Description/Notes/Copyrights:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2145
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tileset title:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "frmGenSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFlags_Click()
  Load frmFlags
  frmFlags.Tag = Prj.eDefOptions
  frmFlags.Show vbModal, Me
  Prj.eDefOptions = Val(frmFlags.Tag)
  frmFlags.Tag = "Y"
  Unload frmFlags
  RefreshSettings
  Prg.bChanged = True
End Sub

Private Sub Form_Load()
  RefreshSettings
  txtTitle.SelLength = Len(txtTitle)
End Sub

Public Sub RefreshSettings()
  txtTitle = Prj.sID
  txtDescr = Prj.sDescr
  cmdFlags.Caption = "Default Animation Flags (" & Prj.eDefOptions & ")"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Prj.sID = txtTitle
  Prj.sDescr = txtDescr
End Sub

Private Sub txtDescr_KeyUp(KeyCode As Integer, Shift As Integer)
  Prg.bChanged = True
End Sub

Private Sub txtTitle_KeyUp(KeyCode As Integer, Shift As Integer)
  Prg.bChanged = True
End Sub

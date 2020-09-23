VERSION 5.00
Begin VB.Form frmGenSet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "General Settings"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFlags 
      Caption         =   "Default Music Flags (0)"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   4455
   End
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
      Caption         =   "Default Sound Effect Flags (0)"
      Height          =   375
      Index           =   0
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
      Caption         =   "Soundset title:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "frmGenSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFlags_Click(index As Integer)
  Select Case index
  Case 0
    Load frmFlags
    frmFlags.Tag = Prj.eSfxFlags
    frmFlags.Show vbModal, Me
    Prj.eSfxFlags = Val(frmFlags.Tag)
    frmFlags.Tag = "Y"
    Unload frmFlags
  Case 1
'    Load frmFlags2
'    frmFlags2.Tag = Prj.eMusFlags
'    frmFlags2.Show vbModal, Me
'    Prj.eMusFlags = Val(frmFlags2.Tag)
'    frmFlags2.Tag = "Y"
'    Unload frmFlags2
    Load frmFlags
    frmFlags.Tag = Prj.eSfxFlags
    frmFlags.Show vbModal, Me
    Prj.eMusFlags = Val(frmFlags.Tag)
    frmFlags.Tag = "Y"
    Unload frmFlags
  End Select
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
  cmdFlags(0).Caption = "Default Sound Effect Flags (" & Prj.eSfxFlags & ")"
  cmdFlags(1).Caption = "Default Music Flags (" & Prj.eMusFlags & ")"
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

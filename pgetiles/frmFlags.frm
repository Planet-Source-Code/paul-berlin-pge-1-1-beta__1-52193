VERSION 5.00
Begin VB.Form frmFlags 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Flags (0)"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Available Flags"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.CheckBox chkFlags 
         Caption         =   "Default to center frames"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CheckBox chkFlags 
         Caption         =   "Default to loop on"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox chkFlags 
         Caption         =   "Include Subframes"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox chkFlags 
         Caption         =   "Include Virtual Size"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Animation behavior:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlags As enumOptions

Private Sub chkFlags_Click(Index As Integer)
  Dim t As enumOptions
  If chkFlags(0) Then t = t Or VirtualSize_1
  If chkFlags(1) Then t = t Or SubImages_2
  If chkFlags(2) Then t = t Or Loop_4
  If chkFlags(3) Then t = t Or Center_8
  Me.Caption = "Flags (" & t & ")"
End Sub

Private Sub Form_Activate()
  lFlags = Val(Me.Tag)
  Me.Caption = "Flags (" & lFlags & ")"
  chkFlags(0).Value = IIf((lFlags And VirtualSize_1) = VirtualSize_1, 1, 0)
  chkFlags(1).Value = IIf((lFlags And SubImages_2) = SubImages_2, 1, 0)
  chkFlags(2).Value = IIf((lFlags And Loop_4) = Loop_4, 1, 0)
  chkFlags(3).Value = IIf((lFlags And Center_8) = Center_8, 1, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Me.Tag <> "Y" Then
    lFlags = 0
    Cancel = 1
    Me.Hide
    If chkFlags(0) Then lFlags = lFlags Or VirtualSize_1
    If chkFlags(1) Then lFlags = lFlags Or SubImages_2
    If chkFlags(2) Then lFlags = lFlags Or Loop_4
    If chkFlags(3) Then lFlags = lFlags Or Center_8
    Me.Tag = lFlags
  End If
End Sub


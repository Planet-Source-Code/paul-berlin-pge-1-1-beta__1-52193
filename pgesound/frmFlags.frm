VERSION 5.00
Begin VB.Form frmFlags 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Flags (0)"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   2535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Available Flags"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.CheckBox chkFlags 
         Caption         =   "Loop"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmFlags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFlags As enumSfx

Private Sub chkFlags_Click(Index As Integer)
  Dim t As enumSfx
  If chkFlags(0) Then t = t Or Sfx1_Loop
  Me.Caption = "Flags (" & t & ")"
End Sub

Private Sub Form_Activate()
  lFlags = Val(Me.Tag)
  Me.Caption = "Flags (" & lFlags & ")"
  chkFlags(0).Value = IIf((lFlags And Sfx1_Loop) = Sfx1_Loop, 1, 0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Me.Tag <> "Y" Then
    lFlags = 0
    Cancel = 1
    Me.Hide
    If chkFlags(0) Then lFlags = lFlags Or Sfx1_Loop
    Me.Tag = lFlags
  End If
End Sub


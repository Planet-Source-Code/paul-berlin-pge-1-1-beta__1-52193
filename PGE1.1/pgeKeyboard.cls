VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Keyboard Class 1.1
'----------------------------------
'Created by Paul Berlin 2002-2003
'
'Note: <Create> The window (for hwnd) must be shown before this can be called,
'so use Me.Show first.
'
'
'OVERVIEW:
'---------
'* Use Create(hWnd) to initialize the keyboard. Note that this sets the key timer to
'  50 ms.
'* Use KeyDown(KeyCode) to check if a key is down.
'* Use SetTimer(lMs) to set an ammount of time to wait after a key has been pressed
'  before it can be pressed again.
'* Use SetTimerEx(KeyCode, lMs) to set the ammount of time to wait for a specified
'  key. You can for example set Enter to 1 seconds wait and Space to 250 ms wait.
'* Use HoldTimer(lMs) to hold the keyboard for a specified ammount of milliseconds.
'* Use bSuspended to disable/enable keyboard.
'
'CHANGELOG:
'----------
'1.1 06-06-2003
'* Changed all subs that took seconds to take milliseconds instead.
'* Cleaned up code & renamed some variables
Option Explicit

Private Type tKeyData
  lKeyTimer As Long
  lKeyInterval As Long
End Type

Public bSuspended As Boolean
Private Key(255) As tKeyData
Private lEndHoldTimer As Long

Private DirectInputDevice As DirectInputDevice8

Private Sub Class_Terminate()
  DirectInputDevice.Unacquire
  Set DirectInputDevice = Nothing
End Sub

Public Sub Create(ByVal hwnd As Long)
  'hWnd = Window handle
  Set DirectInputDevice = DirectInput.CreateDevice("GUID_SysKeyboard")
  DirectInputDevice.SetCommonDataFormat DIFORMAT_KEYBOARD
  DirectInputDevice.SetCooperativeLevel hwnd, DISCL_NONEXCLUSIVE Or DISCL_FOREGROUND
  SetTimer 50
  Aquire
End Sub

Public Function KeyDown(ByVal KeyCode As CONST_DIKEYFLAGS) As Boolean
  'KeyCode = Key to check for
  'Returns: Key is down or not
  On Error Resume Next
  
  If Not bSuspended And timeGetTime >= lEndHoldTimer Then
    Dim KeyState As DIKEYBOARDSTATE
    Call DirectInputDevice.GetDeviceStateKeyboard(KeyState)
    If KeyState.Key(KeyCode) And Key(KeyCode).lKeyTimer - timeGetTime <= 0 Then
      Key(KeyCode).lKeyTimer = timeGetTime + Key(KeyCode).lKeyInterval
      KeyDown = True
    End If
  End If

errh:
  Aquire
End Function

Public Sub SetTimer(ByVal lMs As Long)
  'lMs = Interval in milliseconds
  Dim x As Long
  For x = 0 To 255
    Key(x).lKeyInterval = lMs
  Next
End Sub

Public Sub SetTimerEx(ByVal KeyCode As CONST_DIKEYFLAGS, ByVal lMs As Long)
  'KeyCode = Key to set interval to
  'lMs = Interval in milliseconds
  Key(KeyCode).lKeyInterval = lMs
End Sub

Public Sub HoldTimer(ByVal lMs As Long)
  'lMs = Milliseconds to hold keyboard
  lEndHoldTimer = timeGetTime + lMs
End Sub

Private Sub Aquire()
  On Error GoTo errh
  DirectInputDevice.Acquire
errh:
End Sub

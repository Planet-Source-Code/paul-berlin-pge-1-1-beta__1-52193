VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Mouse Class 1.1
'-------------------------------
'Created by Paul Berlin 2002-2003
'
'NOTE: The window using the mouse should have ScaleMode = 3 Pixel
'
'OVERVIEW:
'---------
'* Use Create(frm) to init the mouse at the specified form.
'* Use UnAquire() to release the mouse.
'* Use Poll(frm) to get data from the mouse.
'* After using Poll(frm), read mouse status from lCursorX, lCursorY, lCursorZ and
'  bButton1-4.
'
'CHANGELOG:
'----------
'1.1 06-06-2003
'* Cleaned up code, made it shorter.
Option Explicit

Public lCursorX As Long
Public lCursorY As Long
Public lCursorZ As Long 'Z is mousewheel
Public lSensitivity As Single
Public bButton1 As Boolean
Public bButton2 As Boolean
Public bButton3 As Boolean
Public bButton4 As Boolean

Private Const BufferSize As Integer = 20

Private DirectInputDevice As DirectInputDevice8

Private Sub Class_Terminate()
  On Error Resume Next
  DirectInputDevice.Unacquire
  Set DirectInputDevice = Nothing
End Sub

Public Sub Create(ByRef frm As Form)
  lCursorX = frm.ScaleWidth / 2
  lCursorY = frm.ScaleHeight / 2
  lSensitivity = 1.5
  
  Set DirectInputDevice = DirectInput.CreateDevice("GUID_SysMouse")
  DirectInputDevice.SetCommonDataFormat DIFORMAT_MOUSE
  DirectInputDevice.SetCooperativeLevel frm.hwnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
  
  ' Set the buffer size
  Dim diProp As DIPROPLONG
  diProp.lHow = DIPH_DEVICE
  diProp.lObj = 0
  diProp.lData = BufferSize
  
  Call DirectInputDevice.SetProperty("DIPROP_BUFFERSIZE", diProp)
  
  Aquire
End Sub

Public Sub Aquire()
  On Error GoTo CANNOTACQUIRE
  DirectInputDevice.Acquire
  On Error GoTo 0

CANNOTACQUIRE:
  Exit Sub
End Sub

Public Sub Poll(ByRef frm As Form)
  Dim diDeviceData(1 To BufferSize) As DIDEVICEOBJECTDATA
  Dim numitems As Long
  Dim i As Integer
  
  On Error GoTo errh
  DirectInputDevice.Poll
  numitems = DirectInputDevice.GetDeviceData(diDeviceData, 0)
  On Error GoTo 0
  For i = 1 To numitems
    Select Case diDeviceData(i).lOfs
      Case DIMOFS_X
        lCursorX = lCursorX + diDeviceData(i).lData * lSensitivity
      Case DIMOFS_Y
        lCursorY = lCursorY + diDeviceData(i).lData * lSensitivity
      Case DIMOFS_Z
        lCursorZ = lCursorZ + diDeviceData(i).lData
      Case DIMOFS_BUTTON0
        bButton1 = (diDeviceData(i).lData And &H80)
      Case DIMOFS_BUTTON1
        bButton2 = (diDeviceData(i).lData And &H80)
      Case DIMOFS_BUTTON2
        bButton3 = (diDeviceData(i).lData And &H80)
      Case DIMOFS_BUTTON3
        bButton4 = (diDeviceData(i).lData And &H80)
    End Select
  Next i

  If lCursorX < 0 Then lCursorX = 0
  If lCursorX >= frm.ScaleWidth Then lCursorX = frm.ScaleWidth - 1
  If lCursorY < 0 Then lCursorY = 0
  If lCursorY >= frm.ScaleHeight Then lCursorY = frm.ScaleHeight - 1

  Exit Sub
errh:
  If (Err.Number = DIERR_INPUTLOST) Or (Err.Number = DIERR_NOTACQUIRED) Then
    Aquire
  End If
End Sub

Public Sub UnAquire()
  DirectInputDevice.Unacquire
End Sub

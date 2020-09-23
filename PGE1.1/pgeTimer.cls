VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Timer Class 1.1
'-------------------------------
'Created by Paul Berlin 2002-2003
'
'Simple way to measure time.
'
'OVERVIEW:
'* Use StartTime() to start/restart the clock.
'* Use GetTime() to get the time since start in milliseconds
'* Use SetStop(lMs) to an ammount of time from now to stop the clock.
'* Use HasStopped() to check if the clock has stopped (with above command).
'* Use Hold(lMs) to hold program execution until specified ammount of seconds
'  has elapsed.
'
'CHANGELOG:
'1.1 06-06-2003
' * Changed everything that took seconds to take milliseconds instead,
'   'cause it's better.
Option Explicit

Private lStartTime As Long
Private lStopTime As Long

Public Sub StartTime()
  lStartTime = timeGetTime
End Sub

Public Function GetTime() As Long
  GetTime = timeGetTime - lStartTime
End Function

Public Sub SetStop(ByVal lMs As Long)
  lStopTime = timeGetTime + lMs
  lStartTime = timeGetTime
End Sub

Public Function HasStopped() As Boolean
  If timeGetTime >= lStopTime Then HasStopped = True
End Function

Public Sub Hold(ByVal lMs As Long)
  Dim lStart As Long
  
  lStart = timeGetTime
  Do
    DoEvents
  Loop Until timeGetTime - lStart >= lMs
End Sub

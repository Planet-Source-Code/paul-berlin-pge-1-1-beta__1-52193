VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine File Reader Class 1.0
'-------------------------------------
'Created by Paul Berlin 2002-2003
'
'This is just a helper class for pgeSound & pgeTileset,
'when they are reading sound- and tilesets.
'Of course, you could use it if you wanted, to write
'your own datafiles, config files, highscore lists, whatever.
'
'This was originally a standalone class called clsDatafile,
'but it is now somewhat modified. The comments for that
'class follows...
'
'* ReadFileMem(b()) was added for Pge.
'  Reads an embedded file into the array.
'
'CHANGELOG:
'
'-------------------------------------------------------------
'clsDatafile v2.0 - 2 December 2003
'This class contains routines for reading and writing to files
'Written by Paul Berlin 2002-2003
'-------------------------------------------------------------
'
'clsDatafile can write and read several types of data to/from a file.
'It CANNOT write single or double values.
'
'This is what it writes:
'[NUMBERS]
'* BYTES (2 bytes, 1 to 255)
'  1. Byte value 50 (used for identification when reading).
'  2. The byte value
'* INTEGERS (3 bytes, -32,768 to 32,767)
'  1. Byte value 51 (used for identification when reading).
'  2. Integer value as 2 bytes.
'* LONGS (5 bytes, -2,147,483,648 to 2,147,483,647)
'  1. Byte value 52 (used for identification when reading)
'  2. Long value as 4 bytes
'[STRINGS]
'* FIXED LENGTH STRING
'  This writes an string of known fixed length (for example a file header)
'  and nothing else. If used in the middle of a file you cannot use SkipField
'  to skip this field! Best used for headers.
'  1. The string itself
'* SHORT STRING (1 to 255 chars) (3-257 bytes)
'  1. Byte value 60 (used for identification when reading)
'  2. Length of string as one byte
'  3. The string itself
'* MEDIUM STRING (256 to 32,767 chars) (259+ bytes)
'  1. Byte value 61 (used for identification when reading)
'  2. Length of string as 2 byte integer
'  3. The string itself
'* LONG STRING (32,768+ chars) (32773+ bytes)
'  1. Byte value 62 (used for identification when reading)
'  2. Length of string as 4 byte long
'  3. The string itself
'[OTHER]
'* DATES AND TIME (9 bytes)
'  1. Byte value 70 (used for identification when reading)
'  2. Date as 8 bytes
'* EMBEDDED FILE (6+ bytes)
'  1. Byte value 71 (used for identification when reading)
'  2. Length of file as 4 byte long
'  3. The entire file itself
'[NUMBERS, STRINGS AND DATES]
'* SKIPPED VALUE (1 byte)
'  The byte value 99 can be entered whenever any number, string length or
'  date is 0. This to save space and speed up a bit.
'  1. Byte value 99 (used for identification when reading)

'FUNCTIONS & SUBROUTINES:
'[PROPERTIES]
'* Filename [String]
'  Assigns an filename to read/write from.
'  This is the first and only thing you have to set before starting to read/write.
'* Position [Long]
'  This is the current byte position in the file you are at.
'  You do not have to touch this property in any simple reading/writing tasks.
'  Perhaps most useful to get the final size of the file when finished writing.
'* XorCode [Boolean]
'  Turn on/off simple Xor coding of strings. Remember to turn on this when
'  reading from Xor coded files.
'[WRITING]
'* WriteStr [String]
'  This simply writes a string to the current position. The string can be any size
'  and the proper writing method (LONG STRING/MEDIUM STRING/SHORT STRING/SKIPPED)
'  will be used depending on the length.
'* WriteNumber [Number]
'  This simply writes the supplied number with the proper writing method
'  (LONG/INTEGER/BYTE/SKIPPED) depending on the length of the number at
'  the current file position.
'* WriteStrFixed [String]
'  Writes an fixed length string at current position. This can only be read with
'  ReadStrFixed, and cannot be skipped with SkipField. This should generally
'  only be used for writing file headers.
'* WriteDate [Date]
'  Writes a date at current file position.
'* WriteFile [File]
'  Writes an entire file into the datafile at current position.
'* There are other writing subroutines, but only the above are needed.
'[READING]
'* ReadStr
'  This reads an string from the current position and returns it.
'  If the value at the current file position is not an string, an error will be
'  raised.
'* ReadNumber
'  This reads an number from the current position and returns it.
'  If the value at the current file position is not an number, an error will be
'  raised.
'* ReadStrFixed [Length]
'  Reads an fixed length string at current position.
'* ReadDate
'  This reads an date from the current position and returns it.
'  If the value at the current file position is not an date, an error will be
'  raised.
'* ReadFile [File]
'  Reads an file embedded into the datafile and saves it to [File].
'  If the value at the current file position is not an file, an error will be
'  raised.
'* There are other writing subroutines, but only the above are needed.
'[OTHER]
'* SkipFiled [NumberSkips]
'  This skips any number of fields, as long as they are recognised.

'CLOSING THE FILE:
'The datafile will be closed whenever you change the filename property or
'you unload the class.

'SIMPLE EXAMPLE:
'--------------
'
'Dim File as new clsDatafile
'
'File.Filename = App.Path & "\file.tmp"
'File.WriteStrFixed "DATA"      'Fixed str header
'File.WriteNumber 100           'Byte number
'File.WriteNumber 100000        'Long number
'File.WriteNumber -100000       'Neg. Long number
'File.WriteNumber 0             'Skipped number
'File.WriteStr "Hello!"         'Byte length string
'File.WriteStr String(500, "X") 'Long length string
'File.WriteStr ""               'Skipped str
'File.WriteDate Now             'Date
'File.WriteFile sFile           'File (must be set to an valid file)
'Set File = Nothing             'Close
'Set File = New clsDatafile
'File.Filename = App.Path & "\file.tmp"
'Debug.Print File.ReadStrFixed 4        'Read fixed str header
'Debug.Print File.ReadNumber            'Byte number
'Debug.Print File.ReadNumber            'Long number
'Debug.Print File.ReadNumber            'Neg. Long number
'Debug.Print File.ReadNumber            'Skipped number
'Debug.Print File.ReadStr               'Byte length string
'Debug.Print File.ReadStr               'Long length string
'Debug.Print File.ReadStr               'Skipped str
'Debug.Print File.ReadDate              'Date
'Debug.Print File.ReadFile sFile        'File (must be set to an valid file)
'

Option Explicit

Const DEF_NUMBER_BYTE As Byte = 50
Const DEF_NUMBER_INTEGER As Byte = 51
Const DEF_NUMBER_LONG As Byte = 52
Const DEF_STRING_SHORT As Byte = 60
Const DEF_STRING_MEDIUM As Byte = 61
Const DEF_STRING_LONG As Byte = 62
Const DEF_DATE As Byte = 70
Const DEF_FILE As Byte = 71
Const DEF_SKIPPED As Byte = 99

Const MaxBuffer As Long = 1048576 'The MAX number of bytes WriteFile can buffer

Private DoXor         As Boolean
Private strFilename   As String 'Filename
Private lngPosition   As Long   'File Position
Private Free          As Integer 'File number
Private IsOpen        As Boolean 'True if file is open, false if not

Public Sub WriteNumber(ByVal wNumber As Long)
  'This sub will examine the length of wNumber and use the
  'appropriate method to write it.
  '------------------------------------------------------
  
  Select Case wNumber
    Case 1 To 255
      WriteByte wNumber
    Case 256 To 32767, -1 To -32768
      WriteInteger wNumber
    Case 0
      WriteNull
    Case Else
      WriteLong wNumber
  End Select
  
End Sub

Public Sub WriteStr(ByVal wString As String)
  'This sub will examine the length of wString and use the
  'appropriate method to write it.
  '------------------------------------------------------

  Select Case Len(wString)
    Case 1 To 255
      WriteShortStr wString
    Case 256 To 32767
      WriteMediumStr wString
    Case 0
      WriteNull
    Case Else
      WriteLongStr wString
  End Select

End Sub

Public Sub WriteShortStr(ByVal wString As String)
  'This sub writes a short string to the current file position
  '-----------------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  If Len(wString) = 0 Then
    WriteNull
    Exit Sub
  End If
  
  'Crop string if it is longer than 255 characters
  If Len(wString) > 255 Then wString = Left(wString, 255)
  
  If DoXor Then wString = XorString(wString)
  
  Put Free, lngPosition, DEF_STRING_SHORT: lngPosition = lngPosition + 1
  Put Free, lngPosition, CByte(Len(wString)): lngPosition = lngPosition + 1
  Put Free, lngPosition, wString: lngPosition = lngPosition + Len(wString)

End Sub

Public Sub WriteMediumStr(ByVal wString As String)
  'This sub writes a medium length string to the current file position
  '-------------------------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  If Len(wString) = 0 Then
    WriteNull
    Exit Sub
  End If
  
  'Crop string if it is longer than 32767 characters
  If Len(wString) > 32767 Then wString = Left(wString, 32767)
  
  If DoXor Then wString = XorString(wString)
  
  Put Free, lngPosition, DEF_STRING_MEDIUM: lngPosition = lngPosition + 1
  Put Free, lngPosition, CInt(Len(wString)): lngPosition = lngPosition + 2
  Put Free, lngPosition, wString: lngPosition = lngPosition + Len(wString)

End Sub

Public Sub WriteLongStr(ByVal wString As String)
  'This sub writes a long string to the current file position
  '----------------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  If Len(wString) = 0 Then
    WriteNull
    Exit Sub
  End If
  
  If DoXor Then wString = XorString(wString)
  
  Put Free, lngPosition, DEF_STRING_LONG: lngPosition = lngPosition + 1
  Put Free, lngPosition, CLng(Len(wString)): lngPosition = lngPosition + 4
  Put Free, lngPosition, wString: lngPosition = lngPosition + Len(wString)

End Sub

Public Sub WriteStrFixed(ByVal wString As String)
  'This sub writes a fixed size string to the current
  'file position. The length can be as long as you wish.
  'No Identification number will be written when writing
  'this string, so you should only use this to write file
  'headers.
  '------------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  Put Free, lngPosition, wString: lngPosition = lngPosition + Len(wString)

End Sub

Public Sub WriteNull()
  'This sub writes the id for an empty value
  '-----------------------------------------

  If Not IsOpen Then OpenFile

  Put Free, lngPosition, DEF_SKIPPED: lngPosition = lngPosition + 1

End Sub

Public Sub WriteByte(ByVal wByte As Byte)
  'This sub writes a byte to the current file position
  '---------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  If wByte = 0 Then
    WriteNull
    Exit Sub
  End If
  
  Put Free, lngPosition, DEF_NUMBER_BYTE: lngPosition = lngPosition + 1
  Put Free, lngPosition, wByte: lngPosition = lngPosition + 1

End Sub

Public Sub WriteInteger(ByVal wInteger As Integer)
  'This sub writes an integer to the current file position
  '-------------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  If wInteger = 0 Then
    WriteNull
    Exit Sub
  End If
  
  Put Free, lngPosition, DEF_NUMBER_INTEGER: lngPosition = lngPosition + 1
  Put Free, lngPosition, wInteger: lngPosition = lngPosition + 2

End Sub

Public Sub WriteLong(ByVal wLong As Long)
  'This sub writes a long to the current file position
  '---------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  If wLong = 0 Then
    WriteNull
    Exit Sub
  End If
  
  Put Free, lngPosition, DEF_NUMBER_LONG: lngPosition = lngPosition + 1
  Put Free, lngPosition, wLong: lngPosition = lngPosition + 4
  
End Sub

Public Sub WriteFile(ByVal sFilename As String)
  'This sub writes an other file into the file
  '-------------------------------------------
  
  Dim sBuffer As String
  Dim lCurLoc As Long
  Dim lFile As Long
  Dim lLen As Long
   
  If Not IsOpen Then OpenFile
  
  lLen = FileLen(sFilename)
  
  Put Free, lngPosition, DEF_FILE: lngPosition = lngPosition + 1
  Put Free, lngPosition, lLen: lngPosition = lngPosition + 4
  
  lFile = FreeFile
  lCurLoc = 1
  Open sFilename For Binary As lFile
  Do
    If lLen - lCurLoc < MaxBuffer Then 'figure out the size of the buffer
      sBuffer = String(lLen - lCurLoc + 1, " ")
    Else
      sBuffer = String(MaxBuffer, " ")
    End If
    Get lFile, lCurLoc, sBuffer 'get from file
    lCurLoc = lCurLoc + Len(sBuffer) 'adjust file position
    Put Free, lngPosition, sBuffer 'put in file
    lngPosition = lngPosition + Len(sBuffer) 'adjust file position
  Loop Until lCurLoc >= lLen

End Sub

Public Sub WriteDate(ByVal wDate As Date)
  'This sub writes a date to the current file position.
  '----------------------------------------------------
 
  If Not IsOpen Then OpenFile
  
  Put Free, lngPosition, DEF_DATE: lngPosition = lngPosition + 1
  Put Free, lngPosition, wDate: lngPosition = lngPosition + 8

End Sub

Public Function ReadStr() As String
  'This function reads a string from the current file
  'position and returns it. It reads short, medium and
  'long strings.
  '---------------------------------------------------
  Dim vID As Byte     'The string id
  Dim vByte As Byte   'The size of the short string
  Dim vInt As Integer 'The size of the medium string
  Dim vLong As Long   'The size of the long string
  
  If Not IsOpen Then OpenFile
  
  Get Free, lngPosition, vID: lngPosition = lngPosition + 1
  Select Case vID
    Case DEF_STRING_SHORT
      Get Free, lngPosition, vByte: lngPosition = lngPosition + 1 'Get size
      ReadStr = Space(vByte) 'Setup the size of the string to read into
      Get Free, lngPosition, ReadStr: lngPosition = lngPosition + vByte
    Case DEF_STRING_MEDIUM
      Get Free, lngPosition, vInt: lngPosition = lngPosition + 2 'Get size
      ReadStr = Space(vInt) 'Setup the size of the string to read into
      Get Free, lngPosition, ReadStr: lngPosition = lngPosition + vInt
    Case DEF_STRING_LONG
      Get Free, lngPosition, vLong: lngPosition = lngPosition + 4 'Get size
      ReadStr = Space(vLong) 'Setup the size of the string to read into
      Get Free, lngPosition, ReadStr: lngPosition = lngPosition + vLong
    Case DEF_SKIPPED
      ReadStr = ""
    Case Else
      Err.Raise 566
  End Select
  
  If DoXor Then ReadStr = XorString(ReadStr)
  
End Function

Public Function ReadStrFixed(ByVal vSize As Long) As String
  'This function reads an fixed size string from the
  'current file position and returns it. vSize is the
  'size of the string to read.
  '---------------------------------------------------
  
  If Not IsOpen Then OpenFile
  
  ReadStrFixed = Space(vSize) 'Setup the size of the string to read into
  Get Free, lngPosition, ReadStrFixed: lngPosition = lngPosition + vSize
  
End Function

Public Sub ReadFile(ByVal sFilename As String)
  'This sub reads and extracts a file and saves it to sFilename
  '------------------------------------------------------------
  Dim sBuffer As String
  Dim lCurLoc As Long
  Dim lFile As Long
  Dim lLen As Long
  Dim vID As Byte
  
  If Not IsOpen Then OpenFile
  
  Get Free, lngPosition, vID: lngPosition = lngPosition + 1 'The ID
  If vID = DEF_FILE Then
      
    Get Free, lngPosition, lLen: lngPosition = lngPosition + 4
    lLen = lngPosition + lLen
    lFile = FreeFile
    lCurLoc = 1
    Open sFilename For Binary As lFile
    Do
      If lLen - lngPosition < MaxBuffer Then
        sBuffer = String(lLen - lngPosition, " ")
      Else
        sBuffer = String(MaxBuffer, " ")
      End If
      Get Free, lngPosition, sBuffer
      lngPosition = lngPosition + Len(sBuffer)
      Put lFile, lCurLoc, sBuffer
      lCurLoc = lCurLoc + Len(sBuffer)
    Loop Until lngPosition >= lLen
    
    Close lFile
    
  Else
    Err.Raise 566
    Exit Sub
  End If

End Sub

Public Function ReadNumber() As Long
  'This function reads a number from the current file
  'position and returns it. It reads bytes, integers
  'and longs, but returns all as longs.
  '---------------------------------------------------
  Dim vID As Byte     'The id of number to read
  Dim vByte As Byte   'Byte number
  Dim vInt As Integer 'Integer number
  Dim vLong As Long   'Long number
  
  If Not IsOpen Then OpenFile
  
  Get Free, lngPosition, vID: lngPosition = lngPosition + 1
  Select Case vID
    Case DEF_NUMBER_BYTE
      Get Free, lngPosition, vByte: lngPosition = lngPosition + 1
      ReadNumber = vByte
    Case DEF_NUMBER_INTEGER
      Get Free, lngPosition, vInt: lngPosition = lngPosition + 2
      ReadNumber = vInt
    Case DEF_NUMBER_LONG
      Get Free, lngPosition, vLong: lngPosition = lngPosition + 4
      ReadNumber = vLong
    Case DEF_SKIPPED
      ReadNumber = 0
    Case Else
      Err.Raise 566
  End Select
  
End Function

Public Function ReadDate() As Date
  'This function reads a date at the current file position and returns it.
  '-----------------------------------------------------------------------
  Dim vID As Byte
  
  If Not IsOpen Then OpenFile
  
  Get Free, lngPosition, vID: lngPosition = lngPosition + 1
  If vID = DEF_DATE Then
    Get Free, lngPosition, ReadDate: lngPosition = lngPosition + 8
  Else
    Err.Raise 566
  End If

End Function

Public Sub SkipField(Optional ByVal lSkips As Integer = 1)
  'This sub skips lSkips number of values, and moves the current
  ' file position after them without reading them.
  '-----------------------------------------------------------------
  Dim x As Integer
  Dim vID As Byte
  Dim vByte As Byte
  Dim vInt As Integer
  Dim vLong As Long
  
  For x = 1 To lSkips
    Get Free, lngPosition, vID: lngPosition = lngPosition + 1 'Check which value this is
    Select Case vID
      Case DEF_NUMBER_BYTE
        lngPosition = lngPosition + 1
      Case DEF_NUMBER_INTEGER
        lngPosition = lngPosition + 2
      Case DEF_NUMBER_LONG
        lngPosition = lngPosition + 4
      Case DEF_STRING_SHORT
        Get Free, lngPosition, vByte
        lngPosition = lngPosition + vByte + 1
      Case DEF_STRING_MEDIUM
        Get Free, lngPosition, vInt
        lngPosition = lngPosition + vInt + 2
      Case DEF_STRING_LONG, DEF_FILE
        Get Free, lngPosition, vLong
        lngPosition = lngPosition + vLong + 4
      Case DEF_DATE
        lngPosition = lngPosition + 8
      Case DEF_SKIPPED
        'Do nothing
      Case Else 'Could not identify an header, position might be off...
        Err.Raise 566
    End Select
  Next x
  
End Sub

Private Function XorString(ByVal Text As String) As String
  Dim A As Integer
  For A = 1 To Len(Text) 'code all
   XorString = XorString & Chr(Asc(Mid(Text, A, 1)) Xor 1) 'Code text
  Next A
End Function

Public Property Get XorCode() As Boolean
  XorCode = DoXor
End Property

Public Property Let XorCode(ByVal bXor As Boolean)
  DoXor = bXor
End Property

Public Property Get Position() As Long
  Position = lngPosition
End Property

Public Property Let Position(ByVal lPos As Long)
  lngPosition = lPos
  If lngPosition < 1 Then lngPosition = 1 'The position cannot be lower than 1
End Property

Public Property Let FileName(ByVal sFilename As String)
  If IsOpen Then
    Close Free
    lngPosition = 1
    IsOpen = False
  End If
  strFilename = sFilename
End Property

Public Property Get FileName() As String
  FileName = strFilename
End Property

Private Sub Class_Initialize()
  'Setup file position
  lngPosition = 1
End Sub

Private Sub OpenFile()
  'This initializes the file
  Free = FreeFile
  Open strFilename For Binary As Free
  IsOpen = True
End Sub

Private Sub Class_Terminate()
  'This closes the file
  Close Free
End Sub

Public Sub ReadFileMem(ByRef b() As Byte)
  'This reads a file from datafile to array
  Dim lLen As Long
  Dim vID As Byte

  Get Free, lngPosition, vID: lngPosition = lngPosition + 1 'The ID
  If vID = DEF_FILE Then
    
    Get Free, lngPosition, lLen: lngPosition = lngPosition + 4
    ReDim b(lLen)
    Get Free, lngPosition, b
    lngPosition = lngPosition + lLen
    
  Else
    Err.Raise 566
    Exit Sub
  End If
End Sub


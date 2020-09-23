Attribute VB_Name = "modMisc"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Enum enumFileNameParts
    efpFileName = 2 ^ 0
    efpFileExt = 2 ^ 1
    efpFilePath = 2 ^ 2
    efpFileNameAndExt = efpFileName + efpFileExt
    efpFileNameAndPath = efpFilePath + efpFileName
End Enum

Public Enum enumKeyPressAllowTypes
    NumbersOnly = 2 ^ 0
    Uppercase = 2 ^ 1
    NoSpaces = 2 ^ 2
    NoSingleQuotes = 2 ^ 3
    NoDoubleQuotes = 2 ^ 4
    AllowDecimal = 2 ^ 5
    AllowNegative = 2 ^ 6
    DatesOnly = 2 ^ 7
    TimesOnly = 2 ^ 8
    LettersOnly = 2 ^ 9
    AllowSpaces = 2 ^ 10
    AllowStars = 2 ^ 11
    AllowPounds = 2 ^ 12
End Enum

Public Function FileExist(ByVal FileName As String) As Boolean
  FileExist = Not (Dir(FileName) = "")
End Function

Public Function ctlKeyPress(ByVal KeyAscii As KeyCodeConstants, ByVal TypeToAllow As enumKeyPressAllowTypes) As Integer
    Dim ltrKeyAscii As Integer
    ltrKeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    ' By default pass the keystroke through and then optionally kill it
    ctlKeyPress = KeyAscii
    
    ' Default Keystrokes to allow (enter, backspace, delete, escape)
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then Exit Function
    
    ' NumbersOnly
    If (TypeToAllow And NumbersOnly) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case (KeyAscii = vbKeySubtract Or KeyAscii = Asc("-")) And (TypeToAllow And AllowNegative)
            Case KeyAscii = Asc("#") And (TypeToAllow And AllowPounds)
            Case KeyAscii = Asc("*") And (TypeToAllow And AllowStars)
            Case KeyAscii = vbKeyDecimal And (TypeToAllow And AllowDecimal)
            Case KeyAscii = vbKeySpace And (TypeToAllow And AllowSpaces)
            Case Else
                KeyAscii = 0
        End Select
    End If
    
    ' DatesOnly
    If (TypeToAllow And DatesOnly) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case KeyAscii = vbKeyDivide Or KeyAscii = Asc("/")
            Case Else
                KeyAscii = 0
        End Select
    End If
    
    ' TimesOnly
    If (TypeToAllow And TimesOnly) Then
        Select Case True
            Case KeyAscii >= vbKey0 And KeyAscii <= vbKey9
            Case KeyAscii = Asc(":") Or KeyAscii = Asc(";")
                ctlKeyPress = Asc(":")
            Case ltrKeyAscii = vbKeyA Or ltrKeyAscii = vbKeyP Or ltrKeyAscii = vbKeyM
            Case Else
                KeyAscii = 0
        End Select
    End If
            
    ' LettersOnly
    If (TypeToAllow And LettersOnly) Then
        Select Case True
            Case ltrKeyAscii >= vbKeyA And ltrKeyAscii <= vbKeyZ
            Case Else
                KeyAscii = 0
        End Select
    End If
            
    ' UpperCase
    If (TypeToAllow And Uppercase) Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
    
    ' No Spaces
    If (TypeToAllow And NoSpaces) And KeyAscii = vbKeySpace Then
        KeyAscii = 0
    End If
    
    ' No Double Quotes
    If (TypeToAllow And NoDoubleQuotes) And KeyAscii = Asc("""") Then
        KeyAscii = Asc("'")
    End If
    
    ' No Single Quotes
    If (TypeToAllow And NoSingleQuotes) And KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
    
    ctlKeyPress = KeyAscii
    
End Function

Public Function sFilename(ByVal sFile As String, ByVal ePortions As enumFileNameParts) As String
    
    Dim lFirstPeriod As Long, lFirstBackSlash As Long
    Dim sName As String, sExt As String
    Dim sRet As String
    
    lFirstPeriod = InStrRev(sFile, ".")
    lFirstBackSlash = InStrRev(sFile, "\")
    
    If ePortions And efpFilePath Then
      If lFirstBackSlash > 0 Then
        sRet = Left(sFile, lFirstBackSlash)
      End If
    End If
    If ePortions And efpFileName Then
      If lFirstPeriod > 0 And lFirstPeriod > lFirstBackSlash Then
        sName = Mid(sFile, lFirstBackSlash + 1, lFirstPeriod - lFirstBackSlash - 1)
      Else
        sName = Mid(sFile, lFirstBackSlash + 1)
      End If
      sRet = sRet & sName
    End If
    If ePortions And efpFileExt Then
      If lFirstPeriod > 0 And lFirstPeriod > lFirstBackSlash Then
        sExt = Mid(sFile, lFirstPeriod + 1)
      End If
      If sRet <> "" Then
        sRet = sRet & "." & sExt
      Else
        sRet = sRet & sExt
      End If
    End If
    
    sFilename = sRet
    
End Function

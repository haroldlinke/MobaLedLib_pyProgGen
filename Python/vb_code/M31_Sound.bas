Attribute VB_Name = "M31_Sound"
Option Explicit
' https://wellsr.com/vba/2019/excel/vba-playsound-to-play-system-sounds-and-wav-files/

#If VBA7 Then
    Public Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As LongPtr, ByVal dwFlags As Long) As Long
#Else
    Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
#End If

Const SND_SYNC = &H0                ' wait for sound to play
Const SND_ASYNC = &H1               ' no wait
Const SND_NODEFAULT = &H2           ' no default sound on error
Const SND_NOSTOP = &H10             ' skip sound if another is playing
Const SND_ALIAS = &H10000           ' play system sound
Const SND_FILENAME = &H20000        ' play WAV file

Public Function BeepThis2(Optional ByVal ThisSound As String = "Beep" _
                        , Optional ByVal ThisValue As Variant _
                        , Optional ByVal ThisCount As Integer = 1 _
                        , Optional ByVal Wait As Boolean = False) As Variant
    Dim sPath As String, flags As Long
    Const sMedia As String = "\Media\"
    If IsMissing(ThisValue) Then ThisValue = ThisSound
    BeepThis2 = ThisValue           ' return value
    If ThisCount > 1 Then Wait = True
    flags = SND_ALIAS
    sPath = StrConv(ThisSound, vbProperCase)
    Select Case sPath
    Case "Beep"
        Beep                        ' ignore ThisCount and Wait
        Exit Function
    Case "Asterisk", "Exclamation", "Hand", "Notification", "Question"
        sPath = "System" + sPath
    Case "Connect", "Disconnect", "Fail"
        sPath = "Device" + sPath
    Case "Mail", "Reminder"
        sPath = "Notification." + sPath
    Case "Text"
        sPath = "Notification.SMS"
    Case "Message"
        sPath = "Notification.IM"
    Case "Fax"
        sPath = "FaxBeep"
    Case "Select"
        sPath = "CCSelect"
    Case "Error"
        sPath = "AppGPFault"
    Case "Close", "Maximize", "Minimize", "Open"
        ' ok
    Case "Default"
        sPath = "." & sPath
    Case "Chimes", "Chord", "Ding", "Notify", "Recycle", "Ringout", "Tada"
        sPath = Environ("SystemRoot") & sMedia & sPath & ".wav"
        flags = SND_FILENAME
    Case Else
        If LCase(right(ThisSound, 4)) <> ".wav" Then ThisSound = ThisSound & ".wav"
        sPath = ThisSound
        If Dir(sPath) = "" Then     ' file is not in working directory
            sPath = ActiveWorkbook.Path & "\" & ThisSound
            If Dir(sPath) = "" Then sPath = Environ("SystemRoot") & sMedia & ThisSound
        End If
        flags = SND_FILENAME
    End Select
    flags = flags + IIf(Wait, SND_SYNC, SND_ASYNC)
    Do While ThisCount > 0          ' skip if ThisCount < 1
        PlaySound sPath, 0, flags   ' if error, .Default sound will play
        ThisCount = ThisCount - 1
    Loop
End Function

'-------------------------------------------------------------------------
Public Function BeepThis1(Optional ByVal ThisSound As String = "Beep" _
                        , Optional ByVal ThisValue As Variant) As Variant
'-------------------------------------------------------------------------
    If IsMissing(ThisValue) Then ThisValue = ThisSound
    BeepThis1 = ThisValue
    Beep
End Function

Private Sub Test_BeepThis1()
  'BeepThis2 "Default"
  'BeepThis2 "Asterisk"
  'BeepThis2 "Fax"
  BeepThis2 "Windows Information Bar.wav", , , True
  
End Sub



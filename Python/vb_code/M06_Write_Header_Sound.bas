Attribute VB_Name = "M06_Write_Header_Sound"
Option Explicit

' Header file generation for the Sound functions
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Dim SoundLines As Scripting.Dictionary
'Dim UseFullPacketMode As Boolean                                           ' 19.10.21: Always use the full packed mode because this mode could be used with both types of JQ6500 modules

'--------------------------------------------------------------
Public Function Init_HeaderFile_Generation_Sound() As Boolean
'--------------------------------------------------------------
  Set SoundLines = New Scripting.Dictionary
  Init_HeaderFile_Generation_Sound = True
  'UseFullPacketMode = false                                                ' 19.10.21: Always use the full packed mode because this mode could be used with both types of JQ6500 modules
End Function

'-------------------------------------------------------------------------------
Public Function Add_SoundPin_Entry(ByRef Cmd As String, Channel As Long) As Boolean

  Dim Parts() As String, Typ As String
  Parts = Split(Replace(Replace(Replace(Trim(Cmd), SF_SERIAL_SOUND_PIN, ""), ")", ""), " ", ""), ",")
  If UBound(Parts) - LBound(Parts) <> 1 Then
    ' todo
    Exit Function
  End If
  Dim Pin As String
  If Set_PinNrLst_if_Matching(SF_SERIAL_SOUND_PIN + Parts(0) + ")", SF_SERIAL_SOUND_PIN, Pin, "O", 1) = False Then Exit Function
  If No_Duplicates_in_two_Lists("Sound", Serial_PinLst, Pin, SF_SERIAL_SOUND_PIN) = False Then Exit Function
  Serial_PinLst = Serial_PinLst + Pin + " "
  If Not Check_Sound_Duplicates Then Exit Function
  
  Dim playerClass As String
  playerClass = GetPlayerClass(Parts(1))
  
  If playerClass = "" Then
        MsgBox Replace(Get_Language_Str("Fehler: Der Soundmodul Typ '#1#' wird nicht unterstützt."), _
                    "#1#", Parts(1)), vbCritical, "Fehler: Soundmodul"
        Exit Function
  End If
  
  
  If SoundLines.Exists(Channel) Then
        MsgBox Replace(Get_Language_Str("Fehler: Der Sound Kanal '#1#' ist schon definiert."), _
                    "#1#", Channel), vbCritical, "Fehler: Soundmodul"
        Exit Function
  End If
                    
  'If Parts(1) = "JQ6500_AA" Then UseFullPacketMode = True                 ' 19.10.21: Always use the full packed mode because this mode could be used with both types of JQ6500 modules
  
  SoundLines.Add Channel, Array(Pin, playerClass)
  
  Cmd = "// " & Cmd
  Add_SoundPin_Entry = True
End Function

Private Function GetPlayerClass(moduleType As String) As String
  Select Case moduleType
    Case "JQ6500":
        GetPlayerClass = "JQ6500SoundPlayer"
    Case "JQ6500_AA":
        GetPlayerClass = "JQ6500SoundPlayer"
        
    Case "MP3-TF-16P":                                                       ' 02.11.21: Juergen add MP3TF16P
        GetPlayerClass = "MP3TF16PSoundPlayer"
    Case "MP3-TF-16P-NO-CRC":                                                ' 02.11.21: Juergen add MP3TF16P
        GetPlayerClass = "MP3TF16PNoCRCSoundPlayer"
    Case Else: GetPlayerClass = ""
  End Select
End Function

'------------------------------------------------------------------
Public Function Write_Header_File_Sound_Before_Config(fp As Integer) As Boolean

  If SoundLines.Count > 0 Then
     If Check_Sound_Duplicates() = False Then Exit Function

     Print #fp, "// ----- Serial Onboard Sound Makros -----"
     Print #fp, "  #include ""SoundChannelMacros.h"""
     Print #fp, ""
     Dim Index As Byte, Key
     Index = 0
     For Each Key In SoundLines.Keys
        Print #fp, "  #define SOUND_CHANNEL_" & Key & " " & Index
        Index = Index + 1
     Next
     Print #fp, ""
  End If
  Write_Header_File_Sound_Before_Config = True
End Function

'------------------------------------------------------------------
Public Function Write_Header_File_Sound_After_Config(fp As Integer) As Boolean
'------------------------------------------------------------------
  If SoundLines.Count > 0 Then
     If Check_Sound_Duplicates() = False Then Exit Function
      
     Print #fp, "// ----- Serial Onboard Sound -----"
     Print #fp, "#ifndef _USE_EXT_PROC"
     Print #fp, "  #error _USE_EXT_PROC must be enabled in MobaLebLib, see file 'Lib_Config.h'"
     Print #fp, "#else"
     Print #fp, "  // includes for Onboard sound processing"
     'If UseFullPacketMode Then                                             ' 19.10.21: Always use the full packed mode because this mode could be used with both types of JQ6500 modules
         Print #fp, "#define _SOUNDPROCCESSOR_SEND_FULL_PACKET"
     'End If
     Print #fp, "  #include ""SoundProcessor.h"""
     Print #fp, ""
     Print #fp, "  #ifndef _ENABLE_EXT_PROC"
     Print #fp, "  #define _ENABLE_EXT_PROC"
     Print #fp, "  #endif"
     Print #fp, "  #ifndef _SOUND_SERBUFFER_SIZE"
     Print #fp, "  #define _SOUND_SERBUFFER_SIZE "; 15 + SoundLines.Count * 5
     Print #fp, "  #endif"
     Print #fp, ""
     
'     Dim modulesArray As String, playersArray As String
'     Dim ChannelToModuleIndex As Scripting.Dictionary
'     Set ChannelToModuleIndex = New Scripting.Dictionary
'     Dim Index As Byte, key
'     Index = 0
      
'     For Each key In SoundLines.Keys
'        If modulesArray <> "" Then modulesArray = modulesArray + ", "
'        modulesArray = modulesArray + "SoundProcessor::CreateSoftwareSerial(" + SoundLines(key)(0) + ", 9600)"
'        'Debug.Print "Channel " & key & " Pin " & SoundLines(key)(0) & " for channel "; SoundLines(key)(1)
'        'ChannelToModuleId.Add key, index
        
'        If playersArray <> "" Then playersArray = playersArray + ", "
'        playersArray = playersArray + "new " & SoundLines(key)(1) & "(" & Str(Index) & ", &serialDispatcher)"
'        Index = Index + 1
'     Next
     
'     Print #fp, "  SOFTWARE_SERIAL_TYPE* mySerial[] { " + modulesArray + "};"
'     Print #fp, "  uint8_t serBuffer[_SOUND_SERBUFFER_SIZE];"
'     Print #fp, "  SoundSerialDispatcher serialDispatcher(serBuffer, _SOUND_SERBUFFER_SIZE, mySerial);"
'     Print #fp, "  SoundPlayer* soundPlayers[] {" + playersArray + "};"
'     Print #fp, "  SoundProcessor soundProcessor;"


' 02.11.2021: Juergen add support of multiple sound module types
' START_CHANGE
     Dim module As String, playersArray As String
     Dim ChannelToModuleIndex As Scripting.Dictionary
     Set ChannelToModuleIndex = New Scripting.Dictionary
     Dim Index As Byte, Key
     Index = 0
      
     For Each Key In SoundLines.Keys
        module = "SoundProcessor::CreateSoftwareSerial(" + SoundLines(Key)(0) + ", 9600)"
        If playersArray <> "" Then playersArray = playersArray + ", "
        playersArray = playersArray + "new " & SoundLines(Key)(1) & "(" & Str(Index) & ", " & module & ")"
        Index = Index + 1
     Next
     
     Print #fp, "  uint8_t serBuffer[_SOUND_SERBUFFER_SIZE];"
     Print #fp, "  SoundPlayer* soundPlayers[] {" + playersArray + "};"
     Print #fp, "  SoundProcessor soundProcessor(serBuffer, _SOUND_SERBUFFER_SIZE, soundPlayers);"
' END_CHANGE
     Print #fp, "#endif"
     Print #fp, ""
  End If
  Write_Header_File_Sound_After_Config = True
End Function

Private Function Check_Sound_Duplicates() As Boolean
    If No_Duplicates_in_two_Lists("LED", Serial_PinLst, LED_PINNr_List, SF_SERIAL_SOUND_PIN) = False Then Exit Function
    If SwitchA_InpCnt Then
        If No_Duplicates_in_two_Lists("Switch A", Serial_PinLst, SwitchA_InpLst, SF_SERIAL_SOUND_PIN) = False Then Exit Function
    End If
    If SwitchB_InpCnt Then
        If No_Duplicates_in_two_Lists("Switch B", Serial_PinLst, SwitchB_InpLst, SF_SERIAL_SOUND_PIN) = False Then Exit Function
    End If
    If SwitchC_InpCnt Then
        If No_Duplicates_in_two_Lists("Switch C", Serial_PinLst, SwitchC_InpLst, SF_SERIAL_SOUND_PIN) = False Then Exit Function
    End If
    If SwitchD_InpCnt Then
        If No_Duplicates_in_two_Lists("Switch D", Serial_PinLst, SwitchD_InpLst, SF_SERIAL_SOUND_PIN) = False Then Exit Function
    End If
    If SwitchB_InpCnt > 0 Or SwitchC_InpCnt > 0 Then
        If No_Duplicates_in_two_Lists("Switch B/C Clock", Serial_PinLst, CLK_Pin_Number, SF_SERIAL_SOUND_PIN) = False Then Exit Function
        If No_Duplicates_in_two_Lists("Switch B/C Reset", Serial_PinLst, RST_Pin_Number, SF_SERIAL_SOUND_PIN) = False Then Exit Function
    End If
    If Read_LDR Then
        If No_Duplicates_in_two_Lists("LDR_Pin_Number", Serial_PinLst, LDR_Pin_Number, SF_SERIAL_SOUND_PIN) = False Then Exit Function
    End If
    If No_Duplicates_in_two_Lists("LED", Serial_PinLst, LED_PINNr_List, SF_SERIAL_SOUND_PIN) = False Then Exit Function
    Check_Sound_Duplicates = True
End Function

Public Function CheckSoundChannelDefined(Channel As Long) As Boolean
  If Not SoundLines.Exists(Channel) Then
        MsgBox Replace(Get_Language_Str("Fehler: Der Sound Kanal '#1#' ist nicht definiert." & vbCr & _
                "Zur Definition muss das Makro " & SF_SERIAL_SOUND_PIN & " vor dieser Zeile verwendet werden"), _
                    "#1#", Channel), vbCritical, "Fehler: Soundmodul"
        Exit Function
  End If
  CheckSoundChannelDefined = True
End Function


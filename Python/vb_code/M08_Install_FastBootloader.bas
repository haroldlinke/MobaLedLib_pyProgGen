Attribute VB_Name = "M08_Install_FastBootloader"
Option Explicit

' Bootloader Programmieren:
' - Jumper setzen unten die linken Pins verbinden
' - ArduinoISP aud DCC spielen
' - Bootloader schreiben
' - Fragen ob noch weitere Arduinos Programmiert werden sollen
' - Wenn Nein, dann wieder das DCC Prog. Installieren (Wichtig)
' - Jumper �ffnen





'----------------------------------------------------------------
Private Function Install_ArduinoISP_to_Right_Arduino() As Boolean
'----------------------------------------------------------------
' Compile and upload the ArduinoISP program to the right Arduino
  Make_sure_that_Col_Variables_match
  Dim InoName As String, SrcDir As String, DstDir As String
  InoName = "ArduinoISP.ino"
  DstDir = ThisWorkbook.Path & "\" & FileName(InoName) & "\"
  SrcDir = FilePath(Find_ArduinoExe()) & "examples\11.ArduinoISP\ArduinoISP\"
  
  CreateFolder DstDir
  
  If Not FileCopy_with_Check(DstDir, InoName, SrcDir & InoName) Then Exit Function
  
  If Compile_and_Upload_Prog_to_Arduino(InoName, COMPrtR_COL, BUILDOpRCOL, DstDir) Then
       Cells(SH_VARS_ROW, R_UPLOD_COL) = "R ISP"
       Install_ArduinoISP_to_Right_Arduino = True
  End If
End Function


'------------------------------------------------------------------------------------
Public Function Create_WriteFastBootloader_cmd(SrcDir As String) As String
'------------------------------------------------------------------------------------
' "C:\Users\Hardi\AppData\Local\Arduino15\packages\arduino\tools\avrdude\6.3.0-arduino17/bin/avrdude" ^
'    "-CC:\Users\Hardi\AppData\Local\Arduino15\packages\arduino\tools\avrdude\6.3.0-arduino17/etc/avrdude.conf"
'    -v -patmega328p -cstk500v1 -PCOM3 -b19200 ^
'    "-Uflash:w:C:\Program Files (x86)\Arduino\hardware\arduino\avr/bootloaders/optiboot/optiboot_atmega328.hex:i" ^
'    -Ulock:w:0x0F:m
  Dim Name As String
  Name = SrcDir & "WriteFastBootloader.cmd"
  'If Dir(Name) <> "" Then                                                  ' 04.11.20: Always write the file
  '   Create_WriteFastBootloader_cmd = Name
  '   Exit Function
  'End If

  Dim fp As Integer
  fp = FreeFile
      
  On Error GoTo WriteError
  Open Name For Output As #fp
  Print #fp, "@ECHO OFF"
  Print #fp, "REM Write the fast Bootloader to the left Arduino"
  Print #fp, "REM ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
  Print #fp, "REM Parameter:               Example"
  Print #fp, "REM  1: Arduino EXE Path:    ""C:\Program Files (x86)\Arduino\"""
  Print #fp, "REM  2: Port:                -PCOM3"
  Print #fp, "REM"
  Print #fp, "REM The program uses the captured and adapted command line from the Arduino IDE"
  Print #fp, "REM"
  Print #fp, "Rem Using the private OptiBoot version 108.1 to indicate that the HFUSE is set to DE"
  Print #fp, "Rem This bootloader is equal to version 8.1"
  Print #fp, "REM"
  Print #fp, "REM This file was automatically generated by the program " & ThisWorkbook.Name & " " & Prog_Version & "      by Hardi"
  Print #fp, "REM File creation: " & Date & " " & Time
  Print #fp, ""
  Print #fp, "SET ArduinoExePath=%~1"
  Print #fp, "SET Port=%2"
  Print #fp, ""
  Print #fp, """%ArduinoExePath%\hardware\tools\avr/bin/avrdude"" ^"
  Print #fp, "   ""-C%ArduinoExePath%hardware\tools\avr\etc\avrdude.conf"" ^"
  Print #fp, "   -v -patmega328p -cstk500v1 %Port%  -b19200 ^"
 'Print #fp, "   ""-Uflash:w:%ArduinoExePath%hardware\arduino\avr/bootloaders/optiboot/optiboot_atmega328.hex:i"" ^" ' Standard Optiboot bootloader
 'Print #fp, "   ""-Uflash:w:" & GetShortPath(SrcDir & "optiboot_atmega328_Ver108.1.hex") & ":i"" ^"                       ' 02.11.20:
  Print #fp, "   ""-Uflash:w:" & GetShortPath(Get_SrcDirInLib() & "ArduinoISP\optiboot_atmega328_Ver108.1.hex") & ":i"" ^" ' 04.11.20:
  
  Print #fp, "   -Ulock:w:0x0F:m ^"
  Print #fp, "   -Uhfuse:w:0xDE:m"                                          ' 29.10.20: Reserve only 512 Byte for the bootloader

  Print #fp, ""
  Print #fp, "if %errorlevel%==1 ("
  Print #fp, "   COLOR 4F" ' Yellow on Red
  Print #fp, "   ECHO *********************************"
  Print #fp, "   ECHO Error writing the boot loader ;-("
  Print #fp, "   ECHO *********************************"
  Print #fp, "   PAUSE"
  Print #fp, ")"
  Print #fp, ""
  Close #fp
  On Error GoTo 0
  Create_WriteFastBootloader_cmd = Name
  Exit Function

WriteError:
  MsgBox Get_Language_Str("Fehler beim schreiben der Datei '") & Name & "'", vbCritical, Get_Language_Str("Fehler beim erzeugen der Compile und Flash Datei")
End Function


'UT---------------------------------------------------------
Private Sub Test_Create_WriteFastBootloader_cmd()
'UT---------------------------------------------------------
  Debug.Print Create_WriteFastBootloader_cmd(ThisWorkbook.Path & "\ArduinoISP\")
End Sub

'---------------------------------------------
Private Function Write_Bootloader() As Boolean
'---------------------------------------------
  #If VBA7 Then
    Dim hWnd As LongPtr: hWnd = Application.hWnd
  #Else
    Dim hWnd As Long:    hWnd = Application.hWnd
  #End If
  
  Dim CmdName As String, CommandStr As String
  CmdName = Create_WriteFastBootloader_cmd(ThisWorkbook.Path & "\ArduinoISP\")
  If CmdName = "" Then Exit Function
  CommandStr = """" & CmdName & """ """ & FilePath(Find_ArduinoExe()) & """" & " -PCOM" & Cells(SH_VARS_ROW, COMPrtR_COL)     ' 07.10.20: Added: Cells(SH_VARS_ROW, COMPrtR_COL)

  Dim Res As ShellAndWaitResult
  Res = ShellAndWait(CommandStr, 0, vbNormalFocus, PromptUser) ' No timeout to be able to study the results in case of an error
  Select Case Res
    Case Success, Timeout: ' No additional error message. They have been shown in the DOS box
    Case Else:             Unload StatusMsg_UserForm
                           MsgBox Get_Language_Str("Fehler ") & Res & Get_Language_Str(" beim Starten des Arduino Programms '") & CommandStr & "'", vbCritical, _
                                  Get_Language_Str("Fehler beim Starten des Arduino programms")
  End Select
  Bring_to_front hWnd
  Write_Bootloader = True
End Function




Private Sub Old_Prog()
  Compile_and_Upload_Prog_to_Right_Arduino
End Sub

'----------------------------------
Public Sub Install_FastBootloader()
'----------------------------------
  Make_sure_that_Col_Variables_match
  If Page_ID <> "DCC" And Page_ID <> "Selectrix" Then
     MsgBox Get_Language_Str("Die schnelle Bootloader kann nur von einer DCC oder Selectrix Seite aus installiert werden."), vbInformation, _
            Get_Language_Str("Falsche Seite zum aktualisieren des Bootloaders ausgew�hlt")
     Exit Sub
  End If
    
  Dim Res As Boolean
  Res = BootJumper_Form.ShowDialog
  Sleep 1000
  If Res Then
     If Not Install_ArduinoISP_to_Right_Arduino() Then Exit Sub
     Do
     Write_Bootloader
     Loop While MsgBox(Get_Language_Str("Installation des Bootloaders abgeschlossen" & vbCr & _
                                        vbCr & _
                                        "Soll der Bootloader auf einen weiteren Arduino geladen werden?" & vbCr & _
                                        vbCr & _
                                        "Wenn ja, dann muss dieser jetzt in den linken Steckplatz gesteckt werden." & vbCr & _
                                        "Mit 'Nein' wird wieder das DCC/Selectrix Programm auf den rechten Nano installiert." & vbCr & _
                                        vbCr & _
                                        "Achtung der rechte Arduino darf nicht entfernt werden!"), vbYesNo + vbDefaultButton2, _
                       Get_Language_Str("Noch einen Arduino aktualisieren?")) = vbYes
     
     Compile_and_Upload_Prog_to_Right_Arduino
     MsgBox Get_Language_Str("Achtung: Die Jumper m�ssen unbedingt wieder entfernt werden sonst geht nichts mehr ;-(" & vbCr & _
                             "Damit sie nicht verloren gehen k�nnen sie so eingesteckt werden, dass sie nur auf einem Pin stecken." & vbCr & _
                             vbCr & _
                             "Das USB Kabel sollte wieder auf den linken Arduino gesteckt werden."), vbInformation, _
            Get_Language_Str("Bootloader Programmierung abgeschlossen")
                             
  End If
End Sub

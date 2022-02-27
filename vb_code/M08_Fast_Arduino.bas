Attribute VB_Name = "M08_Fast_Arduino"
Option Explicit

' Fast way to compile and Flash the programm from Jürgen

' ToDo:
' ~~~~~
' - Test with a different Sketchbook path                                   O.K.
' - Test with old Arduino Bootloader                                        O.K.
'

' Command line captured from the Arduino IDE
'   C:\Program Files (x86)\Arduino\arduino-builder -dump-prefs -logger=machine
'    -hardware C:\Program Files (x86)\Arduino\hardware
'    -hardware C:\Users\Hardi\AppData\Local\Arduino15\packages
'    -hardware C:\Users\Hardi\Documents\Arduino\hardware
'    -tools C:\Program Files (x86)\Arduino\tools-builder
'    -tools C:\Program Files (x86)\Arduino\hardware\tools\avr
'    -tools C:\Users\Hardi\AppData\Local\Arduino15\packages
'    -built-in-libraries C:\Program Files (x86)\Arduino\libraries
'    -libraries C:\Users\Hardi\Documents\Arduino\libraries
'    -fqbn=arduino:avr:nano:cpu=atmega328
'    -vid-pid=1A86_7523
'    -ide-version=10813
'    -build-path C:\Users\Hardi\AppData\Local\Temp\arduino_build_696887
'    -warnings=more
'    -build-cache C:\Users\Hardi\AppData\Local\Temp\arduino_cache_610379
'    -prefs=build.warn_data_percentage=75
'    -prefs=runtime.tools.avr-gcc.path=C:\Program Files (x86)\Arduino\hardware\tools\avr
'    -prefs=runtime.tools.avr-gcc-7.3.0-atmel3.6.1-arduino7.path=C:\Program Files (x86)\Arduino\hardware\tools\avr
'    -prefs=runtime.tools.avrdude.path=C:\Users\Hardi\AppData\Local\Arduino15\packages\arduino\tools\avrdude\6.3.0-arduino17
'    -prefs=runtime.tools.avrdude-6.3.0-arduino17.path=C:\Users\Hardi\AppData\Local\Arduino15\packages\arduino\tools\avrdude\6.3.0-arduino17
'    -prefs=runtime.tools.arduinoOTA.path=C:\Program Files (x86)\Arduino\hardware\tools\avr
'    -prefs=runtime.tools.arduinoOTA-1.3.0.path=C:\Program Files (x86)\Arduino\hardware\tools\avr
'    -verbose
'    C:\Dat\MÃ¤rklin\Arduino\LEDs_Eisenbahn\extras\LEDs_AutoProg\LEDs_AutoProg.ino
'
' Was macht
'  der "-vid-pid" Schalter
'  der "-ide-version" Schalter

' 01.11.20:
' Wenn man nachträglich ein anderes Arduino AVR Board Paket installiert dann steht dieses im
' User Verzeichnis. Das kann man über das "Libraries" Sheet machen indem man die "arduino:avr" aktiviert
' und "Install Selected" drückt. Damit kann man neuere oder ältere Board Pakete installieren (Spalte "Required Version")

' => Die Boards stehen entweder in
'          C:\Program Files (x86)\Arduino\hardware\arduino\avr
' oder in  C:\Users\Hardi\AppData\Local\Arduino15\packages\arduino\hardware\avr\x.y.z   (x.y.z = Versions Nummer. Bsp.: 1.8.1)
' Die Datei "boards.local.txt" definiert eigene Board Varianten. Sie kann in einem der beiden Verzeichnisse stehen.
'
' Es ist auch möglich sie in ein beliebiges anderes Verzeichnis zu kopieren. Dann muss man aber die anderen
' Board Dateien ebenfalls dorthin kopieren und außerdem einen Link auf das Verzeichnis in dem Builder Kommando
' eintragen. Jürgen hat diese Methode verwendet. Dazu hat er das hardware Verzeichnis als Unterverzeichnis in
' das Sketch Verzeichnis kopiert und den folgenden Link hinzugefügt:
'   -hardware ".\hardware"
' Da es einige verschiedene Projekte gibt welche von dem Prog_Generator aus erzeugt werden können
' muss das Verzeichnis x mal kopiert werden.
'   Projekte: LEDs_AutoProg, ArduinoISP, 23_A.DCC_Interface, 23_A.Selectrix_Interface
' Die Arduino IDE hat aber den Link auf die kopierten Verzeichnisse nicht. Darum kennt sie den neuen
' Typ "ATmega328P (New Bootloader full Mem)" nicht.
'
' Neues Konzept:
' Es wird geprüft ob ein eigenes Board Paket existiert.
' Wenn kein eigenes Board Paket vorhanden ist, dann wird das Standard Paket kopiert
' Anschliesend wird die "boards.local.txt" Datei in das Verzeichnis kopiert. Aber nur wenn dort keine
' neuere Datei liegt (xcopy /d). Dadurch kann der Benutzer eigene Pakete hinzufügen.

'--------------------------------------------------
Public Function Packages_Dir_Available() As Boolean                         ' 07.10.21:
'--------------------------------------------------
  Dim Res As String
  Res = Dir(Environ(Env_USERPROFILE) & AppLoc_Ardu, vbDirectory)
  While Res <> "" And LCase(Res) <> "packages"
     Res = Dir
  Wend
  If LCase(Res) <> "" Then Packages_Dir_Available = True
End Function

'------------------------------------------------
Public Sub Create_Packages_Dir_if_not_Available()                           ' 07.10.21:
'------------------------------------------------
  If Not Packages_Dir_Available() Then
     CreateFolder Environ(Env_USERPROFILE) & AppLoc_Ardu & "packages\"
  End If
End Sub

Public Sub Create_Build(BoardName As String, fp As Integer)                                      ' 28.10.20: Jürgen (Old name: Create_PrivateBuild_cmd_if_missing)
'-------------------------------------
    If BoardName = "AM328" Then
        Create_Build_Arduino (fp)
        Exit Sub
    End If
    If BoardName = "PICO" Then                                                                  ' 17.04.21: Jürgen
        Create_Build_Pico (fp)
        Exit Sub
    End If
    Print #fp, "pause Invalid BoardType" & BoardName
    Print #fp, "exit /b 1"
End Sub

'-------------------------------------
Public Sub Create_Build_Arduino(fp As Integer)                                      ' 28.10.20: Jürgen (Old name: Create_PrivateBuild_cmd_if_missing)
'-------------------------------------
  Print #fp, "@ECHO OFF"
  Print #fp, "REM Fast Build command from Juergen"
  Print #fp, "REM ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
  Print #fp, "REM"
  Print #fp, "REM Compile and flash time 10 sec instead of 23 sec ! on Hardis laptop"
  Print #fp, "REM"
  Print #fp, "REM Speed up by"
  Print #fp, "REM - not using the Arduino Core"
  Print #fp, "REM - not checking the flash at the end (saves 3 sec)"
  Print #fp, "REM"
  Print #fp, "REM This file could be modified by the user to support special compiler switches"
  Print #fp, "REM It is called if the switch the ""Schnells Build und Upload verwenden:"" in the 'Config' sheet is enabled"
  Print #fp, "REM"
  Print #fp, "REM Parameter:               Example"
  Print #fp, "REM  1: Arduino EXE Path:    ""C:\Program Files (x86)\Arduino\"""
  Print #fp, "REM  2: Ino Name:            ""LEDs_AutoProg.ino"""
  Print #fp, "REM  3: Com port:            ""\\.\COM3"""
  Print #fp, "REM  4: Build options:       ""arduino:avr:nano:cpu=atmega328"""
  Print #fp, "REM  5: Baudrate:            ""57600"" or ""115200"""
  Print #fp, "REM  6: Arduino Library path ""%USERPROFILE%\Documents\Arduino\libraries"""
  Print #fp, "REM  7: CPU type:            ""atmega328p, atmega4809"                                            ' 28.10.20: Jürgen
  Print #fp, "REM  8: options:             ""noflash|norebuild"""                                     ' 19.12.21: Jürgen: Added noflash option
  Print #fp, "REM"
  Print #fp, "REM The program uses the captured and adapted command line from the Arduino IDE"
  Print #fp, "REM"
  Print #fp, ""
  Print #fp, "SET aHome=%~1"
  Print #fp, "SET fqbn=%~4"
  Print #fp, "SET lib=%~6"
  Print #fp, ""
  Print #fp, ""
  Print #fp, "SET aTemp=%USERPROFILE%\AppData\Local\Temp\MobaLedLib_build\ATMega"   ' 01.12.20: Added: "\ATMega" otherwise the esp32 fastbuild fails if the prog. is conpuled for the Nano
  Print #fp, "SET aCache=%USERPROFILE%\AppData\Local\Temp\MobaLedLib_cache\ATMega"  '    "            "
  Print #fp, "if not exist ""%aTemp%""  md ""%aTemp%"""
  Print #fp, "if not exist ""%aCache%"" md ""%aCache%"""
  Print #fp, ""
  
  Dim PackageDestDir As String ' Example: "C:\Users\Hardi\AppData\Local\Arduino15\packages\arduino\hardware\avr\1.8.3"           ' 01.11.20:
  PackageDestDir = GetShortPath(Environ(Env_USERPROFILE)) & AppLoc_Ardu & "packages\arduino\hardware\avr\" & Get_Std_Arduino_Lib_Ver()
  If Get_User_std_Arduino_Lib_Ver() = "" Then ' No own board package installed => Copy the standard board
     Print #fp, "robocopy ""%aHome%\hardware\arduino\avr"" """ & PackageDestDir & """ /mir /s >nul"
  End If
  Print #fp, "xcopy " & GetShortPath(ThisWorkbook.Path) & "\LEDs_AutoProg\boards.local.txt """ & PackageDestDir & "\"" /d /y >nul"  ' 06.03.21 Juergen: overwrite file if needed, don't promt user, has blocked the build
  
  Create_Packages_Dir_if_not_Available ' Create the 'packages' folder otherwise we get an error in the following 'GetShortPath()' call   07.10.21:
  
  Print #fp, ""
  Print #fp, "REM *** Call the arduino builder ***"
  Print #fp, """%aHome%\arduino-builder"" -compile -logger=human ^"
  Print #fp, "     -hardware ""%aHome%\hardware"" ^"
  Print #fp, "     -hardware """ & GetShortPath(Environ(Env_USERPROFILE) & AppLoc_Ardu & "packages") & """ ^"   ' 28.10.20: Jürgen
  Print #fp, "     -tools ""%aHome%\tools-builder"" ^"
  Print #fp, "     -tools ""%aHome%\hardware\tools\avr"" ^"
  Print #fp, "     -built-in-libraries ""%aHome%\libraries"" -libraries ""%LIB%"" ^"
  Print #fp, "     -fqbn=%fqbn% -build-path ""%aTemp%"" ^"
  Print #fp, "     -warnings=default ^"
  Print #fp, "     -build-cache ""%aCache%"" ^"
  Print #fp, "     -prefs=build.warn_data_percentage=75 ^"
  Print #fp, "     -prefs=runtime.tools.avrdude.path=""%aHome%\hardware\tools\avr"" ^"
  Print #fp, "     -prefs=runtime.tools.avr-gcc.path=""%aHome%\hardware\tools\avr""  ^"
  Print #fp, "     %2"
  Print #fp, ""
  Print #fp, ""
  Print #fp, "if ""%8""==""noflash"" goto :EOF"                                                            ' 19.12.21: Jürgen: add noflash option
  Print #fp, "if %errorlevel%==0 ("
  Print #fp, "   REM *** Flash program ***"
  Print #fp, "   REM -v = Verbose output. -v -v for more."
  Print #fp, "   REM -V = Do not verify.                      => Saves 3 sec"
  Print #fp, "   REM -D = Disable auto erase for flash memory"
  Print #fp, "   set extraArgs="                                                                           ' 28.10.20: Jürgen: New Block
  Print #fp, "   if ""%7""==""atmega4809"" ("
  Print #fp, "      echo Forcing reset using 1200bps open/close on port %3"
  Print #fp, "      mode %3 1200,n,8,1"
  Print #fp, "      set extraArgs=-cjtag2updi -e -Ufuse2:w:0x01:m -Ufuse5:w:0xC9:m -Ufuse8:w:0x00:m"
  Print #fp, "      goto flash"
  Print #fp, "   )"
  Print #fp, "   if ""%7""==""atmega328p"" ("
  Print #fp, "      set extraArgs=-carduino"
  Print #fp, "      goto flash"
  Print #fp, "   )"
  Print #fp, "   :flash"
  Print #fp, "   ""%aHome%\hardware\tools\avr/bin/avrdude"" -C""%aHome%\hardware\tools\avr/etc/avrdude.conf"" ^"
  Print #fp, "      -V -p%7 -P\\.\%3 -b%~5 -D -Uflash:w:""%aTemp%/%~2.hex"":i %extraArgs%"
  Print #fp, ")"
  Exit Sub

End Sub

'-------------------------------------
Public Sub Create_Build_Pico(fp As Integer)                     ' 17.04.21: Jürgen
'-------------------------------------

  Dim Board_Version As String
  Board_Version = Get_Lib_Version("rp2040:rp2040")
  
  Print #fp, "@ECHO OFF"
  Print #fp, "REM Fast Build command from Juergen"
  Print #fp, "REM ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
  Print #fp, "REM"
  Print #fp, "REM Speed up by"
  Print #fp, "REM - not using the Arduino Core"
  Print #fp, "REM - not checking the flash at the end (saves 3 sec)"
  Print #fp, "REM"
  Print #fp, "REM This file could be modified by the user to support special compiler switches"
  Print #fp, "REM It is called if the switch the ""Schnells Build und Upload verwenden:"" in the 'Config' sheet is enabled"
  Print #fp, "REM"
  Print #fp, "REM Parameter:               Example"
  Print #fp, "REM  1: Arduino EXE Path:    ""C:\Program Files (x86)\Arduino\"""
  Print #fp, "REM  2: Ino Name:            ""LEDs_AutoProg.ino"""
  Print #fp, "REM  3: Com port:            ""\\.\COM3"""
  Print #fp, "REM  4: Build options:       ""rp2040:rp2040:rpipico:flash=2097152_0,freq=125,dbgport=Disabled,dbglvl=None"""
  Print #fp, "REM  5: Baudrate:            ""115200"""
  Print #fp, "REM  6: Arduino Library path ""%USERPROFILE%\Documents\Arduino\libraries"""
  Print #fp, "REM  7: CPU type:            ""rp2040"
  Print #fp, "REM  8: options:             ""noflash"""                                     ' 19.12.21: Jürgen: Added noflash option
  Print #fp, "REM"
  Print #fp, "REM The program uses the captured and adapted command line from the Arduino IDE"
  Print #fp, "REM"
  Print #fp, ""
  Print #fp, "SET aHome=%~1"
  Print #fp, "SET fqbn=%~4"
  Print #fp, "SET lib=%~6"
  Print #fp, ""
  Print #fp, ""
  Print #fp, "call :short aTemp ""%USERPROFILE%\AppData\Local\Temp\MobaLedLib_build\Pico"""
  Print #fp, "SET aCache=%aTemp%\cache"
  Print #fp, "call :short packages ""%USERPROFILE%" + AppLoc_Ardu + "packages"""
  Print #fp, "if not exist ""%aTemp%""  md ""%aTemp%"""
  Print #fp, "if not exist ""%aCache%"" md ""%aCache%"""
  Print #fp, ""
  
  Create_Packages_Dir_if_not_Available ' Create the 'packages' folder otherwise we get an error in the following 'GetShortPath()' call   20.10.21:
  
  Print #fp, "REM *** Call the arduino builder ***"
  Print #fp, """%aHome%\arduino-builder"" -compile -logger=human ^"
  Print #fp, "     -hardware ""%aHome%\hardware"" ^"
  Print #fp, "     -hardware """ & GetShortPath(Environ(Env_USERPROFILE) & AppLoc_Ardu & "packages") & """ ^"   ' 28.10.20: Jürgen
  Print #fp, "     -tools ""%aHome%\tools-builder"" ^"
  Print #fp, "     -tools ""%aHome%\hardware\tools\avr"" ^"
  Print #fp, "     -built-in-libraries ""%aHome%\libraries"" -libraries ""%LIB%"" ^"
  Print #fp, "     -fqbn=%fqbn% -build-path ""%aTemp%"" ^"
  Print #fp, "     -warnings=default ^"
  Print #fp, "     -build-cache ""%aCache%"" ^"
  Print #fp, "     -prefs=build.warn_data_percentage=75 ^"
  Print #fp, "     %2"
  Print #fp, ""
  Print #fp, ""
  Print #fp, "if ""%8""==""noflash"" goto :EOF"                                                            ' 19.12.21: Jürgen: add noflash option
  Print #fp, "if %errorlevel%==0 ("
  Print #fp, "   REM *** Flash program ***"
  Print #fp, "   :flash"
  Print #fp, "   ""%packages%\rp2040\tools\pqt-python3\1.0.1-base-3a57aed\python3"" ""%packages%\rp2040\hardware\rp2040\" & Board_Version & "\tools\uf2conv.py"" ^"
  Print #fp, "   --serial %3 --family RP2040 --deploy ""%aTemp%\LEDs_AutoProg.ino.uf2"""
  Print #fp, ")"
  Print #fp, "goto :eof"
  Print #fp, ""
  Print #fp, ":short"
  Print #fp, "set %1=%~s2"
  Print #fp, "goto :eof"
  Print #fp, ""
End Sub


'UT--------------------------------------------------
Private Sub Test_Create_PrivateBuild_cmd_if_missing()
'UT--------------------------------------------------
  Dim Name As String
  Name = ThisWorkbook.Path & "\LEDs_AutoProg\privateBuild.cmd"
  Dim fp As Integer
  Open Name For Output As #fp
  Create_Build "arduino", fp
  Close #fp
End Sub



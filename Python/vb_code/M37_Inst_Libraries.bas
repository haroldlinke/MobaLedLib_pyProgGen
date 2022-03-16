Attribute VB_Name = "M37_Inst_Libraries"
Option Explicit
Option Compare Text ' Case insensitive compare

' Install all required Libraries
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Make sure that all required libraries are installed
' - It's called every time when the program is started
' - All missing libraries are installed, also if they are needed only
'   for a special project. Also the libraries needed for the Pattern_Configurator
'   => more likely that the DetectVers are compatible
' - Hidden sheet "Libraries" contains all data


' ToDo:
' ~~~~~
' - Installation In SketchDir
'   - Test
' - Aus irgend einem Grund funktioniert das Installieren der MobaLedLib mit einer "Required Version" nicht.
'   Bei der "FastLED" und der "NmraDcc" geht es.
'   Es geht auch nicht von Excel aus. Es kommt die Fehlermeldung:
'      "Library MobaLedLib is already installed in: E:\Test Arduino Lib mit Ã¤\libraries\MobaLedLib"
'   => Die Bibliothek muss von Hand gelöscht werden
'   Manchmal geht es aber auch ?!?


Private Const First_Dat_Row = 9

Private Const SelectRow_Col = 2
Private Const Installed_Col = 3
Private Const Lib_Board_Col = 4
Private Const Libr_Name_Col = 5
Private Const Test_File_Col = 6
Private Const Reque_Ver_Col = 7
Private Const DetectVer_Col = 8
Private Const Other_Src_Col = 9

Private Const UPDATE_LIB_CMD_NAME = "Update_Libraries.cmd"
Private Const RESTART_PROGGEN_CMD = "Restart_ProgGen.cmd"

Private UnzipList As String


Private Update_Time As Variant

Public Const WIN7_COMPATIBLE_DOWNLOAD = True                                ' 20.06.20:

'----------------------------------------------------------------------
Public Function Is_Libraries_Select_Column(ByVal Target As Excel.Range)
'----------------------------------------------------------------------
  If Target.CountLarge = 1 Then
     Is_Libraries_Select_Column = Target.Row >= First_Dat_Row And Target.Column = SelectRow_Col
  End If
End Function

' Erkennung des Standard Arduino Boards
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Das Verwendete Nano board hängt von der installierten Arduino IDE ab. Es kann aber auch
' nachträglich eine anderes Board installiert werden. Das macht es kompliziert.
'
' Die installierte Arduino IDE Version kann man aus der Datei auslesen:
'  C:\Program Files (x86)\Arduino\lib\version.txt enthält 1.8.12
' Die Folgende Tabelle enthält die Zusammenhänge.
'
' IDE      Board   GCC                         o.k. FastLed 3.3.3
' ~~~~     ~~~~~   ~~~                         ~~~~~~~~~~~~~~~~~~
' 1.8.13   1.8.3   7.3.0-atmel3.6.1-arduino5
' 1.8.12   1.8.2   7.3.0-atmel3.6.1-arduino5   Yes
' 1.8.11   1.8.2   7.3.0-atmel3.6.1-arduino5   Yes
' 1.8.10   1.8.1   7.3.0-atmel3.6.1-arduino5   Yes
' 1.8.8    1.6.23  5.4.0-atmel3.6.1-arduino2   No   #define FL_FALLTHROUGH __attribute__ ((fallthrough));
'
' Arduino Releases: https://github.com/arduino/Arduino/releases
'
' In der Datei
'   C:\Program Files (x86)\Arduino\hardware\package_index_bundled.json
' findet man die Standard Board Version. Hier für die IDE 1.8.8
'   "version": "1.6.23",
' Bei der IDE Version ist
'   "version": "1.6.23",
' eingetragen
'
' Wenn ein anderes Board installiert wurde, dann findet man die Version hier:
'  "C:\Users\Hardi\AppData\Local\Arduino15\packages\arduino\hardware\avr\1.8.1"
'
' Boards Manager Anzeigen von der Arduino IDE 1.8.12:
' Version 1.8.1:
'   Arduino AVR Boards
'   by Arduino Version 1.8.1
' Version 1.8.2:
'   Built-In by Arduino Version 1.8.2
'
' => Das 'Built-In' zeigt, dass es die Standard mäßig in der Arduino IDE 1.8.12 enthalten Board Version ist

'-------------------------------------------------------
Public Function Get_User_std_Arduino_Lib_Ver() As String
'-------------------------------------------------------
  Dim OtherBoardDir As String
  OtherBoardDir = Environ(Env_USERPROFILE) & AppLoc_Ardu & "packages\arduino\hardware\avr\"
  Get_User_std_Arduino_Lib_Ver = Get_First_SubDir(OtherBoardDir)
End Function

'---------------------------------------------------
Public Function Get_Std_Arduino_Lib_Ver() As String
'---------------------------------------------------
' Std. Boards (Nano, Uno, ...)
' The C:\Users\Hardi\AppData\Local\Arduino15\packages\arduino\hardware\avr\
  Dim BoardVer As String, ArduinoDir As String
  ArduinoDir = FilePath(Find_ArduinoExe())
  BoardVer = Get_User_std_Arduino_Lib_Ver
  If BoardVer = "" Then
     Dim Package_Index_Bundled As String
     Package_Index_Bundled = Read_File_to_String(ArduinoDir & "hardware\package_index_bundled.json")
     BoardVer = Replace(Replace(Get_Ini_Entry(Package_Index_Bundled, """version"": """), """", ""), ",", "")
  End If
  Get_Std_Arduino_Lib_Ver = BoardVer
End Function

'------------------------------------
Private Sub Update_General_Versions()
'------------------------------------
' Update the general versions in the Libraries sheet
' - Arduino IDE
' - Std. Boards (Nano, Uno, ...)
  Dim Sh As Worksheet
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  
  ' Arduino IDE
  Dim ArduinoDir As String, ArduinoVer As String
  ArduinoDir = FilePath(Find_ArduinoExe())
  ArduinoVer = Read_File_to_String(ArduinoDir & "lib\version.txt")
  Sh.Range("Arduino_IDE_Ver") = ArduinoVer
  
    
  ' Std. Boards (Nano, Uno, ...)
  Sh.Range("Std_Boards_Ver") = Get_Std_Arduino_Lib_Ver()
End Sub


'-------------------------------------------------------------------------------------------------------------
Private Function Get_DetectVer_form_library_properties(LibDir As String) As String
'-------------------------------------------------------------------------------------------------------------
  Dim Name As String, FileStr As String
  Name = LibDir & "library.properties"
  If Dir(Name) <> "" Then
     FileStr = Read_File_to_String(Name)
     If FileStr <> "#ERROR#" Then
        Get_DetectVer_form_library_properties = Get_Ini_Entry(FileStr, "version=")
     End If
  Else: Get_DetectVer_form_library_properties = "?"
  End If
End Function

'UT--------------------------------------
Private Sub Test_Get_State_of_Board_Row()
'UT--------------------------------------
  Get_State_of_Board_Row 21
End Sub

'----------------------------------------------
Private Sub Get_State_of_Board_Row(Row As Long)
'----------------------------------------------
' Don't know how the boards are treated. In the boards manager of the Arduino IDE the first (in alphabetical order)
' board is shown. Old "libraries" directories are not (always) deleted if a new version is installed ;-(
' But (sometimes) the old directories are empty.
' The first not empty directory is listed.
'
' We assume the following structure:
'                                                  Name                  Processor  Version             TestFile
'                                                  ~~~~~~~~~~            ~~~~~~~~~  ~~~~~~~             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' "C:\Users\Hardi\AppData\Local\Arduino15\packages\ATTinyCore\ hardware\ avr\       1.3.2\   libraries\ ATTinyCore\src\ATTinyCore.h"
' "C:\Users\Hardi\AppData\Local\Arduino15\packages\esp8266   \ hardware\ esp8266\   2.3.0\   libraries\ ESP8266AVRISP\src\ESP8266AVRISP.h"
' "C:\Users\Hardi\AppData\Local\Arduino15\packages\arduino   \ hardware\ megaavr\   1.6.26\  libraries\ Wire\src\Wire.h"
  Dim Sh As Worksheet, BoardDir As String, TestFile As String, Board_and_Proc As String, Board As String, ProcessorTyp As String
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  Board_and_Proc = Sh.Cells(Row, Libr_Name_Col)
  Board = Split(Board_and_Proc, ":")(0)
  ProcessorTyp = Split(Board_and_Proc, ":")(1)
  
  If Board = "arduino" And ProcessorTyp = "avr" Then
     Sh.Cells(Row, Installed_Col) = 1
     Sh.Cells(Row, DetectVer_Col) = Get_Std_Arduino_Lib_Ver
     Exit Sub
  End If
  
  TestFile = Sh.Cells(Row, Test_File_Col)
  BoardDir = Environ(Env_USERPROFILE) & AppLoc_Ardu & "packages\" & Board & "\hardware\"
  Dim VerList As String
  Dim Res As String
  Res = Dir(BoardDir & ProcessorTyp & "\*.*", vbDirectory) ' The Dir() result seames to be sorted
  While Res <> ""
     If left(Res, 1) <> "." Then
        VerList = VerList & Res & vbTab
     End If
     Res = Dir() ' Mit Excel für Mac 2016 wird der ursprüngliche Dir-Funktionsaufruf erfolgreich ausgeführt. Nachfolgende Aufrufe zum Durchlaufen des angegebenen Verzeichnisses führen jedoch zu einem Fehler. Dies ist leider ein bekanntes Problem.
  Wend
  VerList = DelLast(VerList)
  Dim Ver As Variant
  For Each Ver In Split(VerList, vbTab)
      Dim DirName As String
      DirName = BoardDir & ProcessorTyp & "\" & Ver & "\libraries\"
      If Not Dir_is_Empty(DirName) Then
         With Sh.Cells(Row, Installed_Col)
           If Dir(DirName & TestFile) <> "" Then
                 Sh.Cells(Row, DetectVer_Col) = Ver
                 .Value = 1
           Else: .Value = ""
           End If
         End With
         Exit Sub
      End If
  Next Ver
End Sub

'----------------------------------------------
Private Sub Get_State_of_BoardExtras_Row(Row As Long)
'----------------------------------------------
  Dim Sh As Worksheet, BoardDir As String, TestFile As String, Board_and_Proc As String, Board As String, ExtraType As String
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  Board_and_Proc = Sh.Cells(Row, Libr_Name_Col)
  Board = Split(Board_and_Proc, ":")(0)
  ExtraType = Split(Board_and_Proc, ":")(1)
  
  TestFile = Sh.Cells(Row, Test_File_Col)
  BoardDir = Environ(Env_USERPROFILE) & AppLoc_Ardu & "packages\" & Board
  Dim VerList As String
  Dim Res As String
  Res = Dir(BoardDir & "\" & ExtraType & "\*.*", vbDirectory) ' The Dir() result seames to be sorted
  While Res <> ""
     If left(Res, 1) <> "." Then
        VerList = VerList & Res & vbTab
     End If
     Res = Dir() ' Mit Excel für Mac 2016 wird der ursprüngliche Dir-Funktionsaufruf erfolgreich ausgeführt. Nachfolgende Aufrufe zum Durchlaufen des angegebenen Verzeichnisses führen jedoch zu einem Fehler. Dies ist leider ein bekanntes Problem.
  Wend
  VerList = DelLast(VerList)
  Dim Ver As Variant
  For Each Ver In Split(VerList, vbTab)
      Dim DirName As String
      DirName = BoardDir & "\" & ExtraType & "\" & Ver
      If Not Dir_is_Empty(DirName) Then
         With Sh.Cells(Row, Installed_Col)
           If Dir(DirName & "\" & TestFile) <> "" Then
                 Sh.Cells(Row, DetectVer_Col) = Ver
                 .Value = 1
           Else: .Value = ""
           End If
         End With
         Exit Sub
      End If
  Next Ver
End Sub


'---------------------------------------------------
Private Function Get_All_Library_States() As Boolean
'---------------------------------------------------
' Get the states of all libraries:
' - Installed
' - DetectVer
  If Read_Sketchbook_Path_from_preferences_txt() = False Then Exit Function
  ThisWorkbook.Sheets(LIBRARYS__SH).Range("Sketchbook_Path") = Sketchbook_Path
  Dim LibrariesDir As String
  LibrariesDir = Sketchbook_Path & "\libraries\"
  Dim Row As Long, Sh As Worksheet
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  Row = First_Dat_Row
  While Sh.Cells(Row, Libr_Name_Col) <> ""
       Dim TestFile As String
       TestFile = Sh.Cells(Row, Test_File_Col)
       Sh.Cells(Row, DetectVer_Col) = ""
       Sh.Cells(Row, Installed_Col) = ""
       If InStr(Sh.Cells(Row, Lib_Board_Col), "L") > 0 Then
            ' *** Library ***
            Dim LibDir As String
            LibDir = LibrariesDir & Sh.Cells(Row, Libr_Name_Col) & "\"
            With Sh.Cells(Row, Installed_Col)
              .Value = ""
              On Error GoTo ErrDontExist                                    ' 07.06.20:
              If Dir(LibDir, vbDirectory) <> "" Then
                 If Dir(LibDir & TestFile) <> "" Or Dir(LibDir & "src\" & TestFile) <> "" Then
                    .Value = "1"
                 End If
              End If
DontExist:
              On Error GoTo 0
            End With
            If Sh.Cells(Row, Installed_Col) > 0 Then Sh.Cells(Row, DetectVer_Col) = Get_DetectVer_form_library_properties(LibDir)
       ElseIf InStr(Sh.Cells(Row, Lib_Board_Col), "BE") > 0 Then
            ' *** Board Extras ***
            Get_State_of_BoardExtras_Row Row
       ElseIf InStr(Sh.Cells(Row, Lib_Board_Col), "B") > 0 Then
            ' *** Board ***
            Get_State_of_Board_Row Row
       End If
       Row = Row + 1
  Wend
  Get_All_Library_States = True
  ThisWorkbook.Sheets(LIBRARYS__SH).Range("Last_Update_Time") = Date + Time
  Exit Function

ErrDontExist:                                                               ' 07.06.20:
  MsgBox Get_Language_Str("Fehler beim lesen des Verzeichnisses:") & vbCr & _
         "  '" & LibDir & "'" & vbCr & _
         "Error Nr: " & Err.Number & vbCr & _
         Err.Description, _
         vbCritical, Get_Language_Str("Fehler beim lesen des Verzeichnisses:")
  Resume DontExist
End Function

'---------------------------------------------------------------------------------------------------------------
Public Function Check_if_curl_is_Available_and_gen_Message_if_not(Name As String, InstLink As String) As Boolean
'---------------------------------------------------------------------------------------------------------------
  If Win10_or_newer() Then
     Check_if_curl_is_Available_and_gen_Message_if_not = True
     Exit Function
  End If
  MsgBox Replace(Get_Language_Str("Die Programme 'curl' und 'tar' sind erst ab Win10 verfügbar. " & _
                                  "Darum kann '#1#' nicht automatisch installiert werden ;-(" & vbCr & _
                                  "Es kann manuell von hier installiert werden:"), "#1#", Name) & vbCr & _
                                  "  '" & InstLink & "'" & vbCr & _
                                  vbCr, vbInformation, _
                 Get_Language_Str("Windows Version ist zu alt. Keine automatische Installation möglich")
End Function


'-------------------------------------------------------------------
Private Sub Add_Update_from_Other_Source(fp As Integer, Row As Long)
'-------------------------------------------------------------------
' Creates:
'   powershell Invoke-WebRequest "https://github.com/merose/AnalogScanner/archive/master.zip" -o:AnalogScanner.zip
' Or if WIN7_COMPATIBLE_DOWNLOAD is not defined:
'   curl -LJO https://github.com/merose/AnalogScanner/archive/master.zip
'   tar  -xf AnalogScanner-master.zip
'   ren  AnalogScanner-master AnalogScanner

  Dim Sh As Worksheet, LibName As String, InstLink As String
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  LibName = Sh.Cells(Row, Libr_Name_Col)
  InstLink = Trim(Sh.Cells(Row, Other_Src_Col))
  Print #fp, ""
  Print #fp, "ECHO " & Replicate("*", Len("Updating " & LibName & "..."))
  Print #fp, "ECHO Updating " & LibName & "..."
  Print #fp, "ECHO " & Replicate("*", Len("Updating " & LibName & "..."))
  Print #fp, "if EXIST " & LibName & "\NUL ("
  Print #fp, "   ECHO deleting old directory " & LibName & "\"
  Print #fp, "   rmdir " & LibName & "\ /s /q"
  Print #fp, "   timeout /T 3 /nobreak" ' Wait until the directory is deleted     ' 04.08.20:
  Print #fp, ")"
  Print #fp, "if EXIST " & LibName & "\NUL ("
  Print #fp, "   ECHO Error deleting old directory " & LibName & "\"
  Print #fp, "   ECHO For some reasons the directory could not be deleted ;-("
  Print #fp, "   ECHO Check if an other program is active which prevents the deleting"
  Print #fp, "   ECHO of the directory"
  Print #fp, "   ECHO."
  Print #fp, "   ECHO Going to try a second time"
  Print #fp, "   PAUSE"
  Print #fp, "   rmdir " & LibName & "\ /s /q"
  Print #fp, "   timeout /T 3 /nobreak" ' Wait until the directory is deleted
  Print #fp, ")"
  Print #fp, "if EXIST " & LibName & "\NUL ("
  Print #fp, "   COLOR 4F" ' Yellow on Red"
  Print #fp, "   ECHO Error: Still not able to delete the old directory " & LibName & "\   ;-((("
  Print #fp, "   PAUSE"
  Print #fp, ")"

  If WIN7_COMPATIBLE_DOWNLOAD Then
    Print #fp, "powershell Invoke-WebRequest """ & InstLink & """ -o:" & LibName & ".zip"   ' 20.06.20:
    Print #fp, "ECHO Invoke-WebRequest result: %ERRORLEVEL%"
    Print #fp, "IF ERRORLEVEL 1 Goto ErrorMsg"
    UnzipList = UnzipList & LibName & vbTab  ' The file is unzipped later in excel
  Else
    If Check_if_curl_is_Available_and_gen_Message_if_not(LibName, InstLink) = False Then Exit Sub
    Print #fp, "curl -LJ """ & InstLink & """ --output " & LibName & ".zip"
    Print #fp, "ECHO curl result: %ERRORLEVEL%"
    Print #fp, "IF ERRORLEVEL 1 Goto ErrorMsg"
    Print #fp, "tar -xmf " & LibName & ".zip"
    Print #fp, "ECHO tar  result: %ERRORLEVEL%"
    Print #fp, "IF ERRORLEVEL 1 Goto ErrorMsg"
    Print #fp, "ren " & LibName & "-master " & LibName
    Print #fp, "ECHO ren  result: %ERRORLEVEL%"
    Print #fp, "IF ERRORLEVEL 1 GOTO ErrorMsg"
    Print #fp, "if EXIST " & LibName & ".zip del " & LibName & ".zip"
  End If
  Print #fp, "ECHO."
'  Print #fp, "PAUSE" ' Debug
  Print #fp, ""
End Sub


'---------------------------
Private Sub Proc_UnzipList()                                                ' 20.06.20:
'---------------------------
  UnzipList = DelLast(UnzipList)
  Dim LibName As Variant, LibName_with_path As String
  For Each LibName In Split(UnzipList, vbTab)
      LibName_with_path = Get_Ardu_LibDir() & LibName
      If Not UnzipAFile(LibName_with_path & ".zip", Get_Ardu_LibDir()) Then Exit Sub
      If Dir(LibName_with_path & "-master", vbDirectory) <> "" Then
        On Error GoTo RenameErr
        Name LibName_with_path & "-master" As LibName_with_path
      ElseIf Dir(LibName_with_path & "-beta", vbDirectory) <> "" Then       ' 19.11.21 Juergen support of BETA update directly from github
        On Error GoTo RenameErr
        Name LibName_with_path & "-beta" As LibName_with_path
      Else
        MsgBox Replace(Get_Language_Str("Fehler: Das Verzeichnis '#1#' wurde nicht erzeugt beim entzippen von:"), "#1#", LibName & "-master") & vbCr & _
                  "  '" & LibName_with_path & ".zip", vbCritical, Get_Language_Str("Fehler beim entzippen")
      End If
      On Error Resume Next
      Kill LibName_with_path & ".zip"  ' Delete the ZIP file
      On Error GoTo 0
  Next
  Exit Sub
  
RenameErr:
  MsgBox Get_Language_Str("Fehler beim Umbenennen des Verzeichnisses:") & vbCr & _
                          "  '" & LibName_with_path & "-master'" & vbCr & _
                          "nach '..." & LibName & "'", vbCritical, _
         Get_Language_Str("Verzeichnis kann nicht umbenannt werden")
  Resume Next
End Sub

'-------------------------------
Public Sub Init_Libraries_Page()
'-------------------------------
  ThisWorkbook.Sheets(LIBRARYS__SH).CheckBoxes("Check Box 10") = xlOff
End Sub


' Bei der installation des ATTinys kommt folgende Fehlermeldung:
' Warnung: nicht vertrauenswürdiger Beitrag, Skript-Ausführung wird übersprungen (C:\Users\Hardi\AppData\Local\Arduino15\packages\ATTinyCore\tools\micronucleus\2.5-azd1b\post_install.bat)
'

'------------------------------------------------------------------------
Private Function Create_Do_Update_Script(Pause_at_End As Boolean) As Long
'------------------------------------------------------------------------
' Updates all selected libraries
'
' Arduino parameters see:
'  https://github.com/arduino/Arduino/blob/master/build/shared/manpage.adoc
' Return: -1 in case of an error
'          0 if nothing has to be updated
'          n number of necessary updates
  Dim fp As Integer, Name As String, UpdCnt As Long
  fp = FreeFile
  Name = ThisWorkbook.Path & "\" & UPDATE_LIB_CMD_NAME
  On Error GoTo WriteError
  Open Name For Output As #fp
  Print #fp, "@ECHO OFF"
  Print #fp, "Color 80" ' Black on bright Gray background (See: https://ss64.com/nt/color.html)
  Print #fp, "REM This file was generated by '" & ThisWorkbook.Name & "'  " & Time
  Print #fp, "REM"
  Print #fp, "REM It updates/installs all required libraries for the MobaLedLib projects."
  Print #fp, "REM"
  Print #fp, "REM Attention:"
  Print #fp, "REM This program must be started from the arduino libraries directory"
  Print #fp, "REM"
  Print #fp, ""
  If Win10_or_newer() Then                                                  ' 28.06.20: The find command dosn't work with this code page at Win7 for some reasons. It waits endless ?!?
     Print #fp, "CHCP 65001 >NUL"
  End If
  Dim LibList As String, BrdList As String, URLList As String, Row As Long, Sh As Worksheet, ForceReinstall As Boolean
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  ForceReinstall = False
  If Sh.CheckBoxes("Check Box 10").Value = xlOn Then Pause_at_End = True   ' Wait at End Checkbox
  Row = First_Dat_Row
  While Sh.Cells(Row, Libr_Name_Col) <> ""
        If Sh.Cells(Row, SelectRow_Col) <> "" Then
           UpdCnt = UpdCnt + 1
           If InStr(Sh.Cells(Row, Lib_Board_Col), "L") > 0 Then
                If Sh.Cells(Row, Other_Src_Col) = "" Then
                     LibList = LibList & """" & Sh.Cells(Row, Libr_Name_Col)
                     If Trim(Sh.Cells(Row, Reque_Ver_Col)) <> "" Then LibList = LibList & ":" & Sh.Cells(Row, Reque_Ver_Col)
                     LibList = LibList & ""","
                     
                     ' 19.10.21: Juergen Workaround for problem that libraries with 'unknown' versions are not updated
                     If Sh.Cells(Row, Libr_Name_Col) = "NmraDcc" _
                        And (Sh.Cells(Row, DetectVer_Col) = "2.0.7" Or Sh.Cells(Row, DetectVer_Col) = "2.0.8") Then ' 25.10.21: Added "2.0.7" because this version was detected on Michael computer
                         Del_Folder Sketchbook_Path & "\libraries\" & Sh.Cells(Row, Libr_Name_Col)
                     End If
                     
                     ForceReinstall = True
                Else ' Extract from other source
                     Add_Update_from_Other_Source fp, Row
                     ForceReinstall = True
                End If
           ElseIf InStr(Sh.Cells(Row, Lib_Board_Col), "BE") > 0 Then
                ' skip these extra files
           ElseIf InStr(Sh.Cells(Row, Lib_Board_Col), "B") > 0 Then
                ' Board
                BrdList = BrdList & Sh.Cells(Row, Libr_Name_Col)
                If Trim(Sh.Cells(Row, Reque_Ver_Col)) <> "" Then BrdList = BrdList & ":" & Sh.Cells(Row, Reque_Ver_Col)
                BrdList = BrdList & ","
                ForceReinstall = True
                If Sh.Cells(Row, Other_Src_Col) <> "" Then
                   If InStr(URLList, Sh.Cells(Row, Other_Src_Col) & ",") = 0 Then
                      URLList = URLList & Sh.Cells(Row, Other_Src_Col) & ","
                   End If
                End If
                
                Create_Packages_Dir_if_not_Available ' Create the 'packages' folder. This is necessary if the user has no other packages installed up to now   ' 07.10.21:
                
                Dim BoardDir As String, Board_and_Proc As String, Board As String
                Board_and_Proc = Sh.Cells(Row, Libr_Name_Col)
                Board = Split(Board_and_Proc, ":")(0)
                BoardDir = Environ(Env_USERPROFILE) & AppLoc_Ardu & "packages\" & Board
                If Dir(BoardDir, vbDirectory) <> "" Then
                   Debug.Print "Deleting: " & BoardDir
                   Del_Folder BoardDir ' Deleting the old directory
                End If
           End If

        End If
        Row = Row + 1
  Wend
  If ForceReinstall = True Then                  ' 11.03.21 Juergen: force an ESP32 rebuild
    If Dir(Environ(Env_USERPROFILE) & "\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache") <> "" Then
       Kill Environ(Env_USERPROFILE) & "\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache"
    End If
  End If

  ' *** Libraries ***
  If LibList <> "" Then
     LibList = DelLast(LibList)
     Print #fp, "ECHO ************************************"
     Print #fp, "ECHO  Installing the following libraries"
     Print #fp, "ECHO ************************************"
     Dim Lib As Variant
     For Each Lib In Split(LibList, ",")
        Print #fp, "ECHO   " & Replace(Lib, """", "")
     Next Lib
     Print #fp, "ECHO."
     ' 09.03.21 Juergen: delete cache file to force an ESP32 rebuild, otherwise prebuild library versions would still be used
     Print #fp, "@if exist ""%USERPROFILE%\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache"" del ""%USERPROFILE%\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache"""
     Print #fp, """" & Find_ArduinoExe() & """";
     Print #fp, " --install-library " & LibList;
     Print #fp, " 2>&1 | find /v "" StatusLogger "" | find /v "" INFO c.a"" | find /v "" WARN p.a"" | find /v "" WARN c.a""" ' Hide debug messages
     Print #fp, "ECHO."
     Print #fp, "ECHO Error %ERRORLEVEL%"
     Print #fp, "IF ERRORLEVEL 1 Goto ErrorMsg"
     Print #fp, ""
  End If
 
 ' *** Boards ***
  If BrdList <> "" Then
     BrdList = DelLast(BrdList)
     URLList = DelLast(URLList)
     Print #fp, "ECHO *********************************"
     Print #fp, "ECHO  Installing the following boards"
     Print #fp, "ECHO *********************************"
     Dim Brd As Variant
     For Each Brd In Split(BrdList, ",")
        Print #fp, "ECHO   " & Brd
     Next Brd
     Print #fp, "ECHO."
     ' 09.03.21 Juergen: delete cache file to force an ESP32 rebuild, otherwise prebuild library versions would still be used
     Print #fp, "@if exist ""%USERPROFILE%\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache"" del ""%USERPROFILE%\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache"""
     For Each Brd In Split(BrdList, ",") ' Install each board separately    ' 07.10.21:
        Print #fp, """" & Find_ArduinoExe() & """";
        Print #fp, " --install-boards " & Brd;
        If URLList <> "" Then Print #fp, " --pref ""boardsmanager.additional.urls=" & URLList & """";
        Print #fp, " 2>&1 | find /v "" StatusLogger "" | find /v "" INFO c.a"" | find /v "" WARN p.a"" | find /v "" WARN c.a""" ' Hide debug messages
        Print #fp, "ECHO."
        Print #fp, "ECHO Error %ERRORLEVEL%"
        Print #fp, "IF ERRORLEVEL 1 Goto ErrorMsg"
        Print #fp, ""
     Next Brd
  End If
  'Print #fp, "Pause" ' Debug
  If Pause_at_End Then Print #fp, "Pause"
  Print #fp, "Exit"
  Print #fp, ""
  Print #fp, ":ErrorMsg"
  Print #fp, "   COLOR 4F"
  Print #fp, "   ECHO   ****************************************"
  Print #fp, "   ECHO    Da ist was schief gegangen ;-("
  Print #fp, "   ECHO   ****************************************"
  Print #fp, "   Pause"
  Close #fp
  Create_Do_Update_Script = UpdCnt
  Exit Function

WriteError:
  Close #fp
  MsgBox Get_Language_Str("Fehler beim Schreiben der Datei '") & Name & "'", vbCritical, Get_Language_Str("Fehler beim erzeugen der Arduino Start Datei")
  Create_Do_Update_Script = -1
End Function

'UT---------------------------------------
Private Sub Test_Create_Do_Update_Script()
'UT---------------------------------------
  Create_Do_Update_Script True
End Sub


'---------------------------------------------------------------------------
Private Function Get_Original_Name_from_TestFile(LibDir As String) As String
'---------------------------------------------------------------------------
  Dim Row As Long, Sh As Worksheet
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  Row = First_Dat_Row
  While Sh.Cells(Row, Libr_Name_Col) <> ""
       If InStr(Sh.Cells(Row, Lib_Board_Col), "L") > 0 Then
            Dim TestFile As String
            TestFile = Sh.Cells(Row, Test_File_Col)
            If Dir(LibDir & TestFile) <> "" Or Dir(LibDir & "src\" & TestFile) <> "" Then
               Get_Original_Name_from_TestFile = Sh.Cells(Row, Libr_Name_Col)
               Exit Function
            End If
       End If
       Row = Row + 1
  Wend
End Function

'-----------------------------------------------------------
Private Sub Correct_one_Temp_Arduino_nr_Dir(LibDir As String)
'-----------------------------------------------------------
#If 1 Then ' Das Umbenennen am Ende geht nicht ?!?
  Dim Org_LibName As String
  Org_LibName = Get_Original_Name_from_TestFile(Sketchbook_Path & "\libraries\" & LibDir & "\")
  If Org_LibName <> "" Then
     Dim Org_LibPath As String
     Org_LibPath = Sketchbook_Path & "\libraries\" & Org_LibName
     If Dir(Org_LibPath, vbDirectory) <> "" Then
        Debug.Print "Deleting old library: " & Org_LibPath
        Del_Folder Org_LibPath
     End If
     Debug.Print "Rename directory '" & LibDir & "' to '" & Org_LibName & "'"
     ChDir Sketchbook_Path & "\libraries\"
     Debug.Print Dir(LibDir, vbDirectory)
     On Error GoTo ErrMsg
     Name LibDir As Org_LibName  ' Geht nicht !?!
     On Error GoTo 0
  End If
  Exit Sub
ErrMsg:
  MsgBox Get_Language_Str("Fehler beim umbenennen des temporären Verzeichnisses:") & vbCr & _
                          "  '" & Sketchbook_Path & "\libraries\" & LibDir & "'" & vbCr & _
                          Get_Language_Str("Vermutlich ist irgend eine Datei in dem Verzeichniss " & _
                          "durch ein Programm gesperrt ;-(") & vbCr & _
                          Get_Language_Str("Das Verzeichnis muss von Hand gelöscht werden"), vbCritical, _
                          Get_Language_Str("Temporäres Verzeichnis konnte nicht umbenannt werden")
#Else
     Debug.Print "Deleting Temp directory: " & Sketchbook_Path & "\libraries\" & LibDir & "\"
     Del_Folder Sketchbook_Path & "\libraries\" & LibDir & "\"
#End If
End Sub

'UT--------------------
Private Sub Test_name()
'UT--------------------
  ChDrive "E:"
  ChDir "E:\Test Arduino Lib mit ä\libraries"
  Name "FastLED\" As "Arduino_12345"
End Sub


'------------------------------------------
Private Sub Correct_Temp_Adrduino_nr_Dirs()
'------------------------------------------
' Sometimes the instalation fails an a "Arduino_<nr>" directory is created.
' Unfortunately an update with a new version is not possible
  If CheckArduinoHomeDir() = False Then Exit Sub        ' also sets Sketchbook_Path variable  02.12.21: Juergen
  
  ChDrive Sketchbook_Path
  ChDir Sketchbook_Path
  Dim Res As String, DirList As String
  Res = Dir("libraries\Arduino_*.", vbDirectory)
  While Res <> ""
    DirList = DirList & Res & vbTab
    Res = Dir()
  Wend
  Dim d As Variant
  For Each d In Split(DirList, vbTab)
     Correct_one_Temp_Arduino_nr_Dir (d)
  Next
  Exit Sub
End Sub

' 02.12.21: Juergen see forum post #7085
'------------------------------------------
Public Function CheckArduinoHomeDir()
'------------------------------------------
  CheckArduinoHomeDir = False
  If Read_Sketchbook_Path_from_preferences_txt() = False Then Exit Function
  On Error GoTo DirError
  ChDrive Sketchbook_Path
  ChDir Sketchbook_Path
  CheckArduinoHomeDir = True
  Exit Function
DirError:
  On Error GoTo 0
  Dim Message
  Message = Replace(Get_Language_Str("Das Arduino Sketchbook Verzeichnis #1# existiert nicht." & _
    "Bitte prüfen und korrigieren sie die Einstellungen in der Arduino IDE."), "#1#", _
    Sketchbook_Path & vbCrLf)
      
  MsgBox Message, vbCritical, Get_Language_Str("Es sind Fehler aufgetreten")
End Function

'-----------------------------------------------------------------------------------
Private Function Check_All_Selected_Libraries_Result(Ask_User As Boolean) As Boolean
'-----------------------------------------------------------------------------------
' Return true if the update should be repeated
  Dim Row As Long, Sh As Worksheet, NotInstCnt As Long, List As String
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  Row = First_Dat_Row
  While Sh.Cells(Row, Libr_Name_Col) <> ""
        If Sh.Cells(Row, SelectRow_Col) <> "" Then
           If Sh.Cells(Row, Installed_Col) <> 1 Then
              NotInstCnt = NotInstCnt + 1
              List = List & "   " & Sh.Cells(Row, Libr_Name_Col) & vbCr   ' 02.06.20:
           End If
        End If
        Row = Row + 1
  Wend

  If NotInstCnt > 0 Then
     If Ask_User Then
            If MsgBox(Get_Language_Str("Fehler beim Aktualisieren der Bibliotheken und Boards aufgetreten. " & vbCr & _
                                       "Leider treten beim herunter laden vom Arduino Server manchmal Übertragungsfehler auf. " & vbCr & _
                                       "Oft hilft es wenn man den Prozess noch mal startet." & vbCr & _
                                       vbCr & _
                                       "Nicht installiert:") & vbCr & List & vbCr & _
                                       Get_Language_Str("Soll die Aktualisierung noch mal aufgerufen werden?"), vbQuestion + vbYesNo, _
                                       Get_Language_Str("Es sind Fehler beim Aktualisieren aufgetreten")) = vbYes Then
               Check_All_Selected_Libraries_Result = True
            End If
      Else: Check_All_Selected_Libraries_Result = True
      End If
  End If
End Function

'UT---------------------------------------------------
Private Sub Test_Check_All_Selected_Libraries_Result()
'UT---------------------------------------------------
  Check_All_Selected_Libraries_Result True
End Sub

'---------------------------------------------------
Private Sub Update_Status(Optional Start As Boolean)
'---------------------------------------------------
' Is called by OnTime
  If Update_Time <> 0 Or Start Then
     If Start Then
           Update_Time = Time
     Else: StatusMsg_UserForm.Set_ActSheet_Label Format(Time - Update_Time, "hh:mm:ss")
     End If
     Application.OnTime Now + TimeValue("00:00:01"), "Update_Status"
  End If
End Sub

'--------------------------------
Private Sub Stop_Status_Display()
'--------------------------------
  Update_Time = 0
  Unload StatusMsg_UserForm
End Sub



'----------------------------------------------------------
Private Function Update_All_Selected_Libraries() As Boolean
'----------------------------------------------------------
  Dim Pause_at_End As Boolean, Trials As Long, Ask_User As Boolean, Start_Update As Boolean
  If Read_Sketchbook_Path_from_preferences_txt() = False Then GoTo EndFunc
  Start_Update = True
  
  #If VBA7 Then                                                             ' 05.06.20:
    Dim hWnd As LongPtr: hWnd = Application.hWnd
  #Else
    Dim hWnd As Long:    hWnd = Application.hWnd
  #End If
  
  Do
    UnzipList = ""
    Select Case Create_Do_Update_Script(Pause_at_End)
      Case 0:  MsgBox Get_Language_Str("Es wurden keine Zeilen zur Installation ausgewählt. Die Zeilen müssen mit einem Häkchen in der Spalte 'Select' markiert werden." & vbCr & _
                                       "Für die ausgewählten Zeilen wird die neueste Software installiert, es sei den in der Spalte ""Required Version"" ist eine " & _
                                       "bestimmte Version angegeben."), vbInformation, _
                      Get_Language_Str("Keine Zeilen zur Installation ausgewählt.")
               GoTo EndFunc
      Case -1: GoTo EndFunc
    End Select
  
    StatusMsg_UserForm.ShowDialog Get_Language_Str("Aktualisiere Bibliotheken und Boards"), ""
    Update_Status Start_Update
    Start_Update = False
    Correct_Temp_Adrduino_nr_Dirs
    
    ChDrive Sketchbook_Path
    ChDir Sketchbook_Path
    If Dir("libraries\*.*", vbDirectory) = "" Then MkDir "libraries\"
    ChDir Sketchbook_Path & "\libraries\"
    
    Dim CommandStr As String, Res As Long
    
    CommandStr = ThisWorkbook.Path & "\" & UPDATE_LIB_CMD_NAME
    Res = ShellAndWait(CommandStr, 0, vbNormalFocus, PromptUser) ' No timeout to be able to study the results in case of an error
    Select Case Res
      Case Success, Timeout:
      Case Else:             ' No additional error message. They have been shown in the DOS box
                             MsgBox Replace(Replace(Get_Language_Str("Fehler #1# beim Starten der Update Programms '#2#'"), "#1#", Res), "#2#", CommandStr), vbCritical, _
                                    Get_Language_Str("Fehler beim Aktualisieren der Bibliotheken")
                             GoTo EndFunc
    End Select
    
    If WIN7_COMPATIBLE_DOWNLOAD Then
        Proc_UnzipList                                                      ' 20.06.20:
    End If
    
    Unload StatusMsg_UserForm
    
    ' Bring Excel to the top
    ' Is not working if an other application has be moved above Excel with Alt+Tab
    ' But this is a feature of Windows.
    '   See: https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow
    ' But it brings up excel again after the upload to the Arduino
    ' Without this funchion an other program was activated after the upload for some reasons
    Bring_to_front hWnd
    
    DoEvents
    
    If Get_All_Library_States() = False Then GoTo EndFunc
    
    Trials = Trials + 1
    If Trials >= 2 Then
       Ask_User = True
       Pause_at_End = True
    End If
    Update_General_Versions
    
  Loop While Check_All_Selected_Libraries_Result(Ask_User)
  
  Update_All_Selected_Libraries = True
  
EndFunc:
  Stop_Status_Display
  Unload StatusMsg_UserForm
  ChDrive ThisWorkbook.Path
  ChDir ThisWorkbook.Path
End Function

'----------------------------------------
Private Function Select_Missing() As Long
'----------------------------------------
  Dim Row As Long, Sh As Worksheet, NotInstCnt As Long, FastLED_Ver As String, Arduino_Ver As String, Arduino_row As Long
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  Row = First_Dat_Row
  While Sh.Cells(Row, Libr_Name_Col) <> ""
     With Sh.Cells(Row, SelectRow_Col)
        .Value = ""
        If InStr(Sh.Cells(Row, Lib_Board_Col), "*") = 0 Then ' An '*' is used to mark rows which could only be installed manually
            If Sh.Cells(Row, Installed_Col) <> 1 Then
                  .Value = ChrW(Hook_CHAR)
                  NotInstCnt = NotInstCnt + 1
            ElseIf Sh.Cells(Row, Reque_Ver_Col) <> "" Then
                     If VersionStr_is_Greater(Sh.Cells(Row, Reque_Ver_Col), Sh.Cells(Row, DetectVer_Col)) Then
                        .Value = ChrW(Hook_CHAR)
                        NotInstCnt = NotInstCnt + 1
                     End If
            End If
        End If
        Select Case Sh.Cells(Row, Libr_Name_Col)
           Case "FastLED": FastLED_Ver = Sh.Cells(Row, DetectVer_Col)
           Case "arduino:avr": Arduino_Ver = Sh.Cells(Row, DetectVer_Col)
                Arduino_row = Row
        End Select
        Row = Row + 1
     End With
  Wend
  
  ' Special Check: FastLED >= 3.3.3 require GCC > 7.3.0  => arduino lib >= 1.6.23
  If Sh.Cells(Arduino_row, SelectRow_Col) <> ChrW(Hook_CHAR) Then
     If VersionStr_is_Greater(FastLED_Ver, "3.3.2") And Not VersionStr_is_Greater(Arduino_Ver, "1.6.23") Then
        Sh.Cells(Arduino_row, SelectRow_Col) = ChrW(Hook_CHAR)
        If Not VersionStr_is_Greater(Sh.Cells(Arduino_row, Reque_Ver_Col), "1.6.23") Then Sh.Cells(Arduino_row, Reque_Ver_Col) = ""
        NotInstCnt = NotInstCnt + 1
     End If
  End If
  
  
  Select_Missing = NotInstCnt
End Function

'----------------------------------------------
Private Function Create_Restart_Cmd() As String                             ' 30.05.20:
'----------------------------------------------
' Create a CMD file which restarts the new version of the Prog_Generator
' - Wait until the existing prog generator is closes
' - Restart excel
  If Read_Sketchbook_Path_from_preferences_txt() = False Then Exit Function
  
  Dim fp As Integer, Name As String, UpdCnt As Long
  fp = FreeFile
  Name = ThisWorkbook.Path & "\" & RESTART_PROGGEN_CMD
  On Error GoTo WriteError
  Open Name For Output As #fp
  Print #fp, "@ECHO OFF"
  Print #fp, "Color 79" ' Blue  on bright Gray background (See: https://ss64.com/nt/color.html)
  Print #fp, "REM This file was generated by '" & ThisWorkbook.Name & "'  " & Time
  Print #fp, "REM"
  Print #fp, "Rem Wait until the Prog_Generator_MobaLedLib is closed"
  Print #fp, "REM and restart the new version of the Prog_Generator_MobaLedLib"
  Print #fp, "REM"
  Print #fp, ""
  Print #fp, "ECHO  ~~~~~~~~~~~~~~~~~~"
  Print #fp, "ECHO  Update is finished"
  Print #fp, "ECHO  ~~~~~~~~~~~~~~~~~~"
  Print #fp, "ECHO."
  Print #fp, "ECHO  Going to restarting the new Prog_Generator_MobaLedLib.xlsm"
  Print #fp, "ECHO."
  Print #fp, "ECHO  If the program hangs here the hidden file ""~$Prog_Generator_MobaLedLib.xlsm"""
  Print #fp, "ECHO  is not deleted for some reasons. It has to be deleted manualy."
  Print #fp, "ECHO."
  Print #fp, "ECHO  Make sure that all excel instances are closed if it hangs."
  Print #fp, "ECHO  In case of problems the installation is continued in one minute."
  Print #fp, "ECHO."
  Print #fp, "set /A counter=1"                                             ' 08.10.20: New Block
  Print #fp, "::define a variable containing a single backspace character"
  Print #fp, "for /f %%A in ('""prompt $H &echo on &for %%B in (1) do rem""') do set BS=%%A"
  Print #fp, "echo | set /p=%BS% Waiting until excel is closed" ' ECHO without new line
  Print #fp, ": Wait"
  Print #fp, "@ping localhost -n 3 > NUL" ' Wait 3 seconds
  Print #fp, "echo | set /p=."
  Print #fp, "set /A counter=%counter%+1"                                   ' 08.10.20:
  Print #fp, "if %counter% gtr 20 ( goto :Continue )"                       ' 08.10.20:
  Print #fp, "if exist ""~$Prog_Generator_MobaLedLib.xlsm"" Goto Wait"
  Print #fp, ":Continue"                                                    ' 08.10.20:
  Print #fp, "ECHO."
  Print #fp, "ECHO  Going to start the Prog_Generator_MobaLedLib again"
  Print #fp, "CHCP 65001 > NUL" ' Change the code Page to be able to use special characters like "ä" in the path
  Print #fp, left(Sketchbook_Path, 2)
  Print #fp, "CD """ & ConvertToUTF8Str(Sketchbook_Path) & "\libraries\MobaLedLib\extras\""" ' 13.11.21: Juergen fix issue #6894
'  Print #fp, "CD"    ' Debug
'  Print #fp, "PAUSE" ' Debug
  Print #fp, "@ping localhost -n 1 > NUL" ' Wait 1 second to be shure that excel is closed
  ' 09.03.21 Juergen: delete cache file to force an ESP32 rebuild, otherwise prebuild library versions would still be used
  Print #fp, "@if exist ""%USERPROFILE%\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache"" del ""%USERPROFILE%\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache"""
  Print #fp, "Start Prog_Generator_MobaLedLib.xlsm"
  Print #fp, "EXIT"
  Close #fp
  Create_Restart_Cmd = GetShortPath(ThisWorkbook.Path) & "\" & RESTART_PROGGEN_CMD        ' 13.11.21: Juergen fix issue #6894
  Exit Function

WriteError:
  Close #fp
  MsgBox Get_Language_Str("Fehler beim schreiben der Datei '") & Name & "'", vbCritical, Get_Language_Str("Fehler beim erzeugen der Arduino Start Datei")
End Function


'----------------------------------------------------------------
Private Function Select_from_Range(RangeStr As String) As Integer
'----------------------------------------------------------------
  Dim Row As Long, Sh As Worksheet
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  Row = First_Dat_Row
  While Sh.Cells(Row, Libr_Name_Col) <> ""
     With Sh.Cells(Row, SelectRow_Col)
        .Value = ""
        Row = Row + 1
     End With
  Wend
  On Error GoTo Range_Not_Found
  Sh.Range(RangeStr) = ChrW(Hook_CHAR)
  On Error GoTo 0
  Select_from_Range = Sh.Range(RangeStr).Row
  Exit Function
  
Range_Not_Found:
  MsgBox Replace(Replace(Get_Language_Str("Fehler: Der Bereich '#1#' wurde nicht im Blatt '#2#' gefunden"), "#1#", RangeStr), "#2#", Sh.Name), _
         vbCritical, Get_Language_Str("Fehler beim aktivieren der Update Zeile")
 Select_from_Range = -1
End Function


'--------------------------------------------------------------------
Private Function Show_Close_Message_if_Other_WB_are_Open() As Boolean
'--------------------------------------------------------------------
  Dim wb As Variant
  For Each wb In Workbooks
    If wb.Name <> ThisWorkbook.Name Then
       Close_Other_Workbooks.Start "Start_Update_MobaLedLib_and_Restarte_Excel"
       Show_Close_Message_if_Other_WB_are_Open = True
       Exit Function
    End If
  Next
End Function

'-------------------------------------------------------
Private Sub Start_Update_MobaLedLib_and_Restarte_Excel()
'-------------------------------------------------------
  ' Close all other workbooks without saving (The user has been warned before)
  Dim wb As Variant
  For Each wb In Workbooks
    If wb.Name <> ThisWorkbook.Name Then
       wb.Close Savechanges:=False
    End If
  Next
  
  If Update_All_Selected_Libraries() = False Then Exit Sub
  Dim CommandStr As String
  CommandStr = Create_Restart_Cmd()
  If CommandStr = "" Then Exit Sub
  ThisWorkbook.Save
  Shell "cmd /c start " & CommandStr
'  MsgBox "Warte"
  Application.Quit
End Sub


'-----------------------------------------------------------------------------
Private Sub Update_MobaLedLib_from_Range_and_Restart_Excel(RangeStr As String)
'-----------------------------------------------------------------------------
  Dim Row As Integer
  Row = Select_from_Range(RangeStr)
  If Row < 0 Then Exit Sub
  Dim Ctrl_Pressed As Boolean
  Ctrl_Pressed = GetAsyncKeyState(VK_CONTROL) <> 0  ' Following function must be declared: Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
  If Ctrl_Pressed Then
    Dim frm As New UserForm_SingleInput, currentUrl As String, newUrl As String
    Dim URL As String
    currentUrl = ThisWorkbook.Sheets(LIBRARYS__SH).Cells(Row, Other_Src_Col)
    newUrl = frm.ShowForm(Get_Language_Str("Beta-Test Installation"), Get_Language_Str("Bitte geben sie die URL ein, von der sie die neue Beta Version herunterladen wollen"), currentUrl)
    If newUrl = "<Abort>" Then Exit Sub
    ThisWorkbook.Sheets(LIBRARYS__SH).Cells(Row, Other_Src_Col) = newUrl
  End If

  If Show_Close_Message_if_Other_WB_are_Open Then Exit Sub ' Other Workbooks are opened => "Start_Update_MobaLedLib_and_Restarte_Excel" is called after they ara closed
  Start_Update_MobaLedLib_and_Restarte_Excel
End Sub


'---------------------------
Public Sub Delete_Selected()
'---------------------------
  If Read_Sketchbook_Path_from_preferences_txt() = False Then Exit Sub
  Dim Row As Long, Sh As Worksheet, DidDelete As Boolean
  DidDelete = False
  Set Sh = ThisWorkbook.Sheets(LIBRARYS__SH)
  Row = First_Dat_Row
  While Sh.Cells(Row, Libr_Name_Col) <> ""
        If Sh.Cells(Row, SelectRow_Col) <> "" Then
           If InStr(Sh.Cells(Row, Lib_Board_Col), "L") > 0 Then
                ' *** Library ***
                Dim LibrariesDir As String, LibDir As String
                LibrariesDir = Sketchbook_Path & "\libraries\"
                LibDir = LibrariesDir & Sh.Cells(Row, Libr_Name_Col)
                If Dir(LibDir, vbDirectory) <> "" Then
                   Debug.Print "Deleting: " & LibDir
                   Del_Folder LibDir
                   DidDelete = True
                End If
           ElseIf InStr(Sh.Cells(Row, Lib_Board_Col), "B") > 0 Then
                ' Board
                Dim BoardDir As String, Board_and_Proc As String, Board As String
                Board_and_Proc = Sh.Cells(Row, Libr_Name_Col)
                Board = Split(Board_and_Proc, ":")(0)
                BoardDir = Environ(Env_USERPROFILE) & AppLoc_Ardu & "packages\" & Board
                If Dir(BoardDir, vbDirectory) <> "" Then
                   Debug.Print "Deleting: " & BoardDir
                   Del_Folder BoardDir
                   DidDelete = True
                End If
           End If
        End If
        Row = Row + 1
  Wend
  If DidDelete = True Then                  ' 11.03.21 Juergen: force an ESP32 rebuild
    If Dir(Environ(Env_USERPROFILE) & "\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache") <> "" Then
       Kill Environ(Env_USERPROFILE) & "\AppData\Local\Temp\MobaLedLib_build\ESP32\includes.cache"
    End If
  End If

  Debug.Print "Waiting"
  Dim i As Long
  For i = 1 To 30
      DoEvents
      Debug.Print ".";
      Sleep 100 ' Wait for windows to update the directory structure otherwise we get an error in the following "Dir()" call
  Next
  Debug.Print ""
  Get_All_Library_States
End Sub

'------------------------------------------------------------
Public Sub Update_MobaLedLib_from_Arduino_and_Restart_Excel()
'------------------------------------------------------------
  If MsgBox(Get_Language_Str("Soll die MobaLedLib aktualisiert werden?" & vbCr & _
                             "Wenn die vorhandene Bibliothek die gleiche oder eine neuere Version besitzt, dann " & _
                             "wird die existierend Bibliothek beibehalten."), _
                             vbQuestion + vbYesNo, Get_Language_Str("Aktualisieren der MobaLedLib")) <> vbYes Then Exit Sub
  Update_MobaLedLib_from_Range_and_Restart_Excel "Select_MobaLedLib_Arduino"
End Sub

'---------------------------------------------------------
Public Sub Update_MobaLedLib_from_Beta_and_Restart_Excel()
'---------------------------------------------------------
  If MsgBox(Get_Language_Str("Soll die Beta Test Version der MobaLedLib installiert werden?"), vbQuestion + vbYesNo, _
     Get_Language_Str("Beta Test der MobaLedLib installieren?")) <> vbYes Then Exit Sub
  Update_MobaLedLib_from_Range_and_Restart_Excel "Select_MobaLedLib_Beta"
End Sub


'---------------------------------
Public Sub Check_Actual_Versions()
'---------------------------------
' Is called by the button in the "Libraries" sheet
' It checks all versions and selects the rows which have to be updated
  Update_General_Versions
  Get_All_Library_States
  Select_Missing
End Sub



'----------------------------
Public Sub Install_Selected()
'----------------------------
' Is called by the button in the "Libraries" sheet
  Update_All_Selected_Libraries
End Sub

'-----------------------------------------------
Public Sub Install_Missing_Libraries_and_Board()
'-----------------------------------------------
  StatusMsg_UserForm.ShowDialog Get_Language_Str("Überprüfe Bibliotheken und Boards"), ""
  Update_General_Versions
  Get_All_Library_States
  If Select_Missing() > 0 Then
     Install_Selected
  End If
  Unload StatusMsg_UserForm
End Sub

'------------------------------
Public Sub OpenSketchbookPath()
'------------------------------
  Read_Sketchbook_Path_from_preferences_txt
  Dim Name As String
  Name = Sketchbook_Path
  Shell "Explorer /root,""" & Name & """", vbNormalFocus
End Sub


'-------------------------------------------------------------
Public Function Is_Lib_Installed(LibName As String) As Boolean              ' 11.11.20:
'-------------------------------------------------------------
  Dim Row As Long
  Dim LastRow As Long
  LastRow = LastUsedRowIn(LIBRARYS__SH)
  Row = First_Dat_Row
  With Sheets(LIBRARYS__SH)
     While Row <= LastRow                                                   ' 06.12.2021 Juergen Fix issue with empty lines in sheet
        If .Cells(Row, Libr_Name_Col) = LibName Then
            Is_Lib_Installed = (.Cells(Row, Installed_Col) = 1)
            Exit Function
        End If
        Row = Row + 1
     Wend
  End With
End Function

'-------------------------------------------------------------
Public Function Get_Lib_Version(LibName As String) As String              ' 01.03.21:
'-------------------------------------------------------------
  Dim Row As Long
  Row = First_Dat_Row
  Dim LastRow As Long
  LastRow = LastUsedRowIn(LIBRARYS__SH)
  With Sheets(LIBRARYS__SH)
     While Row <= LastRow                                                   ' 06.12.2021 Juergen Fix issue with empty lines in sheet
        If .Cells(Row, Libr_Name_Col) = LibName Then
            Get_Lib_Version = .Cells(Row, DetectVer_Col)
            Exit Function
        End If
        Row = Row + 1
     Wend
  End With
  Get_Lib_Version = ""
End Function

'-------------------------------------------------------------
Public Function Get_Required_Version(LibName As String) As String              ' 01.03.21:
'-------------------------------------------------------------
  Dim Row As Long
  Row = First_Dat_Row
  Dim LastRow As Long
  LastRow = LastUsedRowIn(LIBRARYS__SH)
  With Sheets(LIBRARYS__SH)
     While Row <= LastRow                                                   ' 06.12.2021 Juergen Fix issue with empty lines in sheet
        If .Cells(Row, Libr_Name_Col) = LibName Then
            Get_Required_Version = .Cells(Row, Reque_Ver_Col)
            Exit Function
        End If
        Row = Row + 1
     Wend
  End With
  Get_Required_Version = ""
End Function

'UT--------------------------------
Private Sub Test_Is_Lib_Installed()
'UT--------------------------------
  Debug.Print "Is_Lib_Installed(esp32:esp32): " & Is_Lib_Installed("esp32:esp32")
  Debug.Print "Is_Lib_Installed(NichtInstal): " & Is_Lib_Installed("NichtInstal")
  Debug.Print "Get_Lib_Version(esp32:tools\esptool_py)" & Get_Lib_Version("esp32:tools\esptool_py")
  Debug.Print "Get_Lib_Version(NichtInstal)" & Get_Lib_Version("NichtInstal")
End Sub

'-----------------------------------------------
Public Function ESP32_Lib_Installed() As Boolean                            ' 11.11.20:
'-----------------------------------------------
  ESP32_Lib_Installed = Is_Lib_Installed("esp32:esp32")
End Function

Public Function PICO_Lib_Installed() As Boolean                             ' 18.04.21: Juergen
'-----------------------------------------------
  PICO_Lib_Installed = Is_Lib_Installed("rp2040:rp2040")
End Function



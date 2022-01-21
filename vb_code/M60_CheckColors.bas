Attribute VB_Name = "M60_CheckColors"
Option Explicit

' Support for the Phython modul MobaLedCheckColors.py from Harold


' ToDo:
' - Verwendet die Datei: LEDs_AutoProg\MobaLedTest_config.json
' - Keine py Dateien unterstützen
' - Download testen
' - Wenn bereits eine ColTab Zeile vorhanden ist, dann sollen die Farben verwendet werden

#Const USE_pyPROGGEN = 2                                                    ' 06.06.20:

#If USE_pyPROGGEN = 2 Then ' New method which uses a ZIP file which has to be extracted           ' 27.01.21:
  Private Const CHECKCOL_EXE_DIR = Ino_Dir_LED & "pyProg_Generator_MobaLedLib\"
  Private Const CHECKCOL_DAT_DIR = Ino_Dir_LED
  Private Const CHECKCOL_ZIP_DIR = Ino_Dir_LED
  Private Const DOWNLOAD_EXEPROG = "https://github.com/haroldlinke/MobaLedLib_pyProgGen/blob/master/pyProg_Generator_MobaLedLib.zip?raw=true" ' 31.01.21:
 'Private Const DOWNLOAD_EXEPROG = "http://www.hlinke.de/files/pyProg_Generator_MobaLedLib.zip"
  Private Const CHECK_COLORS_EXE = "pyProg_Generator_MobaLedLib.exe"
  Private Const CHECK_COLORS_DST = "pyProg_Generator_MobaLedLib.zip"
  Private Const PYCMDLINE_PARAMS = " --startpage ColorCheckPage"
#ElseIf USE_pyPROGGEN = 1 Then
  Private Const CHECKCOL_EXE_DIR = Ino_Dir_LED & "pyProgGen_MobaLedLib\"
  Private Const CHECKCOL_DAT_DIR = Ino_Dir_LED
  Private Const DOWNLOAD_EXEPROG = "https://github.com/haroldlinke/MobaLedLib_pyProgGen/blob/master/pyProg_Generator_MobaLedLib.exe?raw=true" ' 05.06.20:
  Private Const CHECK_COLORS_EXE = "pyProg_Generator_MobaLedLib.exe"
  Private Const CHECK_COLORS_DST = CHECK_COLORS_EXE
  Private Const PYCMDLINE_PARAMS = " --startpage ColorCheckPage"
#Else
  Private Const CHECKCOL_EXE_DIR = Ino_Dir_LED & "CheckColors\"
  Private Const CHECKCOL_DAT_DIR = Ino_Dir_LED & "CheckColors\"
  Private Const DOWNLOAD_EXEPROG = "https://github.com/Hardi-St/MobaLedLib_Docu/blob/master/Tools/CheckColors/MobaLedCheckColors.exe?raw=true"
  Private Const CHECK_COLORS_EXE = "MobaLedCheckColors.exe"
  Private Const CHECK_COLORS_DST = CHECK_COLORS_EXE
  Private Const PYCMDLINE_PARAMS = ""
#End If

Private Const CONFIG_FILE_NAME = "MobaLedTest_config.json"
Private Const DISCONECTED_NAME = "MobaLedTest_disconnect.Txt"
Private Const CLOSE_CHECKCOL_N = "MobaLedTest_Close.Txt"
Private Const COLTEST_ONLYFILE = "ColorTestOnly.txt"

Private Const START_PALETTETXT = "    ""palette"": {"
Private Const ENDE_PALETTE_TXT = "    },"

Private Const START_SERPORT_NR = "    ""serportnumber"":"
Private Const START_SERPO_NAME = "    ""serportname"": """



Private Const EXP_PYTHON_V_STR = "3.7.0"  ' Old 3.8.0
Private Const PYTHON_DOWNLOADP = "https://www.python.org/downloads/"

Private Const FINISHEDTXT_FILE = CHECKCOL_EXE_DIR & "Finished.txt"

Private Const COLTAB_SIZE = 17

Private Const NamesList = "ROOM_COL0," & _
              "ROOM_COL1," & _
              "ROOM_COL2," & _
              "ROOM_COL3," & _
              "ROOM_COL4," & _
              "ROOM_COL5," & _
              "GAS_LIGHT D," & _
              "GAS LIGHT," & _
              "NEON_LIGHT D," & _
              "NEON_LIGHT M," & _
              "NEON_LIGHT," & _
              "ROOM_TV0 A," & _
              "ROOM_TV0 B," & _
              "ROOM_TV1 A," & _
              "ROOM_TV1 B," & _
              "SINGLE_LED," & _
              "SINGLE_LED D"


Private Type RGB_T
  r As Integer
  g As Integer
  b As Integer
End Type

Private PythonVer As Long

Private Smiley_Direction As Long
Private Smiley_Cnt As Long

Private Proc_CheckColors_Form_Callback As String
Private ColTab_Dest_Sheet As String
Private ColTab_Dest_Row As Long


'-----------------------------------------------------
Private Function VerStr_to_long(Str As String) As Long
'-----------------------------------------------------
  Dim Parts() As String
  Parts = Split(Str, ".")
  VerStr_to_long = Parts(0) * 1000000
  If UBound(Parts) >= 1 Then VerStr_to_long = VerStr_to_long + Parts(1) * 1000
  If UBound(Parts) >= 2 Then VerStr_to_long = VerStr_to_long + Parts(2)
End Function


'--------------------------------------------------------------------------------------
Private Function Check_Phyton(ExpVerStr As String, ByRef ExistingVer As String) As Long
'--------------------------------------------------------------------------------------
' Return a positiv number if the detected version is >= the expectet version
'        a negative number if the detected version is < the expectet version
'        The number is the actual version
'        0 if python is not installed
  Const PythonStr = "Python "
  Dim Res As String, ExpVer As Long, ActVer As Long
  ExpVer = VerStr_to_long(ExpVerStr)
  Res = F_shellExec("cmd /c Python -V")
    
  If Res <> "" Then
     If Left(Res, Len(PythonStr)) = PythonStr Then
        ExistingVer = Trim(Replace(Replace(Mid(Res, Len(PythonStr)), vbLf, ""), vbCr, ""))
        ActVer = VerStr_to_long(ExistingVer)
        If ActVer >= ExpVer Then
              Check_Phyton = ActVer
        Else: Check_Phyton = -ActVer
        End If
     End If
  End If
  
End Function

'UT----------------------------
Private Sub Test_Check_Phyton()
'UT----------------------------
  Dim ExistingVer As String
  Debug.Print Check_Phyton("3.8.0", ExistingVer) & " ExistingVer='" & ExistingVer & "'"
End Sub

'--------------------------------------------------------
Private Function Start_MobaLedCheckColors_py() As Boolean
'--------------------------------------------------------
  Dim DstDir As String, OldDir As String
  OldDir = CurDir
  DstDir = ThisWorkbook.Path & "\" & CHECKCOL_EXE_DIR
  On Error GoTo DirError
  ChDrive DstDir
  ChDir DstDir
  On Error GoTo 0
  
  Debug.Print Shell("cmd /c start /min python MobaLedCheckColors.py") ' Program is started in background because otherwise Excel "hangs"
  'Debug.Print Shell("cmd /c start /min Start_py.cmd") ' Program is started in background because otherwise Excel "hangs"
  
  ChDrive OldDir
  ChDir OldDir
  Start_MobaLedCheckColors_py = True
  Exit Function
  
DirError:
  MsgBox Get_Language_Str("Fehler beim Wechsel in das Verzeichnis:") & vbCr & _
         "  '" & DstDir & "'", vbCritical, Get_Language_Str("Fehler beim Start der Farbauswahl")
End Function

'---------------------------------------------------------
Private Function Start_MobaLedCheckColors_exe() As Boolean
'---------------------------------------------------------
  Dim DstDir As String, OldDir As String
  OldDir = CurDir
  DstDir = ThisWorkbook.Path & "\" & CHECKCOL_EXE_DIR
  On Error GoTo DirError
  ChDrive DstDir
  ChDir DstDir
  On Error GoTo 0
  
  Debug.Print Shell("cmd /c start " & CHECK_COLORS_EXE & PYCMDLINE_PARAMS) ' Program is started in background because otherwise Excel "hangs"
  
  ChDrive OldDir
  ChDir OldDir
  Start_MobaLedCheckColors_exe = True
  Exit Function
  
DirError:
  MsgBox Get_Language_Str("Fehler beim Wechsel in das Verzeichnis:") & vbCr & _
         "  '" & DstDir & "'", vbCritical, Get_Language_Str("Fehler beim Start der Farbauswahl")
End Function

#If 0 Then ' 18.01.20: Not Used
'----------------------------------
Public Sub Disconnect_CheckColors()
'----------------------------------
  Dim fp As Integer, Name As String
  Name = ThisWorkbook.Path & "\" & CHECKCOL_DAT_DIR & DISCONECTED_NAME
  fp = FreeFile
  On Error GoTo WriteError
  Open Name For Output As #fp
  Close #fp
  On Error GoTo 0
  Exit Sub
  
WriteError:
  MsgBox Get_Language_Str("Fehler beim erzeugen der Disconnect Datei:") & vbCr & _
         "  '" & Name & "'", vbCritical, Get_Language_Str("Fehler beim trennen der Verbindung zum Arduino")
End Sub
#End If

'-----------------------------
Public Sub Close_CheckColors()
'-----------------------------
  Dim fp As Integer, Name As String
  Name = ThisWorkbook.Path & "\" & CHECKCOL_EXE_DIR & CLOSE_CHECKCOL_N   ' 18.01.21: Old: CHECKCOL_DAT_DIR
  fp = FreeFile
  On Error GoTo WriteError
  Open Name For Output As #fp
  Close #fp
  On Error GoTo 0
  Exit Sub
  
WriteError:
  MsgBox Get_Language_Str("Fehler beim erzeugen der Close Datei:") & vbCr & _
         "  '" & Name & "'", vbCritical, Get_Language_Str("Fehler beim beenden des Farbtest Programms")
End Sub

'-----------------------------------------
Private Sub Delete_CheckColors_CloseFile()
'-----------------------------------------
  Dim Name As String
  Name = ThisWorkbook.Path & "\" & CHECKCOL_EXE_DIR & CLOSE_CHECKCOL_N  ' 18.01.21: Old: CHECKCOL_DAT_DIR
  If Dir(Name) <> "" Then Kill Name
End Sub


'-----------------------------------
Public Sub Write_ColTest_Only_File()                                        ' 18.01.21:
'-----------------------------------
  Dim fp As Integer, Name As String
  Name = ThisWorkbook.Path & "\" & CHECKCOL_EXE_DIR & COLTEST_ONLYFILE
  fp = FreeFile
  On Error GoTo WriteError
  Open Name For Output As #fp
  Close #fp
  On Error GoTo 0
  Exit Sub
  
WriteError:
  MsgBox Get_Language_Str("Fehler beim erzeugen der Datei:") & vbCr & _
         "  '" & Name & "'", vbCritical, Get_Language_Str("Fehler beim anlegen einer Datei")
End Sub

'-------------------------------------
Private Sub Delete_ColTest_Only_File()                                      ' 18.01.21:
'-------------------------------------
  Dim Name As String
  Name = ThisWorkbook.Path & "\" & CHECKCOL_EXE_DIR & COLTEST_ONLYFILE
  If Dir(Name) <> "" Then Kill Name
End Sub


'------------------------------------------------------
Private Sub Set_Default_ColTab(ByRef ColTab() As RGB_T)
'------------------------------------------------------
  ColTab(0).r = 15:   ColTab(0).g = 13:   ColTab(0).b = 3:         ' 0  ROOM_COL0 (very dark warm White)
  ColTab(1).r = 22:   ColTab(1).g = 44:   ColTab(1).b = 27:        ' 1  ROOM_COL1 (cold dark White)
  ColTab(2).r = 155:  ColTab(2).g = 73:   ColTab(2).b = 5:         ' 2  ROOM_COL2 (warm Yellow)
  ColTab(3).r = 39:   ColTab(3).g = 18:   ColTab(3).b = 1:         ' 3  ROOM_COL345 (Dark Yellow)  randomly color 3: 4 or 5 is used
  ColTab(4).r = 30:   ColTab(4).g = 0:    ColTab(4).b = 0:         ' 4  ROOM_COL345 (Dark Red)
  ColTab(5).r = 79:   ColTab(5).g = 39:   ColTab(5).b = 7:         ' 5  ROOM_COL345 (Dark warm White)
  ColTab(6).r = 50:   ColTab(6).g = 50:   ColTab(6).b = 50:        ' 6  Gas light  dark    Bei einzeln adressierten Gas LEDs wird der individuelle Helligkeitswert verwendet (GAS_LIGHT1: GAS_LIGHT2: GAS_LIGHT3)
  ColTab(7).r = 255:  ColTab(7).g = 255:  ColTab(7).b = 255:       ' 7  Gas light  bright  Wenn 3 Kanaele zusammen verwendet werden dan bestimmt der erste Wert die Helligkeit das ist wichtig damit alle Ausgaenge gleich belastet werden (GAS_LIGHT und GAS_LIGHT)
  ColTab(8).r = 20:   ColTab(8).g = 20:   ColTab(8).b = 27:        ' 8  Neon light dark  (Achtung: Muss groesser als 2*MAX_FLICKER_CNT sein)
  ColTab(9).r = 70:   ColTab(9).g = 70:   ColTab(9).b = 80:        ' 9  Neon light medium
  ColTab(10).r = 245: ColTab(10).g = 245: ColTab(10).b = 255:      ' 10 Neon light bright
  ColTab(11).r = 50:  ColTab(11).g = 50:  ColTab(11).b = 20:       ' 11 TV0 and chimney color A randomly color A or B is used
  ColTab(12).r = 70:  ColTab(12).g = 70:  ColTab(12).b = 30:       ' 12 TV0 and chimney color B
  ColTab(13).r = 50:  ColTab(13).g = 50:  ColTab(13).b = 8:        ' 13 TV1 and chimney color A
  ColTab(14).r = 50:  ColTab(14).g = 50:  ColTab(14).b = 8:        ' 14 TV2 and chimney color B
  ColTab(15).r = 255: ColTab(15).g = 255:  ColTab(15).b = 255:     ' 15 Single LED Room:      Fuer einzeln adressierte LEDs wird der individuelle Helligkeitswert verwendet (SINGLE_LED1:  SINGLE_LED2:  SINGLE_LED3)  ' 06.09.19:
  ColTab(16).r = 50:  ColTab(16).g = 50:  ColTab(16).b = 50:       ' 16 Single dark LED Room: Fuer einzeln adressierte LEDs wird der individuelle Helligkeitswert verwendet (SINGLE_LED1D: SINGLE_LED2D: SINGLE_LED3D)
End Sub

'--------------------------------------------------
Private Function Dec_2_Hex2(d As Integer) As String
'--------------------------------------------------
  Dec_2_Hex2 = Right("00" & Hex(d), 2)
End Function

'--------------------------------------------------
Private Function RGB_to_Hex(rgb As RGB_T) As String
'--------------------------------------------------
  RGB_to_Hex = Dec_2_Hex2(rgb.r) & Dec_2_Hex2(rgb.g) & Dec_2_Hex2(rgb.b)
End Function

'---------------------------------------------------------
Private Sub Write_ColTab(fp As Integer, ColTab() As RGB_T)
'---------------------------------------------------------
  Print #fp, START_PALETTETXT
  Dim Name As Variant, NamesArray() As String, Nr As Integer
  NamesArray = Split(NamesList, ",")
  For Each Name In NamesArray
     Print #fp, "        """ & Name & """: ""#" & RGB_to_Hex(ColTab(Nr)) & """";
     If Nr < UBound(NamesArray) Then
           Print #fp, ","
     Else: Print #fp, ""
     End If
     Nr = Nr + 1
  Next
  Print #fp, ENDE_PALETTE_TXT

End Sub

'----------------------------------------------------
Public Sub Write_Default_CheckColors_Parameter_File()
'----------------------------------------------------
  Dim fp As Integer, FileName As String

  FileName = ThisWorkbook.Path & "\" & CHECKCOL_DAT_DIR & CONFIG_FILE_NAME
  fp = FreeFile
  On Error GoTo WriteError
  Open FileName For Output As #fp
  Print #fp, "{"
  Print #fp, "    ""serportnumber"": 0,"
  Print #fp, "    ""serportname"": """","
  Print #fp, "    ""maxLEDcount"": ""256"","
  Print #fp, "    ""lastLedCount"": 1,"
  Print #fp, "    ""lastLed"": 0,"
  Print #fp, "    ""pos_x"": 100,"
  Print #fp, "    ""pos_y"": 100,"
  Print #fp, "    ""colorview"": 1,"
  Print #fp, "    ""startpage"": 1,"
  Print #fp, "    ""led_correction_r"": ""100"","
  Print #fp, "    ""led_correction_g"": ""69"","
  Print #fp, "    ""led_correction_b"": ""94"","
  Print #fp, "    ""use_led_correction"": 1,"
  Print #fp, "    ""old_color"": ""#FFA8C7"","
  
  Dim ColTab(COLTAB_SIZE) As RGB_T
  Set_Default_ColTab ColTab
  Write_ColTab fp, ColTab
  
  Print #fp, "    ""autoconnect"": true"
  Print #fp, "}"
  Close #fp
  On Error GoTo 0
  Exit Sub
  
WriteError:
  MsgBox Get_Language_Str("Fehler beim Schreiben der Parameter Datei:") & vbCr & _
         "  '" & FileName & "'", vbCritical, Get_Language_Str("Fehler beim Schreiben der Parameter Datei")
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
Private Function Open_Cfg_File_and_Get_Sp_and_Ep(ByRef Txt As String, ByRef Sp As Long, ByRef Ep As Long, ByRef FileName As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
  FileName = ThisWorkbook.Path & "\" & CHECKCOL_DAT_DIR & CONFIG_FILE_NAME
  Txt = Read_File_to_String(FileName)
  If Txt = "#ERROR#" Then Exit Function
  
  Sp = InStr(Txt, START_PALETTETXT)
  If Sp = 0 Then
     MsgBox Get_Language_Str("Fehler: Der Text '") & START_PALETTETXT & Get_Language_Str("' existiert nicht in der Datei:") & vbCr & _
            "  '" & FileName & "'", vbCritical, Get_Language_Str("Fehler: Anfang der Farbpalette nicht gefunden")
     Exit Function
  End If
  
  Ep = InStr(Txt, ENDE_PALETTE_TXT)
  If Ep = 0 Then
     MsgBox Get_Language_Str("Fehler: Das Ende der Farbpalette wurde nicht gefunden in") & vbCr & _
            "  '" & FileName & "'", vbCritical, Get_Language_Str("Ende der Farbpalette nicht gefunden")
     Exit Function
  End If
  
  Open_Cfg_File_and_Get_Sp_and_Ep = True
End Function

'-------------------------------------------------------------------------
Private Function Insert_ColTab_to_ConfigFile(ColTab() As RGB_T) As Boolean
'-------------------------------------------------------------------------
  Dim FileName As String, Txt As String, Sp As Long, Ep As Long
  If Not Open_Cfg_File_and_Get_Sp_and_Ep(Txt, Sp, Ep, FileName) Then Exit Function
  
  Dim fp As Integer
  fp = FreeFile
  On Error GoTo WriteError
  Open FileName For Output As #fp
  Print #fp, Left(Txt, Sp - 1);
  
  Write_ColTab fp, ColTab
  
  Print #fp, Mid(Txt, Ep + Len(ENDE_PALETTE_TXT) + 2);
  
  Close #fp
  
  Insert_ColTab_to_ConfigFile = True
  Exit Function

WriteError:
  MsgBox Get_Language_Str("Fehler beim aktualisieren der Parameter Datei:") & vbCr & _
         "  '" & FileName & "'", vbCritical, Get_Language_Str("Fehler beim aktualisieren der Parameter Datei")
End Function




'----------------------------------------------------------------
Private Function Insert_Default_ColTab_to_ConfigFile() As Boolean
'----------------------------------------------------------------
  Dim ColTab(COLTAB_SIZE) As RGB_T
  Set_Default_ColTab ColTab
  Insert_Default_ColTab_to_ConfigFile = Insert_ColTab_to_ConfigFile(ColTab)
  #If False Then
    
  Dim FileName As String, Txt As String, Sp As Long, Ep As Long
  If Not Open_Cfg_File_and_Get_Sp_and_Ep(Txt, Sp, Ep, FileName) Then Exit Function
  
  Dim fp As Integer
  fp = FreeFile
  On Error GoTo WriteError
  Open FileName For Output As #fp
  Print #fp, Left(Txt, Sp - 1);
  
  Dim ColTab(COLTAB_SIZE) As RGB_T
  Set_Default_ColTab ColTab
  Write_ColTab fp, ColTab
  
  Print #fp, Mid(Txt, Ep + Len(ENDE_PALETTE_TXT) + 2);
  
  Close #fp
  
  Insert_Default_ColTab_to_ConfigFile = True
  Exit Function

WriteError:
  MsgBox Get_Language_Str("Fehler beim aktualisieren der Parameter Datei:") & vbCr & _
         "  '" & FileName & "'", vbCritical, Get_Language_Str("Fehler beim aktualisieren der Parameter Datei")
  #End If
End Function


'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Replace_in_String_from_To(ByRef Txt As String, ByVal FromTxt As String, ByVal ToTxt As String, ReplaceTxt As String, FileName As String) As Boolean
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
  Dim Sp, Ep As Long
  Sp = InStr(Txt, FromTxt)
  If Sp = 0 Then
     MsgBox Get_Language_Str("Fehler: Der Text '") & FromTxt & Get_Language_Str("' existiert nicht in der Datei:") & vbCr & _
            "  '" & FileName & "'", vbCritical, Get_Language_Str("Fehler beim aktualisieren der Datei")
     Exit Function
  End If
  
  Ep = InStr(Sp, Txt, ToTxt)
  If Ep = 0 Then
     MsgBox Get_Language_Str("Fehler: Der Text '") & ToTxt & Get_Language_Str("' existiert nicht in der Datei:") & vbCr & _
            "  '" & FileName & "'", vbCritical, Get_Language_Str("Fehler beim aktualisieren der Datei")
     Exit Function
  End If
  
  Txt = Left(Txt, Sp + Len(FromTxt) - 1) & ReplaceTxt & Mid(Txt, Ep)
  
  Replace_in_String_from_To = True
End Function


'-------------------------------------------------------------------------
Private Function Change_Comport_in_ConfigFile(ComNr As Integer) As Boolean
'-------------------------------------------------------------------------
  Dim FileName As String, Txt As String, fp As Integer
  FileName = ThisWorkbook.Path & "\" & CHECKCOL_DAT_DIR & CONFIG_FILE_NAME
  Txt = Read_File_to_String(FileName)
  If Txt = "#ERROR#" Then Exit Function
  
  If Not Replace_in_String_from_To(Txt, START_SERPORT_NR, ",", " " & ComNr, FileName) Then Exit Function
  If Not Replace_in_String_from_To(Txt, START_SERPO_NAME, """,", "COM" & ComNr, FileName) Then Exit Function
  
  fp = FreeFile
  On Error GoTo WriteError
  Open FileName For Output As #fp
  Print #fp, Txt;
  Close #fp
  
  Change_Comport_in_ConfigFile = True
  Exit Function

WriteError:
  MsgBox Get_Language_Str("Fehler beim aktualisieren des COM Ports in der Parameter Datei:") & vbCr & _
         "  '" & FileName & "'", vbCritical, Get_Language_Str("Fehler beim aktualisieren der Parameter Datei")
End Function

'UT--------------------------------------------
Private Sub Test_Change_Comport_in_ConfigFile()
'UT--------------------------------------------
  Change_Comport_in_ConfigFile 4
End Sub


'--------------------------------------------------------------------------------
Private Function Read_ColTab_from_Config_File(ByRef ColTab() As RGB_T) As Boolean
'--------------------------------------------------------------------------------
  Dim FileName As String, Txt As String, Sp As Long, Ep As Long, Nr As Integer
  If Not Open_Cfg_File_and_Get_Sp_and_Ep(Txt, Sp, Ep, FileName) Then Exit Function
  
  Dim ColTabList() As String, line As Variant
  Sp = Sp + Len(START_PALETTETXT) + 1
  ColTabList = Split(Mid(Txt, Sp, Ep - Sp), vbCr)
  For Each line In ColTabList
     Dim ColStr As String
     If Nr < COLTAB_SIZE Then
        ColStr = Replace(Replace(Trim(Split(line, ":")(1)), """", ""), ",", "")
        ColTab(Nr).r = "&H" & Mid(ColStr, 2, 2)
        ColTab(Nr).g = "&H" & Mid(ColStr, 4, 2)
        ColTab(Nr).b = "&H" & Mid(ColStr, 6, 2)
        Nr = Nr + 1
     End If
  Next line
End Function

'UT--------------------------------------------
Private Sub Test_Read_ColTab_from_Config_File()
'UT--------------------------------------------
  Dim ColTab(COLTAB_SIZE) As RGB_T
  Read_ColTab_from_Config_File ColTab
End Sub
  
'---------------------------------------------------------------------
Private Function ColTab_to_C_String(ByRef ColTab() As RGB_T) As String
'---------------------------------------------------------------------
' Generates this string:
'   // Set_ColTab(Red Green Blue)
'   Set_ColTab( 15,  13,   3, // *ROOM_COL0
'               22,  44,  27, //  ROOM_COL1
'              155,  73,   5, //  ROOM_COL2
'               39,  18,   1, // *ROOM_COL3
'               30,   0,   0, //  ROOM_COL4
'               79,  39,   7, //  ROOM_COL5
'               50,  50,  50, //  GAS_LIGHT D
'              255, 255, 255, //  GAS LIGHT
'               20,  20,  27, //  NEON_LIGHT D
'               70,  70,  80, //  NEON_LIGHT M
'              245, 245, 255, //  NEON_LIGHT
'               50,  50,  20, //  ROOM_TV0 A
'               70,  70,  30, //  ROOM_TV0 B
'               50,  50,   8, //  ROOM_TV1 A
'               50,  50,   8, //  ROOM_TV1 B
'              255, 255, 255, //  SINGLE_LED
'               50,  50,  50) //  SINGLE_LED D

  Dim DefColTab(COLTAB_SIZE)  As RGB_T
  Set_Default_ColTab DefColTab
  Dim Names() As String
  Names = Split(NamesList, ",")
  
  Dim Res As String, Nr As Integer, Comment As String
  Res = "// Set_ColTab(Red Green Blue)      " & vbLf & _
        "Set_ColTab("                ' ^ Space are added to show a gap between the two lines if line break is isabled.
  For Nr = 0 To COLTAB_SIZE - 1
      Res = Res & Right("   " & ColTab(Nr).r, 3) & ", "
      Res = Res & Right("   " & ColTab(Nr).g, 3) & ", "
      Res = Res & Right("   " & ColTab(Nr).b, 3)
      
      If DefColTab(Nr).r <> ColTab(Nr).r Or _
         DefColTab(Nr).g <> ColTab(Nr).g Or _
         DefColTab(Nr).b <> ColTab(Nr).b Then
            Comment = " // *" ' Color is changed compared to the standard ColTab
      Else: Comment = " //  "
      End If
      Comment = Comment & Names(Nr)
      
      If Nr < COLTAB_SIZE - 1 Then
            Res = Res & "," & Comment & vbLf & "           "
      Else: Res = Res & ")" & Comment
      End If
  Next Nr
  ColTab_to_C_String = Res
End Function

'UT---------------------------------
Private Sub TestColTab_to_C_String()
'UT---------------------------------
  Dim ColTab(COLTAB_SIZE) As RGB_T
  Read_ColTab_from_Config_File ColTab
  Debug.Print ColTab_to_C_String(ColTab)
  'ActiveCell = ColTab_to_C_String(ColTab)
End Sub

'----------------------------------------------------------------------------
Private Function C_String_to_ColTab(C_Str As String, ByRef ColTab() As RGB_T)
'----------------------------------------------------------------------------
  Dim line As Variant, Nr As Long
  For Each line In Split(C_Str, vbLf)
     If Left(Trim(line), 2) <> "//" Then
        Dim Parts() As String
        Parts = Split(Trim(Replace(line, "Set_ColTab(", "")), ",")
        If Nr < COLTAB_SIZE Then
           With ColTab(Nr)
             .r = val(Trim(Parts(0)))
             .g = val(Trim(Parts(1)))
             .b = val(Trim(Parts(2)))
           End With
        End If
        Nr = Nr + 1
     End If
  Next line
End Function

'UT----------------------------------
Private Sub Test_C_String_to_ColTab()
'UT----------------------------------
  Dim ColTab(COLTAB_SIZE) As RGB_T, C_Str As String
  Read_ColTab_from_Config_File ColTab
  C_Str = ColTab_to_C_String(ColTab)
  C_String_to_ColTab C_Str, ColTab
  'ColTab(1).g = 212 ' Simulate an error
  Debug.Print "Vergleich: " & (C_Str = ColTab_to_C_String(ColTab))
End Sub


'---------------------------------------
Private Sub Show_Wait_CheckColors_Form()
'---------------------------------------
  Application.EnableEvents = True
  Wait_CheckColors_Form.Activity_Label = "J"  ' :-)
  Smiley_Direction = 0
  Smiley_Cnt = 0
  Wait_CheckColors_Form.Show
  Application.OnTime Now + TimeValue("00:00:03"), "Update_Wait_CheckColors_Form"
  'Debug.Print "*Show_Wait_CheckColors_Form()"
End Sub

'-----------------------------------------
Private Sub Update_Wait_CheckColors_Form()
'-----------------------------------------
' Update the wait dialog and check if the finish file is generated
' If the Proc_CheckColors_Form_Callback function is defined it's
' called when the finish file is detected.
  Const Step = 3
  With Wait_CheckColors_Form.Activity_Label
    Select Case Smiley_Direction
      Case 0: ' Smiley to right
              .Caption = " " & .Caption
              If Len(.Caption) >= 10 Then Smiley_Direction = Smiley_Direction + 1
      Case 1: ' Smiley to left
              .Caption = Mid(.Caption, 2)
              If Len(Wait_CheckColors_Form.Activity_Label) <= 1 Then
                 Smiley_Direction = 0
                 Smiley_Cnt = Smiley_Cnt + 1
                 Select Case Smiley_Cnt
                    Case Step * 1: .Caption = "K" ' :-|
                    Case Step * 2: .Caption = "L" ' :-(
                    Case Step * 3: .Caption = "M" ' Bomb
                    Case Step * 4: .Caption = ">" ' Radioactiv
                    Case Step * 5: .Caption = "T" ' Ice
                    Case Step * 6: .Caption = "N" ' Skull
                    Case Step * 7: .Caption = "(" ' Phone
                    Case Step * 8: .Caption = "J" ' :-)    Start again
                             Smiley_Cnt = 0
                 End Select
              End If
    End Select
  End With
  Calculate
  'Debug.Print "*Update_Wait_CheckColors_Form"
  
  If Dir(ThisWorkbook.Path & "\" & FINISHEDTXT_FILE) <> "" Then
    Wait_CheckColors_Form.Hide
    If Proc_CheckColors_Form_Callback <> "" Then Run Proc_CheckColors_Form_Callback
  Else
    If Wait_CheckColors_Form.Visible Then
       Application.OnTime Now + TimeValue("00:00:01"), "Update_Wait_CheckColors_Form"
    End If
  End If
End Sub


'--------------------------------
Private Sub Set_ColTab_Callback()
'--------------------------------
' Is called if the Python ColorCheck program is closed to set the color table
  Debug.Print "Set_ColTab_Callback"  ' Debug
  Dim ColTab(COLTAB_SIZE) As RGB_T
  Read_ColTab_from_Config_File ColTab
  
  ThisWorkbook.Activate
  Sheets(ColTab_Dest_Sheet).Select
  Make_sure_that_Col_Variables_match
  Cells(ColTab_Dest_Row, Config__Col) = ColTab_to_C_String(ColTab)
  Cells(ColTab_Dest_Row, LEDs____Col).ClearContents                         ' 31.12.19:
  Cells(ColTab_Dest_Row, InCnt___Col).ClearContents                         '   "
  Cells(ColTab_Dest_Row, LocInCh_Col).ClearContents                         '   "
End Sub




' Start des Check Color Programms:
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Es gibt eine py Version und eine Exe Version und seit 27.01.21 eine Zip Version.
' Die Exe muss manuell von Harolds Seite herunter geladen werden. Für die py Version benötigt
' man einen Python Interpreter. Die Zip Version muss zusätzlich entzippt werden.
' Wenn das Programm die EXE Version findet, dann wird diese ausgeführt. Dabei wird davon
' ausgangen, dass der Benutzer sich selber um die neueste Version kümmert.
' Wenn die EXE Version nicht vorhanden ist, dann wird geprüft ob Python installiert ist.
' Ist Pyton nicht installiert, dann wird nachgefragt ob die EXE herunter laden will
' Oder Python installieren will.
'


'UT-------------------------------------------------------------
Public Sub Open_MobaLedCheckColors_and_Insert_Set_ColTab_Macro()
'UT-------------------------------------------------------------
  Open_MobaLedCheckColors "Set_ColTab_Callback", ActiveSheet.Name, ActiveCell.Row
End Sub

'---------------------------------------------------------------------------------------------------------------
Public Sub Open_MobaLedCheckColors(Callback As String, Optional Dest_Sheet As String, Optional Dest_Row As Long)
'---------------------------------------------------------------------------------------------------------------
  Proc_CheckColors_Form_Callback = Callback
  ColTab_Dest_Sheet = Dest_Sheet
  ColTab_Dest_Row = Dest_Row
  

  Dim ProgDir As String
  ProgDir = ThisWorkbook.Path & "\" & CHECKCOL_EXE_DIR
  If Dir(ProgDir, vbDirectory) = "" Then
     'MsgBox Get_Language_Str("Fehler das Verzeichnis existiert nicht:") & vbCr & _
            "  '" & ProgDir & "'", vbCritical, Get_Language_Str("CheckColors Verzeichnis nicht vorhanden")
     CreateFolder ProgDir                                                   ' 27.01.21:
     'Exit Sub                                                              '     "      Disabled
  End If
  
  Close_CheckColors ' Write the "Close" file in case an other version of the CheckColors programm is still running  ' 05.06.20: Moved down
  
  Dim Exe_Exists As Boolean, ExistingVer As String
  Exe_Exists = (Dir(ProgDir & CHECK_COLORS_EXE) <> "")
  
  If Exe_Exists Then
     If GetAsyncKeyState(VK_CONTROL) <> 0 Then   ' Following function must be declared: Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
        If MsgBox(Get_Language_Str("Soll das Farbtest Programm neu heruntergeladen werden?"), vbYesNo + vbQuestion, Get_Language_Str("Neue Version des Farbtests herunterladen?")) = vbYes Then
           Exe_Exists = False
        End If
     End If
  End If
  
  If Not Exe_Exists Then
   #If USE_pyPROGGEN = 0 Then
     If PythonVer <= 0 Then ' Don't check again if it already checked before positive ?
        PythonVer = Check_Phyton(EXP_PYTHON_V_STR, ExistingVer)
     End If
     Dim Msg As String
     If PythonVer <= 0 Then
        If PythonVer < 0 Then Msg = vbCr & Get_Language_Str("Momentan ist die ältere Python Version ") & ExistingVer & Get_Language_Str(" installiert. Es kann evtl. Probleme mit existierenden Python Programmen geben wenn ein neuerer Python Interpreter installiert wird.")
        Select Case MsgBox(Get_Language_Str("Für das Farbauswahl Programm benötigt man entweder die Python " & _
                           "Version ") & EXP_PYTHON_V_STR & Get_Language_Str(" oder das davon unabhängige EXE Programm. " & _
                           "Leider kann letzteres nicht mit der MobaLedLib verteilt " & _
                           "werden, weil EXE Programme nicht in einer Arduino Bibliothek erlaubt sind.") & vbCr & _
                           Msg & _
                           vbCr & _
                           vbCr & _
                           Get_Language_Str("Soll das eigenständige Exe Programm verwendet werden ?" & vbCr & _
                           vbCr & _
                           "Ja:    Das Programm wird automatisch herunter geladen" & vbCr & _
                           "Nein: Python muss manuell installiert werden"), vbQuestion + vbYesNoCancel, _
                           Get_Language_Str("Variante des Farbauswahl Programms bestimmen"))
           Case vbYes: ' Download the Exe
   #End If ' USE_pyPROGGEN = 0
                       MsgBox Get_Language_Str("Das Farbauswahl Programm wird von Github herunter geladen. Dazu wird eine Internetverbindung benötigt." & vbCr & _
                              "Nach dem dieses Meldung bestätigt ist wird ein Kommando Fenster geöffnet in dem der Download ausgeführt wird." & vbCr & _
                              vbCr & _
                              "Achtung: Es kann einige Zeit dauern bis die Verbindung aufgebaut ist. In dieser Zeit ist keine Meldung zu sehen"), _
                              vbInformation, Get_Language_Str("Download der EXE Datei")
                       Dim DestName As String
                       DestName = ThisWorkbook.Path & "\" & CHECKCOL_EXE_DIR & CHECK_COLORS_DST
                       
                       If WIN7_COMPATIBLE_DOWNLOAD Then
                            F_shellExec "powershell Invoke-WebRequest """ & DOWNLOAD_EXEPROG & """ ""-o:" & DestName & """"   ' 20.06.20:
                       Else
                            If Check_if_curl_is_Available_and_gen_Message_if_not(CHECK_COLORS_EXE, DOWNLOAD_EXEPROG) = False Then Exit Sub ' 05.06.20:
                            
                            F_shellExec "powershell curl """ & DOWNLOAD_EXEPROG & """ " & _
                                        """-o:" & DestName & """"
                       End If
                       If Dir(DestName) = "" Then
                             MsgBox Get_Language_Str("Beim herunter laden des Farbtest Programms ist etwas schief gegangen ;-("), vbCritical, _
                                    Get_Language_Str("Fehler beim herunter laden des Farbtest Programms")
                             Exit Sub
                       Else:
                             #If USE_pyPROGGEN = 2 Then                     ' 27.01.21:
                                UnzipAFile DestName, ThisWorkbook.Path & "\" & CHECKCOL_ZIP_DIR ' Wenn die Datei bereits existiert, dann wird eine Meldung angezeigt
                                Kill DestName
                             #End If
                             Exe_Exists = True                              ' 06.06.20:
                       End If
   #If USE_pyPROGGEN = 0 Then
           Case vbNo:  ' Download Python
                       MsgBox Get_Language_Str("Die Download Seite von Python wird gleich geöffnet. Dort lädt man die neueste " & _
                              "Python Version für Windows herunter und installiert sie." & vbCr & _
                              "Bei der Installation muss das Häkchen bei 'Add Python ... to Path' gesetzt werden." & vbCr & _
                              "Diese Häkchen befindet sich auf der ersten Seite unten, gleich beim start der Installation." & vbCr & _
                              vbCr & _
                              "Anschließend muss Excel geschlossen werden sonst wird Python nicht gefunden. " & _
                              "Evtl. muss der Rechner auch neu gestartet werden (Windows ist toll)" & vbCr & _
                              "Dann kann die Farbauswahl Funktion benutzt werden."), _
                              vbInformation, Get_Language_Str("Download und Installation von Python")
                             ' Es reicht nicht wenn Excel neu gestartet wird. Es muss auch von einem neuen Explorer aus gestartet werden ;-(
                             ' Ein Neustart über einen Desktop Link scheint zu funktionieren
                      Shell "Explorer """ & PYTHON_DOWNLOADP & """"
                      Exit Sub
            Case Else: Exit Sub
        End Select
     End If
   #End If ' USE_pyPROGGEN = 0
  End If
  
  
  ' Com Port detection
  Make_sure_that_Col_Variables_match
  
  If Check_USB_Port_with_Dialog(COMPort_COL) = False Then Exit Sub          ' 04.05.20: Added exit (Prior Check_USB_Port_with_Dialog ends the program in case of an error)
  
  Change_Comport_in_ConfigFile Cells(SH_VARS_ROW, COMPort_COL)
  
  ' Color table
  If Callback <> "" Then
     Delete_ColTest_Only_File                                               ' 27.01.21: Moved up
     If InStr(Cells(Dest_Row, Config__Col), "Set_ColTab") > 0 Then
        Dim ColTab(COLTAB_SIZE) As RGB_T
        C_String_to_ColTab Cells(Dest_Row, Config__Col), ColTab
        Insert_ColTab_to_ConfigFile ColTab
        ' Moved up
        ' ATTENTION: This Message box is necessary to generate the delay which prevents the CheckColor
        '            program to be closed by the "Close" File which has been written above
        '            But it doesn't help if the user closes the message to fast because the current version
        '            of the CheckColor program in not writing the "Close" imidiately
        Select Case MsgBox(Get_Language_Str("Soll die Standard Farbtabelle geladen werden oder die zuletzt " & _
                           "benutzte Tabelle benutzt werden?" & vbCr & _
                           vbCr & _
                           "Ja: Standard Farbtabelle laden" & vbCr & _
                           "Nein: Letzte Farbtabelle verwenden"), _
                           vbQuestion + vbYesNoCancel + vbDefaultButton2, _
                           Get_Language_Str("Standard oder letzte Benutzer Farbtabelle verwenden?"))
           Case vbYes:    ' Load standard colors
                          If Not Insert_Default_ColTab_to_ConfigFile() Then Exit Sub
           Case vbNo:     ' Do nothing
           Case vbCancel: Exit Sub
        End Select
     End If
  Else
     Write_ColTest_Only_File                                                ' 27.01.21: Moved down
     StatusMsg_UserForm.Set_ActSheet_Label "Please wait" ' Wait to give the color test program time to read the CLOSE_CHECKCOL_N file
     StatusMsg_UserForm.Show
     Sleep (500)
     Unload StatusMsg_UserForm
  End If
  
    
  Delete_CheckColors_CloseFile                                                                                    ' 18.01.21: Old Position
  If Dir(ThisWorkbook.Path & "\" & FINISHEDTXT_FILE) <> "" Then Kill ThisWorkbook.Path & "\" & FINISHEDTXT_FILE
  
  
  ' Start the CheckColors program
  Dim Res As Boolean
  If Exe_Exists Then
        Res = Start_MobaLedCheckColors_exe
  Else: Res = Start_MobaLedCheckColors_py
  End If
  
  If Res Then
     Show_Wait_CheckColors_Form
  End If
End Sub



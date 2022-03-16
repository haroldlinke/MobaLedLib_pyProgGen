Attribute VB_Name = "M02_Public"
Option Explicit

Public Const Lib_Version_Nr = "3.1.0"                             ' If changed check also "Exp_Prog_Gen_Version" in Pattern_Configurator
Private Const Test_Sufix = "D"                                    ' The Excel programs use the same version number than the library to avoid confusion
Public Const Prog_Version = "Ver. " & Lib_Version_Nr & Test_Sufix ' A sufix could be used for beta version
Public Const Prog_Version_Nr = Lib_Version_Nr & Test_Sufix        ' Call Gen_Release_Version() to update all sheets
                                           

Public Const DEBUG_CHANGEEVENT = False ' Debug Events
Public Const DEBUG_DCCSEND = False     ' Debug DCC sending

Public Const InoName_DCC = "23_A.DCC_Interface.ino"
Public Const InoName__SX = "23_A.Selectrix_Interface.ino"

Public Const Ino_Dir_LED = "LEDs_AutoProg\"
Public Const InoName_LED = "LEDs_AutoProg.ino"
Public Const Cfg_Dir_LED = "Configuration\"
Public Const CfgName_LED = "Configuration.cpp"
Public Const CfgBuild_Script = "build.cmd"

#Const USE_SKETCHBOOK_DIR = True                                      ' 30.05.20:
#If USE_SKETCHBOOK_DIR Then
    Private Const SrcDirInLib = "\libraries\MobaLedLib\extras\"
    Private Const DestDir_All = "\MobaLedLib\Ver_" & Lib_Version_Nr & "\"          ' 20.07.20: All versions are located in "\MobaLedLib\Ver_..." to be able to set one exclusion in the virus scan program for all versions
    Private Const MobaUserDir = "\"
    Private Const Ardu_LibDir = "\libraries\"
    Private Const SrcDirExamp = "\libraries\MobaLedLib\examples\"
#Else
    Public Const SrcDirInLib = "\Documents\Arduino\libraries\MobaLedLib\extras\"   ' 06.12.19: Old: "\libraries\MobaLedLib\examples\23_B.LEDs_AutoProg"
    Public Const DestDir_All = "\Documents\Arduino\MobaLedLib_" & Lib_Version_Nr & "\"
    Public Const MobaUserDir = "\Documents\Arduino\"
    Public Const Ardu_LibDir = "\Documents\Arduino\libraries\"
    Public Const SrcDirExamp = "\Documents\Arduino\libraries\MobaLedLib\examples\"
#End If
Public Const MyExampleDir = "Prog_Generator_Data"
Public Const AppLoc_Ardu = "\AppData\Local\Arduino15\"

Public Const DestDir_LED = DestDir_All & "LEDs_AutoProg"                           ' 06.12.19: Old: ... MobaLedLib_" & Lib_Version_Nr & "_Autoprog\23_B.LEDs_AutoProg\"

Public Const INTPROGNAME = "Prog_Generator"
Public Const DSKLINKNAME = "Prog_Generator MobaLedLib"
Public Const DefaultIcon = "Icons\05_Gerald_Prog.ico"

Public Const SECOND_PROG = "Pattern_Configurator"                                  ' 14.06.20:
Public Const SECOND_LINK = "MobaLedLib Pattern_Configurator"
Public Const SECOND_ICON = "Icons\05_Gerald_Patt.ico"



Public Const WikiPg_Icon = "Icons\WikiMLL_v5.ico"
Public Const WikiPg_Link = "https://wiki.mobaledlib.de/"

Public Const Env_USERPROFILE = "USERPROFILE"


Public Const Include_FileName = "LEDs_AutoProg.h"


Public Const LED_CHANNELS = 10   ' Number of LED Channels                  ' 18.02.22: increase to 10 (8 Led, 1 DMX, 1 virtuell)
Public Const SERIAL_CHANNELS = 8 ' Number of Serial Channels               ' 07.10.21: Juergen Sound extensions
Public Const SerialChannelPrefix = "S"
Public Const INTERNAL_COL_CNT = 1 + 4            ' 03.04.21: Juergen Internal used columns in case of compact LedNr display, which should not be modified

' Special function prefixes
Public Const SF_LED_TO_VAR = "LED_to_Var("
Public Const SF_SERIAL_SOUND_PIN = "SOUND_CHANNEL_DEFINITON("

' Sheet Lib_Macros: (The Sheet can not be changed by the USER => We keep the constants)
Public Const SM_DIALOGDATA_ROW1 = 4

Public Const SM_Typ___COL = 1
Public Const SM_Mode__COL = 2
Public Const SM_LEDS__COL = 3
Public Const SM_InCnt_COL = 4
Public Const SM_OutCntCOL = 5
Public Const SM_LocInCCOL = 6
Public Const SM_Tmp8BtCOL = 7
Public Const SM_SngLEDCOL = 8  ' Single LED Cnt
Public Const SM_DefCh_COL = 9  ' Default Channel                            ' 27.04.20:
Public Const SM_Type__COL = 10 ' CounterType                                ' 01.05.20: Jürgen     Mail: Old Name: SM_CountrCOL
Public Const SM_ListS_COL = 11 ' Sort order in the old List based dialog
Public Const SM_TreeS_COL = 12 ' Sort order in the new tree view dialog
Public Const SM_TMode_COL = 13 ' Mode for tree view (Visible at first start, ...)
Public Const SM_Pic_N_COL = 14 ' Names of the pictures show in the tree view
Public Const SM_Macro_COL = 15
Public Const SM_FindN_COL = 16
Public Const SM_Name__COL = 17
Public Const SM_Group_COL = 18 ' First name of the groups the element belongs to (German) The other languages follow (DeltaCol_Lib_Macro_Lang)
Public Const SM_LName_COL = 19 ' First language specific macro name
Public Const SM_ShrtD_COL = 20 ' First Short description
Public Const SM_DetailCOL = 21 ' First detail description

Public Const DeltaCol_Lib_Macro_Lang = 4 ' Column delta between the languages              ' 07.10.21: Old 2


'begin changes 01.05.20:  Jürgen
'Macro Store Types
Public Const MST_None = 0           ' 0 or empty: no special storage type
Public Const MST_CTR_NONE = 1       ' Counter without status
Public Const MST_CTR_ON = 2         ' Counter with Status, storage default on
Public Const MST_CTR_OFF = 3        ' Counter with Status, storage default off
Public Const MST_PREVENT_STORE = 4  ' Functions prevents storage of status     ' 01.05.20: Mail from Jürgen

'Status Storage Types
Public Const SST_NONE = 0           ' don't store status
Public Const SST_COUNTER_ON = 1     ' store counter status with default on
Public Const SST_COUNTER_OFF = 2    ' store counter status with default off
Public Const SST_S_ONOFF = 3        ' store Channel OnOff status
Public Const SST_TRIGGER = 4        ' store Channel Trigger status
Public Const SST_DISABLED = 5       ' store status disabled                    ' 01.05.20: Mail from Jürgen

Public Const AUTOSTORE_ON = "*"     ' store state is forced on
Public Const AUTOSTORE_OFF = "0"    ' store state is turned off
'end changes 01.05.20:  Jürgen

' Valid for all sheets
Public Const Enable_Col = 2

Public Const Header_Row = 2
Public Const FirstDat_Row = Header_Row + 1
         
Public Const SH_VARS_ROW = 1  ' This row contains some sheet specific varaibles. The text uses white color => It' not visible
Public Const PAGE_ID_COL = 2  ' This cell contains a page ID to idetify the sheet

Public Const AllData_PgIDs = " DCC Selectrix CAN "
Public Const Prog_for_Right_Ardu = " DCC Selectrix Loconet "


Public Const MAX_ROWS = 1048576
Public Const MAX_COLUMNS = 16384


Public Const Hook_CHAR = 61692 ' Font Wingdings

Public Const SPARE_ROWS = 3 ' Number of spare rows which are generated if data are entered in a new line


' Sheet names:
Public Const LANGUAGES_SH = "Languages"
Public Const LIBMACROS_SH = "Lib_Macros"
Public Const PAR_DESCR_SH = "Par_Description"
Public Const LIBRARYS__SH = "Libraries"
Public Const PLATFORMS_SH = "Platform_Parameters"           ' 14.10.21 Juergen
Public Const START_SH = "Start"
Public Const ConfigSheet = "Config"

' Variables
Public SelectMacro_Res As String
Public Userform_Res As String
Public Userform_Res_Address As String   ' 10.02.21: 20210208 added by Misha. Used to store Address for Multiplexer.


Public DialogGuideRes As Long ' uses constants like vbOK, vbAbort, vbNo

Public HouseForm_Pos As WinPos_T
Public OtherForm_Pos As WinPos_T


Public Last_SelectedNr_Valid As Boolean
Public Last_SelectedNr As Long

Public Const ComPortfromOnePage = "" ' Com Port is stored on each page


Public Const MouseHook_Store_Page = "Start"  ' This page is used to store the MouseHook. There must be a named range called "MouseHook".


Public Const BOARD_NANO_OLD = "--board arduino:avr:nano:cpu=atmega328old"
Public Const BOARD_NANO_FULL = "--board arduino:avr:nano:cpu=atmega328fullmem"                                  ' 28.10.20: Jürgen
Public Const BOARD_NANO_NEW = "--board arduino:avr:nano:cpu=atmega328"
Public Const BOARD_NANO_EVERY = "--board arduino:megaavr:nona4809:mode=off"     ' without ATMega328 emulation   ' 28.10.20: Jürgen
Public Const BOARD_UNO_NORM = "--board arduino:avr:uno"
Public Const BOARD_ESP32 = "--board esp32:esp32:esp32:PSRAM=disabled,PartitionScheme=default,CPUFreq=240,FlashMode=qio,FlashFreq=80,FlashSize=4M,UploadSpeed=921600,DebugLevel=none"  ' 10.11.20:
Public Const BOARD_PICO = "--board rp2040:rp2040:rpipico:flash=2097152_0,freq=125,dbgport=Disabled,dbglvl=None"  ' 18.04.21: Juergen
Public Const AUTODETECT_STR = "AutoDet"
Public Const DEFARDPROG_STR = "--pref programmer=arduino:arduinoisp"

Public Sketchbook_Path As String

Public Const L2V_COM_OPERATORS = "< > = != & !&"
Public Const MB_LED_NR_STR = "D2 D3 D4 D5 A3 A2 A1 D7 D8 D9 D10 D11 D12 D13 A4 A5 A0"
Public Const MB_LED_PIN_NR = "0  1  2  3  4  5  6  7  8  9  10  11  12  13  14 15 16"                           ' 30.10.20:


' Public Const USE_SWITCH_AND_LED_ARRAY = False ' Enable this to use the new function from Jürgen which uses an array to read in the SwitchD. This switch is important for the ESP32.

'---------------------------------------------------------------------
Public Function Read_Sketchbook_Path_from_preferences_txt() As Boolean
'---------------------------------------------------------------------
' Attention: The file uses UTF8
  Dim Name As String, FileStr As String
  Name = Environ(Env_USERPROFILE) & AppLoc_Ardu & "preferences.txt"
  FileStr = Read_File_to_String(Name)
  If FileStr <> "#ERROR#" Then
     Sketchbook_Path = ConvertUTF8Str(Get_Ini_Entry(FileStr, "sketchbook.path="))
     'ThisWorkbook.Sheets(LIBRARYS__SH).Range("Sketchbook_Path") = Sketchbook_Path
     If Sketchbook_Path = "#ERROR#" Then
        MsgBox Replace(Get_Language_Str("Fehler: beim lesen des 'sketchbook.path' in '#1#'"), "#1#", Name), vbCritical, Get_Language_Str("Fehler beim Lesen der Datei:") & " 'preferences.txt'"
        Exit Function
     End If
     If left(Sketchbook_Path, 2) = "\\" Then
        MsgBox Get_Language_Str("Fehler: Der Arduino 'sketchbook.path' darf kein Netzlaufwerk sein:") & vbCr & _
                                "  '" & Sketchbook_Path & "'", vbCritical, Get_Language_Str("Ungültiger Arduino 'sketchbook.path'")
        Exit Function
     End If
     CreateFolder Sketchbook_Path & "\"
     Read_Sketchbook_Path_from_preferences_txt = True
  End If
End Function

'----------------------------------------------
Public Function Get_Sketchbook_Path() As String
'----------------------------------------------
  Debug.Print "Get_Sketchbook_Path called"
  If Sketchbook_Path = "" Then
     Read_Sketchbook_Path_from_preferences_txt
  End If
  Get_Sketchbook_Path = Sketchbook_Path
End Function


#If USE_SKETCHBOOK_DIR Then
  Public Function Get_SrcDirInLib() As String:  Get_SrcDirInLib = Get_Sketchbook_Path() & SrcDirInLib: End Function
  Public Function Get_DestDir_All() As String:  Get_DestDir_All = Get_Sketchbook_Path() & DestDir_All: End Function
  Public Function Get_MobaUserDir() As String:  Get_MobaUserDir = Get_Sketchbook_Path() & MobaUserDir: End Function
  Public Function Get_Ardu_LibDir() As String:  Get_Ardu_LibDir = Get_Sketchbook_Path() & Ardu_LibDir: End Function
  Public Function Get_SrcDirExamp() As String:  Get_SrcDirExamp = Get_Sketchbook_Path() & SrcDirExamp: End Function
#Else
  Public Function Get_SrcDirInLib() As String:  Get_SrcDirInLib = Environ(Env_USERPROFILE) & SrcDirInLib: End Function
  Public Function Get_DestDir_All() As String:  Get_DestDir_All = Environ(Env_USERPROFILE) & DestDir_All: End Function
  Public Function Get_MobaUserDir() As String:  Get_MobaUserDir = Environ(Env_USERPROFILE) & MobaUserDir: End Function
  Public Function Get_Ardu_LibDir() As String:  Get_Ardu_LibDir = Environ(Env_USERPROFILE) & Ardu_LibDir: End Function
  Public Function Get_SrcDirExamp() As String:  Get_SrcDirExamp = Environ(Env_USERPROFILE) & SrcDirExamp: End Function
#End If

'---------------------------------------
Public Function Get_BoardTyp() As String                                    ' 29.10.20:
'---------------------------------------
' The build options for the ESP32 are something like "esp32:esp32:esp32..."

  If InStr(Cells(SH_VARS_ROW, BUILDOP_COL), "esp32") > 0 Then
        Get_BoardTyp = "ESP32"
  ElseIf InStr(Cells(SH_VARS_ROW, BUILDOP_COL), "rp2040") > 0 Then          ' 17.04.21: Juergen
        Get_BoardTyp = "PICO"
  Else: Get_BoardTyp = "AM328"  ' ATMega328
  End If
  ' Other types:
  ' "Every"           ' Nano Every
End Function


' Used Cmd colors:                 (See: https://ss64.com/nt/color.html)
' ~~~~~~~~~~~~~~~~
' 1F" ' White on Blue                     Arduino Comile DCC
' 2F" ' White on Green                    Arduino Comile SX
' 3F" ' White on Aqua                     Arduino Comile CAN
' 4F" ' Yellow on Red                     Error
' 5F" ' White on Purple                   Create_InstalLib_Cmd_file              Wird das noch gebraucht ? => Nein ==> Ist deaktiviert
' 80  ' Black on bright Gray              Do_Update_Script
' 79" ' Blue  on bright Gray              Restart_Cmd


' Links for 32 and 64 Bit Windows:
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' - General description: https://codekabinett.com/rdumps.php?Lang=2&targetDoc=windows-api-declaration-vba-64-bit
'   "Also new with VBA7 are the two new compiler constants Win64 and VBA7.
'    VBA7 is true if your code runs in the VBA7-Environment (Access/Office 2010 and above).
'    Win64 is true if your code actually runs in the 64-bit VBA environment.
'    Win64 is not true if you run a 32-Bit VBA Application on a 64-bit system."
' - Overview 32 / 64 Bit functions: https://jkp-ads.com/Articles/apideclarations.asp

' Following parts use declared external functions                 (List is not updated)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' - M06_Write_Header:
'    - Detect if CTRL is pressed when the "Arduino" button is pressed
'      Public Sub Write_Header_File_and_Upload_to_Arduino()
' - M40_Mouse_Scroll
'    - Uses a lot of functions to be able to use the scroll wheel
' - M31_Sound
'    - Play a windows sound if the hook is enabled/disabled
'      BeepThis2()
' - M24_Mouse_Insert_Pos
'    - Mouse cursor if lines are moved (Mouse or Keyboard)
' - M40_ShellAndWait
'    - Start the Arduino Compiler
' - M30_Tools
'    - Sleep         => Some locatons
'    - ShellExecute  => EditFile_Click
'    - GetKeyState   => Not used



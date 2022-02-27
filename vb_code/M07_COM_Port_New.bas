Attribute VB_Name = "M07_COM_Port_New"
Option Explicit

' Select the Arduino COM Port
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
' Uses the Modules
' - M07_COM_Port
' - Select_COM_Port_UserForm

Private Const PRINT_DEBUG = False

Public CheckCOMPort As Long
Public CheckCOMPort_Txt As String
Private CheckCOMPort_Res As Long

'------------------------------
Private Sub Blink_Arduino_LED()                                             ' 02.05.20:
'------------------------------
' Is called by OnTime and flashes the LEDs of the Arduino connected to
' the port stored in the global variable "CheckCOMPort"
' It's aborted if CheckCOMPort = 0
' Attention: This function doesn't check if the connected device is an Arduino
' because this would be to slow. In addition the blinking frequence is more visible if
' A baudrate of 50 is used.
  Dim SWMajorVersion As Byte, SWMinorVersion As Byte, HWVersion As Byte
  Dim DeviceSignatur As Long, BaudRate As Long
  BaudRate = 50 ' Using a wrong slow baudrate to get nice flashing LEDs
  If CheckCOMPort > 0 Then
     Select_COM_Port_UserForm.Update_SpinButton 0 ' Rescan COM Ports
     If CheckCOMPort <> 999 Then
           CheckCOMPort_Res = DetectArduino(CheckCOMPort, BaudRate, HWVersion, SWMajorVersion, SWMinorVersion, DeviceSignatur, 1, PrintDebug:=PRINT_DEBUG)
     Else: CheckCOMPort_Res = -9
     End If
     Application.Cursor = xlNorthwestArrow
     'Debug.Print "CheckCOMPort_Res=" & CheckCOMPort_Res & "  CheckCOMPort=" & CheckCOMPort
     If CheckCOMPort_Res < 0 Then
           If CheckCOMPort = 999 Then
                 Select_COM_Port_UserForm.Show_Status True, Get_Language_Str("Kein COM Port erkannt." & vbCr & _
                                                                             "Bitte Arduino an einen USB Anschluss des Computers anschließen")
           Else: Select_COM_Port_UserForm.Show_Status True, Get_Language_Str("Achtung: Der Arduino wird von einem anderen Programm benutzt." & vbCr & _
                                                                             "(Serieller Monitor?)" & vbCr & _
                                                                             "Das Programm muss geschlossen werden! ")
           End If
     Else: Select_COM_Port_UserForm.Show_Status False, CheckCOMPort_Txt ' CheckCOMPort_Res >= 0
     End If
     Sleep 10
     DoEvents ' To be able to abort with Ctrl+Break
     Application.OnTime Now + TimeValue("00:00:00"), "Blink_Arduino_LED" ' Restart it again
  End If
End Sub


'--------------------------------------------------------------------------------------
Public Function Select_Arduino_w_Blinking_LEDs_Dialog(Caption As String, Title As String, Text As String, _
                                                      Picture As String, _
                                                      Buttons As String, _
                                                      ByRef ComPort_IO As Long) As Long                     ' 02.05.20:
'--------------------------------------------------------------------------------------
' This function is called if several Arduinos have been detected
' or if the COM port of the actual Arduino is buzy
' Variables:
'  Caption     Dialog Caption
'  Title       Dialog Title
'  Text        Message in the text box on the top left side
'  Picture     Name of the picture to be shown. Available pictures: "LED_Image", "CAN_Image", "Tiny_Image", "DCC_Image"
'  Buttons     List of 3 buttons with Accelerator. Example "H Hallo; A Abort; O Ok"  Two Buttons: " ; A Abort; "O Ok"
'  ComPort_IO  is used as input and output. Negativ numbe if it's buzy (Used by an other program)
' Return:
'  1: If the left   Button is pressed  (Install, ...)
'  2: If the middle Button is pressed  (Abort)
'  3: If the right  Button is pressed  (OK)

  CheckCOMPort = 999 ' Prevent stopping the Blink_Arduino_LED() function. It's Updated in the dialog
  Application.OnTime Now + TimeValue("00:00:00"), "Blink_Arduino_LED"
  ' Return values of Select_COM_Port_UserForm.ShowDialog:
  '  -1: If Abort is pressed
  '   0: If No COM Port is available
  '  >0: Selected COM Port
  ' The variable "CheckCOMPort_Res" is >= 0 if the Port is available
  Dim Res As Long
  Select_Arduino_w_Blinking_LEDs_Dialog = Select_COM_Port_UserForm.ShowDialog(Caption, Title, Text, Picture, Buttons, "", True, _
                                                 Get_Language_Str("Tipp: Der ausgewählte Arduino blinkt schnell"), ComPort_IO, PRINT_DEBUG)
  If CheckCOMPort_Res < 0 Then ComPort_IO = -ComPort_IO ' Port is buzy
  Application.Cursor = xlDefault
End Function

'UT-----------------------------------------------------
Private Sub Test_Select_Arduino_w_Blinking_LEDs_Dialog()
'UT-----------------------------------------------------
  Dim ComPort As Long: ComPort = 3
  Debug.Print "Res=" & Select_Arduino_w_Blinking_LEDs_Dialog("LED_Image""Auswahl des Arduinos", _
                                            "New Title", _
                                            "Mit diesem Dialog wird der COM Port gewählt an den der Arduino angeschlossen ist.", _
                                            "LED_Image", _
                                            "H Hallo;T Test;O O", _
                                            ComPort)
  Debug.Print "ComPort=" & ComPort
End Sub

'---------------------------------------------------------------------------------------------
Private Function Show_USB_Port_Dialog(ComPortColumn As Long, ByRef ComPort As Long) As Boolean
'---------------------------------------------------------------------------------------------
    Dim Res As Long, Picture As String, ArduName As String
    ComPort = val(Cells(SH_VARS_ROW, ComPortColumn))
    If ComPort < 0 Then ComPort = -ComPort
    Select Case ComPortColumn
        Case COMPort_COL: Picture = "LED_Image":  ArduName = "LED"
        Case COMPrtR_COL: Picture = "DCC_Image":  ArduName = Page_ID
        Case COMPrtT_COL: Picture = "Tiny_Image": ArduName = "ISP"
        Case Else: MsgBox "Internal Error: Unsupported  ComPortColumn=" & ComPortColumn & " in 'USB_Port_Dialog()'", vbCritical, "Internal Error"
                   EndProg
    End Select
    Res = Select_Arduino_w_Blinking_LEDs_Dialog(Get_Language_Str("Überprüfung des USB Ports"), _
                                                Get_Language_Str("Auswahl des Arduino COM Ports"), _
                                                Replace(Get_Language_Str("Mit diesem Dialog wird der COM Port überprüft " & _
                                                "bzw. ausgewählt an den der #1# Arduino angeschlossen ist." & vbCr & _
                                                vbCr & _
                                                "OK, wenn die LEDs am richtigen Arduino schnell blinken."), "#1#", ArduName), _
                                                Picture, _
                                                Get_Language_Str(" ; A Abbruch; O Ok"), ComPort) ' 23.06.20: Added Get_Language_Str()
    Show_USB_Port_Dialog = (Res = 3)
End Function



'----------------------------------------------------------------
Public Function USB_Port_Dialog(ComPortColumn As Long) As Boolean
'----------------------------------------------------------------
  Dim ComPort As Long
  If Show_USB_Port_Dialog(ComPortColumn, ComPort) Then
     If ComPort > 0 Then USB_Port_Dialog = True
     ComPortPage().Cells(SH_VARS_ROW, ComPortColumn) = ComPort ' If the port is buzy a negativ number is written
  End If
End Function

'UT-------------------------------
Private Sub Test_USB_Port_Dialog()
'UT-------------------------------
#If PROG_GENERATOR_PROG Then
  Make_sure_that_Col_Variables_match
  
  USB_Port_Dialog COMPort_COL
  'USB_Port_Dialog COMPrtR_COL
  'USB_Port_Dialog COMPrtT_COL  ' Could only be used im the Pattern_COnfigurator
#End If
End Sub


'--------------------------------------------------------------------------------------------------------
Public Function Detect_Com_Port(Optional RightSide As Boolean, Optional Pic_ID As String = "DCC") As Long
'--------------------------------------------------------------------------------------------------------
  Dim Res As Boolean, ComPortColumn As Long, ComPort As Long
  If RightSide Then
        ComPortColumn = COMPrtR_COL
  Else: ComPortColumn = COMPort_COL
  End If
  If Show_USB_Port_Dialog(ComPortColumn, ComPort) Then Detect_Com_Port = ComPort
End Function

'-------------------------------------------------------------------------
Private Sub Test_Check_If_Arduino_could_be_programmed_and_set_Board_type()
'-------------------------------------------------------------------------
  Dim BuildOptions As String
  
  #If PATTERN_CONFIG_PROG Then
    ' TinyUniProg
    With Cells(SH_VARS_ROW, BuildOT_COL)
      If .Value = "" Then .Value = 115200
    End With
    Dim DeviceSignature As Long
    Debug.Print Check_If_Arduino_could_be_programmed_and_set_Board_type(COMPrtT_COL, BuildOT_COL, BuildOptions, DeviceSignature) & " BuildOptions: " & BuildOptions
  #End If
End Sub


' For some reasons this function was available two times.                              30.10.20:
' The second which is located in "M08_Arduino" was defined as "Private"
' Since the functions are nearely equal only one is active now

''---------------------------------------------------------------------------------------------------------------------------------------------------------------
'Public Function Check_If_Arduino_could_be_programmed_and_set_Board_type(ComPortColumn As Long, BuildOptColumn As Long, ByRef BuildOptions As String) As Boolean ' 04.05.20:
''---------------------------------------------------------------------------------------------------------------------------------------------------------------
'' The "Buzy" check and the automatic board detection is only active if Autodetect is enabled
'' Otherwise the values in the BuildOptColumn are used
'' Result: BuildOptions
'  Dim Start_Baudrate As Long, BaudRate As Long, ComPort As Long, Msg As String, Retry As Boolean, AutoDetect As Boolean
'
'  Do
'    Retry = False
'    If Check_USB_Port_with_Dialog(ComPortColumn) = False Then Exit Function ' Display Dialog if the COM Port is negativ and ask the user to correct it
'
'    ' Now we are sure that the com port is positiv. Check if it could be accesed and get the Baud rate
'    BuildOptions = Cells(SH_VARS_ROW, BuildOptColumn)
'    AutoDetect = InStr(BuildOptions, AUTODETECT_STR) > 0
'    If AutoDetect Then
'       BuildOptions = Trim(Replace(BuildOptions, AUTODETECT_STR, ""))
'       If InStr(BuildOptions, BOARD_NANO_OLD) Or InStr(BuildOptions, BOARD_UNO_NORM) > 0 Then  ' Set the Default Baudrate to speed up the check
'             Start_Baudrate = 57600
'       Else: Start_Baudrate = 115200
'       End If
'    End If
'    ComPort = val(Cells(SH_VARS_ROW, ComPortColumn))
'    Dim DeviceSignatur As Long, FirmwareVer As String
'    BaudRate = Get_Arduino_Baudrate(ComPort, Start_Baudrate, DeviceSignatur, FirmwareVer)  ' 28.10.20: Jürgen: Added: DeviceSignatur
'    If BaudRate <= 0 Then
'          If Check_If_Port_is_Available(ComPort) = False Then
'                Msg = Get_Language_Str("Fehler: Es ist kein Arduino an COM Port #1# angeschlossen.")
'          ElseIf BaudRate = 0 Then
'                Msg = Get_Language_Str("Fehler: Das Gerät am COM Port #1# wurde nicht als Arduino erkannt." & vbCr & _
'                                       "Evtl. ist es ein defekter Arduino oder der Bootloader ist falsch.")
'          Else: Msg = Get_Language_Str("Fehler: Der COM Port #1# wird bereits von einem anderen Programm benutzt." & vbCr & _
'                                       "Das kann z.B. der serielle Monitor der Arduino IDE oder das Farbtestprogramm sein." & vbCr & _
'                                       vbCr & _
'                                       "Das entsprechende Programm muss geschlossen werden.")
'          End If
'          Msg = Replace(Msg, "#1#", ComPort) & vbCr & vbCr & Get_Language_Str("Wollen sie es noch mal mit einem anderen Arduino oder einem anderen COM Port versuchen?")
'          If MsgBox(Msg, vbYesNo + vbQuestion, Get_Language_Str("Fehler bei der Überprüfung des angeschlossenen Arduinos")) = vbYes Then
'                Retry = True
'                With Cells(SH_VARS_ROW, ComPortColumn)
'                   .Value = -val(.Value) ' Set to a negativ number to show the COM Port dialog
'                End With
'          Else: Exit Function
'          End If
'    Else
'          If AutoDetect Then
'             'If BaudRate <> Start_Baudrate Then ' Change the board type to speed up the check the next time ' 30.10.20: Always update the board type and Baud rate
'                Dim NewBrd As String, LeftArduino As Boolean
'                If BaudRate = 57600 Then
'                      NewBrd = BOARD_NANO_OLD
'                Else: ' An UNO can't be detected every time, but it could be programmed like a Nano with the new bootloader
'                      NewBrd = Get_New_Board_Type(FirmwareVer)              ' 29.10.20:
'                End If
'
'                LeftArduino = (ComPortColumn = COMPort_COL)
'                Change_Board_Typ LeftArduino, NewBrd ' Write the new board type
'                BuildOptions = Cells(SH_VARS_ROW, BuildOptColumn) ' Reread the Build options in case the board type was adapted
'                BuildOptions = Trim(Replace(BuildOptions, AUTODETECT_STR, "")) ' Remove the Autodetect flag
'             'End If
'          End If
'    End If
'  Loop While Retry
'
'  Check_If_Arduino_could_be_programmed_and_set_Board_type = True
'End Function




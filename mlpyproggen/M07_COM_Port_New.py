# -*- coding: utf-8 -*-
#
#         Write header
#
# * Version: 1.21
# * Author: Harold Linke
# * Date: January 1st, 2020
# * Copyright: Harold Linke 2020
# *
# *
# * MobaLedCheckColors on Github: https://github.com/haroldlinke/MobaLedCheckColors
# *
# *
# * History of Change
# * V1.00 10.03.2020 - Harold Linke - first release
# *  
# * https://github.com/Hardi-St/MobaLedLib
# *
# * MobaLedCheckColors is free software: you can redistribute it and/or modify
# * it under the terms of the GNU General Public License as published by
# * the Free Software Foundation, either version 3 of the License, or
# * (at your option) any later version.
# *
# * MobaLedCheckColors is distributed in the hope that it will be useful,
# * but WITHOUT ANY WARRANTY; without even the implied warranty of
# * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# * GNU General Public License for more details.
# *
# * You should have received a copy of the GNU General Public License
# * along with this program.  if not, see <http://www.gnu.org/licenses/>.
# *
# *
# ***************************************************************************

from vb2py.vbfunctions import *
from vb2py.vbdebug import *
from vb2py.vbconstants import *

#import mlpyproggen.M02_Public as M02
#import mlpyproggen.P01_Workbook as P01
#import mlpyproggen.M08_ARDUINO as M08
#import mlpyproggen.M09_Language as M09
#import mlpyproggen.M25_Columns as M25
#import mlpyproggen.M30_Tools as M30
#import mlpyproggen.U01_userform as U01

import mlpyproggen.M02_Public as M02
#import mlpyproggen.M03_Dialog as M03
#import mlpyproggen.M06_Write_Header_LED2Var as M06LED
#import mlpyproggen.M06_Write_Header_Sound as M06Sound
#import mlpyproggen.M06_Write_Header as M06
import mlpyproggen.M07_COM_Port as M07
import mlpyproggen.M08_ARDUINO as M08
import mlpyproggen.M09_Language as M09
#import mlpyproggen.M09_Select_Macro as M09SM
#import mlpyproggen.M09_SelectMacro_Treeview as M09SMT
#import mlpyproggen.M10_Par_Description as M10
#import mlpyproggen.M20_PageEvents_a_Functions as M20
import mlpyproggen.M25_Columns as M25
#import mlpyproggen.M27_Sheet_Icons as M27
#import mlpyproggen.M28_divers as M28
import mlpyproggen.M30_Tools as M30
#import mlpyproggen.M31_Sound as M31
#import mlpyproggen.M37_Inst_Libraries as M37
#import mlpyproggen.M60_CheckColors as M60
#import mlpyproggen.M70_Exp_Libraries as M70
#import mlpyproggen.M80_Create_Mulitplexer as M80

import mlpyproggen.P01_Workbook as P01

import mlpyproggen.D08_Select_COM_Port_Userform as D08

from vb2py.vbfunctions import *
from vb2py.vbdebug import *

Select_COM_Port_UserForm = D08.CSelect_COM_Port_UserForm()

""" Select the Arduino COM Port
 ~~~~~~~~~~~~~~~~~~~~~~~~~~~

 Uses the Modules
 - M07_COM_Port
 - Select_COM_Port_UserForm
------------------------------
--------------------------------------------------------------------------------------
UT-----------------------------------------------------
---------------------------------------------------------------------------------------------
----------------------------------------------------------------
UT-------------------------------
--------------------------------------------------------------------------------------------------------
-------------------------------------------------------------------------
 For some reasons this function was available two times.                              30.10.20:
 The second which is located in "M08_Arduino" was defined as "Private"
 Since the functions are nearely equal only one is active now
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Check_If_Arduino_could_be_programmed_and_set_Board_type(ComPortColumn As Long, BuildOptColumn As Long, ByRef BuildOptions As String) As Boolean ' 04.05.20:
'---------------------------------------------------------------------------------------------------------------------------------------------------------------
' The "Buzy" check and the automatic board detection is only active if Autodetect is enabled
' Otherwise the values in the BuildOptColumn are used
' Result: BuildOptions
  Dim Start_Baudrate As Long, BaudRate As Long, ComPort As Long, Msg As String, Retry As Boolean, AutoDetect As Boolean

  Do
    Retry = False
    If Check_USB_Port_with_Dialog(ComPortColumn) = False Then Exit Function ' Display Dialog if the COM Port is negativ and ask the user to correct it

    ' Now we are sure that the com port is positiv. Check if it could be accesed and get the Baud rate
    BuildOptions = Cells(SH_VARS_ROW, BuildOptColumn)
    AutoDetect = InStr(BuildOptions, AUTODETECT_STR) > 0
    If AutoDetect Then
       BuildOptions = Trim(Replace(BuildOptions, AUTODETECT_STR, ""))
       If InStr(BuildOptions, BOARD_NANO_OLD) Or InStr(BuildOptions, BOARD_UNO_NORM) > 0 Then  ' Set the Default Baudrate to speed up the check
             Start_Baudrate = 57600
       Else: Start_Baudrate = 115200
       End If
    End If
    ComPort = val(Cells(SH_VARS_ROW, ComPortColumn))
    Dim DeviceSignatur As Long, FirmwareVer As String
    BaudRate = Get_Arduino_Baudrate(ComPort, Start_Baudrate, DeviceSignatur, FirmwareVer)  ' 28.10.20: Jürgen: Added: DeviceSignatur
    If BaudRate <= 0 Then
          If Check_If_Port_is_Available(ComPort) = False Then
                Msg = Get_Language_Str("Fehler: Es ist kein Arduino an COM Port #1# angeschlossen.")
          ElseIf BaudRate = 0 Then
                Msg = Get_Language_Str("Fehler: Das Gerät am COM Port #1# wurde nicht als Arduino erkannt." & vbCr & _
'                                       "Evtl. ist es ein defekter Arduino oder der Bootloader ist falsch.")
          Else: Msg = Get_Language_Str("Fehler: Der COM Port #1# wird bereits von einem anderen Programm benutzt." & vbCr & _
'                                       "Das kann z.B. der serielle Monitor der Arduino IDE oder das Farbtestprogramm sein." & vbCr & _
'                                       vbCr & _
'                                       "Das entsprechende Programm muss geschlossen werden.")
          End If
          Msg = Replace(Msg, "#1#", ComPort) & vbCr & vbCr & Get_Language_Str("Wollen sie es noch mal mit einem anderen Arduino oder einem anderen COM Port versuchen?")
          If MsgBox(Msg, vbYesNo + vbQuestion, Get_Language_Str("Fehler bei der Überprüfung des angeschlossenen Arduinos")) = vbYes Then
                Retry = True
                With Cells(SH_VARS_ROW, ComPortColumn)
                   .Value = -val(.Value) ' Set to a negativ number to show the COM Port dialog
                End With
          Else: Exit Function
          End If
    Else
          If AutoDetect Then
             'If BaudRate <> Start_Baudrate Then ' Change the board type to speed up the check the next time ' 30.10.20: Always update the board type and Baud rate
                Dim NewBrd As String, LeftArduino As Boolean
                If BaudRate = 57600 Then
                      NewBrd = BOARD_NANO_OLD
                Else: ' An UNO can't be detected every time, but it could be programmed like a Nano with the new bootloader
                      NewBrd = Get_New_Board_Type(FirmwareVer)              ' 29.10.20:
                End If

                LeftArduino = (ComPortColumn = COMPort_COL)
                Change_Board_Typ LeftArduino, NewBrd ' Write the new board type
                BuildOptions = Cells(SH_VARS_ROW, BuildOptColumn) ' Reread the Build options in case the board type was adapted
                BuildOptions = Trim(Replace(BuildOptions, AUTODETECT_STR, "")) ' Remove the Autodetect flag
             'End If
          End If
    End If
  Loop While Retry

  Check_If_Arduino_could_be_programmed_and_set_Board_type = True
End Function
"""

__PRINT_DEBUG = False
CheckCOMPort = Long()
CheckCOMPort_Txt = String()
__CheckCOMPort_Res = Long()

def __Blink_Arduino_LED():
    global CheckCOMPort,__CheckCOMPort_Res,CheckCOMPort_Txt
    SWMajorVersion = Byte()

    SWMinorVersion = Byte()

    HWVersion = Byte()

    DeviceSignatur = Long()

    BaudRate = Long()
    #------------------------------
    # Is called by OnTime and flashes the LEDs of the Arduino connected to
    # the port stored in the global variable "CheckCOMPort"
    # It's aborted if CheckCOMPort = 0
    # Attention: This function doesn't check if the connected device is an Arduino
    # because this would be to slow. In addition the blinking frequence is more visible if
    # A baudrate of 50 is used.
    BaudRate = 50
    if CheckCOMPort > 0:
        D08.Select_COM_Port_UserForm.Update_SpinButton(0)
        if CheckCOMPort != 999:
            __CheckCOMPort_Res = DetectArduino(CheckCOMPort, BaudRate, HWVersion, SWMajorVersion, SWMinorVersion, DeviceSignatur, 1, PrintDebug= __PRINT_DEBUG)
        else:
            __CheckCOMPort_Res = - 9
        #*HLApplication.Cursor = xlNorthwestArrow
        #Debug.Print "CheckCOMPort_Res=" & CheckCOMPort_Res & "  CheckCOMPort=" & CheckCOMPort
        if __CheckCOMPort_Res < 0:
            if CheckCOMPort == 999:
                D08.Select_COM_Port_UserForm.Show_Status(True, M09.Get_Language_Str('Kein COM Port erkannt.' + vbCr + 'Bitte Arduino an einen USB Anschluss des Computers anschließen'))
            else:
                D08.Select_COM_Port_UserForm.Show_Status(True, M09.Get_Language_Str('Achtung: Der Arduino wird von einem anderen Programm benutzt.' + vbCr + '(Serieller Monitor?)' + vbCr + 'Das Programm muss geschlossen werden! '))
        else:
            D08.Select_COM_Port_UserForm.Show_Status(False, CheckCOMPort_Txt)
        #*HL Sleep(10)
        P01.DoEvents()
        P01.Application.OnTime(P01.Now + P01.TimeValue('00:00:00'), 'Blink_Arduino_LED')

# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ComPort_IO - ByRef 
def Select_Arduino_w_Blinking_LEDs_Dialog(Caption, Title, Text, Picture, Buttons, ComPort_IO):
    global CheckCOMPort,__CheckCOMPort_Res,CheckCOMPort_Txt
    fn_return_value = None
    Res = Long()
    #--------------------------------------------------------------------------------------
    # This function is called if several Arduinos have been detected
    # or if the COM port of the actual Arduino is buzy
    # Variables:
    #  Caption     Dialog Caption
    #  Title       Dialog Title
    #  Text        Message in the text box on the top left side
    #  Picture     Name of the picture to be shown. Available pictures: "LED_Image", "CAN_Image", "Tiny_Image", "DCC_Image"
    #  Buttons     List of 3 buttons with Accelerator. Example "H Hallo; A Abort; O Ok"  Two Buttons: " ; A Abort; "O Ok"
    #  ComPort_IO  is used as input and output. Negativ numbe if it's buzy (Used by an other program)
    # Return:
    #  1: If the left   Button is pressed  (Install, ...)
    #  2: If the middle Button is pressed  (Abort)
    #  3: If the right  Button is pressed  (OK)
    CheckCOMPort = 999
    P01.Application.OnTime(P01.Now + P01.TimeValue('00:00:00'), 'Blink_Arduino_LED')
    # Return values of Select_COM_Port_UserForm.ShowDialog:
    #  -1: If Abort is pressed
    #   0: If No COM Port is available
    #  >0: Selected COM Port
    # The variable "CheckCOMPort_Res" is >= 0 if the Port is available
    fn_return_value = D08.Select_COM_Port_UserForm.ShowDialog(Caption, Title, Text, Picture, Buttons, '', True, M09.Get_Language_Str('Tipp: Der ausgewählte Arduino blinkt schnell'), ComPort_IO, __PRINT_DEBUG)
    if __CheckCOMPort_Res < 0:
        ComPort_IO = - ComPort_IO
        # Port is buzy
    P01.Application.Cursor = xlDefault
    return fn_return_value

def __Test_Select_Arduino_w_Blinking_LEDs_Dialog():
    ComPort = Long()
    #UT-----------------------------------------------------
    ComPort = 3
    Debug.Print('Res=' + Select_Arduino_w_Blinking_LEDs_Dialog('LED_Image"Auswahl des Arduinos', 'New Title', 'Mit diesem Dialog wird der COM Port gewählt an den der Arduino angeschlossen ist.', 'LED_Image', 'H Hallo;T Test;O O', ComPort))
    Debug.Print('ComPort=' + ComPort)

# VB2PY (UntranslatedCode) Argument Passing Semantics / Decorators not supported: ComPort - ByRef 
def __Show_USB_Port_Dialog(ComPortColumn, ComPort):
    fn_return_value = False
    Res = Long()

    Picture = String()

    ArduName = String()
    #---------------------------------------------------------------------------------------------
    ComPort = P01.val(P01.Cells(M02.SH_VARS_ROW, ComPortColumn))
    if ComPort < 0:
        ComPort = - ComPort
    if (ComPortColumn == M25.COMPort_COL):
        Picture = 'LED_Image'
        ArduName = 'LED'
    elif (ComPortColumn == M25.COMPrtR_COL):
        Picture = 'DCC_Image'
        ArduName = M25.Page_ID
    elif (ComPortColumn == M25.COMPrtT_COL):
        Picture = 'Tiny_Image'
        ArduName = 'ISP'
    else:
        P01.MsgBox('Internal Error: Unsupported  ComPortColumn=' + ComPortColumn + ' in \'USB_Port_Dialog()\'', vbCritical, 'Internal Error')
        M30.EndProg()
    Res = Select_Arduino_w_Blinking_LEDs_Dialog(M09.Get_Language_Str('Überprüfung des USB Ports'), M09.Get_Language_Str('Auswahl des Arduino COM Ports'), Replace(M09.Get_Language_Str('Mit diesem Dialog wird der COM Port überprüft ' + 'bzw. ausgewählt an den der #1# Arduino angeschlossen ist.' + vbCr + vbCr + 'OK, wenn die LEDs am richtigen Arduino schnell blinken.'), '#1#', ArduName), Picture, M09.Get_Language_Str(' ; A Abbruch; O Ok'), ComPort)
    fn_return_value = ( Res == 3 )
    return fn_return_value

def USB_Port_Dialog(ComPortColumn):
    fn_return_value = False
    ComPort = Long()
    #----------------------------------------------------------------
    if __Show_USB_Port_Dialog(ComPortColumn, ComPort):
        if ComPort > 0:
            fn_return_value = True
        M07.ComPortPage.Cells[M02.SH_VARS_ROW, ComPortColumn] = ComPort
    return fn_return_value

def __Test_USB_Port_Dialog():
    #UT-------------------------------
    ## VB2PY (CheckDirective) VB directive took path 1 on PROG_GENERATOR_PROG
    M25.Make_sure_that_Col_Variables_match()
    USB_Port_Dialog()(M25.COMPort_COL)
    #USB_Port_Dialog COMPrtR_COL
    #USB_Port_Dialog COMPrtT_COL  ' Could only be used im the Pattern_COnfigurator

def Detect_Com_Port(RightSide=VBMissingArgument, Pic_ID='DCC'):
    fn_return_value = None
    Res = Boolean()

    ComPortColumn = Long()

    ComPort = Long()
    #--------------------------------------------------------------------------------------------------------
    if RightSide:
        ComPortColumn = M25.COMPrtR_COL
    else:
        ComPortColumn = M25.COMPort_COL
    if __Show_USB_Port_Dialog(ComPortColumn, ComPort):
        fn_return_value = ComPort
    return fn_return_value

def __Test_Check_If_Arduino_could_be_programmed_and_set_Board_type():
    BuildOptions = String()

    DeviceSignature = Long()
    #-------------------------------------------------------------------------
    ## VB2PY (CheckDirective) VB directive took path 1 on PATTERN_CONFIG_PROG
    # TinyUniProg
    with_0 = P01.Cells(M02.SH_VARS_ROW, M25.BuildOT_COL)
    if with_0.Value == '':
        with_0.Value = 115200
    Debug.Print(M08.Check_If_Arduino_could_be_programmed_and_set_Board_type(COMPrtT_COL, BuildOT_COL, BuildOptions, DeviceSignature) + ' BuildOptions: ' + BuildOptions)

# VB2PY (UntranslatedCode) Option Explicit


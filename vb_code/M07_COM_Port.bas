Attribute VB_Name = "M07_COM_Port"
Option Explicit

' Revision History:
' ~~~~~~~~~~~~~~~~~
' 30.12.19: - Copied from Prog_Generator
'           - Added function ComPortPage() to be able to store the com port in one page for all sheets
' 31.12.19: - Improved the Get_USB_Ports() function because the "find" function didn't work on Norberts compuer for some reasons
' 27.03.20: - method InitComPort to be able to init the serial communication parameter
' 31.03.20: - method InitComPort: also set com timeouts
'           - support com ports > COM9
' 07.04.20: - Jürgen added functions to detect the connected Arduinos
' 04.05.20: - Speed up from 3 sec to 1 sec by
'             - Requesting only the Device Signatur
'             - Reducing the number of requested bytes in Transact()
' 21.04.21: - Add experimantal Pico Support
'           - move generic SendMLLCommand into this module


#If VBA7 Then 'For 64 Bit Systems
    Private Declare PtrSafe Function SetCommState Lib "kernel32.dll" (ByVal hCommDev As Long, ByRef lpDCB As DCB) As Long
    Private Declare PtrSafe Function GetCommState Lib "kernel32.dll" (ByVal nCid As Long, ByRef lpDCB As DCB) As Long
    Private Declare PtrSafe Function BuildCommDCB Lib "kernel32.dll" Alias "BuildCommDCBA" (ByVal lpDef As String, ByRef lpDCB As DCB) As Long
    Private Declare PtrSafe Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
    Private Declare PtrSafe Function GetCommTimeouts Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCommTimeouts As COMMTIMEOUTS) As Long
    Private Declare PtrSafe Function EscapeCommFunction Lib "kernel32.dll" (ByVal hFile As Long, ByVal dwFunc As Long) As Boolean
    Private Declare PtrSafe Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, _
        lpCC As COMMCONFIG, lpdwSize As Long) As Long
#Else 'For 32 Bit Systems
    Private Declare Function SetCommState Lib "kernel32.dll" (ByVal hCommDev As Long, ByRef lpDCB As DCB) As Long
    Private Declare Function GetCommState Lib "kernel32.dll" (ByVal nCid As Long, ByRef lpDCB As DCB) As Long
    Private Declare Function BuildCommDCB Lib "kernel32.dll" Alias "BuildCommDCBA" (ByVal lpDef As String, ByRef lpDCB As DCB) As Long
    Private Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
    Private Declare Function GetCommTimeouts Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpCommTimeouts As COMMTIMEOUTS) As Long
    Private Declare Function EscapeCommFunction Lib "kernel32.dll" (ByVal hFile As Long, ByVal dwFunc As Long) As Boolean
    Private Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
    Private Declare Function GetDefaultCommConfig Lib "kernel32" Alias "GetDefaultCommConfigA" (ByVal lpszName As String, _
        lpCC As COMMCONFIG, lpdwSize As Long) As Long
#End If

Private Type DCB
    DCBlength As Long
    BaudRate As Long
    fBitFields As Long 'See Comments Win32API.Txt
    wReserved As Integer
    XonLim As Integer
    XoffLim As Integer
    ByteSize As Byte
    Parity As Byte
    StopBits As Byte
    XonChar As Byte
    XoffChar As Byte
    ErrorChar As Byte
    EofChar As Byte
    EvtChar As Byte
    wReserved1 As Integer 'Reserved - Do Not Use
End Type

Private Type COMMCONFIG
    dwSize As Long
    wVersion As Integer
    wReserved As Integer
    dcbx As DCB
    dwProviderSubType As Long
    dwProviderOffset As Long
    dwProviderSize As Long
    wcProviderData As Byte
End Type
 

Private Type COMMTIMEOUTS
        ReadIntervalTimeout As Long
        ReadTotalTimeoutMultiplier As Long
        ReadTotalTimeoutConstant As Long
        WriteTotalTimeoutMultiplier As Long
        WriteTotalTimeoutConstant As Long
End Type
'-------------------------------------------------------------------------------
' System Constants
'-------------------------------------------------------------------------------
Private Const ERROR_IO_INCOMPLETE = 996&
Private Const ERROR_IO_PENDING = 997
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const OPEN_EXISTING = 3

Private Const COM_SETXOFF = 1  'Causes transmission to act as if an XOFF character has been received.
Private Const COM_SETXON = 2     'Causes transmission to act as if an XON character has been received.
Private Const COM_SETRTS = 3   'Sends the RTS (request-to-send) signal.
Private Const COM_CLRRTS = 4   'Clears the RTS (request-to-send) signal.
Private Const COM_SETDTR = 5   'Sends the DTR (data-terminal-ready) signal.
Private Const COM_CLRDTR = 6   'Clears the DTR (data-terminal-ready) signal.
Private Const COM_SETBREAK = 8 'Suspends character transmission and places the transmission line in a break state until the ClearCommBreak function is called (or EscapeCommFunction is called with the CLRBREAK extended function code). The SETBREAK extended function code is identical to the SetCommBreak function. Note that this extended function does not flush data that has not been transmitted.
Private Const COM_CLRBREAK = 9 'Restores character transmission and places the transmission line in a nonbreak state. The CLRBREAK extended function code is identical to the ClearCommBreak function.

' COMM Functions
Private Const MS_CTS_ON = &H10&
Private Const MS_DSR_ON = &H20&
Private Const MS_RING_ON = &H40&
Private Const MS_RLSD_ON = &H80&
Private Const PURGE_RXABORT = &H2
Private Const PURGE_RXCLEAR = &H8
Private Const PURGE_TXABORT = &H1
Private Const PURGE_TXCLEAR = &H4


#If VBA7 Then 'For 64 Bit Systems
    '
    ' Creates or opens a communications resource and returns a handle
    ' that can be used to access the resource.
    '
    Private Declare PtrSafe Function CreateFile Lib "kernel32" Alias "CreateFileA" _
        (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
    '
    ' Closes an open communications device or file handle.
    '
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    
    Private Declare PtrSafe Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, _
            ByRef Buffer As Any, ByVal nNumberOfBytesToWrite As Long, _
            ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

    Private Declare PtrSafe Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, _
            ByRef Buffer As Any, ByVal nNumberOfBytesToRead As Long, _
            ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
#Else 'For 32 Bit Systems
    '
    ' Creates or opens a communications resource and returns a handle
    ' that can be used to access the resource.
    '
    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
        (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, lpSecurityAttributes As Any, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
    '
    ' Closes an open communications device or file handle.
    '
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    
    Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, _
            ByRef Buffer As Any, ByVal nNumberOfBytesToWrite As Long, _
            ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

    Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, _
            ByRef Buffer As Any, ByVal nNumberOfBytesToRead As Long, _
            ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
#End If

Private Const Resp_STK_OK = &H10
Private Const Resp_STK_FAILED = &H11
Private Const Resp_STK_INSYNC = &H14
Private Const Sync_CRC_EOP = &H20
Private Const Cmnd_STK_GET_PARAMETER = &H41
Private Const Cmnd_STK_GET_SYNC = &H30
Private Const STK_READ_SIGN = &H75
Private Const Parm_STK_HW_VER = &H80
Private Const Parm_STK_SW_MAJOR = &H81
Private Const Parm_STK_SW_MINOR = &H82
      
'-------------------------
Private Sub Test_Get_COM()
'-------------------------
  Dim Res As String, Line As Variant
  Res = F_shellExec("cmd /c mode")
  For Each Line In Split(Res, vbCr)
      Debug.Print Line;
  Next Line
End Sub


'------------------------------------------
Private Function Get_USB_Ports() As Variant
'------------------------------------------
' The function returns an long array with COM numbers
' COM-10 is allways added because otherwise the array may be empty if no other com port is detected
' The "find" function dosen't work on Norberts computer => Therefore it's replaced by an own find algo
  Dim Res As String, Lines As Variant
  Res = F_shellExec("cmd /c mode") ' Achtung: Der Mode Befehl schickt einen Reset zu allen Ports

  If Res = "" Then ' No COM port available ?
     MsgBox Get_Language_Str("Fehler: Das Abfragen der COM Ports ist fehlgeschlagen ;-("), vbCritical, _
            Get_Language_Str("Fehler beim abfragen der COM Ports")
     EndProg
  End If
  
  Res = Replace(Res, ":", "")
  Dim Line As Variant, p As Long, ResStr As String, Cnt As Long
  ResStr = "-10 "
  For Each Line In Split(Res, vbCr)
     p = InStr(Line, "COM")
     If p > 0 Then
        ResStr = ResStr & Trim(Mid(Line, p + Len("COM"), 255)) & " "
        Cnt = Cnt + 1
     End If
  Next Line
  ResStr = DelLast(ResStr)
  Dim ResSplit As Variant
  ResSplit = Split(ResStr, " ")
  Dim ResArray() As Long, i As Long
  ReDim ResArray(Cnt)
  For i = 0 To Cnt
     ResArray(i) = val(ResSplit(i))
  Next
  Get_USB_Ports = ResArray
End Function



'-----------------------------------------------------------
Public Sub Detect_Com_Port_and_Save_Result(Right As Boolean)
'-----------------------------------------------------------
  Dim ComPortColumn As Long, BuildOptColumn As Long, Pic_ID As String
  If Right Then
        ComPortColumn = COMPrtR_COL: Pic_ID = "DCC": BuildOptColumn = BUILDOpRCOL
  Else: ComPortColumn = COMPort_COL: Pic_ID = "LED": BuildOptColumn = BUILDOP_COL
  End If
  
  Dim Port As Long
  Port = Detect_Com_Port(Right, Pic_ID)
  If Port > 0 Then
     Cells(SH_VARS_ROW, ComPortColumn) = Port
     
     StatusMsg_UserForm.Set_Label Get_Language_Str("Überprüfe den Arduino Typ") ' 30.10.20:
     StatusMsg_UserForm.Show
     Dim BuildOptions As String, DeviceSignature As Long
     Check_If_Arduino_could_be_programmed_and_set_Board_type ComPortColumn, BuildOptColumn, BuildOptions, DeviceSignature
     Unload StatusMsg_UserForm
  End If
End Sub



'UT------------------------------
Private Sub TestDetect_Com_Port()
'UT------------------------------
  Detect_Com_Port False, "LED"
End Sub


'-----------------------------------------
Public Function ComPortPage() As Worksheet                                  ' 30.12.19:
'-----------------------------------------
   If ComPortfromOnePage <> "" Then
         Set ComPortPage = Sheets(ComPortfromOnePage)
   Else: Set ComPortPage = ActiveSheet
   End If
End Function


'---------------------------------------------------------------------------
Public Function Check_USB_Port_with_Dialog(ComPortColumn As Long) As Boolean
'---------------------------------------------------------------------------
  If val(ComPortPage().Cells(SH_VARS_ROW, ComPortColumn)) <= 0 Then
        Check_USB_Port_with_Dialog = USB_Port_Dialog(ComPortColumn)           ' 04.05.20: Prior Check_USB_Port_with_Dialog ends the program in case of an error
  Else: Check_USB_Port_with_Dialog = True
  End If
  
End Function

'UT------------------------------------------
Private Sub Test_Check_USB_Port_with_Dialog()                               ' 30.12.19:
'UT------------------------------------------
  Make_sure_that_Col_Variables_match
  Debug.Print Check_USB_Port_with_Dialog(COMPort_COL)   ' Left Arduino
  'Debug.Print Check_USB_Port_with_Dialog(COMPrtR_COL)   ' Right Arduino
  'Debug.Print Check_USB_Port_with_Dialog(COMPrtT_COL)   ' Tiny_Uniprog
End Sub


'-----------------------------------------------------------------------------
Public Function Get_USB_Port_with_Dialog(Optional Right As Boolean) As Integer  ' 30.12.19:
'-----------------------------------------------------------------------------
  Dim ComPortColumn As Long
  If Right Then
        ComPortColumn = COMPrtR_COL
  Else: ComPortColumn = COMPort_COL
  End If
  
  With ComPortPage().Cells(SH_VARS_ROW, ComPortColumn)
    If val(.Value) <= 0 Then
       If USB_Port_Dialog(ComPortColumn) = False Then
          Get_USB_Port_with_Dialog = -1
          Exit Function
       End If
    End If
    Get_USB_Port_with_Dialog = val(.Value)
  End With
End Function

'-------------------------------------------------------------
Public Sub InitComPort(ByVal Port As Byte, Settings As String)
'-------------------------------------------------------------
  If NativeInitComPort(Port, Settings, 100) Then Exit Sub
  ' if previous method failed use command shell and mode command to set serial options
  F_shellRun "cmd /c mode com" & Port & " " & Settings, 0, True
End Sub

'------------------------------------------------------------------------------------------------
Public Function NativeInitComPort(ByVal Port As Byte, Settings As String, readTimeout As Integer)
'------------------------------------------------------------------------------------------------
  Dim DCB As DCB
  Dim Result, handle As Long
  handle = 0
  Result = 0
  
  On Error GoTo NativeError
  handle = CreateFile("\\.\COM" & Port, 0, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
  If handle > 0 Then
    Result = BuildCommDCB(Settings, DCB)
    If Result = 1 Then Result = SetCommState(handle, DCB)
    If Result = 1 Then
        Dim cts As COMMTIMEOUTS
        If GetCommTimeouts(handle, cts) = 1 Then
            cts.ReadIntervalTimeout = readTimeout
            cts.ReadTotalTimeoutConstant = readTimeout
            cts.ReadTotalTimeoutMultiplier = 0
            SetCommTimeouts handle, cts
            NativeInitComPort = True
        End If
    End If
  End If
        
NativeError:
  On Error GoTo 0
  If handle > 0 Then
    CloseHandle (handle)
  End If
End Function


'********************** 07.04.20: New block from Jürgen ***************************

'------------------------------------------------------------------------------------------------------------------------------
Public Function EnumComPorts(Show_Unknown As Boolean, ByRef ResNames() As String, Optional PrintDebug As Boolean = True) As Byte()
'------------------------------------------------------------------------------------------------------------------------------
' Generate a list of COM ports which have
' "CH340" or "Arduino" in there name
' Doesn't check if the COM Port is used by an other program
    Dim Ports(50) As Byte, NumberOfPorts As Byte, Names(50) As String
    Dim CountOnly As Boolean, objItem, ESP_Inst As Boolean, PICO_Inst As Boolean                ' 21.04.21: Juergen
    CountOnly = True
    NumberOfPorts = 0
    ESP_Inst = (Get_BoardTyp() = "ESP32")
    PICO_Inst = (Get_BoardTyp() = "PICO")
    For Each objItem In GetObject("winmgmts:\\.\root\CIMV2").ExecQuery("SELECT * FROM Win32_PnPEntity WHERE ClassGuid=""{4d36e978-e325-11ce-bfc1-08002be10318}""", , 48)
        If Show_Unknown Or _
           (ESP_Inst = False And PICO_Inst = False And ( _
                InStr(objItem.Caption, "CH340") > 0 Or _
                InStr(objItem.Caption, "Arduino") > 0 Or _
                InStr(objItem.Caption, "USB Serial Port") > 0)) Or _
           (ESP_Inst = True And InStr(objItem.Caption, "Silicon Labs CP210x") > 0) Or _
           (PICO_Inst = True And InStr(objItem.Path_.Path, "USB\\VID_2E8A&PID_000A\") > 0) _
            Then  ' 10.11.20: Added: "Silicon Labs CP210x" for the ESP32  02.05.20: Added: "USB Serial Port" for original Nano (Frank)
                  ' 21.04.21: Added: "Silicon Labs CP210x" for the PICO
            If PrintDebug Then Debug.Print objItem.Caption
            Dim idx1 As Integer, idx2 As Integer
            idx1 = InStr(objItem.Caption, "(COM")
            If idx1 > 0 Then
              idx2 = InStr(idx1 + 3, objItem.Caption, ")")
              If idx2 > 0 Then
                Dim portNumber As Integer
                portNumber = val(Mid(objItem.Caption, idx1 + 4, idx2 - idx1 - 4))
                Ports(NumberOfPorts) = portNumber
                Names(NumberOfPorts) = objItem.Caption
                NumberOfPorts = NumberOfPorts + 1
                If NumberOfPorts >= UBound(Ports) Then Exit For
              End If
            End If
        Else: Debug.Print "Other device (Not added to result): " & objItem.Caption ' Example: "Silicon Labs CP210x USB to UART Bridge (COM10)"
        End If
    Next
    Dim Result() As Byte
    If NumberOfPorts > 0 Then
        ReDim Result(NumberOfPorts - 1)
        ReDim ResNames(NumberOfPorts - 1)
        For idx1 = 0 To NumberOfPorts - 1
          Result(idx1) = Ports(idx1)
        Next
        
        Array_BubbleSort Result ' Sort the Ports
        
        ' find the matching names
        For idx1 = 0 To NumberOfPorts - 1
          For idx2 = 0 To NumberOfPorts - 1
              If Result(idx1) = Ports(idx2) Then
                 ResNames(idx1) = Names(idx2)
                 Exit For
              End If
          Next
        Next
        
        EnumComPorts = Result
    End If
End Function

'--------------------------------------------------------------------------
Public Function Check_If_Port_is_Available(ByVal PortNr As Byte) As Boolean
'--------------------------------------------------------------------------
  Dim Ports() As Byte, ResNames() As String
  Ports = EnumComPorts(False, ResNames)
  Check_If_Port_is_Available = Is_Contained_in_Array(PortNr, Ports)
End Function

'--------------------------------------------------------------------------------------
Public Function Check_If_Port_is_Available_And_Get_Name(ByVal PortNr As Byte) As String
'--------------------------------------------------------------------------------------
  Dim Ports() As Byte, ResNames() As String, Res As Long
  Ports = EnumComPorts(False, ResNames)
  Res = Get_Position_In_Array(PortNr, Ports)
  If Res >= 0 Then
     Check_If_Port_is_Available_And_Get_Name = ResNames(Res)
  End If
End Function



'UT------------------------------------------
Private Sub Test_Check_If_Port_is_Available()
'UT------------------------------------------
  Debug.Print Check_If_Port_is_Available(3)
End Sub

        
'UT---------------------
Private Sub TestDetect()
'UT---------------------
  Dim Ports() As Byte
  Dim Start As Variant: Start = Time
    
  Dim SWMajorVersion As Byte, SWMinorVersion As Byte, HWVersion As Byte
  Dim DeviceSignatur As Long, BaudRate As Long
  Dim i As Integer, ComPort As Variant
  Dim ComPorts() As Byte, Names() As String, Ub As Long
  ComPorts = EnumComPorts(False, Names)
  On Error GoTo IsEmpty
  Ub = UBound(ComPorts)
  On Error GoTo 0
  
  For Each ComPort In ComPorts
    For i = 1 To 2
      If i = 1 Then BaudRate = 57600 Else BaudRate = 115200
      Debug.Print "Trying COM" & ComPort & " with Baudrate " & BaudRate
      Select Case DetectArduino(ComPort, BaudRate, HWVersion, SWMajorVersion, SWMinorVersion, DeviceSignatur)
        Case 1:     Debug.Print "  Serial Port     : COM" & ComPort
                    Debug.Print "  Serial Baudrate : " & BaudRate
                    Debug.Print "  Hardware Version: " & HWVersion
                    Debug.Print "  Firmware Version: " & SWMajorVersion & "." & SWMinorVersion
                    Debug.Print "  Device signature: ";
                    If DeviceSignatur = 2004239 Then Debug.Print "ATMega328 ";
                    If DeviceSignatur = &H1E9651 Then Debug.Print "ATMega4809 ";   ' 28.10.20: Jürgen
                    Debug.Print "0x" & Right("00000" + Hex(DeviceSignatur), 6) & vbCr
                    Exit For
        Case 0:     ' Retry with other baud rate
        Case Else:  Exit For
      End Select
    Next
  Next
  Debug.Print "End"
  Debug.Print "Check duaration: " & Format(Time - Start, "hh:mm:ss")
  Exit Sub
  
IsEmpty:
  MsgBox "No Arduino detected"
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Get_Arduino_Baudrate(ByVal ComPort As Byte, Start_Baudrate As Long, ByRef DeviceSignatur As Long, ByRef FirmwareVer As String, Optional DebugPrint As Boolean) As Long  ' 28.10.20: Jürgen: Added: DeviceSignatur  30.10.20: Added: FirmwareVer
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Return  >0: Baudrate 57600/115200
'          0: if no arduino is detected
'         -1: can't open com port => used by an other program ?
'         -2: can't create com port file
'         -3: can't reset arduino
  Dim SWMajorVersion As Byte, SWMinorVersion As Byte, HWVersion As Byte
  Dim BaudRate As Long
  Select Case Start_Baudrate
     Case 1:    BaudRate = 115200
     Case 2:    BaudRate = 57600
     Case Else: BaudRate = Start_Baudrate
  End Select
  If BaudRate <> 115200 And BaudRate <> 57600 Then BaudRate = 115200
  Dim i As Integer, Res As Long, SleepTime As Long
  SleepTime = 20
  For i = 1 To 8 ' In case of an error we check each baudrate 4 times because sometimes the Baudrate is not detected if started with the wrong Baudrate   13.10.20: old:  For i = 1 To 6
      If DebugPrint Then Debug.Print "Trying COM" & ComPort & " with Baudrate " & BaudRate
      If 0 Then ' Faster                                                    ' 30.10.20: Old: 1 (Seemes to be not much slower)
            Res = DetectArduino(ComPort, BaudRate, DeviceSignatur:=DeviceSignatur, SleepTime:=SleepTime)
      Else: Res = DetectArduino(ComPort, BaudRate, HWVersion, SWMajorVersion, SWMinorVersion, DeviceSignatur:=DeviceSignatur, SleepTime:=SleepTime)
      End If
      Select Case Res
        Case 1:     ' Detected an arduino
                      If DebugPrint Then
                        Debug.Print "  Serial Port     : COM" & ComPort
                        Debug.Print "  Serial Baudrate : " & BaudRate
                        Debug.Print "  Hardware Version: " & HWVersion
                        Debug.Print "  Firmware Version: " & SWMajorVersion & "." & SWMinorVersion
                        Debug.Print "  Device signature: ";
                        If DeviceSignatur = 2004239 Then Debug.Print "ATMega328 ";
                        If DeviceSignatur = &H1E9651 Then Debug.Print "ATMega4809 ";      ' 28.10.20: Jürgen
                        Debug.Print "0x" & Right("00000" + Hex(DeviceSignatur), 6) & vbCr
                    End If
                    FirmwareVer = SWMajorVersion & "." & SWMinorVersion
                    Get_Arduino_Baudrate = BaudRate
                    Exit For
        Case 0:     ' Retry with other baud rate
        Case Else:  Get_Arduino_Baudrate = Res
                    Exit For
      End Select
      If BaudRate = 115200 Then BaudRate = 57600 Else BaudRate = 115200 ' Check again with the other baud rate
      If i = 4 Then
         SleepTime = 150   ' 13.10.20:
      End If
  Next
End Function

'UT------------------------------------
Private Sub Test_Get_Arduino_Baudrate()
'UT------------------------------------
  Dim Start As Variant: Start = Time
  Dim DeviceSignatur As Long, FirmwareVer As String
 'Debug.Print "Get_Arduino_Baudrate=" & Get_Arduino_Baudrate(6, 115200, DeviceSignatur, FirmwareVer, True)
  Debug.Print "Get_Arduino_Baudrate=" & Get_Arduino_Baudrate(6, 57600, DeviceSignatur, FirmwareVer, True)   ' Matching Baudrate 3 Sec, Not matching 6 Sek
 'Debug.Print "Get_Arduino_Baudrate=" & Get_Arduino_Baudrate(8, 115200, DeviceSignatur, FirmwareVer, True)  ' Matching Baudrate 3 Sec, Not matching 6 Sek
  Debug.Print "Check duaration: " & Format(Time - Start, "hh:mm:ss")
End Sub

'---------------------------------------------------------------------------
Public Function DetectArduino(ByVal Port As Byte, ByVal BaudRate As Long, _
    Optional ByRef HWVersion As Byte = 255, _
    Optional ByRef SWMajorVersion As Byte = 255, _
    Optional ByRef SWMinorVersion As Byte = 255, _
    Optional ByRef DeviceSignatur As Long = -1, _
    Optional Trials As Long = 3, _
    Optional PrintDebug As Boolean = True, _
    Optional SleepTime As Long = 20) As Long
'---------------------------------------------------------------------------
' protocol see application note 1AVR061 here http://ww1.microchip.com/downloads/en/Appnotes/doc2525.pdf
' Result:  1: O.K
'          0: Give up after n trials => if no arduino is detected
'         -1: can't open com port
'         -2: can't create com port file
'         -3: can't reset arduino

  Dim handle As Long, i As Integer
  handle = 0
  DetectArduino = 0
  If PrintDebug Then Debug.Print "SleepTime=" & SleepTime
  If Not NativeInitComPort(Port, "BAUD=" & BaudRate & " PARITY=N DATA=8 STOP=1 dtr=off", 500) Then
     If PrintDebug Then Debug.Print "can't open com port"
     DetectArduino = -1
     Exit Function
  End If
  On Error GoTo NativeError
  handle = CreateFile("\\.\COM" & Port, GENERIC_READ Or GENERIC_WRITE, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
  If handle <= 0 Then
     If PrintDebug Then Debug.Print "can't create com port file"
     DetectArduino = -2
     Exit Function
  End If
  
  Dim Result As Boolean
  'if PrintDebug then Debug.Print "reset arduino"
  Result = EscapeCommFunction(handle, COM_SETDTR)   ' Rising edge resets the Arduino
  If Result Then
    DoEvents
    Sleep (SleepTime)      ' 13.10.20: Old 10
    ' Bei einem meiner Arduinos funktioniert die Erkennung nicht zuverlässig mit 10 ms.    13.10.20:
    ' Mit 150ms geht es gut. Aber nur dann wenn die Konfiguration groß ist
    ' (Mainboard Test: "MB Test 1 DCC" 66 LEDs)
    ' Bei einer kleinen Konfiguration mit nur 2 LEDs geht es aber nicht mit 150 ms.
    ' Dann müssen 20 ms verwendet werden
    ' => Ich habe eine umschaltung der Wartetzeit in Get_Arduino_Baudrate() eingebaut.
    '    Jetzt werden jeweils zwei versoche mit jeder Baudrate mit 20 ms gemacht und
    '    danach 2 mit 150 ms.
    '    Damit scheint es zu Funktionieren. Getestet mit 2 verschiedenen Arduinos.
    '    Zwei mit neuem BL und einer mit altem BL.
    '  - Zusätzlich habe ic "Trials" von 5 auf 3 reduziert weil ich beobachtet habe,
    '    dass der Arduino entweder beim ersten oder zweiten Versuch erkannt wird
    '    oder gar nicht erkannt wird.
    
    Result = EscapeCommFunction(handle, COM_CLRDTR) ' Faling edge is ignored
  End If
  If Result = False Then
    If PrintDebug Then Debug.Print "can't reset arduino"   ' I never get this message ?
    CloseHandle (handle)
    DetectArduino = -3
    Exit Function
  End If
  DoEvents
  If Left(CheckCOMPort_Txt, 18) = "Arduino NANO Every" Then ' 29.10.20: Jürgen
       DetectArduino = 1
       DeviceSignatur = 2004561   ' m4809
       HWVersion = 1
       SWMajorVersion = 1
       SWMinorVersion = 7
       If BaudRate = 115200 Then DetectArduino = 1 Else DetectArduino = 0
  Else
       Dim message() As Byte
       For i = 1 To Trials
           message = Transact(handle, Chr(Cmnd_STK_GET_SYNC), 2)
           If UBound(message) = 1 Then
              If message(0) = Resp_STK_INSYNC And message(1) = Resp_STK_OK Then
                 'if PrintDebug then Debug.Print "in sync with arduino"
                 If GetDeviceInformation(handle, HWVersion, SWMajorVersion, SWMinorVersion, DeviceSignatur) Then
                    DetectArduino = 1
                    If PrintDebug Then Debug.Print "Detected after " & i & " trials"
                    Exit For                                                     ' 13.10.20: Moved up
                 End If
                'Exit For                                                           ' 13.10.20: Old position
              End If
           End If
       Next
    End If
    If DetectArduino <> 1 Then
       If PrintDebug Then Debug.Print "Give up after " & Trials & " trials"
    End If
NativeError:
  On Error GoTo 0
  If handle > 0 Then
    CloseHandle (handle)
  End If
End Function

'-----------------------------------------------------------------------------------------------------------------------
Private Function Transact(ByVal handle As Long, message As String, Optional nNumberOfBytesToRead As Long = 10) As Byte()
'-----------------------------------------------------------------------------------------------------------------------
  Dim CbWritten, CbRead, ToSend, j  As Long
  Dim Buffer() As Byte
  message = message + Chr(Sync_CRC_EOP)
  ToSend = Len(message)
  ReDim Buffer(ToSend)
  For j = 0 To ToSend - 1
    Buffer(j) = Asc(Mid(message, j + 1, 1))
  Next
  
  'Debug.Print "nNumberOfBytesToRead = 10"
  'nNumberOfBytesToRead = 10
  
  PurgeComm handle, PURGE_TXABORT Or PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR
  If WriteFile(handle, Buffer(0), ToSend, CbWritten, 0&) = 0 Then Exit Function
  If CbWritten <> ToSend Then Exit Function
  
  Dim Response(0 To 9) As Byte
  Dim EmptyResp(0) As Byte
  Transact = EmptyResp
  Dim Rc As Long
  Do
    Rc = ReadFile(handle, Response(0), nNumberOfBytesToRead, CbRead, 0) ' Slow
    If Rc = 0 Or CbRead < 1 Then
      Exit Function
    End If
    DoEvents
    If Response(0) = Resp_STK_INSYNC And Response(CbRead - 1) = Resp_STK_OK Or Response(CbRead - 1) = Resp_STK_FAILED Then
      Dim resp() As Byte
      ReDim resp(0 To CbRead - 1)
      For j = 0 To CbRead - 1
        resp(j) = Response(j)
      Next
      Transact = resp
      Exit Function
    End If
    'Debug.Print "invalid packet"
    Exit Function
  Loop
End Function

'--------------------------------------------------------------------------------------------------------------------
Private Function GetDeviceInformation(ByVal handle As Long, _
                                      Optional ByRef HWVersion As Byte = 255, _
                                      Optional ByRef SWMajorVersion As Byte = 255, _
                                      Optional ByRef SWMinorVersion As Byte = 255, _
                                      Optional ByRef DeviceSignatur As Long = -1) As Boolean
'--------------------------------------------------------------------------------------------------------------------
' Added option to speed up by not requesting all values                     ' 04.05.20:
' 1 sec instead of 3 if only the DeviceSignatur is requested
' Attention: At least one value has to be requested
  GetDeviceInformation = False
  Dim data() As Byte, Tmp As Long
  If DeviceSignatur <> -1 Then
     data = Transact(handle, Chr(STK_READ_SIGN), 5)
     If UBound(data) <> 4 Then Exit Function
     If data(4) <> Resp_STK_OK Then Exit Function
     Tmp = data(1)
     DeviceSignatur = Tmp * 65536
     Tmp = data(2)
     DeviceSignatur = DeviceSignatur + Tmp * 256
     Tmp = data(3)
     DeviceSignatur = DeviceSignatur + Tmp
  End If
  If HWVersion <> 255 Then
     data = Transact(handle, Chr(Cmnd_STK_GET_PARAMETER) + Chr(Parm_STK_HW_VER), 3)
     If UBound(data) <> 2 Then Exit Function
     If data(2) <> Resp_STK_OK Then Exit Function
     HWVersion = data(1)
  End If
  If SWMinorVersion <> 255 Then
     data = Transact(handle, Chr(Cmnd_STK_GET_PARAMETER) + Chr(Parm_STK_SW_MAJOR), 3)
     If UBound(data) <> 2 Then Exit Function
     If data(2) <> Resp_STK_OK Then Exit Function
     SWMajorVersion = data(1)
  End If
  If SWMinorVersion <> 255 Then
     data = Transact(handle, Chr(Cmnd_STK_GET_PARAMETER) + Chr(Parm_STK_SW_MINOR), 3)
     If UBound(data) <> 2 Then Exit Function
     If data(2) <> Resp_STK_OK Then Exit Function
     SWMinorVersion = data(1)
  End If
  GetDeviceInformation = True
End Function

'-----------------------------------------------------------------------------------------------------------------------
Public Function SendMLLCommand(ByVal ComPort As Integer, message As String, UseHardwareHandshake As Boolean, ShowResult As Boolean) As Boolean
    Dim handle As Long
    Dim CbWritten, CbRead As Long
    Dim t, Repeat As Integer
    Dim by As Byte
    Dim Msg As String
    
    On Error GoTo serialError
    If UseHardwareHandshake Then
      InitComPort ComPort, "BAUD=115200 PARITY=N DATA=8 STOP=1 dtr=off octs=on"
    Else
      If Get_BoardTyp() <> "PICO" Then                                          ' 17.04.21: Juergen
        InitComPort ComPort, "BAUD=115200 PARITY=N DATA=8 STOP=1 dtr=off octs=off"
      Else
        InitComPort ComPort, "BAUD=115200 PARITY=N DATA=8 STOP=1 dtr=on octs=off"
      End If
    End If
    UseHardwareHandshake = Get_Current_Platform_Bool("UseHardwareHandshake")    ' 17.04.21: Juergen ESP and PICO don't need a delay between send chars
    
    handle = CreateFile("\\.\COM" & ComPort, GENERIC_READ Or GENERIC_WRITE, 0, ByVal 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If handle < 0 Then
        Err.Raise 1, , ""
    End If
    If Get_BoardTyp() = "PICO" Then                                          ' 17.04.21: Juergen
        EscapeCommFunction handle, COM_SETDTR                                 ' pico serial USB port needs DTR to be on
    End If
    If UseHardwareHandshake Then Repeat = 1 Else Repeat = 2     ' 03.04.20: Old: to 10
    If DEBUG_DCCSEND Then Debug.Print Format(Time, "hh.mm.ss") & " sending " & Repeat & " times to receiver: '" & message & "'" ' Debug
    ' The interrupts in the Arduino are locked while the LEDs are updatet
    ' => To avoid loosing bits maximal one byte could be send over the RS232 while the interrupts are locked
    ' The delay is calculated by:
    ' 0.9 + 0.35us + 0.3 us = 1.55us / Bit
    ' 24 Bit / LED
    ' Resttime > 50 us
    ' 256 LEDs => Delay 10 ms
    ' Send duration 11 * 10 ms = 110 ms
    
    While Repeat > 0
        Repeat = Repeat - 1
        For t = 1 To Len(message)
            by = Asc(Mid(message, t, 1))
            WriteFile handle, by, 1, CbWritten, 0&
            If UseHardwareHandshake = False Then Sleep (10)       ' 03.04.20: Added delay
        Next
    Wend
    
    ' write response(s) to the debug output
    If 1 And ShowResult Then
        Do
            by = 0
            If ReadFile(handle, by, 1, CbRead, 0) Then
                Msg = Msg + Chr(by)
            End If
        Loop While by > 0
        
        If Len(Msg) > 0 Then
            Debug.Print Format(Time, "hh.mm.ss") & " Response from Arduino:       '" & Replace(Msg, vbLf, "\n") & "'"
        End If
    End If
    'Close #1
    If handle > 0 Then
        CloseHandle (handle)
    End If
    
    SendMLLCommand = True
    Exit Function
    
serialError:
    If handle > 0 Then
        CloseHandle (handle)
    End If

    Msg = Err.Description
    If DEBUG_DCCSEND Then Debug.Print Format(Time, "hh.mm.ss") & " send to receiver failed with " & Msg
    MsgBox Get_Language_Str("Fehler beim senden an die serielle Schnittstelle COM") & ComPort & ":" & vbCr & _
           "  '" & Get_Language_Str(Msg) & "'" & vbCr & _
           Get_Language_Str("Eventuell ist der serielle Monitor noch offen"), _
           vbCritical, Get_Language_Str("Fehler: Decoder senden fehlgeschlagen")
    On Error GoTo 0
    SendMLLCommand = False
End Function
'-----------------------------------------------------------------------------------------------------------------------

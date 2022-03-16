Attribute VB_Name = "M39_Simulator"
Option Explicit

Private Type RECT
  left As Long
  top As Long
  right As Long
  bottom As Long
End Type

#If VBA7 Then 'For 64 Bit Systems
Private Declare PtrSafe Function CreateSampleConfig Lib "MobaLedLibWrapper.dll" () As LongPtr
Private Declare PtrSafe Function CreateSimulator Lib "MobaLedLibWrapper.dll" Alias "Create" ( _
        configData As Any, ByVal configLength As Integer) As LongPtr
#Else
Private Declare PtrSafe Function CreateSampleConfig Lib "MobaLedLibWrapper.dll" () As LongPtr
Private Declare PtrSafe Function CreateSimulator Lib "MobaLedLibWrapper.dll" Alias "Create" ( _
        configData As Any, ByVal configLength As Integer) As LongPtr
#End If
Private Declare Sub SetInput Lib "MobaLedLibWrapper.dll" (ByVal Channel As Byte, ByVal IsOn As Byte)
Private Declare Sub UpdateLeds Lib "MobaLedLibWrapper.dll" Alias "Update" ()
Private Declare Function GetInput Lib "MobaLedLibWrapper.dll" (ByRef Channel As Byte) As Byte
Private Declare Sub ShowLEDWindow Lib "MobaLedLibWrapper.dll" ( _
        ByVal ledsX As Byte, ByVal ledsY As Byte, ByVal ledSize As Byte, ByVal ledOffset As Byte, _
        ByVal windowPosX As Integer, ByVal windowPosY As Integer, automaticUpdate As Boolean)
Private Declare Sub CloseLEDWindow Lib "MobaLedLibWrapper.dll" ()
Private Declare Function IsLEDWindowVisible Lib "MobaLedLibWrapper.dll" () As Byte
Private Declare Function GetLEDWindowRect Lib "MobaLedLibWrapper.dll" (lpRect As RECT) As Long
Private Declare Function GetWrapperVersion Lib "MobaLedLibWrapper.dll" () As Long

Private AddressMapping As Scripting.Dictionary

Public Function IsSimualtorAvailable()
    If Dir(ThisWorkbook.Path & "\" & Cfg_Dir_LED & "MobaLedLibWrapper.dll") = "" Then Exit Function
    ChDir ThisWorkbook.Path & "\" & Cfg_Dir_LED
    On Error GoTo SimError:
    ' call any function to test binding
    Dim WrapperVersion As Long
    WrapperVersion = GetWrapperVersion
    ' major.minor.revision e.g. 10000 = 01.00.00
    ' in case the exported function interface changes you need to increase major/and or minor version
    ' if interface is compatible only revision changes
    IsSimualtorAvailable = Int(WrapperVersion / 100) = 100 ' expect version 01.00.xx
SimError:
End Function
Sub OpenSimulator()
    If Not IsSimualtorAvailable Then Exit Sub
    LoadConfiguration
    StorePosition
    Dim Res
    ShowLEDWindow _
      Get_Num_Config_Var_Range("SimLedsX", 4, 16, 8), _
      Get_Num_Config_Var_Range("SimLedsY", 4, 16, 8), _
      Get_Num_Config_Var_Range("SimLedSize", 4, 32, 24), _
      Get_Num_Config_Var_Range("SimOffset", 0, 255, 1), _
      Get_Num_Config_Var_Range("SimPosX", -16383, 16383, 800), _
      Get_Num_Config_Var_Range("SimPosY", -16383, 16383, 400), _
      Get_Num_Config_Var_Range("SimOnTop", 0, 1, 1)
End Sub

Private Sub LoadFile(FileName As String, ByRef Buffer() As Byte)
    Dim fileInt As Integer: fileInt = FreeFile
    Open FileName For Binary Access Read As #fileInt
    If LOF(fileInt) = 0 Then
        ReDim Buffer(1)
        Buffer(0) = 0
    Else
        ReDim Buffer(0 To LOF(fileInt) - 1)
        Get #fileInt, , Buffer
    End If
    Close #fileInt
End Sub

Private Sub LoadConfiguration()
    Dim Buffer() As Byte
    If Not IsSimualtorAvailable Then Exit Sub
    LoadFile ThisWorkbook.Path & "\" & Cfg_Dir_LED & "LEDConfig.bin", Buffer
    Call CreateSimulator(Buffer(0), UBound(Buffer) - LBound(Buffer) + 1)
    LoadFile ThisWorkbook.Path & "\" & Cfg_Dir_LED & "AddressConfig.bin", Buffer
    Dim Index As Integer, Address As Long, ChannelAndInCntAndType As Long, Length As Integer, Value As Long
    Length = UBound(Buffer) - LBound(Buffer) + 1
    Set AddressMapping = New Scripting.Dictionary
    
    If Length >= 3 Then
        For Index = 0 To (Length / 3) - 1
            Address = Buffer(Index * 3 + 1)
            Address = 256 * Address + Buffer(Index * 3)
            AddressMapping(Address) = Buffer(Index * 3 + 2)
        Next
    End If
End Sub
Public Sub ReloadConfiguration()
    If IsSimualtorAvailable Then
        If IsLEDWindowVisible Then
            LoadConfiguration
        End If
    End If
End Sub

Public Sub UpdateSimulatorIfNeeded()
    If IsSimualtorAvailable Then
        If IsLEDWindowVisible Then
            UploadToSimulator
        End If
    End If
End Sub

Public Sub CloseSimulator()
    If IsSimualtorAvailable Then
        StorePosition
        CloseLEDWindow
    End If
End Sub

Public Function IsSimulatorActive() As Boolean
    If IsSimualtorAvailable Then
        IsSimulatorActive = IsLEDWindowVisible
    End If
End Function

Public Sub ToggleSimulator()
    If IsSimualtorAvailable Then
        If IsLEDWindowVisible Then
            CloseSimulator
        Else
            OpenSimulator
            Sleep (50)
            ' set focus back to excel
            AppActivate ActiveWorkbook.Windows(1).Caption
        End If
    End If
End Sub

Public Function SendToSimulator(ByVal Addr As Integer, ByVal Direction As Byte) As Boolean
    If Not IsSimualtorAvailable Then Exit Function
    If AddressMapping Is Nothing Then Exit Function
    If IsLEDWindowVisible Then
        Dim Channel As Byte, AddressType As Byte, CurrentType As Byte, Address As Integer, InCnt As Byte, Index As Integer, Index2 As Integer
        Channel = 0
        For Index = 0 To AddressMapping.Count - 1
            Address = AddressMapping.Keys(Index) And 16383
            AddressType = ((AddressMapping.Keys(Index) And 49152) / 16385)
            InCnt = AddressMapping.Items(Index)
            CurrentType = AddressType
            For Index2 = 1 To InCnt
                If Address = Addr Then
                    If AddressType = 0 Then
                        SetInput Channel, Direction
                        SendToSimulator = True
                        Exit Function
                    ElseIf CurrentType = 1 And Direction = 0 Then  'RED
                        SetInput Channel, 1
                        Sleep (400)
                        SetInput Channel, 0
                        SendToSimulator = True
                        Exit Function
                    ElseIf CurrentType = 2 And Direction = 1 Then  'GREEN
                        SetInput Channel, 1
                        Sleep (400)
                        SetInput Channel, 0
                        SendToSimulator = True
                        Exit Function
                    End If
                End If
                Select Case CurrentType
                    Case 0:
                        Address = Address + 1
                    Case 1:     ' red
                       CurrentType = 2
                    Case 2:     ' green
                        CurrentType = 1
                        Address = Address + 1
                End Select
                Channel = Channel + 1
            Next
        Next
    End If
End Function

Public Function UploadToSimulator() As Boolean
    If Not IsSimualtorAvailable Then Exit Function
    UploadToSimulator = Create_HeaderFile(True)
    If UploadToSimulator Then
        Dim CommandStr As String
        CommandStr = ThisWorkbook.Path & "\" & Cfg_Dir_LED & CfgBuild_Script
        If Create_Compile_Script Then
            UploadToSimulator = ShellAndWait(CommandStr, 0, vbNormalFocus, PromptUser) = Success
            If UploadToSimulator Then
                OpenSimulator
            End If
        End If
        AppActivate ActiveWorkbook.Windows(1).Caption
    End If
End Function


Private Sub StorePosition()
    Dim currentPos As RECT
    If GetLEDWindowRect(currentPos) = 1 Then
        Set_String_Config_Var "SimPosX", Str(currentPos.left)
        Set_String_Config_Var "SimPosY", Str(currentPos.top)
    End If
End Sub

Private Sub ConfigurationToFile()
    Dim fName As String, fName2 As String, Line As String, fn As Integer, fn2 As Integer, i As Integer, OutputFilename As String
    
    fName = Environ(Env_USERPROFILE) & "\AppData\Local\Temp\MobaLedLib_build\ATMega\LEDs_AutoProg.ino.elf.txt"
    fName2 = Environ(Env_USERPROFILE) & "\AppData\Local\Temp\MobaLedLib_build\ATMega\LEDs_AutoProg.ino.with_bootloader.bin"
    
    If Dir(fName) = "" Then Exit Sub
    If Dir(fName2) = "" Then Exit Sub
    
    fn = FreeFile()
    Open fName For Input As fn
    Do While Not EOF(fn)
        Line Input #fn, Line
        i = InStrRev(Line, " ")
        If i > 0 Then
            Dim TypeName As String
            TypeName = Mid(Line, i + 1)
            
            If TypeName = "_ZL6Config" Or TypeName = "_ZL8Ext_Addr" Then
                Dim Splits
                Line = Replace(Line, "  ", " ")
                Line = Replace(Line, "  ", " ")
                Line = Replace(Line, "  ", " ")
                Splits = Split(Line)
                
                Dim Offset As Integer, Length As Integer
                Offset = CInt("&h0" & Splits(2))
                Length = val(Splits(3))
                fn2 = FreeFile()
                Open fName2 For Binary As fn2
                Seek #fn2, Offset + 1
                
                Dim Buffer() As Byte
                ReDim Buffer(Length - 1)
                Get #fn2, , Buffer
                Close #fn2
                fn2 = FreeFile()
                
                If TypeName = "_ZL6Config" Then
                    OutputFilename = ThisWorkbook.Path & "\LEDs_AutoProg\LEDConfig.bin"
                ElseIf TypeName = "_ZL8Ext_Addr" Then
                    OutputFilename = ThisWorkbook.Path & "\LEDs_AutoProg\DCCConfig.bin"
                End If
                
                If Dir(OutputFilename) <> "" Then Kill (OutputFilename)
                Open OutputFilename For Binary As fn2
                Put #fn2, , Buffer
                Close #fn2
            End If
        End If
    Loop
    Close #fn
    ReloadConfiguration
End Sub

Private Function Create_Compile_Script() As Boolean
  Dim Name As String, fp As Integer
  Name = ThisWorkbook.Path & "\" & Cfg_Dir_LED & CfgBuild_Script
  fp = FreeFile
  On Error GoTo WriteError
  Open Name For Output As #fp
  Print #fp, "@ECHO OFF"
  Print #fp, "REM This file was generated by '" & ThisWorkbook.Name & "'  " & Time
    
  If Win10_or_newer() Then                                              ' 20.06.20: The find command dosn't work with this code page at Win7 for some reasons. It waits endless ?!?
     Print #fp, "CHCP 65001 >NUL" ' change the code page to show the correct german umlauts (ä,ö,ü, ...)
  End If
  Print #fp, ""
  Print #fp, "color 79"
  Print #fp, "set scriptDir=%~d0%~p0"
  Print #fp, "%~d0"
  Print #fp, "cd ""%~p0"""
  Print #fp, "set packagePath=%USERPROFILE%\" & AppLoc_Ardu
  Print #fp, "set toolPath=" & FilePath(Find_ArduinoExe) & "hardware\tools\avr\bin"
  Print #fp, "set platformPath=" & FilePath(Find_ArduinoExe) & "hardware\arduino\avr"
  Print #fp, "set libraryPath=" & GetShortPath(DelLast(Get_Ardu_LibDir()))
  Print #fp, """%toolPath%/avr-g++"" -c -g -Os -Wall -std=gnu++11 -fpermissive -fno-exceptions -ffunction-sections -fdata-sections -fno-threadsafe-statics -Wno-error=narrowing -MMD -flto -mmcu=atmega328p -DF_CPU=16000000L -DARDUINO=10813 -DARDUINO_AVR_NANO -DARDUINO_ARCH_AVR ""-I%platformPath%\cores\\arduino"" ""-I%platformPath%\variants\\eightanaloginputs""  ""-I%libraryPath%\MobaLedLib\src""  ""%scriptDir%Configuration.cpp"" -o ""%scriptDir%Configuration.cpp.o"""
  Print #fp, "if errorlevel 1 goto :error"
  Print #fp, """%toolPath%/avr-gcc"" -Wall -Os -g -flto -fuse-linker-plugin -Wl,--gc-sections -mmcu=atmega328p -o ""%scriptDir%Configuration.elf"" ""%scriptDir%Configuration.cpp.o"" -lm"
  Print #fp, "if errorlevel 1 goto :error"
  Print #fp, ""
  Print #fp, "Rem ""%toolPath%/avr-readelf"" -a ""%scriptDir%Configuration.elf"" >""%scriptDir%Configuration.elf.txt"""
  Print #fp, ""
  Print #fp, """%toolPath%/avr-objcopy"" -O binary -j .MLLLedConfig ""%scriptDir%Configuration.elf"" ""%scriptDir%LEDConfig.bin"""
  Print #fp, "if errorlevel 1 goto :error"
  Print #fp, """%toolPath%/avr-objcopy"" -O binary -j .MLLAddressConfig ""%scriptDir%Configuration.elf"" ""%scriptDir%AddressConfig.bin"""
  Print #fp, "if errorlevel 1 goto :error"
  Print #fp, ""
  Print #fp, "goto :eof"
  Print #fp, ""
  Print #fp, ":error"
  Print #fp, "   COLOR 4F"
  Print #fp, "   ECHO   ****************************************"
  Print #fp, "   ECHO    Da ist was schief gegangen ;-("
  Print #fp, "   ECHO   ****************************************"
  Print #fp, "   Pause"
  Close #fp
  Create_Compile_Script = True
  Exit Function
  
WriteError:
  Close #fp
  MsgBox Get_Language_Str("Fehler beim Schreiben der Datei '") & Name & "'", vbCritical, Get_Language_Str("Fehler beim erzeugen der Arduino Start Datei")
  Create_Compile_Script = False
End Function


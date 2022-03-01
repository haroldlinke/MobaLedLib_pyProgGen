Attribute VB_Name = "M32_DCC"
' Revision History:
' ~~~~~~~~~~~~~~~~~
' 27.03.20: - Initial release: Handle DCCSend button actions and send DCC accessory commands to receiver arduino
' 31.03.20: - add toggle button actions
' 02.04.20: - add optional output of serial messages from receiver arduino
'           - if multiple test buttons use same address toggle all related buttons
' 17.04.20: - add option to use hardware handshake when communicating with LED Arduino to speed up communication
' 18.04.21: - move serial port specific code to M07_COM_Port

Option Explicit

Private Sub DCCSend()
    Dim callerName As String
    Dim Target As Excel.Range
    Dim Button As Object
    On Error Resume Next
    Set Button = ActiveSheet.Shapes(Application.caller)
    callerName = Button.Name
    On Error GoTo 0
    If callerName = "" Then Exit Sub
    If DEBUG_DCCSEND Then Debug.Print Format(Time, "hh.mm.ss") & " click on button " & callerName
    Dim Addr As Integer
    Dim Direction As Byte
    
    'Debug.Print callerName ' Debug
    Addr = val(Mid(callerName, 2, 4))
    Addr = Addr - val(Get_String_Config_Var("DCC_Offset"))                          ' 21.03.20: Juergen
    Direction = val(Mid(callerName, 7, 2))
    If SendDCCAccessoryCommand(Addr, Direction) Then
        For Each Button In ActiveSheet.Shapes
            'Debug.Print Button.Name
            If Button.Name = callerName And Button.AlternativeText <> "" Then
                Dim Tmp As String
                Tmp = Button.Name
                If DEBUG_DCCSEND Then Debug.Print Format(Time, "hh.mm.ss") & _
                    " change from" & vbCrLf & Button.Name & " to" & vbCrLf & Button.AlternativeText
                Button.Name = Button.AlternativeText
                Button.AlternativeText = Tmp
                Button.TextFrame2.TextRange.Text = Mid(Button.Name, 13, 1)
                Button.Fill.ForeColor.rgb = GetButtonColor(val(Mid(Button.Name, 10, 2)))
            End If
        Next
    End If
End Sub

'-------------------------------------------------------------------------------------------------
Public Function SendDCCAccessoryCommand(ByVal Addr As Integer, ByVal Direction As Byte) As Boolean
'-------------------------------------------------------------------------------------------------
    Dim ComPort As Integer
    Dim Output_Buffer As String
    Dim UseHardwareHandshake As Boolean                 ' 17.04.20 new feature (Jürgen)
    ' to be able to use the hardware handshake the Arduino Nano module must be modified
    ' connect A1 pin to pin 9 (CTS) of the CH340 chip
    ' now you may use hardware handshake using CTR flow control
    UseHardwareHandshake = False
    
    Make_sure_that_Col_Variables_match
    If Check_USB_Port_with_Dialog(COMPort_COL) = False Then Exit Function   ' 04.05.20: Added exit (Prior Check_USB_Port_with_Dialog ends the program in case of an error)
    
    If (Addr < 1 Or Addr > 9999) Then
        MsgBox Get_Language_Str("Die Adresse muss im Bereich 1 bis 9999 liegen"), vbCritical, Get_Language_Str("Fehler: Decoder senden fehlgeschlagen")
        Exit Function
    End If
    
    ComPort = Cells(SH_VARS_ROW, COMPort_COL)
    Output_Buffer = "@" & Left(Addr & "   ", 4) & " " & Left(Direction & " ", 2) & " 01" + Chr$(10) ' 29.04.20: Incremented length to be able to use 4 digits 1000-9999
    SendDCCAccessoryCommand = SendMLLCommand(ComPort, Output_Buffer, UseHardwareHandshake, DEBUG_DCCSEND)
End Function



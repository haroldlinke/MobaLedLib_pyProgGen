Attribute VB_Name = "M06_Write_Header"
Option Explicit

' Todo:
' ~~~~~
' - Wichtig: Die Schalter müssen auch in den Inputs der Macros erkannt werden sonst werden sie nur dann definiert wenn
'   sie in der Adress Spalte stehen
' - Warnung generieren wenn Switch C und D Gleichzeitig verwendet werden und die gleichen Pins verwendet werden für C und D


Public InChTxt As String              ' List of all defined DCC (SX or CAN) input channels in the form: "#defines INCH_DCC_1_ONOFF <Nr>"
Public Undefined_Input_Var As String  ' List of all undefined input variables in the first step
Public Undef_Input_Var_Row As String

Private AddrList() As Long
Private LocInChNr As Long
Private CurrentCounterId As Long        ' 01.05.20: Jürgen
Private Ext_AddrTxt As String
Private ConfigTxt As String
Private Err As String
Private Channel As Long
Private LEDNr As Long
Private AddrComment As String
Private Start_Values As String
Private Store_ValuesTxt As String       ' 01.05.20: Jürgen
Private Store_Val_Written As String     ' 01.05.20: Jürgen

Public Start_LED_Channel(LED_CHANNELS - 1) As Long                          ' 26.04.20:
Public LEDs_per_Channel(LED_CHANNELS - 1) As Long
Public Max_Channels As Long                                                 ' 13.03.21 Juergen - for new Farbtest initialisation
Public LEDs_per_ChannelList As String                                       ' 13.03.21 Juergen - for new Farbtest initialisation

Private DayAndNightTimer As String

Private Const MINLEDs = 20                                                  ' 26.10.20:


'-------------------------------------------------------
Private Function Init_HeaderFile_Generation() As Boolean
'-------------------------------------------------------
  Erase AddrList
  Make_sure_that_Col_Variables_match
  LocInChNr = 0
  CurrentCounterId = 0        ' 01.05.20: Jürgen
  Ext_AddrTxt = ""
  Store_ValuesTxt = ""        ' 01.05.20: Jürgen
  Store_Val_Written = ""      ' 01.05.20: Jürgen
  InChTxt = ""
  ConfigTxt = ""
  Err = ""
  Channel = 0
  LEDNr = 0
  AddrComment = ""
  Start_Values = ""
  Undefined_Input_Var = ""
  DayAndNightTimer = ""
  
  If Init_HeaderFile_Generation_SW() = False Then Exit Function
  
  If Init_HeaderFile_Generation_LED2Var() = False Then Exit Function        ' 08.10.20:
  
  If Init_HeaderFile_Generation_Sound() = False Then Exit Function          ' 08.10.21: Juergen
  
  If Init_HeaderFile_Generation_Extension() = False Then Exit Function      ' 31.01.22: Juergen

  ' Fill the array Start_LED_Channel()
  Dim ReserveLeds As Long, NumLeds As Long, Nr As Long
  For Nr = 0 To LED_CHANNELS - 1
      NumLeds = NumLeds + Cells(SH_VARS_ROW, Get_LED_Nr_Column(Nr))
  Next Nr
  If NumLeds < MINLEDs Then  ' To be able to test at least 20 LEDs with the color test program ' 26.10.20:
     ReserveLeds = MINLEDs - NumLeds
     NumLeds = MINLEDs
  End If
  Start_LED_Channel(0) = 0
  For Nr = 0 To LED_CHANNELS - 1
     LEDs_per_Channel(Nr) = Cells(SH_VARS_ROW, Get_LED_Nr_Column(Nr))
     If Nr = 0 And ReserveLeds > 0 Then LEDs_per_Channel(0) = LEDs_per_Channel(0) + ReserveLeds ' To be able to test at least 20 LEDs with the color test program ' 26.10.20:
     If Nr > 0 Then
        Start_LED_Channel(Nr) = Start_LED_Channel(Nr - 1) + LEDs_per_Channel(Nr - 1)
     End If
  Next Nr
  For Nr = LED_CHANNELS - 1 To 0 Step -1
    If LEDs_per_Channel(Nr) > 0 Then
        Max_Channels = Nr
        Exit For
    End If
  Next Nr
  LEDs_per_ChannelList = ""
  For Nr = 0 To Max_Channels
    LEDs_per_ChannelList = LEDs_per_ChannelList & LEDs_per_Channel(Nr)
    If Nr <> Max_Channels Then LEDs_per_ChannelList = LEDs_per_ChannelList & ","
  Next Nr
  
  Init_HeaderFile_Generation = True
End Function

'------------------------------------------------------
Private Function AddressExists(Addr As Long) As Boolean
'------------------------------------------------------
  Dim a As Variant
  ' ToDo: Überlappungen prüfen wenn InCnt > 1
  If Not IsArrayEmpty(AddrList) Then
        For Each a In AddrList
            If a = Addr Then
               AddressExists = True
               Exit Function
            End If
        Next a
        ReDim Preserve AddrList(UBound(AddrList) + 1)
  Else: ReDim Preserve AddrList(0)
  End If
  
  AddrList(UBound(AddrList)) = Addr
End Function

'-------------------------------------------------------------------------------------
Private Function AddressRangeExists(Addr As Long, cnt As Long, ByVal InpTyp As String)
'-------------------------------------------------------------------------------------
' If the InpTyp is a button (Red / Green) two virtual adresses are used.
' One for each button.
' For OnOff switches one address is used twice
' To destinguish the two cases the address is multiplied by 2 and 0/1 is added
'
  Set_Tast_Txt_Var  ' Set the global variables Red_T, Green_T, ...          06.03.20:
  
  Dim Ad As Long, InpTypMod As Long, i As Long
  Select Case InpTyp
    Case Red_T:   InpTypMod = 1
    Case Green_T: InpTypMod = 2
    Case OnOff_T: InpTypMod = 3
    Case Tast_T:  InpTypMod = 3
    Case Else: MsgBox "Internal Error: Unknown InpTyp in AddressRangeExists", vbCritical
               EndProg
  End Select
  Ad = Addr
  For i = 1 To cnt
      If InpTypMod And 1 Then
         If AddressExists(Ad * 2) Then
            AddressRangeExists = True
            Exit Function
         End If
      End If
      If InpTypMod And 2 Then
         If AddressExists(Ad * 2 + 1) Then
            AddressRangeExists = True
            Exit Function
         End If
      End If
      
      Select Case InpTypMod
         Case 1: InpTypMod = 2
         Case 2: InpTypMod = 1: Ad = Ad + 1
         Case 3: Ad = Ad + 1
      End Select
  Next i
End Function

'---------------------------------------------------------------
Public Function Get_Next_Typ(ByVal Inp_Typ As String) As String
'---------------------------------------------------------------
  Set_Tast_Txt_Var  ' Set the global variables Red_T, Green_T, ...          06.03.20:
  
  Select Case Inp_Typ
     Case OnOff_T: Get_Next_Typ = OnOff_T
     Case Red_T:   Get_Next_Typ = Green_T
     Case Green_T: Get_Next_Typ = Red_T
     Case Tast_T:  Get_Next_Typ = Tast_T
     Case Else:    MsgBox "Internal error: Undefined Inp_Typ: '" & Inp_Typ & "' in Get_Next_Typ()", vbCritical, "Internal error in Get_Next_Typ()"
                   EndProg
  End Select
End Function

'----------------------------------------------------------------------------------
Private Function Gen_Address_Define_Name(ByVal Addr As Long, ByVal InTyp As String)
'----------------------------------------------------------------------------------
  If Page_ID = "Selectrix" Then
        Gen_Address_Define_Name = "INCH_SX_" & Int(Addr / 8) & "_" & (Addr Mod 8) + 1 & Replace(Mid(Get_Typ_Const(InTyp), 2, 255), ",", "")
  Else: Gen_Address_Define_Name = "INCH_" & Page_ID & "_" & Addr & Replace(Mid(Get_Typ_Const(InTyp), 2, 255), ",", "")
  End If
End Function

'-------------------------------------------------------------------------------------------------------------------------------
Private Function Generate_Define_Line(ByVal Addr As Long, Row As Long, ByVal Channel As Long, ByVal Comment As String) As String
'-------------------------------------------------------------------------------------------------------------------------------
' Generate defines for the input channels for expert users
Const COMMENT_DEFINE = "   // "
 Dim Name As String, i As Long, InTyp As String
 InTyp = Cells(Row, Inp_Typ_Col)
   
 Set_Tast_Txt_Var  ' Set the global variables Red_T, Green_T, ...          06.03.20:

 For i = 1 To Cells(Row, InCnt___Col)
    'Name = "INCH_" & Page_ID & "_" & Addr & Replace(Mid(Get_Typ_Const(InTyp), 2, 255), ",", "")
    Name = Gen_Address_Define_Name(Addr, InTyp)
    Generate_Define_Line = Generate_Define_Line & "#define " & AddSpaceToLen(Name, 22) & "  " & AddSpaceToLen(Channel, 4) & COMMENT_DEFINE & Comment & vbCr
    If InTyp <> Red_T Then Addr = Addr + 1
    InTyp = Get_Next_Typ(InTyp)
    Channel = Channel + 1
    Comment = "    """
 Next i
End Function

'----------------------------------------------------
Private Function Get_Description(r As Long) As String
'----------------------------------------------------
  Get_Description = Trim(Cells(r, Descrip_Col))
  ' If Get_Description = "" Then Get_Description = Cells(r, Config__Col) ' 02.03.20: Old: Why should the macro be repeated? It gets verry long if "#define" lines are used
  If Get_Description = "" Then Get_Description = "Excel row " & r        ' 02.03.20: New
  Get_Description = Replace(Get_Description, vbLf, "| ")
End Function


'--------------------------------------------------------
Private Function Activate_DayAndNightTimer(Cmd As String)
'--------------------------------------------------------
  Dim Args() As String, Period As Double
  Args = Split(Trim(Replace(Replace(Cmd, "DayAndNightTimer(", ""), ")", "")), ",")
  Period = val(Trim(Args(1)))
  DayAndNightTimer = vbCr & "#define DayAndNightTimer_Period    " & Round(Period * 60 * 1000 / 512, 0) & vbCr
  
  If Trim(Args(0)) <> "SI_1" Then DayAndNightTimer = DayAndNightTimer & "#define DayAndNightTimer_InCh      " & Trim(Args(0)) & vbCr
  
  Activate_DayAndNightTimer = True
End Function

'----------------------------------------------------------------------------------------
Private Function Do_Replace_Sym_Pin_Name(ByVal Cmd As String, PinStr As String) As String      ' 30.10.20:
'----------------------------------------------------------------------------------------
  Dim MB_LED_Nr_Str_Arr() As String: MB_LED_Nr_Str_Arr = Split(MB_LED_NR_STR, " ")
  Dim MB_LED_Pin_Nr_Arr() As String: MB_LED_Pin_Nr_Arr = Split(Replace(MB_LED_PIN_NR, "  ", " "), " ")
  If UBound(MB_LED_Nr_Str_Arr) <> UBound(MB_LED_Pin_Nr_Arr) Then
     MsgBox "Internal Error: Array hafe different size in 'Do_Replace_Sym_Pin_Name()'", vbCritical, "Internal Error"
     EndProg
  End If
  Dim i As Long
  For i = 0 To UBound(MB_LED_Nr_Str_Arr)
     If PinStr = MB_LED_Nr_Str_Arr(i) Then
        Do_Replace_Sym_Pin_Name = Replace(Cmd, "(" & MB_LED_Nr_Str_Arr(i) & ",", "(" & MB_LED_Pin_Nr_Arr(i) & ",")
        Exit Function
     End If
  Next i
  MsgBox "Internal Error: PinStr not found in 'Do_Replace_Sym_Pin_Name()'", vbCritical, "Internal Error"
  EndProg
End Function

'-------------------------------------------------------------------------------------
Private Function Proc_Special_Functions(ByRef Cmd As String, LEDNr As Long, Channel As String) As Boolean  ' 01.10.20:
'-------------------------------------------------------------------------------------
  If left(Cmd, Len("Mainboard_LED(")) = "Mainboard_LED(" Then
     Dim PinOrNr As String, Replace_Sym_Pin_Name As Boolean
     PinOrNr = Split(Replace(Cmd, "Mainboard_LED(", ""), ",")(0)
     If InStr(" " & MB_LED_NR_STR & " ", " " & PinOrNr & " ") > 0 Then Replace_Sym_Pin_Name = True
     
     If (PinOrNr = "4" Or PinOrNr = "D4") And PIN_A3_Is_Used() Then ' Problem if the SwitchB or C is used. If the CAN mode is used the Mainboard LED may be used. It will overwrite the HB
        MsgBox Get_Language_Str("Achtung: Die Mainboard LED 4 kann nicht benutzt werden wenn der PIN A3 an anderer Stelle benutzt wird (CAN, SwitchB oder SwitchC)."), vbCritical, Get_Language_Str("Pin A3 ist bereits benutzt")
        Exit Function
     End If
     
     If PinOrNr = "13" Or PinOrNr = "D13" Then ' The heartbeat LED can't be used if Mainboard LED 13 is used. Attention: Don't use the Mainboard LEDs 10-13 together with the CAN ' 09.10.20:
        Cmd = Cmd & vbCr & "  #undef  LED_HEARTBEAT_PIN  /* Use the heartbeat LED at pin A3 */" & _
                    vbCr & "  #define LED_HEARTBEAT_PIN A3" & Space(79)
     End If
     If Replace_Sym_Pin_Name Then Cmd = Do_Replace_Sym_Pin_Name(Cmd, PinOrNr)
        
     Cmd = Replace(Replace(Replace(Cmd, "Mainboard_LED(", "#define Mainboard_LED"), ",", " "), ")", "")
  End If
  If left(Cmd, Len("DayAndNightTimer(")) = "DayAndNightTimer(" Then
     If Not Activate_DayAndNightTimer(Cmd) Then Exit Function
     Cmd = "// " & Cmd
  End If
  
  If InStr(Cmd, SF_LED_TO_VAR) > 0 Then                                     ' 08.10.20:
     If Not Add_LED2Var_Entry(Cmd, LEDNr) Then Exit Function
  End If
  If InStr(Cmd, SF_SERIAL_SOUND_PIN) > 0 Then                               ' 08.10.21: Juergen
     If Not Add_SoundPin_Entry(Cmd, LEDNr) Then Exit Function
  End If
  Proc_Special_Functions = True
End Function



'-----------------------------------------------------------------------------------------------------------------------------------
Private Function Generate_Config_Line(LEDNr As Long, ByVal Channel_or_define As String, r As Long, Config_Col As Long, Addr As Long) As String
'-----------------------------------------------------------------------------------------------------------------------------------
' ToDo: Add checks like
' - open/closing braket test
' - characters after #LED, #InCh
  Dim Txt As String, lines As Variant, Line As Variant, Res As String, AddDescription As Boolean, Description As String, Inc_LocInChNr As Boolean
  Txt = Cells(r, Config_Col)
  If Trim(Txt) = "" Then Exit Function
  lines = Split(Txt, vbLf)
  Description = Get_Description(r)
  AddDescription = Description <> ""
  For Each Line In lines ' Multiple lines in one cell are possible
      Dim CommentStart As Long, Cmd As String, Comment As String
      Comment = ""
      Cmd = ""
      CommentStart = InStr(Line, "//")
      If CommentStart = 0 Then
         Cmd = Line
      ElseIf CommentStart = 1 Then
         Comment = Line
      Else
         Cmd = left(Line, CommentStart)
         Comment = Mid(Line, CommentStart + 1, 1000)
      End If
      
      If LEDNr < 0 Then                                                     ' 16.11.20:
         Cells(r, Config__Col).Select
         MsgBox Get_Language_Str("Fehler: Die LED Nummer darf nicht negativ werden. Das kann durch eine falsche Angabe bei einem vorangegangenen ""Next_LED"" Befehl passieren."), vbCritical, Get_Language_Str("Fehler: Negative LED Nummer")
         Generate_Config_Line = "#ERROR#"
         Exit Function
      End If
      
      Cmd = Replace(Cmd, "#LED", LEDNr)
      
      If Addr >= 0 Or Addr = -2 Then ' Valid address or INCH_ define
            Cmd = Replace(Cmd, "#InCh", Channel_or_define)
      Else: Cmd = Replace(Cmd, "#InCh", "SI_1")
      End If
      
      If InStr(Cmd, "#LocInCh") > 0 Then
         If Cells(r, LocInCh_Col) = 0 Then
            MsgBox "Interner Fehler: '#LocInCh' wird verwendet aber 'Loc InCh' ist 0 oder leer in Zeile " & r, vbCritical, "Interner Fehler"
            EndProg
         End If
         Cmd = Replace(Cmd, "#LocInCh", "LOC_INCH" & LocInChNr)
         Inc_LocInChNr = True
         ' 18.11.19:
      End If
      
      If Proc_Special_Functions(Cmd, LEDNr, Channel_or_define) = False Then           ' 01.10.20: 08.10.21: Juergen
         Generate_Config_Line = "#ERROR#"
         Cells(r, Config__Col).Select
         Exit Function
      End If
     
      If IsExtensionKey(Cmd) Then                                                      ' 31.01.22: Juergen
         If Not Add_Extension_Entry(Cmd) Then
            Generate_Config_Line = "#ERROR#"
            Cells(r, Config__Col).Select
         End If
         Exit Function
      End If
      
      If Cells(r, LEDs____Col) = SerialChannelPrefix Then                             ' 08.10.20: 08.10.21: Juergen
        If Not CheckSoundChannelDefined(LEDNr) Then
            Generate_Config_Line = "#ERROR#"
            Cells(r, Config__Col).Select
            Exit Function
        End If
      End If
      Dim Add_Backslash_to_End As Boolean
      If right(RTrim(Cmd), 1) = "\" Then ' Macro line which is continued in the folowing line => Add "\" to the end  ' 25.11.19
            Add_Backslash_to_End = True
            Cmd = RTrim(Cmd)
            Cmd = left(Cmd, Len(Cmd) - 1)
      Else: Add_Backslash_to_End = False ' Don't add a '\' to all following lines ' 02.03.20:
      End If
      Cmd = "  " & Cmd & Comment
      If AddDescription Then ' The description is only added to the first line
                                    Cmd = AddSpaceToLen(Cmd, 109) & " /* " & Description
      ElseIf Description <> "" Then
                                    Cmd = AddSpaceToLen(Cmd, 109) & " /*     """
      End If
      Cmd = AddSpaceToLen(Cmd, 300) & " */"
      If Add_Backslash_to_End Then Cmd = Cmd & " \"                         ' 25.11.19:
      
      AddDescription = False
      Res = Res & Cmd & vbCr
  Next Line

    ' Added by Misha 29-03-2020                                             ' 14.06.20: Added from Mishas version
    ' Changed by Misha 20-04-2020
    If InStr(left(Res, InStr(Res, ")")), "Multiplexer") > 0 Then
        Res = vbCrLf & Get_Multiplexer_Group(Res, Description, r) & vbCrLf
    End If
    ' End Changes by Misha
  
  If Inc_LocInChNr Then LocInChNr = LocInChNr + Cells(r, LocInCh_Col)       ' 18.11.19: Moved down
  
  Generate_Config_Line = Res
End Function

'----------------------------------------------------------------
Private Function Get_Typ_Const(ByVal Inp_Typ As String) As String
'----------------------------------------------------------------
  Set_Tast_Txt_Var  ' Set the global variables Red_T, Green_T, ...          06.03.20:
  
  Select Case Inp_Typ
     Case OnOff_T: Get_Typ_Const = "S_ONOFF,"
     Case Red_T:   Get_Typ_Const = "B_RED,  "
     Case Green_T: Get_Typ_Const = "B_GREEN,"
     Case Tast_T:  Get_Typ_Const = "B_TAST, "
     Case Else:    MsgBox "Internal error: Undefined Inp_Typ: '" & Inp_Typ & "' in Get_Typ_Const()", vbCritical, "Internal error in Get_Typ_Const()"
                   EndProg
  End Select
End Function

'------------------------------------------------
Private Sub Add_to_Err(r As Range, Txt As String)
'------------------------------------------------
  If Err = "" Then r.Select ' Marc the first error location
  Err = Err & Txt & vbCr
End Sub

'--------------------------------------------------------------------------------------------
Private Sub Add_Start_Value_Line(r As Long, Mask As Long, Pos As Long, Description As String)
'--------------------------------------------------------------------------------------------
  Start_Values = Start_Values & AddSpaceToLen("  MobaLedLib.Set_Input(" & Channel + Pos & ", 1);", 109) _
                 & " // " & Description & vbCr
End Sub


'----------------------------------------------
Private Sub Create_Start_Value_Entry(r As Long)
'----------------------------------------------
' Fill the global string "Start_Values"
  Dim sv As Long, i As Long, Mask As Long
  sv = val(Cells(r, Start_V_Col))
  If sv = 0 Then Exit Sub
  If sv < 0 Then Add_to_Err Cells(r, Start_V_Col), Get_Language_Str("Negativer Startwert in Zeile ") & r
  Dim Description As String
  Description = Get_Description(r)
  Mask = 1
  For i = 0 To val(Cells(r, InCnt___Col)) - 1
    If (sv And Mask) > 0 Then
       Add_Start_Value_Line r, Mask, i, Description
       Description = "   """
    End If
    Mask = Mask * 2
  Next
  If sv > Mask - 1 Then Add_to_Err Cells(r, Start_V_Col), Get_Language_Str("Startwert in Zeile ") & r & Get_Language_Str(" ist zu groß. Maximal möglicher Wert: ") & Mask - 1
End Sub

'-----------------------------------------------------------------------------------
Private Function Create_Header_Entry(r As Long, ByRef AddrStr As Variant) As Boolean
'-----------------------------------------------------------------------------------
' Fills the global strings
' - "ConfigTxt":    Configuration array "MobaLedLib_Configuration()"
' - "Ext_AddrTxt":  addresses for DCC, Selextrix or CAN: (Array Ext_Addr[])
' - "Start_Values": Initial values for DCC, Selextrix or CAN
' - "InChTxt":      defines like "#defines INCH_DCC_1_ONOFF " for expert user
' Calculate "Channel" = the next input channel number

Const ADDR_BORDER = "           { "
Const COMMENT_START = "      // "
Const STORE_BORDER = "           { "                                       ' 01.05.20: Jürgen
  
  Dim Comment As String
  Comment = Get_Description(r)
  Dim AddrTxt_Line As String, Inp_Typ As String, InCnt As Long, Channel_or_define As String, Addr As Long
  InCnt = val(Cells(r, InCnt___Col))                                        ' 08.10.21: avoid error is cell is empty
  
  If IsNumeric(AddrStr) Then                                                ' 03.04.20:
        Addr = val(AddrStr)
  Else: Addr = -2 ' it's a variable
        Channel_or_define = AddrStr
  End If
  
  If Addr >= 0 Then
     Dim Inp_TypR As Range: Set Inp_TypR = Cells(r, Inp_Typ_Col)
     Complete_Typ Inp_TypR, True ' Check Inp_Typ. If not valid call the dialog
     If Inp_TypR = "" Then
        Exit Function
     End If
     
     If AddressRangeExists(Addr, InCnt, Inp_TypR) Then
           Channel_or_define = Gen_Address_Define_Name(Addr, Inp_TypR)
           If InStr(InChTxt, Channel_or_define) = 0 Then
              Add_to_Err Cells(r, Inp_Typ_Col), Get_Language_Str("Die Adresse '") & Addr & Get_Language_Str("' in Zeile ") & r & Get_Language_Str(" wird bereits mit einem anderen Typ benutzt.")
           End If
           Addr = -2
     Else: Channel_or_define = Channel
     End If
     
  End If
#If True Then                                                               ' 26.04.20:
  Dim LEDs_Channel As Long
  Dim LEDs As String                                ' 10.08.21: Juergen add sound channel
  Dim ErrorMessage As String, ErrorTitle As String
  LEDs_Channel = val(Cells(r, LED_Cha_Col))
  LEDs = Cells(r, LEDs____Col)
  If Trim(LEDs) = SerialChannelPrefix Then                                  ' 10.08.21: Juergen add sound channel
     If LEDs_Channel < 0 Or LEDs_Channel >= SERIAL_CHANNELS Then
        ErrorMessage = Replace(Replace(Get_Language_Str("Fehler: Der 'Sound Kanal' in Zeile #1# ist ungültig." + vbCr + "Es sind die Sound Kanäle 0-#2# erlaubt."), "#1#", r), "#2#", Str(SERIAL_CHANNELS - 1))
        ErrorTitle = Get_Language_Str("Ungültiger Sound Kanal")
     End If
  Else
     If LEDs_Channel < 0 Or LEDs_Channel >= LED_CHANNELS Then
        ErrorMessage = Replace(Replace(Get_Language_Str("Fehler: Der 'LED Kanal' in Zeile #1# ist ungültig." + vbCr + "Es sind die Led Kanäle 0-#2# erlaubt."), "#1#", r), "#2#", Str(LED_CHANNELS - 1))
        ErrorTitle = Get_Language_Str("Ungültiger LED Channel")
     End If
  End If
  If ErrorMessage <> "" Then
     Dim OldEvents As Boolean: OldEvents = Application.EnableEvents
     Application.EnableEvents = False
     Cells(r, LED_Cha_Col).Select
     Application.EnableEvents = OldEvents
     MsgBox ErrorMessage, vbCritical, ErrorTitle
     Exit Function
  End If

  LEDNr = Get_LED_Nr(LEDNr, r, LEDs_Channel)                                ' 04.03.21 Juergen
#Else
  If Cells(r, LED_Nr__Col) <> "" Then
     LEDNr = Cells(r, LED_Nr__Col)
  End If
#End If
  ' Entry for the configuration array which contains the macros
  Dim Res As String
  Res = Generate_Config_Line(LEDNr, Channel_or_define, r, Config__Col, Addr)
  If Res = "#ERROR#" Then Exit Function
  ConfigTxt = ConfigTxt & Res
 
 
 'begin change 01.05.20: Jürgen
 Select Case GetMacroStoreType(r)
   Case MST_CTR_NONE, MST_CTR_ON, MST_CTR_OFF: CurrentCounterId = CurrentCounterId + 1
 End Select
 
  Dim storeStatusType As Integer
  Dim TextLine As String
  
  storeStatusType = Check_And_Get_Store_Status(r, Addr, Inp_TypR, Channel_or_define)
  If storeStatusType < MST_None Then Exit Function
  
  If storeStatusType > MST_None Then
    ' get lastet translated name of channel
    If Not Inp_TypR Is Nothing And Addr >= 0 Then Channel_or_define = Gen_Address_Define_Name(Addr, Inp_TypR)

    If storeStatusType = SST_S_ONOFF Or storeStatusType = SST_TRIGGER Then
      ' avoid duplicate entries
      If (InStr(Store_Val_Written, " " + Channel_or_define + " ")) = 0 Then
        If storeStatusType = SST_S_ONOFF Then
          TextLine = STORE_BORDER & "IS_TOGGLE + "
        Else
          TextLine = STORE_BORDER & "IS_PULSE  + "
        End If
        TextLine = TextLine & AddSpaceToLen(InCnt, 2) & ", "
        TextLine = TextLine & AddSpaceToLen(Channel_or_define, 20) & "},"
        TextLine = TextLine & COMMENT_START & Comment
        Store_ValuesTxt = Store_ValuesTxt & TextLine & vbCr
        Store_Val_Written = Store_Val_Written + " " + Channel_or_define + " "
      End If
    End If
  End If
  If storeStatusType = SST_COUNTER_ON Then
    
    ' diese Variante würde nur ein Byte pro Counter verwenden,
    ' allerdings is der zusätzliche code zum Behandeln der zusätlzichen Liste
    ' in den häufigsten Fällen größer als jene Bytes, die man mit dieser Variante einsparen könnte
    'TextLine = TextLine & AddSpaceToLen(CurrentCounterId, 4) & "},"
    'TextLine = TextLine & COMMENT_START & Comment
    'Store_CountersTxt = Store_CountersTxt & TextLine & vbCr
    
    TextLine = STORE_BORDER & "IS_COUNTER    , "
    TextLine = TextLine & AddSpaceToLen("COUNTER_ID " & CurrentCounterId, 20) & "},"
    TextLine = TextLine & COMMENT_START & Comment
    Store_ValuesTxt = Store_ValuesTxt & TextLine & vbCr
  End If
 'end change 01.05.20: Jürgen
  
  If Addr >= 0 Then
     ' Defines for expert users and duplicate adresses
     InChTxt = InChTxt & Generate_Define_Line(Addr, r, Channel, Comment)
    
     ' Definition of the array with the external adresses for DCC, Selecrix and CAN
     AddrTxt_Line = ADDR_BORDER & AddSpaceToLen(Addr, 5)
     AddrTxt_Line = AddrTxt_Line & "+ " & Get_Typ_Const(Inp_TypR) & " " & AddSpaceToLen(InCnt, 2) & "},"
     Ext_AddrTxt = Ext_AddrTxt & AddrTxt_Line & COMMENT_START
     If AddrComment <> "" Then Ext_AddrTxt = Ext_AddrTxt & AddSpaceToLen(AddrComment, 10)
     Ext_AddrTxt = Ext_AddrTxt & Comment & vbCr
     
     Create_Start_Value_Entry r
    
     ' Calculate the next input channel number
     With Cells(r, InCnt___Col)
        If .Value <> "" Then
           If Not IsNumeric(.Value) Or .Value < 0 Or .Value > 100 Then
                 .Select
                 MsgBox Get_Language_Str("Fehler: Eintrag '") & .Value & Get_Language_Str("' in InCnt Spalte ist ungültig"), vbCritical, Get_Language_Str("Falscher InCnt Eintrag")
                 EndProg
           Else: Channel = Channel + .Value  ' ToDo: Unterstützung für mehrere Zeilen in einer Zelle ?
           End If
        End If
     End With
  End If

  Create_Header_Entry = True
End Function

'begin change 01.05.20: Jürgen
'-----------------------------------------------------------------------------------------------------------------------------
Public Function Check_And_Get_Store_Status(r As Long, Addr As Long, Inp_TypR As Range, Channel_or_define As String) As Integer ' 01.05.20:  Jürgen
'-----------------------------------------------------------------------------------------------------------------------------
' return
' -1 for error
' 0 for Store Status not enabled
' 1 for Counter with status, Default on
' 2 for Counter with status, Default off
' 3 for Channel S_ONOFF
' 4 for Channel TRIGGER

   Check_And_Get_Store_Status = SST_NONE
   With Cells(r, Start_V_Col)
     Check_And_Get_Store_Status = Get_Store_Status(r, Addr, Inp_TypR, Channel_or_define)
     Select Case Check_And_Get_Store_Status
       Case SST_COUNTER_OFF:
         ' user forces status store to on
         If .Value = AUTOSTORE_ON Then Check_And_Get_Store_Status = SST_COUNTER_ON
         Exit Function
  
       Case SST_COUNTER_ON:
         If .Value = AUTOSTORE_OFF Or (.Value <> "" And IsNumeric(.Value)) Then Check_And_Get_Store_Status = SST_COUNTER_OFF   ' 01.05.20: Added from Mail: or IsNumeric(.value)
         Exit Function

       Case SST_S_ONOFF, SST_TRIGGER:
         If .Value = AUTOSTORE_OFF Or (.Value <> "" And IsNumeric(.Value)) Then Check_And_Get_Store_Status = SST_NONE          ' 01.05.20: Added from Mail: or IsNumeric(.value)
         Exit Function
     End Select
   
     ' user is not allow to force status store for functions that don't support this
     If .Value = AUTOSTORE_ON Then
        .Select
        MsgBox Get_Language_Str("Fehler: Eintrag '") & .Value & Get_Language_Str("' in Startwert Spalte ist ungültig"), vbCritical, "Statusspeicherung für diese Funktion nicht möglich"
        Check_And_Get_Store_Status = -1
     End If
  End With
End Function

'-------------------------------------------------------------------------------------------------------------------
Public Function Get_Store_Status(r As Long, Addr As Long, Inp_TypR As Range, Channel_or_define As String) As Integer ' 01.05.20: Jürgen
'-------------------------------------------------------------------------------------------------------------------
' return
' -1 for error
' 0 for Store Status not enabled
' 1 for Counter with status, Default on
' 2 for Counter with status, Default off
' 3 for Channel S_ONOFF
' 4 for Channel TRIGGER

   Get_Store_Status = SST_NONE
   With Cells(r, Start_V_Col)
    
    Dim Message As String
    Dim storeType  As Byte
    
    storeType = GetMacroStoreType(r)
    If storeType = MST_CTR_OFF Then
      Get_Store_Status = SST_COUNTER_OFF
      Exit Function
    End If
    If storeType = MST_CTR_ON Then
      Get_Store_Status = SST_COUNTER_ON
      Exit Function
    End If
    If storeType = MST_PREVENT_STORE Then                                   ' 01.05.20: From Mail
      Get_Store_Status = SST_NONE
      Exit Function
    End If
    Get_Store_Status = GetOnOffStoreType(r, Addr, Inp_TypR, Channel_or_define)
   End With
End Function
'end change 01.05.20: Jürgen

'---------------------------------------------------
Public Function GetMacroStoreType(r As Long) As Byte              ' 17.12.21: Jürgen Split into single line and multi line implementation
'---------------------------------------------------
    GetMacroStoreType = GetMacroStoreTypeLine(Cells(r, Config__Col))
End Function

' return
' 0 for no or undefined storage
' 1 for Button without storage
' 2 for Button with storage and default storage is on
' 3 for Button with storage and default storage is off
' 4 for function which prevents status storage

'---------------------------------------------------
Public Function GetMacroStoreTypeLine(Config_Entry As String) As Byte          ' 01.05.20: Jürgen
'---------------------------------------------------
    GetMacroStoreTypeLine = MST_None                                  ' 01.05.20: From Mail Old: GetMacroStoreType = 0

    Dim Org_Macro_Row As Long
    
    If Trim(Config_Entry) = "" Then Exit Function              ' no macro assigned
    
    Dim Parts() As String, p As Long
    Parts = Split(Config_Entry, vbLf)
    If (LBound(Parts) <> UBound(Parts)) Then
        GetMacroStoreTypeLine = GetMultilineMacroStoreType(Parts)
        Exit Function
    End If
    Parts = Split(Config_Entry, vbCr)
    If (LBound(Parts) <> UBound(Parts)) Then
        GetMacroStoreTypeLine = GetMultilineMacroStoreType(Parts)
        Exit Function
    End If
    
    Parts = Split(Config_Entry, "(")
    If Trim(Parts(0)) = "" Then Exit Function                  ' no macro assigned
    Org_Macro_Row = Find_Macro_in_Lib_Macros_Sheet(Parts(0) & "(")
    If Org_Macro_Row = 0 Then Exit Function                    ' macro not found
    
    Dim OutCntStr As String, Org_Macro As String, Org_Arguments As String
    With Sheets(LIBMACROS_SH)
       GetMacroStoreTypeLine = val(.Cells(Org_Macro_Row, SM_Type__COL))  ' 01.05.20: From Mail Old: GetMacroStoreType = Val(.Cells(Org_Macro_Row, SM_CountrCOL))
    End With
    Exit Function
End Function

'---------------------------------------------------
Public Function GetMultilineMacroStoreType(lines) As Byte             ' 17.12.21: Jürgen
'---------------------------------------------------
    ' find the first macro have a defined store type
    ' otherwise 0 = undefined
    GetMultilineMacroStoreType = MST_None
    Dim Line
    For Each Line In lines
        Dim s As String
        s = Line
        GetMultilineMacroStoreType = GetMacroStoreTypeLine(s)
        If GetMultilineMacroStoreType <> MST_None Then
            Exit Function
        End If
    Next

End Function
'---------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------
Public Function GetOnOffStoreType(r As Long, Addr As Long, Inp_TypR As Range, Channel_or_define As String) As Byte              ' 01.05.20: Jürgen
'-----------------------------------------------------------------------------------------------------------------
  GetOnOffStoreType = SST_NONE
  
  Dim TypConst As String
  If Not Inp_TypR Is Nothing Then
    If Inp_TypR <> "" Then                                                  ' 11.10.20: Prevent error message it the Userform_Selext_Typ* is aborted with ESC
       TypConst = Get_Typ_Const(Inp_TypR)
       If Addr >= 0 Then Channel_or_define = Gen_Address_Define_Name(Addr, Inp_TypR)
    End If
  Else
    TypConst = ""                                                           ' 11.10.20: Not necessary because strings are always initialiced to ""
  End If
  If Channel_or_define = "" Then Exit Function
  
    ' or all functions having Adress
  If Addr >= 0 And TypConst = "S_ONOFF," Then  ' DCC with on/off
      GetOnOffStoreType = SST_S_ONOFF
      Exit Function
  End If
  If Cells(r, InCnt___Col) > 1 And Cells(r, LED_Nr__Col) <> "" Then    ' signals with triggers
      GetOnOffStoreType = SST_TRIGGER
      Exit Function
  End If
End Function

'-----------------------------
Public Function Create_HeaderFile(Optional CreateFilesOnly As Boolean = False) As Boolean   ' 20.12.21: Jürgen add CreateFilesOnly for programatically generation of header files
'-----------------------------
' Is called if the "Z. Arduino schicken" button is pressed

  Create_HeaderFile = False                                                 ' 20.12.21: Jürgen
  Check_Version                                                             ' 21.11.21: Juergen
  Update_Start_LedNr                                                        ' 11.10.20: To prevent problems if the calculation was not called before for some reasons
  Clear_Platform_Parameter_Cache                                            ' 14.10.2021: Juergen force reload of Platofmr Paramters every time a new header is created
  Dim Ctrl_Pressed As Boolean
  Ctrl_Pressed = GetAsyncKeyState(VK_CONTROL) <> 0  ' Following function must be declared: Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
  If Ctrl_Pressed Then UserForm_Header_Created.DontShowAgain = False
  
  Make_sure_that_Col_Variables_match
  
  If Not Init_HeaderFile_Generation() Then Exit Function
  
  ' 04.03.22 Juergen: If shift key is pressed to configuration is sent to the simulator only
  If GetAsyncKeyState(VK_SHIFT) <> 0 And CreateFilesOnly = False And Get_BoardTyp() = "AM328" Then
    Create_HeaderFile = UploadToSimulator
    Exit Function
  End If
  
  Dim r As Long, sx As Boolean, SX_Ch As Long
  sx = Page_ID = "Selectrix"
  For r = FirstDat_Row To LastUsedRow
     If Not Rows(r).EntireRow.Hidden And Cells(r, Enable_Col) <> "" Then
        Dim Addr As Variant
        Addr = -1
        If Address_starts_with_a_Number(r) Then
            If sx Then ' *** Selectrix ***
                Dim Bit_P As Variant
                Bit_P = Cells(r, SX_Bitposi_Col)
                If Bit_P <> "" And val(Cells(r, InCnt___Col)) > 0 Then
                   If Cells(r, SX_Channel_Col) <> "" Then SX_Ch = Get_First_Number_of_Range(r, SX_Channel_Col) ' ToDo: SX_Ch wird nur dann aktualisiert wenn Bit pos vorhanden ist und InCnt > 0. Ist das gut ?
                   If SX_Ch >= 0 And SX_Ch <= 99 Then
                      If Bit_P >= 1 And Bit_P <= 8 Then
                         Addr = SX_Ch * 8 + Bit_P - 1
                         AddrComment = "SX " & AddSpaceToLenLeft(SX_Ch, 2) & "," & Bit_P & ": "
                      Else:   Add_to_Err Cells(r, SX_Bitposi_Col), "Wrong bitpos "" & bp & "" in row " & r
                      End If
                   Else:      Add_to_Err Cells(r, SX_Channel_Col), "Wrong SX channel in row " & r
                   End If
                End If
            Else ' *** DCC or CAN ***
                Dim MaxAddr As Long
                If Page_ID = "DCC" Then
                      MaxAddr = 10240 ' Attention some stations only support adresses up to 9999 => Don't generate a warning. The central station will generate an error
                Else: MaxAddr = 65535 ' 2048? MS2 only 320 ?
                End If
                Addr = Get_First_Number_of_Range(r, DCC_or_CAN_Add_Col)
                If Addr = "" Or val(Cells(r, InCnt___Col)) <= 0 Then
                     If Addr <> "" Then Add_to_Err Cells(r, DCC_or_CAN_Add_Col), Get_Language_Str("Die Ausgewählte Funktion in Zeile ") & r & Get_Language_Str(" ist immer aktiv und kann nicht über DCC oder CAN geschaltet werden.")
                     Addr = -1 ' No address given of InCnt <= 0 or empty
                ElseIf Addr >= 1 And Addr <= MaxAddr Then
                     ' Valid adress range
                Else: Add_to_Err Cells(r, DCC_or_CAN_Add_Col), Get_Language_Str("Die Adresse '") & Replace(Cells(r, DCC_or_CAN_Add_Col), vbLf, " ") & Get_Language_Str("' in Zeile ") & r & Get_Language_Str(" ist ungültig.")
                End If
            End If
        Else ' Address or selectrix channel entry doesn't start with a number           ' 03.04.20:
            Dim VarName As String
            VarName = Get_Address_String(r)
            If VarName <> "" Then
               If Not Valid_Var_Name(VarName, r) Then Exit Function
               Addr = VarName
            End If
        End If
        If Not Create_Header_Entry(r, Addr) Then
            Exit Function
        End If
    End If
  Next r
  
  If Check_Detected_Variables() = False Then Exit Function
  If CheckArduinoHomeDir() = False Then Exit Function           ' 02.12.21: Juergen see forum post #7085
  Create_HeaderFile = Write_Header_File_and_Upload_to_Arduino(CreateFilesOnly)       ' 20.12.21: Jürgen return result of called function
End Function


'----------------------------------------------------
Private Sub DelTailingEmptyLines(ByRef Txt As String)
'----------------------------------------------------
  While right(Txt, 2) = vbCr & vbCr
    Txt = left(Txt, Len(Txt) - 1)
  Wend
End Sub

'--------------------------------------------
Public Function Ext_AddrTxt_Used() As Boolean
'--------------------------------------------
' Check if DCC, SX or CAN is used
  Ext_AddrTxt_Used = (Ext_AddrTxt <> "")
End Function

'begin change 01.05.20: Jürgen
'--------------------------------------------
Public Function Store_ValuesTxt_Used() As Boolean
'--------------------------------------------
  Store_ValuesTxt_Used = (Store_ValuesTxt <> "")
End Function
'end change 01.05.20: Jürgen

'----------------------------------------------------
Private Function Write_Header_File_and_Upload_to_Arduino(Optional CreateFilesOnly As Boolean = False) As Boolean    ' 20.12.21: Jürgen add CreateFilesOnly for programatically generation of header files
'----------------------------------------------------
  Dim NumLeds As Long, Nr As Long, MaxLed As Long
  
  MaxLed = Get_Current_Platform_Int("MaxLed")                '03.04.21: Juergen, max Leds depends on board type
  For Nr = 0 To LED_CHANNELS - 1
      NumLeds = NumLeds + Cells(SH_VARS_ROW, Get_LED_Nr_Column(Nr))   ' 26.04.20: Old: NumLeds = Cells(SH_VARS_ROW, LED_Nr__Col)
  Next Nr
  
  If NumLeds < MINLEDs Then  ' To be able to test at least 20 LEDs with the color test program ' 26.10.20:
     NumLeds = MINLEDs
  End If
  
  If NumLeds > MaxLed Then
    Err = Err & Get_Language_Str("Maximale LED Anzahl überschritten: ") & NumLeds & vbCr & _
    Get_Language_Str("Es sind maximal #1# RGB LEDs möglich") & vbCr  ' Don't check before to be able to temprory add more than 256 LES
    Err = Replace(Err, "#1#", MaxLed)                                ' 03.04.21 Juergen replace with actual number
  End If
  
  If Err <> "" Then
     MsgBox Err & vbCr & vbCr & _
            Get_Language_Str("Ein neues Header file wurde nicht generiert!"), vbCritical, Get_Language_Str("Es sind Fehler aufgetreten")
     Exit Function
  End If
  
  Dim Name As String
  Name = ThisWorkbook.Path & "/" & Ino_Dir_LED & Include_FileName
  
  DelTailingEmptyLines Ext_AddrTxt
  DelTailingEmptyLines Store_ValuesTxt                     ' 01.05.20: Jürgen
  DelTailingEmptyLines InChTxt
  
  Create_Loc_InCh_Defines InChTxt, Channel, LocInChNr
    
  Dim ShortPath As String, p As Long
  p = InStrRev(ThisWorkbook.Path, "\")
  If p = 0 Then p = InStrRev(ThisWorkbook.Path, "/")
  If p > 0 Then ShortPath = Mid(ThisWorkbook.Path, p + 1, 255) & " "
  
  Dim fp As Integer
  fp = FreeFile
  On Error GoTo WriteError
  Open Name For Output As #fp
  Print #fp, "// This file contains the " & Page_ID & " and LED definitions."
  Print #fp, "//"
  Print #fp, "// It was automatically generated by the program " & ThisWorkbook.Name & " " & Prog_Version & "      by Hardi"
  Print #fp, "// File creation: " & Date & " " & Time
  Print #fp, "// (Attention: The display in the Arduino IDE is not updated if Options/External Editor is disabled)"
  Print #fp, ""
  Print #fp, "#ifndef __LEDS_AUTOPROG_H__"
  Print #fp, "#define __LEDS_AUTOPROG_H__"
  Print #fp, ""
  Print #fp, "#ifndef CONFIG_ONLY"          ' 04.03.22 Juergen: add Simulator feature
  Print #fp, "#ifndef ARDUINO_RASPBERRY_PI_PICO"
  Print #fp, "#define FASTLED_INTERNAL       // Disable version number message in FastLED library (looks like an error)"                ' 11.01.20: Added Block
  Print #fp, "#include <FastLED.h>           // The FastLED library must be installed in addition if you got the error message ""..fatal error: FastLED.h: No such file or directory"""
  Print #fp, "                               // Arduino IDE: Sketch / Include library / Manage libraries                    Deutsche IDE: Sketch / Bibliothek einbinden / Bibliothek verwalten"
  Print #fp, "                               //              Type ""FastLED"" in the ""Filter your search..."" field                          ""FastLED"" in das ""Grenzen Sie ihre Suche ein"" Feld eingeben"
  Print #fp, "                               //              Select the entry and click ""Install""                                         Gefundenen Eintrag auswaehlen und ""Install"" anklicken"
  Print #fp, "#else"
  Print #fp, "#include <PicoFastLED.h>       // Juergens minimum version or FastLED for Raspberry Pico"
  Print #fp, "#endif"
  Print #fp, "#endif // CONFIG_ONLY"
  Print #fp, ""
  Print #fp, "#include <MobaLedLib.h>"
  Print #fp, ""
  Print #fp, "#define START_MSG ""LEDs_AutoProg Ver 1: " & ShortPath & Format(Date, "dd.mm.yy") & " " & Format(Time, "hh:mm") & """" ' The version could be read out in a future version of this tool
  Print #fp, ""
  If Page_ID = "Selectrix" Then
         Print #fp, "#define TWO_BUTTONS_PER_ADDRESS 0      // One button is used (Selectrix)"
         Print #fp, "#define USE_SX_INTERFACE               // enable Selectrix protocol on single CPU mainboards"                  ' 06.12.2021 Juergen add SX for ESP
  Else:  Print #fp, "#define TWO_BUTTONS_PER_ADDRESS 1      // Two buttons (Red/Green) are used (DCC/CAN)"
  End If
  Print #fp, "#ifdef NUM_LEDS"
  Print #fp, "  #warning ""'NUM_LEDS' definition in the main program is replaced by the included '" & FileNameExt(Name) & "' with " & NumLeds & """"
  Print #fp, "  #undef NUM_LEDS"
  Print #fp, "#endif"
  Print #fp, ""
  Print #fp, "#define NUM_LEDS " & AddSpaceToLen(NumLeds, 22) & "// Number of LEDs (Maximal 256 RGB LEDs could be used)"
  Print #fp, ""
  Print #fp, "#define LEDS_PER_CHANNEL ""," & LEDs_per_ChannelList & """"     ' 13.03.21 Juergen - for new Farbtest initialisation
  
  ' Set HOUSE_MIN_T and HOUSE_MAX_T
  Dim House_Min_T As String                                                 ' 26.09.19:
  House_Min_T = Get_String_Config_Var("MinTime_House")
  If House_Min_T <> "" Then
        Print #fp, "#undef  HOUSE_MIN_T"
        Print #fp, "#define HOUSE_MIN_T  " & val(House_Min_T)
  Else: House_Min_T = 50 ' Default value used in the library
  End If
  Dim House_Max_T As String
  House_Max_T = Get_String_Config_Var("MaxTime_House")
  If House_Max_T <> "" Then
     Print #fp, "#undef  HOUSE_MAX_T"
     Print #fp, "#define HOUSE_MAX_T " & val(House_Max_T)
  Else: House_Max_T = 150 ' Default value used in the library
  End If
  If val(House_Min_T) > val(House_Max_T) Or val(House_Max_T) = 0 Then
     Sheets(ConfigSheet).Select
     Range("MinTime_House").Select
     Sleep 100
     MsgBox Get_Language_Str("Fehler auf der 'Config' Seite:" & vbCr & _
            "Die 'Minimale Zeit bis zur nächsten Änderung' muss kleiner " & _
            "oder gleich groß wie die Maximale Zeit sein." & vbCr & _
            "Achtung: Wenn nichts eingegeben ist werden die Standard Werte vom 50/150 verwendet. " & _
            "Dadurch kann es ebenfalls zu einem Konflikt kommen."), vbCritical, Get_Language_Str("Falsche Zeiten für die House() Funktion")
     EndProg
  End If
  
  Print #fp, ""
  
  Dim Color_Test_Mode As String: Color_Test_Mode = Get_String_Config_Var("Color_Test_Mode")
  Select Case left(UCase(Color_Test_Mode), 1)
      Case "J", "Y", "1": Print #fp, "#define RECEIVE_LED_COLOR_PER_RS232" & vbCr
  End Select
  
  If Get_Bool_Config_Var("USE_SPI_Communication") Or Page_ID = "CAN" Then         ' 14.05.20: Change the heartbeat LED pin  ' 04.10.20: Added: Page_ID = "CAN"
     If Get_Bool_Config_Var("USE_SPI_Communication") Then
           Print #fp, "#define USE_SPI_COM                    // Use the SPI bus for the communication in addition to the RS232 if J13 is closed. If no DCC commands are configured the A1 pin of the DCC Arduino is disabled"
     End If
     If PIN_A3_Is_Used() Then
           Print #fp, "#define LED_HEARTBEAT_PIN -1           // Disable the heartbeat pin because it's used for the SwitchB or SwitchC"
     Else: Print #fp, "#define LED_HEARTBEAT_PIN A3           // Don't use the internal heartbeat LED because the D13 pins between LED and DCC arduin are connected together"
     End If
  End If
  
  
  If Ext_AddrTxt_Used() Then
  
    If Get_Bool_Config_Var("USE_SPI_Communication") Then                    ' 16.05.20:
       If Check_Switch_Lists_for_SPI_Pins() = False Then
            Close #fp
           Exit Function
       End If
    End If
  
    Print #fp, "#define USE_EXT_ADDR"
    
    If InStr(Prog_for_Right_Ardu, " " & Page_ID & " ") > 0 Then
       Print #fp, "#define USE_RS232_OR_SPI_AS_INPUT      // Use the RS232 or SPI Input to read DCC/SX commands from the second Arduino and from the PC (The SPI is only used if enabled with USE_SPI_COM)"
    End If
    
'    If Get_Bool_Config_Var("USE_SPI_Communication") Then                    ' 14.05.20:
'       Print #fp, "#define USE_SPI_COM                    // Use the SPI bus for the communication in addition to the RS232 if J13 is closed"
'    End If
    
    ' Set DCC Offset                                                        ' 26.09.19:
    If Page_ID = "DCC" Then
          Print #fp, "#define ADDR_OFFSET " & val(Get_String_Config_Var("DCC_Offset"))
    Else: Print #fp, "#define ADDR_OFFSET 0"
    End If
    
    If Page_ID = "CAN" Then
       Print #fp, "#define USE_CAN_AS_INPUT"
    End If
    
    
    Print #fp, ""
    Print #fp, "#define ADDR_MSK  0x3FFF  // 14 Bits are used for the Address"
    Print #fp, ""
    Print #fp, "#define S_ONOFF   (uint16_t)0"
    Print #fp, "#define B_RED     (uint16_t)(1<<14)"
    Print #fp, "#define B_GREEN   (uint16_t)(2<<14)"
    Print #fp, "#define B_RESERVE (uint16_t)(3<<14)    // Not used at the moment"
    Print #fp, "#define B_TAST    B_RED"
    Print #fp, ""
    Print #fp, ""
    Print #fp, "typedef struct"
    Print #fp, "    {"
    Print #fp, "    uint16_t AddrAndTyp; // Addr range: 0..16383. The upper two bytes are used for the type"
    Print #fp, "    uint8_t  InCnt;"
    Print #fp, "    } __attribute__ ((packed)) Ext_Addr_T;"                 ' 05.11.20: Added: __attribute__ ((packed)) to be able to use it on oa 32 Bit platform
    Print #fp, ""
    Print #fp, "// Definition of external adresses"
    Print #fp, "#ifdef CONFIG_ONLY"                                         ' 04.03.22 Juergen: add Simulator feature
    Print #fp, "const Ext_Addr_T Ext_Addr[] __attribute__ ((section ("".MLLAddressConfig""))) ="
    Print #fp, "#else"
    Print #fp, "const PROGMEM Ext_Addr_T Ext_Addr[] ="
    Print #fp, "#endif"
    Print #fp, "         { // Addr & Typ    InCnt"
    Print #fp, Ext_AddrTxt;
    Print #fp, "         };"
    Print #fp, ""
    Print #fp, ""
  End If ' Ext_AddrTxt <> ""
  
  Print #fp, "// Input channel defines for local inputs and expert users" ' 05.10.19: Moved out of the if because the local inputs are also stored here
  Print #fp, InChTxt
  Print #fp, ""
  
  If Write_Switches_Header_File_Part_A(fp, Channel) = False Then
     Close #fp
     Exit Function
  End If
  
  If Write_LowProrityLoop_Header_File(fp) = False Then
     Close #fp
     Exit Function
  End If
  
  If Write_Header_File_LED2Var(fp) = False Then                             ' 08.10.20:
     Close #fp
     Exit Function
  End If
  
   ' 15.10.21: Juergen split creation of sound extensions to ensure that preprocessor defines are corretly compiled
  If Write_Header_File_Sound_Before_Config(fp) = False Then
     Close #fp
     Exit Function
  End If
  ' 31.01.22: Juergen add extension support
  If Write_Header_File_Extension_Before_Config(fp) = False Then
     Close #fp
     Exit Function
  End If

  Print #fp, DayAndNightTimer                                               ' 07.10.20:
  
  Print #fp, ""
  Print #fp, "//*******************************************************************"
  Print #fp, "// *** Configuration array which defines the behavior of the LEDs ***"
  Print #fp, "MobaLedLib_Configuration()"
  Print #fp, "  {"
  Print #fp, ConfigTxt
  Print #fp, "  EndCfg // End of the configuration"
  Print #fp, "  };"
  Print #fp, "//*******************************************************************"
  Print #fp, ""
  Print #fp, "//---------------------------------------------"
  Print #fp, "void Set_Start_Values(MobaLedLib_C &MobaLedLib)"
  Print #fp, "//---------------------------------------------"
  Print #fp, "{"
  Print #fp, Start_Values;
  Print #fp, "}"
  Print #fp, ""
  
  
 'begin change 01.05.20: Jürgen
  If Store_ValuesTxt_Used Then                                              ' 01.05.20: Juergen
    'Print #fp, "#define ENABLE_STORE_STATUS"                               ' 01.05.20: disabled in Mail from Juergen
    Print #fp, ""
    Print #fp, "// if function returns TRUE the calling loop stops"
    Print #fp, "typedef bool(*HandleValue_t) (uint8_t CallbackType, uint8_t ValueId, uint8_t* Value, uint16_t EEPromAddr, uint8_t TargetValueId, uint8_t Options);"
    Print #fp, ""
    Print #fp, ""
    Print #fp, "#define InCnt_MSK  0x0007  // 3 Bits are used for the InCnt"
    Print #fp, "#define IS_COUNTER (uint8_t)0x80"
    Print #fp, "#define IS_PULSE   (uint8_t)0x40"
    Print #fp, "#define IS_TOGGLE  (uint8_t)0x00"
    Print #fp, "#define COUNTER_ID"
    Print #fp, ""
    Print #fp, "typedef struct"
    Print #fp, "    {"
    Print #fp, "    uint8_t TypAndInCnt; // Type bit 7, InCnt bits 0..3, reserved 0 bits 4..6"
    Print #fp, "    uint8_t Channel;"
    Print #fp, "    } __attribute__ ((packed)) Store_Channel_T;"            ' 05.11.20: Added: __attribute__ ((packed)) to be able to use it on oa 32 Bit platform
    Print #fp, ""
    Print #fp, "// Definition of channels and counters that need to store state in EEProm" & vbCr & _
               "const PROGMEM Store_Channel_T Store_Values[] =" & vbCr & _
               "         { // Mode + InCnt , Channel"
    Print #fp, Store_ValuesTxt;
    Print #fp, "         };"
    Print #fp, ""
  Else
    Print #fp, ""
    Print #fp, "// No macros used which are stored to the EEPROM => Disable the ENABLE_STORE_STATUS flag in case it was set in the excel sheet"
    Print #fp, "#ifdef ENABLE_STORE_STATUS"                                 ' 01.05.20: New block in Mail from Juergen
    Print #fp, "  #undef ENABLE_STORE_STATUS"
    Print #fp, "#endif"
    Print #fp, ""
 End If ' Store_ValuesTxt_Used
 'end change 19.04.20 Jürgen
  
  Print #fp, "#ifndef CONFIG_ONLY"          ' 04.03.22 Juergen: add Simulator feature
  
 ' 15.10.21: Juergen move creation of onboard sound code after the configuration struture to ensue that #defines from ProgGenerator are effective
  If Write_Header_File_Sound_After_Config(fp) = False Then
     Close #fp
     Exit Function
  End If
 ' 31.01.22: Juergen add extension support
  If Write_Header_File_Extension_After_Config(fp) = False Then
     Close #fp
     Exit Function
  End If

  Print #fp, "#endif // CONFIG_ONLY"          ' 04.03.22 Juergen: add Simulator feature
  Print #fp, ""
  Print #fp, ""
  Print #fp, ""
  Print #fp, ""
  Print #fp, ""
  Print #fp, "#endif // __LEDS_AUTOPROG_H__"
  
  Close #fp
  On Error GoTo 0
  
  If Channel - 1 > 250 Then
     MsgBox Get_Language_Str("Fehler: Die Anzahl der verwendeten Eingangskanäle ist zu groß!" & vbCr & _
            "Es sind maximal 250 verfügbar. Die Konfiguration enthält aber ") & Channel - 1 & "." & vbCr & _
            vbCr & _
            Get_Language_Str("Die Eingangskanäle werden zum einlesen von DCC, Selectrix und CAN Daten benutzt. " & vbCr & _
            "Außerdem werden sie als interne Zwischenspeicher benötigt."), vbCritical, Get_Language_Str("Anzahl der InCh Variablen überschritten")
     EndProg
  End If
  
  If ConfigTxt = "" Then
     MsgBox Get_Language_Str("Achtung: Es ist keine einzige Zeile in der Spalte ""Beleuchtung, Sound, oder andere Effekte"" aktiv!" & vbCr & _
            "=> Das Programm wird keine LEDs ansteuern"), vbCritical, Get_Language_Str("Achtung: Die Konfiguration ist leer")
     UserForm_Header_Created.DontShowAgain = False
  End If
  
  Application.StatusBar = Time & Get_Language_Str(": Header Datei '") & Name & Get_Language_Str("' wurde erzeugt") ' 14.07.20: Don't use Show_Status_for_a_while because the compile time is shorter with Jürgens new PrivateBuild command
  'Show_Status_for_a_while Time & Get_Language_Str(": Header Datei '") & Name & Get_Language_Str("' wurde erzeugt")
    
  If CreateFilesOnly = False And UserForm_Header_Created.DontShowAgain = False Then  ' 20.12.21: Jürgen add CreateFilesOnly for programatically generation of header files
        UserForm_Header_Created.FileName = Name
        UserForm_Header_Created.Show
  Else: Compile_and_Upload_LED_Prog_to_Arduino CreateFilesOnly
  End If
  ResetTestButtons Store_Status_Enabled                                     ' 19.01.21: Jürgen
  Write_Header_File_and_Upload_to_Arduino = True                            ' 20.12.21: Jürgen add CreateFilesOnly for programatically generation of header files
  Exit Function
  
WriteError:
  ' Attention: This could also be an error some where else in the code
  MsgBox Get_Language_Str("Fehler beim schreiben der Datei '") & Name & "'", vbCritical, Get_Language_Str("Fehler beim erzeugen der Arduino Header Datei")
  Close #fp
End Function





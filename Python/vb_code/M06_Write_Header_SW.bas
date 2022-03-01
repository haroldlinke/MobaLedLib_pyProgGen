Attribute VB_Name = "M06_Write_Header_SW"
Option Explicit

' Header file generation for switches and variables
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Private DstVar_List As String
Private MultiSet_DstVar_List As String

Private StoreVar_List As String      ' 19.04.20 Jürgen

Private MaxUsed_Loc_InCh As Long
Private MaxUsed_Loc_InCh_Row As Long
  
Public SwitchA_InpCnt As Long        ' Number of switch imputs for channel A (A=Analog push buttons)
Public SwitchB_InpCnt As Long        ' Number of switch imputs for channel B (B=Border:  push buttons (or switches) around the border of the layout
Public SwitchC_InpCnt As Long        ' Number of switch imputs for channel C (C=Console: switches (or push buttons) in the console (Weichenstellpult)
Public SwitchD_InpCnt As Long        ' Number of switch imputs for channel D (D=Direct: :  switches (or push buttons) connected direct to the main board
Public SwitchA_InpLst As String      ' List if analog input pins for channel A
Public SwitchB_InpLst As String      ' List if input pins for channel B
Public SwitchC_InpLst As String      ' List if input pins for channel C
Public SwitchD_InpLst As String      ' List if input pins for channel C
Public CLK_Pin_Number As String
Public RST_Pin_Number As String
Public LDR_Pin_Number As String
Public Serial_PinLst As String                                             ' 08.10.21 Juergen
Public DMX_LedChan As Long                                                 ' 19.01.21 Juergen
  
Private CTR_Channels_1 As Long       ' If one Switch (B or C) is used this counter channel is used
Private CTR_Channels_2 As Long       ' If two Switches (B and C) are used this counter channel is used for SwitchC
Private Channel1InpCnt As Long
Private Channel2InpCnt As Long
Private CTR_Cha_Name_1 As String
Private CTR_Cha_Name_2 As String
Private But_Inp_List_1 As String
Private But_Inp_List_2 As String
  
Public LED_PINNr_List As String     ' Pin numbers for controlling the RGB LEDs                ' 26.04.20:
  
Public Read_LDR As Boolean
Private Use_WS2811 As Boolean                                             ' 19.01.21 Juergen
Public Store_Status_Enabled As Boolean                                    '   "
Public Switch_Damping_Fact As String


'------------------------------------------
Public Function PIN_A3_Is_Used() As Boolean
'------------------------------------------
  If Channel1InpCnt > 0 And RST_Pin_Number = "A3" Then
     PIN_A3_Is_Used = True
  End If
End Function

'---------------------------------------------------------------------------
Private Function Add_Logic_InpVars(LogicExp As String, r As Long) As Boolean
'---------------------------------------------------------------------------
  Dim Arglist() As String, Arg As Variant
  Arglist = SplitEx(LogicExp, True, "OR", "AND", "NOT")
  For Each Arg In Arglist
      Arg = Trim(Arg)
      If Arg <> "" Then
         If Valid_Var_Name_and_Skip_InCh_and_Numbers(Arg, r) = False Then Exit Function ' Skip special names like '#InCh' and Numbers
      End If
  Next
  Add_Logic_InpVars = True
End Function

'----------------------------------------------------------------------------------------------
Private Function Check_if_all_Variables_in_sequece_of_N_exists(r As Long, N As Long) As Boolean
'----------------------------------------------------------------------------------------------
  Dim Adr_or_Name As String
  Adr_or_Name = Trim(Cells(r, Get_Address_Col()))
  If Adr_or_Name = "" Then
     Cells(r, Get_Address_Col()).Select
     MsgBox Replace(Get_Language_Str("Fehler: In Zeile #1# ist keine Adresse, kein Schalter oder keine Variable eingetragen"), "#1#", r), vbCritical, _
            Replace(Get_Language_Str("Kein Eintrag in '#1#' Spalte"), "#1#", Get_Address_String(Header_Row))
     Exit Function
  End If
  If Not IsNumeric(Split(Adr_or_Name, " ")(0)) Then ' Adr could be a number "17" or a range "17 - 18"
     Dim Nr As Long, TxtLen As Long, Name As String, i As Long
     Nr = Get_Nr_From_Var(Adr_or_Name, TxtLen)
     If Nr >= 0 Then
        Name = Left(Adr_or_Name, TxtLen)
        For i = Nr To Nr + N - 1
            If Valid_Var_Name(Name & i, r) = False Then Exit Function ' At the moment Valid_Var_Name always returns true
        Next i
     End If
  End If
  Check_if_all_Variables_in_sequece_of_N_exists = True
End Function


'-----------------------------------------------
Function Get_Bin_Inputs(Dec_Cnt As Long) As Long
'-----------------------------------------------
  If Dec_Cnt <= 0 Then
                            Get_Bin_Inputs = -1
  ElseIf Dec_Cnt <= 1 Then
                            Get_Bin_Inputs = 1
  ElseIf Dec_Cnt <= 3 Then
                            Get_Bin_Inputs = 2
  ElseIf Dec_Cnt <= 7 Then
                            Get_Bin_Inputs = 3
  ElseIf Dec_Cnt <= 15 Then
                            Get_Bin_Inputs = 4
  ElseIf Dec_Cnt <= 31 Then
                            Get_Bin_Inputs = 5
  ElseIf Dec_Cnt <= 63 Then
                            Get_Bin_Inputs = 6
  Else:                     Get_Bin_Inputs = -1 ' Error
  End If
End Function

'------------------------------------------------------------------------------------------------------------------
Private Function Check_if_all_Variables_in_sequece_exist(r As Long, Ctr_Name As String, N_Str As String) As Boolean
'------------------------------------------------------------------------------------------------------------------
' Is called if a macro of those is checked
' - InCh_to_TmpVar(InCh, InCh_Cnt)
' - Charlie_Buttons(LED, InCh, States)
' - Charlie_Binary(LED, InCh, BinStates)
' InCh could be a DCC Variable, a Switch or an "Array" variable.
' In the last two cases it has to be checked if all required variables exist.
' Example: InCh_to_TmpVar(SwitchA3, 3)
' => SwitchA3, SwitchA4, SwitchA5 must be defined
  Dim N As Long
  If Ctr_Name = "BinStates" Then
        N = Get_Bin_Inputs(val(N_Str))
        If N <= 0 Then
            MsgBox Get_Language_Str("Fehler: Anzahl der binären Zustände ungültig. Die Anzahl muss zwischen 1 und 63 liegen"), vbCritical, Get_Language_Str("Anzahl der binären Zustände ungültig")
            Exit Function
        End If
  Else: N = val(N_Str)
  End If
  Check_if_all_Variables_in_sequece_exist = Check_if_all_Variables_in_sequece_of_N_exists(r, N)
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Add_InpVars(ByVal MacroName As String, Org_Macro As String, Filled_Macro As String, r As Long, Org_Macro_Row As Long) As Boolean
'------------------------------------------------------------------------------------------------------------------------------------------------
' Die benutzten Schalter Eingänge müssen aus den Parametern der Makros gelesen werden
' Die Eingänge können an verschiedenen Stellen benutzt werden
' - Logic:
' - Zwei Eingänge z.B. "RS_FlipFlop(DstVar, InCh, R_InCh)"
'   Hier gibt es verschiedene mögliche Kandidaten: "R_InCh InReset"
' - Es gibt auch "versteckte" Eingänge bei Makros wie "InCh_to_TmpVar(InCh, InCh_Cnt)"
'   Diese sind in Spalte "InCh" mit "n", "States", "BinStates", "2", "3", "4"
'
' Examples:
' - Logic(TestOr, #InCh OR #InCh+1 OR SwitchA4)              O.K.
' - RS_FlipFlopTimeout(FlipFlop, #InCh+1, #InCh, 30 Sek)     O.K.    Hier muss man das +1 von Hand vertauschen wenn die "Rot= Reset" sein soll
' - RS_FlipFlopTimeout(FlipFlop, #InCh, SwitchA5, 30 Sek)    O.K.
' - RS_FlipFlopTimeout(FlipFlop, #InCh, SI_0, 30 Sek)        O.K.
' - EntrySignal3_RGB(LED, InCh)                              O.K.
' - InCh_to_TmpVar(InCh, InCh_Cnt)                           O.K.
' - Charlie_Buttons(LED, InCh, States)                       O.K.
' - Charlie_Binary(LED, InCh, BinStates)                     O.K.
  Const Second_Input_Names = "InCh R_InCh InReset InCh2"  ' Added InCh to check also manually changed lines. Important for: RS_FlipFlopTimeout(FlipFlop, #InCh+1, #InCh, 30 Sek)
  Dim Arg_List() As String, Fil_List() As String, ArgNr As Long, Name As Variant, Arg As String
  Arg_List = Get_Arguments(Org_Macro)
  Fil_List = Get_Arguments(Filled_Macro)
  Select Case MacroName
    Case "Logic":   ' Locic macro uses several arguments separated by "AND", "OR" and "NOT"
                    If UBound(Fil_List) <> 1 Then
                       Cells(r, Config__Col).Select
                       MsgBox Replace(Get_Language_Str("Fehler: Falsche Parameter Anzahl in 'Logic()' Ausdruck: '#1#'"), "#1#", Filled_Macro), vbCritical, _
                              Get_Language_Str("Fehler: 'Logic()' Ausdruck ist ungültig")
                       Exit Function
                    End If
                    Add_InpVars = Add_Logic_InpVars(Fil_List(1), r)

    Case Else:      ' Other Macros
                    Dim Det_Cnt As Long, Pos_CounterVar As Long, Pos_InCh As Long
                    Pos_CounterVar = -1
                    Pos_InCh = -1
                    For ArgNr = 0 To UBound(Arg_List)
                        Arg = Arg_List(ArgNr)
                        If Arg = "InCh_Cnt" Or Arg = "States" Or Arg = "BinStates" Then Pos_CounterVar = ArgNr
                        If Arg = "InCh" Then Pos_InCh = ArgNr
                        For Each Name In Split(Second_Input_Names, " ")
                            If Arg = Name Then
                               Det_Cnt = Det_Cnt + 1
                               If Valid_Var_Name_and_Skip_InCh_and_Numbers(Fil_List(ArgNr), r) = False Then
                                  Exit Function
                               End If
                            End If
                        Next Name
                    Next ArgNr
                    
                    If Pos_CounterVar >= 0 Then ' Macro like "InCh_to_TmpVar(InCh, InCh_Cnt)"
                        If Check_if_all_Variables_in_sequece_exist(r, Arg_List(Pos_CounterVar), Fil_List(Pos_CounterVar)) = False Then Exit Function
                    Else
                        ' Check if there is a number in column "InCnt" and all arguments have been found
                        Dim InCntStr As String
                        InCntStr = ThisWorkbook.Sheets(LIBMACROS_SH).Cells(Org_Macro_Row, SM_InCnt_COL)
                        If InCntStr <> "" Then                              ' 14.05.20:
                            If IsNumeric(InCntStr) Then
                               If Det_Cnt <> val(InCntStr) Then ' Not all arguments detected
                                  ' It's a function like: "EntrySignal3_RGB(LED, InCh)"
                                  ' => Check if all Variable in the sequece exist: <Name>1..<Name><InCntStr>
                                  If Check_if_all_Variables_in_sequece_of_N_exists(r, val(InCntStr)) = False Then Exit Function
                               End If
                            End If
                            If IsNumeric(Left(InCntStr, Len(InCntStr) - 1)) And Right(InCntStr, 1) = "?" Then      ' 07.05.20:
                               'ToDo Zusätzliche Überprüfung auf "#InCh+2" wenn 3? in Lib_Macros hinzugefügt wird
                               If Fil_List(2) = "#InCh+1" Then              ' 07.06.20:
                                  ' It's a function which may use two inputs like: "RS_FlipFlop(Test, #InCh, #InCh+1)"
                                  ' => Check if all Variable in the sequece exist: <Name>1..<Name><InCntStr>
                                  If Check_if_all_Variables_in_sequece_of_N_exists(r, val(InCntStr)) = False Then Exit Function
                               End If
                            End If
                        End If
                    End If
  End Select
  
  Add_InpVars = True
End Function

'-------------------------------------------------------------------------
Private Function Is_Switch_Var_then_Add_to_Ctr(Var_Name As String) As Long
'-------------------------------------------------------------------------
  Dim Nr As Long
  Select Case Is_in_Nr_String(Var_Name, "Switch?", 1, 250, Nr)
     Case 1:  ' Valid number range
              Select Case Left(Var_Name, Len("Switch?"))
                  Case "SwitchA":
                                  If Nr > SwitchA_InpCnt Then SwitchA_InpCnt = Nr
                  Case "SwitchB":
                                  If Nr > SwitchB_InpCnt Then SwitchB_InpCnt = Nr
                  Case "SwitchC":
                                  If Nr > SwitchC_InpCnt Then SwitchC_InpCnt = Nr
                  Case "SwitchD":
                                  If Nr > SwitchD_InpCnt Then SwitchD_InpCnt = Nr
                  Case Else:      Debug.Print "Unsupported variable '" & Var_Name & "' in 'First_Scan_of_Data_Rows()'"
              End Select
              Is_Switch_Var_then_Add_to_Ctr = 1
     Case -1: MsgBox Replace(Replace(Get_Language_Str("Fehler: Die Nummer der Variable '#1#' ist ungültig!" & vbCr & _
                                     vbCr & _
                                     "Gültiger Bereich: #2#"), "#1#", Var_Name), "#2#", "1..250"), vbCritical, _
                                     Get_Language_Str("Fehler: Ungültige Variable")
              Is_Switch_Var_then_Add_to_Ctr = -1
  End Select
End Function


'----------------------------------------------------
Private Function First_Scan_of_Data_Rows() As Boolean                       ' 04.04.20:
'----------------------------------------------------
' Set the global variables if the corrosponding entries are found
' - in the Var_Col:
'     "Switch?<Nr>"
' - in then Config__Col:
'     "// Set_Switch?_InpLst("
'     "DstVar*"
'     "#define READ_LDR"
'
  Dim r As Long, Var_COL As Long
  Var_COL = Get_Address_Col()
  Switch_Damping_Fact = ""                                                  ' 04.11.21:
  
  For r = FirstDat_Row To LastUsedRow
     If Not Rows(r).EntireRow.Hidden And Cells(r, Enable_Col) <> "" Then
       Dim Var_Name As String
       Var_Name = Cells(r, Var_COL)
       
       If Is_Switch_Var_then_Add_to_Ctr(Var_Name) = -1 Then
          Cells(r, Var_COL).Select
          Exit Function
       End If
       
       Dim Config_Entry As String, line As Variant
       Config_Entry = Cells(r, Config__Col)
       If Trim(Config_Entry) <> "" Then
          For Each line In Split(Config_Entry, vbLf)
              line = Trim(line)
              If Set_PinNrLst_if_Matching(line, "// Set_SwitchA_InpLst(", SwitchA_InpLst, "A", 5) = False Then Exit Function
              If Set_PinNrLst_if_Matching(line, "// Set_SwitchB_InpLst(", SwitchB_InpLst, "I", 12) = False Then Exit Function
              If Set_PinNrLst_if_Matching(line, "// Set_SwitchC_InpLst(", SwitchC_InpLst, "I", 12) = False Then Exit Function
              If Set_PinNrLst_if_Matching(line, "// Set_SwitchD_InpLst(", SwitchD_InpLst, "Pu", 12) = False Then Exit Function
              If Set_PinNrLst_if_Matching(line, "// Set_CLK_Pin_Number(", CLK_Pin_Number, "O", 1) = False Then Exit Function
              If Set_PinNrLst_if_Matching(line, "// Set_RST_Pin_Number(", RST_Pin_Number, "O", 1) = False Then Exit Function
              If Set_PinNrLst_if_Matching(line, "// Set_LDR_Pin_Number(", LDR_Pin_Number, "A", 1) = False Then Exit Function
              If Set_PinNrLst_if_Matching(line, "// Set_LED_OutpPinLst(", LED_PINNr_List, "O", LED_CHANNELS) = False Then Exit Function
              
              If line = "// Use_DMX512()" Then                              ' 19.01.21 Juergen
                DMX_LedChan = val(Cells(r, LED_Cha_Col))
              End If
              
              If line = "#define READ_LDR" Then Read_LDR = True
              If Left(line, Len("#define SWITCH_DAMPING_FACT")) = "#define SWITCH_DAMPING_FACT" Then ' 04.11.21:
                    Switch_Damping_Fact = line
              End If
              If line = "#define USE_WS2811" Then Use_WS2811 = True                        ' 19.01.21 Juergen
              If line = "#define ENABLE_STORE_STATUS" Then Store_Status_Enabled = True     '    "
              
              If Add_Inp_and_DstVars(line, r) = False Then Exit Function ' Add the destination variable to DstVar_List
              
          Next line
       End If
     End If
  Next r
  First_Scan_of_Data_Rows = True
End Function


Sub Test()
  Debug.Print Replace("Aber    Hallo", "  ", " ")
End Sub

'---------------------------------------------------------------------------------------
Private Function Check_one_Switch_Lists_for_SPI_Pins(ByVal Sw_List As String) As Boolean
'---------------------------------------------------------------------------------------
  Dim Pin As Variant
  Sw_List = " " & Sw_List & " "
  For Each Pin In Split("10 11 12", " ")
      If InStr(Sw_List, " " & Pin & " ") > 0 Then
         MsgBox Replace(Get_Language_Str("Fehler: Der Arduino Pin '#1#' kann nicht als Ein- oder Ausgang werden wenn " & _
                "DCC oder Selectrix Daten per SPI Bus gelesen werden. Es muss ein anderer Anschluss verwendet " & _
                "werden oder die SPI Kommunikation in der 'Config' Seite deaktiviert werden." & vbLf & _
                "Achtung: Die beiden Arduinos müssen dann per RS232 verbunden sein."), "#1#", Pin), vbCritical, "Fehler: Ungültiger Arduino Pin erkannt"
         Exit Function
      End If
  Next Pin
  Check_one_Switch_Lists_for_SPI_Pins = True
End Function


'-----------------------------------------------------------
Public Function Check_Switch_Lists_for_SPI_Pins() As Boolean
'-----------------------------------------------------------
  If SwitchA_InpCnt > 0 Then
     If Check_one_Switch_Lists_for_SPI_Pins(SwitchA_InpLst) = False Then Exit Function
  End If
  If SwitchB_InpCnt > 0 Then
     If Check_one_Switch_Lists_for_SPI_Pins(SwitchB_InpLst) = False Then Exit Function
  End If
  If SwitchC_InpCnt > 0 Then
     If Check_one_Switch_Lists_for_SPI_Pins(SwitchC_InpLst) = False Then Exit Function
  End If
  If SwitchD_InpCnt > 0 Then
     If Check_one_Switch_Lists_for_SPI_Pins(SwitchD_InpLst) = False Then Exit Function
  End If
  Check_Switch_Lists_for_SPI_Pins = True
End Function


'--------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Set_PinNrLst_if_Matching(line As Variant, Name As String, ByRef Dest_InpLst As String, PinTyp As String, MaxCnt As Long) As Boolean
'--------------------------------------------------------------------------------------------------------------------------------------------------
' ToDo:
' - Noch mal prüfen ob alle Pins möglich sind
'   Evtl. gibt es auch HW Kombinationen welche verhindern dass bestimmte Pins benutzt werden
'   - A1 kann z.B. dann benutzt werden wenn CAN Benutzt wird oder wenn Kein DCC und Selectrix benutzt ist
'     Evtl. eine Meldung ausgeben
'   - Wenn die SPI Kommunikation zum DCC Arduino verwendet wird, dann können die Pins 10, 11, 12 nicht benutzt werden
  Dim ValidPins As String, SPI_Pins As String, UseA1 As String
  
  ' FIX nach Umstellung auf Platform_Parameters                                ' 17.11.21: Juergen
  ' beim AM328 sind die SPI Pins nur frei, wenn kein CAN Modul angeschlossen ist
  SPI_Pins = ""
  If Page_ID <> "CAN" Then
    If PinTyp = "I" Or PinTyp = "O" Or PinTyp = "Pu" Then
        SPI_Pins = Get_Current_Platform_String("SPI_Pins")                     ' ~08.10.21: Juergen: New function to handle the valid pins. Prior this was handled here
    End If
  End If
  ValidPins = Get_Current_Platform_String(PinTyp + "_Pins", True)              ' ~08.10.21: Juergen:    "
  ValidPins = SPI_Pins + ValidPins
  If Left(line, Len(Name)) = Name Then
     Dim p As Long, NrStr As String
     p = InStr(line, ")")
     If p = 0 Then GoTo PrintError
     NrStr = Mid(line, 1 + Len(Name), p - 1 - Len(Name))
     NrStr = Trim(Replace(NrStr, ",", " "))
     NrStr = Replace(NrStr, "  ", " ")
     If NrStr = "" Then GoTo PrintError
     Dim NrArr() As String, OnePin As Variant
     NrArr = Split(NrStr, " ")
     If UBound(NrArr) + 1 > MaxCnt Then GoTo PrintError
     ' Check if valid pins names / numbers are used
     
     NrStr = ""
     For Each OnePin In NrArr
        If InStr(ValidPins, " " & AliasToPin(OnePin) & " ") = 0 Then                   ' 14.10.21: Juergen
           MsgBox Replace(Replace(Get_Language_Str("Fehler: Der Pin '#1#' ist nicht gültig im" & vbCr & _
           "  '#2#' Befehl"), "#1#", OnePin), "#2#", Replace(line, "// ", "")), _
           vbCritical, Get_Language_Str("Ungültige Arduino Pin Nummer")
           Exit Function
        End If
        ' Check Duplicate Pins
        p = InStr(" " & line & " ", " " & AliasToPin(OnePin) & " ")                     ' 14.10.21: Juergen, 16.05.20: Added space around OnePin (Problem: 2 ... 12)
        If InStr(p + 1, " " & line & " ", " " & OnePin & " ") > 0 Then      '     "        "
           MsgBox Replace(Replace(Get_Language_Str("Fehler: Der Pin '#1#' wird mehrfach verwendet im" & vbCr & _
           "  '#2#' Befehl"), "#1#", OnePin), "#2#", Replace(line, "// ", "")), _
           vbCritical, Get_Language_Str("Mehrfach verwendeter Arduino Pin")
           Exit Function
        End If
        If NrStr <> "" Then NrStr = NrStr + " "                                         ' 14.10.21: Juergen
        NrStr = NrStr + AliasToPin(OnePin)                                              ' build a new list with logical names mapped to physical pins
     Next OnePin
     
     Dest_InpLst = NrStr
  End If
  Set_PinNrLst_if_Matching = True
  If NrStr <> "" Then Debug.Print "Set_PinNrLst_if_Matching(" & Name & "=" & NrStr & ")" ' Debug
  Exit Function
  
PrintError:
  MsgBox Replace(Get_Language_Str("Fehler beim Lesen der Pin Nummern in Zeile:" & vbCr & _
                                  "  '#1#'"), "#1#", line), vbCritical, Get_Language_Str("Fehler beim Lesen der Pin Nummern")
End Function

'---------------------------------------------------------
Private Function Get_Arguments(line As String) As String()
'---------------------------------------------------------
  Dim Arguments As String, Parts() As String, i As Long, p As Long
  If InStr(line, "(") = 0 Then
    MsgBox "Error: Opening bracket not found in '" & line & "'", vbCritical, "Internal Error"
    Exit Function
  End If
  Arguments = Split(line, "(")(1)
  p = InStrRev(Arguments, ")")
  If p = 0 Then
    MsgBox "Error: Closing bracket not found in '" & line & "'", vbCritical, "Internal Error"
    Exit Function
  End If
  Arguments = Left(Arguments, p - 1)
  Parts = Split(Arguments, ",")
  For i = 0 To UBound(Parts)
    Parts(i) = Trim(Parts(i))
  Next i
  Get_Arguments = Parts
End Function

'UT-----------------------------
Private Sub Test_Get_Arguments()
'UT-----------------------------
  Dim Res() As String
  Res = Get_Arguments("Test( A, b, c)")
End Sub

'------------------------------------------------------------------------------------------------------
Private Function Get_Matching_Arg(Org_Macro As String, line As String, DestVarName As String) As String
'------------------------------------------------------------------------------------------------------
' Return the argument in "Line" which matches DestVarName in Org_Macro
   Dim Org_Args() As String, Act_Args() As String
   Org_Args = Get_Arguments(Org_Macro)
   If isInitialised(Org_Args) Then
      Act_Args = Get_Arguments(line)
      If isInitialised(Act_Args) Then
         If UBound(Act_Args) >= UBound(Org_Args) Then ' Org_Args may contain ... => the number of Act_Args may be greater
            Dim i As Long
            For i = 0 To UBound(Org_Args)
               If Org_Args(i) = DestVarName Then
                  Select Case DestVarName
                     Case "...", "OutList":
                                            While i <= UBound(Act_Args)
                                               Get_Matching_Arg = Get_Matching_Arg & Act_Args(i) & ","
                                               i = i + 1
                                            Wend
                                          Get_Matching_Arg = DelLast(Get_Matching_Arg)
                     Case Else: Get_Matching_Arg = Act_Args(i)
                  End Select
                  Exit Function
               End If
            Next
         End If
      End If
   End If
   MsgBox Replace(Get_Language_Str("Fehler bei der Erkennung der Zielvariable in Makro '#1#'"), "#1#", line), vbCritical, Get_Language_Str("Fehler: Zielvariable wurde nicht gefunden")
End Function


'UT--------------------------------
Private Sub Test_Get_Matching_Arg()
'UT--------------------------------
  Debug.Print Get_Matching_Arg("Random(        DstVar, InCh, RandMode, MinTime, MaxTime, MinOn, MaxOn)", "Random( OutA, 1, 2, 3, 4, 5, 6)", "DstVar")
  Debug.Print Get_Matching_Arg("Counter(       CtrMode, InCh, Enable, TimeOut, ...)", "Counter(12, #InCh, Enable, TimeOut, OutA, OutB, OutB)", "...")
  Debug.Print Get_Matching_Arg("RandMux(       DstVar1, DstVarN, InCh, RandMode, MinTime, MaxTime)", "RandMux( Out1, Out10, InCh, RandMode, MinTime, MaxTime)", "DstVarN")
End Sub


'-------------------------------------------------------------------------------
Public Function Add_Variable_to_DstVar_List(ByVal VarName As String) As Boolean
'-------------------------------------------------------------------------------
  Dim Check As String
  Check = " " & VarName & " "
  If InStr(DstVar_List, Check) = 0 Then
        DstVar_List = DstVar_List & VarName & " "
  Else
       If InStr(MultiSet_DstVar_List, Check) = 0 Then MultiSet_DstVar_List = MultiSet_DstVar_List & VarName & " "
  End If
  Add_Variable_to_DstVar_List = True
End Function

'------------------------------------------------------------------------------------------------------------------
Private Function Add_Matching_Arg_to_DstVars(Org_Macro As String, line As String, DestVarName As String) As Boolean
'------------------------------------------------------------------------------------------------------------------
' Locate DestVarName in Org_Macro and add the corrosponding
' argument to the global string DstVar_List
' Example:
'   MonoFlop(DstVar, InCh, Duration)
'   RS_FlipFlop2(DstVar1, DstVar2, InCh, R_InCh)             Called 2 times
  Dim Arg As String
  Arg = Get_Matching_Arg(Org_Macro, line, DestVarName)
  If Arg <> "" Then
     If Arg = "#LocInCh" Then                                               ' 20.06.20: Prevent problems with the random goto activation: Random(#LocInCh, #InCh, RM_NORMAL, 5 Sek,  10 Sek, 1 ms, 1 ms)
           Add_Matching_Arg_to_DstVars = True
     Else: Add_Matching_Arg_to_DstVars = Add_Variable_to_DstVar_List(Arg)
     End If
  End If
End Function

'--------------------------------------------------------------------------------
Public Function Add_Variable_to_StoreVar_List(ByVal VarName As String) As Boolean    ' 01.05.20: Jürgen
'--------------------------------------------------------------------------------
  Dim Check As String
  Check = " " & VarName & " "
  If InStr(StoreVar_List, Check) = 0 Then
        StoreVar_List = StoreVar_List & VarName & " "
  End If
  Add_Variable_to_StoreVar_List = True
End Function

'------------------------------------------------------------------------
Public Function StoreVar_List_Present(ByVal VarName As String) As Boolean            ' 01.05.20: Jürgen
'------------------------------------------------------------------------
  Dim Check As String
  Check = " " & VarName & " "
  StoreVar_List_Present = InStr(StoreVar_List, Check) <> 0
End Function

#If 1 Then
'-----------------------------------------------------------------------------
Private Function Get_Nr_From_Var(Name As String, ByRef TxtLen As Long) As Long       ' 27.12.20:
'-----------------------------------------------------------------------------
  Get_Nr_From_Var = -1
  Dim p As Long
  p = Len(Name)
  While p > 0 And InStr("0123456789", Mid(Name, p, 1)) > 0
    p = p - 1
  Wend
  If p > 0 Then
     p = p + 1
     TxtLen = p - 1
     If IsNumeric(Mid(Name, p)) Then
           Get_Nr_From_Var = val(Mid(Name, p))
     Else: Get_Nr_From_Var = -2
     End If
  End If
End Function

#Else
'-----------------------------------------------------------------------------
Private Function Get_Nr_From_Var(Name As String, ByRef TxtLen As Long) As Long
'-----------------------------------------------------------------------------
  Get_Nr_From_Var = -1
  Dim p As Long
  p = 1
  While p < Len(Name) And InStr("0123456789", Mid(Name, p, 1)) = 0
    p = p + 1
  Wend
  If p <= Len(Name) Then
     TxtLen = p - 1
     If IsNumeric(Mid(Name, p)) Then
           Get_Nr_From_Var = val(Mid(Name, p))
     Else: Get_Nr_From_Var = -2
     End If
  End If
End Function
#End If
'-------------------------------------------------------------------------------------
Private Function Add_N2_Arg_to_DstVars(Org_Macro As String, line As String) As Boolean
'-------------------------------------------------------------------------------------
' Example: RandMux(DstVar1, DstVarN, InCh, RandMode, MinTime, MaxTime)
  Dim Arg1 As String, ArgN As String, TxtLen1 As Long, TxtLenN As Long
  Arg1 = Get_Matching_Arg(Org_Macro, line, "DstVar1")
  If Arg1 <> "" Then
     ArgN = Get_Matching_Arg(Org_Macro, line, "DstVarN")
     If ArgN <> "" Then
        Dim StartNr As Long, EndNr As Long, i As Long
        StartNr = Get_Nr_From_Var(Arg1, TxtLen1)
        If StartNr < 0 Then Exit Function
        EndNr = Get_Nr_From_Var(ArgN, TxtLenN)
        If EndNr < 0 Then Exit Function
        If TxtLen1 <> TxtLenN Or Left(Arg1, TxtLen1) <> Left(ArgN, TxtLenN) Then Exit Function
        For i = StartNr To EndNr
            If Add_Variable_to_DstVar_List(Left(Arg1, TxtLen1) & i) = False Then Exit Function
        Next i
        Add_N2_Arg_to_DstVars = True
     End If
  End If
End Function

'----------------------------------------------------------------------------------------
Private Function Add_VarArgCnt_to_DstVars(Org_Macro As String, line As String) As Boolean
'----------------------------------------------------------------------------------------
' Example: Counter(CtrMode, InCh, Enable, TimeOut, ...)
  Dim Arg As String
  
  ' 20.06.20:
  If Left(Org_Macro, Len("Counter(")) = "Counter(" Then ' If the LED output is disabled the first DestVar contains the destination
     If InStr(line, "CF_ONLY_LOCALVAR") > 0 Then        ' count (Counter => 0 .. n-1)
        Add_VarArgCnt_to_DstVars = True                 ' => There are no Dest vars ==> We don't have to add them
        Exit Function
     End If
  End If
  
  Arg = Get_Matching_Arg(Org_Macro, line, "OutList")   ' Old: "...")
  If Arg <> "" Then
     Dim Name As Variant
     For Each Name In Split(Arg, ",")
         If Not Add_Variable_to_DstVar_List(Name) Then Exit Function
     Next Name
     Add_VarArgCnt_to_DstVars = True
  End If
End Function

'---------------------------------------------------------------------------------------------------------
Private Function Add_Cx_to_DstVars(Org_Macro As String, line As String, Cnt As Long, r As Long) As Boolean
'---------------------------------------------------------------------------------------------------------
' Example: PushButton_w_LED_0_2(B_LED, B_LED_Cx, InCh, DstVar1, Rotate, Timeout)
  Dim Arg1 As String, TxtLen As Long
  Arg1 = Get_Matching_Arg(Org_Macro, line, "DstVar1")
  If Arg1 <> "" Then
     Dim StartNr As Long, EndNr As Long, i As Long
     StartNr = Get_Nr_From_Var(Arg1, TxtLen)
     If StartNr < 0 Then
        MsgBox Replace(Replace(Replace(Get_Language_Str("Fehler: Die Zielvariable '#1#' in Zeile #2# muss eine Zahl am Ende haben " & _
                       "weil sie Teil einer Sequenz ist." & vbCr & _
                       "  Beispiel: #3#"), _
                       "#1#", Arg1), "#2#", r), "#3#", Arg1 & "0"), vbCritical, Get_Language_Str("Fehler: Zielvariable ungültig für Sequenz")
        Exit Function
     End If
     EndNr = StartNr + Cnt - 1
     For i = StartNr To EndNr
         If Add_Variable_to_DstVar_List(Left(Arg1, TxtLen) & i) = False Then Exit Function
     Next i
     Add_Cx_to_DstVars = True
  End If
End Function

'------------------------------------------------------------------------------
Public Function Add_Inp_and_DstVars(ByVal line As String, r As Long) As Boolean
'------------------------------------------------------------------------------
' Following types of macros are defined which generate DstVar's (One example per typ)
'
'  1   Logic(           DstVar, ...)                                         o.k.
'  2   MonoFlop2(       DstVar1, DstVar2, InCh, Duration)                    o.k.
'  n…  Counter(         CtrMode, InCh, Enable, TimeOut, ...)                 o.k.
'  n2  RandMux(DstVar1, DstVarN, InCh, RandMode, MinTime, MaxTime)           o.k.
'
  Add_Inp_and_DstVars = True
  If InStr(line, "(") = 0 Then Exit Function
  
  Dim Parts() As String, p As Long, Org_Macro_Row As Long, Arguments As String, Res As Boolean
  Parts = Split(line, "(")
  Arguments = Parts(1)
  p = InStrRev(Arguments, ")")
  If p = 0 Then Exit Function
    
  Dim SearchMacro As String                                                 ' 10.06.20: Find also HouseT lines
  If Parts(0) = "HouseT" Then
        SearchMacro = "House"
  Else: SearchMacro = Parts(0)
  End If
  Org_Macro_Row = Find_Macro_in_Lib_Macros_Sheet(SearchMacro & "(")
  If Org_Macro_Row = 0 Then
     Debug.Print "Attention: Macro '" & line & " not found in '" & LIBMACROS_SH & "'" ' In case the user has defined own macros some where
     ' ToDo: Wie können Zielvariablen in diesen Makros erkannt werden?
  Else
     Dim OutCntStr As String, Org_Macro As String, Org_Arguments As String
     With Sheets(LIBMACROS_SH)
        OutCntStr = .Cells(Org_Macro_Row, SM_OutCntCOL)
        Org_Macro = .Cells(Org_Macro_Row, SM_Macro_COL)
     End With
     If Add_InpVars(Parts(0), Org_Macro, line, r, Org_Macro_Row) = False Then
        Add_Inp_and_DstVars = False
        Exit Function
     End If
     Select Case OutCntStr
        Case "", "0": Add_Inp_and_DstVars = True
                      Exit Function
        Case "1":     Res = Add_Matching_Arg_to_DstVars(Org_Macro, line, "DstVar")               ' Ex.: MonoFlop(DstVar, InCh, Duration)
        Case "2":     Res = Add_Matching_Arg_to_DstVars(Org_Macro, line, "DstVar1")              ' Ex.: RS_FlipFlop2(DstVar1, DstVar2, InCh, R_InCh)
                      If Res Then Res = Add_Matching_Arg_to_DstVars(Org_Macro, line, "DstVar2")
        Case "n..":   Res = Add_VarArgCnt_to_DstVars(Org_Macro, line)                            ' Ex.: Counter(CtrMode, InCh, Enable, TimeOut, ...)
        Case "n2":    Res = Add_N2_Arg_to_DstVars(Org_Macro, line)                               ' Ex.: RandMux(DstVar1, DstVarN, InCh, RandMode, MinTime, MaxTime)
        Case Else:    ' Other Output Count entries
                      If Left(OutCntStr, 1) = "C" Then ' C1, C2, .. Cn                             Ex.: PushButton_w_LED_0_2(B_LED, B_LED_Cx, InCh, DstVar1, Rotate, Timeout)
                         If IsNumeric(Mid(OutCntStr, 2)) Then
                            Res = Add_Cx_to_DstVars(Org_Macro, line, val(Mid(OutCntStr, 2)), r)
                         End If
                      Else
                           MsgBox "Internal Error: Undefined OutCnt entry '" & OutCntStr & "' in row " & Org_Macro_Row & " in sheet '" & LIBMACROS_SH & "'", vbCritical, "Internal Error"
                           Exit Function
                      End If
     End Select
     If Res = False Then
        Cells(r, Config__Col).Select
        MsgBox Replace(Replace(Get_Language_Str("Fehler in der Definition der Zielvariable(n): '#1#' in Zeile #2#"), "#1#", line), "#2#", r), vbCritical, Get_Language_Str("Fehler in Makro Definition")
        Add_Inp_and_DstVars = False
     End If
  End If
End Function


'-----------------------------------------------------------------
Public Sub Print_DstVar_List(fp As Integer, ByRef Channel As Long)
'-----------------------------------------------------------------
  Dim Var As Variant
  For Each Var In Split(Trim(DstVar_List), " ")
      Print #fp, "#define " & AddSpaceToLen(Var, 22) & "  " & Channel
      Channel = Channel + 1
  Next Var
End Sub


'------------------------------------------------------------------------------------------------------------------------------------------
Public Function Is_in_Nr_String(ByVal Name As String, ByVal ExpectedName As String, MinNr As Long, MaxNr As Long, ByRef Nr As Long) As Long  ' 03.04.20:
'------------------------------------------------------------------------------------------------------------------------------------------
' Check if "Name" starts with "ExpectedName" and has a tailing number in the range form MinNr to MaxNr
  Dim SkipCnt As Long
  While Right(ExpectedName, 1) = "?"
     SkipCnt = SkipCnt + 1
     ExpectedName = DelLast(ExpectedName)
  Wend
  If Left(Name, Len(ExpectedName)) = ExpectedName Then
     Dim NrStr As String, i As Long, WrongChar As String
     NrStr = Mid(Name, 1 + Len(ExpectedName) + SkipCnt)
     If IsNumeric(NrStr) Then
        If Left(NrStr, 1) = "0" Then Exit Function ' Leading 0 are not allowed because they generate the same number
        WrongChar = "-+.,eE"
        For i = 1 To Len(WrongChar)
            If InStr(NrStr, Mid(WrongChar, i, 1)) > 0 Then Exit Function
        Next i
        Nr = val(NrStr)
        If Nr >= MinNr And Nr <= MaxNr Then
              Is_in_Nr_String = 1
        Else: Is_in_Nr_String = -1 ' Exists, but invalide number
        End If
     End If
     
  End If
End Function

'UT-------------------------------
Private Sub Test_Is_in_Nr_String()
'UT-------------------------------
  Dim Nr As Long
  Debug.Print "Is_in_Nr_String=" & Is_in_Nr_String("HalloAB13", "Hallo??", 1, 200, Nr)
End Sub

'-----------------------------------------------------------------------
Private Sub Error_Msg_Varaible_Not_Defined(VarName As String, r As Long)    ' 03.04.20:
'-----------------------------------------------------------------------
  Cells(r, Get_Address_Col()).Select
  MsgBox Replace(Replace(Get_Language_Str("Fehler: Die Variable '#1#' in Zeile #2# ist nicht definiert"), "#1#", VarName), "#2#", r), vbCritical, Get_Language_Str("Fehler: Undefinierte Variable")
End Sub



'---------------------------------------------------------------------------
Public Function Valid_Var_Name(ByVal Name As String, Row As Long) As Boolean                    ' 03.04.20:
'---------------------------------------------------------------------------
' Check if it's a valid variable (Switch<Nr>, Button<Nr>, INCH_DCC_..., LOC_INCH<Nr>, ...)
  Const Std_Names = "SI_1 SI_0 SI_Enable_Sound #LocInCh" ' 20.06.20: Added: "#LocInCh" to prevent problems with the Goto activation Random which uses "Counter(CF_ONLY_LOCALVAR | CF_ROTATE | CF_SKIP0, #LocInCh, #InCh, 0 Sec, 2)"
  Valid_Var_Name = True
  Dim Nr As Long, Std_N As Variant
  
  For Each Std_N In Split(Std_Names, " ")
    If Name = Std_N Then Exit Function
  Next
  
  Select Case Is_Switch_Var_then_Add_to_Ctr(Name)
     Case 1:  Exit Function
     Case -1: Valid_Var_Name = False
              Cells(Row, Config__Col).Select
              Exit Function
  End Select
      
  If InStr(DstVar_List, " " & Name & " ") Then Exit Function
  
  If InStr(InChTxt, "#define " & Name & " ") > 0 Then Exit Function
  Select Case Is_in_Nr_String(Name, "LOC_INCH", 0, 250, Nr)
     Case 1:  ' Valid Number
              If Nr > MaxUsed_Loc_InCh Then
                 MaxUsed_Loc_InCh = Nr       ' The "LOC_INCH" are checked later because they are not generated at this time
                 MaxUsed_Loc_InCh_Row = Row
              End If
              Exit Function
     Case -1: MsgBox Replace(Replace(Get_Language_Str("Fehler: Die Nummer der Variable '#1#' ist ungültig!" & vbCr & _
                                     vbCr & _
                                     "Gültiger Bereich: #2#"), "#1#", Name), "#2#", "0..250"), vbCritical, _
                                     Get_Language_Str("Fehler: Ungültige Variable")
  End Select
  ' Old:
  '  Error_Msg_Varaible_Not_Defined Name, Row
  '  Valid_Var_Name = False
  
  ' Add the undefined input variable to a list which is checked later
  If InStr(Undefined_Input_Var, " " & Name & " ") = 0 Then
     Undefined_Input_Var = Undefined_Input_Var & Name & " "
     Undef_Input_Var_Row = Undef_Input_Var_Row & Row & " "
  End If
End Function

'-----------------------------------------------------------------------------------------------------
Private Function Valid_Var_Name_and_Skip_InCh_and_Numbers(ByVal Arg As String, Row As Long) As Boolean
'-----------------------------------------------------------------------------------------------------
  Dim SubArgList() As String, SubArg As Variant
  SubArgList = SplitMultiDelims(Arg, " +-")
  For Each SubArg In SubArgList
      If SubArg = "#InCh" Then
          If Get_InCh_Number_w_Err_Msg(Arg) < 0 Then Exit Function   ' Check the whole argumnet to make sure that the equation contains only constants
      Else
          If Not IsNumeric(SubArg) Then  ' Skip special names like '#InCh' and Numbers
             If Valid_Var_Name(SubArg, Row) = False Then Exit Function
          End If
      End If
  Next SubArg
  Valid_Var_Name_and_Skip_InCh_and_Numbers = True
End Function

'-------------------------------------------------------------------------------------------------
Public Sub Create_Loc_InCh_Defines(ByRef Dest As String, ByRef Channel As Long, LocInChNr As Long)
'-------------------------------------------------------------------------------------------------
  If LocInChNr > 0 Then
     Dest = Dest & vbCr & "// Local InCh variables" & vbCr
     Dim i As Long
     For i = 0 To LocInChNr - 1
         Dest = Dest & AddSpaceToLen("#define LOC_INCH" & i, 32) & Channel & vbCr
         Channel = Channel + 1
     Next i
  End If
  
  If MaxUsed_Loc_InCh >= LocInChNr Then
     Error_Msg_Varaible_Not_Defined "LOC_INCH" & MaxUsed_Loc_InCh, MaxUsed_Loc_InCh_Row
     Exit Sub
  End If
End Sub

'-------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Print_Keyboard_Defines_for_Type(fp As Integer, Name As String, InpCnt As Long, ByRef Nr As Long, Optional Skip_11_16 As Boolean)    ' 02.04.20:
'-------------------------------------------------------------------------------------------------------------------------------------------
  Dim i As Long
  For i = 1 To InpCnt
      Print #fp, AddSpaceToLen("#define " & Name & i, 32) & Nr
      If Skip_11_16 Then
         If (i - 1 Mod 16) = 9 Then Nr = Nr + 6 ' Skip the IONr 11 and 16 because they are not used for the analog switches, but have to be reserved for MobaLedLib_Copy_to_InpStruct()
      End If
      Nr = Nr + 1
  Next i
  
  ' Es werden immer vielfache von 8 Inp Channels belegt
  
  Dim ResCnt As Long
  ResCnt = (8 - (InpCnt Mod 8)) Mod 8                                       ' 02.05.20: Old: (8 - InpCnt) Mod 8
  If ResCnt > 0 Then
     Nr = Nr + ResCnt
     Print #fp, "// Reserve channels: " & ResCnt & " because MobaLedLib_Copy_to_InpStruct always writes multiple of 8 channels"
  End If
End Sub

'UT----------------------
Private Sub Test_ResCnt()                                                   ' 02,05.20:
'UT----------------------
' Test for the problem above
Dim InpCnt As Long, ResCnt As Long
  Debug.Print "i Ok  Err"
  For InpCnt = 0 To 24
      ResCnt = (8 - (InpCnt Mod 8)) Mod 8
      Debug.Print Left(InpCnt & "   ", 3) & Left(ResCnt & "   ", 3) & (8 - InpCnt) Mod 8
  Next InpCnt
End Sub


'--------------------------------------------------------------------------
Private Sub Make_sure_that_Channel_is_divisible_by_4(ByRef Channel As Long)
'--------------------------------------------------------------------------
  If Channel Mod 4 <> 0 Then ' Attention: Channel number must be divisible by 4 for Copy_Bits_to_InpStructArray()
     Channel = Channel + 4 - Channel Mod 4
  End If
End Sub


'-------------------------------------------------------------------------------------------------
Public Function Write_Switches_Header_File_Part_A(fp As Integer, ByRef Channel As Long) As Boolean
'-------------------------------------------------------------------------------------------------

' #If USE_SWITCH_AND_LED_ARRAY Then                                         ' 04.11.20:
'    Print #fp, "#define USE_SWITCH_AND_LED_ARRAY 1    // Enable the new function which handles the SwitchD and the Mainboard LEDs in the ino file"
' #Else
'    Print #fp, "#define USE_SWITCH_AND_LED_ARRAY 0"
'    If Get_BoardTyp() = "ESP32" Then                                       ' 04.11.20:
'        MsgBox "Internal error: The compiler switch 'USE_SWITCH_AND_LED_ARRAY' must be defined if the ESP32 is used", vbCritical, "Internal error"
'        'EndProg
'     End If
'  #End If
'  Print #fp, ""
  
  Dim Ana_But_Pin_Array() As String, ACh As Long, Used_AButton_Channels As Long, Start_AButtons As Long
  If SwitchA_InpCnt > 0 Or Read_LDR Then
    Ana_But_Pin_Array = Split(SwitchA_InpLst, " ")
    If SwitchA_InpCnt > (UBound(Ana_But_Pin_Array) + 1) * 10 Then
       MsgBox Get_Language_Str("Fehler: Es wurden mehr analoge Taster verwendet als möglich sind. " & _
                               "Es müssen weitere analoge Eingänge zum einlesen definiert werden." & vbCr & _
                               "Das wird mit dem Befehl 'Set_SwitchA_InpLst()' in der Makro Spalte gemacht."), _
                               vbCritical, Get_Language_Str("Fehler: Nicht genügend analoge Eingänge zum einlesen der Taster definiert")
       Exit Function
    End If
    
    Print #fp, "//*** Analog switches ***"
    Print #fp, ""
    If Get_BoardTyp() = "AM328" Then
      If Make_Sure_that_AnalogScanner_Library_Exists() = False Then Exit Function
      Print #fp, "#include <AnalogScanner.h>   // Interrupt driven analog reading library. The library has to be installed manually from https://github.com/merose/AnalogScanner"
      Print #fp, "AnalogScanner scanner;       // Creates an instance of the analog pin scanner."
    ElseIf Get_BoardTyp() = "ESP32" Then
      Print #fp, "#include ""AnalogScannerESP32.h""   "
      Print #fp, "AnalogScannerESP32 scanner;       // Creates an instance of the analog pin scanner."
    End If
    Print #fp, ""
    Print #fp, "#include <Analog_Buttons10.h>"
    Used_AButton_Channels = WorksheetFunction.RoundUp(SwitchA_InpCnt / 10, 0)
    For ACh = 1 To Used_AButton_Channels
        Print #fp, "Analog_Buttons10_C AButtons" & ACh & "(" & Ana_But_Pin_Array(ACh - 1) & ");"
    Next ACh
    Print #fp, ""
    Print #fp, ""

    If Switch_Damping_Fact <> "" Then                                       ' 04.11.21:
       Print #fp, Switch_Damping_Fact
       Print #fp, ""
    End If
    If Read_LDR Then
       Print #fp, "#include ""Read_LDR.h""     // Darkness sensor"
       Print #fp, ""
    End If

    Make_sure_that_Channel_is_divisible_by_4 Channel
    Start_AButtons = Channel
    Dim TmpChannel As Long
    TmpChannel = Channel
    Print_Keyboard_Defines_for_Type fp, "SwitchA", SwitchA_InpCnt, TmpChannel, Skip_11_16:=True
    Channel = Channel + Used_AButton_Channels * 16  ' multiply by 16 because the Copy_Bits_to_InpStructArray() always fills bytes
    Print #fp, ""
  End If ' SwitchA_InpCnt > 0 Or Read_LDR
    
  If Channel1InpCnt > 0 Then
    Print #fp, "//*** Digital switches ***"
    Print #fp, ""
    ' Generate the #define Switch... statements
    Dim StartSwitches1 As Long, StartSwitches2 As Long
    Make_sure_that_Channel_is_divisible_by_4 Channel
    StartSwitches1 = Channel
    If Channel1InpCnt > 0 Then Print_Keyboard_Defines_for_Type fp, CTR_Cha_Name_1, Channel1InpCnt, Channel
    
    
    Make_sure_that_Channel_is_divisible_by_4 Channel
    StartSwitches2 = Channel
    If Channel2InpCnt > 0 Then Print_Keyboard_Defines_for_Type fp, CTR_Cha_Name_2, Channel2InpCnt, Channel
    Print #fp, ""
  End If
  
  If SwitchD_InpCnt > (UBound(Split(SwitchD_InpLst, " ")) + 1) Then         ' 04.11.20: Moved out of the following if because SwitchD_InpCnt should also be checked if USE_SWITCH_AND_LED_ARRAY is enabled
     ' Todo: Activate the corrosponding cell. Therefore a list has to be generated where each switch is used the first time
     MsgBox Replace(Get_Language_Str("Fehler: Es wurden mehr SwitchD Schalter verwendet als Pins definiert sind. " & _
                               "Es müssen weitere Eingänge zum einlesen definiert werden." & vbCr & _
                               "Das wird mit dem Befehl 'Set_SwitchD_InpLst()' in der Makro Spalte gemacht." & vbCr & _
                               "Letzter möglicher Schalter: 'SwitchD#1#'"), "#1#", UBound(Split(SwitchD_InpLst, " ")) + 1), _
                               vbCritical, Get_Language_Str("Fehler: Nicht genügend Eingänge zum einlesen der Schalter definiert")
     Exit Function
  End If
  If SwitchD_InpCnt > 0 Then
    Print #fp, "//*** Direct connected switches ***"
    Print #fp, ""
    
    Print_Keyboard_Defines_for_Type fp, "SwitchD", SwitchD_InpCnt, Channel
    Print #fp, ""
  End If ' SwitchD_InpCnt > 0
  
  If DstVar_List <> " " Then ' #defines for the Output variables
    Print #fp, "//*** Output Channels ***"
    Print_DstVar_List fp, Channel
    Print #fp, ""
  End If
  
  If SwitchD_InpCnt > 0 Then
    Print #fp, "const PROGMEM uint8_t SwitchD_Pins[] = " & AddSpaceToLen("{ " & Replace(SwitchD_InpLst, " ", ",") & " };", 28) & "// Array of pins which read switches 'D'"
    Print #fp, "#define SWITCH_D_INP_CNT sizeof(SwitchD_Pins)"
    Print #fp, ""
  End If
    
  If Channel1InpCnt > 0 Then
    Print #fp, ""
    Print #fp, "#define CTR_CHANNELS_1    " & AddSpaceToLen(CTR_Channels_1, 41) & "// Number of used counter channels for keyboard 1. Up to 10 if one CD4017 is used, up to 18 if two CD4017 are used, ..."
    Print #fp, "#define CTR_CHANNELS_2    " & AddSpaceToLen(CTR_Channels_2, 41) & "// Number of used counter channels for keyboard 2. Up to 10 if one CD4017 is used, up to 18 if two CD4017 are used, ..."
    Print #fp, "#define BUTTON_INP_LIST_1 " & AddSpaceToLen(Replace(But_Inp_List_1, " ", ","), 41) & "// Comma separated list of the button input pins"
    Print #fp, "#define BUTTON_INP_LIST_2 " & AddSpaceToLen(Replace(But_Inp_List_2, " ", ","), 41) & "// Comma separated list of the button input pins"
    Print #fp, "#define CLK_PIN           " & AddSpaceToLen(CLK_Pin_Number, 41) & "// Pin number used for the CD4017 clock"
    Print #fp, "#define RESET_PIN         " & AddSpaceToLen(RST_Pin_Number, 41) & "// Pin number used for the CD4017 reset"
    Print #fp, ""
    Print #fp, "#include <Keys_4017.h>                                             // Keyboard library which uses the CD4017 counter to save Arduino pins. Attention: The pins (CLK_PIN, ...) must be defined prior."
    Print #fp, ""
    Print #fp, "#define START_SWITCHES_1  " & AddSpaceToLen(StartSwitches1, 41) & "// Define the start number for the first keyboard."
    Print #fp, "#define START_SWITCHES_2  " & AddSpaceToLen(StartSwitches2, 41) & "// Define the start number for the second keyboard."
    Print #fp, ""
  End If
  
  Dim dmxDefines As String                                                  ' 19.01.21 Juergen
  If True Then
    ' FastLED initialisation                                                ' 26.04.20:
    Dim LEDCh As Long, LED_PINNr_Arr() As String, Cnt As Long, ExpOutPins As Long, DMX_Pin_Number As String
    LED_PINNr_Arr = Split(LED_PINNr_List, " ")
    Print #fp, "/*********************/"
    Print #fp, "#define SETUP_FASTLED()                                                      \"
    Print #fp, "/*********************/                                                      \"
    For LEDCh = 0 To LED_CHANNELS - 1
        If LEDs_per_Channel(LEDCh) > 0 Then
           ExpOutPins = LEDCh
           If LEDCh <= UBound(LED_PINNr_Arr) Then
             If LEDCh <> DMX_LedChan Then                                   ' 19.01.21 Juergen
              ' Generate: CLEDController& controller0 = FastLED.addLeds<NEOPIXEL,  6 >(leds+   0, 200); \"
                If Use_WS2811 Then                                          ' 19.01.21 Juergen
                  Print #fp, "  CLEDController& controller" & LEDCh & " = FastLED.addLeds<WS2811, " & AddSpaceToLenLeft(LED_PINNr_Arr(LEDCh), 2) & ", RGB>(leds+" & AddSpaceToLenLeft(Cnt, 3) & "," & AddSpaceToLenLeft(LEDs_per_Channel(LEDCh), 3) & "); \" ' 19.01.21 Juergen
                Else
                  Print #fp, "  CLEDController& controller" & LEDCh & " = FastLED.addLeds<NEOPIXEL, " & AddSpaceToLenLeft(LED_PINNr_Arr(LEDCh), 2) & ">(leds+" & AddSpaceToLenLeft(Cnt, 3) & "," & AddSpaceToLenLeft(LEDs_per_Channel(LEDCh), 3) & "); \"
                End If
            Else                                                            ' 19.01.21 Juergen
                dmxDefines = "#define DMX_LED_OFFSET " & Cnt & vbCrLf$ & _
                             "#define DMX_CHANNEL_COUNT " & LEDs_per_Channel(LEDCh) * 3
                DMX_Pin_Number = LED_PINNr_Arr(LEDCh)
                If (LEDs_per_Channel(LEDCh) > 100) Then
                    MsgBox Get_Language_Str("Fehler: Das DMX Senden ist auf 100 Leds (300 DMX Kanäle) limitiert."), vbCritical, Application.ActiveSheet.Name
                    Exit Function
                End If
            End If
           End If
        End If
        Cnt = Cnt + LEDs_per_Channel(LEDCh)
    Next
    If ExpOutPins > UBound(LED_PINNr_Arr) Then
       MsgBox Replace(Get_Language_Str("Fehler: Es sind nicht genügend Ausgangs Pins zur Ansteuerung der LEDs vorhanden. " & _
              "Die LED Pins müssen mit dem Befehl ""Set_LED_OutpPinLst()"" definiert werden." & vbCr & _
              "Es müssen #1# Arduino Ausgänge definiert sein."), "#1#", ExpOutPins + 1), vbCritical, _
              Get_Language_Str("Mehr LED Gruppen verwendet als LED Ausgangspins definiert")
       Exit Function
    End If
    Print #fp, "                                                                             \"
    For LEDCh = 0 To LED_CHANNELS - 1
        If LEDs_per_Channel(LEDCh) > 0 And LEDCh <= UBound(LED_PINNr_Arr) And (LEDCh <> DMX_LedChan) Then  ' 19.01.21 Juergen
           Print #fp, "  controller" & LEDCh & ".clearLeds(256);                                                \"
        End If
    Next
    Print #fp, "  FastLED.setDither(DISABLE_DITHER);       // avoid sending slightly modified brightness values"   ' 05.03.21: Juergen
    Print #fp, "/*End*/"
    Print #fp, ""
    
    If DMX_Pin_Number <> "" And dmxDefines <> "" Then                       ' 19.01.21 Juergen
        Print #fp, "#include ""DmxInterface.h""     // DMX512 Interface"
        Print #fp, "#define USE_DMX_PIN " + DMX_Pin_Number
        Print #fp, dmxDefines
        Print #fp, "DMXInterface dmxInterface;"
    End If
  End If
  
  ' Additional Setup proc
  If SwitchA_InpCnt > 0 Or Read_LDR Or Channel1InpCnt > 0 Or SwitchD_InpCnt > 0 Then
    Print #fp, "#define USE_ADDITIONAL_SETUP_PROC                                  // Activate the usage of the Additional_Setup_Proc()"
    Print #fp, ""
    Print #fp, "//--------------------------"
    Print #fp, "void Additional_Setup_Proc()"
    Print #fp, "//--------------------------"
    Print #fp, "{"
    If SwitchA_InpCnt > 0 Or Read_LDR Then
       Dim PinList As String
       If SwitchA_InpCnt > 0 Then PinList = Replace(Trim(SwitchA_InpLst), " ", ",") & ","
       If Read_LDR Then PinList = PinList & LDR_Pin_Number & ","
       PinList = DelLast(PinList)
       Print #fp, "  int scanOrder[] = {" & PinList & "};"
       Print #fp, "  const int SCAN_COUNT = sizeof(scanOrder) / sizeof(scanOrder[0]);"
       If Get_BoardTyp() = "AM328" Then
         Print #fp, ""
         If Read_LDR Then
            Print #fp, "  Init_DarknessSensor(" & LDR_Pin_Number & ", 50, SCAN_COUNT); // Attention: The analogRead() function can't be used together with the darkness sensor !"
            Print #fp, "  scanner.setCallback(" & LDR_Pin_Number & ", Darkness_Detection_Callback);"
         End If
         Print #fp, "  scanner.setScanOrder(SCAN_COUNT, scanOrder);"
         Print #fp, "  scanner.beginScanning();"
       ElseIf Get_BoardTyp() = "ESP32" Then
         Print #fp, "  scanner.setScanPins(SCAN_COUNT, scanOrder);"
         If Read_LDR Then
            Print #fp, "  Init_DarknessSensor(" & LDR_Pin_Number & ", 50, 50); // Attention: The analogRead() function can't be used together with the darkness sensor !"
            Print #fp, "  scanner.setCallback(" & LDR_Pin_Number & ", Darkness_Detection_Callback);"
         End If
       End If
       Print #fp, ""
    End If
    If Channel1InpCnt > 0 Then
       Print #fp, "  Keys_4017_Setup(); // Initialize the keyboard scanning process"
    End If
    
    If SwitchD_InpCnt > 0 Then
       Print #fp, ""
       Print #fp, "  for (uint8_t i = 0; i < SWITCH_D_INP_CNT; i++)"
       Print #fp, "    pinMode(pgm_read_byte_near(&SwitchD_Pins[i]), INPUT_PULLUP);"
    End If
    Print #fp, "}"
    Print #fp, ""
  End If
  
  ' Generate the "Additional_Loop_Proc()"
  If SwitchA_InpCnt > 0 Or Channel1InpCnt > 0 Or SwitchD_InpCnt > 0 Then
    Print #fp, "/****************************/"
    Print #fp, "#define Additional_Loop_Proc() \"
    Print #fp, "/****************************/ \" ' Attention the function is called with a delay of 3 seconds to wait
    Print #fp, "{                              \" ' until the SPI pins are disabled if they are not used by the program
    If SwitchA_InpCnt > 0 Then
       Dim Act_Start_AButtons As Long
       Act_Start_AButtons = Start_AButtons
       Print #fp, "  uint16_t Button;             \"
       For ACh = 1 To Used_AButton_Channels
           Print #fp, AddSpaceToLen("  Button = AButtons" & ACh & ".Get(); MobaLedLib_Copy_to_InpStruct((uint8_t*)&Button, 2, " & Act_Start_AButtons & ");", 89) & "\"
           Act_Start_AButtons = Act_Start_AButtons + 16 ' add 16 because 2 Bytes add 16 channels
       Next ACh
    End If
    If Channel1InpCnt > 0 Then Print #fp, "  MobaLedLib_Copy_to_InpStruct(Keys_Array_1, KEYS_ARRAY_BYTE_SIZE_1, START_SWITCHES_1);  \"
    If Channel2InpCnt > 0 Then Print #fp, "  MobaLedLib_Copy_to_InpStruct(Keys_Array_2, KEYS_ARRAY_BYTE_SIZE_2, START_SWITCHES_2);  \"
    If SwitchD_InpCnt > 0 Then
       Print #fp, "  for (uint8_t i = 0; i < " & SwitchD_InpCnt & "; i++) \"
       Print #fp, "      MobaLedLib.Set_Input(SwitchD1 + i, !digitalRead(pgm_read_byte_near(&SwitchD_Pins[i])));\"
    End If
    Print #fp, "}"
  End If
  Write_Switches_Header_File_Part_A = True
End Function
  
  
Public Function Write_LowProrityLoop_Header_File(fp As Integer) As Boolean
  If Serial_PinLst <> "" Then
    Print #fp, "/*****************************/"
    Print #fp, "#define Additional_Loop_Proc2() \"
    Print #fp, "/*****************************/ \" ' This function is called in every loop, on ESP32 in the alternate loop (not time critical)
    Print #fp, "{                               \"
    If Serial_PinLst <> "" Then
       Print #fp, "   soundProcessor.process();\"               ' 02.11.2021: Juergen add support of multiple sound module types
    End If
    Print #fp, "}"
  End If
  
  Write_LowProrityLoop_Header_File = True
End Function


'-----------------------------------------------------------------------------------------------------------------------------------------
Private Function No_Duplicates_in_two_InpLists(Letter1 As String, Letter2_or_Name As String, ByVal InpLst1 As String, _
                                               ByVal InpLst2 As String, Optional Set_Funct2 As String = "Set_LED_OutpPinLst()") As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------
' Retutn True if no duplicates are detected
  InpLst2 = " " & InpLst2 & " "
  Dim Pin As Variant
  For Each Pin In Split(InpLst1, " ")
     If InStr(InpLst2, " " & Pin & " ") > 0 Then
        If Len(Letter2_or_Name) = 1 Then
            MsgBox Replace(Replace(Replace(Get_Language_Str("Fehler: Der Eingabe Pin '#1#' des Arduinos wird als Eingabe in den zwei Schalter Funktionen '#2#' und '#3#' benutzt." & vbCr & _
                    "Die beiden Schalter Funktionen können nicht gleichzeitig benutzt werden. Die Pins können mit den Funktionen " & _
                    "'Set_Switch?_InpLst()' angepasst werden." & vbCr & _
                    "Achtung: Dazu muss auch die Hardware angepasst werden!"), _
                    "#1#", Pin), "#2#", Letter1), "#3#", Letter2_or_Name), vbCritical, "Fehler: Doppelte benutzung der Eingangs Pins"
        Else
            MsgBox Replace(Replace(Replace(Replace(Get_Language_Str("Fehler: Der Eingabe Pin '#1#' des Arduinos wird als Eingabe in der Schalter Funktionen '#2#' und gleichzeitig " & _
                    "als #3# Pin benutzt." & vbCr & _
                    "Die Pins können mit den Funktionen 'Set_Switch?_InpLst()' und '#4#' angepasst werden." & vbCr & _
                    "Achtung: Dazu muss auch die Hardware angepasst werden!"), _
                    "#1#", Pin), "#2#", Letter1), "#3#", Letter2_or_Name), "#4#", Set_Funct2), vbCritical, "Fehler: Doppelte benutzung der Arduino Pins"
       End If
        Exit Function
     End If
  Next
  No_Duplicates_in_two_InpLists = True
End Function

Public Function No_Duplicates_in_two_Lists(Pin2 As String, ByVal InpLst1 As String, _
                                               ByVal InpLst2 As String, Optional Set_Funct2 As String = "Set_LED_OutpPinLst()") As Boolean
'-----------------------------------------------------------------------------------------------------------------------------------------
' Retutn True if no duplicates are detected
  InpLst2 = " " & InpLst2 & " "
  Dim Pin As Variant
  For Each Pin In Split(InpLst1, " ")
     If InStr(InpLst2, " " & Pin & " ") > 0 Then
     
        MsgBox Replace(Replace(Replace(Get_Language_Str("Fehler: Der Pin '#1#' in '#3#' wird bereits als '#2#' Pin benutzt." & vbCr & _
                    "Der Pin kann nicht mehrfach benutzt werden."), _
                    "#1#", Pin), "#2#", Pin2), "#3#", Set_Funct2), vbCritical, "Fehler: Doppelte Benutzung eines Pins"
        Exit Function
     End If
  Next
  No_Duplicates_in_two_Lists = True
End Function


'UT------------------------------------------------
Private Sub Test_No_Duplicates_in_two_InpLists()
'UT------------------------------------------------
  SwitchC_InpLst = "2 7 8 9 10 11 12 A5"
  SwitchD_InpLst = "7 8 9"
  'Debug.Print No_Duplicates_in_two_InpLists("C", "D", SwitchC_InpLst, SwitchD_InpLst)
  
  LED_PINNr_List = "6 A4 A5"
  Debug.Print No_Duplicates_in_two_InpLists("C", "LED", SwitchC_InpLst, LED_PINNr_List)
End Sub

'------------------------------------------------------------------------------------------------
Private Function Check_CLK_a_RST_Pin_Usage(Letter1 As String, ByVal InpLst1 As String) As Boolean          ' 04.11.20:
'------------------------------------------------------------------------------------------------
' Check the RST_Pin_Number together with Letter1
  If SwitchB_InpCnt > 0 Or SwitchC_InpCnt > 0 Then
     If False = No_Duplicates_in_two_InpLists(Letter1, Replace(Get_Language_Str("#1# Pin für SwitchB oder SwitchC"), "#1#", "Reset"), InpLst1, RST_Pin_Number, "Set_RST_Pin_Number()") Then Exit Function
     If False = No_Duplicates_in_two_InpLists(Letter1, Replace(Get_Language_Str("#1# Pin für SwitchB oder SwitchC"), "#1#", "Clock"), InpLst1, CLK_Pin_Number, "Set_CLK_Pin_Number()") Then Exit Function
  End If
  Check_CLK_a_RST_Pin_Usage = True
End Function


'-----------------------------------------------------
Public Function No_Duplicates_in_InpLists() As Boolean
'-----------------------------------------------------
  If SwitchA_InpCnt > 0 Then
     If SwitchB_InpCnt > 0 Then
        If No_Duplicates_in_two_InpLists("A", "B", SwitchA_InpLst, SwitchB_InpLst) = False Then Exit Function
     End If
     If SwitchC_InpCnt > 0 Then
        If No_Duplicates_in_two_InpLists("A", "C", SwitchA_InpLst, SwitchC_InpLst) = False Then Exit Function
     End If
     If SwitchD_InpCnt > 0 Then
        If No_Duplicates_in_two_InpLists("A", "D", SwitchA_InpLst, SwitchD_InpLst) = False Then Exit Function
     End If
     If No_Duplicates_in_two_InpLists("A", "LED", SwitchA_InpLst, LED_PINNr_List) = False Then Exit Function
     If Read_LDR Then
        If No_Duplicates_in_two_InpLists("A", "LDR", SwitchA_InpLst, LDR_Pin_Number, "Set_LDR_Pin_Number()") = False Then Exit Function
     End If
     If Check_CLK_a_RST_Pin_Usage("A", SwitchA_InpLst) = False Then Exit Function                 ' 04.11.20:
     If No_Duplicates_in_two_InpLists("A", "SOUND", SwitchA_InpLst, Serial_PinLst, SF_SERIAL_SOUND_PIN) = False Then Exit Function   ' 08.10.21: Juergen
  End If
  
  If SwitchB_InpCnt > 0 Then
     If SwitchC_InpCnt > 0 Then
        If No_Duplicates_in_two_InpLists("B", "C", SwitchB_InpLst, SwitchC_InpLst) = False Then Exit Function
     End If
     If SwitchD_InpCnt > 0 Then
        If No_Duplicates_in_two_InpLists("B", "D", SwitchB_InpLst, SwitchD_InpLst) = False Then Exit Function
     End If
     If No_Duplicates_in_two_InpLists("B", "LED", SwitchB_InpLst, LED_PINNr_List) = False Then Exit Function
     If Read_LDR Then
        If No_Duplicates_in_two_InpLists("B", "LDR", SwitchB_InpLst, LDR_Pin_Number, "Set_LDR_Pin_Number()") = False Then Exit Function
     End If
     If Check_CLK_a_RST_Pin_Usage("B", SwitchB_InpLst) = False Then Exit Function                 ' 04.11.20:
     If No_Duplicates_in_two_InpLists("B", "SOUND", SwitchA_InpLst, Serial_PinLst, SF_SERIAL_SOUND_PIN) = False Then Exit Function   ' 08.10.21: Juergen
  End If
  
  If SwitchC_InpCnt > 0 Then
     If SwitchD_InpCnt > 0 Then
        If No_Duplicates_in_two_InpLists("C", "D", SwitchC_InpLst, SwitchD_InpLst) = False Then Exit Function
     End If
     If No_Duplicates_in_two_InpLists("C", "LED", SwitchC_InpLst, LED_PINNr_List) = False Then Exit Function
     If Read_LDR Then
        If No_Duplicates_in_two_InpLists("C", "LDR", SwitchC_InpLst, LDR_Pin_Number, "Set_LDR_Pin_Number()") = False Then Exit Function
     End If
     If Check_CLK_a_RST_Pin_Usage("C", SwitchC_InpLst) = False Then Exit Function                 ' 04.11.20:
     If No_Duplicates_in_two_InpLists("C", "SOUND", SwitchA_InpLst, Serial_PinLst, SF_SERIAL_SOUND_PIN) = False Then Exit Function   ' 08.10.21: Juergen
  End If
  
  If SwitchD_InpCnt > 0 Then
     If No_Duplicates_in_two_InpLists("D", "LED", SwitchD_InpLst, LED_PINNr_List) = False Then Exit Function
     If Read_LDR Then
        If No_Duplicates_in_two_InpLists("D", "LDR", SwitchD_InpLst, LDR_Pin_Number, "Set_LDR_Pin_Number()") = False Then Exit Function
     End If
     If Check_CLK_a_RST_Pin_Usage("D", SwitchD_InpLst) = False Then Exit Function                 ' 04.11.20:
     If No_Duplicates_in_two_InpLists("D", "SOUND", SwitchA_InpLst, Serial_PinLst, SF_SERIAL_SOUND_PIN) = False Then Exit Function   ' 08.10.21: Juergen
  End If
  
  No_Duplicates_in_InpLists = True
End Function



'---------------------------------------------------------
Public Function Init_HeaderFile_Generation_SW() As Boolean
'---------------------------------------------------------
  MaxUsed_Loc_InCh = -1
  Read_LDR = False
  Store_Status_Enabled = False                                              ' 19.01.21 Juergen
  Use_WS2811 = False                                                        '   "
  
  ' The following variables are read from the data lines
  SwitchA_InpCnt = 0       ' Number of switch imputs for channel A (A=Analog push buttons) Maximal switch number is detected for this and the next two lines.
  SwitchB_InpCnt = 0       ' Number of switch imputs for channel B (B=Border:  push buttons (or switches) around the border of the layout
  SwitchC_InpCnt = 0       ' Number of switch imputs for channel C (C=Console: switches (or push buttons) in the console (Weichenstellpult)
  SwitchD_InpCnt = 0       ' Number of switch imputs for channel D (D=Direct:  switches (or push buttons) connected direct to the main board
  
  DMX_LedChan = -1                                                          ' 19.01.21 Juergen
  Serial_PinLst = ""                                                        ' 08.10.21 Juergen add Serial Sound feature, by default no serial sound
  
  LED_PINNr_List = Get_Current_Platform_String("LED_Pins")                  ' ~08.10.21: Juergen: New function to handle the valid pins. Prior this was handled here
  LDR_Pin_Number = Get_Current_Platform_String("LDR_Pin")
  SwitchA_InpLst = Get_Current_Platform_String("SwitchA_Pins")
  SwitchB_InpLst = Get_Current_Platform_String("SwitchB_Pins")
  SwitchC_InpLst = Get_Current_Platform_String("SwitchC_Pins")
  SwitchD_InpLst = Get_Current_Platform_String("SwitchD_Pins")
  CLK_Pin_Number = Get_Current_Platform_String("CLK_Pin")
  RST_Pin_Number = Get_Current_Platform_String("RST_Pin")
  
  DstVar_List = " "
  MultiSet_DstVar_List = " "
  
  If Not First_Scan_of_Data_Rows() Then Exit Function ' Scan the data rows and fill the variables above if the corrosponding functions are used in the Config__Col
  
  If MultiSet_DstVar_List <> " " Then
     If MsgBox(Get_Language_Str("Achtung: Die folgenden Zielvariablen werden mehrfach gesetzt:") & vbCr & MultiSet_DstVar_List & vbCr & _
               Get_Language_Str("Senden zum Arduino abbrechen?"), vbQuestion + vbYesNo, Get_Language_Str("Warnung: Mehrfach benutzte Zielvariablen")) = vbYes Then
               Exit Function
     End If
  End If
  
  If No_Duplicates_in_InpLists() = False Then Exit Function
  
   
  CTR_Channels_1 = 0
  CTR_Channels_2 = 0
  But_Inp_List_1 = "Unused"
  But_Inp_List_2 = "Unused"
  Channel1InpCnt = 0
  Channel2InpCnt = 0

  
  If SwitchB_InpCnt > 0 And SwitchC_InpCnt > 0 Then
     CTR_Cha_Name_1 = "SwitchB"
     CTR_Cha_Name_2 = "SwitchC"
     CTR_Channels_1 = WorksheetFunction.RoundUp(SwitchB_InpCnt / (UBound(Split(SwitchB_InpLst, " ")) + 1), 0)
     CTR_Channels_2 = WorksheetFunction.RoundUp(SwitchC_InpCnt / (UBound(Split(SwitchC_InpLst, " ")) + 1), 0)
     But_Inp_List_1 = SwitchB_InpLst
     But_Inp_List_2 = SwitchC_InpLst
     Channel1InpCnt = SwitchB_InpCnt
     Channel2InpCnt = SwitchC_InpCnt
  ElseIf SwitchB_InpCnt > 0 Then
     CTR_Cha_Name_1 = "SwitchB"
     CTR_Cha_Name_2 = "Unused"
     CTR_Channels_1 = WorksheetFunction.RoundUp(SwitchB_InpCnt / (UBound(Split(SwitchB_InpLst, " ")) + 1), 0)
     But_Inp_List_1 = SwitchB_InpLst
     Channel1InpCnt = SwitchB_InpCnt
  ElseIf SwitchC_InpCnt > 0 Then
     CTR_Cha_Name_1 = "SwitchC"
     CTR_Cha_Name_2 = "Unused"
     CTR_Channels_1 = WorksheetFunction.RoundUp(SwitchC_InpCnt / (UBound(Split(SwitchC_InpLst, " ")) + 1), 0)
     But_Inp_List_1 = SwitchC_InpLst
     Channel1InpCnt = SwitchC_InpCnt
  End If
  Init_HeaderFile_Generation_SW = True
End Function

'----------------------------------------------------------------------
Private Sub Find_and_Select_Name(ByVal Name As String, UndefNr As Long)
'----------------------------------------------------------------------
  Dim UndefRow As Long
  UndefRow = Split(Undef_Input_Var_Row, " ")(UndefNr)
  Cells(UndefRow, Get_Address_Col()).Select
  On Error Resume Next ' In case it was not found for some reasons
  Rows(UndefRow).Find(What:=Name, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, _
             SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False).Activate
  On Error GoTo 0
End Sub



'----------------------------------------------------
Public Function Check_Detected_Variables() As Boolean
'----------------------------------------------------
' Check undefined input variables
' ToDo:
'  - Check unused or double written desination variables
'  - Variablen müssen geschrieben werden bevor sie gelesen werden sonst wird eine Änderung nicht erkannt.
'    Das ist eigentlich klar => In die Doku
'
  Undefined_Input_Var = Trim(Undefined_Input_Var)
  If Undefined_Input_Var <> "" Then
     Dim UnDefVar As Variant, Found As Boolean, UndefNr As Long
     For Each UnDefVar In Split(Undefined_Input_Var, " ")
         Found = (InStr(InChTxt, " " & UnDefVar & " ") <> 0)                ' 24.04.20: Search also in the list of DCC,SX,CAN defines
         If Not Found Then Found = (InStr(DstVar_List, UnDefVar) <> 0)
         If Not Found And UnDefVar = "[Multiplexer]" Then Found = True      ' Added by Misha 30-5-2020.  ' 14.06.20: Added from Mishas version
         If Not Found Then
            Find_and_Select_Name UnDefVar, UndefNr
            MsgBox Replace(Get_Language_Str("Fehler: Die Variable '#1#' wird als Eingang benutzt, wird aber nirgendwo gesetzt."), "#1#", UnDefVar), _
                   vbCritical, Get_Language_Str("Fehler: Undefinierter Zustand eine Eingangsvariablen")
            Exit Function
         End If
         UndefNr = UndefNr + 1
     Next UnDefVar
  End If
  Check_Detected_Variables = True
End Function




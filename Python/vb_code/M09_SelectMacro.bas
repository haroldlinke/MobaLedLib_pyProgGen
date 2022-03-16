Attribute VB_Name = "M09_SelectMacro"
Option Explicit

' - Bei Effekten welche einzelne LEDs ansteuern muss die Adressierung der
'   LEDs anders gemacht werden.
' - Die LED Nummer darf nicht immer nach einer Zeile erhöht werden.
' - Wenn danach eine RGB Zeile kommt, dann wird die die Nummer erhöht
' - Es muss auch möglich sein, dass mehrere Zeilen den selben LED Kanal ansprechen
'   z.B. bei der Sound Funktion. => Es gibt keine Überprüfung auf doppelte Belegung
' - Mit der "Kommetar" Funktion "End_Single_LEDs"  kann man die Nummer erhöhen kann wenn danach wieder eine
'   Funktion kommt welche einzelne LEDs anspricht.
'   - Das ist alles zu Kompliziert
'   - Es währe besser wenn automatisch die nächste RGB LED Nummer gewählt würde.
'   - Nur bei Funktionen wie der Sound Funktion benötigt man einen Befehl zum weiter schalten
' - Pattern Configurator: Startkanal, Anzahl der Kanäle
' - Eintrag in der LEDs Spalte: Andreaskreuz: C1 => 1-2, C2 => 2-3, C3 => 3-4
'   hier sollen auch größere Nummern möglich sein 5-6
'   Eine größere Startnummer benötigt man dann wenn eine Patternfunktion z.B. 4 LEDs benutzt.
'   Dann kann man mit 5-6 die letzten beiden LEDs des zweiten WS2811 ansprechen.
'   Das ginge aber auch mit NextLED und 2-3

' Neuer Ansatz:
' - Die LED Kanäle müssen wie beim Haus immer in aufsteigender Reihenfolge angegeben werden
'   Wenn eine kleinere Nummer als die Vorangegangene verwendet wird, dann wird damit das nächste
'   WS2811 Modul angesprochen.
' - Funktionen wie die Sound Befehle bekommen ein Flag mit dem verhindert wird, das
'   das nächste Modul verwendet wird. In der Tabelle kann man das so markieren: "^ C1-2"
'   Wenn ein Befehl ausgewählt wurde bei dem die Kanäle mit den vorangegangenen Überlappen,
'   Dann wird der User gefragt ob er die gleiche StartLedNr verwenden will wie die vorangegange Zeile.
'   Das funktioniert aber nur wenn einzel Adressierte LEDs Verwendet werden (C1-2)
' - Befehle welche die RGB LEDs am Stück und nicht als drei einzelne Kanäle ansprechen sorgen immer
'   dafür dass die nächste StartLedNr verwendet wird.


Private Const HeadRow = 3

'--------------------------------------------------------------------------------------------
Private Function Special_ConstrWarnLight(ByVal Res As String, ByRef LEDs As String) As String
'--------------------------------------------------------------------------------------------
' ConstrWarnLight(LED,InCh, LEDcnt, MinBrightness, MaxBrightness, OnT, WaitT, WaitE)
'          Param: 0   1     2       3              4              5    6      7
  Dim Parts As Variant, Param As Variant, Ret As String, FlashFunct As Boolean, LED_Channel As String
  LED_Channel = Split(Res, "$")(1)                                          ' 17.05.20:
  Parts = Split(Replace(Split(Res, "$")(0), ")", ""), "(")                  '    "      Added Split..."$"
  Param = Split(Parts(1), ",")
  If val(Param(6)) > 0 Then
        FlashFunct = True
        Ret = "ConstrWarnLightFlash"
  Else: Ret = "ConstrWarnLight"
  End If
  Ret = Ret & Trim(Param(2)) & "(" & Trim(Param(0)) & ", " & _
                               Trim(Param(1)) & ", " & _
                               Trim(Param(3)) & ", " & _
                               Trim(Param(4)) & ", " & _
                               Trim(Param(5)) & ", "
  If FlashFunct Then Ret = Ret & Trim(Param(6)) & ", "

  Ret = Ret & Trim(Param(7)) & ")" & "$" & LED_Channel                      ' 17.05.20: Added: LED_CHannel
  Special_ConstrWarnLight = Ret
  LEDs = "C1-" & Trim(Param(2))
End Function

'UT---------------------------------------
Private Sub Test_Special_ConstrWarnLight()
'UT---------------------------------------
  Dim Res As String, LEDs As String
       ' ConstrWarnLight(LED,InCh, LEDcnt, MinBrightness, MaxBrightness, OnT, WaitT, WaitE)
  Res = "ConstrWarnLight(#LED,#InCh, 6, 20, 255, 100 ms, 0 ms, 300 ms)"
  Res = Special_ConstrWarnLight(Res, LEDs)
  Debug.Print Res & "LEDs:" & LEDs
End Sub

'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Proc_General_With_Other_Par(ByVal Macro As String, Description As String, LedChannels As Long, Show_Channel As Byte, LED_Channel As Long, Def_Channel As Long) As String ' 27.04.20: Added: Show_LED_Channel, LED_Channel and Def_Channel
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If Macro = "" Then Exit Function
  
  Dim Parts As Variant, Res As String, Param As Variant
  Parts = Split(Replace(Macro, ")", ""), "(")
  Param = Split(Parts(1), ",")
  
  ' IF statement added by Misha 18-4-2020                                   ' 14.06.20: Added from Mishas version
  If left(Macro, Len("Multiplexer")) = "Multiplexer" Then   ' Select which UserForm to show
      UserForm_Create_Multiplexer.Show_UserForm_Other Parts(1), Parts(0), Description, LedChannels
  Else
      UserForm_Other.Show_UserForm_Other Parts(1), Parts(0), Description, LedChannels, Show_Channel, LED_Channel, Def_Channel ' 27.04.20: Added: Show_LED_Channel, LED_Channel and Def_Channel
  End If
  ' End IF statement added by Misha 18-4-2020
   
  Proc_General_With_Other_Par = Userform_Res
End Function

'-------------------------------------------------------------------------------------------------------------
Private Function Get_NamedPar(SearchName As String, MacroWithNames As String, FilledMacro As String) As String
'-------------------------------------------------------------------------------------------------------------
  Dim Names As Variant, Values As Variant, i As Long
  Names = Split(Replace(Split(MacroWithNames, "(")(1), ")", ""), ",")
  Values = Split(Replace(Split(FilledMacro, "(")(1), ")", ""), ",")
  For i = 0 To UBound(Names)
      If Trim(Names(i)) = SearchName Then
         Get_NamedPar = Trim(Values(i))
         Exit Function
      End If
  Next i
End Function

'------------------------------------------------------------------------------
Private Function Cx_to_LED_Channel(Cx As String, LedChannels As Long) As String
'------------------------------------------------------------------------------
  Select Case Cx
    Case "C_ALL": Cx_to_LED_Channel = "1"
    Case "C1":    Cx_to_LED_Channel = "C1-" & 1 + Abs(LedChannels) - 1  ' Negativ LedChannels is used to fource exact one LED in the PushButon function
    Case "C2":    Cx_to_LED_Channel = "C2-" & 2 + Abs(LedChannels) - 1
    Case "C3":    Cx_to_LED_Channel = "C3-" & 3 + Abs(LedChannels) - 1
    Case "C12":   Cx_to_LED_Channel = "C1-2"
    Case "C23":   Cx_to_LED_Channel = "C2-3"
  End Select
End Function


'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Proc_General(LEDs As String, ByVal Macro As String, Description As String, LedChannels As Long, LED_Channel As Long, Def_Channel As Long) As String ' 27.04.20: Added: LED_Channel and Def_Channel
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------
  If Macro = "" Then Exit Function
  
  Dim Parts As Variant, Res As String, Param As Variant
  Parts = Split(Replace(Macro, ")", ""), "(")
  If UBound(Parts) = 0 Then ' Do we have parameters starting with "("
       Res = Macro ' Some macros like '#define READ_LDR" have no parameters ' 06.04.20:
  Else ' Process the parameters in ()
    Param = Split(Parts(1), ",")
    Res = Parts(0) & "("
    Dim p As Variant, Other As Long
    For Each p In Param
       p = Trim(p)
       Select Case p
         Case "LED", "B_LED":   p = "#LED"
         Case "InCh":           p = "#InCh"
         Case "":               ' Nothing
         Case Else:             Other = Other + 1
       End Select
       Res = Res & p & ", "
    Next p
    If UBound(Param) >= 0 Then ' We don't have to delete the tailing ", " if there are no parameters
       Res = DelLast(DelLast(Res))
    End If
    
    Dim Show_Channel As Byte                                                                      ' 07.10.21: Juergen, add channel Type serial
    Show_Channel = CHAN_TYPE_NONE
    If Trim(LEDs) <> "" Then
       If Trim(LEDs) = SerialChannelPrefix Then
            Show_Channel = CHAN_TYPE_SERIAL
       ElseIf IsNumeric(LEDs) Then
             If val(LEDs) >= 0 Then Show_Channel = CHAN_TYPE_LED                              ' 19.01.21: Jürgen: Old ".. <> 0"  07.10.21:Juergen
       Else: Show_Channel = CHAN_TYPE_LED
       End If
    End If
    
    Res = Res & ")"
    If Other > 0 Or Show_Channel <> CHAN_TYPE_NONE Then ' Other parameters then "LED" and "InCh"                   ' 27.04.20: Added: Show_LED_Channel
          Res = Proc_General_With_Other_Par(Res, Description, LedChannels, Show_Channel, LED_Channel, Def_Channel) ' 27.04.20: Added: Show_LED_Channel, LED_Channel and Def_Channel
          
          If Res = "" Then Exit Function
          
          If Parts(0) = "ConstrWarnLight" Then Res = Special_ConstrWarnLight(Res, LEDs) ' 18.09.19
          If left(Parts(0), Len("Multiplexer")) = "Multiplexer" Then Res = Special_Multiplexer_Ext(Res, LEDs)          ' Added by Misha 2020-03-26 ' 14.06.20: Added from Mishas version
    End If
    
    Dim Res_LED_Channel As String                                           ' 27.04.20:
    If InStr(Res, "$") > 0 Then
       Res_LED_Channel = Split(Res, "$")(1)
       Res = Split(Res, "$")(0)
    End If
    
    
    If left(LEDs, Len("LedCnt")) = "LedCnt" Then ' Special treatement if "LEDs" starts with "LedCnt" (Used in "Reserve LEDs" and "Next_LED")
       LEDs = Get_NamedPar(LEDs, Macro, Res)
    End If
    
    Select Case LEDs
       Case "Cx": ' Fill then the LEDs Column
                  Dim par As String
                  par = Get_NamedPar("Cx", Macro, Res)
                  If par = "" Then par = Get_NamedPar("B_LED_Cx", Macro, Res) ' Used in PushButton_w_LED_BL_0..     ' 13.04.20:
                  LEDs = Cx_to_LED_Channel(par, LedChannels)
    End Select
  End If
  
  Proc_General = LEDs & "$" & Res
  If Res_LED_Channel <> "" Then Proc_General = Proc_General & "$" & Res_LED_Channel ' 27.04.20:
End Function

'-----------------------------------------------------
Private Function Get_PriorLine_LEDs(ByRef Row As Long)
'-----------------------------------------------------
' Return the LEDs cell of the first enabled row starting with Row
' Only lines which have an entry in the "LEDs" column are checked
' Row is set to the detected line
  While Row >= FirstDat_Row
    If Row_is_Achtive(Row) Then
       Get_PriorLine_LEDs = Cells(Row, LEDs____Col)
       Exit Function
    End If
    Row = Row - 1
  Wend
End Function

'-----------------------------------------------------------------------------------------------------------
Private Function Get_From_Input_Var(Macro As String, ByVal FilledMacro As String, ParName As String) As Long
'-----------------------------------------------------------------------------------------------------------
  Dim par As String
  par = Get_NamedPar(ParName, Macro, FilledMacro)
  If par = "" Then
     MsgBox "Interner Fehler in Get_From_Input_Var()", vbCritical, "Interner Fehler"
     EndProg
  End If
  Get_From_Input_Var = par
End Function


Private Sub Test()
  Debug.Print Evaluate("1+3")
End Sub

'----------------------------------------------------------------------
Public Function Get_InCh_Number_w_Err_Msg(ByVal Arg As String) As Long
'----------------------------------------------------------------------
  Dim Nr As Long
  Nr = 0
  On Error GoTo ErrMsg
  Nr = Evaluate(Replace(Arg, "#InCh", "0"))
  On Error GoTo 0
  If Nr < 0 Then GoTo ErrMsg
  Get_InCh_Number_w_Err_Msg = Nr
  Exit Function

ErrMsg:
  MsgBox Get_Language_Str("Fehler im Logischen Ausdruck. Es darf nur eine konstante positive Zahl zu '#InCh' addiert werden"), _
         vbCritical, Get_Language_Str("Fehler in Logic() Funktion")
  Get_InCh_Number_w_Err_Msg = -1
End Function



'------------------------------------------------------------------------------------------------------
Private Function Get_Number_of_used_InCh_in_Par(ByVal Statement As String, Mode As String) As Long  ' 10.04.20:
'------------------------------------------------------------------------------------------------------
' Gets the maximal number of "#InCh" which is used in an logic expression if Mode = "Logic":
' The example
'    Logic(TestOr, #InCh OR #InCh+1 OR SwitchA4)
' will return 2
'
' If Mode = "Comma" the staremet is a list of parameters separated by ","
'
  Statement = Trim(Statement)
  If right(Statement, 1) = ")" Then Statement = Trim(DelLast(Statement))
  Dim Arglist() As String, Arg As Variant, Nr As Long, MaxNr As Long
  Select Case Mode
     Case "Logic": Arglist = SplitEx(Statement, True, "OR", "AND", "NOT")
     Case "Comma": Arglist = SplitEx(Statement, True, ",")
     Case Else: MsgBox "Internal Error: Unknown Mode '" & Mode & "' in 'Get_Number_of_used_InCh_in_Par()'", vbCritical, "Internal Error"
                Exit Function
  End Select
  If isInitialised(Arglist) Then
     For Each Arg In Arglist
         If InStr(Arg, "#InCh") > 0 Then
            Nr = Get_InCh_Number_w_Err_Msg(Arg)
            #If 0 Then
            On Error GoTo ErrMsg
            Nr = Evaluate(Replace(Arg, "#InCh", "0"))
            On Error GoTo 0
            If Nr < 0 Then GoTo ErrMsg
            #End If
            If Nr < 0 Then Exit Function
            If Nr > MaxNr Then MaxNr = Nr
         End If
     Next Arg
     Get_Number_of_used_InCh_in_Par = MaxNr + 1
  End If
End Function

'UT----------------------------------------------
Private Sub Test_Get_Number_of_used_InCh_in_Par()
'UT----------------------------------------------
  Debug.Print Get_Number_of_used_InCh_in_Par(" NOT #InCh)", "Logic")
  Debug.Print Get_Number_of_used_InCh_in_Par("#InCh OR #InCh+1 OR SwitchA4", "Logic")
  Debug.Print Get_Number_of_used_InCh_in_Par("#InCh OR #InCh+2 OR SwitchA4", "Logic")
  Debug.Print Get_Number_of_used_InCh_in_Par("#InCh AND Bedigung1 OR #InCh AND Bedingung2", "Logic")
  Debug.Print Get_Number_of_used_InCh_in_Par("#InCh + Bedigung1 + 7", "Logic") ' => Fehler
End Sub

'---------------------------------------------------------
Private Function Get_BinSize(ByVal X As Double) As Integer
'---------------------------------------------------------
' Number of binary bits necessary for x different values
  Get_BinSize = Application.RoundUp(Log(X) / Log(2), 0)
End Function


'----------------------------------------
Public Function SelectMacros() As Boolean                                   ' 07.05.20:
'----------------------------------------
  Dim OldUpdating As Boolean
  OldUpdating = Application.ScreenUpdating
  Application.ScreenUpdating = False
  SelectMacros = SelectMacros_Sub
  Application.ScreenUpdating = OldUpdating
End Function

'-------------------------------------------------------------------------------------------------------------------
Public Sub Add_Icon_and_Name(SelRow As Long, DstRow As Long, Optional Sh As Worksheet, Optional NameOnly As Boolean) ' 22.10.21:
'-------------------------------------------------------------------------------------------------------------------
' SelRow: Row in the Lib_Macros sheet
  If LanName_Col > 0 And SelRow > 0 Then
     With ThisWorkbook.Sheets(LIBMACROS_SH)
        If Sh Is Nothing Then Set Sh = ActiveSheet
        Dim LanName As String ' Language specific name
        LanName = Get_Language_Text(SelRow, SM_LName_COL, Get_ExcelLanguage())
        If LanName = "" Then LanName = .Cells(SelRow, SM_Name__COL) ' Normal non language specific name if the entry ha no language name yet
        Dim OldEvents As Boolean
        OldEvents = Application.EnableEvents
        Application.EnableEvents = False
        Sh.Cells(DstRow, LanName_Col) = LanName
        Application.EnableEvents = OldEvents
        If NameOnly = False Then
            Del_one_Icon_in_IconCol DstRow, Sh
            Dim PicName As String, PicNamesArr() As String
            PicNamesArr = Split(.Cells(SelRow, SM_Pic_N_COL).Value, "|")
            If UBound(PicNamesArr) > 0 Then
               PicName = Trim(PicNamesArr(UBound(PicNamesArr)))
               Add_Icon PicName, DstRow, Sh
            End If
        End If
     End With
  End If
End Sub


'---------------------------------------------
Private Function SelectMacros_Sub() As Boolean
'---------------------------------------------
  Make_sure_that_Col_Variables_match
  
  With ActiveCell
    If .Row < FirstDat_Row Then
       MsgBox Get_Language_Str("Zur Auswahl des (Beleuchtungs-)Effekts muss eine Zeile in der" & vbCr & _
              "Tabelle ausgewählt sein bevor der Dialog Knopf betätigt wird."), vbInformation, Get_Language_Str("Falsche Zielposition ausgewählt")
       Exit Function
    End If
  End With
  Dim OldEvents As Boolean:  OldEvents = Application.EnableEvents
  Application.EnableEvents = False ' Prevent double call if the Dialog button is used   ' 22.10.21:
  Cells(ActiveCell.Row, Config__Col).Select
  Application.EnableEvents = OldEvents
  
  If Get_String_Config_Var("Use_TreeView_for_Macros") = "" Then             ' 06.10.21:
     If MsgBox(Get_Language_Str("Soll die neue Baumansicht zur Auswahl der Makros verwendet werden oder weiter " & _
                                "mit dem alten Listenbasierten Dialog gearbeitet werden?" & vbCr & _
                                "  Ja = Neue Baumansicht" & vbCr & _
                                "  Nein = Alte Listenansicht" & vbCr & _
                                "(Das kann nachträglich auf der 'Config' Seite geändert werden)"), vbQuestion + vbYesNo, _
                                Get_Language_Str("Welcher Makro Auswahl Dialog soll verwendet werden?")) = vbYes Then
           Set_String_Config_Var "Use_TreeView_for_Macros", "1"
     Else: Set_String_Config_Var "Use_TreeView_for_Macros", "0"
     End If
  End If
  
  Dim ActMacro As String
  ActMacro = Replace(Trim(Cells(ActiveCell.Row, Config__Col).Value), "HouseT(", "House(")  ' 10.06.20: Added special treatement for "HouseT" ' 04.11.21: Added: Trim( in case spaces have been added. In this case the macro was not detected
  If Get_Bool_Config_Var("Use_TreeView_for_Macros") Then
       Sort_for_TreeView_based_Makro
       SelectMacro_TreeView ActMacro
  Else
       ' Problem: Wenn der Dialog an der Stelle geöffnet wird an der sich die Maus befindet, dann wird das Element an der Maus Position augewählt ;-(
       ' Hab noch keine Idee wie ich das beheben kann. Die folgenden Zeilen helfen nicht
       '  Sleep 1500 ' Wait until the mouse is released
       '  mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
       '  cancel = True in Proc_DoubleCkick()
       Sort_for_List_based_Makro
       SelectMacros_Form.Show_SelectMacros_Form LIBMACROS_SH, ActMacro
  End If
  
  If SelectMacro_Res <> "" Then
    Dim MacroName As String, SelRow As Long, DlgTyp As String, Macro As String, LEDs As String, Res As String
    Dim Description As String, LedChannels As Long, ActLanguage As Integer, Def_Channel As Long, Act_Channel As Long
    ActLanguage = Get_ExcelLanguage()                                       ' 24.02.20:
    MacroName = Split(SelectMacro_Res, ",")(0)
    SelRow = Split(SelectMacro_Res, ",")(1)
    With ThisWorkbook.Sheets(LIBMACROS_SH)
      DlgTyp = .Cells(SelRow, SM_Typ___COL)
      LEDs = .Cells(SelRow, SM_LEDS__COL)
      Macro = .Cells(SelRow, SM_Macro_COL)
      LedChannels = val(.Cells(SelRow, SM_SngLEDCOL))
      Def_Channel = val(.Cells(SelRow, SM_DefCh_COL))                       ' 26.04.20:
      Act_Channel = val(Cells(ActiveCell.Row, LED_Cha_Col))
      
      Description = Replace(Get_Language_Text(SelRow, SM_DetailCOL, ActLanguage), "|", vbLf)     ' 21.10.21: Old: 'Description = Replace(.Cells(SelRow, SM_DetailCOL + ActLanguage * DeltaCol_Lib_Macro_Lang), "|", vbLf)      ' 24.02.20: added:  + ActLanguage * DeltaCol_Lib_Macro_Lang
      If Description = "" Then Description = Get_Language_Text(SelRow, SM_ShrtD_COL, ActLanguage) ' 21.10.21: Old: Description = .Cells(SelRow, SM_ShrtD_COL + ActLanguage * DeltaCol_Lib_Macro_Lang)  ' 24.02.20:    "
      Select Case DlgTyp
        Case "House":  'UserForm_House.SetMode MacroName
                       UserForm_House.Show_With_Existing_Data MacroName, Cells(ActiveCell.Row, Config__Col), Act_Channel, Def_Channel
                       Res = Userform_Res
        Case "ColTab": 'Cells(ActiveCell.Row, Descrip_Col).Activate ' Move the cursor to get out of the edit mode => Nicht mehr nötig da cancel = True in Proc_DoubleCkick() gesetzt ist
                       'Calculate
                       Open_MobaLedCheckColors_and_Insert_Set_ColTab_Macro
                       Exit Function
        Case "EX.Constructor", "EX.Macro", "":                                                      ' 31.01.22: Juergen add extensions
                       Res = Proc_General(LEDs, Macro, Description, LedChannels, Act_Channel, Def_Channel) ' Empty typ
        Case Else:     MsgBox "Unknown Dialog Typ '" & DlgTyp & "'", vbCritical, "Program Error: SelectMacros_Sub"
      End Select
      
      If Res <> "" Then
         ' If Left(Res, Len("$#define")) = "$#define" Then Res = Replace(Replace(Res, "(", "   "), ")", "") ' Remove the brackets     ' 14.01.20: ' 04.11.21: Commented because the bracets are necessary to parse the argument if the macro should be changed

         Dim Parts As Variant
         Parts = Split(Res, "$")
         Dim DstRow As Long
         DstRow = ActiveCell.Row
         LEDs = Parts(0)
         Dim LED_Channel As String, LEDs_Column As Long                     ' 27.04.20:
         If UBound(Parts) >= 2 Then
               LED_Channel = Parts(2)
         Else: LED_Channel = ""
         End If
         
         If Check_IsSingleChannelCmd(LEDs) Then ' If it's a command which adresses single channels we check if it uses the same channel as above
            Dim PriorLeds As String, PriorLine As Long, PriorChan As Long
            PriorLine = DstRow - 1
            PriorLeds = Trim(Replace(Get_PriorLine_LEDs(PriorLine), "^", "")) ' Get the first enabled line above
            PriorChan = val(Cells(PriorLine, LED_Cha_Col))
            If Check_IsSingleChannelCmd(PriorLeds) Then
               If PriorLeds = LEDs And PriorChan = val(LED_Channel) Then ' Same LEDs used again    ' 08.05.20: Added: And PriorChan = Val(LED_Channel)
                  Dim Answ As Variant
                  Answ = MsgBox(Replace(Get_Language_Str("Achtung: Die LED Kanäle sind gleich wie der vorangegangenen Zeile (#1#)!" & vbCr & _
                             "Das kann zu ungewollten Effekten führen." & vbCr & _
                             "Bei Funktionen welche einen Kanal nur kurzzeitig ansteuern kann das sinnvoll sein." & vbCr & _
                             "Ein Beispiel dafür ist die Ansteuerung von Sound Modulen. Hier steuern mehrere Tasten " & _
                             "den gleichen Kanal mit unterschiedlichen Werten an. Je nach abzuspielendem Sound wird " & _
                             "eine andere 'Helligkeit' gesendet. Da die Tasten werden aber nicht gleichzeitig betätigt " & _
                             "werden ist das unproblematisch." & vbCr & _
                             vbCr & _
                             "Soll der neue Befehl die gleiche LED Adressieren wie der Vorangegangene Befehl?"), "#1#", PriorLine), _
                             vbQuestion + vbYesNoCancel, Get_Language_Str("Überlappende Kanäle entdeckt"))
                  Select Case Answ
                     Case vbYes: LEDs = "^ " & LEDs
                     Case vbCancel: Exit Function
                  End Select
               End If
            End If
         End If
         
         Dim OldEvents1 As Boolean: OldEvents1 = Application.EnableEvents
         Application.EnableEvents = False
         Cells(DstRow, LED_Cha_Col) = LED_Channel
         Cells(DstRow, LEDs____Col) = LEDs
         Application.EnableEvents = OldEvents1
         
         Cells(DstRow, Config__Col) = Parts(1) ' Fill the Macro column
         
         Add_Icon_and_Name SelRow, DstRow                                   ' 22.10.21:
         
         Dim InCnt As Long
         Select Case .Cells(SelRow, SM_InCnt_COL)
             Case "n":         InCnt = Get_From_Input_Var(Macro, Parts(1), "InCh_Cnt")
             Case "States":    InCnt = Get_From_Input_Var(Macro, Parts(1), "States")
             Case "BinStates": InCnt = Get_BinSize(Get_From_Input_Var(Macro, Parts(1), "BinStates"))
             Case "Logic":     InCnt = Get_Number_of_used_InCh_in_Par(Split(Parts(1), ",")(1), "Logic")
             Case "2?":        InCnt = Get_Number_of_used_InCh_in_Par(Split(Parts(1), "(")(1), "Comma")   ' 29.04.20:   07.05.20: Old: "2". For exact match "2!" was used. Tha was not intuitiv. Now numbers always define exakt match.
                               If InCnt = 0 Then ' In case of an error we use the default number
                                  InCnt = 2
                               End If
             Case Else:        InCnt = val(.Cells(SelRow, SM_InCnt_COL)) ' use the Val function to get 0 if the cell is empty in sheet Lib_Macros
         End Select
         Application.EnableEvents = False                                   ' 07.05.20: Prevent drawing the buttons by event (It's called below again)
         Cells(DstRow, InCnt___Col) = InCnt
         Complete_Addr_Column_with_InCnt DstRow                             ' 07.05.20: Update the used adress range ("1" => "1-2" if InCnt = 3)
         Format_Cells_to_Row DstRow + SPARE_ROWS  ' Add some reserve lines  '    "
         Application.EnableEvents = OldEvents1
         
         ' Special Checks                                                   ' 16.11.20:
         If UBound(Parts) > 1 Then
            MacroName = Split(Parts(1), "(")(0)
            Select Case MacroName
                Case "BlueLight1", _
                     "BlueLight2", _
                     "Leuchtfeuer": ' Check if C_All is used because this is not possible for these macros
                                    If InStr(Parts(1), " C_ALL,") > 0 Then
                                       MsgBox Replace(Get_Language_Str("Fehler: Das Makro '#1#' kann nur mit einer oder zwei LEDs benutzt werden."), "#1#", MacroName), vbCritical, Get_Language_Str("Fehler: Makro kann nicht mit 3 LEDs benutzt werden")
                                       ActiveCell = Replace(Parts(1), " C_ALL, ", " C12, ")
                                       Cells(DstRow, LEDs____Col) = "C1-2"
                                    End If
            End Select
         End If
         
         ' Changed by Misha 18-4-2020                                       ' 14.06.20: Added from Mishas version
         Parts = Split(Res, ",")
         If left(MacroName, Len("Multiplexer")) = "Multiplexer" Then
             Cells(DstRow, LocInCh_Col) = Count_Ones(val(Parts(5))) + 1 ' Get the number of selected Patterns in the Multiplexer function. It gives the number of LocInCh variables.
                                                                        ' 10.02.21: 20210206 Misha, Added + 1 because there is an zero pattern added.
             
             Cells(ActiveCell.Row, DCC_or_CAN_Add_Col).Value = Userform_Res_Address ' 10.02.21: 20210208 Added by Misha, to add DCC Address to created Multiplexer.
         
         Else
             Cells(DstRow, LocInCh_Col) = val(.Cells(SelRow, SM_LocInCCOL)) ' use the Val function to get 0 if the cell is empty in sheet Lib_Macros
         End If
         ' End Changed by Misha 18-4-2020
         
         
         Cells(DstRow, Enable_Col) = ChrW(Hook_CHAR) ' Enable the Line
         Update_TestButtons DstRow
         Update_Start_LedNr
         SelectMacros_Sub = True
      End If
    End With
  End If
  
  
  If Res <> "" Then  ' Move the cursor to the next cell                     ' 23.04.20:
        Dim NextRow As Long
        NextRow = ActiveCell.Row + 1
        While Cells(NextRow, 1).EntireRow.Hidden                            ' 20.04.20:
              NextRow = NextRow + 1
        Wend
        
        Move_Cursor_to_visible_Macro_Cell NextRow
  Else: Move_Cursor_to_visible_Macro_Cell ActiveCell.Row
  End If
  
End Function

'---------------------------------------------------------
Private Sub Move_Cursor_to_visible_Macro_Cell(Row As Long)
'---------------------------------------------------------
    Dim OldEvents As Boolean
    OldEvents = Application.EnableEvents
    Application.EnableEvents = False
    
    If Cells(Row, Config__Col).EntireColumn.Hidden Then ' Check if the column is hidden       ' 24.10.21:
          If Cells(Row, LanName_Col).EntireColumn.Hidden Then
                Cells(Row, MacIcon_Col).Select
          Else: Cells(Row, LanName_Col).Select
          End If
    Else: Cells(Row, Config__Col).Select
    End If
    
    Application.EnableEvents = OldEvents
End Sub
'--------------------------------------------------------------------------
Public Function Find_Macro_in_Lib_Macros_Sheet(ByVal Str As String) As Long
'---------------------------------------------------------------------------
  Dim r As Range, c As Range
  With Sheets(LIBMACROS_SH)
  
     Str = Replace(Str, "HouseT(", "House(")
     If InStr(Str, "(") > 0 Then
        Str = Split(Str, "(")(0) & "(" ' Cut of the text after the "("
     End If
  
     Set r = .Range(.Cells(SM_DIALOGDATA_ROW1, SM_Name__COL), .Cells(LastUsedRowIn(ThisWorkbook.Sheets(LIBMACROS_SH)), SM_Name__COL))
     For Each c In r
       ' Find the line
       If .Cells(c.Row, SM_FindN_COL) <> "" Then
          #If True Then  ' 17.04.20: Problems detecting the "Sound_PlayRandom()" macro. The old line found the "Random()" macro.
                         '           Why did the old function use the InStr function ?
            If Str = .Cells(c.Row, SM_FindN_COL) Then
          #Else
            If InStr(Str, .Cells(c.Row, SM_FindN_COL)) <> 0 Then
          #End If
               Find_Macro_in_Lib_Macros_Sheet = c.Row
               Exit Function
            End If
       End If
     Next c
  End With
End Function

'UT----------------------------------------------
Private Sub Test_Find_Macro_in_Lib_Macros_Sheet()
'UT----------------------------------------------
  Debug.Print Find_Macro_in_Lib_Macros_Sheet("Logic(")
  Debug.Print Find_Macro_in_Lib_Macros_Sheet("Const(")
End Sub


#If False Then ' This functions are no longer needed because we don't use links
    '------------------------------------------------------------------------
    Private Sub Change_Links_to_Absolute_In_Col(Sh As Worksheet, Col As Long)   ' 07.10.21:
    '------------------------------------------------------------------------
        Dim c As Variant
        With Sh
            For Each c In .Range(.Cells(HeadRow + 1, Col), .Cells(LastUsedRowIn(Sh), Col))
                If c.Formula <> "" & left(c.Formula, 1) = "=" Then
                     c.Formula = Application.ConvertFormula(c.Formula, xlA1, xlA1, xlAbsolute)
                End If
            Next c
        End With
    End Sub
    
    
    '-------------------------------------
    Private Sub Change_Links_to_Absolute()                                      ' 07.10.21:
    '-------------------------------------
    ' This function must be called before sorting the lines. Otherwise the links get corrupted
        Dim Sh As Worksheet
        Set Sh = ActiveWorkbook.Worksheets(LIBMACROS_SH)
        Dim Col As Long, c As Variant
        With Sh
           Change_Links_to_Absolute_In_Col Sh, SM_Pic_N_COL
           For Col = SM_Group_COL To LastUsedColumnIn(Sh) Step DeltaCol_Lib_Macro_Lang
              Change_Links_to_Absolute_In_Col Sh, Col
           Next Col
        End With
    End Sub
#End If

'----------------------------------------------------------
Private Sub Sort_by_Column(Col As Long, SortFlag As String)                 ' 07.10.21:
'----------------------------------------------------------
    If ThisWorkbook.Worksheets(LIBMACROS_SH).Range("SortByTreeView").Value = SortFlag Then
       Exit Sub
    End If
    Dim OldUpdating As Boolean
    OldUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ' Change_Links_to_Absolute                                              ' 20.10.21: Links are not used anymore because they also create problems if absolute links are used when the lines are sorted ;-(
    Dim Sh As Worksheet
    Set Sh = ActiveWorkbook.Worksheets(LIBMACROS_SH)
    With Sh
      .Sort.SortFields.Clear ' 20.10.21: Replaced ".Add2" by ".Add" in the following line. This solves Ulrichs problem. (See: https://www.mrexcel.com/board/threads/vba-difference-in-sort-with-add2-and-add-sortfields-add2-vs-sortfields-add.1072594/)
      .Sort.SortFields.Add Key:=.Range(.Cells(HeadRow, Col), .Cells(LastUsedRowIn(Sh), Col)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
      .Sort.SetRange .Range(.Cells(HeadRow + 1, 1), .Cells(LastUsedRowIn(Sh), LastUsedColumnIn(Sh)))
      With .Sort
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
      End With
    End With
    ThisWorkbook.Worksheets(LIBMACROS_SH).Range("SortByTreeView").Value = SortFlag
    Application.ScreenUpdating = OldUpdating
End Sub

'-------------------------------------
Public Sub Sort_for_List_based_Makro()                                      ' 07.10.21:
'-------------------------------------
  Sort_by_Column SM_ListS_COL, "L"
End Sub

'-----------------------------------------
Public Sub Sort_for_TreeView_based_Makro()                                  ' 07.10.21:
'-----------------------------------------
  Sort_by_Column SM_TreeS_COL, "T"
End Sub


'-------------------------------------------------------------------------------------------------
Public Function Get_Language_Text(Row As Long, FirstCol As Long, ActLanguage As Integer) As String
'-------------------------------------------------------------------------------------------------
' Get the language specific text
' If the requested text is not available use the englich or german text
  Dim Sh As Worksheet, Txt As String
  Set Sh = ThisWorkbook.Sheets(LIBMACROS_SH)
  With Sh
    Txt = .Cells(Row, FirstCol + ActLanguage * DeltaCol_Lib_Macro_Lang).Value
    If Txt = "" And ActLanguage > 1 Then Txt = .Cells(Row, FirstCol + 1 * DeltaCol_Lib_Macro_Lang).Value ' Use the english name
    If Txt = "" And ActLanguage > 0 Then Txt = .Cells(Row, FirstCol + 0 * DeltaCol_Lib_Macro_Lang).Value ' Use the german name
  End With
  Get_Language_Text = Txt
End Function




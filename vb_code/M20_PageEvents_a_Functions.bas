Attribute VB_Name = "M20_PageEvents_a_Functions"
' Update several cells in the DCC and Selectrix sheet by events
' - Set/delete the Enable hook
' - Toggle the Input Typ
' - Calculate the Start LedNr
' - Format new lines
' -
' 27.03.20: - add automatic creation of DCC/Selectrix test buttons
' 29.03.20: - automatic resize of buttons column
' 31.03.20: - add toggle and improve ButtonColor handling
' 02.04.20: - limit size of test buttons
'           - if multiple test buttons use same address toggle all related buttons
' 26.03.21: - speed up ResetTestButtons function

Option Explicit

Private Global_Rect_List() As String                                        ' 25.03.21:
    
Private PriorCell As Range

'---------------------------------------------------------------------------------
Private Function Get_Parameter_from_Leds_Line(ByVal line As String, ParNr As Long)
'---------------------------------------------------------------------------------
  If Left(line, 1) = "^" Then
     line = Trim(Mid(line, 2, 200))
  End If
  
  If Left(line, 1) <> "C" Then
     MsgBox Get_Language_Str("Fehler: LEDs Eintrag muss mit 'C' beginnen"), vbCritical, Get_Language_Str("Fehler in LEDs Eintrag")
     EndProg
  End If
  
  Dim Parts As Variant, Error As Boolean
  Parts = Split(Mid(line, 2, 200), "-")
  If UBound(Parts) <> 1 Then
     Error = 1
  ElseIf Not IsNumeric(Parts(0)) Or Not IsNumeric(Parts(1)) Then
     Error = 1
  End If
  If Error Then
     MsgBox Get_Language_Str("Fehler: LEDs Eintrag muss zwei mit '-' getrennte Kanäle enthalten"), vbCritical, Get_Language_Str("Fehler in LEDs Eintrag")
     EndProg
  End If
  
  Get_Parameter_from_Leds_Line = val(Parts(ParNr))
End Function

'---------------------------------------------------------------------------------------------
Private Sub Update_LastUsedChannel_in_Row(ByVal line As String, ByRef LastusedChannel As Long)
'---------------------------------------------------------------------------------------------
  line = Trim(line)
  Dim val As Long
  val = Get_Parameter_from_Leds_Line(line, 1)
  If val > LastusedChannel Then LastusedChannel = val
End Sub

'--------------------------------------------------------------------------------
Public Sub Update_LastUsedChannel(LEDs As String, ByRef LastusedChannel As Long)
'--------------------------------------------------------------------------------
  If InStr(LEDs, vbLf) Then
        Dim line As Variant
        For Each line In Split(LEDs, vbLf)
           Update_LastUsedChannel_in_Row line, LastusedChannel
        Next line
  Else: Update_LastUsedChannel_in_Row LEDs, LastusedChannel
  End If
End Sub

'---------------------------------------------------
Public Function Get_FirstUsedChannel(LEDs As String)
'---------------------------------------------------
  If InStr(LEDs, vbLf) Then
        Dim line As Variant
        Get_FirstUsedChannel = 99
        For Each line In Split(LEDs, vbLf)
           Dim val As Long
           val = Get_Parameter_from_Leds_Line(line, 0)
           If val < Get_FirstUsedChannel Then Get_FirstUsedChannel = val
        Next line
  Else: Get_FirstUsedChannel = Get_Parameter_from_Leds_Line(LEDs, 0)
  End If
End Function

'-------------------------------------------------
Private Function UsedModules(LEDs As Long) As Long
'-------------------------------------------------
  Dim d As Double
  d = LEDs / 3
  If d = Round(d) Then
        UsedModules = d
  Else: UsedModules = Round(d + 0.5, 0)
  End If
End Function

'-------------------------------------------------------
Public Function Check_IsSingleChannelCmd(LEDs As String)
'-------------------------------------------------------
  Check_IsSingleChannelCmd = InStr(LEDs, "C") > 0
End Function


'---------------------------------------------------
Public Function Row_is_Achtive(r As Long) As Boolean
'---------------------------------------------------
  Row_is_Achtive = Rows(r).EntireRow.Hidden = False And Cells(r, Enable_Col) <> "" And Trim(Cells(r, LEDs____Col)) <> ""
End Function


'--------------------------------------------------------------------
Public Function Row_Contains_Address_or_VarName(r As Long) As Boolean       ' 20.04.20:
'--------------------------------------------------------------------
  If Rows(r).EntireRow.Hidden = False And Cells(r, Enable_Col) <> "" Then
     Dim Addr_Col As Long
     Addr_Col = Get_Address_Col()
     
     If Trim(Cells(r, Addr_Col)) = "" Then Exit Function    ' Empty Address
     If IsNumeric(Trim(Cells(r, Addr_Col))) Then ' For Variable names this check is not needed
        If Page_ID = "Selextrix" Then
           If Trim(Cells(r, SX_Bitposi_Col)) = "" Then Exit Function
        End If
        If Trim(Cells(r, Inp_Typ_Col)) = "" Then Exit Function
     End If
  End If
  Row_Contains_Address_or_VarName = True
End Function

' Update_Start_LedNr() Berechnet die StartLedNr anhand der vorangegangenen
' Start Nummer und der LEDs der vorangegangenen Zeile.
' Die StartLedNr wird zu bei der Bearbeitung der Zeilen zu beginn
' in das Excel Sheet geschrieben und anschließend die interne Variable für die nächste Zeile erhöht.

' Zeilen welche einzelne LEDs ansprechen (C1, C2, ...) müssen anders behandelt werden.
' Hier wird StartLedNr dann erhöht wenn die neue Zeile:
' - einen LED Kanal anspricht der kleiner ist als der letzte LED Kanal
' - alle Kanäle anspricht (C_ALL, House(), ...)
' Die Überprüfung muss also vor dem schreiben der StartLedNr gemacht werden.

#If True Then ' 26.04.20: New function with support for the additional "LEDs Taster" column
'------------------------------
Public Sub Update_Start_LedNr()
'------------------------------
' Update the Start LedNr in all used rows
  Dim OldEvents As Boolean: OldEvents = Application.EnableEvents
  Dim display_Type As Long
  display_Type = LEDNr_Display_Type
   
  Application.EnableEvents = False  ' Prevent recursive calls ic cells are changed                ' 26.04.20:
  
  
  Make_sure_that_Col_Variables_match

  Dim Row As Variant, LEDNr(LED_CHANNELS) As Long, LastusedChannel(LED_CHANNELS) As Long, MaxLEDNr(LED_CHANNELS) As Long
  Dim r As Long, Max_LEDs_Channel
  
  ' 03.04.21 Juergen - try to find out if only the main LED Channel is in use
  Max_LEDs_Channel = 0
  For Each Row In ActiveSheet.UsedRange.Rows  ' ToDo: Hier könnte man einen Eigenen Range definieren der erst ab der FirstDat_Row beginnt. Dann könnte die abfrage "If r >= FirstDat_Row Then" entfallen
     r = Row.Row
     If r >= FirstDat_Row Then
        If Row_is_Achtive(r) Then
            If val(Cells(r, LED_Cha_Col)) > Max_LEDs_Channel Then
                Max_LEDs_Channel = val(Cells(r, LED_Cha_Col))
            End If
        End If
    End If
  Next
  
  For Each Row In ActiveSheet.UsedRange.Rows  ' ToDo: Hier könnte man einen Eigenen Range definieren der erst ab der FirstDat_Row beginnt. Dann könnte die abfrage "If r >= FirstDat_Row Then" entfallen
     r = Row.Row
     If r >= FirstDat_Row Then
        If Row_is_Achtive(r) Then
           Dim LEDs As String, IsSingleChannelCmd As Boolean, LEDs_Channel As Long
           LEDs_Channel = val(Cells(r, LED_Cha_Col))
           LEDs = Trim(Cells(r, LEDs____Col))
           IsSingleChannelCmd = Check_IsSingleChannelCmd(LEDs)
           
           ' Check if the privious lines adress single channels
           If LastusedChannel(LEDs_Channel) > 0 Then
              If IsSingleChannelCmd Then
                 If Left(LEDs, 1) = "^" Then LastusedChannel(LEDs_Channel) = 0
                 If Get_FirstUsedChannel(LEDs) <= LastusedChannel(LEDs_Channel) Then
                    LEDNr(LEDs_Channel) = LEDNr(LEDs_Channel) + UsedModules(LastusedChannel(LEDs_Channel))
                    LastusedChannel(LEDs_Channel) = 0
                 End If
              Else ' Last lines have been single channel lines, but the actual line adresses full RGB LEDs
                 LEDNr(LEDs_Channel) = LEDNr(LEDs_Channel) + UsedModules(LastusedChannel(LEDs_Channel))
                 LastusedChannel(LEDs_Channel) = 0
              End If
           End If
           
           Clear_LED_Nr_Columns r, LED_Nr__Col
           
           With Cells(r, LED_Nr__Col)
                Dim NewValue
                If LEDs = SerialChannelPrefix Then      ' 07.10.21: Jürgen Serial Channel
                    NewValue = "'" & SerialChannelPrefix & LEDs_Channel
                Else
                    If Max_LEDs_Channel > 0 And display_Type = 1 Then ' 03.04.21: Depending on this setting the LedNumber is displayed differently
                        NewValue = "'" & LEDs_Channel & "-" & LEDNr(LEDs_Channel)
                    Else
                        NewValue = LEDNr(LEDs_Channel)
                    End If
                End If
                If .Value = "" Or .Value <> NewValue Then .Value = NewValue
                If IsSingleChannelCmd Then
                     Update_LastUsedChannel LEDs, LastusedChannel(LEDs_Channel)
                     If LastusedChannel(LEDs_Channel) > 3 Then
                        ' Wenn mehrere einzelne LEDs verwendet werden welche nicht in ein WS2811 Modul passen  ' 06.06.20:
                        ' Beispiel: Drei aufeinanderfolgende "Herz_BiRelais()" Zeilen
                        '           Das zweite Herz_BiRelais() Belegt den blauen Kanal und den Roten des nächsten WS2811
                        '           Darum muss LastusedChannel auf 1 gesetzt und LEDNr um 1 erhöht werden
                        Dim IncLEDNr As Long
                        IncLEDNr = LastusedChannel(LEDs_Channel) \ 3
                        LastusedChannel(LEDs_Channel) = LastusedChannel(LEDs_Channel) Mod 3
                        LEDNr(LEDs_Channel) = LEDNr(LEDs_Channel) + IncLEDNr
                     End If
                Else
                     LEDNr(LEDs_Channel) = LEDNr(LEDs_Channel) + CellLinesSum(LEDs) ' If the cell consists of several lines the sum is calculated
                End If
           End With
        Else ' Row is not activ
           Clear_LED_Nr_Columns r
        End If
     End If
     If LEDNr(LEDs_Channel) > MaxLEDNr(LEDs_Channel) Then MaxLEDNr(LEDs_Channel) = LEDNr(LEDs_Channel)   ' 19.01.21:
  Next Row
  
  ' Write the Last LED number to row 1 (SH_VARS_ROW)
  Dim Nr As Long
  For Nr = 0 To LED_CHANNELS - 1
     Dim Col As Long: Col = Get_LED_Nr_Column(Nr)
     Cells(SH_VARS_ROW, Col) = MaxLEDNr(Nr) + UsedModules(LastusedChannel(Nr))  ' Needed for the NUM_LEDS definition (Not visible because it's printed in white)  19.01.21: Using "MaxLEDNr(Nr)" instead of "LEDNr(Nr)"
  Next Nr
  
  Application.EnableEvents = OldEvents                                      ' 26.04.20:
  Clear_Formula_Errors                                                      ' 03.04.21: remove the green triangles showing cell error
End Sub

'-----------------------------------------------------
'get the led number out of the content of the ledNr cell
'depending on setting the channel may be selected by an own column or as a prefix of the LED Number
Public Function Get_LED_Nr(DefaultLedNr As Long, Row As Long, LED_Channel As Long) As Long                    ' 04.03.21:  Juergen
'-----------------------------------------------------
  Get_LED_Nr = DefaultLedNr
  Dim LEDNr As String                       ' 08.10.21: Juergen
  LEDNr = Cells(Row, LED_Nr__Col)
  If LEDNr <> "" Then
     Dim Pos As Long
     Pos = InStr(LEDNr, "-")
     If (Pos < 1) Then
        If Left(LEDNr, 1) = SerialChannelPrefix Then            ' 08.10.21: Juergen
            Get_LED_Nr = val(Mid(LEDNr, 2))
            Exit Function
        End If
        Get_LED_Nr = LEDNr + Start_LED_Channel(LED_Channel)
     Else
        If Left(LEDNr, Pos - 1) = LED_Channel Then
            Get_LED_Nr = val(Mid(LEDNr, Pos + 1)) + Start_LED_Channel(LED_Channel)
        Else
            ' todo
        End If
     End If
  End If
End Function

'-----------------------------------------------------
Public Function Get_LED_Nr_Column(LED_Channel As Long)                      ' 26.04.20:
'-----------------------------------------------------
     If LED_Channel = 0 Then
           Get_LED_Nr_Column = LED_Nr__Col
     Else: Get_LED_Nr_Column = LED_TastCol + LED_Channel - 1
     End If
End Function

'---------------------------------------------------------------------
Public Sub Clear_LED_Nr_Columns(r As Long, Optional ExceptCol As Long)      ' 26.04.20:
'---------------------------------------------------------------------
  Dim Nr As Long
  For Nr = 0 To LED_CHANNELS - 1
     Dim Col As Long: Col = Get_LED_Nr_Column(Nr)
     If Col <> ExceptCol Then
        With Cells(r, Col)
          If Col = LED_Nr__Col Then
             If .Formula <> "=""""" Then .Formula = "="""""  ' 30.04.20: Using the formula ="" to prevent overlapping texts from the left cell
          Else
             If .Value <> "" Then .Value = ""
          End If
        End With
     End If
  Next Nr
End Sub

#Else ' 26.04.20: Old Function before the additional LEDs Taster column was added
'------------------------------
Public Sub Update_Start_LedNr()
'------------------------------
' Update the Start LedNr in all used rows

  Make_sure_that_Col_Variables_match

  Dim Row As Variant, LEDNr As Long, LastusedChannel As Long
  For Each Row In ActiveSheet.UsedRange.Rows
     Dim r As Long
     r = Row.Row
     With Cells(r, LED_Nr__Col)
        If r >= FirstDat_Row Then
            If Row_is_Achtive(r) Then
               Dim LEDs As String, IsSingleChannelCmd As Boolean
               LEDs = Trim(Cells(r, LEDs____Col))
               IsSingleChannelCmd = Check_IsSingleChannelCmd(LEDs)
               
               ' Check if the privios lines adress single channels
               If LastusedChannel > 0 Then
                  If IsSingleChannelCmd Then
                     If Left(LEDs, 1) = "^" Then LastusedChannel = 0
                     If Get_FirstUsedChannel(LEDs) <= LastusedChannel Then
                        LEDNr = LEDNr + UsedModules(LastusedChannel)
                        LastusedChannel = 0
                     End If
                  Else ' Last lines have been single channel lines, but the actual line adresses full RGB LEDs
                     LEDNr = LEDNr + UsedModules(LastusedChannel)
                     LastusedChannel = 0
                  End If
               End If
               
               If .Value = "" Or .Value <> LEDNr Then .Value = LEDNr
               If IsSingleChannelCmd Then
                    Update_LastUsedChannel LEDs, LastusedChannel
               Else
                    LEDNr = LEDNr + CellLinesSum(LEDs) ' If the cell consists of several lines the sum is calculated
               End If
            Else
               If .Value <> "" Then .Value = ""
            End If
        End If
     End With
  Next Row
  Cells(1, LED_Nr__Col) = LEDNr + UsedModules(LastusedChannel) ' Needed for the NUM_LEDS definition (Not visible because it's printed in white)
End Sub

#End If

'-------------------------------------------------------------
Private Function FirstNotFormatedRow(StartRow As Long) As Long
'-------------------------------------------------------------
' Find the first row where the Enable_Col doesn't use the font "Wingdings"
' The search starts with the giveb row and searches upwards.
' Return 0 if the StartRow is already formated
  Dim Row As Long
  If Cells(StartRow, Enable_Col).Font.Name = "Wingdings" Then
     Exit Function
  End If
  
  Row = StartRow
  While Cells(Row, Enable_Col).Font.Name <> "Wingdings" And Row > FirstDat_Row
    Row = Row - 1
  Wend

  FirstNotFormatedRow = Row
End Function


'--------------------------------------------------------------------------------------------------------------------
Public Sub Format_Cells_to_Row(Row As Long, Optional UpdateAutofilter As Boolean = True, Optional AllRows As Boolean)
'--------------------------------------------------------------------------------------------------------------------
' Formating the unformated rows
' - Wingdings in the Enable column and centeres in both directions
' - For all other columns the horizontal alignement is copied fron the Head line
'   and the vertical alignement is XlTop
' - Border lines

  Dim FirstNFormRow As Long
  If AllRows Then
        FirstNFormRow = FirstDat_Row
  Else: FirstNFormRow = FirstNotFormatedRow(Row)
  End If
  
  If FirstNFormRow > 0 Then
    'Debug.Print "Formating rows " & FirstNFormRow & " to " & Row ' Debug
    With Range(Cells(FirstNFormRow, Enable_Col), Cells(Row, Enable_Col))
      .Font.Name = "Wingdings" ' Enable column
      .VerticalAlignment = xlCenter
      .HorizontalAlignment = xlCenter
    End With
    Dim c As Long
    For c = Enable_Col + 1 To LastColumnDatSheet()                          ' 27.11.21: Old: LastUsedColumn
        With Range(Cells(FirstNFormRow, c), Cells(Row, c))
          '.Select ' Debug
          .HorizontalAlignment = Cells(Header_Row, c).HorizontalAlignment
          .VerticalAlignment = xlTop
          
          'If C = LEDs____Col Then .NumberFormat = "@" ' Text format to be able to enter "1-2"
          
          If c = DCC_or_CAN_Add_Col Or _
             c = SX_Channel_Col Then .NumberFormat = "@" ' Text format to be able to enter "1-2"
          If c = Config__Col Then .Font.Name = "Consolas"
          
          If c = DCC_or_CAN_Add_Col Then .WrapText = True
           
          ' Gray background
          Select Case c
             Case Config__Col, _
                  LED_Nr__Col To LED_Nr__Col + INTERNAL_COL_CNT - 1:
                               With .Interior
                                 .ThemeColor = xlThemeColorDark1
                                 .TintAndShade = -4.99893185216834E-02
                               End With
          End Select

        End With
    Next
    All_Borderlines Range(Cells(FirstNFormRow, Enable_Col), Cells(Row, LastColumnDatSheet()))  ' 27.11.21: Old: LastUsedColumn
    
    If UpdateAutofilter Then Expand_Filter_to_all_used_Rows
    'If UpdateAutofilter Then AutofilterAllColumns Row ' Deletes the filter settings ;-(
  End If
End Sub

'---------------------------
Public Sub Format_All_Rows()
'---------------------------
  Make_sure_that_Col_Variables_match
  Format_Cells_to_Row LastUsedRow(), True, True
End Sub

'---------------------------
Public Sub Update_Sum_Func()
'---------------------------
' The Sum function in the filter column is used to detect changes in the autofilter
'  to update the "Start LedNr"
  Cells(SH_VARS_ROW, Filter__Col).FormulaR1C1 = "=SUM(R[1]C:R[" & LastUsedRow() & "]C)"
End Sub

'------------------------------------------------------
Private Function Col_is_in_Range(c As Long, r As Range)
'------------------------------------------------------
  Col_is_in_Range = c >= r.Column And c <= r.Column + r.Columns.Count - 1
End Function

'-----------------------------------------------------
Private Function Range_is_Empty(r As Range) As Boolean
'-----------------------------------------------------
  Dim c As Variant
  For Each c In r
     If c <> "" Then Exit Function
  Next c
  Range_is_Empty = True
End Function


'------------------------------------------------------
Public Sub Select_Typ_by_Dialog(Target As Excel.Range)
'------------------------------------------------------
  Dim OldEvents As Boolean
  If Not Address_starts_with_a_Number(Target.Row) Then Exit Sub                            ' 03.02.20:
  
  OldEvents = Application.EnableEvents
  Application.EnableEvents = False
  Target.Select
  If Page_ID = "Selectrix" Then
       UserForm_Select_Typ_SX.setFocus Target
       UserForm_Select_Typ_SX.Show
  Else ' "DCC", "CAN"
       UserForm_Select_Typ_DCC.setFocus Target
       UserForm_Select_Typ_DCC.Show
  End If
  Application.EnableEvents = True ' Enable to update the Address column
  If Userform_Res <> "" Then Target = Userform_Res
  Application.EnableEvents = OldEvents
End Sub

'----------------------------------------------------------------------------------------
Public Sub Complete_Typ(Target As Excel.Range, Optional DialogIfEmpty As Boolean = False)
'----------------------------------------------------------------------------------------
  Set_Tast_Txt_Var  ' Set the global variables Red_T, Green_T, ...          06.03.20:
  
  If Len(Target) = 0 Then
     If DialogIfEmpty Then Select_Typ_by_Dialog Target
     Exit Sub
  End If
  Dim Txt As String
  If Page_ID = "Selectrix" Then
     If UCase(Left(Target, 1)) = UCase(Left(Tast_T, 1)) Then Txt = Tast_T ' ToDo: Language
  Else ' "DCC", "CAN"
     If UCase(Left(Target, 1)) = UCase(Left(Red_T, 1)) Then Txt = Red_T
     If UCase(Left(Target, 1)) = UCase(Left(Green_T, 1)) Then Txt = Green_T
  End If
  If UCase(Left(Target, 1)) = UCase(Left(OnOff_T, 1)) Then Txt = OnOff_T
  
  If Txt <> "" Then
        If Target <> Txt Then Target = Txt                                  ' 23.06.20: Speed up (took 45 seconds at a larger configuration with 130 LEDs and a lot od DCC (Harald))
  Else: Select_Typ_by_Dialog Target
  End If
End Sub


'----------------------------------------------------
Private Function AutofilterAllColumns(EndRow As Long)
'----------------------------------------------------
' Activate the autofilter for all columns
  With Range(Cells(Header_Row, Enable_Col), Cells(EndRow, LastColumnDatSheet()))  ' 27.11.21: Old: LastUsedColumn
    If ActiveSheet.AutoFilterMode Then ' already there
        With .Cells
            .AutoFilter ' remove it
            .AutoFilter ' put it back
        End With
    Else
        .Cells.AutoFilter ' first time
    End If
  End With
End Function

'-------------------------------
Private Sub Correct_Autofilter()
'-------------------------------
' Call this manualy in case the filter is corrupted
  AutofilterAllColumns LastUsedRow
End Sub

'-------------------------------
Public Sub Clear_Formula_Errors()               ' 08.10.21: Juergen
'-------------------------------
  Dim rngCell As Range, i As Long
  For Each rngCell In ActiveSheet.UsedRange
      For i = 1 To 7
          rngCell.Errors.Item(i).Ignore = True
      Next i
  Next
End Sub


'-----------------------------
Private Sub Move_Active_Cell()
'-----------------------------
  On Error GoTo JustActivete                                                ' 27.10.20:
  If Not PriorCell Is Nothing Then
     If PriorCell.Row = ActiveCell.Row Then
        Dim Delta As Long
        Delta = ActiveCell.Column - PriorCell.Column
        If Abs(Delta) = 1 Then
           Range(ActiveCell.Address).Offset(0, Delta).Activate
           Exit Sub
        End If
     End If
  End If
JustActivete:
  Range(ActiveCell.Address).Offset(0, 1).Activate
End Sub

'----------------------------------------------
Private Sub Proc_Typ_Col(Target As Excel.Range)
'----------------------------------------------
' Open the Typ Dialog if the "DCC Adresse" or "Bitposition" Cell is filled
' but the "Typ" cell is empty
  Dim RefCol As Long
  Select Case Page_ID
    Case "DCC", "CAN": RefCol = DCC_or_CAN_Add_Col
    Case "Selectrix":  RefCol = SX_Bitposi_Col
    Case Else:         MsgBox "Internal error: Unknown Page_ID in 'Proc_Typ_Col'", vbCritical
                       EndProg
  End Select

  If Cells(Target.Row, RefCol) <> "" And Target = "" Then
     Select_Typ_by_Dialog Target
     Move_Active_Cell
  End If
End Sub

'-----------------------------------------------------
Private Sub Hide_Unhide_Selected_Rows(Hide As Boolean)
'-----------------------------------------------------
' Is called if rows are hidden or unhiden by event´
   Dim OldUppdating As Boolean, OldEvents As Boolean
   OldUppdating = Application.ScreenUpdating
   Application.ScreenUpdating = False
   OldEvents = Application.EnableEvents
   Application.EnableEvents = False
    
   Selection.EntireRow.Hidden = Hide
    
   Update_Start_LedNr ' Update the LedNr
    
   Application.EnableEvents = OldEvents
   Application.ScreenUpdating = OldUppdating
End Sub

'----------------------------
Public Sub myHideRows_Event()
'----------------------------
' Is called on event if rows are hidden
' Must be enabled in Workbook_Open() which is located in "DieseArbeismappe"
  Hide_Unhide_Selected_Rows True
End Sub

'------------------------------
Public Sub myUnhideRows_Event()
'------------------------------
' Is called on event if rows are shown
' Must be enabled in Workbook_Open() which is located in "DieseArbeismappe"
  Hide_Unhide_Selected_Rows False
End Sub

'-------------------------
Public Sub Option_Dialog()
'-------------------------
  Check_Version                                                             ' 21.11.21: Juergen
  UserForm_Options.Show
End Sub


'----------------------
Public Sub ClearSheet()
'----------------------
  If MsgBoxMov(Get_Language_Str("Wollen Sie alle Einträge dieser Seite löschen?"), vbQuestion + vbYesNo, Get_Language_Str("Seite Löschen?")) = vbYes Then
     Dim OldUppdating As Boolean, OldEvents As Boolean
     OldUppdating = Application.ScreenUpdating
     Application.ScreenUpdating = False
     OldEvents = Application.EnableEvents
     Application.EnableEvents = False
     
     Make_sure_that_Col_Variables_match
     Rows(FirstDat_Row & ":" & LastUsedRow).ClearContents ' Hidden Rows are also cleared
     Rows(FirstDat_Row & ":" & LastUsedRow).Hidden = False
     Rows("33:" & LastUsedRow()).Delete Shift:=xlUp
     
     Cells(FirstDat_Row, 1).Activate                                        ' 22.10.21: First column, prior the sheet was shifted to the description col. Old: Descrip_Col
     Application.Goto ActiveCell, True
     ResetTestButtons False                                                 ' 21.03.21 Juergen: clear buttons
     Del_Icons_in_IconCol
     
     Application.EnableEvents = OldEvents
     Application.ScreenUpdating = OldUppdating
  End If
End Sub

'---------------------
Public Sub Show_Help()
'---------------------
  Application.EnableEvents = True ' In case of a previous crash this button enables the events again
  
  #If 1 Then
     Shell "Explorer https://wiki.mobaledlib.de/anleitungen/spezial/programmgenerator"
  #Else
  Dim Name As String
  Name = ThisWorkbook.Path & "\Prog_Generator_MobaLedLib.pdf"
  If Dir(Name) = "" Then
        MsgBox Get_Language_Str("Fehler: Die Hilfe Datei fehlt:") & vbCr & _
               "  '" & Name & "'", vbCritical, Get_Language_Str("Hilfe Datei wurde nicht gefunden")
  Else: Shell "Explorer " & Name
  End If
  #End If
End Sub


'----------------------------------------------------------------------------------------------
Public Sub Proc_DoubleCkick(ByVal Sh As Object, ByVal Target As Range, ByRef Cancel As Boolean)
'----------------------------------------------------------------------------------------------
  Make_sure_that_Col_Variables_match Sh
  If Is_Data_Sheet(ActiveSheet) Then
     If Target.Row >= FirstDat_Row Then
        Select Case Target.Column
          Case Config__Col, MacIcon_Col, LanName_Col
                            Cancel = True: SelectMacros             ' cancel = True to disable the standard function => Don't go into cell edit mode
          Case Inp_Typ_Col: Cancel = True: Select_Typ_by_Dialog Target
        End Select
     End If
  End If
End Sub

'------------------------------------
Public Sub Arduino_Button_StartPage()
'------------------------------------
   UserForm_Protokoll_Auswahl.Show
End Sub


' Typ  InCnt Start - End
' ~~~~ ~~~~~ ~~~~~~~~~
' Rot   1    1 - 1
' Grün  1    1 - 1
' Rot   2    1 - 1
' Grün  2    1 - 2
' Rot   3    1 - 2
' Grün  3    1 - 2
' Rot   4    1 - 2
' Grün  4    1 - 3

' Selextrix
' BitPos InCnt  End  Int(End/8) EndMod8 IncAddr
' 5       3     7    0          7       0
' 6       3     8    1          0       0
' 7       3     9    1          1       1
' 7       4    10    1          2       1
' 7       8    15

'----------------------------------------------------
Public Sub Complete_Addr_Column_with_InCnt(r As Long)
'----------------------------------------------------
' Is called by event if the DCC address or the Selectrix chanel is changed
  Set_Tast_Txt_Var  ' Set the global variables Red_T, Green_T, ...          06.03.20:
  
  Dim Target As Range
  Set Target = Cells(r, DCC_or_CAN_Add_Col + SX_Channel_Col) ' Add both columns because the not used is 0
  'Debug.Print "Complete_Addr_Column_with_InCnt"
  Dim InCnt As Long
  InCnt = val(Cells(r, InCnt___Col))
  If InCnt > 0 Then
     Dim Addr As String
     Addr = Get_First_Number_of_Range(r, Target.Column)
     If Addr <> "" And val(Addr) >= 0 Then
        Dim Inp_Typ As String, SBit_Str As String, EBit_Str As String
        SBit_Str = ""
        EBit_Str = ""
        Inp_Typ = Cells(r, Inp_Typ_Col)
        If Inp_Typ <> "" Then
           Dim i As Long, EndAddr As Long, EBit_Pos As Long
           If Page_ID = "Selectrix" Then
                Dim BitPos As Long
                BitPos = val(Cells(r, SX_Bitposi_Col))
                If BitPos > 0 And BitPos <= 8 Then
                   EBit_Pos = BitPos + InCnt - 1
                   EndAddr = Addr + Int(EBit_Pos / 8)
                   
                   If EBit_Pos Mod 8 = 0 Then
                         EBit_Str = ".8": EndAddr = EndAddr - 1
                   Else: EBit_Str = "." & EBit_Pos Mod 8
                   End If
                   SBit_Str = "." & BitPos
                End If
           Else ' DCC
                If Inp_Typ = Red_T Or Inp_Typ = Green_T Then
                    EndAddr = Addr
                    For i = 1 To InCnt                         ' Unelegant
                        Inp_Typ = Get_Next_Typ(Inp_Typ)
                        If Inp_Typ = Red_T And i < InCnt Then EndAddr = EndAddr + 1
                    Next i
                Else: EndAddr = Addr + InCnt - 1
                End If
           End If
        End If
        If Addr <> EndAddr And EndAddr <> 0 Or (InCnt > 1 And EBit_Str <> "") Then
              Target = Addr & SBit_Str & " - " & EndAddr & EBit_Str
        Else: Target = Addr
        End If
     End If
  End If
End Sub




'-----------------------------------------------------------------------
Public Sub Global_Worksheet_SelectionChange(ByVal Target As Excel.Range)
'-----------------------------------------------------------------------
' Is called by event if the worksheet selection has changed
  If Target.CountLarge = 1 Then
     If Target.Row >= FirstDat_Row And Target.Row <= LastUsedRow() Then
           Make_sure_that_Col_Variables_match
           Application.EnableEvents = False ' Prevent calling the event multiple times
           If DEBUG_CHANGEEVENT Then Debug.Print Format(Time, "hh.mm.ss") & " SelectionChange" ' Debug
           
           Select Case Target.Column
              Case Enable_Col:  ' Enable / Disable the line
                                If Target.Value = "" Then
                                      Target.Value = ChrW(Hook_CHAR)
                                Else: Target.Value = ""
                                End If
                                BeepThis2 "Windows Balloon.wav"
                               
                                Update_Start_LedNr
                                Range(ActiveCell.Address).Offset(0, 1).Activate
              Case Inp_Typ_Col: Proc_Typ_Col Target
              Case Config__Col: ' Configuration column
                                If Target = "" And Not First_Change_in_Line(Target) Then SelectMacros
              Case LED_Nr__Col To LED_Nr__Col + INTERNAL_COL_CNT - 1:
                                ' The LEDs, InCnt, ... columns should not be changed easyly It's only possible if it is selected from the right side
                                Dim PriorCol As Long
                                On Error Resume Next ' In case PriorCell is not set
                                PriorCol = PriorCell.Column
                                On Error GoTo 0
                                If PriorCol <= Target.Column Then
                                   Range(ActiveCell.Address).Offset(0, LEDs____Col - Target.Column - 1).Activate        ' todo: should it not be -2?
                                End If
              
           End Select
           Set PriorCell = ActiveCell
           Application.EnableEvents = True
     ElseIf Target.Row = Header_Row Then
        With Target ' Display the comment below the cell to prevent cut of when first line is fixed and the window is scrolled down  ' 01.05.20:
           If Not .Comment Is Nothing Then
               With .Comment
                   With .Shape
                       .Top = ActiveWindow.VisibleRange.Top
                       .Left = ActiveCell.Left
                   End With
                   '.Visible = True
               End With
           End If
        End With
     End If
  End If
End Sub

' Problem:
' - Wenn eine Zelle Editiert wird und dann zu einem anderen Blatt gewechselt wird,
'   Dann wird der Global_Worksheet_Change Event ausgelöst.
'   Zu dem Zeitpunkt wurde aber bereits auf das neue Sheet umgeschaltet.
'   => Make_sure_that_Col_Variables_match erkennt ein anderes Sheet
'      Und die Änderungen am Ursprünglichen Sheet können nicht mehr überprüft werden
'
' - Wenn man Erkennt, das das Sheet gewechselt wurde, dann könnte man einfach zum alten Sheet
'   Zurück springen.


'--------------------------------------------------------------
Public Sub Global_Worksheet_Change(ByVal Target As Excel.Range)
'--------------------------------------------------------------
' This function is called if the worksheet is changed.
' It performs several checks after a user input depending form the column of the changed cell:

  Dim OldUppdating As Boolean, OldEvents As Boolean
  OldUppdating = Application.ScreenUpdating
  OldEvents = Application.EnableEvents
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  
  Make_sure_that_Col_Variables_match Switch_back_Target:=Target
  
  Dim Row As Long
  Row = Target.Row
  If Target.CountLarge = 1 Then
       If Row >= FirstDat_Row Then
           Dim Dat As String
           If Not IsError(Target.Value) Then ' Prevent problem if an invalid formula is entered (Example: "- Test")  27.10.20:
              If Target.Value <> "" And First_Change_in_Line(Target) Then
                 Format_Cells_to_Row Row + SPARE_ROWS  ' Add some reserve lines
                 Cells(Row, Enable_Col) = ChrW(Hook_CHAR) ' Must be set before the Start LEDNr is calculated
              End If
           End If
    
           Select Case Target.Column
              Case LED_Nr__Col, _
                   LEDs____Col, _
                   LED_Cha_Col To LED_Cha_Col + LED_CHANNELS - 1:           ' 26.04.20:
                                   Update_Start_LedNr
                                   Update_TestButtons Row
              Case SX_Bitposi_Col: Complete_Addr_Column_with_InCnt Row
                                   Update_TestButtons Row
              Case Inp_Typ_Col:
                                   Complete_Typ Target
                                   Complete_Addr_Column_with_InCnt Row
                                   Update_TestButtons Row
                                   Update_StartValue Row                    ' 01.05.20: Jürgen
                                   If Target = "" Then Del_Icons Target     ' 22.10.21:
              Case DCC_or_CAN_Add_Col, _
                   SX_Channel_Col:
                                   If Page_ID = "DCC" Then
                                      Complete_Addr_Column_with_InCnt Row
                                      Update_TestButtons Row
                                      Update_StartValue Row                 ' 01.05.20: Jürgen
                                   Else ' CAN
                                      'ActiveCell.Offset(0, 1).Select        ' 15.10.20:
                                   End If
              Case InCnt___Col:    Complete_Addr_Column_with_InCnt Row
                                   Update_TestButtons Row
           End Select
       End If
  Else ' More than one cell was changed
       'If Target.Rows(1).Cells.Count <> MAX_COLUMNS _
          And Target.Columns(1).Cells.Count <> MAX_ROWS Then  ' Not a whole row/column because it would take long and we dont wan't to get the hook if lines are deleted
            ' Adapt the first row to fit in the data range
            Dim StartRow As Long
            StartRow = Row
            If StartRow < FirstDat_Row Then StartRow = FirstDat_Row
            
            ' Don't fill in hooks if cells in the enable column should be changed/deleted
            If Target.Columns(1).Cells.Count <> MAX_ROWS Then               ' 26.09.19: Speed up deleeting whole columns
                If Range_is_Empty(Target) Then ' all cells in range have been deleted
                     ' Delete the icons                                     ' 22.10.21:
                     If Col_is_in_Range(Inp_Typ_Col, Target) Then
                        If Col_is_in_Range(MacIcon_Col, Target) Then
                              Del_Icons Range(Cells(Target.Row, Inp_Typ_Col), Cells(Target.Row + Target.Rows.Count - 1, MacIcon_Col))
                        Else: Del_Icons Range(Cells(Target.Row, Inp_Typ_Col), Cells(Target.Row + Target.Rows.Count - 1, Inp_Typ_Col))
                        End If
                     Else
                        If Col_is_in_Range(MacIcon_Col, Target) Then
                              Del_Icons Range(Cells(Target.Row, MacIcon_Col), Cells(Target.Row + Target.Rows.Count - 1, MacIcon_Col))
                        End If
                     End If
                Else ' Range is not empty
                     If Not Col_is_in_Range(Enable_Col, Target) Then
                        Range(Cells(StartRow, Enable_Col), Cells(Row + Target.Rows.Count - 1, Enable_Col)) = ChrW(Hook_CHAR) ' Must be set before the Start LEDNr is calculated
                     End If
                End If
            End If
            
            ' Format new cells at the end
            Dim EndRow As Long: EndRow = Row + Target.Rows.Count - 1 + SPARE_ROWS
            If EndRow > LastUsedRow + SPARE_ROWS Then EndRow = LastUsedRow  ' Protection if whole columns are deleted
            If EndRow >= FirstDat_Row Then Format_Cells_to_Row EndRow
            Update_Start_LedNr
       'End If
  End If
  If DEBUG_CHANGEEVENT Then Debug.Print Format(Time, "hh.mm.ss") & " Worksheet_Change " & Target.CountLarge ' Debug
  
  Application.EnableEvents = OldEvents
  Application.ScreenUpdating = OldUppdating
End Sub


'-------------------------------------
Public Sub Global_Worksheet_Activate()
'-------------------------------------
  If DEBUG_CHANGEEVENT Then Debug.Print ActiveSheet.Name & " Global_Worksheet_Activate event"
End Sub


'--------------------------------------
Public Sub Global_Worksheet_Calculate()
'--------------------------------------
' This function is called if a function is entered in the Excel sheet or if data
' are changed which influece a formula
  If DEBUG_CHANGEEVENT Then Debug.Print "Global_Worksheet_Calculate()" ' Debug
  
  ' Update_Start_LedNr                                                      ' 29.04.20: Disabled because I don't understand why it should be called here
End Sub

'-----------------------------------------------
Public Sub Update_StartValue(ByVal Row As Long)                            ' 01.05.20: Jürgen
'-----------------------------------------------
  Dim storeStatusType  As Integer
  Make_sure_that_Col_Variables_match
  Dim Channel_or_define As String, AddrStr As String
  Dim Addr As Long, AddrColumn As Long
  
  If (Page_ID = "DCC") Or (Page_ID = "CAN") Then
      AddrColumn = DCC_or_CAN_Add_Col
  ElseIf (Page_ID = "Selectrix") Then
      AddrColumn = SX_Channel_Col
  Else
      Exit Sub
  End If

  AddrStr = Cells(Row, AddrColumn)

  If IsNumeric(AddrStr) Then                                                ' 03.04.20:
        Addr = val(AddrStr)
  Else: Addr = -2 ' it's a variable
        Channel_or_define = AddrStr
  End If
  
  If Addr >= 0 Then
     Dim Inp_TypR As Range: Set Inp_TypR = Cells(Row, Inp_Typ_Col)
     If Page_ID = "DCC" Then                                                ' 11.10.20: Otherwise the input type is requested before the bit position on the SX page
        Complete_Typ Inp_TypR, True ' Check Inp_Typ. If not valid call the dialog
     End If
  Else
     Channel_or_define = AddrStr
  End If

  storeStatusType = Get_Store_Status(Row, Addr, Inp_TypR, Channel_or_define)
     
  Select Case storeStatusType
    Case SST_COUNTER_OFF:
      Exit Sub
  
    Case SST_COUNTER_ON:
      Exit Sub

    Case SST_S_ONOFF, SST_TRIGGER:
      Exit Sub
    Case Else:
      ' inform user later, when downloading to arduino
      'If Cells(Row, Start_V_Col) = AUTOSTORE_ON Then Cells(Row, Start_V_Col) = ""          '01.05.20: Commneted in Mail from Jürgen
      Exit Sub
  End Select
End Sub
'-----------------------------------------------

#If 1 Then ' 25.03.21: Faster routine
' It takes 12-13 seconds to update all buttons in Haralds example with 154 lines and 108 buttons
' The new function uses 110 ms
'
' Improvenents:
' - Existing buttons are reused and not deleted and created again
'   this saves 5-6 seconds
' - Parse the exisiting buttons only when the function is called the first time
'   The detected buttons are written in the global array Global_Rect_List
'   which has one entry per used excel row. The following calls just read the corrosponding
'   row and don't have to parse al objects again.
'   One "Do While i <= ActiveSheet.Rectangles.Count" loop takes 45 ms
'   for 154 rows that's 7 seconds

'-----------------------------------------------
Public Sub Update_TestButtons(ByVal Row As Long, Optional ByVal onValue As Integer = 0, Optional First_Call As Boolean = True)     ' 19.01.21: Jürgen
'-----------------------------------------------
    ' Debug.Print "Update_TestButtons(" & Row & ")"
    Dim objButton As Shape
    Dim Addr, ButtonPrefix As String
    Dim isSX, isDCC, incAddress, toggle As Boolean
    Dim i, InCnt, BitPos, ButtonCount, TextOffset, TextInc, AltTextOffset, Direction, ColorOffset, ColorInc, PixelOffset  As Integer
    Dim TargetColumn, AddrColumn As Long
    
    toggle = False                              ' 01.08.21: Jürgen bugfix
    ' only for the DCC and Selectrix sheet
    isSX = Page_ID = "Selectrix"
    isDCC = Page_ID = "DCC"
    
    If isDCC Then
        AddrColumn = DCC_or_CAN_Add_Col
        
    ElseIf isSX Then
        AddrColumn = SX_Channel_Col
    Else
        Exit Sub
    End If
    PixelOffset = 35
    
    Dim OldRect_List() As Long, OldRect_Cnt As Long, LastRow As Long            ' 25.03.21:
    LastRow = LastUsedRow()
    
    ButtonPrefix = "B"
    i = 1
    If First_Call Then
        ReDim Global_Rect_List(1 To LastRow)
        Do While i <= ActiveSheet.Rectangles.Count
            '  is it a SendButton Shape?
            With ActiveSheet.Rectangles(i)
                If .OnAction = "DCCSend" Then
                    Dim RectRow As Long
                    RectRow = .TopLeftCell.Row
                    If RectRow <= LastRow Then
                       Global_Rect_List(RectRow) = Global_Rect_List(RectRow) & i & " "
                    End If
                    If .TopLeftCell.Row = Row Then
                        'ActiveSheet.Rectangles(i).Delete                   ' 25.03.21:
                        ReDim Preserve OldRect_List(OldRect_Cnt)
                        OldRect_List(OldRect_Cnt) = i
                        OldRect_Cnt = OldRect_Cnt + 1
                    Else
                        ' find orphans
                        Addr = Get_First_Number_of_Range(.TopLeftCell.Row, AddrColumn)
                        If Addr = "" Then
                            ActiveSheet.Rectangles(i).Delete
                            i = i - 1
                        End If
                    End If
                End If
            End With
            i = i + 1
        Loop
    Else
        If Global_Rect_List(Row) = "" Then Exit Sub
        Dim Rect_List_In_Row() As String
        Rect_List_In_Row = Split(Trim(Global_Rect_List(Row)), " ")
        OldRect_Cnt = UBound(Rect_List_In_Row) + 1
        ReDim OldRect_List(OldRect_Cnt - 1)
        For i = 0 To OldRect_Cnt - 1
            OldRect_List(i) = CLng(Rect_List_In_Row(i))
        Next i
    End If
    
    InCnt = val(Cells(Row, InCnt___Col))
    If InCnt < 1 Then Exit Sub
    
    ButtonCount = 1 * InCnt
    Direction = 1                       ' start with "R"=0 or "G"=1
    ColorOffset = 0                     ' red
    ColorInc = 1                        ' take next color for next address
    TextInc = 1
    
    Make_sure_that_Col_Variables_match
    Set_Tast_Txt_Var  ' Set the global variables Red_T, Green_T, ...        ' 03.04.20:
    
    Addr = Get_First_Number_of_Range(Row, AddrColumn)
    If Addr = "" Or val(Addr) < 0 Then Exit Sub                             ' 04.05.20: Added: Or Val(Addr) < 0 to fix problems with the error returned by Get_First_Number_of_Range() (Mail from Jürgen)
    
    If isSX Then
        If Cells(Row, SX_Bitposi_Col) <> "" Then BitPos = Cells(Row, SX_Bitposi_Col)
        If (BitPos < 1 Or BitPos > 8) Then Exit Sub
        Addr = Addr * 8 + (BitPos - 1)
        TargetColumn = Inp_Typ_Col                                          ' 11.10.20: Old:  SX_Bitposi_Col  Warum diese Spalte ?
        TextOffset = BitPos                    ' starts with bitoffset
        AltTextOffset = BitPos                 ' starts with bitoffset
        Select Case Cells(Row, Inp_Typ_Col).Text
            Case Tast_T:
                            incAddress = True
                            If InCnt = 1 Then          ' sx buttons are always yellow
                                ColorOffset = ColorOffset + 2
                            End If
            Case OnOff_T:
                            toggle = True
                            ColorInc = 0
                            incAddress = True
                            ColorOffset = ColorOffset + 1
            Case Else:      Exit Sub                                        ' 11.10.20:
        End Select
    Else 'DCC
        TargetColumn = Inp_Typ_Col
        TextOffset = 1                      ' starts with this number for button text
        AltTextOffset = 2
        Select Case Cells(Row, Inp_Typ_Col).Text
            Case Red_T:
                            Direction = 0
            Case Green_T:
                            'incAddress = True                              ' 07.05.20: Disabled to be able to start with "Green". Otherwise the second, third, ... button will be shifted by one
                            ColorOffset = ColorOffset + 1
            Case OnOff_T:
                            toggle = True
                            incAddress = True
                            TextOffset = 0
                            AltTextOffset = TextOffset + 1
                            TextInc = 0
                            ColorInc = False
        End Select
        If InCnt > 1 Then      ' button signals start with text '0'
            TextOffset = 0
            AltTextOffset = 1
        End If
    End If

    Dim Height As Integer, NewCreated As Boolean
    With Cells(Row, TargetColumn)
        Height = .Height
        If Height > 13 Then Height = 13
        
        ' increase minimum column size if needed
        i = 1 + PixelOffset + ButtonCount * Height
        If .Width < i Then
            Dim factor
            factor = .Width / Columns(TargetColumn).ColumnWidth
            Range(Cells(, TargetColumn), Cells(, TargetColumn)).ColumnWidth = WorksheetFunction.RoundUp(i / factor, 1)
        End If
    End With

    Dim Used_OldRect As Long
    For i = 1 To ButtonCount
        If Used_OldRect < OldRect_Cnt Then
             Set objButton = ActiveSheet.Rectangles(OldRect_List(Used_OldRect)).ShapeRange.Item(1)
             Used_OldRect = Used_OldRect + 1
        Else
             Set objButton = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0)
             NewCreated = True
        End If
    
        Dim isSetToOn As Boolean                                            ' 19.01.21: Jürgen
        isSetToOn = False                                                   ' 26.03.21: Jürgen, otherwise lines with more than one toggle button are displayed wrong
        If toggle Then
            If (onValue And (2 ^ (i - 1))) <> 0 Then
                isSetToOn = True
            End If
        End If
        
        With objButton
            .Name = ButtonPrefix & Format(Addr, "0000") & "-" & Format(Direction, "00") & "-" & Format(ColorOffset, "00") & "-" & TextOffset
            If NewCreated Then
               .Left = Cells(Row, TargetColumn).Left + PixelOffset + (i - 1) * Height
               .Top = Cells(Row, TargetColumn).Top + 1
               .Height = Height - 2
               .Width = Height - 2
            End If
            If Cells(Row, Inp_Typ_Col).Text = OnOff_T Then ' "AnAus"       ' 03.04.20:
                  If .TextFrame2.TextRange.Text <> TextOffset Then .TextFrame2.TextRange.Text = TextOffset
            Else:
                  If .TextFrame2.TextRange.Text <> " " Then .TextFrame2.TextRange.Text = " " ' No text because it's confusing if 0/1 is used for OnOff switches
            End If
            If NewCreated Then
               .OnAction = "DCCSend"
               .DrawingObject.Border.Color = rgb(0, 0, 0)
               .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            End If
            .Fill.ForeColor.rgb = GetButtonColor(ColorOffset)
            If toggle Then
                Dim altText As String
                altText = ButtonPrefix & Format(Addr, "0000") & "-" & Format(1 - Direction, "00") & "-" _
                    & Format(ColorOffset + 1, "00") & "-" & AltTextOffset
                    
                If isSetToOn Then                                           ' 19.01.21: Jürgen
                    .AlternativeText = .Name
                    .Name = altText
                    .TextFrame2.TextRange.Text = Mid(.Name, 13, 1)
                    .Fill.ForeColor.rgb = GetButtonColor(val(Mid(.Name, 10, 2)))
                Else
                    .AlternativeText = altText
                End If
            Else                                    ' 01.08.21: Jürgen Bugfix
                .AlternativeText = ""
            End If
        End With
        TextOffset = TextOffset + TextInc
        AltTextOffset = AltTextOffset + TextInc
        
        If incAddress Then
            Addr = Addr + 1
        Else
            If (Direction = 1) Then
                Addr = Addr + 1
                Direction = 0
            Else
                Direction = 1
            End If
        End If
        ColorOffset = ValidateColorIndex(ColorOffset + ColorInc)
    Next
    
    For i = Used_OldRect To OldRect_Cnt - 1                                    ' 25.03.21:
        ActiveSheet.Rectangles(OldRect_List(i)).Delete
    Next
End Sub

#Else ' 25.03.21: Old routine
      ' Added "Unused" to be able to test it with the old function
'-----------------------------------------------
Public Sub Update_TestButtons(ByVal Row As Long, Optional ByVal onValue As Integer = 0, Optional Unused As Boolean)      ' 19.01.21: Jürgen
'-----------------------------------------------
    ' Debug.Print "Update_TestButtons(" & Row & ")"
    Dim objButton As Shape
    Dim Addr, ButtonPrefix As String
    Dim isSX, isDCC, incAddress, toggle As Boolean
    Dim i, InCnt, BitPos, ButtonCount, TextOffset, TextInc, AltTextOffset, Direction, ColorOffset, ColorInc, PixelOffset  As Integer
    Dim TargetColumn, AddrColumn As Long
    
    
    ' only for the DCC and Selectrix sheet
    isSX = Page_ID = "Selectrix"
    isDCC = Page_ID = "DCC"
    
    If isDCC Then
        AddrColumn = DCC_or_CAN_Add_Col
        PixelOffset = 35
    ElseIf isSX Then
        AddrColumn = SX_Channel_Col
        PixelOffset = 35                                                    ' 11.10.20: Old: 12
    Else
        Exit Sub
    End If
    
    ButtonPrefix = "B"
    i = 1
    Do While i <= ActiveSheet.Rectangles.Count
        '  is it a SendButton Shape?
        If ActiveSheet.Rectangles(i).OnAction = "DCCSend" Then
            If ActiveSheet.Rectangles(i).TopLeftCell.Row = Row Then
                ActiveSheet.Rectangles(i).Delete
            Else
                ' find orphans
                Addr = Get_First_Number_of_Range(ActiveSheet.Rectangles(i).TopLeftCell.Row, AddrColumn)
                If Addr = "" Then
                    ActiveSheet.Rectangles(i).Delete
                Else
                    i = i + 1
                End If
            End If
        Else
            i = i + 1
        End If
    Loop
    
    
    InCnt = val(Cells(Row, InCnt___Col))
    If InCnt < 1 Then Exit Sub
    
    ButtonCount = 1 * InCnt
    Direction = 1                       ' start with "R"=0 or "G"=1
    ColorOffset = 0                     ' red
    ColorInc = 1                        ' take next color for next address
    TextInc = 1
    
    Make_sure_that_Col_Variables_match
    Set_Tast_Txt_Var  ' Set the global variables Red_T, Green_T, ...        ' 03.04.20:
    
    Addr = Get_First_Number_of_Range(Row, AddrColumn)
    If Addr = "" Or val(Addr) < 0 Then Exit Sub                             ' 04.05.20: Added: Or Val(Addr) < 0 to fix problems with the error returned by Get_First_Number_of_Range() (Mail from Jürgen)
    
    If isSX Then
        If Cells(Row, SX_Bitposi_Col) <> "" Then BitPos = Cells(Row, SX_Bitposi_Col)
        If (BitPos < 1 Or BitPos > 8) Then Exit Sub
        Addr = Addr * 8 + (BitPos - 1)
        TargetColumn = Inp_Typ_Col                                          ' 11.10.20: Old:  SX_Bitposi_Col  Warum diese Spalte ?
        TextOffset = BitPos                    ' starts with bitoffset
        AltTextOffset = BitPos                 ' starts with bitoffset
        Select Case Cells(Row, Inp_Typ_Col).Text
            Case Tast_T:
                            incAddress = True
                            If InCnt = 1 Then          ' sx buttons are always yellow
                                ColorOffset = ColorOffset + 2
                            End If
            Case OnOff_T:
                            toggle = True
                            ColorInc = 0
                            incAddress = True
                            ColorOffset = ColorOffset + 1
            Case Else:      Exit Sub                                        ' 11.10.20:
        End Select
    Else 'DCC
        TargetColumn = Inp_Typ_Col
        TextOffset = 1                      ' starts with this number for button text
        AltTextOffset = 2
        Select Case Cells(Row, Inp_Typ_Col).Text
            Case Red_T:
                            Direction = 0
            Case Green_T:
                            'incAddress = True                              ' 07.05.20: Disabled to be able to start with "Green". Otherwise the second, third, ... button will be shifted by one
                            ColorOffset = ColorOffset + 1
            Case OnOff_T:
                            toggle = True
                            incAddress = True
                            TextOffset = 0
                            AltTextOffset = TextOffset + 1
                            TextInc = 0
                            ColorInc = False
        End Select
        If InCnt > 1 Then      ' button signals start with text '0'
            TextOffset = 0
            AltTextOffset = 1
        End If
    End If

    Dim Height As Integer
    Height = Cells(Row, TargetColumn).Height
    If Height > 13 Then Height = 13
    
    ' increase minimum column size if needed
    i = 1 + PixelOffset + ButtonCount * Height
    If Cells(Row, TargetColumn).Width < i Then
        Dim factor
        factor = Cells(Row, TargetColumn).Width / Columns(TargetColumn).ColumnWidth
        Range(Cells(, TargetColumn), Cells(, TargetColumn)).ColumnWidth = WorksheetFunction.RoundUp(i / factor, 1)
    End If

    For i = 1 To ButtonCount
        Set objButton = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 0, 0, 0, 0)
    
        Dim isSetToOn As Boolean                                            ' 19.01.21: Jürgen
        isSetToOn = False                                                   ' 26.03.21: Jürgen, otherwise lines with more than one toggle button are displayed wrong
        If toggle Then
            If (onValue And (2 ^ (i - 1))) <> 0 Then
                isSetToOn = True
            End If
        End If
        
        With objButton
            .Left = Cells(Row, TargetColumn).Left + PixelOffset + (i - 1) * Height
            .Top = Cells(Row, TargetColumn).Top + 1
            .Name = ButtonPrefix & Format(Addr, "0000") & "-" & Format(Direction, "00") & "-" & Format(ColorOffset, "00") & "-" & TextOffset
            .Height = Height - 2
            .Width = Height - 2
            If Cells(Row, Inp_Typ_Col).Text = OnOff_T Then ' "AnAus"       ' 03.04.20:
                  .TextFrame2.TextRange.Text = TextOffset
            Else: .TextFrame2.TextRange.Text = " " ' No text because it's confusing if 0/1 is used for OnOff switches
            End If
            .OnAction = "DCCSend"
            .Fill.ForeColor.rgb = GetButtonColor(ColorOffset)
            .DrawingObject.Border.Color = rgb(0, 0, 0)
            .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
            If toggle Then
                Dim altText As String
                altText = ButtonPrefix & Format(Addr, "0000") & "-" & Format(1 - Direction, "00") & "-" _
                    & Format(ColorOffset + 1, "00") & "-" & AltTextOffset
                    
                If isSetToOn Then                                           ' 19.01.21: Jürgen
                    .AlternativeText = .Name
                    .Name = altText
                    .TextFrame2.TextRange.Text = Mid(.Name, 13, 1)
                    .Fill.ForeColor.rgb = GetButtonColor(val(Mid(.Name, 10, 2)))
                Else
                    .AlternativeText = altText
                End If
            End If
        End With
        TextOffset = TextOffset + TextInc
        AltTextOffset = AltTextOffset + TextInc
        
        If incAddress Then
            Addr = Addr + 1
        Else
            If (Direction = 1) Then
                Addr = Addr + 1
                Direction = 0
            Else
                Direction = 1
            End If
        End If
        ColorOffset = ValidateColorIndex(ColorOffset + ColorInc)
    Next
End Sub
#End If


Public Sub ResetTestButtons(keepStatus As Boolean)                          ' 19.01.21: Jürgen
'----------------------------------------------------
' resets the on/off state of DCC buttons
' trigger buttons are not affected
'
' sets button to on if StartValue > 0
' sets button to off if StartValue = 0
' if stats value not set and keepStatus = false set button to off
' otherwise keep current on/off state
'
  Dim Row As Variant, First_Call As Boolean
  First_Call = True
  For Row = FirstDat_Row To LastUsedRow
    Dim sv As Long
    If Cells(Row, Start_V_Col) <> "" Then
        sv = val(Cells(Row, Start_V_Col))
        Update_TestButtons Row, sv, First_Call
        First_Call = False
    Else
        If keepStatus = False Then
            Update_TestButtons Row, First_Call:=First_Call             ' set to RED
            First_Call = False
        End If
    End If
  Next Row
End Sub

'UT--------------------------------
Private Sub Test_ResetTestButtons()                                         ' 25.03.21:
'UT--------------------------------
' Org: 12- 13 Sek
  ReDim Global_Rect_List(0)
  Application.ScreenUpdating = False
  Make_sure_that_Col_Variables_match
  Dim Start As Double
  Start_ms_Timer
  Start = Now
  ResetTestButtons False
  Debug.Print Now & " Duration: " & Format(Now - Start, "hh:mm:ss") & " Millis:" & Get_ms_Duration()
End Sub

'----------------------------------------------------------
Public Function GetButtonColor(ByVal Index As Byte) As Long
'----------------------------------------------------------
    Select Case Index
        Case 0: GetButtonColor = rgb(255, 128, 128)
        Case 1: GetButtonColor = rgb(128, 255, 128)
        Case 2: GetButtonColor = rgb(255, 255, 128)
        Case Else: GetButtonColor = rgb(255, 255, 255)
    End Select
End Function

'------------------------------------------------------
Public Function ValidateColorIndex(ByVal Index As Byte)
'------------------------------------------------------
    ' currently color 0..3 are valid
    ValidateColorIndex = Index Mod 4
End Function



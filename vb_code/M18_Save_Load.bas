Attribute VB_Name = "M18_Save_Load"
Option Explicit

' Save and load data from one or several Sheets
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Save the data from DCC, Selectrix and CAN sheets to a text file
'
' File Format
' ~~~~~~~~~~~
' - Tab separated text file
' - Header: Filetype, Version
' - Sub Header for each sheet: Sheet Typ, Sheet Name
' - Ext. MLL_pgf


' ToDo:
' - Es soll möglich sein, dass ein Sheet in ein Sheet geladen wird welches eine andere Page_ID hat
' - Wann sol gefragt werden welche Seiten gespeichert/importiert werden sollen
'   - Beim Import von einer alten Version sollen alle Seiten importiert werden
'   - Beim Speichern Menu werden alle oder alle ausgewählten Seiten gespeichert
'   - Beim Laden Menü wird anfangs gefragt ob alle oder nur bestimmte Seiten importiert werden sollen
'   - Mit einem Copy Befehl können Daten von einem Sheet in ein anderes Sheet kopiert werden.


Public Const PGF_Identification = "Program_Generator configuration file"
Public Const PGF_Version_String = "V1.0"

Private Const Head_ID = "Head:"
Private Const SheetID = "Sheet:"
Private Const Line_ID = "Line:"

Private Import_Page_ID As String
Private First_Line As Boolean
Private AddedToFilterColumn As Long
Private HiddenRows As Long
Private Start_Row As Long
Private LastA_Row As Long

Private Save_Data_FileName As String
Private Changed2SaveDir As Boolean
Private Copy_S2S_SrcSheet  As String
Private ImportFollowingSheets As Boolean

'----------------------------------------------------------
Private Function Create_pgf_File(Name As String) As Integer
'----------------------------------------------------------
  Dim fp As Integer
  fp = FreeFile
  On Error GoTo WriteError
  Open Name For Output As #fp
  Print #fp, Head_ID & vbTab & PGF_Identification & vbTab & PGF_Version_String
  Create_pgf_File = fp
  On Error GoTo 0
  Exit Function
  
WriteError:
  MsgBox Get_Language_Str("Fehler beim erzeugen der Programm Generator Konfigurationsdatei:") & vbCr & _
                          "  '" & Name & "'", vbCritical, Get_Language_Str("Fehler beim erzeugen der Konfigurationsdatei:")
End Function


'----------------------------------------------------------------------------------
Private Function Save_Sheet_to_pgf(fp As Integer, ByVal Sh As Worksheet) As Boolean
'----------------------------------------------------------------------------------
  On Error GoTo WriteError
  Dim OldSh As String
  OldSh = ActiveSheet.Name
  Sh.Select
  Page_ID = Cells(SH_VARS_ROW, PAGE_ID_COL)
  
  Print #fp, SheetID & vbTab; Page_ID & vbTab & Sh.Name
  Dim rng As Range, r As Variant
  If Selection.Cells.Count > 1 Then
        Set rng = Selection
  Else: Set rng = Range(Cells(FirstDat_Row, 1), Cells(LastFilledRowIn_ChkAll(Sh), 1))
  End If
  
  Col_from_Sheet = "" ' Force reloading the col variables                   ' 26.10.21:
  Make_sure_that_Col_Variables_match Sh                                     '  "
  
  For Each r In rng.Rows
      If Not r.Hidden And r.Row >= FirstDat_Row Then
         Dim Col As Long, Enabled As String
         If Cells(r.Row, Enable_Col) = ChrW(Hook_CHAR) Then
               Enabled = "Act"
         Else: Enabled = "-"
         End If
         Print #fp, Line_ID & vbTab & Enabled;
         For Col = Enable_Col + 1 To LastColumnDatSheet()                   ' 27.11.21: Old: LastUsedColumn
             Dim Line As String
             If Col = MacIcon_Col Then Col = Col + 1 ' Skip the Icon column (MacIcon_Col is -1 if the column doesn't exist)      ' 20.10.21
             If Col = LanName_Col Then Col = Col + 1 ' Use two lines to be able to enable both new columns separately                             '
             Print #fp, vbTab & Replace(Cells(r.Row, Col), vbLf, "{NewLine}");
         Next Col
         Print #fp, ""
      End If
  Next r
  On Error GoTo 0
  
  Sheets(OldSh).Select
  Save_Sheet_to_pgf = True
  Exit Function
  
WriteError:
  Sheets(OldSh).Select
End Function


'-----------------------------------------------------------------------------------------
Private Function Save_SingleSheet_to_pgf(Name As String, ByVal Sh As Worksheet) As Boolean
'-----------------------------------------------------------------------------------------
  If Not Is_Data_Sheet(Sh) Then
     MsgBox Replace(Get_Language_Str("Fehler: Die Seite '#1#' ist keine gültige Datenseite"), "#1#", Sh.Name), _
            vbCritical, Get_Language_Str("Ungültige Daten Seite ausgewählt")
     Exit Function
  End If
  
  Dim fp As Integer
  fp = Create_pgf_File(Name)
  If fp > 0 Then
     Save_SingleSheet_to_pgf = Save_Sheet_to_pgf(fp, Sh)
     Close #fp
  End If
End Function


'-----------------------------------------
Private Sub Test_Save_SingleSheet_to_pgf()
'-----------------------------------------
  Debug.Print "Save_SingleSheet_to_pgf: " & Save_SingleSheet_to_pgf(ThisWorkbook.Path & "\Test.MLL_pgf", ActiveSheet)
End Sub


'--------------------------------------------------------------------------------------
Public Function Save_Sheets_to_pgf(Name As String, FromAllSheets As Boolean) As Boolean
'--------------------------------------------------------------------------------------
  Dim Res As Boolean
  If FromAllSheets = False And ActiveWindow.SelectedSheets.Count = 1 Then
       Res = Save_SingleSheet_to_pgf(Name, ActiveSheet)
  Else
       Dim fp As Integer
       Res = True
       fp = Create_pgf_File(Name)
       If fp > 0 Then
          Dim Sh As Variant
          If ActiveWindow.SelectedSheets.Count > 1 Then
               For Each Sh In ActiveWindow.SelectedSheets
                   If Is_Data_Sheet(Sh) Then
                      Res = Save_Sheet_to_pgf(fp, Sh)
                      If Res = False Then Exit For
                   End If
               Next
          Else ' Only one sheet selected but "Save all" Checkbox activated
               For Each Sh In Sheets
                   If Is_Data_Sheet(Sh) And Sh.Visible = xlSheetVisible And Sh.Name <> "Examples" Then ' 17.10.20: Added: And Sh.Name <> "Examples" ' 25.04.20: Added: And Sh.Visible = xlSheetVisible
                      Res = Save_Sheet_to_pgf(fp, Sh)
                      If Res = False Then Exit For
                   End If
               Next Sh
          End If
          Close #fp
       End If
  End If
  
  If Res = False Then
     MsgBox Get_Language_Str("Fehler beim Schreiben der Programm Generator Konfigurationsdatei:") & vbCr & _
                             "  '" & Name & "'", vbCritical, Get_Language_Str("Fehler beim Schreiben der Konfigurationsdatei:")
  End If
  Save_Sheets_to_pgf = Res
End Function

'UT----------------------------------
Private Sub Test_Save_Sheets_to_pgf()
'UT----------------------------------
  Debug.Print "Save_Sheets_to_pgf=" & Save_Sheets_to_pgf(ThisWorkbook.Path & "\Test_All.MLL_pgf", True)
End Sub


'******************************** Load PGF *********************************

'--------------------------------------------------------------------------------
Private Function Find_Sheet_with_matching_Page_ID(Page_ID As String) As Worksheet
'--------------------------------------------------------------------------------
  Dim Sh As Worksheet
  For Each Sh In Sheets
    If Sh.Cells(SH_VARS_ROW, PAGE_ID_COL) = Page_ID Then
       Set Find_Sheet_with_matching_Page_ID = Sh
       Exit Function
    End If
  Next
End Function

'---------------------------------------------------------------------------------------
Private Function Copy_and_Clear_Sheet(SheetName As String, Page_ID As String) As Boolean
'---------------------------------------------------------------------------------------
  Dim Sh As Worksheet
  Set Sh = Find_Sheet_with_matching_Page_ID(Page_ID)
  If Sh Is Nothing Then
       MsgBox Get_Language_Str("Fehler: Es existiert keine passende Seite als Vorlage zum importieren der Daten"), vbCritical, Get_Language_Str("Fehler: Seite kann nicht angelegt werden")
       Exit Function
  Else
       Dim DstSh As Worksheet, s As Worksheet                               ' 10.10.20:
       For Each s In Worksheets ' Find the last data sheet
           If Is_Data_Sheet(s) Then Set DstSh = s
       Next s
       Sh.Copy after:=DstSh
       
       ActiveSheet.Name = SheetName                                         ' 25.04.20:
       Dim First_Row As Long
       First_Row = FirstDat_Row
       While Cells(First_Row, 1).EntireRow.Hidden
           First_Row = First_Row + 1
       Wend
       Rows(FirstDat_Row & ":" & LastUsedRow).ClearContents
       Copy_and_Clear_Sheet = True
  End If
End Function


'---------------------------------------------------------------------------------------------------------------------------------------
Private Function Open_or_Create_Sheet(SheetName As String, Inp_Page_ID As String, ByVal Name As String, ToActiveSheet As Boolean) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------
  Debug.Print "Open_or_Create_Sheet("
  Dim CreatedNewSheet As Boolean
  If ToActiveSheet Then
      Page_ID = ActiveSheet.Cells(SH_VARS_ROW, PAGE_ID_COL)
      If Not Is_Data_Sheet(ActiveSheet) Then
         MsgBox Get_Language_Str("Fehler: Die Ausgewählte Seite ist keine gültige Daten Seite"), vbCritical, Get_Language_Str("Falsche Seite ausgewählt")
         Exit Function
      End If
      Name = Copy_S2S_SrcSheet
      If Page_ID <> Inp_Page_ID And (Page_ID = "Selectrix" Or Inp_Page_ID = "Selectrix") Then
         MsgBox Get_Language_Str("Achtung: Die Adressen werden automatisch konvertiert. Sie müssen im Anschluss manuell überprüft werden."), _
                vbInformation, Get_Language_Str("Achtung: Anpassung der Adressen überprüfen")
      End If
  Else
      If Not SheetEx(SheetName) Then
            If Not Copy_and_Clear_Sheet(SheetName, Inp_Page_ID) Then Exit Function
            CreatedNewSheet = True
      Else: Sheets(SheetName).Activate
      End If
      Page_ID = ActiveSheet.Cells(SH_VARS_ROW, PAGE_ID_COL)
  End If
  Import_Page_ID = Inp_Page_ID
  First_Line = True
  
  Make_sure_that_Col_Variables_match
  Dim Row As Long, Col As Long, i As Long
  Row = LastFilledRowIn_ChkAll(ActiveSheet) + 2
  If CreatedNewSheet = False Then                                           ' 10.10.20:
     Cells(Row, Descrip_Col) = Get_Language_Str("Importiert von:") & Name
     Cells(Row, Enable_Col).ClearContents ' Is set by event => Clear it again
  End If
  
  Open_or_Create_Sheet = True
End Function

'----------------------------------------------------------------------
Private Sub Adapt_Adress_and_Typ_from_Selectrix(ByRef prLst As Variant)
'----------------------------------------------------------------------
' Adapt a DCC or CAN Address from Selectrix to DCC or CAN
  Dim Addr As Long
  If IsNumeric(prLst(DCC_or_CAN_Add_Col - 1)) Then
     Addr = 1 + prLst(DCC_or_CAN_Add_Col - 1) * 8
     If IsNumeric(prLst(DCC_or_CAN_Add_Col - 1 + 1)) Then Addr = Addr + prLst(DCC_or_CAN_Add_Col - 1 + 1)
     prLst(DCC_or_CAN_Add_Col - 1) = Addr
  End If
  
  If prLst(Inp_Typ_Col - 1 + 1) <> "" Then
     Set_Tast_Txt_Var
     If prLst(Inp_Typ_Col - 1 + 1) = Tast_T Then
        prLst(Inp_Typ_Col - 1 + 1) = ""          ' Don't know if Red or Green should be used
     End If
  End If
End Sub

'--------------------------------------------------------------------
Private Sub Adapt_Adress_and_Typ_to_Selectrix(ByRef prLst As Variant)
'--------------------------------------------------------------------
' Adapt a DCC or CAN Address to Selectrix
  Dim Addr As Long
  If IsNumeric(prLst(SX_Channel_Col - 1)) Then
     Addr = prLst(SX_Channel_Col - 1)
     If Addr > 0 Then Addr = Addr - 1
     prLst(SX_Channel_Col - 1) = Int(Addr / 8)
     prLst(SX_Bitposi_Col - 1) = Addr Mod 8
  End If
  
  If prLst(Inp_Typ_Col - 1) <> "" Then
     Set_Tast_Txt_Var
     If prLst(Inp_Typ_Col - 1) <> OnOff_T Then prLst(Inp_Typ_Col - 1) = Tast_T
  End If
End Sub

'----------------------------------------------------
Private Function Read_Line(Line As String) As Boolean
'----------------------------------------------------
  Dim Parts() As String
  Parts = Split(Line, vbTab)
  If Page_ID <> Import_Page_ID Then
     ' Problem "Selectrix" has an additional column "Bitposition"
     If Import_Page_ID = "Selectrix" Then
        Adapt_Adress_and_Typ_from_Selectrix Parts
        DeleteElementAt Inp_Typ_Col - 1, Parts ' delete the "Bitposition" column
     End If
     If Page_ID = "Selectrix" Then
        InsertElementAt Inp_Typ_Col - 2, Parts, "" ' insert the "Bitposition" column
        Adapt_Adress_and_Typ_to_Selectrix Parts
     End If
  End If
  
  Dim SkipLine As Boolean, Row As Long, Col As Long, i As Long
  Row = LastFilledRowIn_ChkAll(ActiveSheet)
  
  If First_Line Then
     First_Line = False
     Dim First_Row As Long
     First_Row = FirstDat_Row
     While Cells(First_Row, 1).EntireRow.Hidden
         First_Row = First_Row + 1
     Wend
     ' Check if the first line in the sheet an in the file is "RGB_Heartbeat(#LED)"
     ' The line is skipped
     If Cells(First_Row, Config__Col) = "RGB_Heartbeat(#LED)" And Cells(First_Row, Config__Col) = Parts(Config__Col - 1) Then
        Read_Line = True
        Exit Function
     End If
  End If
  

  If Not SkipLine Then
    Dim OldEvents As Boolean                                                ' 01.05.20: Prevent poping up dialogs
    OldEvents = Application.EnableEvents
    Application.EnableEvents = False
    Row = Row + 1
    Col = Enable_Col
    Cells(Row, Descrip_Col).Formula = "=""""" ' Otherwise empty rows are overwritten by the following call  ' 01.05.20: Useing formula instead of " "
    For i = 2 To UBound(Parts)
        Col = Col + 1
        If Col = MacIcon_Col Then Col = Col + 1 ' Skip the Icon column (MacIcon_Col is -1 if the column doesn't exist)      ' 20.10.21
        If Col = LanName_Col Then Col = Col + 1 ' Use two lines to be able to enable both new columns separately                             '
        Dim s As String
        s = Replace(Parts(i), "{NewLine}", vbLf)
        If Left(s, 2) = "==" Then s = "'" & s
        If Parts(i) <> "" Then Cells(Row, Col) = s
        
        If Col = Config__Col And s <> "" Then                               ' 22.10.21:
           FindMacro_and_Add_Icon_and_Name s, Row, ActiveSheet
        End If
        If Col = LED_Nr__Col Then                                           ' 25.10.21: (Hopefully) Prevent formating as date
           Cells(Row, Col).NumberFormat = "General"
        End If
    Next i
    
    
    If Cells(Row, Col).EntireRow.Hidden Then HiddenRows = HiddenRows + 1
    If Cells(Row, Filter__Col) <> "" Then AddedToFilterColumn = AddedToFilterColumn + 1
    If Start_Row = 0 Then Start_Row = Row
    LastA_Row = Row
    
    If Parts(1) = "Act" Then Cells(Row, Enable_Col) = ChrW(Hook_CHAR) ' Enable the Line   ' 01.05.20:
    Update_TestButtons Row
    Update_StartValue Row
    Application.EnableEvents = OldEvents
  End If
  
  Read_Line = True
End Function

'---------------------------------------------------------------------
Private Function PGF_has_Multiple_Sheets(Lines() As String) As Boolean
'---------------------------------------------------------------------
  Dim Line As Variant, SheetCnt As Long
  For Each Line In Lines
      Line = Replace(Line, vbLf, "")
      If Left(Line, Len(SheetID)) = SheetID Then
         SheetCnt = SheetCnt + 1
         If SheetCnt > 1 Then
            PGF_has_Multiple_Sheets = True
            Exit Function
         End If
      End If
  Next
End Function

'-------------------------------
Private Sub Check_Hidden_Lines()
'-------------------------------
  If AddedToFilterColumn Then
      Make_sure_that_Col_Variables_match
      If HiddenRows > 0 Then                                                ' 10.10.20:
         Range(Cells(Start_Row, Filter__Col), Cells(LastA_Row, Filter__Col)).Select
         'Application.ScreenUpdating = True
         Proc_Hide_Unhide
      End If
#If 0 Then                                                                  ' 06.08.20: Disabled
      If MsgBox(Get_Language_Str("Achtung es wurden Zeilen hinzugefügt welche durch die Filtereinstellungen " & _
                                 "im aktuellen Blatt ausgeblendet werden können!" & vbCr & _
                                 "Der Filter muss angepasst werden sonst werden diese Zeilen evtl. bei der nächsten " & _
                                 "Änderung wieder ausgeblendet." & vbCr & _
                                 vbCr & _
                                 "Bei dem normalerweise verwendeten Filter löscht man dazu die Einträge in der " & _
                                 "'Filter' Spalte." & vbCr & _
                                 "Alternativ können die Filtereinstellungen im Anschluss angepasst werden." & vbCr & _
                                 vbCr & _
                                 "Sollen die Einträge in der Filterspalte gelöscht werden damit die Einträge immer sichtbar sind ?"), vbQuestion + vbYesNo, _
                                 Get_Language_Str("Zeilen hinzugefügt welche durch den Filter ausgeblendet werden können")) = vbYes Then
          Selection.ClearContents
      End If
#End If
      If HiddenRows > 0 Then
         'Application.ScreenUpdating = False
         Cells(Start_Row, Descrip_Col).Select                               ' 10.10.20:
      End If
  End If
End Sub

'-----------------------------------------------------------------------------------------------------------------
Private Function Read_PGF_from_String_V1_0(Lines() As String, Name As String, ToActiveSheet As Boolean) As Boolean
'-----------------------------------------------------------------------------------------------------------------
  Dim LNr As Long, SkipSheet As Boolean, Inp_Page_ID As String, SheetName As String, Multiple_Sheets As Boolean
  Dim SheetCnt As Long, LineNrInSheet As Long
  Multiple_Sheets = PGF_has_Multiple_Sheets(Lines)
  AddedToFilterColumn = 0
  HiddenRows = 0
  Start_Row = 0
  ImportFollowingSheets = False
  
  Unload UserForm_Options ' Otherwise the status can't be shown             ' 06.08.20:
  
  For LNr = 1 To UBound(Lines) - 1
      Dim Parts() As String
      If Trim(Replace(Lines(LNr), vbLf, "")) <> "" Then                     ' 25.10.21:
          Parts = Split(Replace(Lines(LNr), vbLf, ""), vbTab)
          Select Case Parts(0)
             Case SheetID: ' Read Sheet type and name
                           Check_Hidden_Lines
                           If Is_Data_Sheet(ActiveSheet) Then                                             ' 26.10.20: Prevent crash if it's called after the installation when the "Start" sheet is active
                              Format_Cells_to_Row LastUsedRow() + SPARE_ROWS  ' Add some reserve lines    ' 07.10.20:
                              Update_Start_LedNr                                                          ' 07.10.20:
                           End If
                           Inp_Page_ID = Parts(1)
                           SheetName = Parts(2)
                           AddedToFilterColumn = 0
                           HiddenRows = 0
                           Start_Row = 0
                           If Multiple_Sheets Then
                                 If SheetName = "Examples" Then
                                      SkipSheet = True
                                 Else
                                      If ImportFollowingSheets Then
                                        SkipSheet = False
                                      Else
                                        Select Case MsgBox(Get_Language_Str("Soll die Seite") & " '" & SheetName & "' " & _
                                                           Get_Language_Str("importiert werden?"), vbQuestion + vbYesNoCancel, _
                                                           Get_Language_Str("Seite importieren?"))
                                           Case vbYes:
                                                       If GetAsyncKeyState(VK_CONTROL) <> 0 Then   ' Following function must be declared: Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
                                                           ImportFollowingSheets = True
                                                       End If
                                                       SkipSheet = False
                                           Case vbNo:  SkipSheet = True
                                           Case Else:  ' Abbort
                                                       If SheetCnt > 0 Then Format_Cells_to_Row LastUsedRow() + SPARE_ROWS  ' Add some reserve lines    ' 07.08.20:
                                                       Exit Function
                                        End Select
                                      End If
                                 End If
                           Else: SkipSheet = False
                           End If
    
                           If Not SkipSheet Then
                              If SheetCnt > 0 Then Format_Cells_to_Row LastUsedRow() + SPARE_ROWS  ' Add some reserve lines    ' 06.08.20:
                              SheetCnt = SheetCnt + 1
                              LineNrInSheet = 1
                              StatusMsg_UserForm.ShowDialog Get_Language_Str("Lade Seite '") & SheetName & "'", "..."          ' 06.08.20:
                              If Not Open_or_Create_Sheet(SheetName, Inp_Page_ID, Name, ToActiveSheet) Then Exit Function
                           End If
             Case Line_ID: ' Read sheet line
                           If Not SkipSheet Then
                              If Not Read_Line(Lines(LNr)) Then Exit Function
                              StatusMsg_UserForm.Set_ActSheet_Label "Line: " & LineNrInSheet                                   ' 06.08.20:
                              LineNrInSheet = LineNrInSheet + 1
                           End If
             Case Else:    MsgBox Get_Language_Str("Fehler: Unbekannter Zeilentyp in Zeile:") & " " & LNr & vbCr & _
                                  Get_Language_Str("in der PGF Datei:") & vbCr & _
                                  "  '" & Name & "'", vbCritical, Get_Language_Str("Fehler in PGF Datei")
                                  Exit Function
          End Select
      End If
  Next LNr
  Check_Hidden_Lines
  Format_Cells_to_Row LastUsedRow() + SPARE_ROWS  ' Add some reserve lines    ' 01.05.20:
  Update_Start_LedNr                                                          ' 28.09.20: Added by Jürgen. Otherwise NUM_LEDS is 0
  Read_PGF_from_String_V1_0 = True
End Function



'-------------------------------------------------------------------------------------------
Public Function Read_PGF(ByVal Name As String, Optional ToActiveSheet As Boolean) As Boolean
'-------------------------------------------------------------------------------------------
  Dim FileStr As String, Lines() As String, Parts() As String, Err As Boolean
  ImportFollowingSheets = False
  FileStr = Read_File_to_String(Name)
  If FileStr = "#ERROR#" Then Exit Function
  Lines = Split(FileStr, vbCr)
  If UBound(Lines) <= 1 Then
     MsgBox Get_Language_Str("Fehler: Die PGF Datei enthält keine Daten:") & vbCr & _
            "  '" & Name & "'", vbCritical, Get_Language_Str("Ungültige PGF Datei")
     Exit Function
  End If
  
  Parts = Split(Lines(0), vbTab)
  Err = (UBound(Parts) < 2)
  If Not Err Then Err = (Parts(0) <> Head_ID)
  If Not Err Then Err = (Parts(1) <> PGF_Identification)
  If Not Err Then
     Dim VerStr As String, ScrUpd As Boolean
     ScrUpd = Application.ScreenUpdating
     Application.ScreenUpdating = False
     VerStr = Parts(2)
     Select Case VerStr
         Case PGF_Version_String: Read_PGF = Read_PGF_from_String_V1_0(Lines, Name, ToActiveSheet)
                                  StatusMsg_UserForm.Hide
         Case Else: Err = True
     End Select
     Application.ScreenUpdating = ScrUpd
  End If
  
  If Err Then
     MsgBox Get_Language_Str("Fehler: Die PGF Datei enthält keinen gültigen Header:") & vbCr & _
            "  '" & Name & "'", vbCritical, Get_Language_Str("Ungültige PGF Datei")
  End If
End Function


'UT------------------------
Private Sub Test_Read_PGF()
'UT------------------------
  'Debug.Print "Read_PGF=" & Read_PGF(ThisWorkbook.Path & "\Test.MLL_pgf")
  Application.EnableEvents = False
  Debug.Print "Read_PGF=" & Read_PGF("C:\Dat\Märklin\Arduino\LEDs_Eisenbahn\Doc\von anderen\Dominik\Prog_Gen_Data_06_08_2020.MLL_pgf")
  Debug.Print "Application.EnableEvents:" & Application.EnableEvents
  Application.EnableEvents = True
End Sub


'**********************************************************************************************

'-------------------------------------------
Public Function Get_MyExampleDir() As String
'-------------------------------------------
  Dim Dir As String
  Dir = Environ("USERPROFILE") & "\Documents\" & MyExampleDir
  CreateFolder Dir & "\"
  Get_MyExampleDir = Dir
End Function


'--------------------------------------------------------------------------------------------
Private Sub Save_Data_to_File_CallBack(Do_Import As Boolean, Import_FromAllSheets As Boolean)
'--------------------------------------------------------------------------------------------
  Debug.Print "Save_Data_to_File_CallBack(" & Do_Import & ", " & Import_FromAllSheets & ")"
  
  If Do_Import Then
     If Save_Sheets_to_pgf(Save_Data_FileName, Import_FromAllSheets) Then
        MsgBox Get_Language_Str("Die Datei wurde erfolgreich geschrieben:") & vbCr & _
                                "  '" & Save_Data_FileName & "'", vbInformation, "Datei wurde geschrieben"
     End If
  End If
End Sub

'------------------------------------------------------------
Private Sub Achtivate_MyExampleDir_if_called_the_first_time()
'------------------------------------------------------------
  If Changed2SaveDir = False Then
     ChDir (Get_MyExampleDir())
     Changed2SaveDir = True
  End If
End Sub

'-----------------------------
Public Sub Save_Data_to_File()
'-----------------------------
' Is called from the options dialog
  Achtivate_MyExampleDir_if_called_the_first_time
     
  Dim ExampleName As String
  ExampleName = "Prog_Gen_Data_" & Replace(Date, ".", "_")
  Dim Res As Variant
  Res = Application.GetSaveAsFilename(InitialFileName:=ExampleName, fileFilter:=Get_Language_Str("Program Generator File (*.MLL_pgf), *.MLL_pgf"), Title:=Get_Language_Str("Dateiname zum abspeichern der Daten angeben"))
  If Res <> False Then
     If Dir(Res) <> "" Then
        If MsgBox(Get_Language_Str("Achtung die Datei existiert bereits!" & vbCr & _
                                   vbCr & _
                                   "Soll die Datei überschrieben werden?"), vbQuestion + vbOKCancel, Get_Language_Str("Existierende Datei überschreiben?")) <> vbOK Then
           Exit Sub
        End If
     End If
     Save_Data_FileName = Res
     Remove_Selections_in_all_Data_Sheets
     Import_Hide_Unhide.Start "Save_Data_to_File_CallBack", Import_FromAll:=ActiveWindow.SelectedSheets.Count - 1
  End If
End Sub

'-------------------------------
Public Sub Load_Data_from_File()
'-------------------------------
' Is called from the options dialog
  Achtivate_MyExampleDir_if_called_the_first_time
  
  Dim Res As Variant
  Res = Application.GetOpenFilename(fileFilter:=Get_Language_Str("Program Generator File (*.MLL_pgf), *.MLL_pgf"), Title:=Get_Language_Str("Dateiname zum Importieren der Daten angeben"))
  If Res <> False Then
     Read_PGF Res
  End If
End Sub


'***************************************************************************************************

'--------------------------------------------------------------
Private Sub Copy_to_Selected_Sheet_Callback(Do_Copy As Boolean)
'--------------------------------------------------------------
  If Do_Copy Then
     If ActiveSheet.Name = Copy_S2S_SrcSheet Then
          MsgBox Get_Language_Str("Achtung: Zum Kopieren der Daten von einer Seite auf eine andere Seite müssen zwei " & _
                                  "verschiedene Seiten ausgewählt werden." & vbCr & _
                                  "Dazu wählt man im folgenden Dialog die gewünschte Zielseite über die Reiter am " & _
                                  "unteren Rand der Seite aus BEVOR man 'OK' betätigt."), _
                                  vbInformation, Get_Language_Str("Zielseite wurde nicht ausgewählt")
          Select_Dest_Sheet.Start "Copy_to_Selected_Sheet_Callback"
     Else
          Read_PGF Save_Data_FileName, ToActiveSheet:=True
     End If
  End If
End Sub

'--------------------------------------------------------------------------------------------------------------
Private Sub Save_Data_from_active_Sheet_to_File_CallBack(Do_Import As Boolean, Import_FromAllSheets As Boolean)
'--------------------------------------------------------------------------------------------------------------
  Debug.Print "Save_Data_from_active_Sheet_to_File_CallBack(" & Do_Import & ", " & Import_FromAllSheets & ")"
  
  If Do_Import Then
     ActiveSheet.Select ' Disable multiple selected sheets
     Copy_S2S_SrcSheet = ActiveSheet.Name
     If Save_Sheets_to_pgf(Save_Data_FileName, False) Then
        Select_Dest_Sheet.Start "Copy_to_Selected_Sheet_Callback"
     End If
  End If
End Sub

'------------------------------------
Public Sub Copy_from_Sheet_to_Sheet()
'------------------------------------
' Is called from the options dialog
  If MsgBox(Get_Language_Str("Mit dieser Funktion können ausgewählte Daten aus der " & _
                              "aktuellen Seite in eine andere Seite kopiert werden."), _
                              vbOKCancel, Get_Language_Str("Kopieren von Daten von dieser Seite in eine andere Seite")) <> vbOK Then
     Exit Sub
  End If
  
  Save_Data_FileName = Get_MyExampleDir() & "\Copy_Sheet2Sheet.MLL_pgf"
  Remove_Selection_in_Sheet ActiveSheet
  Import_Hide_Unhide.Start "Save_Data_from_active_Sheet_to_File_CallBack", Import_FromAll:=-2
End Sub



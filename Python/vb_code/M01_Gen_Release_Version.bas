Attribute VB_Name = "M01_Gen_Release_Version"
Option Explicit

' Call the following function to generate a release version
'--------------------------------
Private Sub Gen_Release_Version()
'--------------------------------
  Release_or_Debug_Version True
  Set_Language_Def -1
  Set_Lib_Macros_Test_Language -1                                       ' 28.11.21:
End Sub


'------------------------------
Private Sub Gen_Debug_Version()
'------------------------------
  Release_or_Debug_Version False
End Sub


'------------------------------------------------------------
Private Sub Update_Version_InActSheet(SetCursorPos As String)               ' 06.05.20:
'------------------------------------------------------------
  On Error GoTo ErrMsg
  ActiveSheet.Shapes.Range(Array("Version_TextBox")).Select
  Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = Prog_Version
  ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange(1), Address:="", ScreenTip:=Str(Now)
  On Error GoTo 0
  
  Range("A1").Activate ' Scroll to the top
  
  ' Set Cursor out of the version text box
  If SetCursorPos <> "" Then
        Range(SetCursorPos).Select
  Else
        Dim Row As Long
        Row = FirstDat_Row
        Do While Cells(Row, 1).EntireRow.Hidden
           Row = Row + 1
        Loop
        Cells(Row, Descrip_Col).Select
  End If
  Exit Sub
  
ErrMsg:
  MsgBox "Error: 'Version_TextBox' not found in sheet '" & ActiveSheet.Name & "'", vbCritical, "Internal Error"
End Sub


'-------------------------------------------------------
Private Sub Release_or_Debug_Version(Release As Boolean)
'-------------------------------------------------------
  Dim Sh As Variant, LastSheet As String
  LastSheet = ActiveSheet.Name
  ' AddImagesToTreeForm ".\Icons", True                                     ' 28.11.21: For some reasons this call generates an error. The same function could be called by the "Read Pictures" button in the "Lib_Macros" sheet without problems ?!?
                                                                            '           => It has to be called manually
  
  If Release Then                                               ' 12.03.21: Juergen
     ThisWorkbook.Sheets(LANGUAGES_SH).Visible = False
     ThisWorkbook.Sheets(LIBMACROS_SH).Visible = False
     ThisWorkbook.Sheets(PAR_DESCR_SH).Visible = False
     ThisWorkbook.Sheets(LIBRARYS__SH).Visible = False
     ThisWorkbook.Sheets(PLATFORMS_SH).Visible = False
  End If
  
  For Each Sh In ThisWorkbook.Sheets
      If Is_Data_Sheet(Sh) Then
         Sh.Select
         ActiveWindow.Zoom = 100                                            ' 07.10.20: Not included in Ver 1.9.5F
         ActiveWindow.ScrollColumn = 1                                      ' 18.10.20:
         Make_sure_that_Col_Variables_match Sh
         
         Update_Version_InActSheet ""
         
         ' Show / Hide the internal variables
         With Range(Cells(SH_VARS_ROW, 1), Cells(SH_VARS_ROW, LastUsedColumn())).Font
            If Release Then
                  .ThemeColor = xlThemeColorDark1
            Else: .ColorIndex = xlAutomatic
            End If
         End With
         ' This internal data are always shown
         Cells(SH_VARS_ROW, BUILDOP_COL).Font.ColorIndex = xlAutomatic
         Cells(SH_VARS_ROW, COMPort_COL).Font.ColorIndex = xlAutomatic
         Cells(SH_VARS_ROW, COMPrtR_COL).Font.ColorIndex = xlAutomatic
         Cells(SH_VARS_ROW, BUILDOpRCOL).Font.ColorIndex = xlAutomatic
         
         ' Show / Hide the internal columns
         Cells(1, InCnt___Col).EntireColumn.Hidden = False ' 26.09.19: Don't hide the internal columns because this generates problems when
         Cells(1, LocInCh_Col).EntireColumn.Hidden = False '           lines are copied while the AutoFilter is activ. Prior the columns
                                                           '           have been hidden in release mode
         
         ' Build otions
         Cells(SH_VARS_ROW, BUILDOP_COL) = "'" & AUTODETECT_STR & " " & BOARD_NANO_OLD & " " & DEFARDPROG_STR
         If Page_ID <> "CAN" Then
           Cells(SH_VARS_ROW, BUILDOpRCOL) = "'" & AUTODETECT_STR & " " & BOARD_NANO_OLD & " " & DEFARDPROG_STR
         End If
         
         ' Activate the Filter                                              ' 02.03.20: ' 14.06.20: Filters ar no longer used in the release version
         'Range(Cells(Header_Row, Enable_Col), Cells(LastUsedRow(), LastUsedColumn())).AutoFilter Field:=2, _
               Criteria1:="=B01", Operator:=xlOr, Criteria2:="="
      End If
  Next Sh
  
  ' Show / Hide the internal sheets
  Sheets(LIBMACROS_SH).Visible = Not Release
  Sheets(PAR_DESCR_SH).Visible = Not Release
  Sheets(PLATFORMS_SH).Visible = Not Release
  'Sheets("Farbentest").Visible = Not Release
  
  ' Start sheet
  Sheets(START_SH).Select
  ActiveSheet.Unprotect
  Update_Version_InActSheet "M7" ' Cursor Below the picture
  If Release Then
        Sheets(START_SH).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        ActiveWindow.DisplayHeadings = False
  Else: Sheets(START_SH).Unprotect
  End If

  If Release Then Set_Config_Default_Values_for_Release

  Clear_COM_Port_Check_and_Set_Cursor_in_all_Sheets Release

  If Release Then
        ' move cursor to begin of sheet data                               ' 10.03.21  Juergen
        For Each Sh In ThisWorkbook.Sheets
            If Is_Data_Sheet(Sh) Then
               Sh.Select
               Make_sure_that_Col_Variables_match Sh
               Columns(Enable_Col).ColumnWidth = 5.8
               If Sh.Name = "Examples" Then
                   Columns(Filter__Col).ColumnWidth = 11
                   Columns(Inp_Typ_Col).ColumnWidth = 16
               Else
                   Columns(Filter__Col).ColumnWidth = 5.8
                   Columns(Inp_Typ_Col).ColumnWidth = 12
               End If
               If Page_ID <> "Selectrix" Then
                   Columns(DCC_or_CAN_Add_Col).ColumnWidth = 11.57
               Else
                   Columns(SX_Channel_Col).ColumnWidth = 13.29
                   Columns(SX_Bitposi_Col).ColumnWidth = 9.57
               End If
               Columns(Start_V_Col).ColumnWidth = 4.57
               Columns(Descrip_Col).ColumnWidth = 43.5
               Columns(Dist_Nr_Col).ColumnWidth = 8
               Columns(Conn_Nr_Col).ColumnWidth = 8.86
               Columns(Config__Col).ColumnWidth = 60
               Columns(LED_Nr__Col).ColumnWidth = 4.71
               Columns(LEDs____Col).ColumnWidth = 7
               Columns(InCnt___Col).ColumnWidth = 4.71
               Columns(LocInCh_Col).ColumnWidth = 4.71
               Columns(LED_Cha_Col).ColumnWidth = 4.71
               Cells(Header_Row + 1, Descrip_Col).Select
            End If
        Next
        Sheets(START_SH).Select
  Else: Sheets(LastSheet).Select
  End If
  
  If Release Then Write_Default_CheckColors_Parameter_File                  ' 01.12.19:
  RemoveExistingExtensions                                                  ' 18.02.22: Juergen
End Sub


'------------------------------------------------------
Public Sub Set_Config_Default_Values_at_Program_Start()
'------------------------------------------------------
  Set_Bool_Config_Var "Lib_Installed_other", False
End Sub


'--------------------------------------------------
Private Sub Set_Config_Default_Values_for_Release()
'--------------------------------------------------
  Set_String_Config_Var "MinTime_House", ""
  Set_String_Config_Var "MaxTime_House", ""
  Set_String_Config_Var "DCC_Offset", ""
  Set_String_Config_Var "Color_Test_Mode", "1"
  Set_String_Config_Var "USE_SPI_Communication", "0"         ' 14.05.20: Deactivated per default because it's easier to use a 3.9K Pull-Down (See "TX LED Problem mit Pull Down Widerstand beheben.docx")
  Set_String_Config_Var "Use_Excel_Console", "0"
  Set_String_Config_Var "LEDNr_Display_Type", "1"            ' 03.04.21: Juergen - new option
  Set_String_Config_Var "Expert_Mode_aktivate", ""           ' 06.10.21:
  Set_String_Config_Var "Use_TreeView_for_Macros", ""        ' 06.10.21: "" to ask the user which dialog should be used
  Set_String_Config_Var "Show_Icon_Column", "1"              ' 25.10.21:
  Set_String_Config_Var "Show_Simple_Names", "1"
  Set_String_Config_Var "Show_Macros_Column", "1"
  Set_String_Config_Var "Use_PlatformIO", "0"                ' 21.01.22: Juergen - new option
  Set_String_Config_Var "SimPosX", "800"                     ' 04.03.22: Juergen - new option
  Set_String_Config_Var "SimPosY", "200"                     ' 04.03.22: Juergen - new option
  Set_String_Config_Var "SimLedsX", "8"                      ' 04.03.22: Juergen - new option
  Set_String_Config_Var "SimLedsY", "8"                      ' 04.03.22: Juergen - new option
  Set_String_Config_Var "SimLedSize", "24"                   ' 04.03.22: Juergen - new option
  Set_String_Config_Var "SimAutostart", "0"                  ' 04.03.22: Juergen - new option
  Set_String_Config_Var "SimOffset", "0"                     ' 04.03.22: Juergen - new option
  Set_String_Config_Var "SimOnTop", "1"                      ' 04.03.22: Juergen - new option
End Sub

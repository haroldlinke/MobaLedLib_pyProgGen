Attribute VB_Name = "M09_Language"
Option Explicit


' Attention: One of the following preprcessor constants have to be defined in
' "Extras / Eigenschafteb VBA Projekt"'
'   PATTERN_CONFIG_PROG
'   PROG_GENERATOR_PROG
'

' Other languages could be added in the hidden sheet LANGUAGES_SH = "Languages"
' In addition Get_ExcelLanguage() must be adapted

' Current languages:
'  0 = German
'  1 = English
'  2 = Dutch
'  3 = French
'  4 = Italian  (Prog_Generator only)
'  5 = Spain        "
'  6 = Danish       "

' Strings which have not been translated could be found with
' The seach expression '[!r]("'  and ' "'   (Without ')
' and enabled "Mit Mustervergleich"
'   See also: https://docs.microsoft.com/de-de/office/vba/language/reference/user-interface-help/wildcard-characters-used-in-string-comparisons
' They could be translated by inserting 'Get_Language_Str'
'
'
' There are a lot of different locations where text messages are used:
' - In the sheets
' - Error messages in the sheets which are shown on certain condition
' - Hints in the sheets
' - Dialog boxes
'   - Some of them are changed from the program code
'   - Some messages are located in separate sheets: Special_Mode_Dlg, Par_Description
' - In the program code
' - Buttons
' - The Hotkeys have also to be adapted in the dialogs and single buttons
' - In the configuration files *.MLL_pcf
'   - Variable names
'   - Text messages
' - ...
'
' Some messages in the program are only called in case of an error.
' => They are not added automatically to the 'Languages' sheet
' ==> This is done by calling Add_All_VBA_Strings_to_the_Languages_Sheet()
'
' Location for the translations
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' - The most translations are stored in the hidden "Languages"
' - Some translations are stored in different locations
'   - The start page contains one text box for each language.
'     This is used because here several colors and fonts are used.
'     Normaly only the box which contains the current language
'     is visible. The language number is stored in the field
'     "Alternativtext" which could be changed using the right mouse
'     and the menu "Größe und Eigneschaften".
'     The "Alternativtext" mut contain the keyword "Language:"
'     Example: "Language: 0"
'     If a new language is created one text box has to be copied
'     and translated and the "Alternativtext" has to be changed to
'     the new number.
'   - The text messages in the example sheets in the Pattern_Configurator
'     have an own text box for each language. Here only the text box
'     for the actual language is visible. This method is used to be
'     independand from the Excel program. The number after the
'     keyword "msoTextBox" defines the language number. Old *.MLL_pcf
'     don't have a tailing number they are always shown.
'   - The buttons in the "Morsecode" sheet use a button for each
'     language fo the same reason as the text boxes.
'     (See Activate_Language_in_Example_Sheet(ByVal sh As Worksheet)
'     The Language number is stored in the "AlternativeText".
'   - Dialog functions which select their data form an Excel sheet
'     like the "SelectMacros_Form" or the "UserForm_Other" use
'     separate columns in the sheets "Lib_Macros" and "Par_Description
'
'
'
' Language specific messages in the example sheets                          ' 11.02.20:
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Different languages could be used in the "msoTextBox" lines in the example sheets.
' The language number is added to the token:
'  "msoTextBox0" = German
'  "msoTextBox1" = Englisch
'  "msoTextBox2" = Dutch
'  "msoTextBox3" = French
' All Lines starting with "msoTextBox" are loaded to the sheet, but only the
' text box with the current language is visible. The Language is stored in
' ShapeRange.AlternativeText = "Language: 0"
' If the line line starts with "msoTextBox" it's an old file without different languages.
' This textbox is always shown.
' If the token has a tailing number it's assumed that there are matching lines for all languages.
' If the current language is missing nothing is displayed ;-( => It's importand to update
' the examples if a new language is added
'
'
'
Public Const FirstLangRow = 3  ' Row number 3 in the languages sheet
Public Const LangType_Col = 1  ' This column contains the typ in the languages sheet
Public Const LangParamCol = 2  ' This column contains parameters in the languages sheet
Public Const FirstLangCol = 3  ' This is the first column used for the translations in then languages sheet


' Debuging with othwer languages
Private Const Simulate_Language = -1 ' -1 = Disable, 0 = German, 1 = Englisch, 2 = Dutch
                                     ' If this flag is not set the language could be changed
                                     ' temporary with the function "Test_Translations()"

Private Check_Languages As Boolean   ' Is used to simulate a different language for tests Check_Languages = true
Private Test_Language As Integer     ' The language number which should be used for the test

Public Red_T As String     ' "Rot"
Public Green_T As String   ' "Grün"
Public OnOff_T As String   ' "AnAus"
Public Tast_T As String    ' "Tast"

Public Const Virtual_Channel_T As String = "V"         ' 18.02.22 Juergen

'--------------------------------------------------------------------------------------------------------------
Private Function Update_Language_in_Sheet(ByVal Sh As Worksheet, DestLang As Integer) As Boolean  ' Old Name: Update_Language_in_Pattern_Config_Sheet
'--------------------------------------------------------------------------------------------------------------
#If PROG_GENERATOR_PROG Then
  Make_sure_that_Col_Variables_match Sh   ' Set Page_ID
#End If
  
  Dim LSh As Worksheet ', OldSel As Variant
  'Set OldSel = Selection
  Set LSh = ThisWorkbook.Sheets(LANGUAGES_SH)
  With Sh
    ' Check if the language has to be changed by comparing the first string
    Dim c As Long, FirstMsg As String, ActLang As Integer

    ' Check the actual used language in the sheet
    ActLang = -1
    FirstMsg = .Range(LSh.Cells(FirstLangRow, LangParamCol))
    For c = FirstLangCol To LastUsedColumnInRow(LSh, FirstLangRow)
       If FirstMsg = LSh.Cells(FirstLangRow, c) Then
          ActLang = c - FirstLangCol
          Exit For
        End If
    Next
    If ActLang = -1 Then
       MsgBox "Error: '" & FirstMsg & "' not found in the 'Languages' sheet", vbCritical, "Language Error"
       Exit Function
    End If
    If (ActLang = DestLang) Then Exit Function ' Language is correct
    
    ' Unprotect the sheet
    Dim WasProtected As Boolean
    WasProtected = Sh.ProtectContents                     ' 04.06.20: In Prog_Gen ActiveSheet... was used
    If WasProtected Then Sh.Unprotect
    
    ' Debug
    Dim MaxLang As Integer
    MaxLang = LastUsedColumnInRow(LSh, FirstLangRow) - FirstLangCol
    If DestLang = -1 Then ' -1 could be used for debugging
       DestLang = ActLang + 1
       If ActLang >= MaxLang Then DestLang = 0
    End If
    
    
    ' Replace the texts
    Dim r As Long, LangCol As Integer
    LangCol = FirstLangCol + DestLang
    For r = FirstLangRow To LastUsedRowIn(LSh)
        Dim Param As String, Tmp As String
        Param = LSh.Cells(r, LangParamCol)
        If Param <> "" Then
            Tmp = LSh.Cells(r, LangCol)
            Select Case LSh.Cells(r, LangType_Col)
               Case "":             ' Nothing
               Case "Cell":         .Range(Param).FormulaR1C1 = Tmp ' Cell values and formulas
             #If PROG_GENERATOR_PROG Then
               Case "Cell_DCC":     If Page_ID = "DCC" Then .Range(Param).FormulaR1C1 = Tmp ' Cell values and formulas
               Case "Cell_SX":      If Page_ID = "Selectrix" Then .Range(Param).FormulaR1C1 = Tmp ' Cell values and formulas
               Case "Cell_CAN":     If Page_ID = "CAN" Then .Range(Param).FormulaR1C1 = Tmp ' Cell values and formulas
             #End If
               Case "NumberFormat": .Range(Param).NumberFormat = Tmp
               Case "Button":
             #If PROG_GENERATOR_PROG Then
                                    .Shapes.Range(Array(Param)).Select
                                    Selection.Characters.Text = Tmp
             #Else
                                    .Shapes.Range(Array(Param)).Item(1).DrawingObject.Caption = Tmp '07.03.20:
             #End If
               Case "Comment":      ' Comments
                                    With .Range(Param).Comment
                                      .Text Text:=Replace(Tmp, vbLf, Chr(10))
                                      .Shape.TextFrame.Characters(1, Len(Tmp)).Font.Bold = False  ' ToDo: Bolt aus Sheet lesen
                                    End With
             #If PROG_GENERATOR_PROG Then
               Case "Comment_DCC":  ' Comments DCC
                                    If Page_ID = "DCC" Then
                                        If Not .Range(Param).Comment Is Nothing Then ' 26.10.21:
                                            With .Range(Param).Comment
                                              .Text Text:=Replace(Tmp, vbLf, Chr(10))
                                              .Shape.TextFrame.Characters(1, Len(Tmp)).Font.Bold = False  ' ToDo: Bolt aus Sheet lesen
                                            End With
                                        End If
                                    End If
               Case "Comment_SX":  ' Comments Selectrix
                                    If Page_ID = "Selectrix" Then
                                        If Not .Range(Param).Comment Is Nothing Then ' 26.10.21:
                                            With .Range(Param).Comment
                                              .Text Text:=Replace(Tmp, vbLf, Chr(10))
                                              .Shape.TextFrame.Characters(1, Len(Tmp)).Font.Bold = False  ' ToDo: Bolt aus Sheet lesen
                                            End With
                                        End If
                                    End If
               Case "Comment_CAN":  ' Comments CAN
                                    If Page_ID = "CAN" Then
                                        If Not .Range(Param).Comment Is Nothing Then ' 26.10.21:
                                            With .Range(Param).Comment
                                              .Text Text:=Replace(Tmp, vbLf, Chr(10))
                                              .Shape.TextFrame.Characters(1, Len(Tmp)).Font.Bold = False  ' ToDo: Bolt aus Sheet lesen
                                            End With
                                        End If
                                    End If
             #End If
               Case "ErrorMessage": .Range(Param).Validation.ErrorMessage = Tmp
               Case "ErrorTitle":   .Range(Param).Validation.ErrorTitle = Tmp
               ' ToDo Warnungen
            End Select
        End If
    Next r
    
  End With
  If WasProtected Then Protect_Active_Sheet
  'OldSel.Select
  Update_Language_in_Sheet = True
  'Debug.Print "Updated Language in " & Sh.Name ' Debug
End Function

'UT-------------------------------------------------------
Private Sub Test_Update_Language_in_Sheet()
'UT-------------------------------------------------------
' Switches to the next language
  Dim OldEvents As Boolean, OldUpdate As Boolean
  OldEvents = Application.EnableEvents
  OldUpdate = Application.ScreenUpdating
  Application.EnableEvents = False
  Application.ScreenUpdating = False

  Update_Language_in_Sheet ActiveSheet, -1
  Application.EnableEvents = OldEvents
  Application.ScreenUpdating = OldUpdate
End Sub


#If PROG_GENERATOR_PROG Then
'-------------------------------------------------------------------------------
Private Function Update_Language_in_Config_Sheet(DestLang As Integer) As Boolean    ' 29.02.20:
'-------------------------------------------------------------------------------
  Dim Row As Long, Col As Long, Sh As Worksheet
  Col = 2
  Set Sh = ThisWorkbook.Sheets(ConfigSheet)                                 ' 29.04.20: Added "ThisWorkbook." to prevent problems
  
  With Sh
    If .Range("A1") = DestLang Then Exit Function ' Pos A1 contains the actual language number (White text on white ground)
    For Row = 1 To LastUsedRowIn(Sh)
        With .Cells(Row, Col)
         If .Value <> "" Then
            .Value = Get_Language_Str(.Value)
         End If
         If Not .Comment Is Nothing Then
           .Comment.Text Text:=Get_Language_Str(.Comment.Text)
         End If
        End With
    Next Row
    
    ' Special cells
    With .Range("Lib_Installed_other")
       If .Value <> "" Then
          .Value = Get_Language_Str(.Value)
       End If
    End With
    .Range("A1") = DestLang
  End With
  Update_Language_in_Config_Sheet = True
End Function
#End If

#If 0 Then
'-------------------------------------------------
Private Sub Test_Update_Language_in_Config_Sheet()
'-------------------------------------------------
  Test_Language = 0
  Check_Languages = 1
  Debug.Print Update_Language_in_Config_Sheet(Test_Language)
End Sub
#End If

'--------------------------------------------------------------------
Private Sub Activate_Language_in_Example_Sheet(ByVal Sh As Worksheet)       ' 11.02.20:
 '--------------------------------------------------------------------
  Dim o As Variant, ActLanguage As Integer, LanguageNr As Integer
  ActLanguage = Get_ExcelLanguage()
  For Each o In Sh.Shapes
    Select Case o.Type
      Case msoTextBox:         ' 17: TextBox
                               If left(o.Name, Len("Goto_Graph")) <> "Goto_Graph" And o.Name <> "InternalTextBox" Then  ' "InternalTextBox" = "by Hardi"
                                  If left(o.AlternativeText, Len("Language: ")) = "Language: " Then '  11.02.20:
                                     LanguageNr = val(Mid(o.AlternativeText, Len("Language: ") + 1))
                                     o.Visible = (LanguageNr = ActLanguage)
                                  End If
                               End If
      Case msoFormControl:     ' 8: Button                                  ' 19.10.19:
                               If o.AlternativeText <> "_Internal_Button_" And o.AlternativeText <> "Add_Del_Button" Then
                                  If left(o.AlternativeText, Len("Language: ")) = "Language: " Then
                                     LanguageNr = val(Mid(o.AlternativeText, Len("Language: ") + 1))
                                     o.Visible = (LanguageNr = ActLanguage)
                                  End If
                                  'ToDo     Print_Typ_and_Pos FileNr, "msoFormControl", o, o.AlternativeText & Chr(pcfSep) & Replace(o.OnAction, ThisWorkbook.Name & "!", "")
                               End If
     #If PROG_GENERATOR_PROG Then
      ' Translate the buttons.
      ' Attention: This part is not working in the Pattern_Configurator.
      '            ==> It overwrites the Button texts ;-(
      '            The problem is, that the buttons already had an alternative text. (05.06.20: Deleted the text)
      '            => At the moment (05.06.20) the Buttons "Import vom Prog. Gen.",
      '               "Programm Genarator" and "Zum Modul schicken" are not translated
      Case msoOLEControlObject:
                               If o.AlternativeText = "" Then o.AlternativeText = o.DrawingObject.Object.Caption ' Store the original text
                               o.DrawingObject.Object.Caption = Get_Language_Str(o.AlternativeText)
                               ' Todo: Change also the activation key
     #End If
    End Select
  Next o
End Sub

'UT--------------------------------------------
Private Sub Activate_Language_in_Active_Sheet()
'UT--------------------------------------------
  Activate_Language_in_Example_Sheet ActiveSheet
End Sub



'----------------------------------
Sub Update_Language_in_All_Sheets()                                         ' 25.02.20: Old name: Update_Language_in_All_Pattern_Config_Sheets
'----------------------------------
' Set the language of all sheets to the active display language in excel

  Dim OldEvents As Boolean, OldUpdate As Boolean, OldSheet As Worksheet
  Set OldSheet = ActiveSheet
  OldEvents = Application.EnableEvents
  OldUpdate = Application.ScreenUpdating
  Application.EnableEvents = False
  Application.ScreenUpdating = False
  
  
  Dim Sh As Variant, DestLang As Integer, Initialized As Boolean
  DestLang = Get_ExcelLanguage
  'DestLang = 1 ' debug
  For Each Sh In ThisWorkbook.Sheets
     'If sh.Name <> LANGUAGES_SH And sh.Name <> GOTO_ACTIVATION_SH And sh.Name <> PAR_DESCRIPTION_SH Then ' 20.01.20: Old
   #If PATTERN_CONFIG_PROG Then
     If Sh.Name = MAIN_SH Or Is_Normal_Data_Sheet(Sh.Name, Get_Language_Str("übersetzt")) Then
   #Else
     If Sh.Name = START_SH Or Is_Data_Sheet(Sh) Then
   #End If
        If Not Initialized Then StatusMsg_UserForm.Set_Label Get_Language_Str("Umstellung der Sprache")      '12.02.20:
        StatusMsg_UserForm.Set_ActSheet_Label Sh.Name
        If Not Initialized Then StatusMsg_UserForm.Show
        Initialized = True
        ' 07.03.20: sh.Activate ' Necessary for the button names ;-(
        #If PATTERN_CONFIG_PROG Then
          Translate_Standard_Description_Box_in_Sheet Sh                      ' 07.03.20: Using new function which doesn't need the active sheet
          If Update_Language_in_Sheet(Sh, DestLang) = False And Sh.Name = MAIN_SH Then Exit For
        #End If
        
        #If PROG_GENERATOR_PROG Then                                        ' 23.02.20:
           If Is_Data_Sheet(Sh) Then
              If Update_Language_in_Sheet(Sh, DestLang) Then
                 Translate_Example_Texts_in_Sheet Sh
              End If
           End If
        #End If
     Activate_Language_in_Example_Sheet Sh
     End If
  Next Sh
  
#If PROG_GENERATOR_PROG Then                                                ' 29.02.20:
  Update_Language_in_Config_Sheet DestLang
  
  Update_Language_Name_Column_in_all_Sheets                                 ' 26.10.21:
#End If

  Unload StatusMsg_UserForm
  If Not OldSheet Is Nothing Then                                           ' 01.05.20: Prevent crash if downloaded from the internet when the Savety messages is shown
     OldSheet.Activate
  End If
  Application.EnableEvents = OldEvents
  Application.ScreenUpdating = OldUpdate
End Sub

#If PATTERN_CONFIG_PROG Then ' 04.06.20: Old block in Pattern_config
'-----------------------------------
Public Sub Check_SIMULATE_LANGUAGE()
'-----------------------------------
  If Simulate_Language >= 0 Then
     MsgBox "Attention the compiler switch 'SIMULATE_LANGUAGE' is set to " & Simulate_Language & vbCr & _
            "This is used to test the languages. It must be disabled in the release version!", vbInformation
  End If
End Sub
#End If

#If PROG_GENERATOR_PROG Then
'------------------------------------------
Private Function Get_Language_Def() As Long
'------------------------------------------
  Dim LangStr As String
  LangStr = ThisWorkbook.Sheets(ConfigSheet).Range("Language_Def")
  If IsNumeric(LangStr) Then
        Get_Language_Def = val(LangStr)
  Else: Get_Language_Def = -1
  End If
End Function

'----------------------------------------------
Public Sub Set_Language_Def(LanguageNr As Long)
'----------------------------------------------
  ThisWorkbook.Sheets(ConfigSheet).Range("Language_Def") = LanguageNr
End Sub

#End If


'---------------------------------------------
Public Function Get_ExcelLanguage() As Integer
'---------------------------------------------
' Return a number corrosponding to the actual language used in excel
'  0 = German
'  1 = English and all other languages
'  2 = Dutch
'  3 = French 
'  4 = Italian    (Only in the Prog_Generator)
'  5 = Spain        "
' The number must match to the position in the language strings in M6_Language_Constants
'
' Is working if the office language is changed or the Window language
#If PATTERN_CONFIG_PROG Then
  If Simulate_Language >= 0 Then
    Get_ExcelLanguage = Simulate_Language
    Exit Function
  End If
#End If

  If Check_Languages Then                                                   ' 20.01.20:
     Get_ExcelLanguage = Test_Language
     Exit Function
  End If
  Get_ExcelLanguage = 1

#If PROG_GENERATOR_PROG Then
  Dim Simulate_Language As Long                                             ' 03.03.20:
  Simulate_Language = Get_Language_Def()
  If Simulate_Language >= 0 Then
    Get_ExcelLanguage = Simulate_Language
    Exit Function
  End If
 #End If

'Debug.Print "Achtung: Get_ExcelLanguage() liefert immer 1"
'Exit Function

  Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
     ' Language ID's: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa432635(v=office.12)
     Case msoLanguageIDGerman, msoLanguageIDGermanAustria, msoLanguageIDGermanLiechtenstein, msoLanguageIDGermanLuxembourg, msoLanguageIDSwissGerman
          Get_ExcelLanguage = 0
     Case msoLanguageIDDutch, msoLanguageIDBelgianDutch         '   Added by Misha 24-2-2020
          Get_ExcelLanguage = 2
     Case msoLanguageIDFrench, msoLanguageIDBelgianFrench, msoLanguageIDFrenchCameroon, msoLanguageIDFrenchCanadian, msoLanguageIDFrenchCotedIvoire, _
          msoLanguageIDFrenchHaiti, msoLanguageIDFrenchLuxembourg, msoLanguageIDFrenchMali, msoLanguageIDFrenchMonaco, msoLanguageIDFrenchMorocco, _
          msoLanguageIDFrenchReunion, msoLanguageIDFrenchSenegal, msoLanguageIDFrenchWestIndies, msoLanguageIDSwissFrench  ' Not defined at Exel 2016? msoLanguageIDFranchCongoDRC
          Get_ExcelLanguage = 3 ' ' Added by Misha 24-2-2020
     Case msoLanguageIDItalian, msoLanguageIDSwissItalian
          Get_ExcelLanguage = 4
     Case msoLanguageIDSpanish, msoLanguageIDSpanishArgentina, msoLanguageIDSpanishBolivia, msoLanguageIDSpanishChile, msoLanguageIDSpanishColombia, _
          msoLanguageIDSpanishCostaRica, msoLanguageIDSpanishDominicanRepublic, msoLanguageIDSpanishEcuador, msoLanguageIDSpanishElSalvador, _
          msoLanguageIDSpanishGuatemala, msoLanguageIDSpanishHonduras, msoLanguageIDSpanishModernSort, msoLanguageIDSpanishNicaragua, _
          msoLanguageIDSpanishPanama, msoLanguageIDSpanishParaguay, msoLanguageIDSpanishPeru, msoLanguageIDSpanishPuertoRico, _
          msoLanguageIDSpanishUruguay, msoLanguageIDSpanishVenezuela, msoLanguageIDMexicanSpanish
          Get_ExcelLanguage = 5
     Case msoLanguageIDDanish   ' 16.06.20:
          Get_ExcelLanguage = 6
  End Select
  
End Function

'UT---------------------------------
Private Sub Test_Get_ExcelLanguage()
'UT---------------------------------
  Debug.Print "Get_ExcelLanguage()=" & Get_ExcelLanguage()
End Sub


'-------------------------------------------------------------
Function Find_Cell_Pos_by_Name(ByVal Desc As String) As String
'-------------------------------------------------------------
' Find the given Desc in the language sheet and return the
' destination position as string.
'
' If Desc is not found "" is returned
  Dim LSh As Worksheet, Res As Variant
  Set LSh = ThisWorkbook.Sheets(LANGUAGES_SH)
  With LSh
    Set Res = LSh.Cells.Find(What:=Desc, after:=.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
    If Not Res Is Nothing Then
       Find_Cell_Pos_by_Name = .Cells(Res.Row, LangParamCol).Value
    End If
  End With
  
End Function

'UT-------------------------------------
Private Sub Test_Find_Cell_Pos_by_Name()
'UT-------------------------------------
  Debug.Print "Find_Cell_Pos_by_Name(""Bits pro Wert:"")='" & Find_Cell_Pos_by_Name("Bits pro Wert:") & "'"
  Debug.Print "Find_Cell_Pos_by_Name(""Bits per value:"")='" & Find_Cell_Pos_by_Name("Bits per value:") & "'"
End Sub

'-------------------------------------------------------
Function Get_German_Name(ByVal Desc As String) As String
'-------------------------------------------------------
  Dim LSh As Worksheet, Res As Variant
  Set LSh = ThisWorkbook.Sheets(LANGUAGES_SH)
  With LSh
    Set Res = LSh.Cells.Find(What:=Desc, after:=.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
    If Not Res Is Nothing Then
          Get_German_Name = .Cells(Res.Row, FirstLangCol).Value
    Else: MsgBox "Error '" & Desc & "' not found in 'Get_German_Name'", vbCritical, "Error"
    End If
  End With
End Function

'------------------------------------------------------------
Private Sub Add_Entry_to_Languages_Sheet(GermanTxt As String)
'------------------------------------------------------------
  If Get_ExcelLanguage() = 0 Then
    Dim LSh As Worksheet, Row As Long
    Set LSh = ThisWorkbook.Sheets(LANGUAGES_SH)
    With LSh
      Row = LastUsedRowIn(LSh) + 1
      .Cells(Row, FirstLangCol) = GermanTxt
    End With
  End If
End Sub

'------------------------------------------------------
Private Sub Debug_Find_Diff(s1 As String, s2 As String)
'------------------------------------------------------
  Debug.Print "Len:" & Len(s1) & "  " & Len(s2)
  Dim i As Long
  For i = 1 To Len(s1)
    If Mid(s1, i, 1) <> Mid(s2, i, 1) Then
       Debug.Print "Different pos: " & i
       i = i - 3
       If i < 1 Then i = 1
       Debug.Print Mid(s1, i - 3, 7)
       Debug.Print Mid(s2, i - 3, 7)
       Debug.Print Space(i - 1) & "^"
       Exit Sub
    End If
  Next
  Debug.Print "Strings are equal"
End Sub

'--------------------------------------------------------------------------------------------------------------
Public Function Find_Language_Str_Row(ByVal Desc As String, Optional ByVal Look_At As Integer = xlPart) As Long
'--------------------------------------------------------------------------------------------------------------
  Dim LSh As Worksheet, Res As Variant
  Set LSh = ThisWorkbook.Sheets(LANGUAGES_SH)
  With LSh
    On Error Resume Next ' In case the Description is missing
    ' Maximal length for the find command: 255
    Dim Start As Range, Retry As Boolean, FirstPos As Range
    Set Start = .Range("A1")
    Desc = RTrim(Desc)
    Do
       If Len(Desc) > 90 Then ' For some reasons exel can't find longer strings !?! Maybe it's a problem with strings that contain two vblf? One vblf is no problem
          Look_At = xlPart    ' 27.03.20: After changeing SearchFormat to true it seames to work ? (Sometimes !?!)
       End If
       Set Res = LSh.Cells.Find(What:=left(Desc, 90), after:=Start, LookIn:=xlFormulas, LookAt:=Look_At, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=True) ' 22.01.20: Check only the first 100 characters
       If Look_At = xlPart And Not IsEmpty(Res) Then
          Retry = (RTrim(Res.Value) <> RTrim(Desc))  ' Use RTrim() to ignore tailing space characters
          'Debug_Find_Diff Res.Value, Desc ' Debug
          If Retry Then
             'If Len(Desc) >= 255 Then
             '   Debug.Print "Debug long string in Get_Language_Str"
             'End If
             Set Start = Res
             If FirstPos Is Nothing Then
                  Set FirstPos = Res
             Else
                  If Res.Address = FirstPos.Address Then
                     Retry = False
                     Set Res = Nothing
                  End If
             End If
          End If
       End If
    Loop While Retry
    
    On Error GoTo 0
    
    If Not IsEmpty(Res) Then                                                ' 05.01.20:
        If Not Res Is Nothing Then
           Find_Language_Str_Row = Res.Row
           #If 0 Then
               With Res.Interior ' Debug: Mark the used entries
                  .PatternColorIndex = xlAutomatic
                  .Color = 5296274
               End With
           #End If
        End If
    End If
  End With
End Function



#If True Then                                                               ' 28.01.20:
'------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Get_Language_Str(ByVal Desc As String, Optional GenError As Boolean = False, Optional ByVal Look_At As Integer = xlPart) As String
'------------------------------------------------------------------------------------------------------------------------------------------------
' Replace the vbCr in the input string to be able to find in in the "Languages" sheet      ' 30.01.20:
' Problem:
' vbCr cant be used in an excel table (Languages sheet)
' => It is replaced by '|' in the sheet
' To identify the structure of the message an vbLf is also used.
' - If the original string is stored in a dialog it contains cvCr & vbLf
'   only the vbCr has to be replaced
' - If the original string is stored in the VB code only vbCr is used
'   therefore the cvLf is added in "Add_All_VBA_Strings_to_the_Languages_Sheet()"
'   to get the visual line breaks.
'   In this case the cvLf must also be added in this function

  If Desc = "" Then Exit Function                                           ' 23.01.20:
  
  Dim Use_CrLf As Boolean
  Desc = Replace(Desc, "| ", "|")                                           ' 24.02.20: Added because of Misha's improvement
  If InStr(Desc, vbCr & vbLf) > 0 Then
       Use_CrLf = True
       Desc = Replace(Desc, vbCr, "|")
  Else ' no combination of vbCr & vbLf used => Add vbLf for the check
       Desc = Replace(Desc, vbCr, "|" & vbLf)
  End If
  
  Dim Res As String
  Res = Get_Language_Str_Sub(Desc, GenError, Look_At)
  
  If Use_CrLf Then
        Get_Language_Str = Replace(Res, "|", vbCr)
  Else: Get_Language_Str = Replace(Replace(Res, "|" & vbLf, vbCr), "| " & vbLf, vbCr) ' 07.10.21: Added: replace by additional space '"| " & vbLf, vbCr),' to prevent problems in Display_Navigation_Keys()
  End If

End Function

'-------------------------------------------------------------------------------------------------------------------
Private Function Get_Language_Str_Sub(ByVal Desc As String, GenError As Boolean, ByVal Look_At As Integer) As String
'-------------------------------------------------------------------------------------------------------------------
' Find the given Desc in the language sheet and return the
' string in the actual language
'
  Dim Row As Long
  If IsNumeric(Desc) Then                                                   ' 05.06.20:
     Get_Language_Str_Sub = Desc
     Exit Function
  End If
  
  Row = Find_Language_Str_Row(Desc, Look_At)
  
  Dim LSh As Worksheet
  Set LSh = ThisWorkbook.Sheets(LANGUAGES_SH)
  With LSh
    If Row > 0 Then
       Get_Language_Str_Sub = .Cells(Row, FirstLangCol + Get_ExcelLanguage()).Value
       If Get_Language_Str_Sub <> "" Then                                    ' 25.01.20: Prior the function was left with an empty result ;-( Now a message is generated an the Germal text is used
          Exit Function
       End If
    End If
  
    ' The language string was not found
    If GenError Then
          MsgBox "Error translation missing in sheet 'Languages' for:" & vbCr & _
                 "  '" & Desc & "'", vbCritical, "Error: Translation missing"
    Else:
          Debug.Print "*** Translation not found for:" & vbCr & Desc
    End If
    
    If Row > 0 Then Get_Language_Str_Sub = .Cells(Row, FirstLangCol + 1).Value    ' Use the Enlish text if available  ' 26.01.20:
    If Get_Language_Str_Sub = "" Then Get_Language_Str_Sub = Desc
  End With
  
  Add_Entry_to_Languages_Sheet Desc ' Add the german text to the 'Languages' sheet  ' 24.01.20:
End Function


#Else ' Old 29.01.20:
'-----------------------------------------------------------------------------------------------------------------------------------------
Function Get_Language_Str(ByVal Desc As String, Optional GenError As Boolean = True, Optional ByVal Look_At As Integer = xlPart) As String
'-----------------------------------------------------------------------------------------------------------------------------------------
' Find the given Desc in the language sheet and return the
' string in the actual language
'
  If Desc = "" Then Exit Function                                           ' 23.01.20:
  Dim LSh As Worksheet, Res As Variant
  Set LSh = ThisWorkbook.Sheets(LANGUAGES_SH)
  With LSh
    On Error Resume Next ' In case the Description is missing
    ' Maximal length for the find command: 255
    Dim Start As Range, Retry As Boolean, FirstPos As Range
    Set Start = .Range("A1")
    Desc = RTrim(Desc)
    Do
       If Len(Desc) > 255 Then
          Look_At = xlPart
       End If
       Set Res = LSh.Cells.Find(What:=left(Desc, 255), after:=Start, LookIn:=xlFormulas, LookAt:=Look_At, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False) ' 22.01.20: Check only the first 255 characters
       If Look_At = xlPart And Not IsEmpty(Res) Then
          Retry = (RTrim(Res.Value) <> Desc)  ' Use RTrim() to ignore tailing space characters
          If Retry Then
             'If Len(Desc) >= 255 Then
             '   Debug.Print "Debug long string in Get_Language_Str"
             'End If
             Start = Res
             If FirstPos Is Nothing Then
                  Set FirstPos = Res
             Else
                  If Res.Address = FirstPos.Address Then
                     Retry = False
                     Set Res = Nothing
                  End If
             End If
          End If
       End If
    Loop While Retry
    
    'Set Res = LSh.Cells.Find(What:=Desc, After:=.Range("A1"), LookIn:=xlFormulas, LookAt:=Look_At, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
    On Error GoTo 0
    If Not IsEmpty(Res) Then                                                ' 05.01.20:
        If Not Res Is Nothing Then
           Get_Language_Str = .Cells(Res.Row, FirstLangCol + Get_ExcelLanguage()).Value
           If Get_Language_Str <> "" Then                                    ' 25.01.20: Prior the function was left with an empty result ;-( Now a message is generated an the Germal text is used
              Exit Function
           End If
           
        End If
    End If
  
    ' The language string was not found
    If GenError Then
          MsgBox "Error translation missing in sheet 'Languages' for:" & vbCr & _
                 "  '" & Desc & "'", vbCritical, "Error: Translation missing"
    Else:
          Debug.Print "*** Translation not found for:" & vbCr & Desc
    End If
    
    If Not Res Is Nothing Then Get_Language_Str = .Cells(Res.Row, FirstLangCol + 1).Value    ' Use the Enlish text if available  ' 26.01.20:
    If Get_Language_Str = "" Then Get_Language_Str = Desc
  End With
  
  Add_Entry_to_Languages_Sheet Desc ' Add the german text to the 'Languages' sheet  ' 24.01.20:
End Function
#End If

'-------------------------------------------------
Private Sub Change_Lang_in_MultiPage(o As Variant)                          ' 25.01.20:
'-------------------------------------------------
  Dim Pg As Variant
  For Each Pg In o.Pages
     'Pg.Caption = Replace(Get_Language_Str(Replace(Pg.Caption, vbCr, "|"), False, xlWhole), "|", vbCr)   ' Old 30.01.20:
      Pg.Caption = Get_Language_Str(Pg.Caption, False, xlWhole)
  Next Pg
End Sub

'-------------------------------------------
Sub Change_Language_in_Dialog(dlg As Object)
'-------------------------------------------
Dim o As Variant
   dlg.Caption = Get_Language_Str(dlg.Caption)                              ' 25.01.20:
   For Each o In dlg.Controls
      If o.ControlTipText <> "" Then o.ControlTipText = Get_Language_Str(o.ControlTipText)
      'Debug.Print o.Name
      'If o.Name = "Label12" Then
      '   Debug.Print "Debug"
      'End If
      On Error Resume Next  ' Problem: if o.Caption doesn't exist "<Objekt unterstützt diese Eigenschaft oder Methode nicht>"
        If left(o.Name, Len("MultiPage")) = "MultiPage" Then
           Change_Lang_in_MultiPage o
        ElseIf left(o.Name, Len("ListBox")) = "ListBox" Then
           ' There is no constant text in the ListBox dialog. It is loaded dymanicaly
        ElseIf o.Caption <> "by Hardi" And left(o.Caption, 4) <> "http" And o.Caption <> "*" And o.Caption <> "J" Then   ' "J" = Smilly in "Wait_CheckColors_Form"
          'o.Caption = Replace(Get_Language_Str(Replace(o.Caption, vbCr, "|"), False, xlWhole), "|", vbCr)   ' 30.01.20: Old:
           o.Caption = Get_Language_Str(o.Caption, False, xlWhole)
        End If
      On Error GoTo 0
   Next o
End Sub


'UT---------------------------------
Sub Test_Change_Language_in_Dialog()
'UT---------------------------------
  #If PATTERN_CONFIG_PROG Then
     Copy_Select_GotoAct_Form.Show
  #End If
End Sub

'-----------------------------------------------------------
Public Sub Set_Tast_Txt_Var(Optional ForceUpdate As Boolean)                ' 06.03.20:
'-----------------------------------------------------------
  If ForceUpdate Or Red_T = "" Then
     Red_T = Get_Language_Str("Rot")
     Green_T = Get_Language_Str("Grün")
     OnOff_T = Get_Language_Str("AnAus")
     Tast_T = Get_Language_Str("Tast")
  End If
End Sub




'UT----------------------------
Private Sub Test_Translations()
'UT----------------------------
' Check it the translation works correct
  Check_Languages = True
  Debug.Print vbCr & "-------------------------------------------"
  Dim Res As String
  Res = InputBox("Input the Language number" & vbCr & _
                           " 0 = German" & vbCr & _
                           " 1 = English" & vbCr & _
                           " 2 = Dutch" & vbCr & _
                           " 3 = French" & vbCr & _
                           " 4 = Italian" & vbCr & _
                           " 5 = Spain" & vbCr & _
                           " 6 = Danish")
                        
  If Not IsNumeric(Res) Then Exit Sub
  Test_Language = val(Res)
                           
  Update_Language_in_All_Sheets ' Is also used in the Prog_Generator

#If PATTERN_CONFIG_PROG Then
  Change_Language_in_Dialog Copy_Select_GotoAct_Form
  Change_Language_in_Dialog MainMenu_Form
  Change_Language_in_Dialog Percent_Msg_UserForm
  Change_Language_in_Dialog Select_from_Sheet_Form
  Change_Language_in_Dialog Select_GotoNr_Form
  Change_Language_in_Dialog Select_LED_Address
  Change_Language_in_Dialog StatusMsg_UserForm
  Change_Language_in_Dialog Test_GotoNr_Form
  Change_Language_in_Dialog UserForm_Other
  Change_Language_in_Dialog UserForm_USB_Connection
#End If

#If PROG_GENERATOR_PROG Then
  Change_Language_in_Dialog Select_ProgGen_Dest_Form
  Change_Language_in_Dialog Select_ProgGen_Src_Form
  Change_Language_in_Dialog SelectMacros_Form
  Change_Language_in_Dialog SelectMacrosTreeForm                            ' 07.10.21:
  Change_Language_in_Dialog StatusMsg_UserForm
  Change_Language_in_Dialog UserForm_Connector
  Change_Language_in_Dialog UserForm_Description
  Change_Language_in_Dialog UserForm_DialogGuide1
  Change_Language_in_Dialog UserForm_Header_Created
  Change_Language_in_Dialog UserForm_House
  Change_Language_in_Dialog UserForm_Options
  Change_Language_in_Dialog UserForm_Other
  Change_Language_in_Dialog UserForm_Protokoll_Auswahl
  Change_Language_in_Dialog UserForm_Select_Typ_DCC
  Change_Language_in_Dialog UserForm_Select_Typ_SX
  Change_Language_in_Dialog Wait_CheckColors_Form
  Change_Language_in_Dialog Import_Hide_Unhide
  Change_Language_in_Dialog Select_COM_Port_UserForm
#End If ' PROG_GENERATOR_PROG

  Set_Tast_Txt_Var ForceUpdate:=True    ' 06.03.20:

  'UserForm_DialogGuide1.Show ' Debug
  
End Sub





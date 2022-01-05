VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectMacros_Form 
   Caption         =   "Auswahl des Makros"
   ClientHeight    =   10236
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12765
   OleObjectBlob   =   "SelectMacros_Form.frx":0000
End
Attribute VB_Name = "SelectMacros_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' SelectMacros_Form
' ~~~~~~~~~~~~~~~~~

' Module Description:
' ~~~~~~~~~~~~~~~~~~~
' This module contains functions for the SelectMacros dialog.
' The user could select a Macro with the mouse.
' The result is stored in the public variable SelectMacro_Res

' Revision History:
' ~~~~~~~~~~~~~~~~~
' 19.08.19: - Started

' ToDo:
' - Bei einem Doppelklick in eine bestehede Zeile soll der Cursor auf der
'   entsprechenen Zeile innerhalb der ListBox stehen. Das Funktioniert aber
'   nicht immer.
'   Wenn der Mauszeiger innerhalb der Auswahlliste ist, dann wird anschließend
'   der entsprechende Eintrag gewählt, auch wenn vorher schon der alte Eintrag
'   ausgewählt war.
'   Ich habe die verschiedensten Dinge probiert zum umgehen des Problems aber
'   noch keine Lösung gefunden.
'   Das versetzen des Mauscursors außerhalb der Liste funktioniert, ist aber
'   nicht schön weil man die Maus dann suchen muss.
'   Siehe auch SelectMacros()
'   => Erst mal bleibt es so wie es ist. Wenn der Dialog an einer anderen
'      Stelle erscheint, dann funktioniert es ja.


Option Explicit

#Const AUTO_ACT_FILTER_BY_LETTERS = 1   ' Activate the filter automatically if a letter is typed


Private Enable_Listbox_Changed As Boolean
Private SrcSh As Worksheet

Private ListDataSh As String
Private ListFilter As String

Private Load_Data_to_Listbox_Activ As Boolean

Private SelectByKey As Long                                                 ' 20.04.20:



'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  UnhookFormScroll ' Deactivate the mouse wheel scroll function
  Me.Hide  ' no "Unload Me" to keep the entered data and dialog position
  SelectMacro_Res = ""
  Enable_Listbox_Changed = False
End Sub


'---------------------------------
Private Sub Calc_SelectMacro_Res()
'---------------------------------
' Return the name and the row number in the ListDataSheet
' If MultiSelect is enabled a space separated list is returned
  Dim Nr As Long, Res As String
  For Nr = 0 To ListBox.ListCount - 1
      If ListBox.Selected(Nr) Then
         Res = Res & ListBox.List(Nr, 0) & "," & ListBox.List(Nr, 2) & " " ' Return "Name,Nr"  Nr: row nr in the data sheet
         Last_SelectedNr_Valid = True
         Last_SelectedNr = Nr
         Exit For
      End If
  Next Nr
  SelectMacro_Res = DelLast(Res) ' Delete the tailing space
End Sub

'--------------------------------
Private Sub Select_Button_Click()
'--------------------------------
  UnhookFormScroll ' Deactivate the mouse weel scrol function
  Me.Hide ' no "Unload Me" to keep the entered data.
  Enable_Listbox_Changed = False
  Calc_SelectMacro_Res
End Sub

#If False Then ' nicht mehr gebraucht, der Row Index steht nun in Spalte 2 des jeweiligen Listeneintrags
'--------------------------------------------------------
Private Function Get_Row_with_Mode_Filter(ListNr As Long)
'--------------------------------------------------------
  Dim Nr As Long, r As Long
  r = SM_DIALOGDATA_ROW1
  Nr = -1
  With Sheets(ListDataSh)
    While True
      If .Cells(r, SM_Mode__COL) = "" Or Expert_CheckBox Then
        If TextBoxFilter.Text = "" Then
          Nr = Nr + 1
        Else
            If InStr(1, .Cells(r, SM_Name__COL), TextBoxFilter.Text, VbCompareMethod.vbTextCompare) > 0 Then Nr = Nr + 1
        End If
      End If
      If Nr < ListNr Then
          r = r + 1
      Else
          Get_Row_with_Mode_Filter = r
         Exit Function
      End If
    Wend
  End With
End Function
#End If

#If 1 Then                                                                  ' 20.04.20:
' ListBox:
' Beim drücken einer Taste wird standardmäßig der erste Eintrag angesprungen der den
' entsprechednen Buchstaben enthält.
' Es ist aber nicht möglich zum nächsten Eintrag zu springen indem man die Taste noch mal drückt.
'
' Dummerweise wird die ListBox_KeyPress Funktion vor der normalen Funktion zum anspringen einer Zeile
' Aufgerufen => Wenn man in der ListBox_KeyPress() Funktion die Zeile ändert, dann springt er anschließend
' wieder zu "Seiner" per Taste ausgewählten Zeile
' Darum habe ich die Variable "SelectByKey" eingefügt. Diese wird in ListBox_Change() überprüft und verbiegt die
' Zielzeile wieder.
' Das Funktioniert beim ersten mal. Wenn dann der Cursor an einer Stelle steht welche zu der Taste passt, dann
' wird ListBox_Change() nicht mehr gestartet. Darum wird
'   ListBox.Selected(SelectByKey) = True
' vorher in ListBox_KeyPress aufgerufen

'------------------------------------------
Private Function Get_Selected_Row() As Long
'------------------------------------------
  Dim Nr As Long, Res As String
  For Nr = 0 To ListBox.ListCount - 1
      If ListBox.Selected(Nr) Then
         Get_Selected_Row = Nr
         Exit Function
      End If
  Next Nr
  Get_Selected_Row = -1
End Function

'-------------------------------------------------------------------------------------------------------
Private Function Find_Line_with(ByVal c As String, StartLine As Long, Optional Mode As Long = 1) As Long
'-------------------------------------------------------------------------------------------------------
' Mode: 1 Check first Char
'      -1 Check any Char
  c = UCase(c)
  Dim LineNr As Long
  With ListBox
    For LineNr = StartLine To .ListCount - 1
       Dim Line As String
       Line = .List(LineNr)
       If Line <> "" Then
          Select Case Mode
            Case 1: ' First character
                    If UCase(Left(Line, 1)) = c Then
                       Find_Line_with = LineNr
                       Exit Function
                    End If
            Case -1: ' Find anywhere in the line
                    If InStr(Line, c) > 0 Then
                       Find_Line_with = LineNr
                       Exit Function
                    End If
                    
          End Select
       End If
    Next LineNr
  End With
  
  ' Not found
  If StartLine <> 0 Then
        Find_Line_with = Find_Line_with(c, 0, Mode) ' Try again starting with 0
  Else: Find_Line_with = -1
  End If
End Function
#If 0 Then
'------------------------------------------------------------------------------------------------------------------
Private Function Send_Letters_to_TextBoxFilter(KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) As Boolean
'------------------------------------------------------------------------------------------------------------------
    Select Case KeyCode
        Case Asc("A") To Asc("Z"):
             If Shift < 2 Then
                TextBoxFilter.setFocus
                If Shift = 0 Then KeyCode = KeyCode + 32 ' Convert to lower case
                TextBoxFilter = TextBoxFilter & Chr(KeyCode)
             ElseIf Shift = 4 And KeyCode = Asc("F") Then ' Alt+F => Filter (Attention: The accelerator character can't be displayed in the text because then the excel Toolbox dialog is opened for some reasons)
                TextBoxFilter.setFocus
             End If
             Send_Letters_to_TextBoxFilter = True
       'Case Else: Debug.Print "KeyCode: " & KeyCode.Value & " chr: " & Chr(KeyCode.Value) & " Shift:" & Shift
    End Select
End Function

' Keyboard event functions to send key press to the TextBoxFilter
'Private Sub Expert_CheckBox_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer):  Send_Letters_to_TextBoxFilter KeyAscii, Shift: End Sub

'Private Sub Abort_Button_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer):     Send_Letters_to_TextBoxFilter KeyAscii, Shift: End Sub

'Private Sub Select_Button_KeyUp(ByVal KeyAscii As MSForms.ReturnInteger, ByVal Shift As Integer):    Send_Letters_to_TextBoxFilter KeyAscii, Shift: End Sub
#End If

'--------------------------------------------------------------------
Private Sub ListBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'--------------------------------------------------------------------
    'Debug.Print "ListBox_KeyPress:" & KeyAscii
    If Chr(KeyAscii) Like "[a-zA-Z0-9]" Then
       Dim ActLine As Long
       ActLine = Get_Selected_Row()
       Dim LineNr As Long, Mode As Long
       If Chr(KeyAscii) Like "[0-9]" Then
             Mode = -1  ' Find only in the first position
       Else: Mode = 1   ' Find anywhere in the line
       End If
       
       #If AUTO_ACT_FILTER_BY_LETTERS Then
           TextBoxFilter = TextBoxFilter & Chr(KeyAscii)
           TextBoxFilter.setFocus
       #End If
       
       LineNr = Find_Line_with(Chr(KeyAscii), ActLine + 1, Mode)
       'Debug.Print "ActLine:" & ActLine & " Found Line:" & LineNr
       If LineNr >= 0 Then
          ActLine = LineNr
          SelectByKey = -1
          ListBox.Selected(LineNr) = True   ' This will trigger the ListBox_Change() function
          SelectByKey = LineNr
       End If
    End If
End Sub

'---------------------------------------------------------------------------------
Private Sub CheckKey_and_Activate_Listbox(ByVal KeyAscii As MSForms.ReturnInteger)
'---------------------------------------------------------------------------------
  If Chr(KeyAscii) Like "[a-zA-Z0-9]" Then
     ListBox.setFocus
     ListBox_KeyPress KeyAscii
  End If

End Sub

'----------------------------------------------------------------------------
Private Sub Expert_CheckBox_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'----------------------------------------------------------------------------
  CheckKey_and_Activate_Listbox KeyAscii
End Sub

'-------------------------------------------------------------------------
Private Sub Abort_Button_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'-------------------------------------------------------------------------
  CheckKey_and_Activate_Listbox KeyAscii
End Sub

'--------------------------------------------------------------------------
Private Sub Select_Button_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'--------------------------------------------------------------------------
  CheckKey_and_Activate_Listbox KeyAscii
End Sub

#End If

'---------------------------
Private Sub ListBox_Change()
'---------------------------
  If SelectByKey >= 0 Then                                ' 20.04.20:
     Dim OldEvents As Boolean
     OldEvents = Application.EnableEvents
     Application.EnableEvents = False
     'Debug.Print "ListBox_Change: SelectByKey=" & SelectByKey
     On Error Resume Next                 ' Solve problem if Alt+E ist pressed
     ListBox.Selected(SelectByKey) = True
     On Error GoTo 0
     SelectByKey = -1
     Application.EnableEvents = OldEvents
  End If
  
  Description = ""
  Detail = ""
  If Enable_Listbox_Changed Then
     Dim Nr As Long, Cnt As Long
     For Nr = 0 To ListBox.ListCount - 1
         If ListBox.Selected(Nr) Then
            Cnt = Cnt + 1
            Dim Row As Long
            Row = ListBox.List(Nr, 2)
            Dim Txt As String, ActLanguage As Integer
            ActLanguage = Get_ExcelLanguage()                                                 ' 24.02.20:
            Txt = Replace(SrcSh.Cells(Row, SM_DetailCOL + ActLanguage * DeltaCol_Lib_Macro_Lang), "|", vbLf) ' 24.02.20: added:  + ActLanguage * DeltaCol_Lib_Macro_Lang
            If Txt <> "" Then
                  Description = Description & Txt & vbCr
            Else: Description = Description & SrcSh.Cells(Row, SM_ShrtD_COL + ActLanguage * DeltaCol_Lib_Macro_Lang) & vbCr    ' 24.02.20: added:  + ActLanguage * DeltaCol_Lib_Macro_Lang
            End If
            Detail = Detail & Replace_Multi_Space(SrcSh.Cells(Row, SM_Macro_COL)) & vbCr ' Show one or more selected details
         End If
     Next
     SelectedCnt_Label = "Selected: " & Cnt
  End If
  
End Sub

'------------------------------------------------------------------
Private Sub ListBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'------------------------------------------------------------------
  Select_Button_Click
End Sub

'------------------------------------------------------
Private Function Add_Filtered(Filter As String) As Long
'------------------------------------------------------
' Add elements to the Listbox and return the number of the listbox entry which matches with the given string
  Dim r As Range, c As Range, ActLanguage As Integer
  ActLanguage = Get_ExcelLanguage()                                         ' 24.02.20:
  'Debug.Print "Add_Filtered(" & Filter & "," & TextBoxFilter.Text & ")" ' Debug
  Add_Filtered = -1 ' Nothing found
  Dim Sh As Worksheet
  Set Sh = ThisWorkbook.Sheets(ListDataSh)
  With Sh
     
     Set r = .Range(.Cells(SM_DIALOGDATA_ROW1, SM_Name__COL), .Cells(LastFilledRowIn(Sh, SM_Name__COL), SM_Name__COL)) ' 19.10.21: Old: LastUsedRowIn(sh)
     For Each c In r
       If .Cells(c.Row, SM_Mode__COL) = "" Or Expert_CheckBox Then
          If TextBoxFilter.Text = "" Or _
          InStr(1, c.Value, TextBoxFilter.Text, VbCompareMethod.vbTextCompare) > 0 Or _
          InStr(1, .Cells(c.Row, SM_ShrtD_COL + ActLanguage * DeltaCol_Lib_Macro_Lang).Value, TextBoxFilter.Text, VbCompareMethod.vbTextCompare) > 0 Then            ' 12.02.21: Juergen Filter feature
            ListBox.AddItem c.Value
            ListBox.List(ListBox.ListCount - 1, 1) = .Cells(c.Row, SM_ShrtD_COL + ActLanguage * DeltaCol_Lib_Macro_Lang).Value  ' 24.02.20: added:  + ActLanguage * DeltaCol_Lib_Macro_Lang
            ListBox.List(ListBox.ListCount - 1, 2) = c.Row                       ' 12.02.21: Juergen Filter feature
          End If
       End If
       ' Find the actual line and select it
       If Filter <> "" And .Cells(c.Row, SM_FindN_COL) <> "" Then
            If InStr(Filter, .Cells(c.Row, SM_FindN_COL)) <> 0 Then
               If .Cells(c.Row, SM_Mode__COL) <> "" And Not Expert_CheckBox Then ' Found in Expert mode line, but Expert mode is not active
                    ' Restart again with enabled expert mode
                    ListBox.Clear
                    Load_Data_to_Listbox_Activ = True                       ' 12.04.20: Prevent problem when a function is selected which is in expert mode. Without this the dialog to show the arguments has not been shown the first time
                    Expert_CheckBox = True
                    Load_Data_to_Listbox_Activ = False                      ' 12.04.20:
                    Add_Filtered = Add_Filtered(Filter)
                    Exit Function
               Else
                    Add_Filtered = ListBox.ListCount - 1 ' Last added element
                    'Debug.Print "Detected '" & .Cells(C.Row, SM_FindN_COL) & "' in " & Filter & " at pos " & ListBox.ListCount - 1  ' Debug
               End If
            End If
       End If
     Next c
  End With
End Function

'----------------------------------
Private Sub Expert_CheckBox_Click()
'----------------------------------
  If Not Load_Data_to_Listbox_Activ Then                                    ' 12.04.20:
     Load_Data_to_Listbox ListDataSh, ""
     Set_String_Config_Var "Expert_Mode_aktivate", IIf(Expert_CheckBox, "1", "0")  ' 06.10.21:
  End If
End Sub


'-------------------------------------------------------------------------------------
Private Function Load_Data_to_Listbox(ListDataSh_par As String, Filter As String) As Long
'-------------------------------------------------------------------------------------
' Defines the sheet which contains the data and loads the data into the form.
  ListDataSh = ListDataSh_par
  ListFilter = Filter
  On Error Resume Next ' For some reasons this function is called two times and the second call generates a crash
  Set SrcSh = ThisWorkbook.Sheets(ListDataSh)
  Me.ListBox.Clear
  On Error GoTo 0
  With Me.ListBox
      .ColumnCount = 2
      .ColumnWidths = "130 Pt"
      Load_Data_to_Listbox = Add_Filtered(ListFilter)                    ' 12.02.21: Juergen Filter feature
  End With
End Function

'---------------------------------
Private Sub TextBoxFilter_Change()          ' 12.02.21: Juergen Filter feature
'---------------------------------
    ListBox.Clear
    Add_Filtered ListFilter
End Sub


'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & me.Name & ": UserForm_Initialize"
  'Center_Form Me
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  
  Load_Data_to_Listbox_Activ = True                                         ' 06.10.21:
  Expert_CheckBox = Get_Bool_Config_Var("Expert_Mode_aktivate")
  Load_Data_to_Listbox_Activ = False
  
  
  With Me
        .StartUpPosition = 0
        .Left = Application.Left ' Left side to avoid selecting the wrong entry because the mouse lands in the listbox
        .Top = Application.Top '  + (Application.Height - .Height) / 2
  End With
  
End Sub

'-----------------------------------------------
Public Sub MouseWheel(ByVal lngRotation As Long)
'-----------------------------------------------
' Process the mouse wheel changes
  With ListBox     ' Adapt to the listbox which should be controlled
    If lngRotation > 0 Then
        If .TopIndex > 0 Then
            If .TopIndex > 3 Then
                .TopIndex = .TopIndex - 3
            Else
                .TopIndex = 0
            End If
        End If
    Else
        .TopIndex = .TopIndex + 3
    End If
  End With
End Sub

'------------------------------------------------------------------------------
Public Sub Show_SelectMacros_Form(ListDataSh As String, ByVal Filter As String)
'------------------------------------------------------------------------------
' Use this function to show the dialog.
  Dim SelectedNr As Long
  SelectByKey = -1
  SelectedNr = Load_Data_to_Listbox(ListDataSh, Filter)
  Enable_Listbox_Changed = True
  If SelectedNr < 0 And Last_SelectedNr_Valid Then
        SelectedNr = Last_SelectedNr
  End If
  
  If SelectedNr >= 0 Then
        SelectedCnt_Label = "Selected: 1"
  Else: SelectedCnt_Label = "Selected: 0"
  End If
  
  With ListBox
    If SelectedNr >= 0 And SelectedNr < ListBox.ListCount Then  ' 14.06.20: Added "and SelectedNr < ListBox.ListCount" to prevent crash
          .Selected(SelectedNr) = True
    End If
    .setFocus
  End With
  
  HookFormScroll Me, "ListBox"   ' Initialize the mouse wheel scroll function
  Me.ListBox.setFocus            ' 12.02.21 Juergen Filter feature
  Me.TextBoxFilter.Text = ""     ' 12.02.21 Juergen Filter feature
  Me.Show
End Sub


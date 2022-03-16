Attribute VB_Name = "M28_Diverse"
Option Explicit
  
'--------------------------------------------------------------
Public Function Is_Data_Sheet(ByVal Sh As Worksheet) As Boolean
'--------------------------------------------------------------
  Dim PageId As String
  PageId = Sh.Cells(SH_VARS_ROW, PAGE_ID_COL)
  If PageId = "" Then Exit Function                                         ' 07.10.21:
  Is_Data_Sheet = (InStr(AllData_PgIDs, " " & PageId & " ") > 0) ' 17.10.20: removed: And Sh.Name <> "Examples"  ' 07.08.20: Added: And Sh.Name <> "Examples"
  'Debug.Print "Is_Data_Sheet(" & sh.Name & ")=" & Is_Data_Sheet ' Debug
End Function

'UT-----------------------------
Private Sub Test_Is_Data_Sheet()
'UT-----------------------------
  Debug.Print Is_Data_Sheet(Sheets("Start"))
End Sub


'----------------------------
Public Sub EnableAllButtons()
'----------------------------
' Enable all buttons in case they have been disabled by a crash
  Dim Sh As Variant
  For Each Sh In ThisWorkbook.Sheets
      If Is_Data_Sheet(Sh) Then
        Sh.EnableDisableAllButtons True
      End If
  Next
End Sub


'-------------------------------------------------------------------
Private Sub Clear_COM_Port_Check(r As Range, ReleaseMode As Boolean)
'-------------------------------------------------------------------
' Set to a negativ number.
  With r
    If Not ReleaseMode And IsNumeric(.Value) Then
          .Value = -Abs(.Value)
    Else: .Value = "COM?"
    End If
  End With
End Sub

' 24.12.19:
' Bei Armin stürtzt das Programm beim starten in den beiden folgenden Zeilen ab:
'   sh.Cells(FirstDat_Row, Descrip_Col).Select
'   .Cells(LRow + 1, Descrip_Col).Select
' Siehe Mail vom 23.12.19.
' Bei ihm hat die
'   sh.Select
' Zeile gefehlt.
' Als Work arround habe ich die "On Error.. " Zeilen eingebaut. Damut läuft es.

'-----------------------------------------------------------------------------------
Public Sub Clear_COM_Port_Check_and_Set_Cursor_in_all_Sheets(ReleaseMode As Boolean)    ' 25.12.19: Old: Clear_COM_Port_Check_ans_Set_Cursor_in_all_Sheets
'-----------------------------------------------------------------------------------
  Dim Sh As Variant, OldSh As Worksheet, Skip_Scroll_Down As Boolean
  Set OldSh = ActiveSheet
  If ActiveSheet Is Nothing Then                                            ' 29.10.19:
     Debug.Print "ActiveSheet Is Nothing in Clear_COM_Port_Check_and_Set_Cursor_in_all_Sheets"
     Debug.Print "Tritt beim ersten Start nach dem Download vom Internet auf ('Geschützte Ansicht')"
     Skip_Scroll_Down = True
  End If
  For Each Sh In ThisWorkbook.Sheets
      If Is_Data_Sheet(Sh) Then
        With Sh
          Make_sure_that_Col_Variables_match Sh
          Clear_COM_Port_Check .Cells(SH_VARS_ROW, COMPort_COL), ReleaseMode
          
          If Sh.Cells(SH_VARS_ROW, PAGE_ID_COL) <> "CAN" Then
            Clear_COM_Port_Check .Cells(SH_VARS_ROW, COMPrtR_COL), ReleaseMode
            If ReleaseMode Then .Cells(SH_VARS_ROW, R_UPLOD_COL) = "R not Chk" ' Right arduino software is not checked
          End If
          If Not Skip_Scroll_Down Then                                      ' 29.10.19:
             Dim LRow As Long
             On Error Resume Next                                           ' 24.12.19: Problems with Office 365 ?
             Sh.Select                                                      ' 29.10.19:
             Sh.Cells(FirstDat_Row, Descrip_Col).Select ' Scroll to the top
             LRow = LastFilledRowIn_ChkAll(Sh) + 1
             .Cells(LRow + 1, Descrip_Col).Select     ' Select the first empty row
             'While .Rows(LRow).EntireRow.Hidden                            ' 29.10.19: Disabled
             '   LRow = LRow + 1
             'Wend
             On Error GoTo 0                                                ' 24.12.19:
          End If
        Dim rngCell As Range, i As Long                                     ' 12.10.21: Jürgen Clear errors
        For Each rngCell In Sh.UsedRange
          For i = 1 To 7
              rngCell.Errors.Item(i).Ignore = True
          Next i
        Next
        End With
      End If
  Next
  If Not OldSh Is Nothing Then OldSh.Select
End Sub

'UT-----------------------------------------------------------------
Private Sub Test_Clear_COM_Port_Check_and_Set_Cursor_in_all_Sheets()
'UT-----------------------------------------------------------------
  Clear_COM_Port_Check_and_Set_Cursor_in_all_Sheets True
End Sub

'-------------------------------------------------------------
Public Function Get_Bool_Config_Var(Name As String) As Boolean
'-------------------------------------------------------------
  On Error GoTo NotFound
  With ThisWorkbook.Sheets(ConfigSheet).Range(Name)
    Select Case UCase(left(Trim(.Value), 1))                     ' Languages (DE,  EN, NL,  FR,   IT, ES)
       Case "", "N", "G", "A", "0":  Get_Bool_Config_Var = False '            Nein No  geen aucun no  no          ' 14.05.20: Added "0"
       Case Else:     Get_Bool_Config_Var = True                 '            Ja   Yes ja   oui   sì sì
    End Select
  End With
  Exit Function
  
NotFound:
  MsgBox "Interner Fehler: Die Konfigurationsvariable '" & Name & "' wurde nicht im Sheet '" & ConfigSheet & "' gefunden", _
         vbCritical, "Interner Fehler in Get_Bool_Config_Var"
  EndProg
End Function

'------------------------------------------
Public Function Get_Num_Config_Var(Name As String) As Long
'------------------------------------------
  On Error GoTo NotFound
  Dim Str As String
  Str = ThisWorkbook.Sheets(ConfigSheet).Range(Name)
  If IsNumeric(Str) Then
        Get_Num_Config_Var = val(Str)
  Else: Get_Num_Config_Var = -1
  End If
  Exit Function
NotFound:
  MsgBox "Interner Fehler: Die Konfigurationsvariable '" & Name & "' wurde nicht im Sheet '" & ConfigSheet & "' gefunden", _
         vbCritical, "Interner Fehler in Get_Num_Config_Var"
  EndProg
End Function

'------------------------------------------   04.03.22 Juergen
Public Function Get_Num_Config_Var_Range(Name As String, Min As Long, Max As Long, Optional Default As Long = 0) As Long
'------------------------------------------
  On Error GoTo NotFound
  Dim Str As String
  Str = ThisWorkbook.Sheets(ConfigSheet).Range(Name)
  If IsNumeric(Str) Then
        Get_Num_Config_Var_Range = val(Str)
  Else: Get_Num_Config_Var_Range = Default
  End If
  If Get_Num_Config_Var_Range < Min Then Get_Num_Config_Var_Range = Min
  If Get_Num_Config_Var_Range > Max Then Get_Num_Config_Var_Range = Max
  Exit Function
NotFound:
  MsgBox "Interner Fehler: Die Konfigurationsvariable '" & Name & "' wurde nicht im Sheet '" & ConfigSheet & "' gefunden", _
         vbCritical, "Interner Fehler in Get_Num_Config_Var"
  EndProg
End Function

'UT----------------------------------
Private Sub TestGet_Bool_Config_Var()
'UT----------------------------------
  Debug.Print "Get_Bool_Config_Var=" & Get_Bool_Config_Var("Lib_Installed_other")
End Sub

'-------------------------------------------------------------
Public Sub Set_Bool_Config_Var(Name As String, val As Boolean)
'-------------------------------------------------------------
  On Error GoTo NotFound
  With ThisWorkbook.Sheets(ConfigSheet).Range(Name)
    If val Then
          .Value = Get_Language_Str("Ja")
    Else: .Value = Get_Language_Str("Nein")
    End If
  End With
  Exit Sub
  
NotFound:
  MsgBox "Interner Fehler: Die Konfigurationsvariable '" & Name & "' wurde nicht im Sheet '" & ConfigSheet & "' gefunden", _
         vbCritical, "Interner Fehler in Set_Bool_Config_Var"
  EndProg
End Sub

'--------------------------------------------------------------
Public Function Get_String_Config_Var(Name As String) As String
'--------------------------------------------------------------
  On Error GoTo NotFound
  Get_String_Config_Var = ThisWorkbook.Sheets(ConfigSheet).Range(Name)
  Exit Function
  
NotFound:
  MsgBox "Interner Fehler: Die Konfigurationsvariable '" & Name & "' wurde nicht im Sheet '" & ConfigSheet & "' gefunden", _
         vbCritical, "Interner Fehler in Get_String_Config_Var"
  EndProg
End Function

'--------------------------------------------------------------
Public Sub Set_String_Config_Var(Name As String, val As String)
'--------------------------------------------------------------
  On Error GoTo NotFound
  ThisWorkbook.Sheets(ConfigSheet).Range(Name) = val
  Exit Sub
  
NotFound:
  MsgBox "Interner Fehler: Die Konfigurationsvariable '" & Name & "' wurde nicht im Sheet '" & ConfigSheet & "' gefunden", _
         vbCritical, "Interner Fehler in Set_String_Config_Var"
  EndProg
End Sub

'--------------------------------------------------------------
Public Function Get_Old_Board(LeftArduino As Boolean) As String ' 04.05.20: Extracted from Get_Arduino_Typ()
'--------------------------------------------------------------
  Dim Col As Integer
  If LeftArduino Then
        Col = BUILDOP_COL
  Else: Col = BUILDOpRCOL
  End If
  Dim BuildOpt As String
  BuildOpt = Cells(SH_VARS_ROW, Col)
  If InStr(BuildOpt, BOARD_NANO_OLD) > 0 Then
                                                   Get_Old_Board = BOARD_NANO_OLD
  ElseIf InStr(BuildOpt, BOARD_UNO_NORM) > 0 Then
                                                   Get_Old_Board = BOARD_UNO_NORM
  ElseIf InStr(BuildOpt, BOARD_NANO_EVERY) > 0 Then
                                                   Get_Old_Board = BOARD_NANO_EVERY   ' 28.10.20: Jürgen
  ElseIf InStr(BuildOpt, BOARD_NANO_FULL) > 0 Then
                                                   Get_Old_Board = BOARD_NANO_FULL    ' 28.10.20: Jürgen
  ElseIf InStr(BuildOpt, BOARD_NANO_NEW) > 0 Then
                                                   Get_Old_Board = BOARD_NANO_NEW
  End If
End Function


'--------------------------------------------------------------------
Public Sub Change_Board_Typ(LeftArduino As Boolean, NewBrd As String)
'--------------------------------------------------------------------
'  If Disable_Set_Arduino_Typ Then Exit Sub
  
  Dim Col As Long, Brd As Integer
  If LeftArduino Then
        Col = BUILDOP_COL: Brd = 0
  Else: Col = BUILDOpRCOL: Brd = 1
  End If
  Dim BuildOpt As String, Old_Board As String
  Old_Board = Get_Old_Board(LeftArduino)
  BuildOpt = Cells(SH_VARS_ROW, Col)
  If Old_Board = "" Then
        BuildOpt = NewBrd  ' & " " & BuildOpt               28.10.20: Jürgen: Disabled "& " " & BuildOpt"
  Else: BuildOpt = Replace(BuildOpt, Old_Board, NewBrd)
  End If
  Cells(SH_VARS_ROW, Col) = Trim(BuildOpt)
End Sub

Public Function Is_Named_Range(rng As Range) As Boolean
    On Error Resume Next
    Is_Named_Range = rng.Name.Name <> ""
End Function



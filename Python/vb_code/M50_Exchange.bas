Attribute VB_Name = "M50_Exchange"
Option Explicit

' Exchange functions for the Pattern_Configuarator                          ' 16.09.19:

Private Const Default_Pattern_Configurator_Name = "Pattern_Configurator.xlsm"


'----------------------------------------------
Public Function Get_Prog_Version_Nr() As String
'----------------------------------------------
' Return a number like: 1.0.8b
  Get_Prog_Version_Nr = Prog_Version_Nr
End Function

'----------------------------------------------
Public Function Selected_Row_Valid() As Boolean
'----------------------------------------------
  Make_sure_that_Col_Variables_match
  Selected_Row_Valid = (ActiveCell.Row >= FirstDat_Row)
End Function

'------------------------------------------------------------
Public Function Get_Description_Range_from_Act_Row() As Range
'------------------------------------------------------------
  Make_sure_that_Col_Variables_match
  Set Get_Description_Range_from_Act_Row = Cells(ActiveCell.Row, Descrip_Col)
End Function



'------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Write_Macro_to_Act_Row(MacroTxt As String, LEDs As String, InCnt As String, LocInCh As String, Optional Comment As String, Optional WrapText As Boolean)
'------------------------------------------------------------------------------------------------------------------------------------------------------------------
  Make_sure_that_Col_Variables_match
  Dim Row As Long
  Row = ActiveCell.Row
  If Cells(Row, LED_Cha_Col) = "" Then
     Dim Res As String, ResOk As Boolean
     While Not ResOk
       Res = InputBox(Get_Language_Str("Welcher LED Kanal soll verwendet werden?" & vbCr & _
                                       "  0 = Standard LEDs" & vbCr & _
                                       "  1 = Taster LEDs" & vbCr & _
                                       "  2 = Optionale LED Gruppe 2" & vbCr & _
                                       "  3 = Optionale LED Gruppe 2" & vbCr & _
                                       vbCr & _
                                       "LED Kanal (0..3):"), _
                      Get_Language_Str("Eingabe des LED Kanals"), 0)
       Res = Trim(Res)
       If Res = "" Then Exit Sub
       If IsNumeric(Res) And val(Res) >= 0 And val(Res) < LED_CHANNELS Then ResOk = True
     Wend
     Cells(Row, LED_Cha_Col) = val(Res)
  End If
  
  Dim LEDs_Channel As Long
  LEDs_Channel = val(Cells(Row, LED_Cha_Col))
  Cells(Row, Enable_Col) = ChrW(Hook_CHAR) ' Enable the Line
  Cells(Row, Config__Col) = MacroTxt
  Cells(Row, Config__Col).WrapText = WrapText
  Cells(Row, LEDs____Col) = LEDs
  Cells(Row, InCnt___Col) = InCnt
  Cells(Row, LocInCh_Col) = LocInCh
  If Comment <> "" Then Cells(Row, Descrip_Col) = Comment
End Sub

'---------------------------------------------------------
Private Function Get_WinStateName(State As Long) As String
'---------------------------------------------------------
  Select Case State
    Case xlMinimized: Get_WinStateName = "xlMinimized"
    Case xlMaximized: Get_WinStateName = "xlMaximized"
    Case xlNormal: Get_WinStateName = "xlNormal"
    Case Else: Get_WinStateName = "Unknown " & State
  End Select
End Function

'--------------------------------------------------------------------------
Public Sub NotMinimizedWindow(NewState As Long, Optional Fource As Boolean)
'--------------------------------------------------------------------------
  If Application.WindowState = xlMinimized Or Fource Then
        Debug.Print "Set Application.WindowState to" & Get_WinStateName(NewState)
        Application.WindowState = NewState
  Else: Debug.Print "Don't change Application.WindowState=" & Get_WinStateName(Application.WindowState)
  End If
End Sub


'-----------------------------------------------------------------------------------------------------
Public Sub Select_Line_for_Patern_Config_and_Call_Macro(Get_Dest As Boolean, Macro_Callback As String)
'-----------------------------------------------------------------------------------------------------
' Select the destination / Source row
    If Get_Dest Then
          Select_ProgGen_Dest_Form.Check_and_Start Macro_Callback
    Else: Select_ProgGen_Src_Form.Check_and_Start Macro_Callback
    End If
End Sub


'---------------------------------------------------------
Private Function Get_Pattern_Configurator_Name() As String                  ' 14.06.20:
'---------------------------------------------------------

  If Same_Name_already_open(Default_Pattern_Configurator_Name) Then
       Get_Pattern_Configurator_Name = Workbooks(Default_Pattern_Configurator_Name).FullName
  Else ' not opened
       Dim Path As String, FullPath As String
       Path = Get_DestDir_All()
       ' Check if it exists in the user dir
       FullPath = Path & Default_Pattern_Configurator_Name
       If Dir(FullPath) = "" Then
             ' Check if it exists in the lib dir
             Path = Get_SrcDirInLib()
             FullPath = Path & Default_Pattern_Configurator_Name
             If Dir(FullPath) = "" Then
                MsgBox Get_Language_Str("Fehler: Das Programm '") & Default_Pattern_Configurator_Name & "'" & vbCr & _
                       Get_Language_Str("existiert nicht im Standard Verzeichnis:") & vbCr & _
                       "  '" & Path & "'", vbCritical, Get_Language_Str("Fehler ") & Default_Pattern_Configurator_Name & Get_Language_Str(" nicht vorhanden")
                Exit Function
             End If
       End If
       Get_Pattern_Configurator_Name = FullPath
  End If
End Function


'--------------------------------------
Public Sub Start_Pattern_Configurator()
'--------------------------------------
' Make sure that the Pattern_Configurator excel sheet is opened
' and activated.
' If the program is already opened it's shown normal (Not minimized) and brought to the top
' If not it's opened from
'   1.: %USERPROFILE%\Documents\Arduino\MobaLedLib_ <Lib Ver>
'   2.: %USERPROFILE%\Documents\Arduino\libraries\MobaLedLib\extras\
  Dim Pattern_Configurator_Name As String
  Pattern_Configurator_Name = Get_Pattern_Configurator_Name
  If Pattern_Configurator_Name = "" Then Exit Sub
  
  If Same_Name_already_open(Pattern_Configurator_Name) Then
        Workbooks(FileNameExt(Pattern_Configurator_Name)).Activate
        Application.WindowState = xlNormal ' In case it was minimized. Unfortunately we don't know the state before
  Else: Workbooks.Open(Pattern_Configurator_Name).RunAutoMacros (xlAutoOpen) ' 14.06.20: Old: Workbooks.Open Filename:=Pattern_Configurator_Name
  End If

End Sub


'-------------------------------
Public Sub Copy_Pattern_Config()                                            ' 14.06.20:
'-------------------------------
'
  Dim Pattern_Configurator_Name As String
  Pattern_Configurator_Name = Get_Pattern_Configurator_Name
  ChDrive Pattern_Configurator_Name
  ChDir FilePath(Pattern_Configurator_Name)
  Run FileNameExt(Pattern_Configurator_Name) & "!Copy_Prog_If_in_LibDir"
End Sub



VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Options 
   Caption         =   "Optionen und Spezielle Funktionen"
   ClientHeight    =   5784
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   6885
   OleObjectBlob   =   "UserForm_Options.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Explicit


'Private Old_Board(2) As String
Private Disable_Set_Arduino_Typ As Boolean


'-------------------------------
Private Sub Close_Button_Click()
'-------------------------------
  Me.Hide
End Sub


'-----------------------------------
Private Sub ColorTest_Button_Click()
'-----------------------------------
  Me.Hide
  Open_MobaLedCheckColors ""
End Sub


'-----------------------------------------
Private Sub Detect_LED_Port_Button_Click()
'-----------------------------------------
  Me.Hide
  Detect_Com_Port_and_Save_Result False
  Me.Show
End Sub

'-------------------------------------------
Private Sub Detect_Right_Port_Button_Click()
'-------------------------------------------
  Me.Hide
  Detect_Com_Port_and_Save_Result True
  Me.Show
End Sub


'----------------------------------------
Private Sub FastBootloader_Button_Click()
'----------------------------------------
  Me.Hide
  Install_FastBootloader
End Sub

'------------------------------------
Private Sub HardiForum_Button_Click()
'------------------------------------
  Me.Hide
  If MsgBox(Get_Language_Str("Öffnet das Profil von Hardi im Stummi Forum." & vbCr & _
            vbCr & _
            "Dort findet man, wenn man im Forum angemeldet ist, einen Link zur " & _
            "Email Adresse des Autors zum senden einer E-Mail oder PN." & vbCr & _
            vbCr & _
            "Alternativ kann auch eine Mail an 'MobaLedlib@gmx.de' geschickt werden wenn " & _
            "es Fragen oder Anregungen zu dem Programm oder zur MobaLedLib gibt."), vbOKCancel, _
            Get_Language_Str("Profil des Autors öffnen")) = vbOK Then
     Shell "Explorer ""https://www.stummiforum.de/memberlist.php?mode=viewprofile&u=26419"""
  End If
End Sub

'----------------------------------------
Private Sub Pattern_Config_Button_Click()
'----------------------------------------
  Me.Hide
  Start_Pattern_Configurator
End Sub

'------------------------------------
Private Sub ProInstall_Button_Click()
'------------------------------------
  Ask_to_Upload_and_Compile_and_Upload_Prog_to_Right_Arduino
End Sub


Private Sub Nano_Normal_L_Click(): Change_Board True, BOARD_NANO_OLD:   End Sub
Private Sub Nano_New_L_Click():    Change_Board True, BOARD_NANO_NEW:   End Sub
Private Sub Nano_Full_L_Click():   Change_Board True, BOARD_NANO_FULL:  End Sub  ' 28.10.20: Jürgen


Private Sub Uno_L_Click():         Change_Board True, BOARD_UNO_NORM:   End Sub
Private Sub Board_IDE_L_Click():   Change_Board True, "":               End Sub
Private Sub ESP32_L_Click():       Change_Board True, BOARD_ESP32:      Autodetect_Typ_L_CheckBox.Value = False: End Sub  ' 11.11.20:
Private Sub Pico_L_Click():        Change_Board True, BOARD_PICO:       Autodetect_Typ_L_CheckBox.Value = False: End Sub  ' 18.04.21: Juergen

Private Sub Nano_Normal_R_Click(): Change_Board False, BOARD_NANO_OLD:  End Sub
Private Sub Nano_New_R_Click():    Change_Board False, BOARD_NANO_NEW:  End Sub
Private Sub Nano_Full_R_Click():   Change_Board False, BOARD_NANO_FULL: End Sub  ' 28.10.20: Jürgen
Private Sub Uno_R_Click():         Change_Board False, BOARD_UNO_NORM:  End Sub
Private Sub Board_IDE_R_Click():   Change_Board False, "":              End Sub

Private Sub Autodetect_Typ_L_CheckBox_Click(): Change_Autodetect True:  Check_Board (Autodetect_Typ_L_CheckBox): End Sub
Private Sub Autodetect_Typ_R_CheckBox_Click(): Change_Autodetect False: End Sub

'----------------------------------------------------------------
Public Sub Change_Board(LeftArduino As Boolean, NewBrd As String)
'----------------------------------------------------------------
  If Disable_Set_Arduino_Typ Then Exit Sub
  Change_Board_Typ LeftArduino, NewBrd
End Sub

'----------------------------------------------------
Private Sub Check_Board(AutodetectChecked As Boolean)
'----------------------------------------------------
    If AutodetectChecked Then
        If ESP32_L Or Pico_L Or Uno_L Or Board_IDE_L Then
            ' change back to Nano
            Nano_New_L = True
        End If
    End If
End Sub


'----------------------------------------------------
Private Sub Change_Autodetect(LeftArduino As Boolean)
'----------------------------------------------------
  Dim Side As String, Col As Long
  If LeftArduino Then
        Col = BUILDOP_COL: Side = "L"
  Else: Col = BUILDOpRCOL: Side = "R"
  End If
  Set_Autodetect_Value Col, Controls("Autodetect_Typ_" & Side & "_CheckBox")
End Sub

'--------------------------------------------------------------------------
Private Sub Set_Autodetect_Value(BuildOpt_Col As Long, Value As Boolean)
'--------------------------------------------------------------------------
  With Cells(SH_VARS_ROW, BuildOpt_Col)
       If Value Then
          If InStr(.Value, AUTODETECT_STR) = 0 Then .Value = AUTODETECT_STR & " " & Trim(.Value)
       Else
          .Value = Replace_Multi_Space(Trim(Replace(.Value, AUTODETECT_STR, "")))
       End If
  End With
End Sub


'---------------------------------------------------
Private Sub Get_Arduino_Typ(LeftArduino As Boolean)
'---------------------------------------------------
  Dim Side As String, Col As Integer
  If LeftArduino Then
        Col = BUILDOP_COL: Side = "L"
  Else: Col = BUILDOpRCOL: Side = "R"
  End If
  Dim BuildOpt As String
  BuildOpt = Cells(SH_VARS_ROW, Col)
  
  Controls("Autodetect_Typ_" & Side & "_CheckBox") = (InStr(BuildOpt, AUTODETECT_STR) > 0)
  
  If InStr(BuildOpt, BOARD_NANO_OLD) > 0 Then Controls("Nano_Normal_" & Side).Value = True: Exit Sub
  If InStr(BuildOpt, BOARD_NANO_FULL) > 0 Then Controls("Nano_Full_" & Side).Value = True:  Exit Sub  ' 28.10.20:
  If InStr(BuildOpt, BOARD_NANO_NEW) > 0 Then Controls("Nano_New_" & Side).Value = True:    Exit Sub
  If InStr(BuildOpt, BOARD_UNO_NORM) > 0 Then Controls("Uno_" & Side).Value = True:         Exit Sub
  If InStr(BuildOpt, BOARD_NANO_EVERY) > 0 Then Exit Sub        ' currently no option in GUI, but that's ok, as ATMEGA4809 is currently unsupported 28.10.20: Jürgen
  If InStr(BuildOpt, BOARD_ESP32) > 0 And Side = "L" And ESP32_Lib_Installed() Then Controls("ESP32_L").Value = True: Exit Sub                              ' 11.11.20:
  If InStr(BuildOpt, BOARD_PICO) > 0 And Side = "L" And PICO_Lib_Installed() Then Controls("PICO_L").Value = True: Exit Sub                                 ' 18.04.21: Juergen
  If InStr(BuildOpt, "--board ") > 0 Then
        Controls("Nano_Normal_" & Side).Value = False
        Controls("Nano_New_" & Side).Value = False
        Controls("Uno_" & Side).Value = False
        Controls("Board_IDE_" & Side).Value = False
        If Side = "L" Then Controls("ESP32_L").Value = False
        If Side = "L" Then Controls("PICO_L").Value = False
        MsgBox Get_Language_Str("Unbekannte Board Option: ") & vbCr & _
               BuildOpt, vbInformation, Get_Language_Str("Unbekanntes Board")
               
  End If
  Controls("Board_IDE_" & Side).Value = True ' Default value
End Sub


'UT-------------------------------
Private Sub Test_Get_Arduino_Typ()
'UT-------------------------------
  Get_Arduino_Typ True
End Sub


'--------------------------------
Private Sub Import_Button_Click()                                           ' 17.03.20:
'--------------------------------
  Me.Hide ' Must be hidden to be able to show the non modal dialog "Import_Hide_Unhide"
  Import_from_Old_Version
End Sub

'------------------------------
Private Sub Save_Button_Click()                                             ' 17.03.20:
'------------------------------
  Me.Hide ' Must be hidden to be able to show the non modal dialog "Import_Hide_Unhide"
  Save_Data_to_File
End Sub

'------------------------------
Private Sub Load_Button_Click()                                             ' 17.03.20:
'------------------------------
  Load_Data_from_File
End Sub

'-----------------------------------
Private Sub Copy_Page_Button_Click()
'-----------------------------------
  Me.Hide ' Must be hidden to be able to show the non modal dialog "Import_Hide_Unhide"
  Copy_from_Sheet_to_Sheet
End Sub


'-------------------------------------------
Private Sub Update_to_Arduino_Button_Click()
'-------------------------------------------
  Me.Hide ' Must be hidden to be able to show the non modal dialog "Import_Hide_Unhide"
  Update_MobaLedLib_from_Arduino_and_Restart_Excel
End Sub

'-------------------------------------
Private Sub Update_Beta_Button_Click()
'-------------------------------------
  Me.Hide ' Must be hidden to be able to show the non modal dialog "Import_Hide_Unhide"
  Update_MobaLedLib_from_Beta_and_Restart_Excel
End Sub


'-------------------------------------------------
Private Sub Show_Lib_and_Board_Page_Button_Click()
'-------------------------------------------------
  Me.Hide
  With ThisWorkbook.Sheets(LIBRARYS__SH)
     .Visible = True
     .Select
  End With
End Sub


'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
' Is called once to initialice the form
  'Debug.Print vbCr & Me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Center_Form Me
  Me.MultiPage1.Value = 0
End Sub


'------------------------------
Private Sub UserForm_Activate()
'------------------------------
' Is called every time when the form is shown
  Make_sure_that_Col_Variables_match
  ESP32_L.Visible = ESP32_Lib_Installed()                                   ' 11.11.20:
  Pico_L.Visible = PICO_Lib_Installed()                                    ' 18.04.21: Juergen

  'Debug.Print vbCr & Me.Name & ": UserForm_Activate"
  Me.MultiPage1.Page2.Visible = Page_ID <> "CAN" ' Not visible if the CAN Sheet is active
  Me.MultiPage1.Page2.Caption = Page_ID & " Arduino"
  
  Disable_Set_Arduino_Typ = True
  Get_Arduino_Typ True
  Get_Arduino_Typ False
  Disable_Set_Arduino_Typ = False
End Sub





VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_House 
   Caption         =   "House: Simulation eines ""belebten"" Hauses  in dem zufällig und abwechselnd  nur einige der Räume beleuchtet sind"
   ClientHeight    =   12024
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16620
   OleObjectBlob   =   "UserForm_House.frx":0000
End
Attribute VB_Name = "UserForm_House"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Mode As String
Private LED_CntList As String    ' String which contains the channel for each LED: "RGB RGB 1 2 3 1 "
Private MinCntChanged As Boolean
Private MaxCntChanged As Boolean

Private Const LED_ChannelsList = _
" ROOM_DARK:RGB         ROOM_BRIGHT:RGB       ROOM_WARM_W:RGB       ROOM_RED:RGB         " & _
" ROOM_D_RED:RGB        ROOM_COL0:RGB         ROOM_COL1:RGB         ROOM_COL2:RGB        " & _
" ROOM_COL3:RGB         ROOM_COL4:RGB         ROOM_COL5:RGB         ROOM_COL345:RGB      " & _
" FIRE:RGB              FIRED:RGB             FIREB:RGB             ROOM_CHIMNEY:RGB     " & _
" ROOM_CHIMNEYD:RGB     ROOM_CHIMNEYB:RGB     ROOM_TV0:RGB          ROOM_TV0_CHIMNEY:RGB " & _
" ROOM_TV0_CHIMNEYD:RGB ROOM_TV0_CHIMNEYB:RGB ROOM_TV1:RGB          ROOM_TV1_CHIMNEY:RGB " & _
" ROOM_TV1_CHIMNEYD:RGB ROOM_TV1_CHIMNEYB:RGB NEON_LIGHT:RGB        NEON_LIGHT1:1        " & _
" NEON_LIGHT2:2         NEON_LIGHT3:3         NEON_LIGHTD:RGB       NEON_LIGHT1D:1       " & _
" NEON_LIGHT2D:2        NEON_LIGHT3D:3        NEON_LIGHT3M:3        NEON_LIGHTM:RGB      " & _
" NEON_LIGHT1M:1        NEON_LIGHT2M:2        NEON_LIGHT3L:3        NEON_LIGHTL:RGB      " & _
" NEON_LIGHT1L:1        NEON_LIGHT2L:2        SINGLE_LED1:1         SINGLE_LED2:2        " & _
" NEON_DEF_D:RGB        NEON_DEF1D:1          NEON_DEF2D:2          NEON_DEF3D:3         " & _
" SINGLE_LED3:3         GAS_LIGHT3D:3         GAS_LIGHT1:1          GAS_LIGHT:RGB        " & _
" CANDLE:RGB            CANDLE1:1             CANDLE2:2             CANDLE3:3            " & _
" GAS_LIGHT2:2          GAS_LIGHT3:3          GAS_LIGHTD:RGB        GAS_LIGHT1D:1        " & _
" GAS_LIGHT2D:2         SKIP_ROOM:RGB         SKIP_ROOM:RGB         SKIP_ROOM:RGB        " & _
" SKIP_ROOM:RGB         SKIP_ROOM:RGB         SKIP_ROOM:RGB         SKIP_ROOM:RGB        " & _
" SKIP_ROOM:RGB         SKIP_ROOM:RGB         SKIP_ROOM:RGB         SKIP_ROOM:RGB        " & _
" SKIP_ROOM:RGB         SINGLE_LED1D:1        SINGLE_LED2D:2        SINGLE_LED3D:3       " & _
" SINGLE_LED3D:3        SINGLE_LED3D:3        SINGLE_LED3D:3"

' 12.010.20: Corrected the NEON_DEF2D entry. Channel 1 was used instead of channel 2. This caused the occupation
'            of a new RGB channel if NEON_DEF1D and NEON_DEF2D was used in a sequence

#If False Then
'-------------------------------------
Private Sub Debug_Print_LED_Channels()
'-------------------------------------
' Used to generate LED_ChannelsList (Prior the Channel have been stored in the "Tag" of the button, but this is no longer used)
  Dim o As Variant, All As String
  
  For Each o In Me.Controls
    Dim Txt As String
    If Left(o.Name, Len("CommandButton")) = "CommandButton" Then
       Txt = " " & o.Caption & ":"
       If o.Tag <> "" Then
             Txt = Txt & o.Tag
       Else: Txt = Txt & "RGB"
       End If
       Txt = Txt & "                "
    End If
    All = All & Left(Txt, 22)
    If Len(All) > 80 Then
       Debug.Print All & """ & _"
       All = """"
    End If
  Next o
  Debug.Print All & """ & _"
End Sub
#End If

'--------------------------------------------------------------------------------
Private Sub Set_Color_in_y_Range(StartY As Long, EndY As Long, ForeColor As Long)
'--------------------------------------------------------------------------------
  Dim c As Variant
  For Each c In Controls
     If c.Top >= StartY And c.Top < EndY Then
        c.ForeColor = ForeColor  ' &H80000012&
     End If
  Next c
End Sub

'---------------------------------------------------------------------------------
Private Sub Set_Bold_in_y_Range(StartY As Double, EndY As Double, Bold As Boolean) ' 11.10.20:   ' 02.01.21: Old: long
'---------------------------------------------------------------------------------
  Dim c As Variant
  For Each c In Controls
     If c.Top >= StartY And c.Top < EndY Then
        c.Font.Bold = Bold
     End If
  Next c
End Sub


Private Sub CommandButton58_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton59_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton60_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton61_Click()
  Add_Room Me.ActiveControl.Caption
End Sub



'-------------------------------------------
Private Sub IndividualTimes_CheckBox_Click()
'-------------------------------------------
  MinTime_TextBox.Enabled = IndividualTimes_CheckBox
  MaxTime_TextBox.Enabled = IndividualTimes_CheckBox
End Sub

'--------------------------------------------------------------------------------
Private Sub LimmitActivInput(ct As Control, MinVal As Integer, MaxVal As Integer)
'--------------------------------------------------------------------------------
  With ct
    If Not IsNumeric(.Value) Then
       While Len(.Value) > 0 And Not IsNumeric(.Value)
         .Value = DelLast(.Value)
       Wend
    Else
         If val(.Value) < MinVal Then .Value = MinVal
         If val(.Value) > MaxVal Then .Value = MaxVal
         If Round(.Value, 0) <> .Value Then .Value = Round(.Value, 0)
    End If
  End With
End Sub

'---------------------------------------
Private Sub LED_Channel_TextBox_Change()
'---------------------------------------
  LimmitActivInput LED_Channel_TextBox, 0, LED_CHANNELS - 1
End Sub

'----------------------------------
Private Sub MinCnt_TextBox_Change()
'----------------------------------
  MinCntChanged = True
  LimmitActivInput MinCnt_TextBox, 0, 127                                   ' 13.01.20: Old 255
End Sub

'----------------------------------
Private Sub MaxCnt_TextBox_Change()
'----------------------------------
  MaxCntChanged = True
  LimmitActivInput MaxCnt_TextBox, 1, 255
End Sub

'-------------------------------------------------------
Private Function Get_Calc_MinCnt(LEDCnt As Long) As Long
'-------------------------------------------------------
  Dim val As Long
  val = Round(LEDCnt / 3, 0)
  If val < 1 Then val = 1
  If val > 127 Then val = 127                                               ' 13.01.20:
  Get_Calc_MinCnt = val
End Function

'-------------------------------------------------------
Private Function Get_Calc_MaxCnt(LEDCnt As Long) As Long
'-------------------------------------------------------
  Get_Calc_MaxCnt = Round(2 * LEDCnt / 3, 0)
End Function

'----------------------------------------
Private Sub Set_MinMaxCnt(LEDCnt As Long)
'----------------------------------------
' Set MinCnt to 1/3 LEDCnt
' And MaxCnt to 2/3 LEDCnt
#If 1 Then
  Application.EnableEvents = False
  If Not MinCntChanged Then
     MinCnt_TextBox = Get_Calc_MinCnt(LEDCnt)
     MinCntChanged = False
  End If
  If Not MaxCntChanged Then
     MaxCnt_TextBox = Get_Calc_MaxCnt(LEDCnt)
     MaxCntChanged = False
  End If
  Application.EnableEvents = True
#End If
End Sub


'-----------------------------------
Private Sub MinTime_TextBox_Change()
'-----------------------------------
  LimmitActivInput MinTime_TextBox, 0, 254
End Sub

'-----------------------------------
Private Sub MaxTime_TextBox_Change()
'-----------------------------------
  LimmitActivInput MaxTime_TextBox, 0, 254
End Sub

'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  Dim Cnt As Long
  Cnt = Count_Used_RGB_Channels()
  If Cnt = 0 Then
     MsgBox Get_Language_Str("Das Haus enthält noch keine Räume" & vbCr & _
            "Bitte wählen die mindestens einen Raumtyp aus"), vbInformation, Get_Language_Str("Kein Raum ausgewählt")
     Exit Sub
  End If
  
  Userform_Res = Cnt & "$"                                                  ' 07.05.20: Old: "|"
  If Mode = "House" Then
    If IndividualTimes_CheckBox Then
          Userform_Res = Userform_Res & "HouseT"
    Else: Userform_Res = Userform_Res & "House"
    End If
  Else:   Userform_Res = Userform_Res & "GasLights"
  End If
  If InpInversBox Then Userform_Res = Userform_Res & "_Inv"                 ' 13.01.20:
  Userform_Res = Userform_Res & "(#LED, #InCh, "
  If Mode = "House" Then
     Userform_Res = Userform_Res & MinCnt_TextBox & ", " & MaxCnt_TextBox & ", "
     If IndividualTimes_CheckBox Then Userform_Res = Userform_Res & MinTime_TextBox & ", " & MaxTime_TextBox & ", "
  End If
  Userform_Res = Userform_Res & SelectedRooms_TextBox & ")$" & LED_Channel_TextBox ' 27.04.20:  ' 07.05.20: Old: "|"
  
  Store_Pos Me, HouseForm_Pos
  Unload Me ' Don't keep the entered data.
End Sub

'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  Userform_Res = ""
  Store_Pos Me, HouseForm_Pos
  Unload Me ' Don't keep the entered data.
End Sub

'-----------------------------------------
Private Function Count_Used_RGB_Channels()
'-----------------------------------------
' Single LEDs (Connected to a WS2811) are counted as one LED as long as they are in
' assending order.
' 1 2 3 = one RGB LED
' 1 1 1 = three RGB LEDs
  Dim Cnt As Long, X As Variant, SingleLED As Integer
  For Each X In Split(LED_CntList, " ")    ' LED_CntList = "RGB RGB 1 2 3 1 " for example
      If X = "RGB" Then
              If SingleLED = 0 Then
                    Cnt = Cnt + 1
              Else: Cnt = Cnt + 2
                    SingleLED = 0
              End If
      ElseIf X <> "" Then
              If val(X) <= SingleLED Then Cnt = Cnt + 1
              SingleLED = val(X)
      End If
  Next X
  If SingleLED > 0 Then Cnt = Cnt + 1
  Count_Used_RGB_Channels = Cnt
End Function

'-----------------------------
Private Function Count_Rooms()
'-----------------------------
  Dim Cnt As Long
  If SelectedRooms_TextBox <> "" Then
     Cnt = UBound(Split(SelectedRooms_TextBox, ",")) + 1
  End If
  Count_Rooms = Cnt
End Function

'-------------------------------------------------------
Private Sub Set_Fokus_To_Selected_Rooms_to_Show_Cursor()
'-------------------------------------------------------
  SelectedRooms_TextBox.setFocus ' To show the cursor in the SelectedRooms_TextBox
End Sub

'-------------------------------------
Private Sub Set_RoomCount(Cnt As Long)
'-------------------------------------
  RoomCnt_Label = Get_Language_Str("Anzahl: ") & Cnt
  Used_RGB_LEDs_Label = Get_Language_Str("RGB LED Kanäle: ") & Count_Used_RGB_Channels()
  Set_MinMaxCnt Cnt
  
  Set_Fokus_To_Selected_Rooms_to_Show_Cursor ' To show the cursor in the SelectedRooms_TextBox
End Sub

Private Sub CommandButton1_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton10_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton11_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton12_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton13_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton14_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton15_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton16_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton17_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton18_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton19_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton2_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton20_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton21_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton22_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton23_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton24_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton25_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton26_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton27_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton28_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton29_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton3_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton30_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton31_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton32_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton33_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton34_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton35_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton36_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton37_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton38_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton39_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton4_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton40_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton41_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton42_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton43_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton44_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton45_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton46_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton47_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton48_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton49_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton5_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton50_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton51_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton52_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton53_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton54_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton6_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton7_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton8_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton9_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton55_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton56_Click()
  Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton57_Click()
    Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton62_Click()
    Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton63_Click()
    Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton64_Click()
    Add_Room Me.ActiveControl.Caption
End Sub

Private Sub CommandButton65_Click()
    Add_Room Me.ActiveControl.Caption
End Sub


Private Sub SelectedRooms_TextBox_Click()
  Debug.Print "Textbox" & SelectedRooms_TextBox.SelStart
End Sub

'---------------------------------------------------------------------------------------------------------------------------------
Private Sub SelectedRooms_TextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'---------------------------------------------------------------------------------------------------------------------------------
  With SelectedRooms_TextBox
    Dim Word As Variant, ChrCnt As Long, Txt As String
    For Each Word In Split(.Text, ",")
      ChrCnt = Len(Txt) + 0.5 + Len(Word) / 2
      If .SelStart < ChrCnt Then
         .SelStart = Len(Txt)
         Exit Sub
      End If
      Txt = Txt & Word & ","
    Next Word
    .SelStart = Len(.Text)
  End With
End Sub

'-----------------------------------
Private Function Correct_Selection()
'-----------------------------------
  If SelectedRooms_TextBox.SelLength > 0 Then
     SelectedRooms_TextBox.SelLength = 0
     Do While SelectedRooms_TextBox.SelStart > 0
        If Mid(SelectedRooms_TextBox, SelectedRooms_TextBox.SelStart, 1) = "," Then Exit Do
        SelectedRooms_TextBox.SelStart = SelectedRooms_TextBox.SelStart - 1
     Loop
     Correct_Selection = True
     Set_Fokus_To_Selected_Rooms_to_Show_Cursor ' To show the cursor in the SelectedRooms_TextBox
  End If
End Function

'---------------------------------
Private Sub DelRoom_Button_Click()
'---------------------------------
  If SelectedRooms_TextBox = "" Then
     Set_Fokus_To_Selected_Rooms_to_Show_Cursor ' To show the cursor in the SelectedRooms_TextBox
     Exit Sub
  End If

  If Correct_Selection() Then Exit Sub
  
  Dim p As Long, e As Long, Txt As String
  e = SelectedRooms_TextBox.SelStart
  If e = 0 Then
     Set_Fokus_To_Selected_Rooms_to_Show_Cursor ' To show the cursor in the SelectedRooms_TextBox
     Exit Sub
  End If
  
  p = e - 1
  Txt = SelectedRooms_TextBox
  Do While Mid(Txt, p, 1) <> ","
    p = p - 1
    If p <= 0 Then Exit Do
  Loop
  Txt = Trim(Left(Txt, p) & Mid(Txt, e + 1))
  If Right(Txt, 1) = "," Then Txt = DelLast(Txt)
  SelectedRooms_TextBox = Txt
  SelectedRooms_TextBox.SelStart = p

  ' Delete the last element in LED_CntList
  Txt = LED_CntList ' LED_CntList = "RGB RGB 1 2 3 1 " for example
  If Right(Txt, 1) = " " Then Txt = Left(Txt, Len(Txt) - 1)
  While Right(Txt, 1) <> " " And Len(Txt) > 0
    Txt = Left(Txt, Len(Txt) - 1)
  Wend
  LED_CntList = Txt
  
  Set_RoomCount Count_Rooms()
  
End Sub

'-------------------------------------------------------------------
Private Function Get_LED_Channel_from_Name(Name As String) As String
'-------------------------------------------------------------------
  Dim p As Long, e As Long
  p = InStr(LED_ChannelsList, " " & Name & ":")
  If p = 0 Then
        MsgBox "Internal Error: '" & Name & "' not found in 'LED_ChannelsList'", vbCritical, "Internal Error"
        End
  Else: p = p + 1 + Len(Name) + 1
        e = InStr(p, LED_ChannelsList, " ")
        Get_LED_Channel_from_Name = Mid(LED_ChannelsList, p, e - p)
  End If
End Function

'--------------------------------------
Private Sub Add_Room(Caption As String)
'--------------------------------------
  Dim Cnt As Long
  Cnt = Count_Rooms()
  If Cnt >= 250 Then
     MsgBox Get_Language_Str("Es können keine weiteren Räume hinzugefügt werden"), vbInformation, Get_Language_Str("Maximale Raumanzahl erreicht")
     Set_Fokus_To_Selected_Rooms_to_Show_Cursor ' To show the cursor in the SelectedRooms_TextBox
     Exit Sub
  End If
  
  If Correct_Selection() Then Exit Sub
  
  LED_CntList = LED_CntList & Get_LED_Channel_from_Name(Caption) & " "
  
  If SelectedRooms_TextBox = "" Then
        SelectedRooms_TextBox = Caption
  Else: ' Insert to existing string
        Dim p As Long
        p = SelectedRooms_TextBox.SelStart
        If p >= Len(SelectedRooms_TextBox) Then
            SelectedRooms_TextBox = SelectedRooms_TextBox & ", " & Caption
            p = p + 1
        Else
            If p > 0 Then
               If Mid(SelectedRooms_TextBox, p, 1) = "," Then p = p + 1
            End If
            SelectedRooms_TextBox = Left(SelectedRooms_TextBox, p) & Caption & ", " & Mid(SelectedRooms_TextBox, p + 1)
        End If
        SelectedRooms_TextBox.SelStart = p + Len(Caption) + 1
  End If
  Set_RoomCount Count_Rooms()
End Sub

'------------------------------------------
Private Sub Change_Height(factor As Double)
'------------------------------------------
' Factor 0.85 => 712  (width 1110)
' factor 0.9  => 745
  Dim X As Byte
  Me.Height = (Me.Height + 5) * factor
  Dim obj As Control
  For Each obj In Controls
    obj.Top = obj.Top * factor
    obj.Height = obj.Height * factor
  Next
  Description_Label.Font.Size = Description_Label.Font.Size * factor
  Label_NotchangableCol.Font.Size = Label_NotchangableCol.Font.Size * factor
End Sub



'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & Me.Name & ": UserForm_Initialize"
  
  Dim Monitor_Pixel_Cnt_Y As Long                                           '23.04.20:
  Monitor_Pixel_Cnt_Y = Get_Primary_Monitor_Pixel_Cnt_Y()
  If Monitor_Pixel_Cnt_Y <= 720 Then
                                        Change_Height 0.85 ' Scale to a height of 712  (Fit into 768)
  ElseIf Monitor_Pixel_Cnt_Y <= 833 Then
                                        Change_Height 0.9  ' Scale to a height of 745  (Fit into 768)
  End If
  
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Restore_Pos_or_Center_Form Me, HouseForm_Pos
  
  #If False Then ' Debug
    SelectedRooms_TextBox.Text = "ROOM_D_RED, ROOM_COL4, ROOM_CHIMNEYD, ROOM_TV1, NEON_LIGHTD, NEON_LIGHTL, SINGLE_LED2D, GAS_LIGHTD, GAS_LIGHT1D, SINGLE_LED3D, NEON_LIGHT1L, NEON_LIGHT1D, ROOM_TV1_CHIMNEY, ROOM_CHIMNEYB"
    LED_CntList = "               RGB         RGB        RGB            RGB       RGB          RGB          2             RGB         1            3             1             1             RGB               RGB"
    Set_RoomCount Count_Rooms()
  #End If
  
  OK_Button.setFocus
  
End Sub

'---------------------------------------
Private Sub SetMode(MacroName As String)
'---------------------------------------
  Mode = MacroName
  Select Case MacroName
    Case "House":     'Set_Color_in_y_Range CommandButton27.Top, CommandButton54.Top, &H80000011   ' 11.10.20: Disabled
                      Set_Bold_in_y_Range CommandButton1.Top, CommandButton27.Top, True            ' 11.10.20:
                      
    Case "GasLights": 'Set_Color_in_y_Range CommandButton35.Top, CommandButton27.Top, &H80000011   ' 11.10.20: Disabled
                      Set_Bold_in_y_Range CommandButton27.Top, CommandButton54.Top, True           ' 11.10.20:
                      Hide_and_Move_up Me, "CommandButton1", "CommandButton35"
                      Hide_and_Move_up Me, "MinCnt_TextBox", "Abort_Button"
                      Label_NotchangableCol.Visible = False
                      Me.Caption = Get_Language_Str("Gaslights: Die Gaslaternen werden zufällig, nacheinander aktiviert. Sie erreichen erst nach einiger Zeit die volle Helligkeit und flackern manchmal.")
                      Me.Description_Label = Get_Language_Str("Straßenlaternen sind ein wichtiger Bestandteil einer virtuellen Stadt. Sie beleuchten die nächtlichen Straßen und erzeugen eine warme Atmosphäre insbesondere, wenn es sich um Gaslaternen handelt. " & _
                                             "Die Lampen gehen zufällig an und werden dann langsam heller bis sie die volle Helligkeit erreichen. Außerdem ist noch ein zufälliges Flackern implementiert welches durch Schwankungen im Gasdruck oder durch Windböen entstehen kann.")
  End Select
End Sub

'----------------------------------------------------------------------------------------------------------------------
Public Sub Show_With_Existing_Data(MacroName As String, ConfigLine As String, LED_Channel As Long, Def_Channel As Long) ' 27.04.20: Added: LED_Channel and Def_Channel
'----------------------------------------------------------------------------------------------------------------------
  SetMode MacroName
  LED_CntList = ""
  IndividualTimes_CheckBox = False
  
  Dim Txt As String
  LED_Channel_TextBox = Def_Channel                                         ' 27.04.20:
  If Len(ConfigLine) > Len(MacroName) Then
     If Left(ConfigLine, Len(MacroName)) = MacroName Then
        Dim Parts() As String, Nr As Long, i As Long
        Parts = Split(Replace(Split(ConfigLine, "(")(1), ")", ""), ",")
        If MacroName = "House" Then
              MinCnt_TextBox = val(Parts(2))
              MaxCnt_TextBox = val(Parts(3))
              If Left(ConfigLine, Len("HouseT")) = "HouseT" Then
                    MinTime_TextBox = val(Parts(4))
                    MaxTime_TextBox = val(Parts(5))
                    IndividualTimes_CheckBox = True
                    Nr = 6
              Else: Nr = 4
              End If
        Else: Nr = 2
        End If
        InpInversBox = (Right(Split(ConfigLine, "(")(0), Len("_Inv")) = "_Inv")   ' 13.01.20:
        For i = Nr To UBound(Parts)
           Txt = Txt & Trim(Parts(i))
           If i < UBound(Parts) Then Txt = Txt & ", "
           LED_CntList = LED_CntList & Get_LED_Channel_from_Name(Trim(Parts(i))) & " "
        Next i
        LED_Channel_TextBox = LED_Channel                                   ' 27.04.20:
     End If
  End If
  SelectedRooms_TextBox.Text = Txt
  Set_RoomCount Count_Rooms()
  If LED_CntList <> "" Then
        Dim LEDCnt As Long
        LEDCnt = Count_Rooms()
        MinCntChanged = (Get_Calc_MinCnt(LEDCnt) <> MinCnt_TextBox)
        MaxCntChanged = (Get_Calc_MaxCnt(LEDCnt) <> MaxCnt_TextBox)
  Else: MinCntChanged = False
        MaxCntChanged = False
  End If
  Set_Fokus_To_Selected_Rooms_to_Show_Cursor ' To show the cursor in the SelectedRooms_TextBox
  Me.Show
End Sub

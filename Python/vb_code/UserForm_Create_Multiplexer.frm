VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Create_Multiplexer 
   Caption         =   "UserForm1"
   ClientHeight    =   8772
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13830
   OleObjectBlob   =   "UserForm_Create_Multiplexer.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm_Create_Multiplexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Binary ' Use case sensitive compare. Important for the "Sec" compare below

' Following parameters exist:
' LED Cx  InCh    Val0    Val1    On_Min  On_Limit    LedCnt  Brightness  DstVar  Mode    Enable  TimeOut DstVar1 DstVarN MinTime MaxTime MinOn   MaxOn   Var SrcLED  EnableCh    Start   End GlobVarNr   FirstInCh   InCh_Cnt    Duration    Period  Pause   Act Off S_InCh  R_InCh  Timeout T_InCh  DstVar0 MinBrightness   MaxBrightness   B_LED_Cx    B_LED   InNr    TmpNr   Rotate  Steps   MaxSoundNr  InReset
'
'
' Maximal 6 named parameter in one macro
'
' ToDo:
' - Alte eingaben speichern und beim nächsten mal wieder verwenden
'   - Geht nicht in dem man das Unload Me weglässt
'   - Man könnte ein Array anlegen in dem die Parameter Namen und der Letzte Wert enthalten ist
'     das sollte aber für jedes Makro spezifisch sein

Private ParList As Variant
Private FuncName As String
Private NamesA As Variant

Const MAX_PAR_CNT = 14                                                      ' 14.01.20: Old 6 (Warum ncht 7? Es waren doch 7 Felder verfügbar)
Private TypA(MAX_PAR_CNT) As String
Private MinA(MAX_PAR_CNT) As Variant
Private MaxA(MAX_PAR_CNT) As Variant
Private ParName(MAX_PAR_CNT) As String


'-------------------------------------------------------------------------------
Private Function Check_Limit_to_MinMax(ParNr As Long, Value As Variant) As Boolean
'-------------------------------------------------------------------------------
' Return true if its within the alowed range
  Dim Msg As String
  'With Controls("Par" & ParNr)
    Dim ValidRangeTxt As String
    ValidRangeTxt = vbCr & Get_Language_Str("Bitte einen Wert zwischen ") & MinA(ParNr - 1) & Get_Language_Str(" und ") & MaxA(ParNr - 1) & Get_Language_Str(" eingeben.")
    If Value = "" Then
                                                                       Msg = Get_Language_Str("leer.") & ValidRangeTxt
    ElseIf Not IsNumeric(Value) Then
                                                                       Msg = Get_Language_Str("keine gültige Zahl.") & ValidRangeTxt
    ElseIf CStr(Round(Value, 0)) <> CStr(Value) Then
                                                                       Msg = Get_Language_Str("nicht Ganzzahlig.") & ValidRangeTxt

    ElseIf MinA(ParNr - 1) <> "" And Value < val(MinA(ParNr - 1)) Then
                                                                       Msg = Get_Language_Str("zu klein!" & vbCr & "Der Minimal zulässiger Wert ist: ") & MinA(ParNr - 1)
    ElseIf MaxA(ParNr - 1) <> "" And Value > val(MaxA(ParNr - 1)) Then
                                                                       Msg = Get_Language_Str("zu groß!" & vbCr & "Der Maximal zulässige Wert ist: ") & MaxA(ParNr - 1)
    End If

    If Msg <> "" Then
          Controls("Par" & ParNr).setFocus
          MsgBox Get_Language_Str("Der Parameter '") & Controls("LabelPar" & ParNr).Caption & Get_Language_Str("' ist ") & Msg, vbInformation, Get_Language_Str("Bereichsüberschreitung")
    Else: Check_Limit_to_MinMax = True
    End If
  'End With
End Function

'-----------------------------------------------------------
Private Function Check_Time_String(ParNr As Long) As Boolean
'-----------------------------------------------------------
  Dim ValidRangeTxt As String
  ValidRangeTxt = vbCr & Get_Language_Str("Bitte einen Zeit zwischen ") & MinA(ParNr - 1) & Get_Language_Str(" ms und ") & MaxA(ParNr - 1) & Get_Language_Str(" ms eingeben." & vbCr & _
                  "Die Zeitangabe kann auch eine der folgenden Einheit enthalten:" & vbCr & _
                  " Min, Sec, ms " & vbCr & _
                  "Achtung: Zwischen Zahl und Einheit muss ein Leerzeichen stehen." & vbCr & _
                  "Beispiel: 3 Sec") ' ToDo: Erlaubte Zeiten zusätzlich in Minuten Angeben
  Dim Parts As Variant
  With Controls("Par" & ParNr)
    Parts = Split(.Value, " ")
    Const ValidUnits = " Min Sek sek Sec sec Ms ms "
    If UBound(Parts) <> 1 Or Not IsNumeric(Parts(0)) Or InStr(ValidUnits, " " & Parts(1) & " ") = 0 Then
         MsgBox Get_Language_Str("Der Parameter '") & ParName(ParNr - 1) & Get_Language_Str("' ist ungültig"), vbInformation, Get_Language_Str("Ungültiger Parameter") & ValidRangeTxt
         Exit Function
    Else ' Two parameter detected. First is numeric, the second is a valid Unit
         Dim val As Double
         Select Case LCase(Parts(1))
           Case "min":             val = Parts(0) * 60 * 1000
           Case "sec", "sek": val = Parts(0) * 1000
           Case "ms":              val = Parts(0)
           Case Else:              MsgBox "Internal error: Unknown unit '" & Parts(1) & "' in Check_Time_String()", vbCritical, "Internal error"
                                   EndProg
         End Select
         If Check_Limit_to_MinMax(ParNr, val) = False Then Exit Function
    End If
  End With
  Check_Time_String = True
End Function


'-------------------------------------------------------------------------------------
Private Function Check_Par_with_ErrMsg(ParNr As Long, ByRef val As Variant) As Boolean
'-------------------------------------------------------------------------------------
  If ParNr > MAX_PAR_CNT Then
     MsgBox "Internal error in Chek_Range()"
     EndProg
  End If
  With Controls("Par" & ParNr)
    .Value = Trim(.Value)
    'Debug.Print "Check_Range " & ParName(ParNr - 1) & ": " & .Value
    Select Case TypA(ParNr - 1)
      Case "":     ' Normal Numeric parameter
                   If Check_Limit_to_MinMax(ParNr, .Value) = False Then Exit Function
      Case "Time": ' time could have a tailing "Min", "Sek", "sek", "Sec", "sec", "Ms", "ms"
                   If IsNumeric(.Value) Then
                        .Value = Int(.Value)
                        If Check_Limit_to_MinMax(ParNr, .Value) = False Then Exit Function
                   Else ' The parameter is NOT numeric
                        If Check_Time_String(ParNr) = False Then Exit Function
                   End If
                   .Value = Replace(.Value, ",", ".") ' Replace the german comma
      Case "Var":  ' No Check at the moment                                                        ' 14.01.20:
                   Debug.Print "Check parameter typ 'Var' for '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "Txt":  ' No Check at the moment
                   Debug.Print "Check parameter typ 'Txt' for '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "Mode": ' No Check at the moment
                   Debug.Print "Check parameter typ 'Mode' for '" & TypA(ParNr - 1) & "'" ' ToDo
      Case Else:   MsgBox "Internal error: Unknown parameter Typ: '" & TypA(ParNr - 1) & "'", vbCritical, "Internal Error"
                   EndProg
    End Select
    val = .Value
  End With
  Check_Par_with_ErrMsg = True
End Function

#If 0 Then                                                                  ' 14.06.20: Disabled because there is no "ScrollBar1"
'------------------------------
Private Sub ScrollBar1_Change()
'------------------------------
' ToDo: Die Steuerung ist noch nicht gut
' - Slider Size anpassen
' - Die Description_TextBox sollte nicht editierbar sein
'
  'Debug.Print ScrollBar1.Value
  With Description_TextBox
    .setFocus
    'Debug.Print .LineCount & " " & .MaxLength & " " & .TextLength
    .SelStart = ScrollBar1.Value
  End With
  ScrollBar1.setFocus
End Sub

'-----------------------------------------------
Public Sub MouseWheel(ByVal lngRotation As Long)
'-----------------------------------------------
' Process the mouse wheel changes
  'Debug.Print "MouseWheel" & lngRotation ' Debug
  If lngRotation < 0 Then
        ScrollBar1 = Application.Min(ScrollBar1 + 100, ScrollBar1.Max)
  Else: ScrollBar1 = Application.Max(ScrollBar1 - 100, ScrollBar1.Min)
  End If
End Sub
#End If

'--------------------------------------------------------------------------------------------------------------------
'Private Sub Description_TextBox_MouseWheel(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles Description_TextBox.MouseWheel
'--------------------------------------------------------------------------------------------------------------------
'    Dim v As Integer
'    v = Me.VerticalScroll.Value
'    v -= e.Delta \ 10
'        v = Math.Max(v, 0)
'    Me.VerticalScroll.Value = v
'End Sub

'--------------------------------------------------------------------------------------------------------------------
Public Sub Show_UserForm_Other(ByVal par As String, ByVal Name As String, Description As String, LedChannels As Long)
'--------------------------------------------------------------------------------------------------------------------
  FuncName = Name
  Const CNames = "Val0 Val1 Period Duration Timeout DstVar MinTime MaxTime Par1"
  If Description = "" Then
        Description_TextBox = Get_Language_Str("Noch keine Beschreibung zur Funktion '") & Name & Get_Language_Str("' vorhanden ;-(")
  Else: Description_TextBox = Description
  End If
#If 0 Then                                                                  ' 14.06.20: Disabled
'  ScrollBar1.Max = Len(Description_TextBox)
  HookFormScroll Me, "Description_TextBox"   ' Initialize the mouse wheel scroll function
#End If
  Me.Caption = Get_Language_Str("Parametereingabe der '") & Name & Get_Language_Str("' Funktion")

'  Debug.Print                         ' Debug
'  Debug.Print Name & " (" & Par & ")" ' Debug

  ParList = Split(par, ",")

  Dim OldCmdLine As String, UseOldParams As Boolean, OldParams() As String
  OldCmdLine = Cells(ActiveCell.Row, Config__Col)
  If Len(OldCmdLine) > Len(Name) And left(OldCmdLine, Len(Name)) = Name Then
     UseOldParams = True
     OldParams = Split(Trim(Replace(Mid(OldCmdLine, Len(Name) + 2), ")", "")), ",")
  End If

  ' Add parameters
  Dim p As Variant, UsedParNr As Long, Nr As Long
  For Each p In ParList
    p = Trim(p)
    If UsedParNr = 7 Then
        Debug.Print "Achtung"
    End If
    If left(p, 1) <> "#" Then
        If UsedParNr >= MAX_PAR_CNT Then
            MsgBox "Internal error: The number of parameters is to large in Show_UserForm_Other()"
            EndProg
        End If
        Dim Typ As String, Min As String, Max As String, Def As String, Opt As String, InpTxt As String, Hint As String
        Get_Par_Data p, Typ, Min, Max, Def, Opt, InpTxt, Hint
        If UseOldParams Then
            Def = Trim(OldParams(Nr))
        End If
        TypA(UsedParNr) = Typ
        MinA(UsedParNr) = Min
        MaxA(UsedParNr) = Max
        ParName(UsedParNr) = p

        UsedParNr = UsedParNr + 1
        Me.Controls("Par" & UsedParNr) = Def
        Me.Controls("LabelPar" & UsedParNr) = InpTxt
        Me.Controls("LabelPar" & UsedParNr).ControlTipText = Hint
    End If
    Nr = Nr + 1
  Next p
  
  Call Get_Multiplexer_Names
  For Nr = 1 To 6
      If Me.Controls("LabelPar" & Nr).Caption = "ControlNr" Then Me.Controls("LabelPar" & Nr) = Get_Language_Str("Kontroll Nummer")
      If Me.Controls("LabelPar" & Nr).Caption = "Groups" Then Me.Controls("LabelPar" & Nr) = Get_Language_Str("Anzahl der Gruppen im Multiplexer")
      If Me.Controls("LabelPar" & Nr).Caption = "RndMinTime" Then Me.Controls("LabelPar" & Nr) = Get_Language_Str("Minimale Umschaltzeit zwischen Patronen")
      If Me.Controls("LabelPar" & Nr).Caption = "RndMaxTime" Then Me.Controls("LabelPar" & Nr) = Get_Language_Str("Maximale Umschaltzeit zwischen Patronen")
      If Me.Controls("LabelPar" & Nr).Caption = "NumOfLEDs" Then Me.Controls("LabelPar" & Nr) = Get_Language_Str("Anzahl der LEDs in den Patronen")
  Next Nr
  Me.Controls("LabelPar7").Caption = Get_Language_Str("LED-Typ")
  
  If UseOldParams Then
      Call Set_CheckBox(val(OldParams(5)))
      Call Set_CtrMode(Trim(OldParams(8)))
  Else
      Me.Controls("CheckBox1") = True           ' Default First option
      Me.Controls("OptionButtonSEQ") = True     ' Default Sequentieel
  End If

  Center_Form Me                                                            ' 18.01.20:
  Me.Show

End Sub


'--------------------------------------------------------------------------------------------------------------------
Private Sub Get_Multiplexer_Names()
'--------------------------------------------------------------------------------------------------------------------
    Dim ProgDir As String
    Dim IniFileName As String
    Dim FileName, Map As String, Nr As Integer
    
    Map = Environ("USERPROFILE") & "\Documents\" & "MyPattern_Config_Examples"
    IniFileName = Map & "\" & Multiplexer_INI_FILE_NAME
    
    ProgDir = IniFileName
    If Dir(ProgDir, vbDirectory) = "" Then
       MsgBox Get_Language_Str("Fehler das Verzeichnis existiert nicht:") & vbCr & _
              "  '" & ProgDir & "'", vbCritical, Get_Language_Str("Multiplexer Verzeichnis nicht vorhanden")
       Exit Sub
    End If
  
    FileName = IniFileName
    If Not Dir(FileName) <> "" Then
      MsgBox Get_Language_Str("Fehler die Datei existiert nicht:") & vbCr & _
              "  '" & FileName & "'", vbCritical, Get_Language_Str("Kirmes-Datei nicht gefunden!")
      Exit Sub
    End If
    
    ' Syntax for reading INI file
    ' Section = "Multiplexer_RGB_Ext4"              ' MultiplexerName
    ' KeyName = "Option 1 Name"                     ' Variable Name
    ' Value   = ReadIniFileString(Section, KeyName) ' Variable Value
    
    Dim OptionName As String
    
    Dim sectnNames() As String, strBuffer As String
    Dim intx As Integer, strfullpath As String
    
    Erase sectnNames
    
    strfullpath = IniFileName                        ' Get Path and Name of "Multiplexer.ini" file.
    
    Let strBuffer$ = String$(1000, Chr$(0&))         ' Size of strBuffer$ = 1000, filled with 0 (zero's).
    Call GetPrivateProfileSectionNames(strBuffer, Len(strBuffer), strfullpath)
    sectnNames = Split(strBuffer, vbNullChar)
    
    For intx = LBound(sectnNames) To UBound(sectnNames)
        If sectnNames(intx) = vbNullString Then Exit For
        If intx > 2 Then ComboBox1.AddItem Mid(sectnNames(intx), Len("Mutliplexer_") + 1)
        If sectnNames(intx) = "Multiplexer_" & Cells(ActiveCell.Row, Descrip_Col).Value Then
            ComboBox1.Value = Cells(ActiveCell.Row, Descrip_Col).Value
        End If
    Next intx

End Sub

'--------------------------------------
Private Sub DCC_Button_Checkbox_Click()
'--------------------------------------
    If DCC_Button_Checkbox = True Then
        Me.DCC_Address_Or_Name_Label.Enabled = True
        Me.DCC_Address_Or_Name_TextBox.Enabled = True
    Else
        Me.DCC_Address_Or_Name_Label.Enabled = False
        Me.DCC_Address_Or_Name_TextBox.Enabled = False
        Me.DCC_Address_Or_Name_TextBox.Value = ""           ' 10.02.21: 20201102 Misha. DCC address cleared.
    End If
End Sub


'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  
  'Debug.Print vbCr & me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Restore_Pos_or_Center_Form Me, OtherForm_Pos
End Sub


'--------------------------------------------------------------------------------------------------------------------
Private Sub ComboBox1_Change()
'--------------------------------------------------------------------------------------------------------------------
    
    Dim i, SelectItem, Pos, CbxNr As Integer, DCC_Address_Val As String
    
    ' Find selected item in Combolist
    If ComboBox1.ListIndex <> -1 Then
        SelectItem = ComboBox1.List(ComboBox1.ListIndex)
        Pos = ComboBox1.ListIndex + 1
    End If
    
    Me.Controls("Par5").Value = ReadIniFileString("Multiplexer_" & SelectItem, "ControlNr")
    Me.Controls("Par6").Value = ReadIniFileString("Multiplexer_" & ComboBox1.Value, "Number_Of_LEDs")
    Me.Controls("Par7").Value = ReadIniFileString("Multiplexer_" & ComboBox1.Value, "LED_Type")
    
    If ReadIniFileString("Multiplexer_" & ComboBox1.Value, "Enable_DCC_Button") = 1 Then
        Me.DCC_Button_Checkbox.Enabled = True                                                   ' 10.02.21:
        DCC_Address_Val = Cells(ActiveCell.Row, DCC_or_CAN_Add_Col).Value                       ' 20201102 Misha. Changes for filling DCC Address.
        If DCC_Address_Val <> "[Multiplexer]" And DCC_Address_Val <> "" Then                    ' 20201102 Misha. Check if DCC Address is filled with DCC address.
            Me.DCC_Button_Checkbox = True                                                       ' 20201102 Misha. Checkbox checked.
            Me.DCC_Address_Or_Name_TextBox.Value = DCC_Address_Val                              ' 20201102 Misha. Fill DCC address.
        Else
            Me.DCC_Button_Checkbox = False                                                      ' 20201102 Misha. Checkbox unchecked.
            Me.DCC_Address_Or_Name_TextBox.Value = ""                                           ' 20201102 Misha. DCC address cleared.
        End If
    Else
        Me.DCC_Button_Checkbox = False                                                          ' 20201102 Misha. Checkbox unchecked.
        Me.DCC_Address_Or_Name_TextBox.Value = ""                                               ' 20201102 Misha. DCC address cleared.
        Me.DCC_Button_Checkbox.Enabled = False
        Me.DCC_Address_Or_Name_Label.Enabled = False
        Me.DCC_Address_Or_Name_TextBox.Enabled = False
    End If
    
    ' Fill the Checkboxes with the Multiplexer Options
    For CbxNr = 1 To 8
        Me.Controls("CheckBox" & CbxNr).Caption = ReadIniFileString("Multiplexer_" & SelectItem, "Option " & CbxNr & " Name")
    Next CbxNr
        
End Sub


'--------------------------------------------------------------------------------------------------------------------
Private Sub Set_CheckBox(Options As Integer)
'--------------------------------------------------------------------------------------------------------------------
    Dim CbxNr As Integer, binOptions As String

    For CbxNr = 1 To 8
        binOptions = DecToBin(Options)
        If right(binOptions, 1) = 1 Then
            ' LSB = 1
            Me.Controls("CheckBox" & CbxNr) = True
        Else
            Me.Controls("CheckBox" & CbxNr) = False
        End If
        Options = Application.WorksheetFunction.Bitrshift(Options, 1)
    Next CbxNr

End Sub


'-------------------------------------------------------------------------------------
Private Function Get_Options_Byte(ByRef val As Variant)
'-------------------------------------------------------------------------------------
    Dim ParNr As Integer
    val = 0
    For ParNr = 1 To 8
        With Controls("Checkbox" & ParNr)
            Debug.Print "Checkbox " & ParNr & ": " & .Value
            If .Value And Not .Caption = "To be filled!" Then
                val = val + (2 ^ (ParNr - 1))
            ElseIf .Value And .Caption = "To be filled!" Then
                ' Not used option selected
                MsgBox Get_Language_Str("Fehler: Eine nicht verwendete Pattern wurde ausgewählt!") & vbCr, _
                                        vbCritical, Get_Language_Str("Kirmes Einstellung hat sich nicht geändert!")
                .Value = False
            End If
        End With
    Next ParNr
    Get_Options_Byte = val
    
End Function


'-------------------------------------------------------------------------------------
Private Function Set_CtrMode(ByRef CtrMode As Variant)
'-------------------------------------------------------------------------------------
    If CtrMode = "CF_ROTATE" Or CtrMode = "CF_ROTATE|CF_SKIP0" Then Me.Controls("OptionButtonSEQ") = True ' 10.02.21: Misha added Or ...
    If CtrMode = "CF_RANDOM" Or CtrMode = "CF_RANDOM|CF_SKIP0" Then Me.Controls("OptionButtonRnd") = True '   "
   
End Function


'-------------------------------------------------------------------------------------
Private Function Get_CtrMode(ByRef val As Variant)
'-------------------------------------------------------------------------------------
    val = 0
    If Me.Controls("OptionButtonSEQ") Then val = "CF_ROTATE|CF_SKIP0" ' 10.02.21: Misha added "|CF_SKIP0"
    If Me.Controls("OptionButtonRnd") Then val = "CF_RANDOM|CF_SKIP0" '    "
    Get_CtrMode = val
    
End Function


'-------------------------------------------------------------
Private Function Create_Result(ByRef Res As String) As Boolean
'-------------------------------------------------------------
' Return True if sucessfully checked alt inputs
  
  Res = ""
  Dim p As Variant
  For Each p In ParList
      Dim val As Variant
      val = "Not Found"
      p = Trim(p)
      If left(p, 1) = "#" Then
        If p = "#Options" Then
            val = Get_Options_Byte(val)
        ElseIf p = "#CtrMode" Then
            val = Get_CtrMode(val)
        Else
            val = p
        End If
      Else
#If 0 Then                                                                  ' 14.06.20: Disabled
           If p = "Cx" Or p = "B_LED_Cx" Then
                MsgBox "Disabled 'Get_OptionButton_Res()'"
                val = Get_OptionButton_Res()
           Else ' Not a standard parameter
#End If
                Dim Nr As Long
                For Nr = 1 To MAX_PAR_CNT
                    If ParName(Nr - 1) = p Then
                       If Check_Par_with_ErrMsg(Nr, val) = False Then Exit Function
                       Exit For
                    End If
                Next
           End If
#If 0 Then                                                                  ' 14.06.20: Disabled
      End If
#End If
      If val = "Not Found" Then MsgBox Get_Language_Str("Fehler der Parameter '") & p & Get_Language_Str("' wurde nicht gefunden"), vbCritical, Get_Language_Str("Programm Fehler")
      Res = Res & val & ", "
  Next p
  Res = FuncName & "(" & DelLast(Res, 2) & ")"
  Create_Result = True
End Function


'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
#If 0 Then                                                                  ' 14.06.20: Disabled
  UnhookFormScroll ' Deactivate the mouse wheel scroll function
#End If
  Userform_Res = ""
  Store_Pos Me, OtherForm_Pos
  
  Unload Me ' Don't keep the entered data. Importand because the positions of the controlls and the visibility have been changed
End Sub


'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  If Me.ComboBox1.Value <> "" Then ' Anything selected ?                     ' 14.06.20:
     ' Todo: Check all other fields
     
     If Create_Result(Userform_Res) Then
        ' Place MultiplexerName on active worksheet in Description Column
        Cells(ActiveCell.Row, Descrip_Col).Value = Me.ComboBox1.Value
'        Cells(ActiveCell.Row, DCC_or_CAN_Add_Col).Value = Me.DCC_Address_Or_Name_TextBox.Value ' 10.02.21: 20210208 Disabled by Misha,
                                                                                                ' This is trigering a sheet change event and after that an 'Res' error in
                                                                                                ' the Multiplexer module 'M80_Create_Multiplexer/Function Special_Multiplexer_Ext'.
        Userform_Res_Address = Me.DCC_Address_Or_Name_TextBox.Value                             ' 20210208 Added by Misha, Store Address to fill address cell (Col 4) after ending of the sheet change event.
     End If
#If 0 Then                                                                  ' 14.06.20: Disabled
     UnhookFormScroll ' Deactivate the mouse wheel scroll function
#End If
     Store_Pos Me, OtherForm_Pos

     Unload Me ' Don't keep the entered data. Importand because the positions of the controlls and the visibility have been changed
  End If
End Sub

#If 0 Then                                                                  ' 14.06.20: Disabled
'------------------------------------------------
Private Function Get_OptionButton_Res() As String
'------------------------------------------------
  Dim val As String
  If OptionButton_All Then
                               val = "C_ALL"
  ElseIf OptionButton_C1 Then: val = "C1"
  ElseIf OptionButton_C2 Then: val = "C2"
  ElseIf OptionButton_C3 Then: val = "C3"
  ElseIf OptionButton_12 Then: val = "C12"
  ElseIf OptionButton_23 Then: val = "C23"
  Else: MsgBox Get_Language_Str("LED Auswahl Fehler"), vbCritical
  End If
  Get_OptionButton_Res = val
End Function


'------------------------------------------
Private Sub Set_OptionButton(val As String)
'------------------------------------------
  Select Case val
    Case "C_ALL": OptionButton_All = True
    Case "C1":    OptionButton_C1 = True
    Case "C2":    OptionButton_C2 = True
    Case "C3":    OptionButton_C3 = True
    Case "C12":   OptionButton_12 = True
    Case "C23":   OptionButton_23 = True
    Case Else:    OptionButton_All = True
                  MsgBox Get_Language_Str("Fehler beim Lesen der bestehenden Kanalbezeichnung '") & val & "'", vbCritical, Get_Language_Str("Unbekannte Kanalbezeichnung")
  End Select
End Sub
#End If



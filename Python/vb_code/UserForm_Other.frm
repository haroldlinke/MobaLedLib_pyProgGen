VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Other 
   Caption         =   "UserForm1"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   OleObjectBlob   =   "UserForm_Other.frx":0000
End
Attribute VB_Name = "UserForm_Other"
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
Private Show_Channel_Type As Long
Private CurWidth As Long                                                    ' 05.11.21: Juergen form resize feature
Private CurHeight As Long
Private MinFormHeight As Long
Private MinFormWidth As Long

Const MAX_PAR_CNT = 14                                                      ' 14.01.20: Old 6 (Warum ncht 7? Es waren doch 7 Felder verfügbar)
Private TypA(MAX_PAR_CNT) As String
Private MinA(MAX_PAR_CNT) As Variant
Private MaxA(MAX_PAR_CNT) As Variant
Private ParName(MAX_PAR_CNT) As String
Private Invers(MAX_PAR_CNT) As Boolean

Const DEFAULT_PAR_WIDTH = 48

#If VBA7 Then
  Private Declare PtrSafe Function SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long) As Long
  Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
  Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
  Private Declare Function SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long) As Long
  Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If


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
    Dim Err As Boolean
    Err = UBound(Parts) <> 1
    If Not Err Then Err = Not IsNumeric(Parts(0))
    If Not Err Then Err = InStr(ValidUnits, " " & Parts(1) & " ") = 0
    If Err Then
         MsgBox Get_Language_Str("Der Parameter '") & ParName(ParNr - 1) & Get_Language_Str("' ist ungültig") & ValidRangeTxt, vbInformation, Get_Language_Str("Ungültiger Parameter")
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

'------------------------------------------------------------------------------------------------
Private Function Check_RGB_List(TypStr As String, ByVal Value As String, ParNr As Long) As String
'------------------------------------------------------------------------------------------------
  Dim s As Variant, Err As Boolean, ExpCnt As Long
  If TypStr = "RGB" Then
        ExpCnt = 3
  Else: ExpCnt = 6
  End If
  
  If UBound(Split(Value, ",")) <> ExpCnt - 1 Then
     MsgBox Replace(Get_Language_Str("Fehler: Es müssen #1# Farbwerte zwischen 0 und 255 angegeben werden"), "#1#", ExpCnt), _
            vbCritical, Get_Language_Str("Anzahl der angegebenen Farbwerte ist falsch")
     Exit Function
  End If
  
  For Each s In Split(Value, ",")
      If Not IsNumeric(Trim(s)) Then Err = True
      If Not Err Then Err = val(s) < 0 Or val(s) > 255 Or InStr(Value, ".") > 0
      If Err Then
         MsgBox Replace(Get_Language_Str("Fehler der Parameter '#1#' ist ungültig." & vbCr & _
                "Die Farbwerte müssen im Bereich von 0 bis 255 liegen"), "#1#", Controls("LabelPar" & ParNr)), _
         vbCritical, Get_Language_Str("Ungültiger RGB Parameter")
         Exit Function
      End If
  Next s
  Check_RGB_List = Value
End Function
                      

'-------------------------------------------------------------------------------------
Private Function Check_Par_with_ErrMsg(ParNr As Long, ByRef val As Variant) As Boolean
'-------------------------------------------------------------------------------------
  If ParNr > MAX_PAR_CNT Then
     MsgBox "Internal error in Chek_Range()"
     EndProg
  End If
  Dim VarLabel As String, ShowErr As Boolean
  VarLabel = Controls("LabelPar" & ParNr)
  With Controls("Par" & ParNr)
    .Value = Trim(.Value)
    'Debug.Print "Check_Range " & ParName(ParNr - 1) & ": " & .Value
    Select Case TypA(ParNr - 1)
      Case "":        ' Normal Numeric parameter
                      If Check_Limit_to_MinMax(ParNr, .Value) = False Then Exit Function
      Case "Time":    ' time could have a tailing "Min", "Sek", "sek", "Sec", "sec", "Ms", "ms"
                      If IsNumeric(.Value) Then
                           .Value = Int(.Value)
                           If Check_Limit_to_MinMax(ParNr, .Value) = False Then Exit Function
                      Else ' The parameter is NOT numeric
                           If Check_Time_String(ParNr) = False Then Exit Function
                      End If
                      .Value = Replace(.Value, ",", ".") ' Replace the german comma
      Case "Var":     ' Check the variable name
                      ShowErr = (.Value = "")
                      If Not ShowErr Then
                         If left(.Value, Len("#InCh")) <> "#InCh" Then _
                            ShowErr = Not left(.Value, 1) Like "[_a-zA-Z]"
                      End If
                      If ShowErr Then
                         MsgBox Replace(Get_Language_Str("Fehler: Der Parameter '#1#' muss einen gültigen Variablennamen enthalten"), _
                                "#1#", VarLabel), vbCritical, Get_Language_Str("Ungültiger Variablenname")
                         .setFocus
                         Exit Function
                      End If
      Case "Txt":     ' No Check at the moment
                      Debug.Print "ToDo: Check parameter typ '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "Mode":    ' No Check at the moment
                      Debug.Print "ToDo: Check parameter typ '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "List":    ' No Check at the moment
                      Debug.Print "ToDo: Check parameter typ '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "PinList": ' No Check at the moment
                      Debug.Print "ToDo: Check parameter typ '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "PinNr":   ' No Check at the moment
                      Debug.Print "ToDo: Check parameter typ '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "Logic":   ' No Check at the moment
                      Debug.Print "ToDo: Check parameter typ '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "OutList": ' No Check at the moment
                      Debug.Print "ToDo: Check parameter typ '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "InpVar":  ' No Check at the moment
                      Debug.Print "ToDo: Check parameter typ '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "GVarNr":  ' No Check at the moment
                      Debug.Print "ToDo: Check parameter typ '" & TypA(ParNr - 1) & "'" ' ToDo
      Case "RGB", _
           "RGB2": '  Color lists
                      val = Check_RGB_List(TypA(ParNr - 1), Replace(Trim(Replace_Multi_Space(Replace(.Value, ",", " "))), " ", ", "), ParNr)
                      Check_Par_with_ErrMsg = (val <> "")
                      Exit Function
      Case "CmpMod":  ' Compare mode for the LED_to_Var function
                      If InStr(" " & L2V_COM_OPERATORS & " ", " " & .Value & " ") = 0 Then
                         MsgBox Replace(Replace(Get_Language_Str("Fehler: Der Parameter '#1#' muss eine der folgenden Vergleichsoperatoren enthalten:" & vbCr & _
                                "  #2#"), "#1#", VarLabel), "#2#", L2V_COM_OPERATORS), vbCritical, _
                                Get_Language_Str("Ungültiger Vergleichsoperator")
                         Exit Function
                      End If
      Case "MB_LED":  ' Mainboard LED Number
                      If (IsNumeric(.Value) And (.Value < 0 Or .Value > 16)) Or _
                         (Not IsNumeric(.Value) And InStr(" " & MB_LED_NR_STR & " ", " " & .Value & " ") = 0) Then  ' 21.10.20: Jürgen added: (not IsNumeric(.Value) and
                         MsgBox Replace(Replace(Get_Language_Str("Fehler: Der Parameter '#1#' muss eine der folgenden Werte enthalten:" & vbCr & _
                                "  #2#"), "#1#", VarLabel), "#2#", "0 - 16, D2 - D5, D7 - D13, A0 - A5"), vbCritical, _
                                Get_Language_Str("Ungültige LED Bezeichnung")
                         Exit Function
                      End If
      Case Else:      MsgBox "Internal error: Unknown parameter Typ: '" & TypA(ParNr - 1) & "'", vbCritical, "Internal Error"
                      EndProg
    End Select
    val = .Value
  End With
  Check_Par_with_ErrMsg = True
End Function


'---------------------------------------
Private Sub LED_Channel_TextBox_Change()                                    ' 27.04.20:
'---------------------------------------
  If Show_Channel_Type = CHAN_TYPE_LED Then
    LimmitActivInput LED_Channel_TextBox, 0, LED_CHANNELS - 1
  ElseIf Show_Channel_Type = CHAN_TYPE_SERIAL Then
    LimmitActivInput LED_Channel_TextBox, 0, SERIAL_CHANNELS - 1
  End If
End Sub


Private Sub Par10Select_Change()
    Par10.Value = Par10Select.Value
End Sub

Private Sub Par11Select_Change()
    Par11.Value = Par11Select.Value
End Sub

Private Sub Par12Select_Change()
    Par12.Value = Par12Select.Value
End Sub

Private Sub Par13Select_Change()
    Par13.Value = Par13Select.Value
End Sub

Private Sub Par14Select_Change()
    Par14.Value = Par14Select.Value
End Sub

Private Sub Par1Select_Change()
    Par1.Value = Par1Select.Value
End Sub

Private Sub Par2Select_Change()
    Par2.Value = Par2Select.Value
End Sub

Private Sub Par3Select_Change()
    Par3.Value = Par3Select.Value
End Sub

Private Sub Par4Select_Change()
    Par4.Value = Par4Select.Value
End Sub

Private Sub Par5Select_Change()
    Par5.Value = Par5Select.Value
End Sub

Private Sub Par6Select_Change()
    Par6.Value = Par6Select.Value
End Sub

Private Sub Par7Select_Change()
    Par7.Value = Par7Select.Value
End Sub

Private Sub Par8Select_Change()
    Par8.Value = Par8Select.Value
End Sub

Private Sub Par9Select_Change()
    Par9.Value = Par9Select.Value
End Sub

' Scrollbar:
' ~~~~~~~~~~
' Wir können nicht den Scrollbar der Textboy verwenden weil diese nicht aktiv ist ;-(
' Darum ist ein eigener Balken daneben platziert. Dieser muss aber von Hand angesteuert werden.
'
' - Die Höhe des Scroll handles kann über ScrollBar1.LargeChange gesetzt werden. Diese kann aber
'   maximal die halbe Höhe haben
' - Die .LineCount Eigenschaft der Textbox kann nur verwendet werden wenn sie den Fokus hat.
' - Die Position der Textbox wird über SelStart bestimmt. Der Wert wird Buchstabenweise gesetzt.
'   Der Wert müsste immer um die Länge einer Zeile erhöht werden.
'   Aber das ist kompliziert

#If False Then ' 04.11.21: Disabled
'------------------------------
Private Sub ScrollBar1_Change()
'------------------------------
' ToDo: Die Steuerung ist noch nicht gut
' - Slider Size anpassen
' - Die Description_TextBox sollte nicht editierbar sein
'
  'Debug.Print ScrollBar1.Value
  Dim LinesDialog As Long
  LinesDialog = 5
  With Description_TextBox
    .setFocus
    #If 0 Then
    If .LineCount > LinesDialog Then
        ScrollBar1.Min = 0
        ScrollBar1.Max = .LineCount - LinesDialog - 1
        ScrollBar1.LargeChange = .LineCount - LinesDialog - 1 ' Height of the bar (must be <= Max - Min)
    End If
    #End If
    .SelStart = .TextLength / .LineCount * ScrollBar1.Value
    Debug.Print "Lines:" & .LineCount & " MaxLength:" & .MaxLength & " TextLength:" & .TextLength & " ScrollBar1.Value:" & ScrollBar1.Value & " SelStart:" & .SelStart
  End With
  ScrollBar1.setFocus
End Sub
#End If

'-----------------------------------------------
Public Sub MouseWheel(ByVal lngRotation As Long)
'-----------------------------------------------
' Process the mouse wheel changes
  'Debug.Print "MouseWheel" & lngRotation ' Debug
  Dim ctlCurrentControl
  Set ctlCurrentControl = Me.ActiveControl
  With Description_TextBox
    .setFocus
    #If 1 Then ' Simulate keys
        If lngRotation < 0 Then
              Application.SendKeys "{HOME}{DOWN}{DOWN}{DOWN}" ' {HOME} to move the cursor to the start of the line
        Else: Application.SendKeys "{HOME}{UP}{UP}{UP}"
        End If
    #Else
        Dim ScrollChars As Long
        If .LineCount > 0 Then
          ScrollChars = 3 * Len(.Value) / .LineCount ' Not Optimal, should be the size of 3 lines
        End If
        If lngRotation < 0 Then
              .SelStart = .SelStart + ScrollChars
        Else: .SelStart = Application.Max(.SelStart - ScrollChars, 0)
        End If
    #End If
  End With
  DoEvents
  ctlCurrentControl.setFocus
End Sub

'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  MinFormHeight = 0                                                         ' 05.11.21: Juergen form resize feature
  MinFormWidth = 0
  Restore_Pos_or_Center_Form Me, OtherForm_Pos
End Sub

'-----------------------------------------------------------------------------------------------------------
' 06.11.21: Juergen make form resizeable feature
Private Sub UserForm_Resize()
  If MinFormHeight <> 0 And MinFormWidth <> 0 Then
    If Me.Height < MinFormHeight Then Me.Height = MinFormHeight
    If Me.Width < MinFormWidth Then Me.Width = MinFormWidth
    
    Dim DiffHeight As Long, DiffWidth As Long
    DiffHeight = Me.Height - CurHeight
    DiffWidth = Me.Width - CurWidth
    Dim c
    For Each c In Me.Controls
      If c.Name = Description_TextBox.Name Then
         c.Height = c.Height + DiffHeight
         c.Width = c.Width + DiffWidth
      ElseIf c.Name = ScrollBar1.Name Then
         c.Height = c.Height + DiffHeight
         c.left = c.left + DiffWidth
      ElseIf c.Name = OK_Button.Name Or c.Name = Abort_Button.Name Then
         c.top = c.top + DiffHeight
         c.left = c.left + DiffWidth
      ElseIf c.Parent.Name <> LED_Kanal_Frame.Name Then  ' Don't move the elements in LED_Kanal_Frame ' 10.11.21: Hardi
          c.top = c.top + DiffHeight
      End If
    Next
    CurWidth = Me.Width
    CurHeight = Me.Height
  End If
End Sub

'-----------------------------------------------------------------------------------------------------------
' 06.11.21: Juergen make form resizeable feature
Private Sub Resize_Description(heightCorr As Integer)
    Dim c
    For Each c In Me.Controls
      If c.Name <> Description_TextBox.Name And c.Name <> ScrollBar1.Name Then
         If c.Parent.Name <> LED_Kanal_Frame.Name Then  ' Don't move the elements in LED_Kanal_Frame ' 10.11.21: Hardi
            c.top = c.top + heightCorr
         End If
      End If
    Next
    Description_TextBox.Height = Description_TextBox.Height + heightCorr
    ScrollBar1.Height = ScrollBar1.Height + heightCorr
    Me.Height = Me.Height + heightCorr
End Sub

'-----------------------------------------------------------------------------------------------------------
' 06.11.21: Juergen make form resizeable feature
'
' Written: August 02, 2010
' Author:  Leith Ross
' Summary: Makes the UserForm resizable by dragging one of the sides. Place a call
'         to the macro MakeFormResizable in the UserForm'
' from  https://www.mrexcel.com/board/threads/resize-a-userform.485489/
Private Sub MakeFormResizable(oForm As Object)


  Dim lStyle As Long
  Dim RetVal
  
#If VBA7 Then
   Dim hWnd  As LongPtr
#Else
   Dim hWnd As Long
#End If

  hWnd = FindWindow("ThunderDFrame", oForm.Caption)
  
  Const WS_THICKFRAME = &H40000
  Const GWL_STYLE As Long = (-16)
  
  'Get the basic window style
  lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME
 
  'Set the basic window styles
  RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)
    
  'Clear any previous API error codes
  SetLastError 0
End Sub

'-----------------------------------------------------------------------------------------------------------
Private Function Get_RGB_from_OldParams(OldParams() As String, ByVal StartNr As Long, cnt As Long) As String
'-----------------------------------------------------------------------------------------------------------
  Dim Nr As Long, Res As String
  For Nr = StartNr To StartNr + cnt - 1
      Res = Res & Trim(OldParams(Nr)) & " "
  Next Nr
  Get_RGB_from_OldParams = Res
End Function


'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Show_UserForm_Other(ByVal par As String, ByVal Name As String, Description As String, LedChannels As Long, Show_Channel As Byte, LED_Channel As Long, Def_Channel As Long)  ' 27.04.20: Added: LED_Channel and Def_Channel
'-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  FuncName = Name
  Show_Channel_Type = Show_Channel
  'Const CNames = "Val0 Val1 Period Duration Timeout DstVar MinTime MaxTime Par1"
  If Description = "" Then
        Description_TextBox = Get_Language_Str("Noch keine Beschreibung zur Funktion '") & Name & Get_Language_Str("' vorhanden ;-(")
  Else: Description_TextBox = Description
  End If

  Description_TextBox.setFocus ' To show the scroll bar                     ' 04.11.21:
  Description_TextBox.SelStart = 0
  
  If Len(Description_TextBox) >= 500 Then                                   ' 05.11.21: Juergen make descrition filed bigger in case of longer texts
    Resize_Description 100
  End If
  
  Me.Caption = Get_Language_Str("Parametereingabe der '") & Name & Get_Language_Str("' Funktion")

  'Debug.Print                         ' Debug
  'Debug.Print Name & " (" & Par & ")" ' Debug

  '*** Hide all entrys in the dialog which are not needed ***
  ParList = Split(par, ",")

  ' Hide CX selection if it's not used
  If Not Is_Contained_in_Array("Cx", ParList) And Not Is_Contained_in_Array("B_LED_Cx", ParList) Then
       Hide_and_Move_up Me, "LED_Kanal_Frame", "Par1"
  Else ' Prepare the Cx selection
       Select Case LedChannels
          #If 1 Then ' deactivated because then the other functions like "Const" can't use more than one LED.    17.04.20:
                     ' On the other hand it's o.K. if the PushButton function could use more than one LED,
                     ' => It's enabled again
          Case -1:  ' Funktions with a single LED like PushButton_w_LED_BL_0_2    ' 13.04.20:
                    OptionButton_All.Enabled = False ' can't access all LEDs
                    OptionButton_12.Enabled = False
                    OptionButton_23.Enabled = False
                    OptionButton_C1 = True  ' Default value                 ' 06.05.20:
          Case -2:  ' Funktions with two LED like Herz_BiRelais             ' 19.05.20:
                    OptionButton_All.Enabled = False ' can't access all LEDs
                    OptionButton_C1 = True  ' Default value
          #End If
          Case 2:   ' Functions with two single LEDs like the Andreaskreuz
                    OptionButton_All.Enabled = False ' can't access all LEDs
                    If OptionButton_All Then OptionButton_23 = True
       End Select
  End If
  
  
  Dim OldCmdLine As String, UseOldParams As Boolean, OldParams() As String
  OldCmdLine = Cells(ActiveCell.Row, Config__Col)
  If Len(OldCmdLine) > Len(Name) And left(OldCmdLine, Len(Name) + 1) = Name & "(" Then   ' 18.04.20: Added +1 and "(" to prevent mix-up of MonoFlop2 <-> MonoFlop
     UseOldParams = True
     OldParams = Split(Trim(Replace(Mid(OldCmdLine, Len(Name) + 2), ")", "")), ",")
  End If


  LED_Channel_TextBox.Visible = Show_Channel <> CHAN_TYPE_NONE                       ' 07.10.21: Jürgen Add serial Channel support
  LED_Channel_Label.Visible = LED_Channel_TextBox.Visible
  If Show_Channel = CHAN_TYPE_SERIAL Then                                            ' 07.10.21: Jürgen Add serial Channel support
        LED_Channel_TextBox = LED_Channel
        LED_Channel_Label.Caption = Get_Language_Str("Sound Kanal")
  ElseIf Show_Channel = CHAN_TYPE_LED Then
     If UseOldParams Then
           LED_Channel_TextBox = LED_Channel
     Else: LED_Channel_TextBox = Def_Channel
     End If
  End If

  ' Add parameters
  Dim p As Variant, UsedParNr As Long, Nr As Long, UsedArgNr As Long
  Dim SelectValues As Variant
  
  OK_Button.setFocus                                                        ' 07.10.21: Focus on the OK Button (Prior sometimes the Abort button was selected)
  For Each p In ParList
    p = Trim(p)
    If left(p, 1) <> "#" And InStr(" Cx B_LED_Cx ", " " & p & " ") = 0 Then
        If UsedParNr >= MAX_PAR_CNT Then
            MsgBox "Internal error: The number of parameters is to large in Show_UserForm_Other()"
            EndProg
        End If
        Dim Typ As String, Min As String, Max As String, Def As String, Opt As String, InpTxt As String, Hint As String, ParVal As String
        ParVal = ""                                                         ' 20.10.21: Hardi: Added because otherwise the default values are always equal than the prior
        Get_Par_Data p, Typ, Min, Max, Def, Opt, InpTxt, Hint
        If UseOldParams Then
            Select Case Typ
                Case "RGB":  ParVal = Get_RGB_from_OldParams(OldParams, Nr, 3): Nr = Nr + 2
                Case "RGB2": ParVal = Get_RGB_from_OldParams(OldParams, Nr, 3) & "  " & Get_RGB_from_OldParams(OldParams, Nr + 3, 3): Nr = Nr + 5
                Case Else:   ' Fill in the old value
                             If Nr <= UBound(OldParams) Then ParVal = OldParams(Nr)  ' 20.05.20: Added if()... to prevent crash if the number of parameters have been changed in a new version of the library  ' 07.06.20: Old "<" => "<="
            End Select
        End If
        ParVal = Trim(ParVal)
        TypA(UsedParNr) = Typ
        MinA(UsedParNr) = Min
        MaxA(UsedParNr) = Max
        ParName(UsedParNr) = p
        UsedParNr = UsedParNr + 1
        If OK_Button Is Me.ActiveControl Then
            Me.Controls("Par" & UsedParNr).setFocus ' 07.10.21:
        End If
        SelectValues = Null
        Me.Controls("Par" & UsedParNr) = ParVal
        Dim w As Variant
        w = DEFAULT_PAR_WIDTH
        If Opt <> "" Then                                                  ' 18.04.20:
            Dim OptionValue As Variant                                      ' 23.04.21
            For Each OptionValue In Split(Opt, ";")
                OptionValue = Trim(OptionValue)
                If (LCase(left$(OptionValue, 7)) = "values:") Then ' Value list for parameter found
                    SelectValues = Split(Trim(Mid$(OptionValue, 8)), ",") ' Values delimited by comma
                Else
                    Opt = OptionValue
                    Dim o As Variant, Parts() As String
                    For Each o In Split(Opt, " ")
                        Parts = Split(o, "=")
                        Select Case Parts(0)
                            Case "w":   w = val(Parts(1))
                                        Me.Controls("Par" & UsedParNr).Width = w
                                        With Me.Controls("LabelPar" & UsedParNr)
                                            .left = .left + w - DEFAULT_PAR_WIDTH
                                        End With
                            Case "Inv": ' Invert input                      ' 30.04.20: Added to invert the Skip0 parameter in the PushButton function
                                        Invers(UsedParNr - 1) = True        '           because its difficult to understand the negative logic of Skip0
                                        If ParVal = 0 Then                  '           But it's not used because Skip0 has been chanded to Use0 in the C code
                                              ParVal = 1
                                        Else: ParVal = 0
                                        End If
                                        Me.Controls("Par" & UsedParNr) = ParVal
                            Case "...", "…": ' Join all following parameters         ' 21.04.20:
                                        If UseOldParams Then
                                            Dim i As Long
                                            For i = UsedParNr + 1 To UBound(OldParams)
                                                ParVal = ParVal & ", " & Trim(OldParams(i))
                                            Next
                                            Me.Controls("Par" & UsedParNr) = ParVal
                                        End If
                            Case Else:  MsgBox "Internal Error: Unknown option '" & o & "' in parameter '" & InpTxt & "' in sheet '" & PAR_DESCR_SH & "'", vbCritical, "Internal Error"
                                        EndProg
                        End Select
                    Next o
                End If
            Next OptionValue
        End If
        Me.Controls("LabelPar" & UsedParNr) = InpTxt
        Me.Controls("LabelPar" & UsedParNr).ControlTipText = Hint
        Me.Controls("Par" & UsedParNr).ControlTipText = Hint
        Me.Controls("Par" & UsedParNr & "Select").ControlTipText = Hint     ' 11.10.21: Juergen
            
        Dim validListEntry As Boolean
        validListEntry = False                                              ' 20.10.21: Hardi: In case we have several List entries
        
        If (Not IsNull(SelectValues)) Then
            Dim SelectValue As Variant
            For Each SelectValue In SelectValues
                SelectValue = Trim(SelectValue)
                If (SelectValue <> "") Then
                    ' Debug.Print (SelectValue)
                    If SelectValue = ParVal Then validListEntry = True
                    With Me.Controls("Par" & UsedParNr & "Select")
                        .AddItem SelectValue
                    End With
                End If
            Next SelectValue
            Me.Controls("Par" & UsedParNr & "Select").Visible = True
            Me.Controls("Par" & UsedParNr).Width = w
            Me.Controls("Par" & UsedParNr).left = 12
            Me.Controls("Par" & UsedParNr & "Select").left = 12 + w + 4
            Me.Controls("Par" & UsedParNr & "Select").Width = w
            With Me.Controls("LabelPar" & UsedParNr)
                .left = 12 + 2 * w + 8
                .Width = 410 - (12 + 2 * w + 8)
            End With
        Else
            Me.Controls("Par" & UsedParNr & "Select").Visible = False
        End If
        
        If ParVal = "" Then
            ParVal = Def                                            ' 11.10.21: Juergen
            validListEntry = True
        End If
        Select Case Typ
            Case "List":
                Me.Controls("Par" & UsedParNr & "Select").Style = 2
                Me.Controls("Par" & UsedParNr).Enabled = False
                If Not validListEntry Then
                        MsgBox Replace(Replace(Replace(Get_Language_Str("Fehler: Der Parameter '#1#' hat einen ungültigen Wert #2#. Der Parameter wird auf den Standardwert #3# zurückgesetzt."), _
                    "#1#", p), "#2#", ParVal), "#3#", Def), vbCritical, Get_Language_Str("Parameter Fehler")
                    ParVal = Def
                End If
        End Select
        
        Me.Controls("Par" & UsedParNr & "Select") = ParVal
        Me.Controls("Par" & UsedParNr) = ParVal
      End If
      If (p = "Cx" Or p = "B_LED_Cx") Then
        If UseOldParams Then
              Set_OptionButton Trim(OldParams(Nr)), p = "B_LED_Cx"
        Else: OptionButton_C1.setFocus                              ' 07.10.21:
        End If
      End If
      Nr = Nr + 1
  Next p
  'Debug.Print "UsedParNr=" & UsedParNr ' Debug

  Hide_and_Move_up Me, "Par" & UsedParNr + 1, "Abort_Button" ' Hide the not needed controlls
  Dim c1 As Variant
  
  ' set focus to first visible and enabeld control                          ' 09.10.21 Jürgen
  Dim minIndex As Integer
  Dim minControl As Variant
  
  On Error Resume Next
  minIndex = 9999
  For Each c1 In Me.Controls
      If c1.Visible And c1.TabIndex > 0 And Not left(c1.Name, 13) = "OptionButton_" Then
        Dim isTabStop As Boolean
        isTabStop = False
        isTabStop = c1.TabStop
        'If isTabStop And c1.TabIndex > 2 And c1.TabIndex < minIndex Then     ' 06.11.21 Jürgen  First par not selected
        If isTabStop And c1.TabIndex >= 2 And c1.TabIndex < minIndex Then
            minIndex = c1.TabIndex
            Set minControl = c1
        End If
      End If
  Next
  If minIndex <> 9999 Then minControl.setFocus
  On Error GoTo 0

  Center_Form Me                                                            ' 18.01.20:
  
  'hook at end of initialisation code, because if an error and exception thrown AFTER HookFormScroll Excel will crash soon      ' 09.10.21 Juergen
  HookFormScroll Me, "Description_TextBox"   ' Initialize the mouse wheel scroll function
  CurWidth = Me.Width                                                       ' 05.11.21: Juergen form resize feature
  CurHeight = Me.Height
  MinFormHeight = Me.Height
  MinFormWidth = Me.Width
  MakeFormResizable Me
  Me.Show
End Sub

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

'------------------------------------------------------------------
Private Sub Set_OptionButton(val As String, Only_Single As Boolean)
'------------------------------------------------------------------
  Select Case val
    Case "C_ALL":
                  If Only_Single Then
                        OptionButton_C1 = True:  OptionButton_C1.setFocus    ' 07.10.21: Added ...setFocus to all lines
                  Else: OptionButton_All = True: OptionButton_All.setFocus
                  End If
    Case "C1":    OptionButton_C1 = True:        OptionButton_C1.setFocus
    Case "C2":    OptionButton_C2 = True:        OptionButton_C2.setFocus
    Case "C3":    OptionButton_C3 = True:        OptionButton_C3.setFocus
    Case "C12":
                  If Only_Single Then
                        OptionButton_C1 = True:  OptionButton_C1.setFocus
                  Else: OptionButton_12 = True:  OptionButton_12.setFocus
                  End If
    Case "C23":
                  If Only_Single Then
                        OptionButton_C2 = True:  OptionButton_C2.setFocus
                  Else: OptionButton_23 = True:  OptionButton_23.setFocus
                  End If
    Case Else:    OptionButton_All = True:       OptionButton_All.setFocus
                  MsgBox Get_Language_Str("Fehler beim Lesen der bestehenden Kanalbezeichnung '") & val & "'", vbCritical, Get_Language_Str("Unbekannte Kanalbezeichnung")
  End Select
End Sub


'-------------------------------------------------------------
Private Function Create_Result(ByRef Res As String) As Boolean
'-------------------------------------------------------------
' Return True if sucessfully checked all inputs
  Res = ""
  Dim p As Variant
  For Each p In ParList
      Dim val As Variant
      val = "Not Found"
      p = Trim(p)
      If left(p, 1) = "#" Then
           val = p
      Else
           If p = "Cx" Or p = "B_LED_Cx" Then
                val = Get_OptionButton_Res()
           Else ' Not a standard parameter
                Dim Nr As Long
                For Nr = 1 To MAX_PAR_CNT
                    If ParName(Nr - 1) = p Then
                       If Check_Par_with_ErrMsg(Nr, val) = False Then
                          Controls("Par" & Nr).setFocus
                          Exit Function
                       End If
                       If Invers(Nr - 1) Then
                          If val = 0 Then
                                val = 1
                          Else: val = 0
                          End If
                       End If
                       Exit For
                    End If
                Next
           End If
      End If
     If val = "Not Found" Then MsgBox Get_Language_Str("Fehler der Parameter '") & p & Get_Language_Str("' wurde nicht gefunden"), vbCritical, Get_Language_Str("Programm Fehler")
      Res = Res & val & ", "
  Next p
  Res = FuncName & "(" & DelLast(Res, 2) & ")"
  
  If LED_Channel_TextBox.Visible Then                                       ' 27.04.20:
     Res = Res & "$" & LED_Channel_TextBox
  End If
  
  Create_Result = True
End Function


'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  UnhookFormScroll ' Deactivate the mouse wheel scroll function
  Userform_Res = ""
  Store_Pos Me, OtherForm_Pos
  
  Unload Me ' Don't keep the entered data. Importand because the positions of the controlls and the visibility have been changed
End Sub

'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  If Create_Result(Userform_Res) Then
     UnhookFormScroll ' Deactivate the mouse wheel scroll function
     Store_Pos Me, OtherForm_Pos

     Unload Me ' Don't keep the entered data. Importand because the positions of the controlls and the visibility have been changed
  End If
End Sub








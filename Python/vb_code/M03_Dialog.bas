Attribute VB_Name = "M03_Dialog"
Option Explicit

' Dialog guided input

' - Kurze Erklärung am Anfang
'   mit Auswahl der Zeile.
'   Wenn bereits Daten in der Zeile sind, dann wird mit der ausgewählten Spalte weitergemacht
' - Abfrage ob der Effekt von DCC gesteuert werden soll.
'   - Adresse
' -
Private Ask_Input_NextRow As Boolean
Private Input_NextRow As Boolean

'-------------------------------
Public Sub Dialog_Guided_Input()
'-------------------------------
  Make_sure_that_Col_Variables_match
  
  If First_Change_in_Line(ActiveCell) Then
     UserForm_DialogGuide1.Show
    
     While UserForm_DialogGuide1.IsActive
       DoEvents
     Wend
     If DialogGuideRes = vbAbort Then Exit Sub
     
     Dim i As Long
     For i = 1 To 5
       DoEvents ' Redraw the screen
       Sleep 50 ' Time to update the display
     Next
         
     Do
         Input_NextRow = False
         Ask_Input_NextRow = True
         Ask_External_Control
         
         If Input_NextRow Then
            Debug.Print "ToDo: Prüfen of die nächste Zeile leer ist und geg. eine Zeile einfügen"
         End If
     Loop While Input_NextRow
  Else
     Ask_Input_NextRow = False
     Dim r As Long
     r = ActiveCell.Row
     Select Case ActiveCell.Column
       Case DCC_or_CAN_Add_Col, _
            SX_Channel_Col: ' DCC adress oder SX Channel
                            Dim SX_DataAvailable As Boolean
                            If SX_Bitposi_Col > 0 Then SX_DataAvailable = Cells(r, SX_Bitposi_Col) <> ""
                            If Cells(r, DCC_or_CAN_Add_Col + SX_Channel_Col) = "" And Not SX_DataAvailable Then
                                  Ask_External_Control
                            Else: Input_Address
                            End If
       Case Inp_Typ_Col:    Input_Typ
       Case SX_Bitposi_Col: Input_BitPos
       Case Start_V_Col:    Input_Start_Val
       Case Descrip_Col:    Input_Description
       Case Dist_Nr_Col, _
            Conn_Nr_Col:    ' Distributor or connector number
                            Input_Connector
       Case Config__Col, _
            MacIcon_Col, _
            LanName_Col:    ' Macro selection
                            SelectMacros
       Case Else            ' Unsupported column
                            If ActiveCell.Column > Config__Col Then
                                  MsgBox Get_Language_Str("Die ausgewählte Spalte sollte nur von erfahrenen Benutzern verändert werden." & vbCr & _
                                         "Es existiert keine Dialog gestützte Eingabe."), vbInformation, Get_Language_Str("Spalte sollte nur von Experten verändert werden")
                            Else: MsgBox Get_Language_Str("Für die Ausgewählte Spalte existiert noch kein Dialog"), vbInformation, Get_Language_Str("Kein Dialog vorhanden")
                            End If
     End Select
  End If
End Sub

'--------------------------------
Public Sub Ask_External_Control()
'--------------------------------
Make_sure_that_Col_Variables_match
  Select Case MsgBox(Get_Language_Str("Soll die LED Gruppe über ") & Page_ID & Get_Language_Str(" gesteuert werden?" & vbCr & _
                     vbCr & _
                     "  Ja:     Der Effekt kann über eine Zentrale geschaltet werden." & vbCr & _
                     "           Im Folgenden wird die Adresse zur Steuerung der" & vbCr & _
                     "           Funktion abgefragt. Das ist z.B. bei einem Haus oder" & vbCr & _
                     "           einem Signal sinnvoll." & vbCr & _
                     "  Nein: Der Effekt ist dauerhaft aktiv. Das kann man z.B. bei" & vbCr & _
                     "           einer Ampel auswählen. Die Steuerung über ") & Page_ID & vbCr & Get_Language_Str( _
                     "           kann auch nachträglich aktiviert werden."), _
                     vbQuestion + vbYesNoCancel, Get_Language_Str("Steuerung über ") & Page_ID & "?") ' 03.09.19: Corected with problem detected by Alf
     Case vbYes: Input_Address
     Case vbNo:  Input_Description
     Case vbCancel: Exit Sub
  End Select
End Sub


'-------------------------
Public Sub Input_Address()
'-------------------------
  Dim Txt As String, This_Addr_Channel As String, Addr_Channel As String, MinVal As Long, MaxVal As Long, Adresses_Channels As String
  If Page_ID = "Selectrix" Then
        Txt = Get_Language_Str(" Kanal eingeben"):   This_Addr_Channel = Get_Language_Str("Der Kanal"):   Addr_Channel = Get_Language_Str("Kanal"):   MinVal = 0: MaxVal = 99:    Adresses_Channels = Get_Language_Str("Kanäle")    ' 24.02.20: Added some more "Get_Language_Str"
  Else: Txt = Get_Language_Str(" Adresse eingeben"): This_Addr_Channel = Get_Language_Str("Die Adresse"): Addr_Channel = Get_Language_Str("Adresse"): MinVal = 1: MaxVal = 10240: Adresses_Channels = Get_Language_Str("Adressen")  '   "
  End If
  If Page_ID = "CAN" Then MaxVal = 65535 ' ??
  
  Dim Inp As String, Valid As Boolean
  Inp = Get_First_Number_of_Range(ActiveCell.Row, DCC_or_CAN_Add_Col + SX_Channel_Col)
  Do
    Inp = InputBox(Get_Language_Str("Bitte ") & Page_ID & Txt & " [" & MinVal & ".." & MaxVal & "]" & vbCr & _
                   vbCr & _
                   This_Addr_Channel & Get_Language_Str(" muss bei der Zentrale zur Steuerung der Funktion angegeben werden." & vbCr & _
                   vbCr & _
                   "Achtung: Bei manchen Funktionen werden mehrere ") & Adresses_Channels & Get_Language_Str(" belegt. " & _
                   "Das Programm ergänzt den Bereich automatisch (Beispiel: 23 - 24)" & vbCr & _
                   "Es muss nur der Startwert ohne '- 24' eingegeben werden.") & vbCr & _
                   vbCr & _
                   Page_ID & " " & Addr_Channel & ": ", Page_ID & Txt, Default:=Inp)
    
    'Debug.Print "Res='" & Inp & "'" ' Debug
    If InStr(Inp, "-") > 1 Then Inp = Left(Inp, InStr(Inp, "-"))
    If IsNumeric(Inp) Then Valid = val(Inp) >= MinVal And val(Inp) <= MaxVal And Int(Inp) = Inp
    If Inp <> "" And Not Valid Then
       BeepThis2 "Windows Balloon.wav"
       Show_Status_for_a_while Get_Language_Str("Falsche Eingabe. ") & This_Addr_Channel & Get_Language_Str(" muss zwischen ") & MinVal & Get_Language_Str(" und ") & MaxVal & Get_Language_Str(" liegen. ")
    End If
  Loop Until Inp = "" Or Valid
  Show_Status_for_a_while ""
  
  If Valid Then
     With Cells(ActiveCell.Row, DCC_or_CAN_Add_Col + SX_Channel_Col)
       .Value = val(Inp)
       Application.EnableEvents = False ' Prevent opening the Typ Dialog per Event
       .Offset(0, 1).Select
       Application.EnableEvents = True
     End With
     Select Case Page_ID                                                    ' 15.10.20: Adapted
       Case "Selectrix": Input_BitPos
       Case "CAN":       Input_Typ
       Case "DCC":       Cells(ActiveCell.Row, Inp_Typ_Col).Offset(0, 1).Select ' 15.10.20: Added
                         Input_Start_Val
       Case Else:        MsgBox "Internal error in 'Input_Address()': Unknown Page_ID '" & Page_ID & "'", vbCritical, "Internal Error"
     End Select

  End If
End Sub

'------------------------
Public Sub Input_BitPos()
'------------------------
  Dim Inp As String, Valid As Boolean
  Inp = Cells(ActiveCell.Row, SX_Bitposi_Col)
  Do
    Inp = InputBox(Get_Language_Str("Bitte die Bitposition eingeben [1..8]" & vbCr & _
                   vbCr & _
                   "Die Bitposition muss bei der Zentrale zur Steuerung der Funktion angegeben werden." & vbCr & _
                   "Achtung: Bei manchen Funktionen werden mehrere Bits belegt. Die Eingabe definiert das erste benutzte Bit." & vbCr & _
                   vbCr & _
                   "Bitposition: "), Page_ID & Get_Language_Str("Bitposition eingeben"), Inp)
    
    'Debug.Print "Res='" & Inp & "'" ' Debug
    If IsNumeric(Inp) Then Valid = val(Inp) >= 1 And val(Inp) <= 8 And Int(Inp) = Inp
    If Inp <> "" And Not Valid Then
       BeepThis2 "Windows Balloon.wav"
       Show_Status_for_a_while Get_Language_Str("Falsche Eingabe. Die Bitposition muss zwischen 1 und 8 liegen. ")
    End If
  Loop Until Inp = "" Or Valid
  
  If Valid Then
     With Cells(ActiveCell.Row, SX_Bitposi_Col)
       .Value = val(Inp)
       Application.EnableEvents = False ' Prevent opening the Typ Dialog per Event
       .Offset(0, 1).Select
       Application.EnableEvents = True
     End With
     Input_Typ
  End If
End Sub

'---------------------
Public Sub Input_Typ()
'---------------------
  Select_Typ_by_Dialog ActiveCell
  If Userform_Res <> "" Then
     Cells(ActiveCell.Row, Inp_Typ_Col).Offset(0, 1).Select
     Input_Start_Val
  End If
End Sub

'---------------------------
Public Sub Input_Start_Val()
'---------------------------
  Const MinVal = 1
  Const MaxVal = 255
  'Debug.Print "Inp_Typ_Col=" & Inp_Typ_Col
  Dim Valid As Boolean, Inp As Variant
  Inp = ActiveCell
  Do
    Inp = InputBox(Get_Language_Str("Startwert des Eingangs eingeben" & vbCr & _
                   vbCr & _
                   "Der Startwert bestimmt das Verhalten nach dem Einschalten in Verbindung mit DCC, " & _
                   "CAN oder Selectrix. " & vbCr & _
                   "Normalerweise sind die Funktionen beim Start deaktiviert. " & _
                   "Erst wenn der erste ") & Page_ID & Get_Language_Str(" Einschaltbefehl von der Zentrale kommt wird " & _
                   "die Zeile aktiviert. " & vbCr & _
                   "Wenn eine bestimmte Funktion bereits beim Einschalten der " & _
                   "Anlage einen definierten Wert haben soll kann das über den " & _
                   "Startwert vorgegeben werden. Die meisten Funktionen haben einen Eingang mit dem sie " & _
                   "Ein- oder Ausgeschaltet werden. Hier wird eine 1 zum Einschalten angegeben." & vbCr & _
                   "Bei Funktionen mit mehreren Eingängen (z.B. Signale) ist der Wert ist Bitkodiert. " & _
                   "Hier wird der erste Eingang mit einer 1, zweite Eingang mit einer 2 und der dritte Eingang " & _
                   "mit einer 4 aktiviert." & vbCr & _
                   vbCr & _
                   "Startwert:  (Keine Eingabe wenn nicht benötigt)"), Get_Language_Str("Definition des Startwerts"), Inp)
    
    If IsNumeric(Inp) Then Valid = val(Inp) >= MinVal And val(Inp) <= MaxVal And Int(Inp) = val(Inp)
    
    If Inp <> "" And Not Valid Then
       BeepThis2 "Windows Balloon.wav"
       Show_Status_for_a_while Get_Language_Str("Falsche Eingabe. ") & Inp & Get_Language_Str(" muss zwischen ") & MinVal & Get_Language_Str(" und ") & MaxVal & Get_Language_Str(" liegen. ")
    End If
    
  Loop Until Inp = "" Or Valid
  ActiveCell = Inp
  Show_Status_for_a_while ""
  ActiveCell.Offset(0, 1).Select
  Input_Description
End Sub

'-----------------------------
Public Sub Input_Description()
'-----------------------------
  Dim Res As String
  Res = UserForm_Description.ShowForm(ActiveCell.Value)
  If Res <> "<Abort>" Then
     With Cells(ActiveCell.Row, Descrip_Col)
       .Value = Res
       .Offset(0, 1).Select
     End With
     Input_Connector
  End If
End Sub


'---------------------------
Public Sub Input_Connector()
'---------------------------
  Dim r As Long, Res As Boolean
  r = ActiveCell.Row
  If UserForm_Connector.Start(Cells(r, Dist_Nr_Col), Cells(r, Conn_Nr_Col)) Then
     Application.EnableEvents = False
     Cells(r, Conn_Nr_Col + 1).Select
     Application.EnableEvents = True
    
     If MsgBox(Get_Language_Str("Im folgenden Dialog wird die Funktion ausgewählt welche mit dieser Zeile verknüpft ist. " & _
               "Je nach Funktion müssen weitere Parameter angegeben werden."), vbOKCancel, Get_Language_Str("Fast geschafft")) = vbOK Then
        Debug.Print "Start SelectMacros() from Input_Connector"
        If SelectMacros() Then
           If Ask_Input_NextRow Then
              Ask_Input_NextRow = False
              Cells(r + 1, DCC_or_CAN_Add_Col + SX_Channel_Col).Select
              If MsgBox(Get_Language_Str("Eingabe einer weiteren Zeile?"), vbYesNo + vbQuestion, Get_Language_Str("Nächste Zeile Eingeben")) = vbYes Then
                 Input_NextRow = True
              End If
           End If
        End If
     End If
  End If
End Sub

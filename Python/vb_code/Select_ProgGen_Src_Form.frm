VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Select_ProgGen_Src_Form 
   Caption         =   "Quellzeile im Prog_Generator ausw�hlen"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5085
   OleObjectBlob   =   "Select_ProgGen_Src_Form.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "Select_ProgGen_Src_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Callback_Proc As String

'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
  If Callback_Proc <> "" Then Run Callback_Proc, False, "", "", 0
End Sub

'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  Dim Message As String
  If ActiveWorkbook.Name <> ThisWorkbook.Name Then
       Message = Get_Language_Str("Fehler: Die Zeile muss im Prog_Generator Excel Programm ausgew�hlt werden")
  Else
       If Is_Data_Sheet(ActiveSheet) = False Then
          Message = Get_Language_Str("Fehler: Das Ausgew�hlte Excel Blatt ist keine g�ltige Prog_Generator Konfigurationsseite")
       End If
  End If
  
  If Message = "" Then
     If Not Selected_Row_Valid() Then
          Message = Get_Language_Str("Fehler: Die Ausgew�hlte Zeile ist nicht innerhalb des g�ltigen Bereichs")
     Else
          Make_sure_that_Col_Variables_match
          If InStr(Cells(ActiveCell.Row, Config__Col), "PatternT") = 0 Then
             Message = Get_Language_Str("Achtung: Die ausgew�hlte Zeile enth�lt kein Makro welches vom Pattern_Generator importiert werden kann.")
          End If
     End If
  End If
  
  If Message <> "" Then
       Select Case MsgBox(Message, vbCritical + vbRetryCancel, Get_Language_Str("Fehler bei der Auswahl der Zielzeile"))
          Case vbCancel: Me.Hide
          Case vbRetry:  ' Retry
                         If ActiveWorkbook.Name <> ThisWorkbook.Name Then ThisWorkbook.Activate
                         Exit Sub
       End Select
  Else
       Me.Hide
       If Callback_Proc <> "" Then Run Callback_Proc, True, Cells(ActiveCell.Row, Descrip_Col), _
                                                            Cells(ActiveCell.Row, Config__Col), _
                                                            ActiveCell.Row
  End If
End Sub

'---------------------------------------------
Public Sub Check_and_Start(Callback As String)
'---------------------------------------------
  Callback_Proc = Callback
  Me.Show
End Sub

'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
' Center the dialog if it's called the first time.
' On a second call without me.close this function is not called
' => The last position ist used
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Center_Form Me
End Sub

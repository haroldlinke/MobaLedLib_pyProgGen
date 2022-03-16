VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Select_ProgGen_Dest_Form 
   Caption         =   "Zielzeile im Prog_Generator auswählen"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5085
   OleObjectBlob   =   "Select_ProgGen_Dest_Form.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "Select_ProgGen_Dest_Form"
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
  If Callback_Proc <> "" Then Run Callback_Proc, False, False, False
End Sub

'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  Dim Message As String
  If ActiveWorkbook.Name <> ThisWorkbook.Name Then
       Message = Get_Language_Str("Fehler: Die Zeile muss im Prog_Generator Excel Programm ausgewählt werden")
  Else
       If Is_Data_Sheet(ActiveSheet) = False Then
          Message = Get_Language_Str("Fehler: Das Ausgewählte Excel Blatt ist keine gültige Prog_Generator Konfigurationsseite")
       End If
  End If
  
  If Message = "" Then
     If Not Selected_Row_Valid() Then _
        Message = Get_Language_Str("Fehler: Die Ausgewählte Zeile ist nicht innerhalb des gültigen Bereichs")
  End If
  
  If Message <> "" Then
       Select Case MsgBox(Message, vbCritical + vbRetryCancel, "Fehler bei der Auswahl der Zielzeile")
          Case vbCancel: Me.Hide
          Case vbRetry:  ' Retry
                         If ActiveWorkbook.Name <> ThisWorkbook.Name Then ThisWorkbook.Activate
                         Exit Sub
       End Select
  Else
       Me.Hide
       If Callback_Proc <> "" Then
          Run Callback_Proc, True, SendToArduino_CheckBox, GoBack_CheckBox
          
          ThisWorkbook.Activate                                             ' 20.10.21: Add the Pattern Configurator Icon
          With ActiveSheet.Cells(ActiveCell.Row, Config__Col)
               If .Value <> "" Then
                  FindMacro_and_Add_Icon_and_Name .Value, ActiveCell.Row, ThisWorkbook.ActiveSheet, False
               End If
          End With
       End If
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

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Protokoll_Auswahl 
   Caption         =   "Protokoll Auswahl"
   ClientHeight    =   4284
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   OleObjectBlob   =   "UserForm_Protokoll_Auswahl.frx":0000
End
Attribute VB_Name = "UserForm_Protokoll_Auswahl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
 Me.Hide ' no "Unload Me" to keep the entered data and dialog position
End Sub


'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  Me.Hide
  If DCC_Button Then Sheets("DCC").Select
  If Selectrix_Button Then Sheets("Selectrix").Select
  If CAN_Button Then Sheets("CAN").Select
End Sub

'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Center_Form Me
End Sub


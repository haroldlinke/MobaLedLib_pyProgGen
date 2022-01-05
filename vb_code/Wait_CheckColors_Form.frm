VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Wait_CheckColors_Form 
   Caption         =   "Warte auf das Ende des Farbtest Programms"
   ClientHeight    =   4470
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5985
   OleObjectBlob   =   "Wait_CheckColors_Form.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "Wait_CheckColors_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
  Close_CheckColors
End Sub

'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Center_Form Me                                                            ' 06.05.20:
End Sub


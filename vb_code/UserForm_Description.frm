VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Description 
   Caption         =   "Eingabe der Beschreibung"
   ClientHeight    =   5796
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8370
   OleObjectBlob   =   "UserForm_Description.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm_Description"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  TextBox = "<Abort>"
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
End Sub


'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  'Userform_Res = TextBox.Value
  Me.Hide
End Sub

'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & Me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Center_Form Me
End Sub

'------------------------------------------------------
Public Function ShowForm(ByVal Txt As String) As String
'------------------------------------------------------
  TextBox.setFocus
  TextBox.Value = Txt
  Me.Show
  ShowForm = TextBox
End Function

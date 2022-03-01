VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_SingleInput 
   ClientHeight    =   1695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15510
   OleObjectBlob   =   "UserForm_SingleInput.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm_SingleInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Abort_Button_Click()
'-------------------------------
  TextBox = "<Abort>"
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
End Sub


'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  Me.Hide
End Sub

'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  Change_Language_in_Dialog Me
  Center_Form Me
End Sub

'------------------------------------------------------
Public Function ShowForm(ByVal Caption As String, ByVal Label As String, ByVal Text As String) As String
'------------------------------------------------------
  TextBox.setFocus
  TextBox.Text = Text
  Label1.Caption = Label
  Me.Caption = Caption
  Me.Show
  ShowForm = TextBox.Text
End Function

Private Sub UserForm_QueryClose(CloseMode As Integer, Cancel As Integer)
    Abort_Button_Click
End Sub


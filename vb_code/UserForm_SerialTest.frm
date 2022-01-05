VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_SerialTest 
   Caption         =   "Serial Test"
   ClientHeight    =   3144
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4590
   OleObjectBlob   =   "UserForm_SerialTest.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm_SerialTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
  End_Serial_Test = True
  Me.Hide
End Sub

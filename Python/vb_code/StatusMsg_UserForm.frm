VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StatusMsg_UserForm 
   Caption         =   "Bitte etwas Geduld..."
   ClientHeight    =   1452
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3090
   OleObjectBlob   =   "StatusMsg_UserForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "StatusMsg_UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & me.Name & ":Load_All_Examples Initialize"
  Change_Language_in_Dialog Me
  Center_Form Me
End Sub

'----------------------------------
Public Sub Set_Label(Msg As String)
'----------------------------------
  Label = Msg
End Sub


'-------------------------------------------
Public Sub Set_ActSheet_Label(Txt As String)
'-------------------------------------------
  ActSheet_Label = Txt
  DoEvents
End Sub


'----------------------------------------------------
Public Sub ShowDialog(Label As String, Txt As String)
'----------------------------------------------------
  Set_Label Label
  Set_ActSheet_Label Txt
  Show
End Sub

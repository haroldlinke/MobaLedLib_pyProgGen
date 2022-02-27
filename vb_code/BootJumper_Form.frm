VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BootJumper_Form 
   Caption         =   "Installieren des schnellen Bootloaders"
   ClientHeight    =   5004
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   9090
   OleObjectBlob   =   "BootJumper_Form.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "BootJumper_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Res As Boolean

'--------------------------------
Private Sub Abbort_Button_Click()
'--------------------------------
  Res = False
  Me.Hide ' no "Unload Me" to keep position
End Sub

'-------------------------------
Private Sub Start_Button_Click()
'-------------------------------
  Res = True
  Me.Hide ' no "Unload Me" to keep position
End Sub


'--------------------------------------
Public Function ShowDialog() As Boolean
'--------------------------------------
  Res = False
  Me.Show
  ShowDialog = Res
End Function


'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me
  Center_Form Me
End Sub


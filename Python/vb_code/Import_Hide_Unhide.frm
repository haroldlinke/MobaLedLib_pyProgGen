VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Import_Hide_Unhide 
   Caption         =   "Auswahl der gewünschten Zeilen"
   ClientHeight    =   4572
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5985
   OleObjectBlob   =   "Import_Hide_Unhide.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Import_Hide_Unhide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Callback_Proc As String

'--------------------------------
Private Sub Abbort_Button_Click()
'--------------------------------
   Me.Hide
   If Callback_Proc <> "" Then Run Callback_Proc, False, Import_FromAllSheets_CheckBox
End Sub

'----------------------------
Private Sub OK_Button_Click()
'----------------------------
   Me.Hide
   If Callback_Proc <> "" Then Run Callback_Proc, True, Import_FromAllSheets_CheckBox
End Sub

'----------------------------------------------------------------------------
Public Sub Start(Callback As String, Optional Import_FromAll As Integer = -1)
'----------------------------------------------------------------------------
  Callback_Proc = Callback
  If Import_FromAll = -2 Then
      Import_FromAllSheets_CheckBox = False
      Import_FromAllSheets_CheckBox.Enabled = False
      MultiSelectSheets_Label.Visible = False
  Else
      Import_FromAllSheets_CheckBox.Enabled = True
      MultiSelectSheets_Label.Visible = True
      If Import_FromAll > 0 Then Import_FromAllSheets_CheckBox = True
      If Import_FromAll = 0 Then Import_FromAllSheets_CheckBox = False
  End If
  Me.Show
End Sub



'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
' Center the dialog if it's called the first time.
' On a second call without me.close this function is not called
' => The last position ist used
  Change_Language_in_Dialog Me
  Center_Form Me
End Sub


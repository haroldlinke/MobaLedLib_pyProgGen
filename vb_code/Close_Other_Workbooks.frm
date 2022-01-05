VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Close_Other_Workbooks 
   Caption         =   "Schlieﬂen aller Excel Fenster"
   ClientHeight    =   2790
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5985
   OleObjectBlob   =   "Close_Other_Workbooks.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Close_Other_Workbooks"
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
End Sub

'----------------------------
Private Sub OK_Button_Click()
'----------------------------
   Me.Hide
   If Callback_Proc <> "" Then Run Callback_Proc
End Sub

'-----------------------------------
Public Sub Start(Callback As String)
'-----------------------------------
  Callback_Proc = Callback
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


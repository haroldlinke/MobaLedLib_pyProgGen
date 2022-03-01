VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_DialogGuide1 
   Caption         =   "Einführung und Auswahl der Zielzeile"
   ClientHeight    =   6912
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8880
   OleObjectBlob   =   "UserForm_DialogGuide1.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "UserForm_DialogGuide1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  DialogGuideRes = vbOK
  Me.Hide
End Sub

'------------------------------
Private Sub UserForm_Activate()
'------------------------------
  DialogGuideRes = -1
End Sub

'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  DialogGuideRes = vbAbort
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
End Sub

'------------------------------------
Public Function IsActive() As Boolean
'------------------------------------
  IsActive = Me.Visible
End Function


'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & Me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Center_Form Me
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Connector 
   Caption         =   "Verteiler und Stecker Nummer"
   ClientHeight    =   6072
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   OleObjectBlob   =   "UserForm_Connector.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserForm_Connector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private OK_Pressed As Boolean

'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
End Sub


'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  OK_Pressed = True
  Me.Hide
End Sub

'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Center_Form Me
End Sub



'-----------------------------------------------------------------------------------
Public Function Start(Dist_Nr_R As Excel.Range, Conn_Nr_R As Excel.Range) As Boolean
'-----------------------------------------------------------------------------------
  Dist_Nr = Dist_Nr_R
  Conn_Nr = Conn_Nr_R
  OK_Pressed = False
  
  Dist_Nr.setFocus
  
  Me.Show
  
  If OK_Pressed Then
    Dist_Nr_R = Dist_Nr
    Conn_Nr_R = Conn_Nr
  End If
  Start = OK_Pressed
End Function


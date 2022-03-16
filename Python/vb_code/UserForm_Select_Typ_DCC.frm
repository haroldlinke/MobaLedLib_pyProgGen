VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Select_Typ_DCC 
   Caption         =   "Auswahl des Eingabe Typs"
   ClientHeight    =   5664
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5910
   OleObjectBlob   =   "UserForm_Select_Typ_DCC.frx":0000
End
Attribute VB_Name = "UserForm_Select_Typ_DCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  Userform_Res = ""
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
End Sub


'----------------------------
Private Sub OK_Button_Click()
'----------------------------
  Set_Tast_Txt_Var  ' Set the global variables Red_T, Green_T, ...          06.03.20:
  
  If Button_OnOff Then
                       Userform_Res = OnOff_T
  ElseIf Button_Red Then
                       Userform_Res = Red_T
  Else:                Userform_Res = Green_T
  End If
  Me.Hide
End Sub


'-----------------------------------------
Public Sub setFocus(Target As Excel.Range)
'-----------------------------------------
  Set_Tast_Txt_Var  ' Set the global variables Red_T, Green_T, ...          06.03.20:^
  
  Select Case left(Target, 1)
    Case UCase(left(OnOff_T, 1)): Button_OnOff = True ' "AnAus"
    Case UCase(left(Green_T, 1)): Button_Green = True ' "Grün"
    Case UCase(left(Red_T, 1)):   Button_Red = True   ' "Rot"
    ' In any other cases the last state is used
  End Select
  
  If Button_OnOff Then
                       Button_OnOff.setFocus
  ElseIf Button_Red Then
                       Button_Red.setFocus
  Else:                Button_Green.setFocus
  End If
End Sub


'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & me.Name & ": UserForm_Initialize"
  
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Center_Form Me
End Sub


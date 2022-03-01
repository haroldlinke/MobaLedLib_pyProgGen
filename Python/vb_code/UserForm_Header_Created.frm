VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Header_Created 
   Caption         =   "Header Datei erzeugt"
   ClientHeight    =   4170
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7605
   OleObjectBlob   =   "UserForm_Header_Created.frx":0000
End
Attribute VB_Name = "UserForm_Header_Created"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
End Sub

'---------------------------
Private Sub EditFile_Click()
'---------------------------
  Me.Hide
  Call ShellExecute(0, "open", "Notepad", """" & FileName & """", "", SW_NORMAL)
End Sub

'-------------------------
Private Sub Image1_Click()
'-------------------------
  OK_Click
End Sub

'---------------------
Private Sub OK_Click()
'---------------------
  Me.Hide
  Compile_and_Upload_LED_Prog_to_Arduino
End Sub

'---------------------------
Private Sub OpenPath_Click()
'---------------------------
  Me.Hide
  'Call ShellExecute(0, "open", "Explorer", "/e,""" & ThisWorkbook.Path & """", "", SW_NORMAL)
  
  Dim Name As String
  Name = ThisWorkbook.Path & "\" & Ino_Dir_LED & Include_FileName
  If Dir(Name) <> FileNameExt(Name) Then
        Shell "Explorer /root,""" & FilePath(Name) & """", vbNormalFocus
  Else: Shell "Explorer /Select,""" & Name & """", vbNormalFocus
  End If
End Sub

'---------------------------------------
Private Sub Right_Arduino_Button_Click()
'---------------------------------------
  Ask_to_Upload_and_Compile_and_Upload_Prog_to_Right_Arduino
End Sub

'-----------------------------
Private Sub USB_Button_Click()
'-----------------------------
  USB_Port_Dialog COMPort_COL
End Sub

'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & Me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me                                              ' 20.02.20:
  Center_Form Me
End Sub


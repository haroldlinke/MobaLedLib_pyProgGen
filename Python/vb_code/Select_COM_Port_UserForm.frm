VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Select_COM_Port_UserForm 
   ClientHeight    =   6615
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   10125
   OleObjectBlob   =   "Select_COM_Port_UserForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Select_COM_Port_UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private LocalComPorts() As Byte
Private OldL_ComPorts() As Byte
Private PortNames() As String
Private OldSpinButton As Long
Private Pressed_Button As Long
Private LocalPrintDebug As Boolean
Private LocalShow_ComPort As Boolean

'-------------------------------
Private Sub Check_Button_Click()
'-------------------------------
' Left Button
  Pressed_Button = 1
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
  CheckCOMPort = 0 ' Stop Blink_Arduino_LED()
End Sub

'-------------------------------
Private Sub Abort_Button_Click()
'-------------------------------
' Middle Button
  Pressed_Button = 2
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
  CheckCOMPort = 0 ' Stop Blink_Arduino_LED()
End Sub

'---------------------------------
Private Sub Default_Button_Click()
'---------------------------------
' Right Button
  Pressed_Button = 3
  Me.Hide ' no "Unload Me" to keep the entered data and dialog position
  CheckCOMPort = 0 ' Stop Blink_Arduino_LED()
End Sub

'------------------------------
Private Sub SpinButton_Change()
'------------------------------
  'Debug.Print "Update_SpinButton"
  Update_SpinButton 0
End Sub

'------------------------------------------------------
Public Sub Update_SpinButton(ByVal DefaultPort As Long)
'------------------------------------------------------
' Is also called by the OnTime proc which checks the available ports
  LocalComPorts = EnumComPorts(Show_Unknown_CheckBox, PortNames, PrintDebug:=LocalPrintDebug) ' Read the available COM ports where an Arduino is connected
  If isInitialised(LocalComPorts) Then
        SpinButton.Max = UBound(LocalComPorts)
        If DefaultPort > 0 Then ' DefaultPort is > 0 when it's called the first time
            Dim i As Long
            For i = 0 To UBound(LocalComPorts)
               If DefaultPort = LocalComPorts(i) Then
                  SpinButton = i
               End If
            Next
        Else ' Check if a new com port was detected and select it
            If isInitialised(OldL_ComPorts) Then
               If UBound(LocalComPorts) > UBound(OldL_ComPorts) Then
                  Dim ix As Long
                  For ix = 0 To UBound(OldL_ComPorts)
                      If LocalComPorts(ix) <> OldL_ComPorts(ix) Then
                         SpinButton = ix
                         Exit For
                      End If
                  Next ix
                  If ix > UBound(OldL_ComPorts) Then SpinButton = ix
               End If
            End If
        End If
        If SpinButton > SpinButton.Max Then SpinButton = SpinButton.Max
        CheckCOMPort_Txt = PortNames(SpinButton)
        CheckCOMPort = LocalComPorts(SpinButton)
        COM_Port_Label = " COM" & CheckCOMPort
        If SpinButton <> OldSpinButton Then
           Show_Status False, Get_Language_Str("Aktualisiere Status ...")
           OldSpinButton = SpinButton
        End If
        Dim Port As Variant, PortsStr As String
        For Each Port In LocalComPorts
            PortsStr = PortsStr & Port & " "
        Next
        Available_Ports_Label = DelLast(PortsStr)
        OldL_ComPorts = LocalComPorts
  Else: CheckCOMPort = 999
        Available_Ports_Label = ""
        COM_Port_Label = " -"
  End If
End Sub

'-------------------------------------------------------
Public Sub Show_Status(ErrBox As Boolean, Msg As String)
'-------------------------------------------------------
  If ErrBox Then
       If Error_Label <> Msg Then Error_Label = Msg     ' "If" is used to prevent flickering
  Else
       If Status_Label <> Msg Then Status_Label = Msg
  End If
  If Error_Label.Visible <> ErrBox Then Error_Label.Visible = ErrBox
  If Status_Label.Visible = ErrBox Then Status_Label.Visible = Not ErrBox
End Sub

'----------------------------------------------------------------------------------------------
Public Function ShowDialog(Caption As String, Title As String, Text As String, _
                           Picture As String, _
                           Buttons As String, _
                           FocusButton As String, _
                           Show_ComPort As Boolean, _
                           Red_Hint As String, _
                           ByRef ComPort_IO As Long, _
                           Optional PrintDebug As Boolean = False) As Long
'----------------------------------------------------------------------------------------------
' Variables:
'  Caption     Dialog Caption
'  Title       Dialog Title
'  Text        Message in the text box on the top left side
'  Picture     Name of the picture to be shown. Available pictures: "LED_Image", "CAN_Image", "Tiny_Image", "DCC_Image"
'  Buttons     List of 3 buttons with Accelerator. Example "H Hallo; A Abort; O Ok"  Two Buttons: " ; A Abort; "O Ok"
'  ComPort_IO  is used as input and output
' Return:
'  1: If the left   Button is pressed  (Install, ...)
'  2: If the middle Button is pressed  (Abort)
'  3: If the right  Button is pressed  (OK)
  Me.Caption = Caption
  Title_Label = Title
  Text_Label = Text
  Error_Label = ""
  Status_Label = ""
  
  Dim ButtonArr() As String, BNr
  ButtonArr = Split(Buttons, ";")
  If UBound(ButtonArr) <> 2 Then
     MsgBox "Internal Error in Select_COM_Port_UserForm: 'Buttons' must be a string with 3 buttons separated by ';'" & vbCr & _
            "Wrong: '" & Buttons & "'", vbCritical, "Internal Error (Wrong translation?)"
     EndProg
  End If
  
  Button_Setup Check_Button, ButtonArr(0)
  Button_Setup Abort_Button, ButtonArr(1)
  Button_Setup Default_Button, ButtonArr(2)
  If FocusButton <> "" Then Controls(FocusButton).setFocus
  
  LocalPrintDebug = PrintDebug
  OldSpinButton = -1
  Pressed_Button = 0
  Update_SpinButton ComPort_IO
  SpinButton.Visible = Show_ComPort
  If Show_ComPort Then SpinButton.setFocus
  LocalShow_ComPort = Show_ComPort
  
  ' Show / Hide the COM Port
  COM_Port_Label.Visible = Show_ComPort
  Error_Label.Visible = Show_ComPort
  Status_Label.Visible = Show_ComPort
  AvailPortsTxt_Label.Visible = Show_ComPort
  Available_Ports_Label.Visible = Show_ComPort
  Show_Unknown_CheckBox.Visible = Show_ComPort
  Hint_Label.Visible = Show_ComPort
  If Show_ComPort Then ' Set the height of the main text box
        Text_Label.Height = Error_Label.top - Text_Label.top   ' 78
  Else: Text_Label.Height = Hint_Label.top + Hint_Label.Height - Text_Label.top
  End If
  
  Dim c As Variant, Found As Boolean
  For Each c In Me.Controls
     If right(c.Name, Len("Image")) = "Image" Then
        If Picture = c.Name Then
              c.Visible = True: Found = True
        Else: c.Visible = False
        End If
     End If
  Next
  If Not Found Then MsgBox "Internal Error: Unknown picture: '" & Picture & "'", vbCritical, "Internal Error"
  
  Red_Hint_Label = Red_Hint
  Me.Show
  
  ' Store the results
  If Show_ComPort Then
    If isInitialised(LocalComPorts) Then
        ComPort_IO = LocalComPorts(SpinButton)
    End If
  End If
  ShowDialog = Pressed_Button
End Function

'--------------------------------
Private Sub UserForm_Initialize()
'--------------------------------
  'Debug.Print vbCr & me.Name & ": UserForm_Initialize"
  Change_Language_in_Dialog Me
  Center_Form Me
End Sub

Private Sub UserForm_QueryClose(CloseMode As Integer, Cancel As Integer)
    Abort_Button_Click
End Sub


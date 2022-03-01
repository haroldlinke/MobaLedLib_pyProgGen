VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'---------------------------------
Private Sub Arduino_Button_Click()
'---------------------------------
  Button_Pressed_Proc
  Create_HeaderFile
End Sub

'------------------------------------
Private Sub ClearSheet_Button_Click()
'------------------------------------
  Button_Pressed_Proc
  ClearSheet
End Sub

'-------------------------------
Private Sub Dialog_Button_Click()
'-------------------------------
  Button_Pressed_Proc
  Dialog_Guided_Input
End Sub

'------------------------------
Private Sub Help_Button_Click()
'------------------------------
  Button_Pressed_Proc
  Show_Help
End Sub


'--------------------------------
Private Sub Button_Pressed_Proc()
'--------------------------------
  Selection.Select ' Remove the focus from the button
  Application.EnableEvents = True ' In case the program crashed before
  Correct_Buttonsizes
End Sub


Private Sub Name_with_LF()
  Hide_Button.Caption = "Aus- oder" & vbLf & "Einblenden"
End Sub

'-----------------------------------------
Private Sub Correct_Create_Buttonsize(obj)
'-----------------------------------------
  obj.Height = 160
  obj.Width = 100
  obj.Height = 76
  obj.Width = 60
End Sub

'------------------------
Sub Correct_Buttonsizes()
'------------------------
' There is a bug in excel which changes the size of the buttons
' if the resolution of the display is changed. This happens
' fore instance if the computer is connected to a beamer.
' To prevent this the buttons are resized with this function.
  Dim OldScreenupdating As Boolean
  OldScreenupdating = Application.ScreenUpdating
  Application.ScreenUpdating = False

  Correct_Create_Buttonsize Arduino_Button
  Correct_Create_Buttonsize Dialog_Button
  Correct_Create_Buttonsize Insert_Button
  Correct_Create_Buttonsize Del_Button
  Correct_Create_Buttonsize Move_Button
  Correct_Create_Buttonsize Copy_Button
  Correct_Create_Buttonsize Hide_Button
  Correct_Create_Buttonsize UnHideAll_Button
  
  Correct_Create_Buttonsize ClearSheet_Button
  Correct_Create_Buttonsize Options_Button
  Correct_Create_Buttonsize Help_Button
  
  Application.ScreenUpdating = OldScreenupdating
End Sub

'------------------------------------------------------
Public Sub EnableDisableAllButtons(Enab As Boolean)
'------------------------------------------------------
  Arduino_Button.Enabled = Enab
  Dialog_Button.Enabled = Enab
  Insert_Button.Enabled = Enab
  Del_Button.Enabled = Enab
  Move_Button.Enabled = Enab
  Copy_Button.Enabled = Enab
  Hide_Button.Enabled = Enab
  UnHideAll_Button.Enabled = Enab
  
  ClearSheet_Button.Enabled = Enab
  Options_Button.Enabled = Enab
End Sub

'-----------------------------
Private Sub EnableAllButtons()
'-----------------------------
' Could be called manually after a crash
  EnableDisableAllButtons True
End Sub


'------------------------------
Private Sub Hide_Button_Click()
'------------------------------
  Button_Pressed_Proc
  Proc_Hide_Unhide
End Sub

Private Sub Insert_Button_Click()
  Button_Pressed_Proc
  Proc_Insert_Line
End Sub

'-----------------------------
Private Sub Del_Button_Click()
'-----------------------------
  Button_Pressed_Proc
  Proc_Del_Row
End Sub

'------------------------------
Private Sub Move_Button_Click()
'------------------------------
  Button_Pressed_Proc
  Proc_Move_Row
End Sub

'------------------------------
Private Sub Copy_Button_Click()
'------------------------------
  Button_Pressed_Proc
  Proc_Copy_Row
End Sub

'----------------------------------
Private Sub Options_Button_Click()
'----------------------------------
  Button_Pressed_Proc
  Option_Dialog
End Sub

'-----------------------------------
Private Sub UnHideAll_Button_Click()
'-----------------------------------
  Button_Pressed_Proc
  Proc_UnHide_All
End Sub

'--------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Excel.Range)
'--------------------------------------------------------
' This function is called if the worksheet is changed.
' It performs several checks after a user input depending form the column of the changed cell:
  Global_Worksheet_Change Target
End Sub

'-----------------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
'-----------------------------------------------------------------
' Is called by event if the worksheet selection has changed
  Global_Worksheet_SelectionChange Target
End Sub

'-------------------------------
Private Sub Worksheet_Activate()
'-------------------------------
  Global_Worksheet_Activate
End Sub

'--------------------------------
Private Sub Worksheet_Calculate()
'--------------------------------
' This proc is also called if an other workbook (from a mail/internet) is opened    24.02.20:
  If DEBUG_CHANGEEVENT Then Debug.Print "Worksheet_Calculate() in sheet 'DCC' called"
  If ActiveSheet Is Nothing Then Exit Sub
  
  If Cells.Parent.Name = ActiveSheet.Name Then ' At program start the Worksheet_Calculate proc is called without activating the sheet.
     Global_Worksheet_Calculate                ' This causes a crash in Make_sure_that_Col_Variables_match
  End If
End Sub


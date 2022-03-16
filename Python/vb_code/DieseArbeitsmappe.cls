VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Const ENABLE_CRITICAL_EVENTS_WB = False ' Deactivate this for tests. Call "Remove_Special_Events" below if disabled
' If this is disabled the LED numbers are not ubdated if lines are hidden or unhidden
' But the numbers are corrected if oter things are changed (Enable hook changed, macros added, ...
' => I have disabled it on 12.10.20.

'----------------------------------
Private Sub Remove_Special_Events()                                         ' 11.10.20:
'----------------------------------
' Events which generate a link to Prog_Generator_MobaLedLib.xlsm
' They get active if hidden lines are unhiden
    Application.CommandBars("row").FindControl(Id:=883).OnAction = ""  ' "myHideRows_Event"
    Application.CommandBars("row").FindControl(Id:=884).OnAction = ""  ' "myUnhideRows_Event"
End Sub



'--------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
'--------------------------------------------------
  If ENABLE_CRITICAL_EVENTS_WB Then
    Application.CommandBars("row").FindControl(Id:=883).OnAction = ""
    Application.CommandBars("row").FindControl(Id:=884).OnAction = ""
  End If
End Sub

'-------------------------
Public Sub Workbook_Open()
'-------------------------
  Application.ScreenUpdating = False
  EnableAllButtons
  Application.EnableEvents = False ' 28.10.19.
  
  Cleare_Mouse_Hook                                                         ' 28.04.20:
  
  Debug.Print "Workbook_Open() called"                                      ' 23.02.20:
  ThisWorkbook.Sheets(LANGUAGES_SH).Visible = False                         ' 29.04.20: Added: ThisWorkbook to hopefully prevent problems at startup
  ThisWorkbook.Sheets(LIBMACROS_SH).Visible = False
  ThisWorkbook.Sheets(PAR_DESCR_SH).Visible = False
  ThisWorkbook.Sheets(LIBRARYS__SH).Visible = False                         ' 28.05.20:
  ThisWorkbook.Sheets(PLATFORMS_SH).Visible = False                         ' 14.10.21: Juergen
  
  Check_Version                                                             ' 21.11.21: Juergen
  
  If ActiveSheet.Name = ConfigSheet Then Sheets(START_SH).Select            ' 30.05.20: In case one of the hidden sheets was active before
  Init_Libraries_Page
  
  Update_Language_in_All_Sheets
  
  
  Clear_COM_Port_Check_and_Set_Cursor_in_all_Sheets False
  Set_Config_Default_Values_at_Program_Start
  
  Install_Missing_Libraries_and_Board                                       ' 28.05.20:
  
  Load_Extensions                                                           ' 26.01.22 Juergen add Extension feature
  
  Application.ScreenUpdating = True
  
  If ENABLE_CRITICAL_EVENTS_WB Then
    ' Generate events for special actions
    On Error GoTo ErrorDetectEvents
    Application.CommandBars("row").FindControl(Id:=883).OnAction = "myHideRows_Event"
    Application.CommandBars("row").FindControl(Id:=884).OnAction = "myUnhideRows_Event"
    On Error GoTo 0
  Else
    Remove_Special_Events                                                     ' 13.10.20:
  End If

  Dim DidCopy As Boolean                                                      ' 12.11.21: Juergen ImportOld version also in case of Beta Update
  DidCopy = False
  If Copy_Prog_If_in_LibDir_WithResult(DidCopy) Then
    Import_from_Old_Version_If_exists (Not DidCopy)
  End If
  
  If Get_Num_Config_Var_Range("SimAutostart", 0, 1, 0) = 1 Then               ' 04.04.22: Juergen Simulator
    OpenSimulator
  End If
  
  Application.EnableEvents = True ' 28.10.19.
  Exit Sub
  
ErrorDetectEvents:
  MsgBox "Interner Fehler: Die Event Routinen wurden nicht gefunden", vbCritical, "Interner Fehler"
End Sub




'****************************** Mouse Events **************************

#If 0 Then ' Disabled because it's enabled and disabled in M24_Mouse_Inset_Pos

    '************ Left Mouse button ************
    '------------------------------
    Private Sub Workbook_Activate()
    '------------------------------
      'hTimer = SetTimer(0, 0, 50, AddressOf M24_Mouse_Inset_Pos.MausUeberwachen)
    End Sub
    
    '--------------------------------
    Private Sub Workbook_Deactivate()
    '--------------------------------
      'KillTimer 0, hTimer
    End Sub
#End If


' Es gibt zwei Methoden wie man die RECHTE Maustaste abfängt:
' Die eine muss in "Dieser Arbeitsmappe" stehen:

#If 0 Then
    '-------------------------------------------------------------------------------------------------------
    Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    '-------------------------------------------------------------------------------------------------------
      Cancel = True
      Debug.Print "DieserArbeitsmappe: Workbook_SheetBeforeRightClick" & Target.Row
    End Sub
#End If


' Die andere muss in der Code Seite es entsprechenden Sheets stehen:

'  Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
'  ' https://docs.microsoft.com/de-de/office/vba/api/excel.worksheet.beforerightclick
'  ' !! RECHTE Taste nicht die linke [X]<- [ ]
'    Debug.Print Activesheet.Name ":Worksheet_BeforeRightClick " & Target.Row
'  End Sub


'--------------------------------------------------------------------------------------------------------------
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, ByRef Cancel As Boolean)
'--------------------------------------------------------------------------------------------------------------
  If DEBUG_CHANGEEVENT Then Debug.Print "DieserArbeitsmappe: Workbook_SheetBeforeDoubleClick:" & Target.Row & " Cancel=" & Cancel
  Proc_DoubleCkick Sh, Target, Cancel
  'Cancel = True
End Sub

#If 0 Then ' Not used at the moment
'------------------------------------------------------------
Private Sub Workbook_WindowActivate(ByVal Wn As Excel.Window)
'------------------------------------------------------------
' Is called if switched from an other excel Workbook to this workbook
  If DEBUG_CHANGEEVENT Then Debug.Print ThisWorkbook.Name & " Workbook_WindowActivate event"
End Sub
#End If



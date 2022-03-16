Attribute VB_Name = "M35_Mouse_Scroll"
Option Explicit
' Mouse wheel support for list boxes
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
'
' Program combined from
'  https://stackoverflow.com/questions/34911413/mouse-scroll-on-a-listbox
'  https://www.ozgrid.com/forum/forum/help-forums/excel-general/128354-mouse-wheel-scroll-userform
'
' Hier eine weitere Lösung:
'  http://www.office-loesung.de/ftopic174250_0_0_asc.php

' Hier wird von Problemen mit 32/64 Bit berichtet
' https://stackoverflow.com/questions/36621795/use-mouse-wheel-in-excel-dynamic-combobox-not-working-on-excel-2010

' Revision History:
' ~~~~~~~~~~~~~~~~~
' 28.04.20: - Hopefully saved the MouseHook problem by storing the MouseHook to the worksheet
'


' Usage:
' ~~~~~~
' - Define the page where the MouseHook is stored in a public modul:
'    Public Const MouseHook_Store_Page = "Start"  ' This page is used to store the MouseHook. There must be a named range called "MouseHook".
'
' Add the following calls to the UserForm code:
' - Call Cleare_Mouse_Hook() when the program is started !!!
' - To the initialisation of the user form:
'     HookFormScroll Me, "ListBox"   ' Initialize the mouse wheel scroll function
'
' - When the form is closed:
'     UnhookFormScroll ' Deactivate the mouse weel scrol function
'
' - The following sub is called if the mouse wheel is changed
'   It must be copied to the user form and adapted (Remove the #if ...)

#Const SPEEDUP_MOUSE_SCROLL = False                                         ' 09.11.20: Added but it dosn't speed up the scrolling ;-(

#If SPEEDUP_MOUSE_SCROLL Then
  Private Is_Mouse_Hook_Stored_in_RAM As Boolean
  Private Is_Mouse_Hook_Stored_RAM As Boolean
#End If



#If False Then
'-----------------------------------------------
Public Sub MouseWheel(ByVal lngRotation As Long)
'-----------------------------------------------
' Process the mouse wheel changes
  With ListBox  ' Adapt to the listbox which should be controlled
    If lngRotation > 0 Then
        If .TopIndex > 0 Then
            If .TopIndex > 3 Then
                .TopIndex = .TopIndex - 3
            Else
                .TopIndex = 0
            End If
        End If
    Else
        .TopIndex = .TopIndex + 3
    End If
  End With
End Sub
#End If

' 02.10.20: The mouse scroll is always disabled for Office <= 2007 because here excel generates a crash
#Const ENABLE_CRITICAL_EVENTS_MOUSE = True  ' Disable this for debugging    29.04.20:


Private Const PRINT_MOUSE_DEBUG = False   ' Enable the Debug messages. The Error messages printed to the Debug window are always shown


' Overview 32 / 64 Bit functions: https://jkp-ads.com/Articles/apideclarations.asp
' https://stackoverflow.com/questions/45324586/using-setwindowshookex-in-excel-2010
#If VBA7 Then 'For 64 Bit Systems
   Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
   Private Declare PtrSafe Function GetForegroundWindow Lib "user32.dll" () As LongPtr
   Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
   Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
   Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hHook As LongPtr, ByVal nCode As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
   Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
#Else 'This will compile in 32 bit Excel only
   Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
   Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
   Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
   Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
   Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
   Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If



#If False Then ' Das braucht man nicht
    ' This is one of the few API functions that requires the Win64 compile constant:
    #If VBA7 Then
        #If Win64 Then
            Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        #Else
            Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        #End If
    #Else
        Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    #End If
#End If

   
Private Const HC_ACTION = 0
Private Const WH_MOUSE_LL = 14
Private Const WM_MOUSEWHEEL = &H20A
Private Const GWL_HINSTANCE = (-6)

Public Const nMyControlTypeNONE = 0
Public Const nMyControlTypeUSERFORM = 1
Public Const nMyControlTypeFRAME = 2
Public Const nMyControlTypeCOMBOBOX = 3
Public Const nMyControlTypeLISTBOX = 4


#If VBA7 Then
  Private mLngMouseHook              As LongPtr
  Private mFormHwnd                  As LongPtr
  Type MouseHook_T
      v As LongPtr
  End Type
#Else
  Private mLngMouseHook              As Long
  Private mFormHwnd                  As Long
  Type MouseHook_T
      v As Long
  End Type
#End If

Private mbHook                     As Boolean
Dim mForm                          As Object

Private mControllName              As String


'   MSLLHOOKSTRUCT
'   0    pt.X Long
'   4    pt.Y Long
'   8    mouseData Long  Holds Forward\Backward flag
'   12   flags Long
'   16   time Long
'   20   dwExtraInfo Long

#If VBA7 Then
    Function GetMouseData(ByVal lParam As LongPtr) As Long
      Dim Value As Long
      ' offset of MouseData in MSLLHOOKSTRUCT is 8
      CopyMemory Value, ByVal lParam + 8, 4
      GetMouseData = Value
    End Function
#Else
    Function GetMouseData(ByVal lParam As Long) As Long
      Dim Value As Long
      ' offset of MouseData in MSLLHOOKSTRUCT is 8
      CopyMemory Value, ByVal lParam + 8, 4
      GetMouseData = Value
    End Function
#End If

#If VBA7 Then
  '-----------------------------------------------------------------------------------------------------------
  Function LowLevelMouseProc(ByVal nCode As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
  '-----------------------------------------------------------------------------------------------------------
#Else
  Function LowLevelMouseProc(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If
  'Avoid XL crashing if RunTime error occurs due to Mouse fast movement

  Dim iDirection As Long

  On Error Resume Next
  
  ' Unhook & get out in case the application is deactivated
  If mbHook = False And Is_Mouse_Hook_Stored() Then                         ' 28.04.20:
     Debug.Print "!!! Disabling MouseHook after debug break !!!"
     mbHook = True
     mLngMouseHook = Read_Mouse_Hook.v
     UnhookFormScroll
     Exit Function
  End If
  
  If GetForegroundWindow = mFormHwnd Then ' Hardi
      If (nCode = HC_ACTION) Then
        If wParam = WM_MOUSEWHEEL Then
          iDirection = GetMouseData(lParam)
          mForm.MouseWheel iDirection
    
          ' Don't process Default WM_MOUSEWHEEL Window message
          LowLevelMouseProc = True
        End If
    
        Exit Function
      End If
  End If
  LowLevelMouseProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)
End Function


'----------------------------
Private Sub Save_Mouse_Hook()                                               ' 28.04.20:
'----------------------------
  ThisWorkbook.Sheets(MouseHook_Store_Page).Range("MouseHook") = mLngMouseHook
  #If SPEEDUP_MOUSE_SCROLL Then                                             ' 09.11.20:
    Is_Mouse_Hook_Stored_in_RAM = True
    Is_Mouse_Hook_Stored_RAM = True
  #End If
End Sub

'-----------------------------
Public Sub Cleare_Mouse_Hook()                                              ' 28.04.20:
'-----------------------------
  ThisWorkbook.Sheets(MouseHook_Store_Page).Range("MouseHook") = ""
  #If SPEEDUP_MOUSE_SCROLL Then                                             ' 09.11.20:
    Is_Mouse_Hook_Stored_in_RAM = True
    Is_Mouse_Hook_Stored_RAM = False
  #End If
End Sub

'------------------------------------------------
Private Function Read_Mouse_Hook() As MouseHook_T                            ' 28.04.20:
'------------------------------------------------
  Read_Mouse_Hook.v = ThisWorkbook.Sheets(MouseHook_Store_Page).Range("MouseHook")
End Function

'-------------------------------------------------
Private Function Is_Mouse_Hook_Stored() As Boolean                          ' 28.04.20:
'-------------------------------------------------
  #If SPEEDUP_MOUSE_SCROLL Then                                             ' 09.11.20:
        If Not Is_Mouse_Hook_Stored_in_RAM Then
           Is_Mouse_Hook_Stored_RAM = ThisWorkbook.Sheets(MouseHook_Store_Page).Range("MouseHook") <> "" And IsNumeric(ThisWorkbook.Sheets(MouseHook_Store_Page).Range("MouseHook"))
           Is_Mouse_Hook_Stored_in_RAM = True
        End If
        Is_Mouse_Hook_Stored = Is_Mouse_Hook_Stored_RAM
  #Else
        Is_Mouse_Hook_Stored = ThisWorkbook.Sheets(MouseHook_Store_Page).Range("MouseHook") <> "" And IsNumeric(ThisWorkbook.Sheets(MouseHook_Store_Page).Range("MouseHook"))
  #End If
End Function

'UT------------------
Private Sub Test_MH()
'UT------------------
  Debug.Print Is_Mouse_Hook_Stored()
  Debug.Print CLng(Read_Mouse_Hook().v)
End Sub



'---------------------
Sub UnhookFormScroll()
'---------------------
  If PRINT_MOUSE_DEBUG Then Debug.Print "Mouse: UnhookFormScroll() called mbHook=" & mbHook
  If mbHook Then
      UnhookWindowsHookEx mLngMouseHook
      mLngMouseHook = 0
      mFormHwnd = 0
      mbHook = False
      Cleare_Mouse_Hook                                                     ' 28.04.20:
   End If
End Sub


'--------------------------------------------------
Private Function Newer_than_Office2007() As Boolean                         ' 05.11.20:
'--------------------------------------------------
  Static Checked As Boolean
  Static Res As Boolean
  If Checked Then
     Newer_than_Office2007 = Res
     Exit Function
  End If
  Checked = True
  Res = (val(Application.Version) > 12)  ' > Office 2007
  Newer_than_Office2007 = Res
End Function


'----------------------------------------------------------
Sub HookFormScroll(oForm As Object, ControllName As String)
'----------------------------------------------------------
#If ENABLE_CRITICAL_EVENTS_MOUSE Then
  If Not Newer_than_Office2007() Then Exit Sub    ' <= Office 2007

#If VBA7 Then
   Dim lngAppInst       As LongPtr
   Dim hwndUnderCursor  As LongPtr
#Else
   Dim lngAppInst       As Long
   Dim hwndUnderCursor  As Long
#End If

   Set mForm = oForm
   hwndUnderCursor = FindWindow("ThunderDFrame", oForm.Caption)
   If PRINT_MOUSE_DEBUG Then Debug.Print "Mouse: HookFormScroll() Form window: " & hwndUnderCursor ' Debug
   If mFormHwnd <> hwndUnderCursor Then
      mControllName = ControllName
      UnhookFormScroll
      mFormHwnd = hwndUnderCursor

      #If 0 Then  ' Das braucht man nicht. Es geht auch wenn lngAppInst = 0 ist
        #If VBA7 Then
           lngAppInst = GetWindowLongPtr(mFormHwnd, GWL_HINSTANCE)  ' Geht das so ???
           'lngAppInst = GetWindowLongPtr(FindWindow("XLMAIN", Application.Caption), GWL_HINSTANCE) ' Geht auch nicht
        #Else
           lngAppInst = GetWindowLong(mFormHwnd, GWL_HINSTANCE)
        #End If
      #End If

      If mbHook Then
         Debug.Print "!!! Mouse: Mouse was already hooked !!! => Don't hook it again "
      Else
         mLngMouseHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf LowLevelMouseProc, lngAppInst, 0)
         Save_Mouse_Hook ' Save the MouseHook to the worksheet to be able to restore it in case the program was aborted  ' 28.04.20:
         mbHook = mLngMouseHook <> 0
         If mbHook Then
              If PRINT_MOUSE_DEBUG Then Debug.Print "Mouse: Form hooked"
         Else: Debug.Print "!!! Mouse: Error hook !!!"
         End If
      End If
   Else: Debug.Print "!!! Mouse: Error mFormHwnd = hwndUnderCursor !!!"
   End If
#End If ' ENABLE_CRITICAL_EVENTS_MOUSE
End Sub



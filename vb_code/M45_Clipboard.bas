Attribute VB_Name = "M45_Clipboard"
Option Explicit
#If False Then  ' 29.04.20: Not used any more

' https://www.spreadsheet1.com/how-to-copy-strings-to-clipboard-using-excel-vba.html

#If VBA7 Then
  #If Win64 Then                                                            ' 18.12.19:
    Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As LongPtr
    Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As Long
  #Else
    Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As LongPtr, ByVal dwBytes As LongPtr) As Long
    Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As LongPtr, ByVal hMem As LongPtr) As Long
  #End If
#Else
  Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
  Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
  Declare Function CloseClipboard Lib "user32" () As Long
  Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
  Declare Function EmptyClipboard Lib "user32" () As Long
  Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
  Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
#End If

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MaxSize = 4096


'-----------------------------------------------------------------
Public Function ClipBoard_SetData(sPutToClip As String) As Boolean
'-----------------------------------------------------------------

    ' www.msdn.microsoft.com/en-us/library/office/ff192913.aspx
#If Win64 Then                                                              ' 18.12.19:
    Dim hGlobalMemory As LongPtr
    Dim lpGlobalMemory As LongPtr
    Dim hClipMemory As LongLong
#Else
    Dim hGlobalMemory As Long
    Dim lpGlobalMemory As Long
    Dim hClipMemory As Long
#End If
    Dim X As Long
    
    On Error GoTo ExitWithError_

    ' Allocate moveable global memory
    hGlobalMemory = GlobalAlloc(GHND, Len(sPutToClip) + 1)

    ' Lock the block to get a far pointer to this memory
    lpGlobalMemory = GlobalLock(hGlobalMemory)

    ' Copy the string to this global memory
    lpGlobalMemory = lstrcpy(lpGlobalMemory, sPutToClip)

    ' Unlock the memory
    If GlobalUnlock(hGlobalMemory) <> 0 Then
        MsgBox "Memory location could not be unlocked. Clipboard copy aborted", vbCritical, "API Clipboard Copy"
        GoTo ExitWithError_
    End If

    ' Open the Clipboard to copy data to
    If OpenClipboard(0&) = 0 Then
        MsgBox "Clipboard could not be opened. Copy aborted!", vbCritical, "API Clipboard Copy"
        GoTo ExitWithError_
    End If

    ' Clear the Clipboard
    X = EmptyClipboard()

    ' Copy the data to the Clipboard
    hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)
    ClipBoard_SetData = True
    
    If CloseClipboard() = 0 Then
        MsgBox "Clipboard could not be closed!", vbCritical, "API Clipboard Copy"
    End If
    Exit Function
ExitWithError_:
    On Error Resume Next
    If Err.Number > 0 Then MsgBox "Clipboard error: " & Err.Description, vbCritical, "API Clipboard Copy"
    ClipBoard_SetData = False

End Function

'UT---------------------------------
Private Sub Test_ClipBoard_SetData()
'UT---------------------------------
  ClipBoard_SetData "Hallo Armin"
End Sub
#End If

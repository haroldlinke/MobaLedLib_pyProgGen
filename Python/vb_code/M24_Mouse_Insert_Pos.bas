Attribute VB_Name = "M24_Mouse_Insert_Pos"
Option Explicit

' Example from
'  - http://www.herber.de/forum/archiv/1124to1128/1126102_GetKeyState_Taste_abfragen.html
'    http://www.herber.de/bbs/user/66853.xls
'  - https://stackoverflow.com/questions/47271141/vba-get-cursor-position-as-cell-address

' https://stackoverflow.com/questions/20269844/api-timers-in-vba-how-to-make-safe
#If VBA7 And Win64 Then    ' 64 bit Excel under 64-bit windows
                           ' Use LongLong and LongPtr
                           
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongLong, ByVal lpTimerFunc As LongPtr) As LongLong
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As LongLong
    
    Private hTimer As LongPtr ' Müsste es hier nicht LongLong heißen ?

#ElseIf VBA7 Then           ' 64 bit Excel in all environments
                            ' Use LongPtr only, LongLong is not available
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr ' 01.12.20: Old: nIDEvent As Long
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long                                                        '    "              "

    Private hTimer As LongPtr
    
#Else    ' 32 bit Excel
    Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

    Private hTimer As Long
#End If


#If VBA7 Then
  Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
  
  Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

  Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
#Else
  Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
  
  Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

  Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
#End If

Public Const VK_SHIFT = &H10   ' 12.03.21 Juergen
Public Const VK_CONTROL = &H11 ' used in GetAsyncKeyState


' Create custom variable that holds two integers
Type POINTAPI
    Xcoord As Long
    Ycoord As Long
End Type

' https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_MBUTTON = &H4

Private Const VK_UP = &H26
Private Const VK_DOWN = &H28
Private Const VK_RETURN = &HD
Private Const VK_ESCAPE = &H1B

Private LeftMousePressed As Boolean
Private ESCButtonPressed As Boolean
Private EnterKey_Pressed As Boolean
Private LastRow As Long
Private Col1 As Long
Private ColN As Long

'--------------------------------
Private Sub MouseCheckTimerProc()
'--------------------------------
Dim Result%
KillTimer 0, hTimer

Result = GetAsyncKeyState(VK_LBUTTON)
If Result <> 0 Then LeftMousePressed = True

hTimer = SetTimer(0, 0, 50, AddressOf MouseCheckTimerProc)
End Sub


'---------------------------------------
Private Sub Show_Insert_Pos(Row As Long)
'---------------------------------------
  If Row > 0 Then
    With Range(Cells(Row, Col1), Cells(Row, ColN)).Borders(xlEdgeTop)
        .ThemeColor = 10
        .TintAndShade = -0.249977111117893
        .Weight = xlThick
    End With
  End If
End Sub

'----------------------------------------------------
Private Sub Normal_Line(Sh As Worksheet, Row As Long)
'----------------------------------------------------
  With Sh
    With .Range(.Cells(Row, Col1), .Cells(Row, ColN)).Borders(xlEdgeTop)
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
  End With
End Sub

'-------------------------------------------------------
Private Function GetRange(X As Long, Y As Long) As Range
'-------------------------------------------------------
  On Error Resume Next
    Set GetRange = ActiveWindow.RangeFromPoint(X, Y)
  On Error GoTo 0
End Function

'--------------------------------------------------------------------------------------------------
Private Function Show_InsertLine_until_Mousepressed(MinRow As Long, SheetName As String) As Boolean
'--------------------------------------------------------------------------------------------------
    Dim llCoord As POINTAPI, rng As Range
    
    GetCursorPos llCoord ' Get the cursor positions
    'Debug.print "X Position: " & llCoord.Xcoord & vbNewLine & "Y Position: " & llCoord.Ycoord ' Display the cursor position coordinates
    
    DoEvents
    
    If GetAsyncKeyState(VK_ESCAPE) <> 0 Then ESCButtonPressed = True
    
    Dim MoveByKey As Boolean
    If GetAsyncKeyState(VK_UP) <> 0 Then MoveByKey = True
    If GetAsyncKeyState(VK_DOWN) <> 0 Then MoveByKey = True
    If GetAsyncKeyState(VK_RETURN) <> 0 Then EnterKey_Pressed = True
    
    If MoveByKey Then
          SetCursorPos ActiveWindow.ActivePane.PointsToScreenPixelsX(ActiveCell.left + (ActiveCell.Width / 2)), _
              ActiveWindow.ActivePane.PointsToScreenPixelsY(ActiveCell.top + (ActiveCell.Height / 2))
          Set rng = ActiveCell
    Else: Set rng = GetRange(llCoord.Xcoord, llCoord.Ycoord)
    End If
    
    Dim Row As Long
    If Not rng Is Nothing Then
       Row = rng.Row
       If Row < MinRow Then Row = MinRow
    End If
    

    If (Row <> 0 And Row <> LastRow) Or LeftMousePressed Or EnterKey_Pressed Or ESCButtonPressed Then
       Dim OldUpdating As Boolean
       OldUpdating = Application.ScreenUpdating
       Application.ScreenUpdating = False
       If LastRow > 0 Then Normal_Line Sheets(SheetName), LastRow
       If Row <> 0 Then LastRow = Row
       
       If LeftMousePressed Or EnterKey_Pressed Or ESCButtonPressed Or SheetName <> ActiveSheet.Name Then
             Show_InsertLine_until_Mousepressed = True ' Abort
       Else: Show_Insert_Pos Row
       End If
       Application.ScreenUpdating = OldUpdating
    End If
End Function

'-----------------------------------------------------------------------------------
Public Function Select_Move_Dest_by_Mouse(FirstCol As Long, LastCol As Long) As Long
'-----------------------------------------------------------------------------------
' End when Left Mouse, Enter or ESC is pressed
' Return the destination Row
' Return 0 if aborted with ESC
'
   Col1 = FirstCol
   ColN = LastCol
   
   hTimer = SetTimer(0, 0, 50, AddressOf MouseCheckTimerProc)
   LeftMousePressed = False
   EnterKey_Pressed = False
   ESCButtonPressed = False
   
   Dim ShName As String
   ShName = ActiveSheet.Name
   
   While Show_InsertLine_until_Mousepressed(FirstDat_Row, ShName) = False
   Wend
   
   KillTimer 0, hTimer
   
   If ActiveSheet.Name = ShName And (LeftMousePressed Or EnterKey_Pressed) Then
      Select_Move_Dest_by_Mouse = LastRow
   End If
End Function

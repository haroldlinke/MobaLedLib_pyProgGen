Attribute VB_Name = "M30_Tools"
Option Explicit

Const PlatformKey_ROW = 3
Const PlatformKey_COL = 3
Const PlatformParName_COL = 1

Private PlatformParams As Scripting.Dictionary

' Module Description:
' ~~~~~~~~~~~~~~~~~~~
' This module contains general tools.

' Overview 32 / 64 Bit functions: https://jkp-ads.com/Articles/apideclarations.asp

#If VBA7 Then 'For 64 Bit Systems
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'    Public Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else 'For 32 Bit Systems
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'    Public Declare Function GetKeyState Lib "user32" (ByVal vKey As Integer) As Integer
#End If

' Test whether you are using the 64-bit version of Office 2010.
#If Win64 Then
   Public Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
   Private StartTime_for_ms_Timer As LongLong
#Else
   #If VBA7 Then
      Public Declare PtrSafe Function getTickCount Lib "kernel32" Alias "GetTickCount" () As Long
   #Else
      Public Declare Function getTickCount Lib "kernel32" Alias "GetTickCount" () As Long
   #End If
   Private StartTime_for_ms_Timer As Long
#End If


' https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getsystemmetrics    ' 23.04.20:
#If VBA7 Then 'For 64 Bit Systems
  Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
  Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If


#If Win64 Then                                                              ' 20.05.20:
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
#Else
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

#If Mac Then                                                                ' 07.10.21:
 ' don't compile the APIs in Mac
#ElseIf VBA7 Then
    Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" ( _
            ByRef lpPoint As POINTAPI) As Long
    Private Declare PtrSafe Function SetCursorPos Lib "user32.dll" ( _
            ByVal X As Long, _
            ByVal Y As Long) As Long
#Else
    Private Declare Function GetCursorPos Lib "user32.dll" ( _
                                          ByRef lpPoint As POINTAPI) As Long
    Private Declare Function SetCursorPos Lib "user32.dll" ( _
                                          ByVal X As Long, _
                                          ByVal Y As Long) As Long
#End If

Private Type POINTAPI
    X As Long
    Y As Long
End Type


#If VBA7 Then
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, _
    ByVal lpUsedDefaultChar As LongPtr) As Long                             ' 07.06.20: Using LongPtr instead of long
 #Else
 Private Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
 #End If
 
''' Maps a character string to a UTF-16 (wide character) string             ' 26.05.20:
#If VBA7 Then
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long _
    ) As Long
#Else
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long _
    ) As Long
#End If

Private Const CP_UTF8 As Long = 65001
    
'Some constants for GetSystemMetrics
Public Const SM_CXSCREEN = 0            ' The width of the screen of the primary display monitor, in pixels.
Public Const SM_CYSCREEN = 1            ' The height of the screen of the primary display monitor, in pixels.
Public Const SM_CMONITORS = 80          ' The number of display monitors on a desktop. (counts only visible display monitors)


Public Const SW_NORMAL = 1

'Public Const KS_SHIFT_KEY = 16
'Public Const KS_CTRL_KEY = 17
'Public Const KS_ALT_KEY = 18


Public Type WinPos_T
  Valid As Boolean
  Left As Double
  Top As Double
End Type


'--------------------------
Public Sub Start_ms_Timer()
'--------------------------
' Timer for debugging
  #If Win64 Then
    StartTime_for_ms_Timer = GetTickCount64()                               ' 29.04.20:
  #Else
    StartTime_for_ms_Timer = getTickCount()
  #End If
End Sub

'----------------------------------------
Public Function Get_ms_Duration() As Long
'----------------------------------------
' Timer for debugging
  Get_ms_Duration = getTickCount() - StartTime_for_ms_Timer
End Function



'----------------------------------------------------------------------------
Public Function AddSpaceToLen(ByVal s As String, MinLength As Long) As String
'----------------------------------------------------------------------------
  While Len(s) < MinLength
     s = s & " "
  Wend
  AddSpaceToLen = s
End Function

'--------------------------------------------------------------------------------
Public Function AddSpaceToLenLeft(ByVal s As String, MinLength As Long) As String
'--------------------------------------------------------------------------------
  While Len(s) < MinLength
     s = " " & s
  Wend
  AddSpaceToLenLeft = s
End Function


'---------------------------------------------------
Function IsArrayEmpty(anArray As Variant) As Boolean
'---------------------------------------------------
On Error GoTo IS_EMPTY
If (UBound(anArray) >= 0) Then Exit Function
IS_EMPTY:
    IsArrayEmpty = True
End Function

'-----------------------------
Function LastUsedRow() As Long
'-----------------------------
' Return the last used row in the active sheet.
' Attention: Rows containing only format informations are also 'used' rows.
  LastUsedRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
End Function

'--------------------------------
Function LastUsedColumn() As Long
'--------------------------------
  LastUsedColumn = ActiveSheet.UsedRange.Columns(ActiveSheet.UsedRange.Columns.Count).Column
End Function

'------------------------------------
Function LastColumnDatSheet() As Long ' 27.11.21:
'------------------------------------
' Last column containing data in the data sheets
  LastColumnDatSheet = LED_Nr__Col + INTERNAL_COL_CNT - 1                   ' 03.04.21 Juergen - ignore hidden column storing the Max Led number per channel
End Function

'-----------------------------------------------
Function LastUsedRowIn(Sheet As Variant) As Long
'-----------------------------------------------
' return the last used row in the given sheet.
' The sheet could be given as sheet name or as worksheets variable.
Dim Sh As Variant
  If VarType(Sheet) = vbString Then
        Set Sh = Sheets(Sheet)
  Else: Set Sh = Sheet
  End If
  LastUsedRowIn = Sh.UsedRange.Rows(Sh.UsedRange.Rows.Count).Row
  Set Sh = Nothing
End Function

'---------------------------------------------------------------
Function LastUsedColumnInRow(Sh As Worksheet, Row As Long) As Long
'---------------------------------------------------------------
  LastUsedColumnInRow = Sh.Cells(Row, Sh.Columns.Count).End(xlToLeft).Column
End Function



'--------------------------------------------------
Function LastUsedColumnIn(Sheet As Variant) As Long
'--------------------------------------------------
Dim Sh As Variant
  If VarType(Sheet) = vbString Then
        Set Sh = Sheets(Sheet)
  Else: Set Sh = Sheet
  End If

  LastUsedColumnIn = Sh.UsedRange.Columns(Sh.UsedRange.Columns.Count).Column
  Set Sh = Nothing
End Function

'------------------------------------------------------------------
Function LastFilledRowIn(Sh As Worksheet, CheckCol As Long) As Long
'------------------------------------------------------------------
  Dim Row As Long
  Row = LastUsedRowIn(Sh)
  With Sh
    While .Cells(Row, CheckCol) = "" And Row > 0
      Row = Row - 1
    Wend
    LastFilledRowIn = Row
  End With
End Function

'-----------------------------------------------
Function LastFilledRow(CheckCol As Long) As Long
'-----------------------------------------------
  LastFilledRow = LastFilledRowIn(ActiveSheet, CheckCol)
End Function

'---------------------------------------------------------------------
Function LastFilledColumnIn(Sh As Worksheet, CheckRow As Long) As Long
'---------------------------------------------------------------------
  Dim Column As Long
  Column = LastUsedColumnIn(Sh)
  With Sh
    While .Cells(CheckRow, Column) = ""
      Column = Column - 1
      If Column = 0 Then Exit Function
    Wend
    LastFilledColumnIn = Column
  End With
End Function

'--------------------------------------------------
Function LastFilledColumn(CheckRow As Long) As Long
'--------------------------------------------------
  LastFilledColumn = LastFilledColumnIn(ActiveSheet, CheckRow)
End Function

'--------------------------------------------------------------
Function First_Change_in_Line(Target As Excel.Range) As Boolean
'--------------------------------------------------------------
' Check if the target cell is the only cell which contains data
  'First_Change_in_Line = Target.End(xlToLeft).Column = 1 And Target.End(xlToRight).Column = Target.Parent.Columns.Count ' 10.09.19: Removed: (Target = "") And  ' 28.10.19: Replaced ActiveSheet with Target.Parent
  First_Change_in_Line = Target.End(xlToLeft).Column = 1 And Target.End(xlToRight).Column = LED_Nr__Col ' 07.05.20: Adapted because the Start LedNr column is filled with ="" since 30.04.20
End Function

'-------------------------------------------------------------
Function LastFilledRowIn_ChkAll(ByVal Sh As Worksheet) As Long
'-------------------------------------------------------------
  Dim Row As Long
  Row = LastUsedRowIn(Sh)
  With Sh
    While First_Change_in_Line(.Cells(Row, 1))
      Row = Row - 1
      If Row = 0 Then Exit Function
    Wend
    LastFilledRowIn_ChkAll = Row
  End With
End Function

'UT--------------------------------------
Private Sub Test_LastFilledRowIn_ChkAll()
'UT--------------------------------------
  Debug.Print LastFilledRowIn_ChkAll(ActiveSheet)
End Sub


'----------------------------------------------------------------------
Function DelLast(ByVal s As String, Optional Cnt As Long = 1) As String
'----------------------------------------------------------------------
  If Len(s) > 0 Then
     DelLast = Left(s, Len(s) - Cnt)
  End If
End Function



'----------------------------------------------------------------
Function DelAllLast(ByVal s As String, Chars As String) As String
'----------------------------------------------------------------
  While InStr(Chars, Right(s, 1)) > 0
    s = Left(s, Len(s) - 1)
  Wend
  DelAllLast = s
End Function

'---------------------------
Sub Center_Form(f As Object)
'---------------------------
  With f
        .StartUpPosition = 0
        .Left = Application.Left + (Application.Width - .Width) / 2
        .Top = Application.Top + (Application.Height - .Height) / 2
        If .Top < Application.Top Then .Top = Application.Top               ' 02.03.20
        If .Left < Application.Left Then .Left = Application.Left
  End With
End Sub

'--------------------------------------------------------------
Sub Restore_Pos_or_Center_Form(f As Object, OldPos As WinPos_T)
'--------------------------------------------------------------
  If OldPos.Valid Then
        f.StartUpPosition = 0
        f.Left = OldPos.Left
        f.Top = OldPos.Top
  Else: Center_Form f
  End If
End Sub

'---------------------------------------------------
Sub Store_Pos(f As Object, ByRef PosVar As WinPos_T)
'---------------------------------------------------
     PosVar.Valid = True
     PosVar.Left = f.Left
     PosVar.Top = f.Top
End Sub



'UT---------------------------
Private Sub Test_Center_Form()
'UT---------------------------
  Center_Form UserForm_Other
  UserForm_Other.Show
End Sub

'----------------------------------------------------------
Function Replace_Multi_Space(ByVal Txt As String) As String
'----------------------------------------------------------
  Dim Res As String
  Res = Txt
  While InStr(Res, "  ") > 0
    Res = Replace(Res, "  ", " ")
  Wend
  Replace_Multi_Space = Res
End Function

'--------------------------------------------------
Function CellLinesSum(ByVal c As String) As Variant
'--------------------------------------------------
  If InStr(c, vbLf) Then
        Dim Line As Variant
        For Each Line In Split(c, vbLf)
           CellLinesSum = CellLinesSum + val(Line)
        Next Line
  Else: CellLinesSum = val(c)
  End If
End Function

#If 1 Then ' 05.05.20: New Function which uses a variant as search expression
'-------------------------------------------------------------------------------------
Public Function Is_Contained_in_Array(ByVal Txt As Variant, Arr As Variant) As Boolean
'-------------------------------------------------------------------------------------
  'If StrPtr(Arr) = 0 Then Exit Function                                     ' 06.05.20: Jürgen
  If Not isInitialised(Arr) Then Exit Function                              ' 06.05.20:
  Dim e As Variant
  Txt = Trim(Txt)
  For Each e In Arr
     If Trim(e) = Txt Then
        Is_Contained_in_Array = True
        Exit Function
     End If
  Next e
End Function
#Else
'------------------------------------------------------------------------------------
Public Function Is_Contained_in_Array(ByVal Txt As String, Arr As Variant) As Boolean
'------------------------------------------------------------------------------------
  Dim e As Variant
  Txt = Trim(Txt)
  For Each e In Arr
     If Trim(e) = Txt Then
        Is_Contained_in_Array = True
        Exit Function
     End If
  Next e
End Function
#End If

'----------------------------------------------------------------------------------
Public Function Get_Position_In_Array(ByVal Txt As Variant, Arr As Variant) As Long  ' 13.11.20:
'----------------------------------------------------------------------------------
  Get_Position_In_Array = -1
  If Not isInitialised(Arr) Then Exit Function
  Dim e As Variant, Nr As Long
  Nr = LBound(Arr)
  Txt = Trim(Txt)
  For Each e In Arr
     If Trim(e) = Txt Then
        Get_Position_In_Array = Nr
        Exit Function
     End If
     Nr = Nr + 1
  Next e
End Function


'--------------------------------------------------------------------------
Function IsInArray(ByVal stringToBeFound As String, Arr As Variant) As Long
'--------------------------------------------------------------------------
  Dim i As Long
  ' default return value if value not found in array
  IsInArray = -1

  For i = LBound(Arr) To UBound(Arr)
    If StrComp(stringToBeFound, Arr(i), vbTextCompare) = 0 Then
      IsInArray = i
      Exit For
    End If
  Next i
End Function

'--------------------------------------------------------------------------------------------------
Sub Hide_and_Move_up(dlg As Object, ByVal StartHide_Name As String, ByVal StartMove_Name As String)
'--------------------------------------------------------------------------------------------------
' Hide the controls where StartHide_y <= controls.Top < StartMove_y
' Move the controls up where controls.Top >= StartMove_y

  Dim MoveDelta As Long, StartHide_y As Single, StartMove_y As Single ' 06.12.20: Changed to single to fix problems with hight reduced dialog (Proc: Change_Height)
  StartHide_y = dlg.Controls(StartHide_Name).Top
  StartMove_y = dlg.Controls(StartMove_Name).Top
  MoveDelta = StartMove_y - StartHide_y
  
  'Debug.Print "Hide_and_Move_up from '" & StartHide_Name & "' to '" & StartMove_Name & "' " & MoveDelta ' Debug
  
  Dim c As Variant
  For Each c In dlg.Controls
      If c.Top >= StartMove_y Then
         c.Top = c.Top - MoveDelta
      ElseIf c.Top >= StartHide_y Then
         c.Visible = False
      End If
  Next c
  dlg.Height = dlg.Height - MoveDelta
End Sub


#If False Then ' 24.02.20: No longer used becase the head columns are translated also
'-------------------------------------------------------------------------
Function FindHeadCol(Sh As Worksheet, Row As Long, Name As Variant) As Long
'-------------------------------------------------------------------------
  Dim r As Range
  With Sh
    Set r = .Range(.Cells(Row, 1), .Cells(Row, LastUsedColumnIn(Sh)))
  End With
  
  #If 1 Then ' Use Match instead of find because the find command dosn't look in hidden cells (InCh)    ' 26.09.19
        Dim p As Variant
        p = Application.Match(Name, r, 0)
        If IsError(p) Then
           MsgBox Get_Language_Str("Fehler: Die Spalte '") & Name & Get_Language_Str("' wurde nicht im Sheet '") & Sh.Name & Get_Language_Str("' gefunden!" & vbCr & _
                  vbCr & _
                  "Die Spaltennamen dürfen nicht verändert werden"), vbCritical, Get_Language_Str("Fehler Spaltenname nicht gefunden")
           EndProg
        Else
           FindHeadCol = p
        End If
  #Else
        Dim f As Variant
        Set f = r.Find(What:=Name, after:=r.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
        
        If f Is Nothing Then
             MsgBox Get_Language_Str("Fehler: Die Spalte '") & Name & Get_Language_Str("' wurde nicht im Sheet '") & Sh.Name & Get_Language_Str("' gefunden!" & vbCr & _
                    vbCr & _
                    "Die Spaltennamen dürfen nicht verändert werden"), vbCritical, Get_Language_Str("Fehler Spaltenname nicht gefunden")
             EndProg
        Else
             FindHeadCol = f.Column
        End If
   #End If
End Function

'UT---------------------------
Private Sub Test_FindHeadCol()
'UT---------------------------
  Debug.Print FindHeadCol(ActiveSheet, 2, "Beschreibung")
End Sub
#End If

'--------------------------------------------------------------------------------------------------------
Function InputBoxMov(prompt As String, Optional Title As Variant, Optional Default As Variant, _
                     Optional Left As Variant, Optional Top As Variant, Optional helpfile As Variant, _
                     Optional HelpContextID As Variant) As Variant
'--------------------------------------------------------------------------------------------------------
' InputBox which could be moved with correct screen update even if screenupdating is disabled
  Dim OldUpdate As Boolean
  OldUpdate = Application.ScreenUpdating
  Application.ScreenUpdating = True
  InputBoxMov = InputBox(prompt, Title, Default, Left, Top, helpfile, HelpContextID)
  Sleep 50 ' Time to update the display
  Application.ScreenUpdating = OldUpdate
End Function

'-----------------------------
Private Sub Test_InputBoxMov()
'-----------------------------
  Application.ScreenUpdating = False

  Debug.Print InputBoxMov("Hallo", "Title", "Dafault")

  Application.ScreenUpdating = True
End Sub

'----------------------------------------------------------------------------------------
Function MsgBoxMov(prompt As String, Optional Buttons As Long, Optional Title As Variant, _
         Optional helpfile As Variant, Optional context As Variant) As Variant
'----------------------------------------------------------------------------------------
' MsgBox which could be moved with correct screen update even if screenupdating is disabled
  Dim OldUpdate As Boolean
  OldUpdate = Application.ScreenUpdating
  Application.ScreenUpdating = True
  MsgBoxMov = MsgBox(prompt, Buttons, Title, helpfile, context)
  Sleep 50 ' Time to update the display
  Application.ScreenUpdating = OldUpdate
End Function

'UT-------------------------
Private Sub Test_MsgBoxMov()
'UT-------------------------
  Application.ScreenUpdating = False

  Debug.Print MsgBoxMov("Hallo", vbYesNoCancel, "Titel")

  Application.ScreenUpdating = True
End Sub


'-----------------------------------------
Sub ShowHourGlassCursor(bApply As Boolean)                                  ' 07.10.21:
'-----------------------------------------
    #If HostProject = "Access" Then
        Application.DoCmd.Hourglass bApply
    #ElseIf HostProject = "Word" Then
        System.Cursor = IIf(bApply, wdCursorWait, wdCursorNormal)
    #Else
        Application.Cursor = IIf(bApply, xlWait, xlDefault)
    #End If

    #If Mac = False Then
        Dim pt As POINTAPI
        If Not bApply Then
            ' in some systems the cursor may fail to reset to default, this forces it
            GetCursorPos pt
            SetCursorPos pt.X, pt.Y
        End If
    #End If
End Sub

'---------------------------------------------
Public Function IsHourGlassCursor() As Boolean
'---------------------------------------------
  IsHourGlassCursor = (Application.Cursor = xlWait)
End Function


'-------------------
Public Sub EndProg()
'-------------------
' Is called in case of an fatal error
' Normaly this function should not be called because the
' global variables and dialog positions are cleared.
  ShowHourGlassCursor False                                                 ' 07.10.21:
  Application.EnableEvents = True
  Application.ScreenUpdating = True
  End
End Sub

'--------------------------
Public Sub ClearStatusbar()
'--------------------------
' Is called by onTime to clear the status bar after a while
  Application.StatusBar = ""
End Sub

'-------------------------------------------------------------------------------------------
Public Sub Show_Status_for_a_while(Txt As String, Optional Duration As String = "00:00:15")
'-------------------------------------------------------------------------------------------
  Application.StatusBar = Txt
  If Txt <> "" Then
        Application.OnTime Now + TimeValue(Duration), "ClearStatusbar"
  Else: Application.OnTime Now + TimeValue("00:00:00"), "ClearStatusbar"
  End If
End Sub


'-------------------------------------
Public Sub All_Borderlines(r As Range)
'-------------------------------------
    'r.Select ' Debug
    r.Borders(xlDiagonalDown).LineStyle = xlNone
    r.Borders(xlDiagonalUp).LineStyle = xlNone
    With r.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With r.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub

'---------------------------------------------------
Function FileNameExt(ByVal Name As String) As String
'---------------------------------------------------
' Return name and extention without path
Dim Pos As Long, Pos2 As Long, Temp As String
  Pos = InStrRev(Name, "\")
  Pos2 = InStrRev(Name, "/")
  If Pos2 > Pos Then Pos = Pos2
  If Pos > 0 Then
     Temp = Mid(Name, Pos + 1)
  Else
     Pos = InStrRev(Name, ":")
     If Pos > 0 Then
          Temp = Mid(Name, Pos + 1)
     Else
          Temp = Name
     End If
  End If
  FileNameExt = Temp
End Function

'------------------------------------------------
Function FilePath(ByVal Name As String) As String
'------------------------------------------------
  FilePath = Left(Name, Len(Name) - Len(FileNameExt(Name)))
End Function

'---------------------------------------------
Function NoExt(ByVal Name As String) As String
'---------------------------------------------
' Cut of the extention of a filename
Dim Pos As Long
  Pos = InStrRev(Name, ".")
  If Pos > 0 Then
        NoExt = Left(Name, Pos - 1)
  Else: NoExt = Name
  End If
End Function

'------------------------------------------------
Function FileName(ByVal Name As String) As String
'------------------------------------------------
' Return name without extention and path
  FileName = NoExt(FileNameExt(Name))
End Function

'-------------------------------------------------------------------
Function Same_Name_already_open(ByVal FullName As String) As Boolean
'-------------------------------------------------------------------
' Check if a workbook with the same name is already opened
Dim w, Name As String
  Name = FileNameExt(FullName)
  For Each w In Workbooks
    If UCase(w.Name) = UCase(Name) Then
       Same_Name_already_open = True
       Exit Function
    End If
  Next w
End Function

'-------------------------------
Function SheetEx(Name As String)
'-------------------------------
  Dim s As Variant
  For Each s In Sheets
     If s.Name = Name Then
        SheetEx = True
        Exit Function
     End If
  Next
End Function


'-------------------------
Sub Protect_Active_Sheet()
'-------------------------
  ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingCells:=True, AllowInsertingHyperlinks:=True
End Sub


'-------------------------------------------
Function ColumnLetters(r As Range) As String
'-------------------------------------------
  ColumnLetters = Replace(Replace(r.Address, "$", ""), r.Row, "")
End Function

'-------------------------------------------------
Function ColumnLettersFromNr(ByVal Colunm As Long)
'-------------------------------------------------
  ColumnLettersFromNr = ColumnLetters(Cells(1, Colunm))
End Function

'----------------------------------------
Sub DisableFiltersInSheet(s As Worksheet)
'----------------------------------------
    If s.AutoFilterMode Then
        If s.FilterMode Then
            On Error Resume Next ' Generates an error if all rows have been deleted
            s.ShowAllData
            On Error GoTo 0
            Exit Sub
        End If
    End If
        
    ' Check if a table is used. In this case the filter can't be disabled if the active cell is not located in the table
    Dim obj As Variant
    For Each obj In s.ListObjects
        If Not obj.AutoFilter Is Nothing Then
           Dim OldActCell As Range
           Set OldActCell = ActiveCell
           s.Cells(obj.Range.Row, obj.Range.Column).Select
           On Error Resume Next ' Generates an error if all rows have been deleted
           s.ShowAllData
           On Error GoTo 0
           OldActCell.Select
           Exit Sub
        End If
    Next obj
End Sub

'-----------------------------------------------
Function isVariantArray(v As Variant) As Boolean
'-----------------------------------------------
  On Error GoTo ErrorDet
  Dim i As Long
  i = UBound(v)
  isVariantArray = True

ErrorDet:
  On Error GoTo 0
End Function



'----------------------------------------------------
Public Function F_shellExec(sCmd As String) As String
'----------------------------------------------------
' Excecute command and get the output as string
' Example call:
'   MsgBox F_shellExec("cmd /c dir c:\")
'
' Requires ref to Windows Script Host Object Model
' To do this go to Extras -> References in the VBA IDE's menu bar.
' See:
'   https://stackoverflow.com/questions/2784367/capture-output-value-from-a-shell-command-in-vba

    Dim oShell As New WshShell 'requires ref to Windows Script Host Object Model
    F_shellExec = oShell.Exec(sCmd).StdOut.ReadAll
End Function

'----------------------------------------------------------------------------------------
Public Sub F_shellRun(sCmd As String, ByVal WindowMode As Integer, ByVal Wait As Boolean)
'----------------------------------------------------------------------------------------
' Excecute command
' Example call:
'   MsgBox F_shellRun("cmd /c dir c:\", 0, true)
'
    '0: versteckt das Fenster und aktiviert ein anderes
    '1: aktiviert und zeigt ein Fenster
    '2: aktiviert und minimiert das Fenster
    '3: aktiviert und maximiert das Fenster
    '4: zeigt das Fenster in seiner letzen Position, das aktive Fenster bleibt aktiv
    '5: zeigt das Fenster in seiner letzen grösse und Position
    '6: minimiert das Fenster und aktiviert ein anderes
    '7: minimiert das Fenster, das aktive Fenster bleibt aktiv
    '8: zeigt das Fenster in seiner letzen Position, das aktive Fenster bleibt aktiv
    '9: stellt ein minimiertes Fenster wieder in seinen ursprünglichen Zustand
    '10: setzt das Fenster gleich dem Programm
    
    Dim oShell As New WshShell 'requires ref to Windows Script Host Object Model
    oShell.Run sCmd, WindowMode, Wait
End Sub


'UT---------------------
Private Sub Test_Shell()
'UT---------------------
  MsgBox F_shellExec("cmd /c Dir C:\")
End Sub


'----------------------------------------------------------------
Public Function Read_File_to_String(FileName As String) As String
'----------------------------------------------------------------
  Dim strFileContent As String
  Dim fp As Integer: fp = FreeFile
  On Error GoTo ReadError
  Open FileName For Input As #fp
  Read_File_to_String = Input(LOF(fp), fp)
  Close #fp
  Exit Function

ReadError:
  MsgBox Get_Language_Str("Fehler beim lesen der Datei:") & vbCr & _
         "  '" & FileName & "'", vbCritical, Get_Language_Str("Fehler beim Datei lesen")
  Read_File_to_String = "#ERROR#"
End Function

'------------------------------------------------------------------------------
Public Function Get_Ini_Entry(FileStr As String, EntryName As String) As String
'------------------------------------------------------------------------------
  Dim p As Long, e As String
  Get_Ini_Entry = "#ERROR#"
  p = InStr(FileStr, EntryName)
  If p = 0 Then Exit Function
  p = p + Len(EntryName)
  e = InStr(p, FileStr, vbCr)
  If e = 0 Then e = InStr(p, FileStr, vbLf)
  If e = 0 Then Exit Function
  Get_Ini_Entry = Mid(FileStr, p, e - p)
End Function


'-----------------------------------------------------------------------------------
Public Sub Debug_Print_Arr(ByRef Arr As Variant, Optional ArrName As String = "arr")
'-----------------------------------------------------------------------------------
  Dim i As Long
  For i = 0 To UBound(Arr)
      Debug.Print ArrName & "(" & i & ")=" & Arr(i)
  Next i
End Sub

'-------------------------------------------------------------------------
Public Sub DeleteElementAt(ByVal Index As Integer, ByRef prLst As Variant)
'-------------------------------------------------------------------------
   Dim i As Long

   ' Move all element back one position
   For i = Index + 1 To UBound(prLst)
       prLst(i - 1) = prLst(i)
   Next

   ' Shrink the array by one, removing the last one
   ReDim Preserve prLst(UBound(prLst) - 1)
End Sub

'UT-------------------------------
Private Sub Test_DeleteElementAt()
'UT-------------------------------
  Dim Arr() As String
  Arr = Split("A B C D E", " ")
  DeleteElementAt 1, Arr
  Debug.Print ""
  Debug_Print_Arr Arr, "arr"
End Sub



'-----------------------------------------------------------------------------------------------
Public Sub InsertElementAt(ByVal Index As Integer, ByRef prLst As Variant, InsertVal As Variant)
'-----------------------------------------------------------------------------------------------
' index could be 0 .. UBound(prLst)+1
' If index is > UBound(prLst)+1 the function will crash
' Tested with prList() string and prList() interger
  Dim i As Long
  ReDim Preserve prLst(UBound(prLst) + 1)
  For i = UBound(prLst) To Index + 1 Step -1
       prLst(i) = prLst(i - 1)
  Next
  prLst(Index) = InsertVal
End Sub

'UT-------------------------------
Private Sub Test_InsertElementAt()
'UT-------------------------------
  Dim Arr() As String
  Arr = Split("1 2 3 4 5", " ")
  InsertElementAt 0, Arr, 0
  Debug.Print ""
  Debug_Print_Arr Arr, "arr"
  
  
  Dim iarr() As Integer, i As Long
  ReDim iarr(UBound(Arr))
  For i = 0 To UBound(Arr)
      iarr(i) = Arr(i)
  Next i
  InsertElementAt 3, iarr, -1
  Debug_Print_Arr iarr, "iarr"
End Sub

'------------------------------------------------------
Private Function GetPathOnly(sPath As String) As String
'------------------------------------------------------
 GetPathOnly = Left(sPath, InStrRev(sPath, "\", Len(sPath)) - 1)
End Function


'--------------------------------------------------------
Public Function CreateFolder(sFolder As String) As String
'--------------------------------------------------------
' http://www.freevbcode.com/ShowCode.asp?ID=257
' sFolder must have an "\" at the end
On Error GoTo ErrorHandler
Dim s As String
s = GetPathOnly(sFolder)
If Dir(s, vbDirectory) = "" Then
  s = CreateFolder(s)
  MkDir s
End If
CreateFolder = sFolder
Exit Function

ErrorHandler:
Exit Function
End Function


'-------------------------------------------------------------------------------------------------------
Public Function UnzipAFile(ByVal zippedFileFullName As Variant, ByVal unzipToPath As Variant) As Boolean
'-------------------------------------------------------------------------------------------------------
' The Destination directory must exist
' The Arguments must be "byVal" and "Variant" otherwise the program fails
  
  Dim ShellApp As Object
    
  'Copy the files & folders from the zip into a folder
  Set ShellApp = CreateObject("Shell.Application")
  On Error GoTo ErrMsg
  ShellApp.Namespace(unzipToPath).CopyHere ShellApp.Namespace(zippedFileFullName).Items
  On Error GoTo 0
  UnzipAFile = True
  Exit Function

ErrMsg:
  MsgBox Replace(Replace(Get_Language_Str("Fehler beim entpacken der ZIP-Datei:" & vbCr & _
                          "  '#1#'" & vbCr & _
                          "nach" & vbCr & _
                          "  '#2#'"), "#1#", zippedFileFullName), "#2#", unzipToPath), _
                          vbCritical, Get_Language_Str("Fehler: Zip-Datei konnte nicht entpackt werden")
End Function

'----------------------------------------------------
Function isInitialised(ByRef a As Variant) As Boolean
'----------------------------------------------------
' Check if an array in initialized
' This is usefull for functions which return an array
' in case they fail
    On Error Resume Next
    isInitialised = IsNumeric(UBound(a))
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------------
Function SplitMultiDelims(ByVal Text As String, DelimChars As String) As String()
'--------------------------------------------------------------------------------
' SplitMutliChar
' This function splits Text into an array of substrings, each substring
' delimited by any character in DelimChars. Only a single character
' may be a delimiter between two substrings, but DelimChars may
' contain any number of delimiter characters. It returns
' an unallocated array it Text is empty, a single element array
' containing all of text if DelimChars is empty, or a 1 or greater
' element array if the Text is successfully split into substrings.
'
' http://www.cpearson.com/excel/splitondelimiters.aspx
'
' Adapted by Hardi to
' - skip multiple delimiters between two parts
' - generate an array starting wit 0 like the split function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Pos1 As Long
Dim N As Long
Dim M As Long
Dim Arr() As String
Dim i As Long
Dim TextLen As Long

TextLen = Len(Text)

''''''''''''''''''''''''''''''''
' if Text is empty, get out
''''''''''''''''''''''''''''''''
If TextLen = 0 Then
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''
' if DelimChars is empty, return original text
'''''''''''''''''''''''''''''''''''''''''''''
If DelimChars = vbNullString Then
    SplitMultiDelims = Array(Text)
    Exit Function
End If

'''''''''''''''''''''''''''''''''''''''''''''''
' oversize the array, we'll shrink it later so
' we don't need to use Redim Preserve
'''''''''''''''''''''''''''''''''''''''''''''''
ReDim Arr(0 To Len(Text) - 1)

i = 0
N = 1

While N <= TextLen And InStr(DelimChars, Mid(Text, N, 1)) > 0 ' Skip leading delimiters
   N = N + 1
Wend
Pos1 = N

N = N + 1
While N <= TextLen
  If InStr(DelimChars, Mid(Text, N, 1)) > 0 Then
     Arr(i) = Mid(Text, Pos1, N - Pos1)
     i = i + 1
     While N <= TextLen And InStr(DelimChars, Mid(Text, N, 1)) > 0 ' Skip leading delimiters
       N = N + 1
     Wend
     Pos1 = N
  End If
  N = N + 1
Wend

If Pos1 <= Len(Text) Then
    Arr(i) = Mid(Text, Pos1)
    i = i + 1
End If

''''''''''''''''''''''''''''''''''''''
' chop off unused array elements
''''''''''''''''''''''''''''''''''''''
If i >= 1 Then
   ReDim Preserve Arr(0 To i - 1)
   SplitMultiDelims = Arr
End If
    
End Function


'UT--------------------------------
Private Sub Test_SplitMultiDelims()
'UT--------------------------------
  Debug.Print "Test_SplitMultiDelims"
  Dim Res() As String, i As Long
  Res = SplitMultiDelims("  Text  with+several|delimmiters +|", " +|")
  'Res = SplitMultiDelims("a", " +|")
  If isInitialised(Res) Then
     For i = 0 To UBound(Res)
        Debug.Print "'" & Res(i) & "'"
     Next i
  End If
  Debug.Print "---"
End Sub


'-----------------------------------------------------------------------------------------------------------------------
Function SplitEx(ByVal InString As String, IgnoreDoubleDelmiters As Boolean, ParamArray Delims() As Variant) As String()
'-----------------------------------------------------------------------------------------------------------------------
' http://www.cpearson.com/excel/splitondelimiters.aspx
    Dim Arr() As String
    Dim Ndx As Long
    Dim N As Long
    
    If Len(InString) = 0 Then
        SplitEx = Arr
        Exit Function
    End If
    If IgnoreDoubleDelmiters = True Then
        For Ndx = LBound(Delims) To UBound(Delims)
            N = InStr(1, InString, Delims(Ndx) & Delims(Ndx), vbTextCompare)
            Do Until N = 0
                InString = Replace(InString, Delims(Ndx) & Delims(Ndx), Delims(Ndx))
                N = InStr(1, InString, Delims(Ndx) & Delims(Ndx), vbTextCompare)
            Loop
        Next Ndx
    End If
    
    
    ReDim Arr(1 To Len(InString))
    For Ndx = LBound(Delims) To UBound(Delims)
        InString = Replace(InString, Delims(Ndx), Chr(1))
    Next Ndx
    Arr = Split(InString, Chr(1))
    SplitEx = Arr
End Function

'UT-----------------------
Private Sub Test_SplitEx()
'UT-----------------------
' Attention: The result contains space characters
    Dim s As String
    Dim t() As String
    Dim N As Long
    'S = "A AND #InCh OR A AND NOT #InCh + 1 OR D"
    'S = "#InCh"
    t = SplitEx(s, True, "OR", "AND", "NOT")
    If isInitialised(t) Then
       For N = LBound(t) To UBound(t)
           Debug.Print N, t(N)
       Next N
    Else: Debug.Print "Empty"
    End If
End Sub


'--------------------------------------------------------
Public Function Get_Primary_Monitor_Pixel_Cnt_X() As Long                   ' 23.04.20:
'--------------------------------------------------------
  Get_Primary_Monitor_Pixel_Cnt_X = GetSystemMetrics(SM_CXSCREEN)
End Function

'--------------------------------------------------------
Public Function Get_Primary_Monitor_Pixel_Cnt_Y() As Long                   ' 23.04.20:
'--------------------------------------------------------
  Get_Primary_Monitor_Pixel_Cnt_Y = GetSystemMetrics(SM_CYSCREEN)
End Function


'------------------
Sub ResetComments()
'------------------
' https://www.heise.de/ct/hotline/Excel-Kommentare-automatisch-positionieren-2055961.html
    Dim objComment As Comment

    ' Alle Kommentare des aktuellen Arbeitsblatts
    ' durchlaufen
    For Each objComment In ActiveSheet.Comments
        With objComment
            ' Top-Wert des Kommentars auf Top-Wert
            ' der verknüpften Zelle setzen
            .Shape.Top = .Parent.Top + .Parent.Height - .Shape.Height
            ' Left-Wert des Kommentars auf Left-Wert
            ' der verknüpften Zelle plus Zellbreite
            ' mal zwei setzen
            .Shape.Left = .Parent.Left + (.Parent.Width * 2)
        End With
    Next
End Sub

'---------------------------------------------------------
Public Sub Array_BubbleSort(ByRef vArrayName As Variant, _
                   Optional ByVal lUpper As Long = -1, _
                   Optional ByVal lLower As Long = -1)
'---------------------------------------------------------
' https://bettersolutions.com/vba/arrays/sorting-bubble-sort.htm
Dim vtemp As Variant
Dim i As Long
Dim j As Long

   If IsEmpty(vArrayName) = True Then Exit Sub
   If lLower = -1 Then lLower = LBound(vArrayName, 1)
   If lUpper = -1 Then lUpper = UBound(vArrayName, 1)

   For i = lLower To (lUpper - 1)
      For j = i To lUpper
         If (vArrayName(j) < vArrayName(i)) Then
            vtemp = vArrayName(i)
            vArrayName(i) = vArrayName(j)
            vArrayName(j) = vtemp
         End If
      Next j
   Next i
End Sub

'---------------------------------------------------------------------
Public Sub Button_Setup(Button As CommandButton, ByVal Text As String)
'---------------------------------------------------------------------
' If text is empty the button is not shown
  Text = Trim(Text)
  Dim Err As Boolean
  Button.Visible = (Text <> "")
  If Text <> "" Then
     Err = (Len(Text) < 3)
     If Not Err Then Err = (Mid(Text, 2, 1) <> " ")
     If Err Then
        MsgBox "Internal Error: Button text is wrong '" & Text & "'." & vbCr & _
               "It must contain an Accelerator followed by the text." & vbCr & _
               "Example: 'H Hallo'", vbCritical, "Internal Error (Wrong translation?)"
        EndProg
     End If
     Button.Caption = Mid(Text, 3, 255)
     Button.Accelerator = Left(Text, 1)
  End If
End Sub


#If VBA7 Then
'----------------------------------------
Public Sub Bring_to_front(hwnd As LongPtr)                                   ' 05.06.20: Added hWnd => Now it's working (Thanks to Jürgen)
'----------------------------------------
#Else
Public Sub Bring_to_front(hwnd As Long)
#End If
' Is not working if an other application has be moved above Excel with Alt+Tab
' But this is a feature od Windows.
' See: https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow
' But it brings up excel again after the upload to the Arduino
' Without this funchion an other program was activated after the upload for some reasons
    ThisWorkbook.Activate
    SetForegroundWindow hwnd
End Sub

'--------------------------------------------------------------------
Public Function Replicate(RepeatString As String, NumOfTimes As Long)
'--------------------------------------------------------------------
    Dim s As String
    Dim c As Long
    Dim l As Long
    Dim i As Long

    l = Len(RepeatString)
    c = l * NumOfTimes
    s = Space$(c)

    For i = 1 To c Step l
        Mid(s, i, l) = RepeatString
    Next

    Replicate = s
End Function


' https://www.herber.de/forum/archiv/916to920/917548_Textdateien_von_UTF8_nach_Ansi_konvertieren.html

'---------------------------------------------------------------
Public Function ConvertToUTF8(ByRef Source As String) As Byte()
'---------------------------------------------------------------
    Dim Length As Long
    #If VBA7 Then
      Dim Pointer As LongPtr
    #Else
      Dim Pointer As Long
    #End If
    Dim Size As Long
    Dim Buffer() As Byte

    Length = Len(Source)
    Pointer = StrPtr(Source)
    Size = WideCharToMultiByte(CP_UTF8, 0, Pointer, Length, 0, 0, 0, 0)
    ReDim Buffer(0 To Size - 1)

    WideCharToMultiByte CP_UTF8, 0, Pointer, Length, VarPtr(Buffer(0)), Size, 0, 0

    ConvertToUTF8 = Buffer
End Function

'-----------------------------------------------------------------
Public Function ConvertToUTF8Str(ByRef Source As String) As String
'-----------------------------------------------------------------
  Dim Bytes() As Byte, Res As String, i
  Bytes = ConvertToUTF8(Source)
  For i = 0 To UBound(Bytes)
      Res = Res & Chr(Bytes(i))
  Next
  ConvertToUTF8Str = Res
End Function


'----------------------------------------------------------------
Public Function ConvertFromUTF8(ByRef Source() As Byte) As String           ' 26.05.20:
'----------------------------------------------------------------
    Dim Size As Long
    #If VBA7 Then                                                           ' 28.05.20:
      Dim Pointer As LongPtr
    #Else
      Dim Pointer As Long
    #End If
    Dim Length As Long
    Dim Buffer As String

    Size = UBound(Source) - LBound(Source) + 1
    Pointer = VarPtr(Source(LBound(Source)))
    Length = MultiByteToWideChar(CP_UTF8, 0, Pointer, Size, 0, 0)
    Buffer = Space$(Length)
    MultiByteToWideChar CP_UTF8, 0, Pointer, Size, StrPtr(Buffer), Length
    ConvertFromUTF8 = Buffer
End Function

'----------------------------------------------------------
Public Function ConvertUTF8Str(UTF8Str As String) As String                 ' 26.05.20:
'----------------------------------------------------------
  Dim bStr() As Byte, i As Long
  ReDim bStr(Len(UTF8Str) - 1)
  For i = 1 To Len(UTF8Str)
      bStr(i - 1) = Asc(Mid(UTF8Str, i, 1))
  Next
  ConvertUTF8Str = ConvertFromUTF8(bStr)
End Function


'---------------------------------------------------------
Public Function Dir_is_Empty(DirName As String) As Boolean                  ' 27.05.20:
'---------------------------------------------------------
' Return true it the directoriy contains at least one subdirectory or one file

  Dim Res As String
  Res = Dir(DirName & "\*.*", vbDirectory)
  While Res <> ""
     If Res <> "" And Left(Res, 1) <> "." Then
        Exit Function ' It's not empty => return false
     End If
     Res = Dir() ' Mit Excel für Mac 2016 wird der ursprüngliche Dir-Funktionsaufruf erfolgreich ausgeführt. Nachfolgende Aufrufe zum Durchlaufen des angegebenen Verzeichnisses führen jedoch zu einem Fehler. Dies ist leider ein bekanntes Problem.
  Wend
  
  Res = Dir(DirName & "\*.*")
  While Res <> ""
     If Res <> "" And Left(Res, 1) <> "." Then
        Exit Function ' It's not empty => return false
     End If
     Res = Dir() ' Mit Excel für Mac 2016 wird der ursprüngliche Dir-Funktionsaufruf erfolgreich ausgeführt. Nachfolgende Aufrufe zum Durchlaufen des angegebenen Verzeichnisses führen jedoch zu einem Fehler. Dies ist leider ein bekanntes Problem.
  Wend
  
  Dir_is_Empty = True
End Function

'------------------------------------------------------------
Public Function Get_First_SubDir(DirName As String) As String               ' 27.05.20:
'------------------------------------------------------------
  Dim Res As String
  Res = Dir(DirName & "\*.*", vbDirectory)
  While Res <> ""
     If Res <> "" And Left(Res, 1) <> "." Then
        Get_First_SubDir = Res
        Exit Function
     End If
     Res = Dir() ' Mit Excel für Mac 2016 wird der ursprüngliche Dir-Funktionsaufruf erfolgreich ausgeführt. Nachfolgende Aufrufe zum Durchlaufen des angegebenen Verzeichnisses führen jedoch zu einem Fehler. Dies ist leider ein bekanntes Problem.
  Wend
End Function

'--------------------------------------------------------------------------------------------------------------------
Public Function VersionStr_is_Greater(Ver1 As String, Ver2 As String, Optional Delimmiter As String = ".") As Boolean
'--------------------------------------------------------------------------------------------------------------------
' Compares two version strings like
'  "1.0.7"
' If one string is shorter than the other the missing digits are replaced by 0
' "1.0" => "1.0.0"
  Dim Ver1A() As String, Ver2A() As String, EndNr As Long, Nr As Long
  Ver1A = Split(Ver1, Delimmiter)
  Ver2A = Split(Ver2, Delimmiter)
  EndNr = WorksheetFunction.Max(UBound(Ver1A), UBound(Ver2A))
  For Nr = 0 To EndNr
     Dim v1 As Long, v2 As Long
     If UBound(Ver1A) >= Nr Then
           v1 = val(Ver1A(Nr))
     Else: v1 = 0
     End If
     If UBound(Ver2A) >= Nr Then
           v2 = val(Ver2A(Nr))
     Else: v2 = 0
     End If
     If v1 <> v2 Then
        VersionStr_is_Greater = v1 > v2
        Exit Function
     End If
  Next Nr
End Function

'UT-------------------------------------
Private Sub Test_VersionStr_is_Greater()
'UT-------------------------------------
  'Debug.Print VersionStr_is_Greater("1.0.7", "1.03.1")
  'Debug.Print VersionStr_is_Greater("1.0.7", "1.0.1")
  'Debug.Print VersionStr_is_Greater("1.0.7", "1.0.8")
  'Debug.Print VersionStr_is_Greater("2.0.7", "")
  Debug.Print VersionStr_is_Greater("1.0.8", "1.0.7b")
End Sub

'---------------------------------------------------------------------------------------------------
Public Function Del_Folder(ByVal DirName As String, Optional ShowError As Boolean = True) As Boolean
'---------------------------------------------------------------------------------------------------
  On Error GoTo ErrMsg
  CreateObject("Scripting.FileSystemObject").DeleteFolder DirName
  On Error GoTo 0
  Del_Folder = True
  Exit Function

ErrMsg:
  If ShowError Then
     MsgBox Replace(Get_Language_Str("Fehler beim Löschen des Verzeichnisses:" & vbCr & _
                             "  '#1#'" & vbCr & _
                             "Evtl. enthält es Dateien welche in einem anderen Programm geöffnet sind oder es ist " & _
                             "das Arbeitsverzeichnis eines Programms und darf darum nicht gelöscht werden."), "#1#", DirName), _
                             vbCritical, Get_Language_Str("Verzeichnis konnte nicht (vollständig) gelöscht werden")
  End If
End Function


'----------------------------------------------
Public Function Get_OperatingSystem() As String
'----------------------------------------------
' Returns something like:
'   objOperatingSystem.Caption        objOperatingSystem.Version
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~        ~~~~~~~~~~~~~~~~~~~~~~~~~~
'  "Microsoft Windows 10 Home         10.0.18362"
'  "Microsoft Windows 8.1 Pro         6.3.9600"
'  "Microsoft Windows 7 Home Premium  6.1.7601"

  Dim localHost           As String
  Dim objWMIService       As Variant
  Dim colOperatingSystems As Variant
  Dim objOperatingSystem  As Variant
  On Error GoTo Error_Handler
  localHost = "." 'Technically could be run against remote computers, if allowed
  Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & localHost & "\root\cimv2")
  Set colOperatingSystems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
  For Each objOperatingSystem In colOperatingSystems
      Get_OperatingSystem = objOperatingSystem.Caption & vbTab & objOperatingSystem.Version
      Exit Function
  Next

Error_Handler_Exit:
  On Error Resume Next
  Exit Function

Error_Handler:
  MsgBox "The following error has occured." & vbCrLf & vbCrLf & _
         "Error Number: " & Err.Number & vbCrLf & _
         "Error Source: getOperatingSystem" & vbCrLf & _
         "Error Description: " & Err.Description, _
          vbCritical, "An Error has Occured!"
  Resume Error_Handler_Exit
End Function

'------------------------------------------
Public Function Win10_or_newer() As Boolean
'------------------------------------------
  Dim OpSys As String, Nr As Double
  OpSys = Get_OperatingSystem
  Nr = val(Mid(OpSys, Len("Microsoft Windows ")))
  Win10_or_newer = (Nr >= 10)
End Function

                                                   
' 22.11.21: Juergen
'------------------------------------------
Public Function Check_Version() As Boolean
  If Not Valid_Excel Then
      Dim message
      message = Replace(Get_Language_Str("Diese Excel Version wird nicht unterstützt." & _
      "Bitte besuchen sie die Webseite #1# für weitergehende Informationen." & _
      "Das Programm wird weiter ausgeführt, es kann jedoch zu unerwarteten Fehlfunktionen" & _
      ", Fehlermeldung und Abstürzen kommen."), "#1#", _
      vbCrLf & vbCrLf & "https://wiki.mobaledlib.de/anleitungen/programmgenerator" & vbCrLf & vbCrLf)
      
      MsgBox message, vbCritical, Get_Language_Str("Versionsprüfung")
  End If
End Function
                                                   
' 22.11.21: Juergen
'------------------------------------------
Public Function Valid_Excel() As Boolean
'------------------------------------------
  ' see details on https://wiki.mobaledlib.de/anleitungen/programmgenerator
  ' Excel Version on https://de.wikipedia.org/wiki/Microsoft_Excel
  Dim exVer
  
#If VBA7 Then
  exVer = val(Application.Version)
  Valid_Excel = exVer >= 15
#Else
  Valid_Excel = False
#End If
  
End Function

'UT-----------------------------------
Private Sub Test_Get_OperatingSystem()
'UT-----------------------------------
  Debug.Print "Get_OperatingSystem:" & Get_OperatingSystem
End Sub


'------------------------------------
Private Sub Test_Select_LastusedCol()
'------------------------------------
  Cells(1, LastUsedColumn()).Select
End Sub


' 14.10.21 Juergen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Clear_Platform_Parameter_Cache()
    Set PlatformParams = Nothing
End Sub


' 14.10.21 Juergen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Get_Current_Platform_String(ByVal ParName As String, Optional ByVal EmptyCheck As Boolean = False, Optional Silent As Boolean = False) As String
    Get_Current_Platform_String = Get_Platform_String(Get_BoardTyp(), ParName, EmptyCheck, Silent)
End Function

' 14.10.21 Juergen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Get_Current_Platform_Bool(ByVal ParName As String, Optional Silent As Boolean = False) As Boolean
    Get_Current_Platform_Bool = Get_Platform_Bool(Get_BoardTyp(), ParName, Silent)
End Function

' 14.10.21 Juergen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Get_Current_Platform_Int(ByVal ParName As String, Optional Silent As Boolean = False) As Integer
    Get_Current_Platform_Int = Get_Platform_Int(Get_BoardTyp(), ParName, Silent)
End Function


' 14.10.21 Juergen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Get_Platform_String(ByVal PlatformKey As String, ByVal ParName As String, Optional ByVal EmptyCheck As Boolean = False, Optional ByVal Silent As Boolean = False) As String
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
  Get_Platform_String = ""
  Dim Row As Long, Sh As Worksheet, ActLanguage As Integer, Offs As Long
  
  If PlatformParams Is Nothing Then
    Set PlatformParams = New Scripting.Dictionary
  End If
  If PlatformParams.Exists(PlatformKey + "|" + ParName) Then
    Get_Platform_String = PlatformParams(PlatformKey + "|" + ParName)
    Exit Function
  End If
  
  Set Sh = Sheets(PLATFORMS_SH)
  Dim r As Range, f As Variant
  Set r = Sh.Range(Sh.Cells(PlatformKey_ROW, PlatformKey_COL), Sh.Cells(PlatformKey_ROW, PlatformKey_COL + 99))
  Set f = r.Find(What:=PlatformKey, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
    
  If f Is Nothing Then
       Debug.Print "Fehlende Plattform " & PlatformKey
       If Not Silent Then
            MsgBox Replace(Get_Language_Str("Fehler: Die Plattform '#1#' ist nicht definiert."), _
            "#1#", PlatformKey), vbCritical, Get_Language_Str("Internal Error")
       End If
       Exit Function
  End If
  
  Dim Platform_COL As Integer
  Platform_COL = f.Column
  
  Set r = Sh.Range(Sh.Cells(PlatformKey_ROW + 1, PlatformParName_COL), Sh.Cells(LastUsedRowIn(Sh), PlatformParName_COL))
  Set f = r.Find(What:=ParName, LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=True, SearchFormat:=False)
  
  If f Is Nothing Then
      If Not Silent Then
         Show_Missing_Platform_Parameter_Error PlatformKey, ParName
      End If
      Exit Function
  End If
  If EmptyCheck And Sh.Cells(f.Row, Platform_COL) = "" Then
       Debug.Print "Der Parameter" & ParName & " für Plattform " & PlatformKey & " darf nicht leer sein"
       If Not Silent Then
          Show_Invalid_Platform_Parameter_Error PlatformKey, ParName
       End If
       Exit Function
  End If
  
  Get_Platform_String = Sh.Cells(f.Row, Platform_COL)
  
  ' values starting with "=" indicate an indirection, get the referenced value
  If Left(Get_Platform_String, 1) = "=" Then
     Get_Platform_String = Get_Platform_String(PlatformKey, Mid(Get_Platform_String, 2), EmptyCheck, Silent)
  End If
  PlatformParams.Add PlatformKey + "|" + ParName, Get_Platform_String

End Function

' 14.10.21 Juergen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Get_Platform_Bool(ByVal PlatformKey As String, ByVal ParName As String, Optional Silent As Boolean = False) As Boolean
    Dim Value As String
    Get_Platform_Bool = False
    Value = Get_Platform_String(PlatformKey, ParName, True, Silent)
    If UCase$(Value) = "TRUE" Or Value = "1" Then
        Get_Platform_Bool = True
    ElseIf Not UCase$(Value) = "FALSE" And Value <> "0" Then
       Debug.Print "Der Parameter" & ParName & " für Plattform " & PlatformKey & " ist weder true noch false"
       Show_Invalid_Platform_Parameter_Error PlatformKey, ParName
       Exit Function
    End If

End Function

' 14.10.21 Juergen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function Get_Platform_Int(ByVal PlatformKey As String, ByVal ParName As String, Optional Silent As Boolean = False) As Integer
    Get_Platform_Int = 0
    Dim Value As String
    Value = Get_Platform_String(PlatformKey, ParName, True, Silent)
    If Not IsNumeric(Value) Then
       Debug.Print "Der Parameter" & ParName & " für Plattform " & PlatformKey & " ist nicht numerisch"
       Show_Invalid_Platform_Parameter_Error PlatformKey, ParName
       Exit Function
    End If
    Get_Platform_Int = val(Value)
End Function

' 14.10.21 Juergen
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function AliasToPin(ByVal Pin As String) As String

    AliasToPin = Get_Current_Platform_String("PIN_ALIAS_" + UCase(Pin), False, True)
    
    If AliasToPin = "" Then
        ' not a valid alias definition
        AliasToPin = Pin
    End If
    
End Function

Private Sub Show_Invalid_Platform_Parameter_Error(ByVal PlatformKey As String, ByVal ParName As String)
    MsgBox Replace(Replace(Get_Language_Str("Fehler: Der Parameter '#2#' für die Plattform '#1#' hat keinen gültigen Wert."), _
       "#1#", PlatformKey), "#2#", ParName), vbCritical, Get_Language_Str("Parameter Fehler")
End Sub

Private Sub Show_Missing_Platform_Parameter_Error(ByVal PlatformKey As String, ByVal ParName As String)
    Debug.Print "Fehlender Parameter: " & ParName & " für Plattform " & PlatformKey
    MsgBox Replace(Replace(Get_Language_Str("Fehler: Der Parameter '#2#' ist für die Plattform '#1#' nicht definiert."), _
        "#1#", PlatformKey), "#2#", ParName), vbCritical, Get_Language_Str("Parameter Fehler")
End Sub

Sub Test_Get_Platform_String()
' Test values for parameter sheet
'   SPI_Pins    Pin list must have leading & trailing blanks    10 11 12    5 19 23
'   a       Test
'   TRUE        TRUE
'   TRue        TRue
'   0       0
'   1       1
'   false       false
'   intGood         -22
'   intBad          a44

    Get_Platform_String "ESP32", "SPI_Pins"
    Get_Platform_Bool "AM328", "TRUE"
    Get_Platform_Bool "AM328", "TRue"
    Get_Platform_Bool "AM328", "false"
    Get_Platform_Bool "AM328", "0"
    Get_Platform_Bool "AM328", "1"
    Get_Platform_Int "PICO", "intGood"
    Get_Platform_Int "PICO", "intBad"
End Sub


#If Win64 Then
   Public Function Get_Act_ms() As LongLong
      Get_Act_ms = GetTickCount64()
   End Function
#Else
   Public Function Get_Act_ms() As Long
      Get_Act_ms = getTickCount()
   End Function
#End If

#If False Then
'--------------------------------------------
Public Sub HourGlassCursor(bApply As Boolean)
'--------------------------------------------
    #If HostProject = "Access" Then
        Application.DoCmd.Hourglass bApply
    #ElseIf HostProject = "Word" Then
        System.Cursor = IIf(bApply, wdCursorWait, wdCursorNormal)
    #Else
        Application.Cursor = IIf(bApply, xlWait, xlDefault)
    #End If

    #If Mac = False Then
        Dim pt As POINTAPI
        If Not bApply Then
            ' in some systems the cursor may fail to reset to default, this forces it
            GetCursorPos pt
            SetCursorPos pt.X, pt.Y
        End If
    #End If
End Sub
#End If

Attribute VB_Name = "M40_ShellAndWait"
' M40_ShellAndWait:
' ~~~~~~~~~~~~~~~~~

' Module Description:
' ~~~~~~~~~~~~~~~~~~~
' This module provides a function to call external programs and wait for a certain time.
' In addition the windows style (Hidden, Maximized, Minimized, ...) and the
' "Ctrl Break" behavior could be defined.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Chip Pearson
' http://www.cpearson.com/excel/ShellAndWait.aspx
' This module contains code for the ShellAndWait function that will Shell to a process
' and wait for that process to End before returning to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Option Explicit
Option Compare Text


#If Win64 Then ' 28.09.19: New 64 Bit definition (Test für Armins MoBa Rechner)
  ' https://foren.activevb.de/forum/vba/thread-25588/beitrag-25588/VBA7-Win64-CreateProcess-WaitFo/
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As LongLong) As LongLong
#Else
  #If VBA7 Then
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As LongPtr ' LongPtr oder LongLong?
  #Else
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
  #End If
#End If


#If VBA7 Then
  Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
#Else
  Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
  Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
#End If

Private Const SYNCHRONIZE = &H100000

Public Enum ShellAndWaitResult
    Success = 0
    Failure = 1
    Timeout = 2
    InvalidParameter = 3
    SysWaitAbandoned = 4
    UserWaitAbandoned = 5
    UserBreak = 6
End Enum

Public Enum ActionOnBreak
    IgnoreBreak = 0
    AbandonWait = 1
    PromptUser = 2
End Enum

Public TaskId As Long

#If Win64 Then ' 28.09.19:
    Private Const STATUS_ABANDONED_WAIT_0 As LongLong = &H80
    Private Const STATUS_WAIT_0 As LongLong = &H0
    Private Const WAIT_ABANDONED As LongLong = (STATUS_ABANDONED_WAIT_0 + 0)
    Private Const WAIT_OBJECT_0 As LongLong = (STATUS_WAIT_0 + 0)
    Private Const WAIT_TIMEOUT As LongLong = 258&
    Private Const WAIT_FAILED As LongLong = &HFFFFFFFF
    Private Const WAIT_INFINITE = -1&
#Else
    Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
    Private Const STATUS_WAIT_0 As Long = &H0
    Private Const WAIT_ABANDONED As Long = (STATUS_ABANDONED_WAIT_0 + 0)
    Private Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)
    Private Const WAIT_TIMEOUT As Long = 258&
    Private Const WAIT_FAILED As Long = &HFFFFFFFF
    Private Const WAIT_INFINITE = -1&
#End If

Public Function ShellAndWait(ShellCommand As String, _
                    TimeOutSeconds As Double, _
                    ShellWindowState As VbAppWinStyle, _
                    BreakKey As ActionOnBreak) As ShellAndWaitResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShellAndWait
'
' This function calls Shell and passes to it the command text in ShellCommand. The function
' then waits for TimeOutSeconds (in Seconds) to expire.
'
'   Parameters:
'       ShellCommand
'           is the command text to pass to the Shell function.
'
'       TimeOutSeconds Hardi: Changed type to double (Old: Long)
'
'       TimeOutMs
'           is the number of milliseconds to wait for the shell'd program to wait. If the
'           shell'd program terminates before TimeOutMs has expired, the function returns
'           ShellAndWaitResult.Success = 0. If TimeOutMs expires before the shell'd program
'           terminates, the return value is ShellAndWaitResult.TimeOut = 2.
'
'       ShellWindowState
'           is an item in VbAppWinStyle specifying the window state for the shell'd program.
'
'       BreakKey
'           is an item in ActionOnBreak indicating how to handle the application's cancel key
'           (Ctrl Break). If BreakKey is ActionOnBreak.AbandonWait and the user cancels, the
'           wait is abandoned and the result is ShellAndWaitResult.UserWaitAbandoned = 5.
'           If BreakKey is ActionOnBreak.IgnoreBreak, the cancel key is ignored. If
'           BreakKey is ActionOnBreak.PromptUser, the user is given a ?Continue? message. If the
'           user selects "do not continue", the function returns ShellAndWaitResult.UserBreak = 6.
'           If the user selects "continue", the wait is continued.
'
'   Return values:
'            ShellAndWaitResult.Success = 0
'               indicates the the process completed successfully.
'            ShellAndWaitResult.Failure = 1
'               indicates that the Wait operation failed due to a Windows error.
'            ShellAndWaitResult.TimeOut = 2
'               indicates that the TimeOutMs interval timed out the Wait.
'            ShellAndWaitResult.InvalidParameter = 3
'               indicates that an invalid value was passed to the procedure.
'            ShellAndWaitResult.SysWaitAbandoned = 4
'               indicates that the system abandoned the wait.
'            ShellAndWaitResult.UserWaitAbandoned = 5
'               indicates that the user abandoned the wait via the cancel key (Ctrl+Break).
'               This happens only if BreakKey is set to ActionOnBreak.AbandonWait.
'            ShellAndWaitResult.UserBreak = 6
'               indicates that the user broke out of the wait after being prompted with
'               a ?Continue message. This happens only if BreakKey is set to
'               ActionOnBreak.PromptUser.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim TimeOutMs As Long

#If Win64 Then                                                              ' 28.09.19:
  Dim WaitRes As LongLong
#Else
  Dim WaitRes As Long
#End If

Dim ms As Long
Dim MsgRes As VbMsgBoxResult
Dim SaveCancelKey As XlEnableCancelKey
Dim ElapsedTime As Long
Dim Quit As Boolean
Dim ProcHandle As Long
Const ERR_BREAK_KEY = 18
Const DEFAULT_POLL_INTERVAL = 500

If Trim(ShellCommand) = vbNullString Then
    ShellAndWait = ShellAndWaitResult.InvalidParameter
    Exit Function
End If

TimeOutMs = 1000 * TimeOutSeconds
If TimeOutMs < 0 Then
    ShellAndWait = ShellAndWaitResult.InvalidParameter
    Exit Function
ElseIf TimeOutMs = 0 Then
    ms = WAIT_INFINITE
Else
    ms = TimeOutMs
End If

Select Case BreakKey
    Case AbandonWait, IgnoreBreak, PromptUser
        ' valid
    Case Else
        ShellAndWait = ShellAndWaitResult.InvalidParameter
        Exit Function
End Select

Select Case ShellWindowState
    Case vbHide, vbMaximizedFocus, vbMinimizedFocus, vbMinimizedNoFocus, vbNormalFocus, vbNormalNoFocus
        ' valid
    Case Else
        ShellAndWait = ShellAndWaitResult.InvalidParameter
        Exit Function
End Select

On Error Resume Next
Err.Clear
TaskId = Shell(ShellCommand, ShellWindowState)
If (Err.Number <> 0) Or (TaskId = 0) Then
    ShellAndWait = ShellAndWaitResult.Failure
    Exit Function
End If

ProcHandle = OpenProcess(SYNCHRONIZE, False, TaskId)
If ProcHandle = 0 Then
    ShellAndWait = ShellAndWaitResult.Failure
    Exit Function
End If

On Error GoTo errH:
SaveCancelKey = Application.EnableCancelKey
Application.EnableCancelKey = xlErrorHandler
WaitRes = WaitForSingleObject(ProcHandle, DEFAULT_POLL_INTERVAL)
Do Until WaitRes = WAIT_OBJECT_0
    DoEvents
    Select Case WaitRes
        Case WAIT_ABANDONED
            ' Windows abandoned the wait
            ShellAndWait = ShellAndWaitResult.SysWaitAbandoned
            Exit Do
        Case WAIT_OBJECT_0
            ' Successful completion
            ShellAndWait = ShellAndWaitResult.Success
            Exit Do
        Case WAIT_FAILED
            ' attach failed
            ShellAndWait = ShellAndWaitResult.Failure
            Exit Do
        Case WAIT_TIMEOUT
            ' Wait timed out. Here, this time out is on DEFAULT_POLL_INTERVAL.
            ' See if ElapsedTime is greater than the user specified wait
            ' time out. If we have exceed that, get out with a TimeOut status.
            ' Otherwise, reissue as wait and continue.
            ElapsedTime = ElapsedTime + DEFAULT_POLL_INTERVAL
            If ms > 0 Then
                ' user specified timeout
                If ElapsedTime > ms Then
                    ShellAndWait = ShellAndWaitResult.Timeout
                    Exit Do
                Else
                    ' user defined timeout has not expired.
                End If
            Else
                ' infinite wait -- do nothing
            End If
            ' reissue the Wait on ProcHandle
            WaitRes = WaitForSingleObject(ProcHandle, DEFAULT_POLL_INTERVAL)
            
        Case Else
            ' unknown result, assume failure
            ShellAndWait = ShellAndWaitResult.Failure
            Exit Do
            Quit = True
    End Select
Loop

CloseHandle ProcHandle
Application.EnableCancelKey = SaveCancelKey
Exit Function

errH:
Debug.Print "ErrH: Cancel: " & Application.EnableCancelKey
If Err.Number = ERR_BREAK_KEY Then
    If BreakKey = ActionOnBreak.AbandonWait Then
        CloseHandle ProcHandle
        ShellAndWait = ShellAndWaitResult.UserWaitAbandoned
        Application.EnableCancelKey = SaveCancelKey
        Exit Function
    ElseIf BreakKey = ActionOnBreak.IgnoreBreak Then
        Err.Clear
        Resume
    ElseIf BreakKey = ActionOnBreak.PromptUser Then
        MsgRes = MsgBoxMov("User Process Break." & vbCrLf & _
                           "Continue to wait?", vbYesNo)
        If MsgRes = vbNo Then
            CloseHandle ProcHandle
            ShellAndWait = ShellAndWaitResult.UserBreak
            Application.EnableCancelKey = SaveCancelKey
        Else
            Err.Clear
            Resume Next
        End If
    Else
        CloseHandle ProcHandle
        Application.EnableCancelKey = SaveCancelKey
        ShellAndWait = ShellAndWaitResult.Failure
    End If
Else
    ' some other error. assume failure
    CloseHandle ProcHandle
    ShellAndWait = ShellAndWaitResult.Failure
End If
Application.EnableCancelKey = SaveCancelKey

End Function





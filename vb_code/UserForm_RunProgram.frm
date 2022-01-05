VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_RunProgram 
   Caption         =   "..."
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12555
   OleObjectBlob   =   "UserForm_RunProgram.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "UserForm_RunProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' UserForm_RunProgram:
' ~~~~~~~~~~~~~~~~~
' Jürgen Winkler
' Description:
' ~~~~~~~~~~~~~~~~~~~
' This form provides a function to call external programs and wait for a certain time.
' It writes console output to a dialog
' with help of https://devblogs.microsoft.com/oldnewthing/20131209-00/?p=2433

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const DEFAULT_CANCEL_MESSAGE = "Vorgang abbrechen?"     ' default message, may be overriden

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_EX_COMPOSITED As Long = &H2000000

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20

Private Const Modal = 0
Private Const Modeless = 1
Private tSA_CreateProcessPrc As SECURITY_ATTRIBUTES
Private tSA_CreateProcessThrd As SECURITY_ATTRIBUTES
Private tSA_CreateProcessPrcInfo As PROCESS_INFORMATION
Private UserBreak As Boolean
Private BreakKey As ActionOnBreak
Private CancelMessage As String


#If Win64 Then ' 28.09.19: New 64 Bit definition (Test für Armins MoBa Rechner)
  ' https://foren.activevb.de/forum/vba/thread-25588/beitrag-25588/VBA7-Win64-CreateProcess-WaitFo/
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As LongLong) As LongLong
    Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As LongPtr, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As LongPtr, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As LongPtr
    Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Boolean
    Private Declare PtrSafe Function QueryInformationJobObject Lib "kernel32" (ByVal hJob As LongLong, ByVal JobObjectInfoClass As Long, lpJobObjectInfo As JOBOBJECT_EXTENDED_LIMIT_INFORMATION, ByVal cbJobObjectInfoLength As Long, ByRef cbLength As LongPtr) As Boolean
    Private Declare PtrSafe Function AssignProcessToJobObject Lib "kernel32" (ByVal hJob As LongLong, ByVal hProcess As LongLong) As Boolean
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function ResumeThread Lib "kernel32" (ByVal hThread As LongPtr) As Long
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
    Private Declare PtrSafe Function CreatePipe Lib "kernel32" (phReadPipe As LongPtr, phWritePipe As LongPtr, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
    Private Declare PtrSafe Function GetFileSize Lib "kernel32" (ByVal hFile As LongPtr, lpFileSizeHigh As Long) As Long
    Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
    Private Declare PtrSafe Function WriteFile Lib "kernel32.dll" (ByVal hFile As LongPtr, Buffer As String, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As LongPtr) As Long
    Private Declare PtrSafe Function EnableWindow Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal fEnable As Long) As Long
    Private Declare PtrSafe Function CreateJobObject Lib "kernel32" Alias "CreateJobObjectA" (ByVal lpJobAttributes As Long, ByVal lpName As String) As LongPtr
    Private Declare PtrSafe Function SetInformationJobObject Lib "kernel32" (ByVal hJob As LongPtr, ByVal JobObjectInfoClass As Long, ByRef lpJobObjectInfo As Any, ByVal cbJobObjectInfoLength As Long) As Boolean
#Else
  #If VBA7 Then
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As LongPtr ' LongPtr oder LongLong?
    Private Declare PtrSafe Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
    Private Declare PtrSafe Function QueryInformationJobObject Lib "kernel32" (ByVal hJob As Long, ByVal JobObjectInfoClass As Long, ByRef lpJobObjectInfo As Any, ByVal cbJobObjectInfoLength As Long, ByRef cbLength As Long) As Boolean
    Private Declare PtrSafe Function AssignProcessToJobObject Lib "kernel32" (ByVal hJob As Long, ByVal hProcess As Long) As Boolean
    Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Declare PtrSafe Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
    Private Declare PtrSafe Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
    Private Declare PtrSafe Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
    Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
    Private Declare PtrSafe Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef Buffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
    Private Declare PtrSafe Function EnableWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
    Private Declare PtrSafe Function CreateJobObject Lib "kernel32" Alias "CreateJobObjectA" (ByVal lpJobAttributes As Long, ByVal lpName As String) As Long
    Private Declare PtrSafe Function SetInformationJobObject Lib "kernel32" (ByVal hJob As Long, ByVal JobObjectInfoClass As Long, ByRef lpJobObjectInfo As Any, ByVal cbJobObjectInfoLength As Long) As Boolean
  #Else
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
    Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
    Private Declare Function QueryInformationJobObject Lib "kernel32" (ByVal hJob As Long, ByVal JobObjectInfoClass As Long, ByRef lpJobObjectInfo As Any, ByVal cbJobObjectInfoLength As Long, ByRef cbLength As Long) As Boolean
    Private Declare Function AssignProcessToJobObject Lib "kernel32" (ByVal hJob As Long, ByVal hProcess As Long) As Boolean
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
    Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
    Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
    Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
    Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
    Private Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef Buffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
    Private Declare Function EnableWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
    Private Declare Function CreateJobObject Lib "kernel32" Alias "CreateJobObjectA" (ByVal lpJobAttributes As Long, ByVal lpName As String) As Long
    Private Declare Function SetInformationJobObject Lib "kernel32" (ByVal hJob As Long, ByVal JobObjectInfoClass As Long, ByRef lpJobObjectInfo As Any, ByVal cbJobObjectInfoLength As Long) As Boolean
  #End If
#End If

#If VBA7 Then
  Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
  Private Declare PtrSafe Function GetLastError Lib "kernel32.dll" () As Long
  Private Declare PtrSafe Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
  Private Declare PtrSafe Sub LockWindowUpdate Lib "user32.dll" (ByVal hwnd As LongPtr)
  Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
  Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#Else
  Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
  Private Declare Function GetLastError Lib "kernel32.dll" () As Long
  Private Declare Sub GetStartupInfo Lib "kernel32" Alias "GetStartupInfoA" (lpStartupInfo As STARTUPINFO)
  Private Declare Sub LockWindowUpdate Lib "user32.dll" (ByVal hwnd As Long)
  Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Private Type LARGE_INTEGER
    lowPart As Long
    highPart As Long
End Type

Private Type TIOCounters
    ReadOperationCount As LARGE_INTEGER
    WriteOperationCount As LARGE_INTEGER
    OtherOperationCount As LARGE_INTEGER
    ReadTransferCount As LARGE_INTEGER
    WriteTransferCount As LARGE_INTEGER
    OtherTransferCount As LARGE_INTEGER
End Type

#If Win64 Then ' 28.09.19:
    Private Const STATUS_ABANDONED_WAIT_0 As LongLong = &H80
    Private Const STATUS_WAIT_0 As LongLong = &H0
    Private Const WAIT_ABANDONED As LongLong = (STATUS_ABANDONED_WAIT_0 + 0)
    Private Const WAIT_OBJECT_0 As LongLong = (STATUS_WAIT_0 + 0)
    Private Const WAIT_TIMEOUT As LongLong = 258&
    Private Const WAIT_FAILED As LongLong = &HFFFFFFFF
    Private Const WAIT_INFINITE = -1&
    
    Private hWriteOut As LongPtr
    Private hJob As LongPtr
    
    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As LongPtr
        bInheritHandle As Long
    End Type
    
    Private Type PROCESS_INFORMATION
        hProcess As LongPtr
        hThread As LongPtr
        dwProcessId As Long
        dwThreadId As Long
    End Type
    
    Private Type STARTUPINFO
        cb As Long
        lpReserved As LongPtr
        lpDesktop As LongPtr
        lpTitle As LongPtr
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As LongPtr
        hStdInput As LongPtr
        hStdOutput As LongPtr
        hStdError As LongPtr
    End Type
    
    Private Type JOBOBJECT_BASIC_LIMIT_INFORMATION
        PerProcessUserTimeLimit As LARGE_INTEGER
        PerJobUserTimeLimit As LARGE_INTEGER
        LimitFlags As Long
        MinimumWorkingSetSize As LongLong
        MaximumWorkingSetSize As LongLong
        ActiveProcessLimit As Long
        ByteArray(23) As Byte
    End Type
    
    Private Type JOBOBJECT_EXTENDED_LIMIT_INFORMATION
      BasicLimitInformation As JOBOBJECT_BASIC_LIMIT_INFORMATION
      IoInfo As TIOCounters
      ProcessMemoryLimit As LongLong
      JobMemoryLimit As LongLong
        PeakProcessMemoryUsed As LongLong
        PeakJobMemoryUsed As LongLong
    End Type
    
#Else
    Private Const STATUS_ABANDONED_WAIT_0 As Long = &H80
    Private Const STATUS_WAIT_0 As Long = &H0
    Private Const WAIT_ABANDONED As Long = (STATUS_ABANDONED_WAIT_0 + 0)
    Private Const WAIT_OBJECT_0 As Long = (STATUS_WAIT_0 + 0)
    Private Const WAIT_TIMEOUT As Long = 258&
    Private Const WAIT_FAILED As Long = &HFFFFFFFF
    Private Const WAIT_INFINITE = -1&
    
    Private hWriteOut As Long
    Private hJob As Long
    
    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
    End Type
    
    Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
    End Type
     
    Private Type STARTUPINFO
        cb As Long
        lpReserved As Long
        lpDesktop As Long
        lpTitle As Long
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
    End Type
    
    Private Type JOBOBJECT_BASIC_LIMIT_INFORMATION
        PerProcessUserTimeLimit As LARGE_INTEGER
        PerJobUserTimeLimit As LARGE_INTEGER
        LimitFlags As Long
        MinimumWorkingSetSize As Long
        MaximumWorkingSetSize As Long
        ActiveProcessLimit As Long
        ByteArray(15) As Byte
    End Type
    
    Private Type JOBOBJECT_EXTENDED_LIMIT_INFORMATION
      BasicLimitInformation As JOBOBJECT_BASIC_LIMIT_INFORMATION
      IoInfo As TIOCounters
      ProcessMemoryLimit As Long
      JobMemoryLimit As Long
        PeakProcessMemoryUsed As Long
        PeakJobMemoryUsed As Long
    End Type
#End If

Private Const JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = &H2000

Private Const JobObjectInfoType_AssociateCompletionPortInformation = 7
Private Const JobObjectInfoType_BasicLimitInformation = 2
Private Const JobObjectInfoType_BasicUIRestrictions = 4
Private Const JobObjectInfoType_EndOfJobTimeInformation = 6
Private Const JobObjectInfoType_ExtendedLimitInformation = 9
Private Const JobObjectInfoType_SecurityLimitInformation = 5
Private Const JobObjectInfoType_GroupInformation = 11

Private Const STARTF_USESHOWWINDOW  As Long = &H1
Private Const STARTF_USESTDHANDLES  As Long = &H100
Private Const SW_HIDE               As Long = 0&
Private Const SW_SHOW               As Long = 5&
Private Const CREATE_SUSPENDED = &H4
Private Const CREATE_NEW_PROCESS_GROUP = &H200
Private Const POLL_INTERVAL = 50

#If Win64 Then
    Private mlnghWnd As LongPtr
    Public Property Get hwnd() As LongPtr
        hwnd = mlnghWnd
    End Property
#Else
    Private mlnghWnd As Long
    Public Property Get hwnd() As Long
        hwnd = mlnghWnd
    End Property
#End If

Private Sub ShowLastError()
    Debug.Print ("lastError=" & GetLastError())
End Sub


Private Sub StorehWnd()
 
    Dim strCaption As String
    Dim strClass As String
 
    'class name changed in Office 2000
    If val(Application.Version) >= 9 Then
        strClass = "ThunderDFrame"
    Else
        strClass = "ThunderXFrame"
    End If
 
    'remember the caption so we can
    'restore it when we're done
    strCaption = Me.Caption
 
    'give the userform a random
    'unique caption so we can reliably
    'get a handle to its window
    Randomize
    Me.Caption = CStr(RND)
 
    'store the handle so we can use
    'it for the userform's lifetime
    mlnghWnd = FindWindow(strClass, Me.Caption)
 
    'set the caption back again
    Me.Caption = strCaption
 
End Sub

Private Sub UserForm_Initialize()
   Me.Top = Application.Top + (Application.UsableHeight / 2) - (Me.Height / 2)
   Me.Left = Application.Left + (Application.UsableWidth / 2) - (Me.Width / 2)
   hJob = 0
   StorehWnd
   BreakKey = AbandonWait
End Sub

Private Sub UserForm_Activate()
    FormResizable
    UserForm_Resize
End Sub

Public Sub FormResizable()

Dim lStyle As Long
lStyle = GetWindowLong(hwnd, GWL_STYLE) Or WS_THICKFRAME
SetWindowLong hwnd, GWL_STYLE, lStyle

lStyle = GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_COMPOSITED           ' to avoid flickering
SetWindowLong hwnd, GWL_EXSTYLE, lStyle

End Sub

Private Sub console_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Debug.Print "ConsoleKD=" & KeyCode & " Shift=" & Shift
    Dim written
    If KeyCode = 67 And Shift = 2 Then
        If PromptUserToAbort <> vbYes Then Exit Sub
        Debug.Print ("crtl-c")
        
        userform_terminate
        UserBreak = True
    End If
    If (Shift = 0 And KeyCode > 0 And KeyCode < 31) Then WriteFile hWriteOut, Chr(KeyCode), 1, written, 0
    
End Sub
Private Function PromptUserToAbort()
    PromptUserToAbort = vbNo
    Select Case BreakKey
       Case ActionOnBreak.PromptUser
            PromptUserToAbort = MsgBox(CancelMessage, vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption)
            Exit Function
       Case ActionOnBreak.AbandonWait
            PromptUserToAbort = vbYes
            Exit Function
    End Select
End Function
Private Sub console_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim written As Long
    If (KeyAscii > 0) Then WriteFile hWriteOut, Chr(KeyAscii), 1, written, 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        If PromptUserToAbort <> vbYes Then
            Cancel = 1
        Else
            UserBreak = True
        End If
    End If
End Sub

Private Sub UserForm_Resize()
    If Me.InsideWidth > 8 Then console.Width = Me.InsideWidth - 8           ' 19.01.21: Jürgen: Added: If
    If Me.InsideHeight > 24 Then                                            '     "
        console.Height = Me.InsideHeight - 24
        LabelTime.Top = Me.InsideHeight - 18
    End If
End Sub

Private Sub userform_terminate()
    EnableWindow Application.hwnd, Modeless
    If tSA_CreateProcessPrcInfo.hProcess <> 0 Then
        TerminateProcess tSA_CreateProcessPrcInfo.hProcess, -1
    End If
End Sub

Public Function ShellExecute(ShellCommand As String, _
                    TimeOutSeconds As Double, _
                    WindowTitle As String, _
                    Optional BreakAction As ActionOnBreak = ActionOnBreak.PromptUser, _
                    Optional BackgroundColor As Long = vbBlack, _
                    Optional TextColor As Long = vbWhite, _
                    Optional userCancelMessage As String = "" _
                    ) As ShellAndWaitResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Execute
'
' This function calls Shell and passes to it the command text in ShellCommand. The function
' then waits for TimeOutSeconds (in Seconds) to expire. If captures outputs of shell and displays it in a text box
'
'   Parameters:
'       ShellCommand
'           is the command text to pass to the Shell function.
'
'       TimeOutSeconds Hardi: Changed type to double (Old: Long)
'
'       WindowTitle: The caption of the window
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
  Dim lngResult As LongLong
  Dim hRead As LongLong
  Dim hWrite As LongLong
  Dim hReadOut As LongLong
#Else
  Dim WaitRes As Long
  Dim lngResult As Long
  Dim hRead As Long
  Dim hWrite As Long
  Dim hReadOut As Long
#End If

Dim ms As Long
Dim MsgRes As VbMsgBoxResult
Dim SaveCancelKey As XlEnableCancelKey
Dim ElapsedTime As Long
Dim tSA_CreatePipe As SECURITY_ATTRIBUTES
Dim tStartupInfo As STARTUPINFO
Dim lngSizeOf As Long
Dim bRead As Long
Dim abytBuff() As Byte
Dim lngExitCode As Long

Const ERR_BREAK_KEY = 18

If Trim(ShellCommand) = vbNullString Then
    ShellExecute = ShellAndWaitResult.InvalidParameter
    Exit Function
End If

If Trim(WindowTitle) = vbNullString Then
    ShellExecute = ShellAndWaitResult.InvalidParameter
    Exit Function
End If

TimeOutMs = 1000 * TimeOutSeconds
If TimeOutMs < 0 Then
    ShellExecute = ShellAndWaitResult.InvalidParameter
    Exit Function
ElseIf TimeOutMs = 0 Then
    ms = WAIT_INFINITE
Else
    ms = TimeOutMs
End If

BreakKey = BreakAction
Select Case BreakKey
    Case AbandonWait, IgnoreBreak, PromptUser
        ' valid
    Case Else
        ShellExecute = ShellAndWaitResult.InvalidParameter
        Exit Function
End Select

If userCancelMessage <> "" Then
    CancelMessage = userCancelMessage
Else
    CancelMessage = DEFAULT_CANCEL_MESSAGE
End If

tSA_CreatePipe.nLength = Len(tSA_CreatePipe)
tSA_CreatePipe.lpSecurityDescriptor = 0&
tSA_CreatePipe.bInheritHandle = True
UserBreak = False
console.ForeColor = TextColor
console.BackColor = BackgroundColor

Me.Show vbModeless                                                          ' 19.01.21: Added vbModeless
EnableWindow Application.hwnd, Modal
Application.EnableCancelKey = xlDisabled
Application.ScreenUpdating = False

If CreatePipe(hRead, hWrite, tSA_CreatePipe, 0&) <> 0& And CreatePipe(hReadOut, hWriteOut, tSA_CreatePipe, 0&) <> 0 Then
    tStartupInfo.cb = Len(tStartupInfo)
    GetStartupInfo tStartupInfo
 
    With tStartupInfo
        .hStdOutput = hWrite
        .hStdError = hWrite
        .hStdInput = hReadOut
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
    End With
    On Error Resume Next
    Err.Clear
    
    Rem create a list of null terminated environment variables
    Dim EnvString, Indx
    Dim ProcEnv As String
    Indx = 1
    Do
        EnvString = Environ(Indx)    ' Get environment
        ProcEnv = ProcEnv + EnvString + Chr(0)
        Indx = Indx + 1
    Loop Until EnvString = ""
    ProcEnv = ProcEnv + Chr(0)
          
    Dim Buffer() As Byte
    Buffer = StrConv(ProcEnv, vbFromUnicode)    ' Convert string to singleByte
   
    Dim basePath As String
    basePath = Left(ShellCommand, InStrRev(ShellCommand, Application.PathSeparator))  ' get the base directory
    If Left(basePath, 1) = """" Then basePath = Mid(basePath, 2)
    
    lngResult = CreateProcess(0, ShellCommand, tSA_CreateProcessPrc, tSA_CreateProcessThrd, True, _
    CREATE_SUSPENDED + CREATE_NEW_PROCESS_GROUP + NORMAL_PRIORITY_CLASS, _
    VarPtr(Buffer(0)), basePath, tStartupInfo, tSA_CreateProcessPrcInfo)
    
    If (lngResult <> 0&) Then
        Me.Caption = WindowTitle
        Dim startupOk As Boolean
        startupOk = False

        Dim info As JOBOBJECT_EXTENDED_LIMIT_INFORMATION
        hJob = CreateJobObject(0, 0)
        Dim cbLen
        If QueryInformationJobObject(hJob, JobObjectInfoType_ExtendedLimitInformation, info, Len(info), cbLen) <> 0 Then
            Debug.Print "ActiveProcessLimit=" & CStr(info.BasicLimitInformation.ActiveProcessLimit)
            info.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
            If SetInformationJobObject(hJob, JobObjectInfoType_ExtendedLimitInformation, info, Len(info)) Then
                If AssignProcessToJobObject(hJob, tSA_CreateProcessPrcInfo.hProcess) Then
                    startupOk = ResumeThread(tSA_CreateProcessPrcInfo.hThread)
                End If
            End If
        End If
        If startupOk Then
            Do
                WaitRes = WaitForSingleObject(tSA_CreateProcessPrcInfo.hProcess, POLL_INTERVAL)
                lngSizeOf = GetFileSize(hRead, 0&)
                If (lngSizeOf > 0) Then
                    ReDim abytBuff(lngSizeOf - 1)
                    If ReadFile(hRead, abytBuff(0), UBound(abytBuff) + 1, bRead, 0) Then
                        Dim Res As String
                        Res = StrConv(abytBuff, vbUnicode)
                        Me.AddToConsole (Res)
                    End If
                End If
                DoEvents
                ElapsedTime = ElapsedTime + POLL_INTERVAL
                Me.LabelTime = Format(ElapsedTime / 86400 / 1000, "hh:mm:ss")
                If ms > 0 Then
                    ' user specified timeout
                    If ElapsedTime > ms Then
                        ShellExecute = ShellAndWaitResult.Timeout
                        Exit Do
                    Else
                        ' user defined timeout has not expired.
                    End If
                Else
                    ' infinite wait -- do nothing
                End If
            Loop While WaitRes = WAIT_TIMEOUT
            tSA_CreateProcessPrcInfo.hProcess = 0
            Call GetExitCodeProcess(tSA_CreateProcessPrcInfo.hProcess, lngExitCode)
            CloseHandle tSA_CreateProcessPrcInfo.hThread
            CloseHandle tSA_CreateProcessPrcInfo.hProcess
            tSA_CreateProcessPrcInfo.hProcess = 0
            CloseHandle hJob
            hJob = 0
            
            If UserBreak Then
                ShellExecute = ShellAndWaitResult.UserBreak
            ElseIf (lngExitCode <> 0&) Then
                ShellExecute = ShellAndWaitResult.Failure
            End If
        Else
            ShellExecute = ShellAndWaitResult.Failure
        End If
    Else
        ShellExecute = ShellAndWaitResult.Failure
    End If
    CloseHandle hWriteOut
    CloseHandle hReadOut
    CloseHandle hWrite
    CloseHandle hRead
Else
    ShellExecute = ShellAndWaitResult.Failure
End If
Application.EnableCancelKey = SaveCancelKey
Unload UserForm_RunProgram

End Function

Public Sub AddToConsole(message As String)
'----------------------------------------------------
  message = Replace(message, vbCrLf, vbCr)
  message = Replace(message, vbLf, vbCr)
  message = ConvertUTF8Str(message)
  LockWindowUpdate hwnd
  With console
      .setFocus '//required
      If Left(message, 1) = Chr$(13) And Len(message) > 1 Then ' start with newline only
          If Left(message, 2) <> Chr$(10) And InStrRev(.Text, Chr$(13) + Chr$(10)) > 0 Then
            .Text = Left(.Text, InStrRev(.Text, Chr$(13))) + Mid(message, 2)
          Else
              .Text = .Text + message
          End If
         
      Else
          .Text = .Text + message
      End If
      .SelStart = Len(.Text)
      .HideSelection = False
  End With
  LockWindowUpdate 0
End Sub





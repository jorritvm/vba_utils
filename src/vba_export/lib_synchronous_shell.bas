Attribute VB_Name = "lib_synchronous_shell"
Option Explicit
Option Compare Text

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modShellAndWait
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
' This page on the web site: www.cpearson.com/Excel/ShellAndWait.aspx
' 9-September-2008
'
' This module contains code for the ShellAndWait function that will Shell to a process
' and wait for that process to end before returning to the caller.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" ( _
    ByVal hHandle As LongPtr, _
    ByVal dwMilliseconds As LongPtr) As LongPtr

Private Declare PtrSafe Function OpenProcess Lib "kernel32.dll" ( _
    ByVal dwDesiredAccess As LongPtr, _
    ByVal bInheritHandle As LongPtr, _
    ByVal dwProcessId As LongPtr) As LongPtr

Private Declare PtrSafe Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As LongPtr) As LongPtr
    
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

Private Const SYNCHRONIZE = &H100000

Public Enum ShellAndWaitResult
    Success = 0
    Failure = 1
    TimeOut = 2
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

Private Const STATUS_ABANDONED_WAIT_0 As LongPtr = &H80
Private Const STATUS_WAIT_0 As LongPtr = &H0
Private Const WAIT_ABANDONED As LongPtr = (STATUS_ABANDONED_WAIT_0 + 0)
Private Const WAIT_OBJECT_0 As LongPtr = (STATUS_WAIT_0 + 0)
Private Const WAIT_TIMEOUT As LongPtr = 258&
Private Const WAIT_FAILED As LongPtr = &HFFFFFFFF
Private Const WAIT_INFINITE = -1&


Public Function ShellAndWait(ShellCommand As String, _
       TimeOutMs As LongPtr, _
       ShellWindowState As VbAppWinStyle, _
       BreakKey As ActionOnBreak) As ShellAndWaitResult
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ShellAndWait
'
' This function calls Shell and passes to it the command text in ShellCommand. The function
' then waits for TimeOutMs (in milliseconds) to expire.
'
'   Parameters:
'       ShellCommand
'           is the command text to pass to the Shell function.
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
    
    Dim TaskID      As LongPtr
    Dim ProcHandle  As LongPtr
    Dim WaitRes     As LongPtr
    Dim ms          As LongPtr
    Dim MsgRes      As VbMsgBoxResult
    Dim SaveCancelKey As XlEnableCancelKey
    Dim ElapsedTime As LongPtr
    Dim Quit        As Boolean
    Const ERR_BREAK_KEY = 18
    Const DEFAULT_POLL_INTERVAL = 500
    
    If Trim(ShellCommand) = vbNullString Then
        ShellAndWait = ShellAndWaitResult.InvalidParameter
        Exit Function
    End If
    
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
    TaskID = shell(ShellCommand, ShellWindowState)
    If (Err.Number <> 0) Or (TaskID = 0) Then
        ShellAndWait = ShellAndWaitResult.Failure
        Exit Function
    End If
    
    ProcHandle = OpenProcess(SYNCHRONIZE, False, TaskID)
    If ProcHandle = 0 Then
        ShellAndWait = ShellAndWaitResult.Failure
        Exit Function
    End If
    
    On Error GoTo ErrH:
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
                ShellAndWait = ShellAndWaitResult.Success
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
                        ShellAndWait = ShellAndWaitResult.TimeOut
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
                Quit = True
        End Select
    Loop
    
    CloseHandle ProcHandle
    Application.EnableCancelKey = SaveCancelKey
    Exit Function
    
ErrH:
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
            MsgRes = MsgBox("User Process Break." & vbCrLf & _
                     "Continue To wait?", vbYesNo)
            If MsgRes = vbNo Then
                CloseHandle ProcHandle
                ShellAndWait = ShellAndWaitResult.UserBreak
                Application.EnableCancelKey = SaveCancelKey
            Else
                Err.Clear
                Resume Next
            End If
        Else
            'Debug.Print "Unknown value of 'BreakKey': " & CStr(BreakKey)
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


'***************************************************************************
'Purpose: run an external program in a synchronous shell
'         vba invokes a cmd prompt which runs a powershell which in turn run the desired app
'         the apps stout and stderr output can be piped to a logfile
'Inputs:  executable
'         dbug: boolean indicating whether the script should halt before closing the powershell
'         logfile: string containing location of logfile, empty string means no log
'         success_string: string that has to appear in the last 5 lines of the logfile to check if the script terminated successfully
'         args: list of parameters to be passed on to the executable, these will be space-separated
'Outputs: 0: success
'         1: fail
'         2: no success_string passed
'         3: synchronous shell threw error
'         4: executable not found
'***************************************************************************
Function run_app_in_synchronous_shell(ByVal executable As String, _
                                        ByVal dbug As Boolean, _
                                        ByVal logfile As String, _
                                        ByVal success_string As String, _
                                        ParamArray args() As Variant) As ShellAndWaitResult

    Dim batchCommand As String, cmd_switch As String
    Dim arg As Variant
    Dim return_value As Integer
    
    return_value = 0
    If (Not file_exists(executable)) Then return_value = 4
    
    cmd_switch = "/C"
    If (dbug) Then cmd_switch = "/K"
       
    If (return_value < 4) Then
        'init the command
        batch_command = "cmd " & cmd_switch & " " & _
                        Chr(34) & "powershell.exe -command " & _
                        Chr(34) & "& " & _
                        Chr(39) & executable & Chr(39)
        'add all arguments
        For Each arg In args
            batch_command = batch_command & _
                           " " & Chr(39) & arg & Chr(39)
        Next arg
        'pipe command output
        If (logfile <> "") Then
            batch_command = batch_command & _
                           " 2>&1 | tee " & _
                           Chr(39) & logfile & Chr(39)
        End If
        'close and show full command
        batch_command = batch_command & Chr(34) & Chr(34)
        Debug.Print ("batch_command: " & batch_command)
        
        'run the external app and get the shell result
        shell_result = ShellAndWait(batch_command, 0, vbMaximizedFocus, AbandonWait)
        If shell_result > 0 Then
            return_value = 3
        End If
        
        'verify the logfile to get an even better shell result
        If (shell_result = 0 And success_string <> "") Then
            Dim result As Boolean
            result = look_for_string_in_last_n_lines(logfile, success_string)
            
            If (result) Then
                return_value = 0
            Else
                return_value = 1
            End If
        End If
    End If
    
    run_app_in_synchronous_shell = return_value
End Function


Function look_for_string_in_last_line_of_file(ByVal text_file As String, ByVal string_to_look_for As String) As booolean
'***************************************************************************
'Purpose: returns true/false depending whether a given string is found at the end of the text file
'Inputs text_file: string with text file path
'       string_to_look_for
'Outputs: boolean
'***************************************************************************
    Dim return_value As Boolean
    Dim text_line As String
    
    return_value = False
    
    text_line = get_last_line_of_textfile(text_file)
    If InStr(1, text_line, string_to_look_for, vbTextCompare) > 0 Then return_value = True
    
    look_for_string_in_last_line_of_file = return_value
End Function


Function get_last_line_of_textfile(text_file As String) As String
'***************************************************************************
'Purpose: returns the last line of text in a text file
'Inputs text_file: string with text file path
'Outputs: string
'***************************************************************************
    Const ForReading = 1
    Const TristateTrue = 0        '-1: UNICODE / 0: ASCII
    Const IOMode = 0
    
    Dim FSO          As Object
    Dim txt_fs       As Object
    Dim temp_string       As String
    Dim result_string As String
    
    ' Open the text file as a textstream
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set txt_fs = FSO.OpenTextFile(text_file, ForReading, IOMode, TristateTrue)
    
    ' read all lines untill we end up with the last non empty one
    Do Until txt_fs.AtEndOfStream
        temp_string = txt_fs.ReadLine
        If Len(temp_string) > 0 Then
            result_string = temp_string
        End If
    Loop
    
    get_last_line_of_textfile = result_string
End Function

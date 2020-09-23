Attribute VB_Name = "modGeneral"
'***************************************************************************************
'*                                VB General Functions                                 *
'*                               Last Updated 26/03/2002                               *
'*                                                                                     *
'*  ExecCmd      -  Executes an external program and waits for it to close             *
'*  Wait         -  Waits a specified time without hanging the OS                      *
'*  DirExists    -  Checks for the existance of a directory                            *
'*  LoadResAVI   -  Loads and runs an AVI into the Animation Control                   *
'*  LPad         -  Pads a string with a  specified character to the right             *
'*  RPad         -  Pads a string with a  specified character to the left              *
'*                                                                                     *
'*       If you have any problems or queries: martin.sidgreaves@wrigley.co.uk          *
'***************************************************************************************
Option Explicit

'Declare API functions
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

'Define Constants
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const WM_USER = &H400
Private Const INFINITE = -1&
Public Const ACM_OPEN = WM_USER + 100&


Public Enum AnimationType
    Busy = 101
End Enum

'Define Types
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
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
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type


Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type


Public fnum As Integer
Public pings As Long
Public delay As Long
Public lResID As Long


Sub Main()

  Dim check As String

    'Check to see if registry values exist
    If GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger", "Installed") = "Error" Then
        
        'They don't so install with default settings
        CreateKey "HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger"
        SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger", "Installed", "True"
        SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger", "Pings", "10"
        SetBinaryValue "HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger", "Delay", "5"
        pings = 10
        delay = 5
        frmMain.Show
        
      Else
        'Read the values from the registry
        pings = GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger", "Pings")
        delay = GetBinaryValue("HKEY_LOCAL_MACHINE\SOFTWARE\ICMP Pinger", "Delay")
        frmMain.Show
        
    End If

End Sub


'=========================================================================
  'Function ExecCmd(cmdline$)
  'Executes an external program and waits for it to close
  'before continuing
  
  'Inputs:    cmdline  -  The command line to be executed
  
  'Returns:   None
'=========================================================================
Public Function ExecCmd(cmdline As String)

  Dim proc As PROCESS_INFORMATION
  Dim start As STARTUPINFO
  Dim ret As Long

    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)

    ' Start the shelled application:
    ret = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)

    ' Wait for the shelled application to finish:
    ret = WaitForSingleObject(proc.hProcess, INFINITE)
    Call GetExitCodeProcess(proc.hProcess, ret&)
    Call CloseHandle(proc.hThread)
    Call CloseHandle(proc.hProcess)
    ExecCmd = ret

End Function


'=========================================================================
  'Function Wait(ByVal TimeToWait As Long)
  'Waits a specified time without hanging the OS
  
  'Inputs:    TimeToWait - Pause in seconds
  
  'Returns:   None
'=========================================================================
Public Function Wait(ByVal TimeToWait As Long)
    Dim EndTime As Long
    
    '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
    EndTime = GetTickCount + TimeToWait * 1000
    
    Do Until GetTickCount > EndTime
        DoEvents
    Loop
End Function


'=========================================================================
'Function DirExists(DirName As String) As Boolean
'Checks for the existance of a DIRECTORY
  
'Inputs:    Dirname  -  The Directory path being checked
  
'Returns:   DirExists as boolean
'           Contains TRUE or FALSE depending on the result
'=========================================================================
Public Function DirExists(DirName As String) As Boolean

    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False

End Function


'=========================================================================
  'Sub LoadResAVI(aniControl As Animation, resAniID As AnimationType)
  'Loads an animation control with an AVI
  
  'Inputs:    aniControl - The name of the Amination control being used
  '           resAniID  - The Resource ID in the RES file
  
  'Returns:   None
'=========================================================================
Public Sub LoadResAVI(aniControl As Animation, resAniID As AnimationType)
    SendMessage aniControl.hwnd, ACM_OPEN, ByVal App.hInstance, ByVal resAniID
End Sub


'=========================================================================
  'Function Lpad(strToPad As String, strPadChar As String, intLength As Integer)
  'Pads a string with a specified character from the left
  
  'Inputs:    strToPad         -  The string to be padded out
  '           strPadChar  -  The character to pad with
  '           intLength  -  The final length of the string
  
  'Returns:   Lpad            -  The padded string
'=========================================================================
Function Lpad(strToPad As String, strPadChar As String, intLength As Integer) As String

  Dim padlength As Long
  Dim intCount As Integer
  Dim PadString As String
     
    padlength = (intLength - Len(strToPad))
    For intCount = 1 To padlength
        PadString = PadString & strPadChar
    Next intCount
    Lpad = PadString + strToPad

End Function


'=========================================================================
  'Function Rpad(strToPad As String, strPadChar As String, intLength As Integer)
  'Pads a string with a specified character to the right
  
  'Inputs:    strToPad         -  The string to be padded out
  '           strPadChar  -  The character to pad with
  '           intLength  -  The final length of the string
  
  'Returns:   Lpad            -  The padded string
'=========================================================================
Function Rpad(strToPad As String, strPadChar As String, intLength As Integer) As String

  Dim padlength As Long
  Dim intCount As Integer
  Dim PadString As String
     
    padlength = (intLength - Len(strToPad))
  
    For intCount = 1 To padlength
        PadString = strPadChar & PadString
    Next intCount
    Rpad = strToPad + PadString

End Function

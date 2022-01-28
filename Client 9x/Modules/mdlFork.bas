Attribute VB_Name = "mdlFork"
Option Explicit
Option Private Module

Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

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

Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
   (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Public Function ExecCmd(cmdline$)
   Dim proc As PROCESS_INFORMATION
   Dim start As STARTUPINFO
   
   
   Dim ret As Variant
   
   ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    
    With start
        .lpReserved = vbNull
        .lpReserved2 = 0
        .cbReserved2 = 0
        .lpDesktop = vbNull
        .lpTitle = vbNull
        .dwFlags = 0
        .dwX = 0
        .dwY = 0
        .dwFillAttribute = 0
    End With

   ' Start the shelled application:
'   ret = CreateProcessA(vbNullString, _
                        cmdline$, _
                        0&, _
                        0&, _
                        1&, _
                        NORMAL_PRIORITY_CLASS, _
                        0&, _
                        vbNullString, _
                        start, _
                        proc)

   ret = CreateProcessA(cmdline$, _
                        vbNull, _
                        vbNull&, _
                        vbNull&, _
                        0, _
                        NORMAL_PRIORITY_CLASS, _
                        vbNull, _
                        vbNull, _
                        start, _
                        proc)

    frmTest.Caption = ret
    
   ' Wait for the shelled application to finish:
'      ret = WaitForSingleObject(proc.hProcess, INFINITE)
'      Call GetExitCodeProcess(proc.hProcess, ret)
'      Call CloseHandle(proc.hThread)
'      Call CloseHandle(proc.hProcess)
      ExecCmd = ret
End Function

Attribute VB_Name = "mdlSystemOperations"
Option Explicit
Option Private Module

Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
Private Const TOKEN_QUERY As Long = &H8
Private Const SE_PRIVILEGE_ENABLED As Long = &H2

Private Const EWX_LOGOFF As Long = &H0
Private Const EWX_SHUTDOWN As Long = &H1
Private Const EWX_REBOOT As Long = &H2
Private Const EWX_FORCE As Long = &H4
Private Const EWX_POWEROFF As Long = &H8
Private Const EWX_FORCEIFHUNG As Long = &H10 '2000/XP only

Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type

Private Type LUID
   dwLowPart As Long
   dwHighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
   udtLUID As LUID
   dwAttributes As Long
End Type

Private Type TOKEN_PRIVILEGES
   PrivilegeCount As Long
   laa As LUID_AND_ATTRIBUTES
End Type
      
Private Declare Function ExitWindowsEx Lib "USER32" _
   (ByVal dwOptions As Long, _
   ByVal dwReserved As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenProcessToken Lib "advapi32" _
  (ByVal ProcessHandle As Long, _
   ByVal DesiredAccess As Long, _
   TokenHandle As Long) As Long

Private Declare Function LookupPrivilegeValue Lib "advapi32" _
   Alias "LookupPrivilegeValueA" _
  (ByVal lpSystemName As String, _
   ByVal lpName As String, _
   lpLuid As LUID) As Long

Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
  (ByVal TokenHandle As Long, _
   ByVal DisableAllPrivileges As Long, _
   NewState As TOKEN_PRIVILEGES, _
   ByVal BufferLength As Long, _
   PreviousState As Any, _
   ReturnLength As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" _
   Alias "GetVersionExA" _
  (lpVersionInformation As OSVERSIONINFO) As Long

Public Sub ShutDown()
   Dim uflags As Long
   Dim success As Long
   
   uflags = EWX_SHUTDOWN
   
   uflags = uflags Or EWX_FORCE
   
  'assume success
   success = True
  
  'if running under NT or better,
  'the shutdown privledges need to
  'be adjusted to allow the ExitWindowsEx
  'call. If the adjust call fails on a NT+
  'system, success holds False, preventing shutdown.
   If IsWinNTPlus Then
        uflags = EWX_POWEROFF
        success = EnableShutdownPrivledges()
   End If

   If success Then Call ExitWindowsEx(uflags, 0&)
   
End Sub

Public Sub ReBoot()
   Dim uflags As Long
   Dim success As Long
   
   uflags = EWX_REBOOT
   
   uflags = uflags Or EWX_FORCE
   
  'assume success
   success = True
  
  'if running under NT or better,
  'the shutdown privledges need to
  'be adjusted to allow the ExitWindowsEx
  'call. If the adjust call fails on a NT+
  'system, success holds False, preventing shutdown.
   If IsWinNTPlus Then
      success = EnableShutdownPrivledges()
   End If

   If success Then Call ExitWindowsEx(uflags, 0&)
   
End Sub

Public Sub LogOff()
   Dim uflags As Long
   Dim success As Long
   
   uflags = EWX_LOGOFF
   
   uflags = uflags Or EWX_FORCE
   
  'assume success
   success = True
  
  'if running under NT or better,
  'the shutdown privledges need to
  'be adjusted to allow the ExitWindowsEx
  'call. If the adjust call fails on a NT+
  'system, success holds False, preventing shutdown.
   If IsWinNTPlus Then
      success = EnableShutdownPrivledges()
   End If

   If success Then Call ExitWindowsEx(uflags, 0&)
   
End Sub

Private Function IsWinNTPlus() As Boolean

  'returns True if running Windows NT,
  'Windows 2000, Windows XP, or .net server
   #If Win32 Then
  
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
      If GetVersionEx(OSV) = 1 Then

         IsWinNTPlus = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And _
                      (OSV.dwVerMajor >= 4)
      End If

   #End If

End Function


Private Function EnableShutdownPrivledges() As Boolean

   Dim hProcessHandle As Long
   Dim hTokenHandle As Long
   Dim lpv_la As LUID
   Dim token As TOKEN_PRIVILEGES
   
   hProcessHandle = GetCurrentProcess()
   
   If hProcessHandle <> 0 Then
   
     'open the access token associated
     'with the current process. hTokenHandle
     'returns a handle identifying the
     'newly-opened access token
      If OpenProcessToken(hProcessHandle, _
                        (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), _
                         hTokenHandle) <> 0 Then
   
        'obtain the locally unique identifier
        '(LUID) used on the specified system
        'to locally represent the specified
        'privilege name. Passing vbNullString
        'causes the api to attempt to find
        'the privilege name on the local system.
         If LookupPrivilegeValue(vbNullString, _
                                 "SeShutdownPrivilege", _
                                 lpv_la) <> 0 Then
         
           'TOKEN_PRIVILEGES contains info about
           'a set of privileges for an access token.
           'Prepare the TOKEN_PRIVILEGES structure
           'by enabling one privilege.
            With token
               .PrivilegeCount = 1
               .laa.udtLUID = lpv_la
               .laa.dwAttributes = SE_PRIVILEGE_ENABLED
            End With
   
           'Enable the shutdown privilege in
           'the access token of this process.
           'hTokenHandle: access token containing the
           '  privileges to be modified
           'DisableAllPrivileges: if True the function
           '  disables all privileges and ignores the
           '  NewState parameter. If FALSE, the
           '  function modifies privileges based on
           '  the information pointed to by NewState.
           'token:  TOKEN_PRIVILEGES structure specifying
           '  an array of privileges and their attributes.
           '
           'Since were just adjusting to shut down,
           'BufferLength, PreviousState and ReturnLength
           'can be passed as null.
            If AdjustTokenPrivileges(hTokenHandle, _
                                     False, _
                                     token, _
                                     ByVal 0&, _
                                     ByVal 0&, _
                                     ByVal 0&) <> 0 Then
                                     
              'success, so return True
               EnableShutdownPrivledges = True
   
            End If  'AdjustTokenPrivileges
         End If  'LookupPrivilegeValue
      End If  'OpenProcessToken
   End If  'hProcessHandle

End Function



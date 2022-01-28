Attribute VB_Name = "mdlNotify"
Option Explicit
Option Private Module

Private Const APP_SYSTRAY_ID = 999 'unique identifier

Private Const NOTIFYICON_VERSION = &H3

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIM_VERSION = &H5

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

'icon flags
Private Const NIIF_NONE = &H0
Private Const NIIF_INFO = &H1
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_GUID = &H5
Private Const NIIF_ICON_MASK = &HF
Private Const NIIF_NOSOUND = &H10
   
Private Const WM_USER = &H400
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)

'shell version / NOTIFIYICONDATA struct size constants
Private Const NOTIFYICONDATA_V1_SIZE As Long = 88  'pre-5.0 structure size
Private Const NOTIFYICONDATA_V2_SIZE As Long = 488 'pre-6.0 structure size
Private Const NOTIFYICONDATA_V3_SIZE As Long = 504 '6.0+ structure size
Private NOTIFYICONDATA_SIZE As Long

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" _
   Alias "Shell_NotifyIconA" _
  (ByVal dwMessage As Long, _
   lpData As NOTIFYICONDATA) As Long

Private Declare Function GetFileVersionInfoSize Lib "version.dll" _
   Alias "GetFileVersionInfoSizeA" _
  (ByVal lptstrFilename As String, _
   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" _
   Alias "GetFileVersionInfoA" _
  (ByVal lptstrFilename As String, _
   ByVal dwHandle As Long, _
   ByVal dwLen As Long, _
   lpData As Any) As Long
   
Private Declare Function VerQueryValue Lib "version.dll" _
   Alias "VerQueryValueA" _
  (pBlock As Any, _
   ByVal lpSubBlock As String, _
   lpBuffer As Any, _
   nVerSize As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Public Sub ShellTrayAdd()
   
   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
  'set up the type members
   With nid
   
      .cbSize = NOTIFYICONDATA_SIZE
      .hWnd = frmSocket.hWnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
      .dwState = NIS_SHAREDICON
      .hIcon = frmSocket.Icon
      
      'szTip is the tooltip shown when the
      'mouse hovers over the systray icon.
      'Terminate it since the strings are
      'fixed-length in NOTIFYICONDATA.
      .szTip = "Virtual Client" & vbNullChar
      .uTimeoutAndVersion = NOTIFYICON_VERSION
   End With
   
  'add the icon ...
   Call Shell_NotifyIcon(NIM_ADD, nid)
   
  '... and inform the system of the
  'NOTIFYICON version in use
   Call Shell_NotifyIcon(NIM_SETVERSION, nid)
       
End Sub


Public Sub ShellTrayRemove()

   Dim nid As NOTIFYICONDATA
   
   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
      
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hWnd = frmSocket.hWnd
      .uID = APP_SYSTRAY_ID
   End With
   
   Call Shell_NotifyIcon(NIM_DELETE, nid)

End Sub


Private Sub ShellTrayModifyTip(nIconIndex As Long, nTitle As String, nMessage As String)

   Dim nid As NOTIFYICONDATA

   If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With nid
      .cbSize = NOTIFYICONDATA_SIZE
      .hWnd = frmSocket.hWnd
      .uID = APP_SYSTRAY_ID
      .uFlags = NIF_INFO
      .dwInfoFlags = nIconIndex
      
      'InfoTitle is the balloon tip title,
      'and szInfo is the message displayed.
      'Terminating both with vbNullChar prevents
      'the display of the unused padding in the
      'strings defined as fixed-length in NOTIFYICONDATA.
      .szInfoTitle = nTitle & vbNullChar
      .szInfo = nMessage
   End With

   Call Shell_NotifyIcon(NIM_MODIFY, nid)

End Sub


Private Sub SetShellVersion()

   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select

End Sub


Private Function IsShellVersion(ByVal version As Long) As Boolean

  'returns True if the Shell version
  '(shell32.dll) is equal or later than
  'the value passed as 'version'
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         IsShellVersion = nVerMajor >= version
      
      End If  'VerQueryValue
    
   End If  'nBufferSize
  
End Function

Public Sub showMessage(nMessage As String, Optional nTitle As String = vbNullString, Optional nIcon As Long = 0)
    Call ShellTrayModifyTip(nIcon, nTitle, nMessage)
End Sub


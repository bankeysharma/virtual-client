Attribute VB_Name = "mdlComm"
Option Explicit
Option Private Module

Public Const cSignalDelimiter = ":"
Public Const cClauseDelimiter = "|"

Public Const cAckHandshaking As String = "%HI%"
Public Const cAckPositive As String = "%SUCCEED%"

Public Const cSigHandshaking As String = "%HELLO%"
Public Const cSigSendDesktop As String = "%DESKTOP%"
Public Const cSigSendFileSystem As String = "%FILESYSTEM%"
Public Const cSigSendComplete As String = "%DONE%"
Public Const cSigNotify As String = "%NOTIFICATION%"
Public Const cSigClientComputerName As String = "%COMPUTERNAME%"
Public Const cSigClientUserName As String = "%USERNAME%"
Public Const cSigShutDown As String = "%SHUTDOWN%"
Public Const cSigReboot As String = "%REBOOT%"
Public Const cSigLogOff As String = "%LOGOFF%"

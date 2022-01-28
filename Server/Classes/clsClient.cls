VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Enum enmSocketState
    
    enmSckConnected = 0
    enmSckConnecting = 1
    enmSckClosed = 2
    
End Enum


Public Enum enmSignals

    enmSigNone = -1
    enmSigNotify = 0
    enmSigHandShaking = 1
    enmSigSendDesktop = 2
    enmSigSendFileSystem = 3
    enmSigSendClientComputerName = 4
    enmSigSendClientUserName = 5
    enmSigShutDown = 6
    enmSigReboot = 7
    enmSigLogOff = 8
    
End Enum

'*************************************
'*************************************
' Attributes
'*************************************
'*************************************
Private vID As String * 5
Private vAlias As String
Private vHostName As String
Private vIPAddress As String
Private vCommunicationPort As Integer
Private vUserName As String

Private eSocketState As enmSocketState
Private eLastSignal As enmSignals

Private vIsBeingDisplayed As Boolean
Private vIsConnected As Boolean
Private vIsON As Boolean
Private vIsResponding As Boolean

Private oTreeNode As MSComctlLib.Node
Private oTerminal As Form
Private oDesktopImage As New stdole.StdPicture
Private oSocket As Winsock

Private Sub Class_Initialize()
    vAlias = vbNullString
    vHostName = vbNullString
    vUserName = vbNullString
    vIPAddress = vbNullString
    vCommunicationPort = 0
    
    vIsBeingDisplayed = False
    vIsConnected = False
    vIsON = False
    eLastSignal = enmSigNone
    
    Set oDesktopImage = Nothing
    
End Sub

Private Sub Class_Terminate()
    
    On Error Resume Next
    
    frmMDI.trvClients.Nodes.Remove (oTreeNode.key)
    Set oTreeNode = Nothing
    
    If (Not oTerminal Is Nothing) Then
        Unload oTerminal
        Set oTerminal = Nothing
    End If
    
    Set oDesktopImage = Nothing
    If (Not oSocket Is Nothing) Then
        Unload oSocket
        Set oSocket = Nothing
    End If

End Sub

'*************************************
'*************************************
' Proerties
'*************************************
'*************************************

'*********************************************
' NAME: Get IsResponding
' PURPOSE: 2 know whether client is in work
' progress against previous signal or not?
'*********************************************
Public Property Get IsResponding() As Boolean
    IsResponding = vIsResponding
End Property

Private Property Let IsResponding(Value As Boolean)
    vIsResponding = Value
    Call mdlMsg.showStaus(Me)
End Property

'*************************************
' NAME: Let/Get Last Signal
' PURPOSE: 2 know/set what was the last
' signal 2 client
'*************************************
Public Property Get LastSignal() As enmSignals
    LastSignal = eLastSignal
End Property

'Public Property Let LastSignal(Signal As enmSignals)
'    eLastSignal = Signal
'End Property

'*************************************
' NAME: Get SocketState
' PURPOSE: 2 know the state of socket
'*************************************
Public Property Get SocketState() As enmSocketState
    SocketState = eSocketState
End Property

'*************************************
' NAME: Get IsBeingDisplayed
' PURPOSE: 2 know whether Client
' is selected 2 display its terminal or
' not?
'*************************************
Public Property Get IsBeingDisplayed() As Boolean
    IsBeingDisplayed = vIsBeingDisplayed
End Property

'*************************************
' NAME: Get ShouldDisplay
' PURPOSE: 2 determine whether Client
' is selected 2 display its terminal or
' not?
'*************************************
Public Property Let shouldDisplay(Value As Boolean)
    vIsBeingDisplayed = Value
    
    If (Not oTreeNode Is Nothing) Then _
        oTreeNode.Checked = Value
    
End Property

'*************************************
' NAME: Let/ Get ID
'*************************************
Public Property Let ID(ByVal Val As String)
    vID = Val
End Property

Public Property Get ID() As String
    ID = vID
End Property

'*************************************
' NAME: Let/ Get Alias 4 Client
'*************************************
Public Property Let Alias(ByVal pAlias As String)
    
    vAlias = StrConv(pAlias, vbProperCase)
    
    If (Not oTreeNode Is Nothing) Then
        
        oTreeNode.text = vAlias
         
        'Following line intended
        'just to remove unwanted (non printable)
        'characters from the alias
        
        vAlias = oTreeNode.text
        
    End If
    
End Property

Public Property Get Alias() As String
    Alias = vAlias
End Property

'*************************************
' NAME: Get User Name working at Client
'*************************************
Public Property Get UserName() As String
    UserName = vUserName
End Property

'*******************************************************************

'*************************************
' NAME: Let/ Get Host Name of Client
'*************************************
Public Property Let hostName(ByVal pHostName As String)
    vHostName = pHostName
End Property

Public Property Get hostName() As String
    hostName = vHostName
End Property

'*******************************************************************

'*************************************
' NAME: Let/ Get IP Address of Client
'*************************************
Public Property Let IPAddress(ByVal pIPAddress As String)
    vIPAddress = pIPAddress
End Property

Public Property Get IPAddress() As String
    IPAddress = vIPAddress
End Property

'*******************************************************************

'**************************************
' NAME: Let/ Get Communication Port of Client
'**************************************
Public Property Let communicationPort(ByVal pCommPort As String)
    vCommunicationPort = pCommPort
End Property

Public Property Get communicationPort() As String
    communicationPort = vCommunicationPort
End Property

'*******************************************************************

'*****************************************
' NAME: IsConnected
' PURPOSE: Get IsConnected 2 determine
' whether this client is connected or not
'*****************************************

Public Property Get IsConnected() As Boolean
    IsConnected = vIsConnected
End Property

'*****************************************
' NAME: IsON
' PURPOSE: Get IsON 2 determine
' whether this client is Visible or not
'*****************************************

Public Property Get IsON() As Boolean
    IsON = vIsON
End Property

'*******************************************************************

'*****************************************
' NAME: Let/ Get Tree Node
'*****************************************
Public Property Set treeNode(ByRef Node As MSComctlLib.Node)
    Set oTreeNode = Node
End Property

Public Property Get treeNode() As MSComctlLib.Node
    Set treeNode = oTreeNode
End Property

'*******************************************************************

'*****************************************
' NAME: Set/ Get Terminal
'*****************************************
Public Property Set Terminal(ByVal Form As Form)
    Set oTerminal = Form
End Property

Public Property Get Terminal() As Form
    Set Terminal = oTerminal
End Property

'*****************************************
' NAME: Set/ Get DesktopImage
' PURPOSE: Object, 2 to keep desktop Image
' of the client
'*****************************************
Public Property Set desktopImage(ByVal Image As stdole.StdPicture)
    Set oDesktopImage = Image
End Property

Public Property Get desktopImage() As stdole.StdPicture
    Set desktopImage = oDesktopImage
End Property

'*****************************************
' NAME: Set/ Get Socket
' PURPOSE: Keeps the refrence of the socket
' connected 2 client
'*****************************************
Public Property Set Socket(ByRef pSocket As Winsock)
    Set oSocket = pSocket
End Property

Public Property Get Socket() As Winsock
    Set Socket = oSocket
End Property

'****************************************
'****************************************
' Methods
'****************************************
'****************************************

'************************************************
' NAME:
' PURPOSE: To display desktop image of the client
' on respective terminal
'************************************************
Public Sub showDesktop(Optional Image As stdole.StdPicture = Nothing)
    
    If (Not Image Is Nothing) Then Set oDesktopImage = Image
    
    If (mdlViews.IsFullScreen) Then
        Set frmFullScreen.picDesktop.Picture = oDesktopImage
    Else
        If (Me.IsON) Then
            If (oTerminal Is Nothing) Then Exit Sub
            Set oTerminal.imgDesktop.Picture = oDesktopImage
        ElseIf (activeClient.ID = vID) Then
            Set frmServices.imgDesktop.Picture = oDesktopImage
            frmServices.frameDesktop.Visible = True
        End If
End If
    
End Sub


'************************************************
' NAME:
' PURPOSE: To place terminal at work space
'************************************************
Public Sub showTerminal()
    
    If (oTerminal Is Nothing) Then Exit Sub
    
    Load oTerminal
    With oTerminal
        .Caption = vAlias
        .Tag = vID
        Call .Show
    End With
    vIsON = True
    
    Call Me.sendSignal(enmSigSendDesktop)

End Sub

'************************************************
' NAME:
' PURPOSE: To make the client terminal off
'************************************************
Public Sub Off()
    
    vIsON = False
    Unload oTerminal

End Sub

'************************************************
' NAME:
' PURPOSE: To establish communication session
'************************************************
Public Sub Connect()

    If (oSocket Is Nothing) Then Exit Sub

    If (vIPAddress = vbNullString) Then Exit Sub

    If (vCommunicationPort = 0) Then Exit Sub

    eSocketState = enmSckConnecting
    Call frmSockets.connectClient(Me)

    If (oSocket.State = sckConnected) Then
        eSocketState = enmSckConnected
        vIsConnected = True
    Else
        eSocketState = enmSckClosed
        vIsConnected = False
    End If

    If (Not oTreeNode Is Nothing) Then
        Call placeImage2Node(oTreeNode)
        vIsBeingDisplayed = oTreeNode.Checked
    End If
    
    vIsResponding = False
    
    Call Me.sendSignal(enmSigSendClientComputerName)
    DoEvents
    Call Me.sendSignal(enmSigSendClientUserName)
    
End Sub

'************************************************
' NAME:
' PURPOSE: To end the communication session
'************************************************
Public Sub Disconnect()
    
    If (oSocket Is Nothing) Then Exit Sub

    If (oSocket.State <> sckClosed) Then oSocket.Close

    vIsConnected = False
    
    If (Not oTreeNode Is Nothing) Then _
        Call placeImage2Node(oTreeNode)
    
    Unload oTerminal
    
    IsResponding = False
    
End Sub


'************************************************
' NAME:
' PURPOSE: To send a signal
'************************************************
Public Sub sendSignal(Signal As enmSignals, Optional auxMessage As String = vbNullString)
    
    Dim vMessage As String
        
    If (vIsResponding) Then
        
        mdlMsg.showMessage (vAlias & " is already busy...")
        
    Else
        
        auxMessage = cSignalDelimiter & auxMessage
        
        If (oSocket Is Nothing) Then
        
            Exit Sub
            
        ElseIf (oSocket.State <> sckConnected) Then
        
            Exit Sub
            
        End If
            
        IsResponding = True
        
        Select Case Signal
        
            Case enmSigHandShaking
            
                Call oSocket.SendData(cSigHandshaking & auxMessage)
            
            Case enmSigSendFileSystem
                
                Call oSocket.SendData(cSigSendFileSystem & auxMessage)
            
            Case enmSigSendDesktop
                
                Call oSocket.SendData(cSigSendDesktop & auxMessage)
                Call Me.showDesktop
            
            Case enmSigNotify
                
                Call oSocket.SendData(cSigNotify & auxMessage)
            
            Case enmSigSendClientComputerName
                
                Call oSocket.SendData(cSigClientComputerName & auxMessage)
                
            Case enmSigSendClientUserName
                
                Call oSocket.SendData(cSigClientUserName & auxMessage)
                
            Case enmSigLogOff
            
                vMessage = "You are forcing LogOff to the Client: " & vAlias & vbCrLf & vbCrLf & _
                            "WARNING: " & vbCrLf & vbCrLf & _
                            "Client will not be given the chance to save its state or" & vbCrLf & _
                            "data of running processes and applications." & vbCrLf & vbCrLf & _
                            "Proceed with LOGOFF?"
                    
                If (mdlMsg.showConfirm(vMessage) = vbNo) Then
                    IsResponding = False
                    Exit Sub
                End If
                
                vUserName = vbNullString
                Call oSocket.SendData(cSigLogOff & auxMessage)
                
            Case enmSigReboot
            
                vMessage = "You are forcing Reboot to the Client: " & vAlias & vbCrLf & vbCrLf & _
                            "WARNING: " & vbCrLf & vbCrLf & _
                            "Client will not be given the chance to save its state or" & vbCrLf & _
                            "data of running processes and applications." & vbCrLf & vbCrLf & _
                            "Proceed with REBOOT?"
                
                If (mdlMsg.showConfirm(vMessage) = vbNo) Then
                    IsResponding = False
                    Exit Sub
                End If
                
                Call oSocket.SendData(cSigReboot & auxMessage)
                
            Case enmSigShutDown
            
                vMessage = "You are forcing ShutDown to the Client: " & vAlias & vbCrLf & vbCrLf & _
                            "WARNING: " & vbCrLf & vbCrLf & _
                            "Client will not be given the chance to save its state or" & vbCrLf & _
                            "data of running processes and applications." & vbCrLf & vbCrLf & _
                            "Proceed with SHUTDOWN?"
                
                If (mdlMsg.showConfirm(vMessage) = vbNo) Then
                    IsResponding = False
                    Exit Sub
                End If
                
                Call oSocket.SendData(cSigShutDown & auxMessage)
                
        End Select
        
        eLastSignal = Signal
    
    End If
    
    '****************************
    '*** Let rest events executed
    '****************************
    'DoEvents

End Sub

'************************************************
' NAME:
' PURPOSE: To collect arriving packets
'************************************************

Public Sub collectPackets()
    
    Dim vRecvData As Variant
    
    Call oSocket.PeekData(vRecvData, vbString)
    
    If (eLastSignal = enmSigSendClientComputerName) Then
        
        Dim vComputerName As String
        
        Call oSocket.GetData(vComputerName, vbString)
        
        vHostName = vComputerName
        Me.Alias = vHostName

        IsResponding = False
        
    ElseIf (eLastSignal = enmSigSendClientUserName) Then
        
        Call oSocket.GetData(vUserName, vbString)
        
        IsResponding = False
        
    ElseIf (VBA.Right(vRecvData, Len(cSigSendComplete)) = cSigSendComplete) Then

        Select Case (eLastSignal)
            Case enmSigSendDesktop
            
                Call getImage
                Call showDesktop
            Case enmSigSendFileSystem
            
                'flushing whatever is in the buffer
                Call oSocket.GetData(vRecvData, vbString)
                
                Call buildExplorerTree(VBA.left(vRecvData, Len(vRecvData) - _
                    Len(cSigSendComplete)))
        End Select
    
        IsResponding = False
    
    ElseIf (vRecvData = cAckPositive) Then
        
        'flushing whatever is in the buffer
        Call oSocket.GetData(vRecvData, vbString)
        
        IsResponding = False
    
    End If

End Sub

'************************************************
' NAME:
' PURPOSE: To make image from collected packets
'************************************************

Private Sub getImage()
    
    On Error GoTo Err_Handle
    Err.Clear
    
    Dim vFileNumber As Byte
    Dim vImageFileName As String
    Dim oFSO As New FileSystemObject
    Dim aryRecvData() As Byte

    vFileNumber = FreeFile()

    ReDim aryRecvData(oSocket.BytesReceived) As Byte

    Call oSocket.GetData(aryRecvData, vbArray + vbByte)
    
    vImageFileName = oFSO.BuildPath(App.Path, oFSO.GetBaseName(oFSO.GetTempName) & ".Jpg")

    Open vImageFileName For Binary Access Write As #vFileNumber
        
        Put #vFileNumber, , aryRecvData
        'Dim vByteCount As Long

        'For vByteCount = 0 To UBound(aryRecvData) - Len(cSigSendComplete)
        '    Put #vFileNumber, , aryRecvData(vByteCount)
        'Next vByteCount

    Close #vFileNumber

    Set oDesktopImage = LoadPicture(vImageFileName)
    
    Call oFSO.DeleteFile(vImageFileName)

Exit_Lable:
    
    Exit Sub
    
Err_Handle:
    
    Select Case (Err.Number)
        Case 481 'Invalid Picture
            Call mdlMsg.showMessage("Invalid picture format received from " & vAlias)
            Resume Next
        Case Else
            Call oDB.showError(Err)
    End Select
    
End Sub

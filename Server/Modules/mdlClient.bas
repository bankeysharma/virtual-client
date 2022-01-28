Attribute VB_Name = "mdlClient"
Option Explicit
Option Private Module

'********** Refrence to Selected Client
Private oActClient As clsClient

Public colClients As New VBA.Collection

Public Property Set activeClient(Obj As clsClient)
    
    If (Not oActClient Is Nothing) Then _
        oActClient.treeNode.Bold = False
    
    Set oActClient = Obj
        
    If (Not oActClient Is Nothing) Then
    
        frmMDI.lblClientName.Caption = oActClient.Alias
        If (oActClient.UserName = vbNullString) Then
            Call oActClient.sendSignal(enmSigSendClientUserName)
        End If
        frmMDI.lblClientInfo.Caption = "User: " & oActClient.UserName
        
        oActClient.treeNode.Bold = True
        Call mdlMsg.showStaus(activeClient)
    Else
        frmMDI.lblClientName.Caption = vbNullString
        frmMDI.lblClientInfo.Caption = vbNullString
    End If

End Property

Public Property Get activeClient() As clsClient

    Set activeClient = oActClient

End Property


'***********************************************
' NAME:
' PURPOSE: To build Client Collection
'***********************************************

Public Sub buildClientCollection()
    
    On Error GoTo Err_Handle
    Err.Clear
        
    Dim oClient As clsClient
    Dim oTerminal As frmDesktop
    
    Dim vClientID As Byte
    
    For vClientID = 1 To colClients.Count
        colClients.Remove (vClientID)
    Next vClientID
    
'    vClientID = 1
'    Do While (vClientID <= 25)

    If (orsIPAddresses.BOF And orsIPAddresses.EOF) Then
    
        mdlMsg.showError ("I am unable to fetch information" & _
                            vbCrLf & "regarding even a single client")
        Exit Sub
        
    End If
    Call orsIPAddresses.moveFirst
    
    While (Not orsIPAddresses.EOF)
        
        Set oClient = New clsClient
        Set oTerminal = New frmDesktop
        
        With oClient
            .ID = "C" & CStr(vClientID)
            .IPAddress = orsIPAddresses!IPAddress 'vbNullString
            .communicationPort = 1001
            .Alias = "Anonymous " & vClientID
            Set .Terminal = oTerminal
            
            If (vClientID > 1) Then Load frmSockets.sckServer(vClientID)
            Set .Socket = frmSockets.sckServer(vClientID)
        
        End With
                    
        'Call colClients.Add(Item:=oClient, key:=oClient.hostName)
        Call colClients.Add(Item:=oClient, key:=oClient.ID)
        
        vClientID = vClientID + 1
        
        orsIPAddresses.moveNext
        
    Wend
    'Loop
    orsIPAddresses.moveFirst
    
    '******** For test only
'    With colClients(1)
'        .IPAddress = frmSockets.sckServer(1).LocalIP
'        .communicationPort = 1001
'    End With
    
    Exit Sub

Err_Handle:
    
    Select Case (Err.Number)
        Case 380 'Invalid Value
            Resume Next
        Case Else
            Call oDB.showError(Err)
    End Select
End Sub

'***********************************************
'******* To build Client Addressing Tree
'***********************************************
Public Sub buildClientTree()

    On Error GoTo Err_Handle
    Err.Clear
    
    Set frmMDI.trvClients.ImageList = frmMDI.imgList16
        
    '**************** First clear existing nodes
    frmMDI.trvClients.Nodes.Clear
    
    Dim vClientID As Byte
    Dim oClient As clsClient
    Dim oNewNodeParent As MSComctlLib.Node
    
    Set oNewNodeParent = frmMDI.trvClients.Nodes.Add
    With oNewNodeParent
        .text = "MCA Lab"
        .key = "MCALAB"
    End With
    
    Call placeImage2Node(oNewNodeParent)
    
    vClientID = colClients.Count
    Do While (vClientID > 0)
        
        Set oClient = colClients.Item(vClientID)
        
        Set oClient.treeNode = frmMDI.trvClients.Nodes.Add
        With oClient.treeNode
            .text = oClient.Alias
            '.key = oClient.hostName
            .key = oClient.ID
            Set .Parent = oNewNodeParent
        End With
        
        Call placeImage2Node(oClient.treeNode)
    
        vClientID = vClientID - 1
    
    Loop
    
    Exit Sub

Err_Handle:
    
    Select Case (Err.Number)
        Case 380 'Invalid Value
            Resume Next
        Case Else
            Call oDB.showError(Err)
    End Select
    
        
End Sub


Public Sub placeImage2Node(ByRef Node As MSComctlLib.Node)
    
    If (Node.Parent Is Nothing) Then
        Node.Image = "imgNetwork"
    ElseIf (colClients.Item(Node.key).IsConnected = False) Then
        Node.Image = "imgMyComputerOffline2"
        Node.Checked = False
    Else
        Node.Image = "imgMyComputer"
    End If

End Sub


VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSockets 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrStateVerifier 
      Interval        =   1000
      Left            =   3015
      Top             =   540
   End
   Begin VB.Timer tmrRequestBreaker 
      Interval        =   15000
      Left            =   3915
      Top             =   450
   End
   Begin VB.Timer tmrProcess 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3960
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar prgBarConnection 
      Height          =   1050
      Left            =   945
      TabIndex        =   5
      Top             =   -45
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   1852
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   960
      TabIndex        =   4
      Top             =   0
      Width           =   960
      Begin VB.Image imgLogo 
         Height          =   615
         Left            =   90
         Picture         =   "frmConnection.frx":0000
         Stretch         =   -1  'True
         Top             =   180
         Width           =   780
      End
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   1
      Left            =   2970
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connecting"
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   540
      Width           =   810
   End
   Begin VB.Label lblClient 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client 1"
      Height          =   195
      Left            =   1800
      TabIndex        =   2
      Top             =   270
      Width           =   525
   End
   Begin VB.Label lblHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   195
      Index           =   1
      Left            =   1215
      TabIndex        =   1
      Top             =   540
      Width           =   495
   End
   Begin VB.Label lblHeading 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client:"
      Height          =   195
      Index           =   0
      Left            =   1275
      TabIndex        =   0
      Top             =   270
      Width           =   435
   End
End
Attribute VB_Name = "frmSockets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vRespondedData As String
Private eRequestStatus As enmResponse
Private vIsConnectingPool As Boolean

Public Sub setProgressBar(Max As Byte)
    Me.prgBarConnection.Max = Max
    Me.prgBarConnection.Min = 0
    Me.prgBarConnection.Value = 0
End Sub

Public Sub connectPool()
        
    vIsConnectingPool = True
    frmMDI.Enabled = False
    setProgressBar (colClients.Count)
    Load Me
    Me.Show
        
    Dim oClient As clsClient
    For Each oClient In colClients
        oClient.Connect
    Next oClient
        
    vIsConnectingPool = False
    frmMDI.Enabled = True
    Unload Me

End Sub

Public Sub connectClient(ByRef Client As clsClient)
        
    On Error GoTo Err_Handle
    Err.Clear
        
    If (Not vIsConnectingPool) Then
        frmMDI.Enabled = False
        setProgressBar (1)
        Load Me
        Me.Show
    End If
    
    With Client
        
        Me.lblClient.Caption = .Alias
        Me.lblStatus.Caption = "Connecting"
        
        If (.Socket.State <> sckClosed) Then .Socket.Close
        
        '***** Connecting
        .Socket.RemotePort = .communicationPort
        Call .Socket.Connect(.IPAddress)
                            
        'Waiting 4 response whether will connect
        'to client or not?
        eRequestStatus = resWaiting
        Do While (eRequestStatus = resWaiting)
            DoEvents
        Loop
        '**************************************
        
        If (eRequestStatus = resPositive) Then
            
            '****** Handshaking
            Me.lblStatus.Caption = "Handshaking"
            Call .sendSignal(enmSigHandShaking)
            
            'Waiting 4 acknowledgement
            eRequestStatus = resWaiting
            Do While (eRequestStatus = resWaiting)
                DoEvents
            Loop
                    
            If (eRequestStatus = resPositive) Then
                'Validating Handshaking
                If (vRespondedData = cAckHandshaking) Then
                    
                    Me.lblStatus.Caption = "Connected"
                
                Else
                    Me.lblStatus.Caption = "Handshaking failed"
                    .Socket.Close
                End If
            
            Else
                Me.lblStatus.Caption = "Handshaking failed"
                .Socket.Close
            End If
                    
        Else
            Me.lblStatus.Caption = "Failed"
            .Socket.Close
        End If
                    
        Me.prgBarConnection.Value = Me.prgBarConnection.Value + 1
    
    End With
    
    Pause (0.5) 'Wait 4 a while
    
    If (Not vIsConnectingPool) Then
        frmMDI.Enabled = True
        Unload Me
    End If

Exit_Lable:
    Exit Sub

Err_Handle:
    Select Case (Err.Number)
        Case 40020 'Invalid IP Address
            Call mdlMsg.showMessage(Err.Description & " IP Address: " & Chr(34) & Client.IPAddress & Chr(34))
        Case 10049 'IP Address not available
            Call mdlMsg.showMessage(Err.Description & " IP Address: " & Chr(34) & Client.IPAddress & Chr(34))
        Case Else
            Call oDB.showError(Err)
    End Select
End Sub

Private Sub Form_Load()
    vIsConnectingPool = False
    'eLastSignal = enmSigNone
    ReDim aryRecvData(0) As Byte
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = Not Cancel
    Me.Hide
End Sub

Private Sub sckServer_Connect(Index As Integer)
    If (colClients(Index).SocketState = enmSckConnecting) Then _
        eRequestStatus = resPositive
End Sub

Private Sub tmrProcess_Timer()
    Static vDots As String
    Dim vLen As Byte
    
    vDots = vDots & "."
    vLen = InStr(1, Me.lblStatus.Caption, ".")
    If (vLen = 0) Then vLen = Len(Me.lblStatus.Caption)
    
    Me.lblStatus.Caption = VBA.Strings.left(Me.lblStatus.Caption, vLen) & vDots
    
    If (Len(vDots) >= 5) Then vDots = vbNullString
End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    On Error Resume Next
    
    If (colClients(Index).SocketState = enmSckConnecting) Then
        Call Me.sckServer(Index).GetData(vRespondedData, vbString)
        eRequestStatus = resPositive
    Else
        Call colClients(Index).collectPackets
    End If

End Sub

Private Sub sckServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    If (colClients(Index).SocketState = enmSckConnecting) Then
        eRequestStatus = resNegative
    Else
        Call showMessage(colClients(Index).Alias & "'s Socket:" & Description)
        Call colClients(Index).Disconnect
    End If

End Sub

Private Sub tmrRequestBreaker_Timer()
    If (eRequestStatus = resWaiting) Then eRequestStatus = resNegative
End Sub

Private Sub tmrStateVerifier_Timer()
    Me.tmrStateVerifier.Enabled = False
    
    Dim oClient As clsClient
    
    For Each oClient In colClients
        If (oClient.IsConnected And oClient.Socket.State <> sckConnected) Then
            Call oClient.Disconnect
            mdlMsg.showMessage (oClient.Alias & " has been disconnected.")
        End If
    Next oClient
    
    Me.tmrStateVerifier.Enabled = True
    
End Sub

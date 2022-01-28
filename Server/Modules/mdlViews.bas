Attribute VB_Name = "mdlViews"
Option Explicit
Option Private Module

Private Const cRandomViewTerminalCount As Byte = 7
Private Const cRandomViewTimeSpan As Byte = 3

Private eView As enmView

Private eCurrentService As enmServices
Private vIsWorking As Boolean

Public IsFullScreen As Boolean

Private Property Let View(Value As enmView)
    
    eView = Value

    With frmMDI
        .mnuViewBroadSpectrum.Checked = False
        .mnuViewRandom.Checked = False
        .mnuViewsFullScreen.Checked = False
    End With
    
    
    Select Case (eView)
        Case vuBroadSpectrum
            IsFullScreen = False
            Unload frmFullScreen
            Call offSolitaryView
            If (frmMDI.tmrRandomView.Enabled) Then _
                Call offAllTerminals
            
            frmMDI.tlbStandard.Buttons("keyBroadSpectrumView").Value = tbrPressed
            frmMDI.mnuViewBroadSpectrum.Checked = True
        Case vuRandom
            IsFullScreen = False
            Unload frmFullScreen
            Call offSolitaryView
            frmMDI.tlbStandard.Buttons("keyRandomView").Value = tbrPressed
            frmMDI.mnuViewRandom.Checked = True
        Case vuSolitary
            IsFullScreen = False
            Unload frmFullScreen
            Call offAllTerminals
            
            frmMDI.tlbStandard.Buttons("keySolitaryView").Value = tbrPressed
            'frmMDI.mnuViewSolitary.Checked = True
        Case vuFullScreen
            Call offAllTerminals
            Call offSolitaryView
            
            'frmMDI.tlbStandard.Buttons("keyFullScreen").Value = tbrPressed
            frmMDI.mnuViewsFullScreen.Checked = True
        Case vuNone
            With frmMDI.tlbStandard
                .Buttons("keyBroadSpectrumView").Value = tbrUnpressed
                .Buttons("keyRandomView").Value = tbrUnpressed
                .Buttons("keySolitaryView").Value = tbrUnpressed
                .Buttons("keyFullScreen").Value = tbrUnpressed
            End With
            Call offAllTerminals
            Call offSolitaryView
            
            Unload frmFullScreen
            IsFullScreen = False
    End Select
    
End Property

Public Property Get View() As enmView
    View = eView
End Property

Public Property Let CurrentService(Value As enmServices)
    
    If (eCurrentService = Value) Then Exit Property
    
    With frmMDI
        
        .cbStandard.Bands("keyExplorer").Visible = False
        .cbStandard.Bands("keyProcess").Visible = False
        
        .mnuExplorer.Visible = False
        .mnuProcess.Visible = False
        
        .mnuViewSolitaryDesktop.Checked = False
        .mnuViewSolitaryProcesses.Checked = False
        .mnuViewSolitaryFileSystem.Checked = False
    
    End With
        
    eCurrentService = Value
    If (Value = srvNone) Then Exit Property
    
    If (activeClient Is Nothing) Then
        mdlMsg.showAlert ("Please! Select a client first!")
        View = vuNone
        Exit Property
    
    ElseIf (Not activeClient.IsConnected) Then
        mdlMsg.showError (UCase(activeClient.Alias) & " is no longer connected, can not be monitored.")
        activeClient.treeNode.Checked = False
        
        View = vuNone
        
        Exit Property
    
    End If
    
    eCurrentService = Value
    
    Select Case (eCurrentService)
        Case srvDesktopCapture
            
            frmServices.tabServices.Tab = 0
            With frmMDI
                .tlbServices.Buttons("keyDesktop").Value = tbrPressed
                .mnuViewSolitaryDesktop.Checked = True
            End With
            
        Case srvProcessMonitoring
            
            frmServices.tabServices.Tab = 1
            With frmMDI
                .mnuViewSolitaryProcesses.Checked = True
                .mnuProcess.Visible = True
                .cbStandard.Bands("keyProcess").Visible = True
                .tlbServices.Buttons("keyProcesses").Value = tbrPressed
            End With
            
        Case srvExplorer
            
            frmServices.tabServices.Tab = 2
            
            With frmMDI
                .mnuViewSolitaryFileSystem.Checked = True
                .mnuExplorer.Visible = True
                .cbStandard.Bands("keyExplorer").Visible = True
                .tlbServices.Buttons("keyExplorer").Value = tbrPressed
            End With
            
    End Select
    
End Property

Public Property Get CurrentService() As enmServices
    CurrentService = eCurrentService
End Property


'*****************************************
'****** Solitary View
'*****************************************

Public Sub solitaryView()
    
    '*********** Check whether 2 enter or not?
    If (vIsWorking) Then Exit Sub
    
    If (eCurrentService = srvNone) Then
        Exit Sub
        
    ElseIf (activeClient Is Nothing) Then
        mdlMsg.showAlert ("Please! Select a client first!")
        View = vuNone
        
        Exit Sub
    
    ElseIf (Not activeClient.IsConnected) Then
        
        mdlMsg.showError (UCase(activeClient.Alias) & " is no longer connected, can not be monitored.")
        activeClient.treeNode.Checked = False
        
        View = vuNone
        
        Exit Sub
    End If
    
    vIsWorking = True
    
    '**********************************************
        
    Load frmServices
    View = vuSolitary
    
    With frmServices
    
        If (eCurrentService = srvProcessMonitoring) Then
            'Nothing to do at startup
            
        ElseIf (eCurrentService = srvExplorer) Then
            
            If (frmServices.trvDirectories.Nodes.Count = 0) Then
                frmServices.cbXplorer.Visible = False
            End If
        
        ElseIf (eCurrentService = srvDesktopCapture) Then
            
            If (activeClient.desktopImage Is Nothing) Then
            
                frmServices.frameDesktop.Visible = False
                
            End If
        
        End If
        
        .Show
        .Caption = activeClient.Alias
    
    End With
    
    vIsWorking = False
    
End Sub

'*******************************************
'***** Broad Spectrum View
'*******************************************
Public Sub broadSpectrumView()
    
    On Error GoTo Err_Handle
    Err.Clear
    
    '*********** Do not enter, if any previous job is going on
    If (vIsWorking) Then Exit Sub
    
    vIsWorking = True
    
    View = vuBroadSpectrum
    
    If (colClients.Count = 0) Then Exit Sub
    
    '***************** ProgressBar
    Call setPrgBar("Portraying Desktop...", 0, colClients.Count)
    '******************************
    
    Dim oClient As clsClient
    For Each oClient In colClients
        prgBar.Value = prgBar.Value + 1
        DoEvents
        If (oClient.IsBeingDisplayed) Then
            Call oClient.showTerminal
        ElseIf (oClient.IsON) Then
            Call oClient.Off
        ElseIf (oClient.treeNode.Checked) Then
            oClient.treeNode.Checked = False
        End If
    Next oClient
    
    If (Not frmMDI.ActiveForm Is Nothing) Then
        If (frmMDI.tlbStandard.Buttons("keyTileHorizontal").Value = tbrPressed) Then
            frmMDI.standardAction ("keyTileHorizontal")
        Else
            frmMDI.standardAction ("keyTileVertical")
        End If
    End If
    
    If (frmMDI.ActiveForm Is Nothing) Then View = vuNone
    
    '*****************************************
    'Closing Progress bar
    '*****************************************
    Call offPrgBar
    
    vIsWorking = False
    
    Exit Sub
    
Err_Handle:
    
    Select Case (Err.Number)
        Case 380 'Invalid Value
            Resume Next
        Case Else
            Call oDB.showError(Err)
    End Select
    
    vIsWorking = False
    
End Sub

'*******************************************
'***** Random View
'*******************************************

Public Sub randomView()
    
    '*********** Do not enter, if any previous job is going on
    If (vIsWorking) Then Exit Sub
    
    vIsWorking = True
    
'    Dim oParent As MSComctlLib.Node
'    Dim oChild As MSComctlLib.Node
'    Dim vTerminalID As Byte
'    Dim vID As Byte
'    Dim vHasLooped As Boolean
    
    frmMDI.tmrRandomView.Interval = cRandomViewTimeSpan * 500
    frmMDI.tmrRandomView.Enabled = False
    
    View = vuRandom

' Code

    If (frmMDI.tlbStandard.Buttons("keyTileHorizontal").Value = tbrPressed) Then
        frmMDI.standardAction ("keyTileHorizontal")
    Else
        frmMDI.standardAction ("keyTileVertical")
    End If
    
    frmMDI.tmrRandomView.Enabled = True

    vIsWorking = False
    
End Sub

'*****************************************
'******** Off all the Terminals
'*****************************************

Private Sub offAllTerminals()
        
    Dim vTmpKey As String
    Dim i As Byte
    Dim oClient As clsClient
    
    For Each oClient In colClients
        If (oClient.IsON) Then oClient.Off
    Next oClient
    
    '********** Disable it if working
    frmMDI.tlbStandard.Buttons("keyTileVertical").Value = tbrUnpressed
    frmMDI.tlbStandard.Buttons("keyTileHorizontal").Value = tbrUnpressed

End Sub

'*****************************************
'******** Off all the Terminals
'*****************************************

Public Sub offSolitaryView()
    Unload frmServices
    
    Set frmServices.oProcesses = Nothing
    CurrentService = srvNone
    
    With frmMDI
        .tlbStandard.Buttons("keySolitaryView").Value = tbrUnpressed
        .cbStandard.Bands("keyExplorer").Visible = False
        .cbStandard.Bands("keyProcess").Visible = False
        .mnuViewSolitaryDesktop.Checked = False
        .mnuViewSolitaryProcesses.Checked = False
        .mnuViewSolitaryFileSystem.Checked = False
    End With
    
    With frmMDI.tlbServices
        .Buttons("keyDesktop").Value = tbrUnpressed
        .Buttons("keyProcesses").Value = tbrUnpressed
        .Buttons("keyExplorer").Value = tbrUnpressed
    End With

End Sub


Public Sub fullScreenView()
    If (activeClient Is Nothing) Then
        mdlMsg.showAlert ("Please! Select a client first!")
        View = vuNone
        
        Exit Sub
    
    ElseIf (Not activeClient.IsConnected) Then
        
        mdlMsg.showError (UCase(activeClient.Alias) & " is no longer connected, can not be monitored.")
        activeClient.treeNode.Checked = False
        
        View = vuNone
        
        Exit Sub
    End If
        
    View = vuFullScreen
    
    IsFullScreen = True
    
    Call activeClient.sendSignal(enmSigSendDesktop)
    
    Load frmFullScreen
    With frmFullScreen
        .WindowState = vbMaximized
        Call .Show(vbModal)
    End With

End Sub


'public sub

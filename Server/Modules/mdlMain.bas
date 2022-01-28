Attribute VB_Name = "mdlMain"
Option Explicit
Option Private Module

Public Enum enmView
    vuNone = -1
    vuSolitary = 0
    vuBroadSpectrum = 1
    vuRandom = 2
    vuFullScreen = 3
End Enum

Public Enum enmServices
    srvNone = -1
    srvDesktopCapture = 0
    srvExplorer = 1
    srvProcessMonitoring = 2
End Enum

Public Enum enmPackageMode
    softDevelopment = 0
    softAlpha = 1
    softBeta = 2
    softInstalled = 3
End Enum

Public Enum enmResponse
    resWaiting = 0
    resNegative = 1
    resPositive = 2
End Enum

Public Const cHeightRatio As Byte = 72 ' actual 71.78

Public gPackageMode As enmPackageMode
Public oDB As New clsDB

Public oConn As New ADODB.Connection
Public orsIPAddresses As New ADODB.RecordSet

Public Property Get prgBar() As MSComctlLib.progressBar
    Set prgBar = frmMDI.progressBar
End Property

Public Sub setPrgBar(Caption As String, minVal As Variant, maxVal As Variant)
    
    If (minVal = maxVal) Then Exit Sub
    
    With frmMDI.cbStandard.Bands("keyProgress")
        .Caption = Caption
        .Visible = True
    End With
    
    With frmMDI.progressBar
        .Min = minVal
        .Max = maxVal + 1
        .Value = 1
    End With
End Sub

Public Sub offPrgBar()
    frmMDI.cbStandard.Bands("keyProgress").Visible = False
End Sub

Public Sub Main()
    On Error GoTo Err_Handle
    Err.Clear
    
    '*************** Allow only single instance
    If (App.PrevInstance) Then End
    '******************************************
    
    Dim vSQL As String
    
    gPackageMode = softDevelopment
'    gPackageMode = softInstalled
    
    Call flushTmp
    
    vSQL = "SELECT IPAddress " & _
            "FROM IPAddresses " & _
            "ORDER BY [IPAddress]"
            
    Set oConn = oDB.getConnection
    
    If (oConn Is Nothing) Then
        
        Call mdlMsg.showError("Possibily! database is missing or currupted.")
        Call Quit
        
    End If
    
    Set orsIPAddresses = oDB.getRecordSet(vSQL, oConn)
    orsIPAddresses.Open
    
    '****************** Loading MDI Form
    Load frmMDI
    
    With frmMDI
        .Show
        .WindowState = vbMaximized
        
        .Caption = ClientName 'Just to truncate any non printable character
        .Caption = "Virtual Client: Server (" & frmMDI.Caption & ") "
    End With
    
    '******************************************
    'Updating IP Address table for a
    'NON-NT base Server
    '******************************************
    Call mdlIPAddresses.updateIPAddressTable
    
    Call buildClientCollection
    
    Call buildClientTree
        
    If (gPackageMode <> softDevelopment) Then Call LogIn
    
    Call frmSockets.connectPool
            
    CurrentService = srvNone
    
    '************************************
    
Exit_Lable:
    Exit Sub

Err_Handle:
    Select Case (Err.Number)
        Case -2147217865
            Call mdlMsg.showError("Currupted database encountered.")
            Call Quit
        Case Else
            Call oDB.showError(Err)
    End Select
    
End Sub

Public Function Quit() As Boolean

    Quit = True
    
    Unload frmFullScreen
    
    If (gPackageMode <> softDevelopment) Then
    
        If (mdlMsg.showConfirm("Wanna Quit?") = vbNo) Then Quit = False
        
    End If

    If (Quit) Then End

End Function


Private Sub deactivateInterface()
    With frmMDI
        .mnuMain.Enabled = False
        .mnuView.Enabled = False
        .mnuUtility.Enabled = False
        '.cbComponents.Enabled = False
        .cbStandard.Enabled = False
    
        .picComponents.Visible = False
        '.mnuMain.Visible = False
        '.mnuView.Visible = False
        '.cbComponents.Visible = False
        '.cbStandard.Visible = False
    
    End With
End Sub

Private Sub activateInterface()
    With frmMDI
        .mnuMain.Enabled = True
        .mnuView.Enabled = True
        .mnuUtility.Enabled = True
        '.cbComponents.Enabled = True
        .cbStandard.Enabled = True
        
        .picComponents.Visible = True
        '.mnuMain.Visible = True
        '.mnuView.Visible = True
        '.cbComponents.Visible = True
        '.cbStandard.Visible = True
    
    End With
End Sub


Public Function LogIn() As Boolean
    Call deactivateInterface
    
    Dim vLoginChance As Byte
    
    For vLoginChance = 0 To 2
        
        LogIn = frmAuthentication.DoAuthentication
        
        If (Not LogIn) Then
            Call mdlMsg.showMessage("Login denied.")
            If (vLoginChance = 2) Then End 'Call Quit
        Else
            Call mdlMsg.showMessage("Login succeed.")
            Exit For
        End If
    
    Next vLoginChance
    
    Unload frmAuthentication
    
    Call activateInterface
End Function


Private Sub flushTmp()
    Dim oFile As Scripting.File
    Dim oFiles As Scripting.Files
    
    Dim oFolder As Scripting.Folder
    
    Dim oFSO As New Scripting.FileSystemObject
    
    Set oFolder = oFSO.GetFolder(App.Path)

    Set oFiles = oFolder.Files
    
    For Each oFile In oFiles
        Select Case (UCase(oFile.Type))
            Case "TMP FILE"
                Call oFile.Delete(True)
            Case "JPEG IMAGE"
                Call oFile.Delete(True)
            Case "BITMAP IMAGE"
                Call oFile.Delete(True)
        End Select
    Next oFile
    
End Sub


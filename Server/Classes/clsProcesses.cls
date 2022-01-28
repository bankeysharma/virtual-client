VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcesses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Sink As SWbemSink
Attribute Sink.VB_VarHelpID = -1
Private Services As SWbemServices

Private oClient As clsClient
Private oContainer As VB.Frame
Private oListView As MSComctlLib.ListView

Public Property Set Client(pClient As clsClient)
    Set oClient = pClient
End Property

Public Property Set containerObject(pObject As VB.Frame)
    Set oContainer = pObject
End Property

Public Property Set listingObject(pObject As MSComctlLib.ListView)
    Set oListView = pObject
End Property

Public Sub enumerateProcesses()
    
    On Error Resume Next
    
    If (oListView Is Nothing Or oClient Is Nothing) Then
        Call mdlMsg.showMessage("Process object is not initialized with Client or display object")
        Exit Sub
    End If
        
    oContainer.Visible = True
        
    With oListView
        .View = lvwReport
        .ListItems.Clear
        .MousePointer = vbHourglass
        .ListItems.Clear
    End With
    
    Dim sComputerName As String
    sComputerName = oClient.IPAddress
    
    ' Create a sink to recieve the results of the enumeration
    Set Sink = New SWbemSink
        
    ' Connect to root\cimv2.
    Set Services = GetObject("winmgmts://" & sComputerName)
    
    ' Perform the asynchronous enumeration of processes
    Call Services.InstancesOfAsync(Sink, "Win32_process")
    
    oListView.MousePointer = vbNormal
    
End Sub

Private Sub Sink_OnCompleted(ByVal iHResult As WbemScripting.WbemErrorEnum, ByVal objWbemErrorObject As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
   ' This event handler is called when there are no more instances to
    ' be returned
    oListView.MousePointer = vbDefault
    
    If (iHResult <> wbemNoErr) Then
        Call mdlMsg.showError(Err.Description)
    End If
    
End Sub

Private Sub Sink_OnObjectReady(ByVal objWbemObject As WbemScripting.ISWbemObject, _
                                ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)

    ' This event handler is called once for every process returned by the
    ' enumeration

    Dim key As String
    Dim Item As MSComctlLib.ListItem

    key = "Handle:" & objWbemObject.Handle

    Set Item = oListView.ListItems.Add(, key, StrConv(objWbemObject.Name, vbProperCase))
    With Item
        .SmallIcon = frmMDI.imgList16.ListImages("imgProcess").Index
    End With
    
'    Item.SubItems(1) = objWbemObject.Handle
    
    Item.SubItems(1) = objWbemObject.Handle

    If vbNull <> VarType(objWbemObject.Priority) Then
        Item.SubItems(2) = objWbemObject.Priority
    End If

    If vbNull <> VarType(objWbemObject.WorkingSetSize) Then
        Item.SubItems(3) = CStr(objWbemObject.WorkingSetSize / 1024) + " K"
    End If

End Sub

Public Sub Kill()
    On Error Resume Next
    
    If (oListView Is Nothing Or oClient Is Nothing) Then
        Call mdlMsg.showMessage("Process object is not initialized with Client or display object")
        Exit Sub
    ElseIf (oListView.SelectedItem Is Nothing) Then
        Call mdlMsg.showMessage("No process is selected.")
        Exit Sub
    End If
    
    Dim sComputerName As String
    Dim vMessage As String
    Dim key As String
    Dim vPid As Variant
    Dim Process As Object ' WbemScripting.SWbemObject
    
    key = oListView.SelectedItem.key
    
    vPid = oListView.SelectedItem.SubItems(1) 'Process ID
    
    sComputerName = oClient.IPAddress
    
    For Each Process In GetObject("winmgmts:{impersonationLevel=impersonate}!//" & sComputerName).ExecQuery("select * from Win32_Process where Handle=" & vPid)
        
        vMessage = "You are going to kill the process:" & vbCrLf & vbCrLf & _
                    "PID : " & oListView.SelectedItem.SubItems(1) & vbCrLf & _
                    "Name: " & oListView.SelectedItem.text & vbCrLf & vbCrLf & _
                    "WARNING: Terminating a process can cause undesired" & vbCrLf & _
                    "results including loss of data and system instability. The" & vbCrLf & _
                    "process will not be given the chance to save its state or" & vbCrLf & _
                    "data before it is terminated." & vbCrLf & vbCrLf & _
                    "Proceed with KILL?"
        
        If (mdlMsg.showConfirm(vMessage) = vbNo) Then Exit Sub
        
        If (Process.Terminate <> 0) Then
            
            vMessage = "The operation could not be completed." & vbCrLf & vbCrLf & _
                        "The operation is not valid for this Process."
            
            Call mdlMsg.showError(vMessage, "Unable to Terminate Process")
            
            Exit Sub
        End If
    
    Next Process
    
    Call oListView.ListItems.Remove(key)
    
    DoEvents
    
End Sub

Public Sub showAssociators()
    On Error Resume Next
    
    If (oListView Is Nothing Or oClient Is Nothing) Then
        Call mdlMsg.showMessage("Process object is not initialized with Client or display object")
        Exit Sub
    ElseIf (oListView.SelectedItem Is Nothing) Then
        Call mdlMsg.showMessage("No process is selected.")
        Exit Sub
    End If
    
    Dim sComputerName As String
    Dim vMessage As String
    Dim vPid As Long
    Dim Process As WbemScripting.SWbemObject
    Dim Processes As WbemScripting.SWbemObjectSet
        
    vPid = oListView.SelectedItem.SubItems(1) 'Process ID
    
    sComputerName = oClient.IPAddress
    
    For Each Process In GetObject("winmgmts:{impersonationLevel=impersonate}!//" & sComputerName).ExecQuery("select * from Win32_Process where Handle=" & vPid)
    
        Set Processes = Process.Associators_
            
    Next Process
    
    
    Dim vInfo As Variant
    vMessage = vbNullString
    
    For Each Process In Processes
    
        vInfo = Split(Process.Path_.DisplayName, Chr(34))
        
        If (UBound(vInfo) >= 1) Then vMessage = vMessage & vInfo(1) & "|"
        
    Next Process
    
    vInfo = Split(vMessage, "|")
    
    If (UBound(vInfo) >= 1) Then
        Call frmShowAssociators.showInfo(vInfo)
    End If
    
    DoEvents
    
End Sub

Public Sub showProperties()
    
    If (oListView Is Nothing Or oClient Is Nothing) Then
        Call mdlMsg.showMessage("Process object is not initialized with Client or display object")
        Exit Sub
    ElseIf (oListView.SelectedItem Is Nothing) Then
        Call mdlMsg.showMessage("No process is selected.")
        Exit Sub
    End If
    
    Dim sComputerName As String
    Dim vMessage As String
    Dim vPid As Long
    Dim Process As WbemScripting.SWbemObject
        
    vPid = oListView.SelectedItem.SubItems(1) 'Process ID
    
    sComputerName = oClient.IPAddress
    
    For Each Process In GetObject("winmgmts:{impersonationLevel=impersonate}!//" & sComputerName).ExecQuery("select * from Win32_Process where Handle=" & vPid)
        Dim a As WbemScripting.WbemImpersonationLevelEnum
        Dim b As WbemScripting.WbemAuthenticationLevelEnum
        
        vMessage = "Authentication Level: "
        
        Select Case (Process.Security_.AuthenticationLevel)
            Case wbemAuthenticationLevelCall
                vMessage = vMessage & "Call"
            Case wbemAuthenticationLevelConnect
                vMessage = vMessage & "Connect"
            Case wbemAuthenticationLevelDefault
                vMessage = vMessage & "Default"
            Case wbemAuthenticationLevelNone
                vMessage = vMessage & "None"
            Case wbemAuthenticationLevelPkt
                vMessage = vMessage & "Pkt"
            Case wbemAuthenticationLevelPktIntegrity
                vMessage = vMessage & "PktIntigrity"
            Case wbemAuthenticationLevelPktPrivacy
                vMessage = vMessage & "PktPrivacy"
        End Select
    
        vMessage = vMessage & vbCrLf & _
                        "Impersonation Level: "
                        
        Select Case (Process.Security_.ImpersonationLevel)
            Case wbemImpersonationLevelAnonymous
                vMessage = vMessage & "Anonymous"
            Case wbemImpersonationLevelDelegate
                vMessage = vMessage & "Delegate"
            Case wbemImpersonationLevelIdentify
                vMessage = vMessage & "Identify"
            Case wbemImpersonationLevelImpersonate
                vMessage = vMessage & "Impersonate"
        End Select
        
    Next Process
    
    mdlMsg.showInfo (vMessage)
    
End Sub

'Public Sub Info()
'    'On Error Resume Next
'
'    If (oListView Is Nothing Or oClient Is Nothing) Then
'        Call mdlMsg.showMessage("Process object is not initialized with Client or display object")
'        Exit Sub
'    End If
'
'    Dim sComputerName As String
'    Dim vMessage As String
'    Dim vPid As Long
'    Dim P2 As WbemScripting.SWbemObject
'    Dim Process As WbemScripting.SWbemObject
'    Dim Processes As WbemScripting.SWbemObjectSet
'
'    vPid = oListView.SelectedItem.SubItems(1) 'Process ID
'
'    sComputerName = oClient.IPAddress
'
'    For Each Process In GetObject("winmgmts:{impersonationLevel=impersonate}!//" & sComputerName).ExecQuery("select * from Win32_Process where Handle=" & vPid)
'        Dim a As WbemScripting.WbemImpersonationLevelEnum
'        Dim b As WbemScripting.WbemAuthenticationLevelEnum
'
'    Next Process
'
''    vMessage = vbNullString
''
''    For Each Process In Processes
''
''        vMessage = vMessage & " * " & Process.Path_.DisplayName & vbCrLf
''
''    Next Process
''
''    MsgBox vMessage
''
'    DoEvents
'
'End Sub



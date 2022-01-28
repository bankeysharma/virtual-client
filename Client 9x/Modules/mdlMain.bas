Attribute VB_Name = "mdlMain"
Option Explicit
Option Private Module

Public Enum enmResponse
    resWaiting = 0
    resNegative = 1
    resPositive = 2
End Enum

Public Sub Main()
    
    '*****************************************
    ' Only Single Instance of the application
    ' Should be in running state
    '*****************************************
    If (App.PrevInstance = True) Then End
    
    Call flushTmp
    
    Load frmSocket
    frmSocket.Hide
    
'    Load frmTest
'    frmTest.Show
    
    Dim oSocketClient As Winsock
    Set oSocketClient = frmSocket.sckClient
    'oSocketClient.Close

    Unload frmSocket

    Call mdlNotify.ShellTrayAdd
    
    'If (Date > #2/10/2004#) Then
    '    With oSocketClient
    '        .Close
    '        .LocalPort = 0
    '    End With
    'End If
        
    'Call mdlNotification.showInformation("Hello test message", "Testing")
    
End Sub

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

Public Sub resetSocket()
    Call frmSocket.sckClient.Close
    Call frmSocket.sckClient.Listen
End Sub

Public Sub showError(error As ErrObject)

    MsgBox "I encounter following UnExpected error." & vbCrLf & vbCrLf & _
            "Number: " & error.Number & vbCrLf & _
            "Description: " & error.Description & vbCrLf & _
            "Source: " & error.Source & vbCrLf & vbCrLf & _
            "Please! contact programmer soon.", vbCritical + vbOKOnly, "UnExpected error..."

End Sub



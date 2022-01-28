Attribute VB_Name = "mdlMsg"
Option Explicit
Option Private Module

Public Function showConfirm(text As String, Optional btn As VbMsgBoxStyle = vbYesNo + vbDefaultButton2) As VbMsgBoxResult
    showConfirm = MsgBox(text, btn + vbQuestion, "Asking...")
End Function

Public Sub showAlert(text As String)
     Call MsgBox(text, vbOKOnly + vbExclamation, "Alert...")
End Sub

Public Sub showError(text As String, Optional Title As String = vbNullString)
    
    If (Title = vbNullString) Then Title = "Error..."
    
    Call MsgBox(text, vbOKOnly + vbCritical, Title)
    
End Sub

Public Sub showInfo(text As String)
    Call MsgBox(text, vbOKOnly + vbInformation, "Info...")
End Sub

' ********************************************
' Procedure to place message at status bar
' ********************************************
Public Sub showMessage(msg As String)
    frmMDI.statusBar.Panels("keyMessage").text = msg
    frmMDI.tmrMessageCleaner.Enabled = True
End Sub

' ********************************************
' Procedure to place status at status bar
' ********************************************
Public Sub showStaus(ByVal oClient As clsClient)
    
    If (oClient.IsResponding) Then
        frmMDI.statusBar.Panels("keyStatus").text = "Responding..."
    Else
        frmMDI.statusBar.Panels("keyStatus").text = "Idle..."
    End If

End Sub



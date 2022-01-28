VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum enmActionType
    enmAcAddSuccess = 0
    enmAcSaveSuccess = 1
    enmAcEditSuccess = 2
    enmAcDeleteSuccess = 3
    enmAcRevert = 4
    enmAcNone = 5
    enmAcSaveFailure = 6
    enmAcDeleteFailure = 7
    enmAcAddFailure = 8
    enmAcEditFailure = 9
End Enum

Public Enum enmMethodFailureCause
    enmExclusivelyLockedDB = 0
    enmExclusivelyLockedTable = 1
    enmUnrecognizedSQL = 2
    enmUnspecifiedCause = 3
End Enum

Private vMethodFailureCause As enmMethodFailureCause

Private vDBPath As String
Private vConnectionString As String

Public Property Get MethodFailureCause() As enmMethodFailureCause
    MethodFailureCause = vMethodFailureCause
End Property

Private Sub Class_Initialize()
    DBPath = App.Path & "\Database\VirtualClient.mdb"
    vConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vDBPath & ";Mode=ReadWrite;Persist Security Info=False"
End Sub

'***************************************************
'******************** Properties
'***************************************************

Public Property Get DBPath() As String
    DBPath = vDBPath
End Property

Public Property Let DBPath(pDBPath As String)
    vDBPath = pDBPath
End Property

Public Property Get ConnectionString() As String
    ConnectionString = vConnectionString
End Property

Public Property Let ConnectionString(pSQL As String)
    vConnectionString = pSQL
End Property

'***********************************************************
'****************** Functions
'***********************************************************
Public Function getConnection(Optional pCursorLocation As ADODB.CursorLocationEnum = adUseClient) As ADODB.Connection
    On Error GoTo Err_Handle
    Err.Clear
    
    Dim myConn As New ADODB.Connection
    
    myConn.ConnectionString = vConnectionString
    myConn.CursorLocation = pCursorLocation
    myConn.Open
    
    Set getConnection = myConn 'Returning Connection Object
    
    Exit Function
    
Err_Handle:
    
    Select Case (Err.Number)
        Case -2147467259
            mdlMsg.showError ("Requested DataStore can't be open" & vbCrLf & _
                                "Probably another one may have exclusive rights over DataStore")
            Set getConnection = Nothing
            vMethodFailureCause = enmExclusivelyLockedDB
        Case Else
            Call Me.showError(Err)
    End Select
End Function

Public Function getCommand(ByRef pConnection As ADODB.Connection, Optional pCommandType As ADODB.CommandTypeEnum = adCmdText) As ADODB.Command
    Dim myCmd As New ADODB.Command
    
    Set myCmd.ActiveConnection = pConnection
    myCmd.CommandType = pCommandType
    myCmd.Prepared = True

    Set getCommand = myCmd 'Returning command object
End Function

Public Function getRecordSet(Optional pSource As String = "", Optional pActConnection As ADODB.Connection = Nothing, Optional pCursorType As ADODB.CursorTypeEnum = adOpenDynamic, Optional pLockType As ADODB.LockTypeEnum = adLockOptimistic, Optional pCursorLocation As ADODB.CursorLocationEnum = adUseClient) As ADODB.RecordSet
    Dim myRs As New ADODB.RecordSet
    
    If Not (pActConnection Is Nothing) Then
        myRs.ActiveConnection = pActConnection
    End If
    
    myRs.CursorLocation = pCursorLocation
    myRs.CursorType = pCursorType
    myRs.LockType = pLockType
    myRs.Source = pSource
    
    Set getRecordSet = myRs 'Returning recordset object
End Function

Public Function Execute(SQLStatement As String, Optional ByRef ADOConnection As ADODB.Connection = Nothing) As ADODB.RecordSet
    Dim myConn As New ADODB.Connection
  
    If (ADOConnection Is Nothing) Then
        Set myConn = Me.getConnection
        If (myConn.State = adStateClosed) Then
            Set Execute = Nothing
            Exit Function
        End If
    Else
        Set myConn = ADOConnection
    End If
    
    Set Execute = myConn.Execute(SQLStatement)  'Returns
    

End Function

Public Function AddSave(ByRef callerForm As Form, ByRef RecordSet As ADODB.RecordSet, ByRef cmdAdd As CommandButton, ByRef cmdEdit As CommandButton, Optional ByVal doCheck As Boolean = True) As enmActionType
    On Error GoTo ErrHandle
    Err.Clear
    
    
    If (Trim(UCase(cmdAdd.Caption)) = "&ADD") Then
        
        AddSave = enmAcAddFailure
        
        If (RecordSet.Supports(adAddNew)) Then
            AddSave = enmAcAddSuccess
            RecordSet.AddNew
            cmdAdd.Caption = "&Commit"
            cmdEdit.Caption = "&Revert"
            callerForm.makeEditable
            callerForm.clearUI
        End If
    Else
        If (MsgBox("Commit changes?", vbYesNo + vbQuestion + vbDefaultButton1, "Asking...") = vbYes) Then
            
            AddSave = enmAcSaveFailure  ' Initially assuming that successfull
                                        ' updation would not take place
            
            If (doCheck) Then If (Not callerForm.chk) Then Exit Function
            
            callerForm.letRec
            RecordSet.Update
            cmdAdd.Caption = "&Add"
            cmdEdit.Caption = "&Edit"
            callerForm.makeReadOnly
            
            AddSave = enmAcSaveSuccess  ' Flaging as Successfull updation
                                        ' Committed
            
        End If
    End If
    
    Exit Function

ErrHandle:
    Select Case Err.Number
        Case -2147467259
            Dim msgStr As String
            msgStr = "Probably! this record already exists"
'            If (gHasVoiceAid = True) Then Speek (msgStr)
            MsgBox msgStr, vbOKOnly + vbCritical, "Fault..."
        Case Else
            MsgBox "Unexpected Error raised." & vbCrLf & _
                   "Error number: " & Err.Number & vbCrLf & _
                   "Messgae: " & Err.Description & vbCrLf & _
                   "Source: " & Err.Source
    End Select
End Function

Public Function EditRevert(ByRef callerForm As Form, ByRef RecordSet As ADODB.RecordSet, ByRef cmdAdd As CommandButton, ByRef cmdEdit As CommandButton) As enmActionType
    
    If (Trim(UCase(cmdEdit.Caption)) = "&EDIT") Then
        If (RecordSet.BOF Or RecordSet.EOF) Then
            MsgBox "Sorry! You can't modify this record", vbOKOnly + vbInformation + vbDefaultButton1, "Info..."
            EditRevert = enmAcEditFailure
        Else
            cmdAdd.Caption = "&Commit"
            cmdEdit.Caption = "&Revert"
            callerForm.makeEditable
            EditRevert = enmAcEditSuccess
        End If
    Else
        RecordSet.CancelUpdate
        callerForm.getRec
        cmdAdd.Caption = "&Add"
        cmdEdit.Caption = "&Edit"
        callerForm.makeReadOnly
        EditRevert = enmAcRevert
    End If
        
End Function

Public Function Delete(ByRef callerForm As Form, ByRef RecordSet As ADODB.RecordSet) As enmActionType
    On Error GoTo Err_Handle
    Err.Clear
    
    Delete = enmAcDeleteFailure
    
    If (RecordSet.BOF Or RecordSet.EOF) Then
        MsgBox "Sorry! You can't delete this record.", vbOKOnly + vbInformation + vbDefaultButton1, "Info..."
        Exit Function
    End If
    
    If Not (IsNull(RecordSet!ISARCHIVED)) Then
        If (RecordSet!ISARCHIVED) Then
            MsgBox "Sorry! You can't delete an archived record.", vbOKOnly + vbInformation + vbDefaultButton1, "Info..."
            Exit Function
        End If
    End If

SKIP_ARCHIVE:
    
    If (MsgBox("You are going to delete the record." & vbCrLf & vbCrLf & "Are you sure?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm...") = vbNo) Then Exit Function
    
    RecordSet.Delete
    Delete = enmAcDeleteSuccess
    Me.movePrevious callerForm, RecordSet
    
    Exit Function

Err_Handle:
    Select Case Err.Number
        Case 3265   'Attribute does not found in the collection(ISARCHIVED).
            GoTo SKIP_ARCHIVE
        Case Else
            Me.showError Err
    End Select
End Function

Public Function moveNext(ByRef callerForm As Form, ByRef RecordSet As ADODB.RecordSet) As Boolean
    moveNext = False
    If (Not RecordSet.EOF) Then
        moveNext = True
    
        RecordSet.moveNext
    End If
    
    If (RecordSet.EOF) Then
        Me.moveLast callerForm, RecordSet
    Else
        callerForm.getRec
    End If
End Function

Public Function movePrevious(ByRef callerForm As Form, ByRef RecordSet As ADODB.RecordSet) As Boolean
    movePrevious = False
    If (Not RecordSet.BOF) Then
        movePrevious = True
        RecordSet.movePrevious
    End If
    
    If (RecordSet.BOF) Then
        Me.moveFirst callerForm, RecordSet
    Else
        callerForm.getRec
    End If
End Function

Public Function moveFirst(ByRef callerForm As Form, ByRef RecordSet As ADODB.RecordSet) As Boolean
    moveFirst = False
    If (RecordSet.BOF And RecordSet.EOF) Then
        callerForm.clearUI
        Exit Function
    End If
    moveFirst = True
    
    RecordSet.moveFirst
    callerForm.getRec
End Function

Public Function moveLast(ByRef callerForm As Form, ByRef RecordSet As ADODB.RecordSet) As Boolean
    moveLast = False
    If (RecordSet.BOF And RecordSet.EOF) Then
        callerForm.clearUI
        Exit Function
    End If
    moveLast = True
    
    RecordSet.moveLast
    callerForm.getRec
End Function

Public Function keyTraversal(KeyCode As Integer, Shift As Integer, ByRef callerForm As Form, ByRef RecordSet As ADODB.RecordSet) As Boolean
    keyTraversal = True
    Select Case KeyCode
        Case 33 ' PageUP
            Me.movePrevious callerForm, RecordSet
        Case 34 ' PageDown
            Me.moveNext callerForm, RecordSet
        Case 35 ' End
            Me.moveLast callerForm, RecordSet
        Case 36 ' Home
            Me.moveFirst callerForm, RecordSet
        Case Else
            keyTraversal = False
    End Select
End Function

Public Function keyManipulation(KeyCode As Integer, Shift As Integer, ByRef callerForm As Form, ByRef RecordSet As ADODB.RecordSet, ByRef cmdAdd As CommandButton, ByRef cmdEdit As CommandButton, Optional ByVal doCheck As Boolean = True) As enmActionType
    'MsgBox KeyCode & " : " & Shift
    Select Case KeyCode
        Case Is = 45 And Shift = 0 ' Insert
            If (UCase(cmdAdd.Caption) = "&ADD") Then
                keyManipulation = Me.AddSave(callerForm, RecordSet, cmdAdd, cmdEdit, doCheck)
            End If
        Case Is = 45 And Shift = 1 ' Shift + Insert
            If (UCase(cmdEdit.Caption) = "&EDIT") Then
                keyManipulation = Me.EditRevert(callerForm, RecordSet, cmdAdd, cmdEdit)
            End If
        Case Is = 123 And Shift = 0 ' F12
            If (UCase(cmdAdd.Caption) = "&COMMIT") Then
                keyManipulation = Me.AddSave(callerForm, RecordSet, cmdAdd, cmdEdit, doCheck)
            End If
        Case Is = 46 And Shift = 0 ' Delete
            If (UCase(cmdAdd.Caption) = "&ADD") Then
                keyManipulation = Me.Delete(callerForm, RecordSet)
            End If
        Case Else
            keyManipulation = enmAcNone
    End Select
    
End Function

Public Sub showError(error As ErrObject)

    MsgBox "I encounter following UnExpected error." & vbCrLf & vbCrLf & _
            "Number: " & error.Number & vbCrLf & _
            "Description: " & error.Description & vbCrLf & _
            "Source: " & error.Source & vbCrLf & vbCrLf & _
            "Please! contact programmer soon.", vbCritical + vbOKOnly, "UnExpected error..."

End Sub

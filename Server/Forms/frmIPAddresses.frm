VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmIPAddresses 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IP Addresses:"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid dbgIPAddress 
      Height          =   3885
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   6853
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbAction 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   635
      ButtonWidth     =   1746
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Look-Up"
            Key             =   "keyLookUp"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Key             =   "keyDelete"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "keyExit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmIPAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dbgIPAddress_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    On Error Resume Next
'
'    Call orsIPAddresses.moveFirst
'    Me.dbgIPAddress.Col = 1
'    Call orsIPAddresses.Find("IPAddress = '" & _
'                                Me.dbgIPAddress.text & _
'                                "' ", , adSearchForward)
End Sub

Private Sub Form_Load()
    Call Form_Resize
    
    With Me.tlbAction
        Set .ImageList = frmMDI.imgList16
        .Buttons("keyLookUp").Image = frmMDI.imgList16.ListImages("imgFind").Index
        .Buttons("keyDelete").Image = frmMDI.imgList16.ListImages("imgDelete").Index
        .Buttons("keyExit").Image = frmMDI.imgList16.ListImages("imgClose").Index
    End With
    
    
    If (orsIPAddresses.State <> adStateClosed) Then orsIPAddresses.Close
    
    'orsIPAddresses.Requery
    orsIPAddresses.Open
    
    If (Not orsIPAddresses.BOF) Then Call orsIPAddresses.moveFirst
    Set Me.dbgIPAddress.DataSource = orsIPAddresses
    
    
    
End Sub

Private Sub Form_Resize()
    With Me.dbgIPAddress
        .left = Me.ScaleLeft
        .tOp = Me.ScaleTop + Me.tlbAction.Height
        .Height = Me.ScaleHeight - Me.tlbAction.Height
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    If (orsIPAddresses.State = adStateOpen) Then
        orsIPAddresses.Update
    End If
    
End Sub

Private Sub tlbAction_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    Select Case (Button.key)
        Case "keyLookUp"
            If (orsIPAddresses.BOF And orsIPAddresses.EOF) Then
                mdlMsg.showError ("Nothing to be looked in.")
                Exit Sub
            Else
                Dim vIPAddress As String
                vIPAddress = InputBox("Enter IP Address to be looked up.", _
                                        "Look-Up", vbNullString)
                
                orsIPAddresses.moveFirst
                If (vIPAddress <> vbNullString) Then
                    Call orsIPAddresses.Find("IPAddress = '" & vIPAddress & "' ", , adSearchForward)
                End If
            End If
        Case "keyDelete"
            If (orsIPAddresses.BOF Or orsIPAddresses.EOF) Then
                mdlMsg.showError ("Can not delete this record.")
                Exit Sub
            Else
                If (mdlMsg.showConfirm(orsIPAddresses!IPAddress & vbCrLf & _
                    "Wish to delete?") = vbYes) Then
                    Call orsIPAddresses.Delete(adAffectCurrent)
                    Me.dbgIPAddress.Refresh
                End If
            End If
        Case "keyExit"
            Unload Me
    End Select

End Sub

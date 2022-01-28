VERSION 5.00
Begin VB.Form frmNotify 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Notification:"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBtn 
      Caption         =   "&Send"
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   4
      Top             =   3060
      Width           =   1410
   End
   Begin VB.CommandButton cmdBtn 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Index           =   0
      Left            =   4680
      TabIndex        =   5
      Top             =   3060
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Height          =   2940
      Left            =   90
      TabIndex        =   6
      Top             =   0
      Width           =   6045
      Begin VB.TextBox txtTitle 
         Height          =   330
         Left            =   1080
         TabIndex        =   2
         Top             =   1035
         Width           =   4875
      End
      Begin VB.TextBox txtMessage 
         Height          =   1410
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1440
         Width           =   4875
      End
      Begin VB.ComboBox cmbNotificationType 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   630
         Width           =   2265
      End
      Begin VB.ComboBox cmbClients 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   225
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ttitle:"
         Height          =   195
         Index           =   3
         Left            =   630
         TabIndex        =   10
         Top             =   1035
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Text:"
         Height          =   195
         Index           =   2
         Left            =   675
         TabIndex        =   9
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   8
         Top             =   630
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Destination:"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   225
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cMessageLength As Byte = 200
Private Const cTitleLength As Byte = 50

Private Sub cmdBtn_Click(Index As Integer)
    Select Case (Index)
        Case 0 'Close
            Unload Me
        Case 1 'Send
            If (Me.cmbClients.ListIndex = -1) Then
                Call mdlMsg.showAlert("Select a valid client.")
                Call Me.cmbClients.SetFocus
            ElseIf (Me.cmbNotificationType.ListIndex = -1) Then
                Call mdlMsg.showAlert("Select a notification type.")
                Call Me.cmbNotificationType.SetFocus
            ElseIf (Len(Trim(Me.txtMessage.text)) = 0) Then
                Call mdlMsg.showAlert("Hey buddy! why u want to send null message?")
                Call Me.txtMessage.SetFocus
            Else
                
                '******* Message Format
                '   Message text
                '   Message Title
                '   Message Icon
                '       0 = No icon
                '       1 = Information
                '       2 = Warning
                '       3 = Error
                '**********************
                Dim vMessage As String
                
                vMessage = Me.txtMessage.text
                vMessage = vMessage & cClauseDelimiter & Me.txtTitle.text
                vMessage = vMessage & cClauseDelimiter & Me.cmbNotificationType.ListIndex
                
                Call colClients(Me.cmbClients.ListIndex + 1).sendSignal(enmSigNotify, vMessage)
                Call mdlMsg.showMessage(colClients(Me.cmbClients.ListIndex + 1).Alias & " has been notifed.")
                                
                Unload Me
                
            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim oClient As New clsClient
    
    Me.cmbClients.Clear
    For Each oClient In colClients
        Me.cmbClients.AddItem (oClient.Alias)
    Next oClient
    Me.cmbClients.ListIndex = 0
    
    With Me.cmbNotificationType
        .Clear
        .AddItem ("Normal")
        .AddItem ("Information")
        .AddItem ("Warning")
        .AddItem ("Error")
        .ListIndex = 0
    End With

    With Me.txtTitle
        .text = vbNullString
        .MaxLength = cTitleLength
    End With

    With Me.txtMessage
        .text = vbNullString
        .MaxLength = cMessageLength
    End With

End Sub

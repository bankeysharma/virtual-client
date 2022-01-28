VERSION 5.00
Begin VB.Form frmAuthentication 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7620
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBottom 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   360
      ScaleHeight     =   1050
      ScaleWidth      =   5235
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4725
      Width           =   5235
      Begin VB.Label lblMessage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   765
         Width           =   480
      End
      Begin VB.Label lblQuit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   315
         Left            =   4005
         TabIndex        =   6
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.PictureBox picHeader 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   0
      ScaleHeight     =   870
      ScaleWidth      =   5685
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   5685
      Begin VB.Image imgHeader 
         Height          =   735
         Left            =   0
         Picture         =   "frmAuthentication.frx":0000
         Top             =   0
         Width           =   4125
      End
   End
   Begin VB.PictureBox picBase 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   45
      ScaleHeight     =   3615
      ScaleWidth      =   7080
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   990
      Width           =   7080
      Begin VB.PictureBox picInputbase 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   1770
         Left            =   1035
         ScaleHeight     =   1740
         ScaleWidth      =   5205
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   990
         Width           =   5235
         Begin VB.TextBox txtInput 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   180
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   855
            Width           =   4830
         End
         Begin VB.Label lblHeading 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   4
            Top             =   495
            Width           =   675
         End
      End
      Begin VB.Label lblWelcomeGreeting 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   480
         Left            =   2160
         TabIndex        =   8
         Top             =   135
         Width           =   135
      End
      Begin VB.Image imgBase 
         Height          =   285
         Left            =   225
         Picture         =   "frmAuthentication.frx":138B
         Stretch         =   -1  'True
         Top             =   180
         Visible         =   0   'False
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmAuthentication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cBottomHeight As Long = 750

Private vWait As Boolean

Private Sub Form_Resize()

    On Error Resume Next
    
    With Me.picHeader
        
        .Left = Me.ScaleLeft
        .Top = Me.ScaleTop
        .Width = Me.ScaleWidth
         
        .BackColor = vbWhite
    
    End With
    
    With Me.imgHeader
        .Left = Me.picHeader.Width - .Width 'Me.picHeader.ScaleLeft
        .Top = Me.picHeader.ScaleTop
        
        Me.picHeader.Height = .Height
    End With

    With Me.picBottom
        .Top = Me.ScaleHeight - cBottomHeight
        .Left = Me.ScaleLeft
        .Height = cBottomHeight
        .Width = Me.ScaleWidth
    End With
    
    With Me.lblQuit
        .Left = Me.picBottom.ScaleWidth - .Width - 100
        .Top = Me.picBottom.ScaleHeight - .Height - 100
    End With
    
    With Me.lblMessage
        .Left = Me.picBottom.ScaleLeft + 20
        .Top = Me.picBottom.ScaleHeight - .Height - 20
    End With
    
    With Me.picBase
        
        .Top = Me.picHeader.Top + Me.picHeader.Height
        .Left = Me.ScaleLeft
        .Height = Me.ScaleHeight - Me.picHeader.Height - cBottomHeight
        .Width = Me.ScaleWidth
    
    End With

    With Me.imgBase
        
        .Top = Me.picBase.ScaleTop
        .Left = Me.picBase.ScaleLeft
        .Height = Me.picBase.ScaleHeight
        .Width = Me.picBase.ScaleWidth
        
    End With

    With Me.picInputbase
        
        .Top = (Me.picBase.ScaleHeight - Me.picInputbase.Height) / 2
        .Left = (Me.picBase.ScaleWidth - Me.picInputbase.Width) / 2
    
    End With

    With Me.lblWelcomeGreeting
        '.Caption = "WelCome" & vbCrLf & _
                    "to" & vbCrLf & _
                    "Virtual Client"
        .Caption = "W e l C o m e" & vbCrLf & _
                    "t o" & vbCrLf & _
                    "V i r t u a l  C l i e n t"
        '.Caption = "WelCome to Virtual Client"
        '.Caption = "W e l C o m e  t o  V i r t u a l  C l i e n t"
        
        .Top = (Me.picInputbase.Top - Me.lblWelcomeGreeting.Height) / 2 'Me.lblWelcomeGreeting.Height
        .Left = Me.picBase.ScaleLeft
        .Width = Me.picBase.ScaleWidth
    End With

End Sub

Public Function DoAuthentication() As Boolean
    
    Load Me
            
    Call Me.reSetDimensions
            
    Dim vUser As String
    Dim vPassword As String
    Dim vDomain As String
    
    Me.lblMessage.Caption = "* Enter NT authenticated IDs only, here."
    
    '************* User Name
    vWait = True
    Me.lblHeading.Caption = "User Name"
    With Me.txtInput
        .PasswordChar = vbNullString
        .text = vbNullString
        .SetFocus
    End With
    
    Do While (vWait)
        DoEvents
    Loop
    
    vUser = Me.txtInput.text
    '****************************************
    
    '************* Password
    vWait = True
    Me.lblHeading.Caption = "Password"
    With Me.txtInput
        .PasswordChar = "*"
        .text = vbNullString
        .SetFocus
    End With

    Do While (vWait)
        DoEvents
    Loop
    
    vPassword = Me.txtInput.text
    '*****************************************
    
    '************* Domain Name
    vWait = True
    Me.lblHeading.Caption = "Domain"
    With Me.txtInput
        .PasswordChar = vbNullString
        .text = vbNullString
        .SetFocus
    End With

    Do While (vWait)
        DoEvents
    Loop
    
    vDomain = Me.txtInput.text
    '*****************************************
    
    DoAuthentication = AuthenticateUser(vDomain, vUser, vPassword)

End Function

Private Sub lblQuit_Click()
    Call Quit
End Sub

Private Sub lblQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.lblQuit.ForeColor = &H80FFFF

End Sub

Private Sub picBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.lblQuit.ForeColor = &HC0C0&
    
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        vWait = False
        KeyAscii = 0
    End If
End Sub

Public Sub reSetDimensions()
    With Me
        .Left = 0
        .Top = 0
        .Height = frmMDI.ScaleHeight
        .Width = frmMDI.ScaleWidth
        .Show
    End With
End Sub

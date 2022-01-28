VERSION 5.00
Begin VB.Form frmDesktop 
   Caption         =   "Form1"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4200
   Visible         =   0   'False
   Begin VB.Frame frameDesktop 
      Height          =   2310
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   4200
      Begin VB.PictureBox picDesktop 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   2130
         Left            =   45
         ScaleHeight     =   2070
         ScaleWidth      =   4050
         TabIndex        =   1
         Top             =   135
         Width           =   4110
         Begin VB.Image imgDesktop 
            Height          =   1770
            Left            =   135
            Stretch         =   -1  'True
            Top             =   135
            Width           =   3750
         End
      End
   End
End
Attribute VB_Name = "frmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cframeBottomMargin As Byte = 15 '495
Private Const cframeRightMargin As Byte = 12 '270

Private Const cpicBottomMargin As Byte = 180
Private Const cpicRightMargin As Byte = 90

Private Sub Form_Resize()
    On Error GoTo Err_Handle
    Err.Clear
    
    '************* Desktop
    Me.frameDesktop.Height = Me.ScaleHeight - (cframeBottomMargin + Me.frameDesktop.tOp)
    Me.frameDesktop.Width = Me.ScaleWidth - cframeRightMargin
    
    Me.picDesktop.Height = Me.frameDesktop.Height - cpicBottomMargin
    Me.picDesktop.Width = Me.frameDesktop.Width - cpicRightMargin
    
    '******* Image
    Dim vHeight As Long
    Dim vWidth As Long
    
    vWidth = Me.picDesktop.ScaleWidth
    vHeight = vWidth * cHeightRatio / 100
    
    If (vHeight > Me.picDesktop.ScaleHeight) Then
        vHeight = Me.picDesktop.ScaleHeight
        vWidth = vHeight * 100 / cHeightRatio
    End If
        
    Me.imgDesktop.Height = vHeight
    Me.imgDesktop.Width = vWidth
    
    Me.imgDesktop.left = (Me.picDesktop.ScaleWidth - Me.imgDesktop.Width) / 2
    Me.imgDesktop.tOp = (Me.picDesktop.ScaleHeight - Me.imgDesktop.Height) / 2
    '*******
    
    '**************************
    
    Exit Sub
    
Err_Handle:
    
    Select Case (Err.Number)
        Case 380
            Exit Sub
        Case Else
            Call oDB.showError(Err)
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If (Not colClients(Me.Tag) Is Nothing) Then _
        colClients(Me.Tag).Off
End Sub

Private Sub imgDesktop_DblClick()
    picDesktop_DblClick
End Sub

Private Sub imgDesktop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picDesktop_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub picDesktop_DblClick()
    CurrentService = srvDesktopCapture
    Call solitaryView
    Call activeClient.showDesktop
End Sub

Private Sub picDesktop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.SetFocus
    Set activeClient = colClients(Me.Tag)
    If (Button = vbRightButton) Then
        frmMDI.showDesktopPopup
    End If
End Sub

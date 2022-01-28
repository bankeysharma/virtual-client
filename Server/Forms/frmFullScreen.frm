VERSION 5.00
Begin VB.Form frmFullScreen 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   4680
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picDesktop 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2130
      Left            =   0
      ScaleHeight     =   2130
      ScaleWidth      =   4110
      TabIndex        =   0
      Top             =   0
      Width           =   4110
      Begin VB.Image imgDesktop 
         Height          =   1770
         Left            =   135
         Stretch         =   -1  'True
         Top             =   135
         Width           =   3750
      End
   End
   Begin VB.Menu mnuViews 
      Caption         =   "Views"
      Begin VB.Menu mnuViewsBroadSpectrum 
         Caption         =   "Broad-Spectrum"
      End
      Begin VB.Menu mnuViewSolitary 
         Caption         =   "Solitary View"
      End
      Begin VB.Menu mnuViewRandom 
         Caption         =   "Random View"
      End
   End
End
Attribute VB_Name = "frmFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyF5
            Call activeClient.sendSignal(enmSigSendDesktop)
    End Select
End Sub

Private Sub Form_Load()
    Me.mnuViews.Visible = False
    
    'Call activeClient.showDesktop
    'Call Form_KeyUp(vbKeyF5, 0)
End Sub

Private Sub Form_Resize()
    On Error GoTo Err_Handle
    Err.Clear
    
    '************* Desktop
    Me.picDesktop.Height = Me.ScaleHeight
    Me.picDesktop.Width = Me.ScaleWidth
    
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

Private Sub imgDesktop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbRightButton) Then
        Call Me.PopupMenu(Me.mnuViews)
    End If
End Sub

Private Sub mnuViewRandom_Click()
    Call mdlViews.randomView
End Sub

Private Sub mnuViewsBroadSpectrum_Click()
    Call mdlViews.broadSpectrumView
End Sub

Private Sub mnuViewSolitary_Click()
    Call mdlViews.solitaryView
End Sub

VERSION 5.00
Begin VB.Form frmShowAssociators 
   Caption         =   "Associators:"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstInfo 
      Height          =   3180
      Left            =   135
      TabIndex        =   0
      Top             =   720
      Width           =   4965
   End
   Begin VB.Label lblHead2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   45
      TabIndex        =   2
      Top             =   360
      Width           =   585
   End
   Begin VB.Label lblHead1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   690
   End
End
Attribute VB_Name = "frmShowAssociators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call Form_Resize
End Sub

Private Sub Form_Resize()
    Me.lstInfo.Height = Me.ScaleHeight - Me.lstInfo.tOp
    Me.lstInfo.Width = Me.ScaleWidth
    Me.lstInfo.left = Me.ScaleLeft
End Sub

Public Sub showInfo(pInfo As Variant)
    
    Me.lstInfo.Clear
    
    Dim i As Integer
        
    Me.lblHead1 = pInfo(0)
    Me.lblHead2 = "Path: " & pInfo(1)
    
    For i = LBound(pInfo) + 2 To UBound(pInfo) - 1
        Me.lstInfo.AddItem (pInfo(i))
    Next i
    
    Call Me.show(vbModal)
    
End Sub

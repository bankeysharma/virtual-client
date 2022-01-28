VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form frmServices 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   5820
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab tabServices 
      Height          =   2805
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   4948
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Desktop"
      TabPicture(0)   =   "frmServices.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblDesktopSubstitute"
      Tab(0).Control(1)=   "frameDesktop"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Processes"
      TabPicture(1)   =   "frmServices.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameProcesses"
      Tab(1).Control(1)=   "lblProcessesSubstitute"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Explorer"
      TabPicture(2)   =   "frmServices.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblExplorerSubstitute"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cbXplorer"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.Frame frameProcesses 
         Height          =   2130
         Left            =   -74865
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   5055
         Begin MSComctlLib.ListView lstvProcesses 
            Height          =   1905
            Left            =   45
            TabIndex        =   18
            Top             =   135
            Width           =   4875
            _ExtentX        =   8599
            _ExtentY        =   3360
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "imgList16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Process Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Priority"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Memory Usage"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame frameDesktop 
         Height          =   2310
         Left            =   -74865
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
         Width           =   5370
         Begin VB.PictureBox picDesktop 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            Height          =   2130
            Left            =   45
            ScaleHeight     =   2070
            ScaleWidth      =   5220
            TabIndex        =   2
            Top             =   135
            Width           =   5280
            Begin VB.Image imgDesktop 
               Height          =   1635
               Left            =   90
               Stretch         =   -1  'True
               Top             =   90
               Width           =   2400
            End
         End
      End
      Begin ComCtl3.CoolBar cbXplorer 
         Height          =   1965
         Left            =   135
         TabIndex        =   4
         Top             =   450
         Visible         =   0   'False
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   3466
         BandCount       =   2
         FixedOrder      =   -1  'True
         BandBorders     =   0   'False
         VariantHeight   =   0   'False
         _CBWidth        =   5355
         _CBHeight       =   1965
         _Version        =   "6.0.8169"
         BandBackColor1  =   -2147483638
         Child1          =   "trvDirectories"
         MinWidth1       =   15
         MinHeight1      =   1905
         Width1          =   2775
         UseCoolbarColors1=   0   'False
         Key1            =   "keyTreeView"
         NewRow1         =   0   'False
         BandBackColor2  =   -2147483638
         Child2          =   "picExplorereListView"
         MinHeight2      =   915
         Width2          =   465
         UseCoolbarColors2=   0   'False
         Key2            =   "keyListView"
         NewRow2         =   0   'False
         Begin VB.PictureBox picExplorereListView 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   915
            Left            =   2940
            ScaleHeight     =   885
            ScaleWidth      =   2295
            TabIndex        =   6
            Top             =   525
            Width           =   2325
            Begin MSComctlLib.ListView lstvXplorer 
               Height          =   1275
               Left            =   1710
               TabIndex        =   14
               Top             =   45
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   2249
               Arrange         =   2
               LabelEdit       =   1
               MultiSelect     =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               Icons           =   "imgList32"
               SmallIcons      =   "imgList16"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
            Begin VB.PictureBox picObjectInfo 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   1635
               Left            =   0
               ScaleHeight     =   1605
               ScaleWidth      =   2280
               TabIndex        =   7
               Top             =   45
               Width           =   2310
               Begin MSChart20Lib.MSChart chartHDD 
                  Height          =   1005
                  Left            =   135
                  OleObjectBlob   =   "frmServices.frx":0054
                  TabIndex        =   8
                  Top             =   855
                  Width           =   690
               End
               Begin VB.Label lblDescription 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Description Value"
                  Height          =   195
                  Left            =   1350
                  TabIndex        =   13
                  Top             =   450
                  Visible         =   0   'False
                  Width           =   1245
               End
               Begin VB.Label lblDescriptionHead 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Description"
                  Height          =   195
                  Left            =   360
                  TabIndex        =   12
                  Top             =   675
                  Visible         =   0   'False
                  Width           =   795
               End
               Begin VB.Label lblObjectType 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Object Type"
                  Height          =   240
                  Left            =   315
                  TabIndex        =   11
                  Top             =   405
                  Visible         =   0   'False
                  Width           =   1320
               End
               Begin VB.Label lblObjectName 
                  BackStyle       =   0  'Transparent
                  Caption         =   "ObjectName"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   270
                  TabIndex        =   10
                  Top             =   225
                  UseMnemonic     =   0   'False
                  Visible         =   0   'False
                  Width           =   1230
                  WordWrap        =   -1  'True
               End
               Begin VB.Label lblParentName 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Parent Name"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   270
                  TabIndex        =   9
                  Top             =   -45
                  Visible         =   0   'False
                  Width           =   2085
                  WordWrap        =   -1  'True
               End
               Begin VB.Image imgObjectIcon 
                  Height          =   375
                  Left            =   90
                  Top             =   135
                  Visible         =   0   'False
                  Width           =   465
               End
            End
         End
         Begin MSComctlLib.TreeView trvDirectories 
            Height          =   1905
            Left            =   30
            TabIndex        =   5
            Top             =   30
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   3360
            _Version        =   393217
            Indentation     =   88
            LabelEdit       =   1
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label lblProcessesSubstitute 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FFC0C0&
         Height          =   270
         Left            =   -71985
         MouseIcon       =   "frmServices.frx":1D16
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   2475
         Width           =   690
      End
      Begin VB.Label lblDesktopSubstitute 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FFC0C0&
         Height          =   270
         Left            =   -72210
         MouseIcon       =   "frmServices.frx":2020
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   2520
         Width           =   690
      End
      Begin VB.Label lblExplorerSubstitute 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FFC0C0&
         Height          =   270
         Left            =   2565
         MouseIcon       =   "frmServices.frx":232A
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   2340
         Width           =   690
      End
   End
End
Attribute VB_Name = "frmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cInfoBarWidth As Integer = 2800
Private Const cIsFile As Byte = 2

Private oActiveListItem As MSComctlLib.ListItem

Private Const cframeBottomMargin As Integer = 495
Private Const cframeRightMargin As Integer = 270

Private Const cpicBottomMargin As Byte = 180
Private Const cpicRightMargin As Byte = 90

Private Const clstvRightMargin As Byte = 95
Private Const clstvBottomMargin As Byte = 185

Public oProcesses As New clsProcesses

Private Sub Form_Load()
    
    With oProcesses
        
        Set .listingObject = Me.lstvProcesses
        Set .containerObject = Me.frameProcesses
        Set .Client = activeClient
        
    End With
    
    Me.picObjectInfo.BorderStyle = 0
    Me.lstvXplorer.BorderStyle = 0
    Me.lblParentName.Caption = vbNullString
    
    Me.lblExplorerSubstitute.Caption = " Click on me " & vbCrLf & _
                                            "to" & vbCrLf & _
                                            "Build Xplorer"
    
    Me.lblDesktopSubstitute.Caption = " Click on me " & vbCrLf & _
                                            "to" & vbCrLf & _
                                            "Fetch Desktop"
    
    Me.lblProcessesSubstitute.Caption = " Click on me " & vbCrLf & _
                                            "to" & vbCrLf & _
                                            "Enumerate Processes"
    
    Call hideParentInfo
    Call hideItemInfo
    
End Sub

Private Sub Form_Resize()
    
    On Error GoTo Err_Handle
    Err.Clear
    
    Me.tabServices.Width = Me.ScaleWidth - (Me.tabServices.left * 2)
    Me.tabServices.Height = Me.ScaleHeight - (Me.tabServices.tOp * 2)
    
    '************* Desktop
    Me.frameDesktop.Height = Me.tabServices.Height - cframeBottomMargin
    Me.frameDesktop.Width = Me.tabServices.Width - cframeRightMargin
    
    With Me.frameDesktop
        
        Me.picDesktop.Height = .Height - cpicBottomMargin
        Me.picDesktop.Width = .Width - cpicRightMargin
        
        Me.lblDesktopSubstitute.tOp = .tOp + (.Height - Me.lblDesktopSubstitute.Height) / 2
        Me.lblDesktopSubstitute.left = .left + (.Width - Me.lblDesktopSubstitute.Width) / 2
        
    End With
    
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
    
    '************* Processes
    With Me.frameProcesses
        
        .Height = Me.frameDesktop.Height
        .Width = Me.frameDesktop.Width
    
        Me.lblProcessesSubstitute.tOp = .tOp + (.Height - Me.lblProcessesSubstitute.Height) / 2
        Me.lblProcessesSubstitute.left = .left + (.Width - Me.lblProcessesSubstitute.Width) / 2
                
        'Me.lstvProcesses.left = 0
        'Me.lstvProcesses.tOp = 0
        Me.lstvProcesses.Height = .Height - clstvBottomMargin
        Me.lstvProcesses.Width = .Width - clstvRightMargin
        
    End With
    '***************************
    
    '************** Explorer
    
    With Me.cbXplorer
        
        .Width = Me.tabServices.Width - cframeRightMargin
        
        .Bands("keyTreeView").MinHeight = Me.tabServices.Height - cframeBottomMargin - 150
        .Bands("keyListView").MinHeight = Me.tabServices.Height - cframeBottomMargin - 150
        
        .Bands("keyTreeView").Width = Me.cbXplorer.Width / 4
        .Bands("keyListView").Width = Me.cbXplorer.Width - .Bands("keyTreeView").Width
        
        .Bands("keyTreeView").Style = cc3BandNormal
        .Bands("keyListView").Style = cc3BandNormal
        
        Me.lblExplorerSubstitute.tOp = Me.frameDesktop.tOp + (Me.frameDesktop.Height - Me.lblExplorerSubstitute.Height) / 2
        Me.lblExplorerSubstitute.left = .left + (.Width - Me.lblExplorerSubstitute.Width) / 2
    
    End With

Exit_Label:

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
    Call offSolitaryView
End Sub

Private Sub lblDesktopSubstitute_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblDesktopSubstitute.ForeColor = vbBlue
End Sub

Private Sub lblDesktopSubstitute_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then _
        Call activeClient.sendSignal(enmSigSendDesktop)
End Sub

Private Sub lblExplorerSubstitute_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblExplorerSubstitute.ForeColor = vbBlue
End Sub

Private Sub lblExplorerSubstitute_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then _
        Call activeClient.sendSignal(enmSigSendFileSystem, cClauseDelimiter & "0")
End Sub

Private Sub lblProcessesSubstitute_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblProcessesSubstitute.ForeColor = vbBlue
End Sub

Private Sub lblProcessesSubstitute_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbLeftButton) Then
        Call oProcesses.enumerateProcesses
    End If
End Sub

Private Sub lstvProcesses_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbRightButton) Then
        Call Me.PopupMenu(frmMDI.mnuPadProcess)
    End If
End Sub

Private Sub lstvXplorer_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyReturn And Shift = 0) Then
        Call lstvXplorer_DblClick
    End If
End Sub

Private Sub tabServices_Click(PreviousTab As Integer)
    Call tabServices_DblClick
End Sub

Private Sub tabServices_DblClick()
    
    If (Me.tabServices.Tab = 0) Then
        CurrentService = srvDesktopCapture
        'frmMDI.tlbServices.Buttons("keyDesktop").Value = tbrPressed
    ElseIf (Me.tabServices.Tab = 1) Then
        CurrentService = srvProcessMonitoring
        'frmMDI.tlbServices.Buttons("keyProcesses").Value = tbrPressed
    ElseIf (Me.tabServices.Tab = 2) Then
        CurrentService = srvExplorer
        'frmMDI.tlbServices.Buttons("keyExplorer").Value = tbrPressed
    End If
    
    'Call solitaryView
    
End Sub

Private Sub lstvXplorer_DblClick()
    If (oActiveListItem Is Nothing) Then Exit Sub
    
    If (oActiveListItem.Tag = cIsFile) Then Exit Sub
    
    Call trvDirectories_NodeClick(Me.trvDirectories.Nodes(oActiveListItem.key))
    
End Sub

Private Sub lstvXplorer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set oActiveListItem = Item
    Call showItemInfo(Item)
End Sub

Private Sub picExplorereListView_Resize()
    On Error Resume Next
    
    With Me.picObjectInfo
        .left = Me.picExplorereListView.ScaleLeft
        .tOp = Me.picExplorereListView.ScaleTop
        .Height = Me.picExplorereListView.ScaleHeight
        .Width = cInfoBarWidth 'Me.picExplorereListView.ScaleWidth / 3.5
    End With
    
    With Me.lstvXplorer
        .left = Me.picObjectInfo.Width
        .tOp = Me.picExplorereListView.ScaleTop
        .Height = Me.picExplorereListView.ScaleHeight
        .Width = Me.picExplorereListView.ScaleWidth - Me.picObjectInfo.Width
    End With
End Sub

Private Sub picObjectInfo_Resize()

    On Error Resume Next
    
    Const cIconLeftMargin As Integer = 50
    Const cIconTopMargin As Integer = 200
    
    Me.imgObjectIcon.left = cIconLeftMargin
    Me.imgObjectIcon.tOp = cIconTopMargin
    
    Me.lblParentName.left = cIconLeftMargin
    Me.lblParentName.Width = Me.ScaleWidth - cIconLeftMargin
    
    Me.lblObjectName.left = cIconLeftMargin
    Me.lblObjectName.Width = Me.ScaleWidth - cIconLeftMargin
    
    Me.lblObjectType.left = cIconLeftMargin
    Me.lblObjectType.Width = Me.ScaleWidth - cIconLeftMargin
    
    Me.lblDescription.left = cIconLeftMargin
    Me.lblDescription.Width = Me.ScaleWidth - cIconLeftMargin
    
    Me.lblDescriptionHead.left = cIconLeftMargin
    Me.lblDescriptionHead.Width = Me.ScaleWidth - cIconLeftMargin
    
End Sub

Private Sub tabServices_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Me.tabServices.Tab = 2) Then
        Me.lblExplorerSubstitute.ForeColor = &HFFC0C0
    ElseIf (Me.tabServices.Tab = 1) Then
        Me.lblProcessesSubstitute.ForeColor = &HFFC0C0
    ElseIf (Me.tabServices.Tab = 0) Then
        Me.lblDesktopSubstitute.ForeColor = &HFFC0C0
    End If
End Sub

Private Sub trvDirectories_Collapse(ByVal Node As MSComctlLib.Node)
    If (Node.Image = frmMDI.imgList16.ListImages("imgFolderOpen1").Index) Then _
        Node.Image = frmMDI.imgList16.ListImages("imgFolderClose1").Index
End Sub

Private Sub trvDirectories_Expand(ByVal Node As MSComctlLib.Node)
    If (Node.Image = frmMDI.imgList16.ListImages("imgFolderClose1").Index) Then _
        Node.Image = frmMDI.imgList16.ListImages("imgFolderOpen1").Index
End Sub

Private Sub trvDirectories_NodeClick(ByVal Node As MSComctlLib.Node)
    
    If (Node.Child Is Nothing) Then
        
        Static vIsBuilding As Boolean
        
        If (vIsBuilding) Then Exit Sub
        vIsBuilding = True
        
        Me.lstvXplorer.MousePointer = ccArrowHourglass
        Me.trvDirectories.MousePointer = ccArrowHourglass
        
        'Call buildExplorerTree(ScanFS(Node.Key), Node)
        Node.Expanded = True
        
        Me.trvDirectories.MousePointer = ccDefault
        Me.lstvXplorer.MousePointer = ccDefault
        
        vIsBuilding = False
        
    End If
    
    Call buildExplorerListView(Node)
    
End Sub



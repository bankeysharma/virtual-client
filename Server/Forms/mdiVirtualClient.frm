VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Virtual Client:"
   ClientHeight    =   6225
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8430
   Icon            =   "mdiVirtualClient.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   5850
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "keyInfo1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "keyInfo2"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "keyStatus"
            Object.ToolTipText     =   "State of Active Client."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6641
            Key             =   "keyMessage"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picComponents 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   5460
      Left            =   0
      ScaleHeight     =   5460
      ScaleWidth      =   4290
      TabIndex        =   2
      Top             =   390
      Width           =   4290
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   2040
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer tmrAutoRefresh 
         Enabled         =   0   'False
         Left            =   3735
         Top             =   1035
      End
      Begin VB.Timer tmrMessageCleaner 
         Enabled         =   0   'False
         Left            =   3765
         Top             =   4770
      End
      Begin VB.Timer tmrRandomView 
         Enabled         =   0   'False
         Interval        =   15000
         Left            =   3765
         Top             =   45
      End
      Begin ComCtl3.CoolBar cbComponents 
         Height          =   4125
         Left            =   45
         TabIndex        =   3
         Top             =   855
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   7276
         BandCount       =   2
         FixedOrder      =   -1  'True
         BandBorders     =   0   'False
         Orientation     =   1
         _CBWidth        =   2055
         _CBHeight       =   4125
         _Version        =   "6.0.8169"
         Child1          =   "trvClients"
         MinHeight1      =   1995
         Width1          =   1995
         BandPicture1    =   "mdiVirtualClient.frx":038A
         Key1            =   "keyClientList"
         NewRow1         =   0   'False
         Child2          =   "tlbServices"
         MinHeight2      =   1605
         Width2          =   495
         BandPicture2    =   "mdiVirtualClient.frx":2A7AE
         Key2            =   "keyServices"
         NewRow2         =   0   'False
         Begin MSComctlLib.Toolbar tlbServices 
            Height          =   990
            Left            =   225
            TabIndex        =   8
            Top             =   2160
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   1746
            ButtonWidth     =   2223
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            TextAlignment   =   1
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Desktop"
                  Key             =   "keyDesktop"
                  Object.ToolTipText     =   "Bring Desktop. (Ctrl+D)"
                  Style           =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "   Processes"
                  Key             =   "keyProcesses"
                  Object.ToolTipText     =   "Ride running processes. (Ctrl+P)"
                  Style           =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Explorer"
                  Key             =   "keyExplorer"
                  Object.ToolTipText     =   "Explore client's file system. (Ctrl+F)"
                  Style           =   2
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.TreeView trvClients 
            Height          =   1935
            Left            =   30
            TabIndex        =   4
            Top             =   30
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   3413
            _Version        =   393217
            Indentation     =   88
            LabelEdit       =   1
            Sorted          =   -1  'True
            Style           =   7
            Checkboxes      =   -1  'True
            HotTracking     =   -1  'True
            SingleSel       =   -1  'True
            Appearance      =   1
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
      Begin MSComctlLib.ImageList imgList16 
         Left            =   3105
         Top             =   1935
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   52
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4715B
               Key             =   "imgNetwork1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":475AD
               Key             =   "imgClient1"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":479FF
               Key             =   "imgClient2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":47E51
               Key             =   "imgExplorer"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":482A3
               Key             =   "imgHierarchy"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":486F5
               Key             =   "imgSecondaryRoot1"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":48B47
               Key             =   "imgDesktop"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":48F99
               Key             =   "imgClient3"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":493EB
               Key             =   "imgGraph1"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4983D
               Key             =   "imgClose"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":49997
               Key             =   "imgFind"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":49DE9
               Key             =   "imgUsers"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4AAC3
               Key             =   "imgExit"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4ADDD
               Key             =   "imgOfflineClient"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4B0F7
               Key             =   "imgMyComputer"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4B491
               Key             =   "imgNetwork"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4B82B
               Key             =   "imgDisconnectedClient"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4C505
               Key             =   "imgMyComputerOffline"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4C89F
               Key             =   "imgClient4"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4CCF1
               Key             =   "imgClient5"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4CE4B
               Key             =   "imgTileHorizontal"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4D1E5
               Key             =   "imgTileVertical"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4D57F
               Key             =   "imgMyComputerOffline2"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4D91B
               Key             =   "imgZoomIn"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4E5F7
               Key             =   "imgZoomOut"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4F2D3
               Key             =   "imgRandom"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":4F727
               Key             =   "imgDownArrow1"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":50401
               Key             =   "imgLeftArrow1"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":510DB
               Key             =   "imgRightArrow1"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":51DB5
               Key             =   "imgUpArrow1"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":52A8F
               Key             =   "imgFolderClose1"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":53769
               Key             =   "imgFolderOpen1"
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":560FB
               Key             =   "imgHardDrive"
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5714D
               Key             =   "imgFile2"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":57E27
               Key             =   "imgFile1"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":58B01
               Key             =   "imgListView"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":58C13
               Key             =   "imgSmallIconView"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":58D25
               Key             =   "imgIconView"
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":58E37
               Key             =   "imgMoveUpLevel"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":591D1
               Key             =   "imgMoveBack"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5956B
               Key             =   "imgMoveNext"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":59905
               Key             =   "imgPaste"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":59C9F
               Key             =   "imgDelete"
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5A039
               Key             =   "imgCut"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5A3D3
               Key             =   "imgCopy"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5A76D
               Key             =   "imgProperties"
            EndProperty
            BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5AB07
               Key             =   "imgFolders"
            EndProperty
            BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5AEA1
               Key             =   "imgSearch"
            EndProperty
            BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5B23B
               Key             =   "imgRefresh"
            EndProperty
            BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5B5D5
               Key             =   "imgReload"
            EndProperty
            BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5C2AF
               Key             =   "imgFullScreen"
            EndProperty
            BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5C649
               Key             =   "imgProcess"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList imgList32 
         Left            =   3105
         Top             =   2610
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   28
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5D323
               Key             =   "imgGraph1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5D775
               Key             =   "imgDesktop"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5DBC7
               Key             =   "imgNetwork3"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5E019
               Key             =   "imgExplorer2"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5ECF3
               Key             =   "imgDesktop2"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":5F9CD
               Key             =   "imgNetwork"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":606A7
               Key             =   "imgProcess"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":61381
               Key             =   "imgMyComputer"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":6205B
               Key             =   "imgMyDocuments"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":62D35
               Key             =   "imgExit"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":63187
               Key             =   "imgOfflineClient"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":634A1
               Key             =   "imgPerformance1"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":6417B
               Key             =   "imgPerformance2"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":64E55
               Key             =   "imgServices"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":65B2F
               Key             =   "imgTools1"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":66809
               Key             =   "imgTools2"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":674E3
               Key             =   "imgMessage"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":681BD
               Key             =   "imgUserSecurity"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":68E97
               Key             =   "imgClient3"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":692E9
               Key             =   "imgProperties"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":69443
               Key             =   "imgRefresh"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":6959D
               Key             =   "imgZoomIn"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":6A279
               Key             =   "imgZoomOut"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":6AF55
               Key             =   "imgHardDrive"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":6BC2F
               Key             =   "imgFolderClose1"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":6C909
               Key             =   "imgFolderOpen1"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":6F29B
               Key             =   "imgFile2"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiVirtualClient.frx":6FF75
               Key             =   "imgFile1"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Auto refresh "
         Height          =   195
         Index           =   2
         Left            =   3300
         TabIndex        =   12
         Top             =   1485
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Timer Random View"
         Height          =   195
         Index           =   1
         Left            =   2715
         TabIndex        =   11
         Top             =   495
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Timer Message Cleaner"
         Height          =   195
         Index           =   0
         Left            =   2445
         TabIndex        =   10
         Top             =   5220
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label lblClientInfo 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   510
         Left            =   45
         TabIndex        =   6
         Top             =   360
         Width           =   2865
      End
      Begin VB.Label lblClientName 
         Caption         =   "Label1"
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
         Left            =   45
         TabIndex        =   5
         Top             =   45
         Width           =   2715
      End
   End
   Begin ComCtl3.CoolBar cbStandard 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   688
      BandCount       =   4
      _CBWidth        =   8430
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      Child1          =   "tlbStandard"
      MinHeight1      =   330
      Width1          =   4005
      Key1            =   "keyStandard"
      NewRow1         =   0   'False
      Child2          =   "tlbNavigation"
      MinHeight2      =   330
      Width2          =   8010
      Key2            =   "keyExplorer"
      NewRow2         =   0   'False
      Child3          =   "tlbProcess"
      MinWidth3       =   105
      MinHeight3      =   330
      Width3          =   555
      Key3            =   "keyProcess"
      NewRow3         =   0   'False
      Caption4        =   "<Caption>"
      Child4          =   "progressBar"
      MinHeight4      =   195
      Width4          =   1005
      Key4            =   "keyProgress"
      NewRow4         =   0   'False
      BandStyle4      =   1
      BandEmbossPicture4=   -1  'True
      Begin MSComctlLib.Toolbar tlbProcess 
         Height          =   330
         Left            =   7200
         TabIndex        =   14
         Top             =   30
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyKill"
               Object.ToolTipText     =   "Kill"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyProperties"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "keyAssociators"
                     Text            =   "Associators"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbNavigation 
         Height          =   330
         Left            =   4200
         TabIndex        =   13
         Top             =   30
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "keyMovePrev"
               Object.ToolTipText     =   "Go to last folder visited. (Ctrl+B)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyMoveNext"
               Object.ToolTipText     =   "Go to next folder visited. (Ctrl+N)"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "keyMoveUp"
               Object.ToolTipText     =   "Move one level up. (Ctrl+U)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyFolders"
               Object.ToolTipText     =   "Show folders."
               Style           =   1
               Value           =   1
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keySearch"
               Object.ToolTipText     =   "Search..."
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyCut"
               Object.ToolTipText     =   "Cut (Ctrl+X)"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyCopy"
               Object.ToolTipText     =   "Copy (Ctrl+C)"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyPaste"
               Object.ToolTipText     =   "Paste (Ctrl+V)"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyDelete"
               Object.ToolTipText     =   "Delete (Del)"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyIconView"
               Object.ToolTipText     =   "Icon View"
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keySmallIconView"
               Object.ToolTipText     =   "Small Icon View."
               Style           =   2
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyListView"
               Object.ToolTipText     =   "List View."
               Style           =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar progressBar 
         Height          =   195
         Left            =   8400
         TabIndex        =   9
         Top             =   90
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.Toolbar tlbStandard 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyBroadSpectrumView"
               Object.ToolTipText     =   "Broad-spectrum View (F6)"
               Style           =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keySolitaryView"
               Object.ToolTipText     =   "Solitary View."
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyRandomView"
               Object.ToolTipText     =   "Random view. (F7)"
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyFullScreen"
               Object.ToolTipText     =   "Switch to Full Screen Mode. (F8)"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyTileVertical"
               Object.ToolTipText     =   "Arrange Vertically. (F11)"
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyTileHorizontal"
               Object.ToolTipText     =   "Arrange Horizontally. (F12)"
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "keyRefresh"
               Object.ToolTipText     =   "Refresh (F5)"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuMainIPAddresses 
         Caption         =   "&IP Addresses"
      End
      Begin VB.Menu mnuMainBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainLogout 
         Caption         =   "&Logout"
      End
      Begin VB.Menu mnuMainBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewAutoRefresh 
         Caption         =   "Auto Refresh"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "Re&fresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSolitary 
         Caption         =   "&Solitary View"
         Begin VB.Menu mnuViewSolitaryDesktop 
            Caption         =   "&Desktop"
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuViewSolitaryProcesses 
            Caption         =   "&Processes"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuViewSolitaryFileSystem 
            Caption         =   "&File System"
            Shortcut        =   ^F
         End
      End
      Begin VB.Menu mnuViewBroadSpectrum 
         Caption         =   "&Broad Spectrum View"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuViewRandom 
         Caption         =   "&Random View"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuViewsFullScreen 
         Caption         =   "F&ull Screen"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTileVertically 
         Caption         =   "Tile &Vertically"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuViewTileHorizontally 
         Caption         =   "Tile &Horizontally"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuUtility 
      Caption         =   "&Utility"
      Begin VB.Menu mnuUtilityNotification 
         Caption         =   "&Notification"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuExplorer 
      Caption         =   "E&xplorer"
      Begin VB.Menu mnuExplorerNext 
         Caption         =   "&Next"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuExplorerBack 
         Caption         =   "&Back"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuExplorerUp 
         Caption         =   "&Up"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuExplorerBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExplorerCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuExplorerCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuExplorerPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuExplorerDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuExplorerBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExplorerFolders 
         Caption         =   "&Folders"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExplorerSearch 
         Caption         =   "&Search"
      End
   End
   Begin VB.Menu mnuProcess 
      Caption         =   "&Process"
      Begin VB.Menu mnuProcessKill 
         Caption         =   "&Kill"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuDesktopPopup 
      Caption         =   "DesktopPopup"
      Begin VB.Menu mnuDesktopPopupSolitaryView 
         Caption         =   "Solitary View"
         Begin VB.Menu mnuDesktopPopupSolitaryViewServices 
            Caption         =   "Desktop"
            Index           =   1
         End
         Begin VB.Menu mnuDesktopPopupSolitaryViewServices 
            Caption         =   "Processes"
            Index           =   2
         End
         Begin VB.Menu mnuDesktopPopupSolitaryViewServices 
            Caption         =   "File System"
            Index           =   3
         End
      End
      Begin VB.Menu mnuDesktopPopupBroadSpectrumView 
         Caption         =   "Broad-Spectrun View"
      End
      Begin VB.Menu mnuDesktopPopupRandomView 
         Caption         =   "Random View"
      End
      Begin VB.Menu mnuDesktopPopupFullScreen 
         Caption         =   "Full Screen View"
      End
      Begin VB.Menu mnuDesktopPopupBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesktopPopupConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuDesktopPopupDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuDesktopPopupOn 
         Caption         =   "On"
      End
      Begin VB.Menu mnuDesktopPopupOff 
         Caption         =   "Off"
      End
      Begin VB.Menu mnuDesktopPopupBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesktopPopupNotify 
         Caption         =   "Notification"
      End
      Begin VB.Menu mnuDesktopPopupBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesktopPopupLogout 
         Caption         =   "LogOff"
      End
      Begin VB.Menu mnuDesktopPopupReboot 
         Caption         =   "Reboot"
      End
      Begin VB.Menu mnuDesktopPopupShutDown 
         Caption         =   "Shut Down"
      End
      Begin VB.Menu mnuDesktopPopupBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDesktopPopupEliminate 
         Caption         =   "Eliminate"
      End
   End
   Begin VB.Menu mnuPadProcess 
      Caption         =   "PadProcess"
      Begin VB.Menu mnuPadProcessKill 
         Caption         =   "Kill"
      End
      Begin VB.Menu mnuPadBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPadProcessAssiciators 
         Caption         =   "Associators"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cCbComponentWidth As Integer = 2000
Private Const cCbComponentTopMargin As Integer = 800
Private Const cCbComponentBottomMargin As Integer = 0 '500

'Private Const cClientListBandHeight As Integer = 4450
Private Const cClientListBandMinHeight As Integer = 1000

Private Const cServicesBandHeight As Integer = 1950
Private Const cServicesBandMinHeight As Integer = 1000

Private Const cStandardBandWidth As Integer = 3000
Private Const cProgressBarBandWidth As Integer = 2500
Private Const cExplorerBandWidth As Integer = 4000
Private Const cProcessBandWidth As Integer = 1000

Private Const cMessageDisplayTimeout As Byte = 30
Private Const cAutoRefreshTimeSpan As Byte = 1

Private Sub cbComponents_Resize()

    Me.cbComponents.left = Me.picComponents.ScaleLeft
    Me.picComponents.Width = Me.cbComponents.Width
    
End Sub

Private Sub setToolBar()
    
    With Me.tlbStandard
        Set .ImageList = Me.imgList16
        .Buttons("keyRefresh").Image = Me.imgList16.ListImages("imgRefresh").Index
        .Buttons("keySolitaryView").Image = Me.imgList16.ListImages("imgMyComputer").Index
        .Buttons("keyBroadSpectrumView").Image = Me.imgList16.ListImages("imgHierarchy").Index
        .Buttons("keyRandomView").Image = Me.imgList16.ListImages("imgRandom").Index
        .Buttons("keyTileVertical").Image = Me.imgList16.ListImages("imgTileVertical").Index
        .Buttons("keyTileHorizontal").Image = Me.imgList16.ListImages("imgTileHorizontal").Index
        .Buttons("keyFullScreen").Image = Me.imgList16.ListImages("imgFullScreen").Index
    End With

End Sub

Private Sub setToolBarProcess()
    With Me.tlbProcess
        Set .ImageList = Me.imgList16
        .Buttons("keyKill").Image = Me.imgList16.ListImages("imgDelete").Index
        .Buttons("keyProperties").Image = Me.imgList16.ListImages("imgProperties").Index
    End With
End Sub

'*************************************************
'******** To set the entire layout i.e.
'******** Visual interface
'*************************************************

Private Sub setInterface()

On Error GoTo Err_Handle
Err.Clear

    Me.mnuDesktopPopup.Visible = False
    Me.mnuExplorer.Visible = False
    Me.mnuProcess.Visible = False
    Me.mnuPadProcess.Visible = False
    
    Call setToolBar
    Call setToolBarExplorer
    Call setToolBarProcess
    
    With Me.cbStandard.Bands("keyStandard")
        .Width = cStandardBandWidth
    End With
    
    With Me.cbStandard.Bands("keyExplorer")
        .Width = cExplorerBandWidth
        .Visible = False
    End With
    
    With Me.cbStandard.Bands("keyProgress")
        .MinWidth = cProgressBarBandWidth
        .Visible = False
    End With
    
    With Me.cbStandard.Bands("keyProcess")
        .MinWidth = cProcessBandWidth
        .Visible = False
    End With
    
    With Me.cbComponents.Bands("keyServices")
        .Width = cServicesBandHeight
    End With

    With Me.cbComponents.Bands("keyClientList")
        .MinHeight = cCbComponentWidth
        .MinWidth = cClientListBandMinHeight
        '.Width = cClientListBandHeight
    End With
    
    Me.lblClientInfo.Caption = vbNullString
    Me.lblClientName.Caption = vbNullString
    
    Set Me.tlbServices.ImageList = Me.imgList32
    With Me.tlbServices
        .Buttons("keyDesktop").Image = Me.imgList32.ListImages("imgDesktop2").Index
        .Buttons("keyProcesses").Image = Me.imgList32.ListImages("imgProcess").Index
        .Buttons("keyExplorer").Image = Me.imgList32.ListImages("imgExplorer2").Index
    End With
    
    Me.tmrMessageCleaner.Interval = cMessageDisplayTimeout * 500
    Me.tmrMessageCleaner.Enabled = False
    
    Me.tmrAutoRefresh.Interval = cAutoRefreshTimeSpan * 500
    Me.tmrAutoRefresh.Enabled = False
    
    '***********************************
    '************* Initially no view is there
    'View = vuNone
       

Exit_Lable:

    Exit Sub

Err_Handle:
    Select Case (Err.Number)
        Case 380 'invalid property value
            'Error due to providing height/ width
            'while form is minimizing or minimized
            Resume Next
        Case Else
            Call oDB.showError(Err)
    End Select
    
End Sub

'Private Sub MDIForm_Click()
'    Dim a As String
'    Dim i As Byte
'
'    For i = 33 To 100
'        a = a & "(" & i & " : " & Chr(i) & ") "
'    Next i
'
'    MsgBox a
'
'End Sub

Private Sub MDIForm_Load()
    
    Call setInterface

End Sub

Private Sub MDIForm_Resize()

    On Error Resume Next
    
    If (Me.WindowState = vbMinimized) Then Exit Sub
    
    Me.cbComponents.tOp = Me.picComponents.ScaleTop + cCbComponentTopMargin
    Me.cbComponents.Height = Me.ScaleHeight - (cCbComponentTopMargin + cCbComponentBottomMargin)

    With Me.cbComponents.Bands("keyServices")
        .Width = cServicesBandHeight
    End With
    
    With Me.cbComponents.Bands("keyClientList")
        .Width = Me.cbComponents.Height - cServicesBandHeight
    End With
    
'    If (Not Me.ActiveForm Is Nothing) Then
'        If (Me.ActiveForm.Name = "frmAuthentication") Then _
'            Call frmAuthentication.reSetDimensions
'    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    Call Quit
    
    Cancel = 1
    
End Sub

Private Sub mnuDesktopPopupBroadSpectrumView_Click()
    Call mdlViews.broadSpectrumView
End Sub

Private Sub mnuDesktopPopupConnect_Click()
    Call activeClient.Connect
End Sub

Private Sub mnuDesktopPopupDisconnect_Click()
    Call activeClient.Disconnect
End Sub

Private Sub mnuDesktopPopupEliminate_Click()
    
    Call colClients.Remove(CStr(activeClient.ID))
    Set activeClient = Nothing

End Sub

Private Sub mnuDesktopPopupFullScreen_Click()
    Call mdlViews.fullScreenView
End Sub

Private Sub mnuDesktopPopupLogout_Click()
    Call activeClient.sendSignal(enmSigLogOff)
End Sub

Private Sub mnuDesktopPopupNotify_Click()
    Call showNotification(activeClient)
End Sub

Private Sub mnuDesktopPopupOff_Click()

    Call activeClient.Off
    
End Sub

Private Sub mnuDesktopPopupOn_Click()
    
    With activeClient
        .treeNode.Checked = True
        .shouldDisplay = True
        .showTerminal
    End With
    
    If (Not View = vuBroadSpectrum) Then
        'View = vuBroadSpectrum
        Call mdlViews.broadSpectrumView
        
        If (frmMDI.tlbStandard.Buttons("keyTileHorizontal").Value = tbrPressed) Then
            frmMDI.standardAction ("keyTileHorizontal")
        Else
            frmMDI.standardAction ("keyTileVertical")
        End If
    End If

End Sub

Private Sub mnuDesktopPopupRandomView_Click()

    Call mdlViews.randomView
    
End Sub

Private Sub mnuDesktopPopupReboot_Click()

    Call activeClient.sendSignal(enmSigReboot)
    
End Sub

Private Sub mnuDesktopPopupShutDown_Click()

    Call activeClient.sendSignal(enmSigShutDown)
    
End Sub

Private Sub mnuDesktopPopupSolitaryViewServices_Click(Index As Integer)
    
    If (Index = 1) Then
        
        Call mnuViewSolitaryDesktop_Click
        Call activeClient.showDesktop
    
    ElseIf (Index = 2) Then
        
        Call mnuViewSolitaryProcesses_Click
    
    ElseIf (Index = 3) Then
        
        Call mnuViewSolitaryFileSystem_Click
    
    End If
    
End Sub

Private Sub mnuExplorerBack_Click()
    Call navigatoryAction("keyMovePrev")
End Sub

Private Sub mnuExplorerCut_Click()
    Call navigatoryAction("keyCut")
End Sub

Private Sub mnuExplorerFolders_Click()
    Call navigatoryAction("keyFolders")
End Sub

Private Sub mnuExplorerNext_Click()
    Call navigatoryAction("keyMoveNext")
End Sub

Private Sub mnuExplorerSearch_Click()
    Call navigatoryAction("keySearch")
End Sub

Private Sub mnuExplorerUp_Click()
    Call navigatoryAction("keyMoveUp")
End Sub

Private Sub mnuHelpAbout_Click()
    
    Load frmAbout
    Call frmAbout.Show(vbModal)
    
End Sub

Private Sub mnuMainExit_Click()
    
    Unload Me
    
End Sub

Private Sub mnuMainIPAddresses_Click()
'    MsgBox Me.Winsock1.LocalIP
    Load frmIPAddresses
    Call frmIPAddresses.Show(vbModal)
End Sub

Private Sub mnuMainLogout_Click()
    Call LogIn
End Sub

Private Sub mnuPadProcessAssiciators_Click()
    Call processAction("keyAssociators")
End Sub

Private Sub mnuPadProcessKill_Click()
    Call processAction("keyKill")
End Sub

Private Sub mnuProcessKill_Click()
    Call processAction("keyKill")
End Sub

Private Sub mnuUtilityNotification_Click()
    Call showNotification(activeClient)
End Sub

Private Sub mnuViewAutoRefresh_Click()
    Me.mnuViewAutoRefresh.Checked = Not Me.mnuViewAutoRefresh.Checked
    If (Not Me.mnuViewAutoRefresh.Checked) Then
        Me.tmrAutoRefresh.Enabled = False
    Else
        'Me.tmrAutoRefresh.Interval = cAutoRefreshTimeSpan * 500
        Me.tmrAutoRefresh.Enabled = True
    End If
End Sub

Private Sub mnuViewBroadSpectrum_Click()
    Call standardAction("keyBroadSpectrumView")
End Sub

Private Sub mnuViewRandom_Click()
    Call standardAction("keyRandomView")
End Sub

Private Sub mnuViewRefresh_Click()
    Call standardAction("keyRefresh")
End Sub

Private Sub mnuViewsFullScreen_Click()
    Call standardAction("keyFullScreen")
End Sub

Private Sub mnuViewSolitaryDesktop_Click()
    CurrentService = srvDesktopCapture
    Call solitaryView
End Sub

Private Sub mnuViewSolitaryFileSystem_Click()
    CurrentService = srvExplorer
    Call solitaryView
End Sub

Private Sub mnuViewSolitaryProcesses_Click()
    CurrentService = srvProcessMonitoring
    Call solitaryView
End Sub

Private Sub mnuViewTileHorizontally_Click()
    Call standardAction("keyTileHorizontal")
End Sub

Private Sub mnuViewTileVertically_Click()
    Call standardAction("keyTileVertical")
End Sub

Private Sub tlbNavigation_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call navigatoryAction(Button.key)
End Sub

Private Sub tlbProcess_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call processAction(Button.key)
End Sub

Private Sub tlbProcess_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Call processAction(ButtonMenu.key)
End Sub

Private Sub tlbServices_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case (Button.key)
        Case "keyDesktop"
            CurrentService = srvDesktopCapture
        Case "keyProcesses"
            CurrentService = srvProcessMonitoring
        Case "keyExplorer"
            CurrentService = srvExplorer
    End Select
    
    Call solitaryView
    
End Sub

Private Sub tlbStandard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call standardAction(Button.key)
End Sub

Private Sub tmrAutoRefresh_Timer()
    Call standardAction("keyRefresh")
End Sub

Private Sub tmrMessageCleaner_Timer()
    frmMDI.statusBar.Panels("keyMessage").text = vbNullString
    Me.tmrMessageCleaner.Enabled = False
End Sub

Private Sub tmrRandomView_Timer()
    If (View = vuRandom) Then
        Call randomView
    Else
        Me.tmrRandomView.Enabled = False
    End If
End Sub

Private Sub trvClients_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = vbRightButton) Then Me.showDesktopPopup
End Sub

Private Sub trvClients_NodeCheck(ByVal Node As MSComctlLib.Node)
    
    Dim oChildNode As MSComctlLib.Node
    Dim oParentNode As MSComctlLib.Node
    Dim vCheck As Boolean
    
    If (Node.Parent Is Nothing) Then
        Set oChildNode = Me.trvClients.Nodes(Node.key).Child
        While (Not oChildNode Is Nothing)
            If (colClients(oChildNode.key).IsConnected) Then
                colClients(oChildNode.key).shouldDisplay = Node.Checked
            Else
                colClients(oChildNode.key).shouldDisplay = False
            End If
            Set oChildNode = oChildNode.Next
        Wend
    
    Else
        If (colClients(Node.key).IsConnected) Then _
            colClients(Node.key).shouldDisplay = Node.Checked
        
        
        '********* Validating parent's check prop.
        Set oParentNode = Me.trvClients.Nodes(Node.key).Parent
        vCheck = False
        
        If (Not oParentNode Is Nothing) Then
            
            Set oChildNode = oParentNode.Child
            
            While (Not oChildNode Is Nothing)
                vCheck = (oChildNode.Checked Or vCheck)
                Set oChildNode = oChildNode.Next
            Wend
            
            oParentNode.Checked = vCheck
        End If
    End If

End Sub

Private Sub trvClients_NodeClick(ByVal Node As MSComctlLib.Node)
    If (Node.Parent Is Nothing) Then
        Set activeClient = Nothing
        'Me.trvClients.Nodes(Node.Key).Expanded = True
        
    Else
        
        Set activeClient = colClients(Node.key)
        
    End If
End Sub

Public Sub showDesktopPopup()
    
    If (activeClient Is Nothing) Then Exit Sub
        
    If (Not activeClient.IsConnected) Then
        
        Me.mnuDesktopPopupConnect.Visible = True
        
        Me.mnuDesktopPopupSolitaryView.Visible = False
        Me.mnuDesktopPopupBar2.Visible = False
        Me.mnuDesktopPopupDisconnect.Visible = False
        Me.mnuDesktopPopupBar3.Visible = False
        Me.mnuDesktopPopupLogout.Visible = False
        Me.mnuDesktopPopupReboot.Visible = False
        Me.mnuDesktopPopupShutDown.Visible = False
        Me.mnuDesktopPopupOff.Visible = False
        Me.mnuDesktopPopupOn.Visible = False
        Me.mnuDesktopPopupNotify.Visible = False
        Me.mnuDesktopPopupBar5.Visible = False
        
        Me.mnuDesktopPopupBroadSpectrumView.Visible = False
        Me.mnuDesktopPopupFullScreen.Visible = False
        Me.mnuDesktopPopupRandomView.Visible = False
        Me.mnuDesktopPopupSolitaryView.Visible = False
    
    Else
        
        Me.mnuDesktopPopupConnect.Visible = False
        
        Me.mnuDesktopPopupSolitaryView.Visible = True
        Me.mnuDesktopPopupBar2.Visible = True
        Me.mnuDesktopPopupDisconnect.Visible = True
        Me.mnuDesktopPopupBar3.Visible = True
        Me.mnuDesktopPopupLogout.Visible = True
        Me.mnuDesktopPopupReboot.Visible = True
        Me.mnuDesktopPopupShutDown.Visible = True
        Me.mnuDesktopPopupNotify.Visible = True
        Me.mnuDesktopPopupBar5.Visible = True
        
        If (Not activeClient.IsON) Then
            Me.mnuDesktopPopupOff.Visible = False
            Me.mnuDesktopPopupOn.Visible = True
        Else
            Me.mnuDesktopPopupOff.Visible = True
            Me.mnuDesktopPopupOn.Visible = False
        End If
        Me.mnuDesktopPopupBroadSpectrumView.Visible = True
        Me.mnuDesktopPopupFullScreen.Visible = True
        Me.mnuDesktopPopupRandomView.Visible = True
        Me.mnuDesktopPopupSolitaryView.Visible = True
        
        If (View = vuBroadSpectrum) Then
            Me.mnuDesktopPopupBroadSpectrumView.Visible = False
        ElseIf (View = vuFullScreen) Then
            Me.mnuDesktopPopupFullScreen.Visible = False
        ElseIf (View = vuRandom) Then
            Me.mnuDesktopPopupRandomView.Visible = False
        ElseIf (View = vuSolitary) Then
            Me.mnuDesktopPopupSolitaryView.Visible = False
        End If
    End If
    
    Call Me.PopupMenu(frmMDI.mnuDesktopPopup)
End Sub

'*******************************************
'***** Action Sequence for Standard Toolbar
'*******************************************

Public Sub standardAction(keyValue As String)
    
    Me.tmrAutoRefresh.Enabled = False
    
    Select Case (keyValue)
        Case "keyRefresh"
            If (View = vuSolitary) Then
                
                If (activeClient Is Nothing) Then Exit Sub
                
                If (CurrentService = srvDesktopCapture) Then
                    Call activeClient.sendSignal(enmSigSendDesktop)
                ElseIf (CurrentService = srvProcessMonitoring) Then
                    Call frmServices.oProcesses.enumerateProcesses
                ElseIf (CurrentService = srvExplorer) Then
                    Call activeClient.sendSignal(enmSigSendFileSystem, _
                                                    cClauseDelimiter & "0")
                End If
                
            ElseIf (View = vuBroadSpectrum) Then
                Call standardAction("keyBroadSpectrumView")
                
            ElseIf (View = vuFullScreen) Then
                Call activeClient.sendSignal(enmSigSendDesktop)
                
            End If
            
        Case "keyBroadSpectrumView"
            Call broadSpectrumView
            
        Case "keySolitaryView"
        
            CurrentService = srvDesktopCapture
            Call solitaryView
            
        Case "keyRandomView"
            Call randomView
            frmMDI.tmrRandomView.Enabled = True
            
        Case "keyTileVertical"
            If (View = vuSolitary) Then Exit Sub
            Call frmMDI.Arrange(vbTileVertical)
            frmMDI.tlbStandard.Buttons("keyTileVertical").Value = tbrPressed
            frmMDI.mnuViewTileHorizontally.Checked = False
            frmMDI.mnuViewTileVertically.Checked = True
            
        Case "keyTileHorizontal"
            If (View = vuSolitary) Then Exit Sub
            
            Call frmMDI.Arrange(vbTileHorizontal)
            
            frmMDI.tlbStandard.Buttons("keyTileHorizontal").Value = tbrPressed
            frmMDI.mnuViewTileHorizontally.Checked = True
            frmMDI.mnuViewTileVertically.Checked = False
            
        Case "keyFullScreen"
            Call mdlViews.fullScreenView
            
    End Select
    
    If (Me.mnuViewAutoRefresh.Checked) Then _
        Me.tmrAutoRefresh.Enabled = True
        
End Sub

Public Sub navigatoryAction(key As String)
    Select Case (key)
        Case "keyMoveUp"
            Call moveUp
        Case "keyMovePrev"
            Call moveBack
        Case "keyFolders"
            
            Me.mnuExplorerFolders.Checked = Not Me.mnuExplorerFolders.Checked
            
            If (Me.mnuExplorerFolders.Checked) Then
                Me.tlbNavigation.Buttons("keyFolders").Value = tbrPressed
            Else
                Me.tlbNavigation.Buttons("keyFolders").Value = tbrUnpressed
            End If
            
            frmServices.cbXplorer.Bands("keyTreeView").Visible = Not _
                frmServices.cbXplorer.Bands("keyTreeView").Visible
        
        Case "keySearch"
            'Call buildExplorerTree(ScanFS(vbNullString))
        Case "keyIconView"
            frmServices.lstvXplorer.Arrange = lvwAutoTop
            frmServices.lstvXplorer.View = lvwIcon
        Case "keyListView"
            frmServices.lstvXplorer.Arrange = lvwAutoTop
            frmServices.lstvXplorer.View = lvwList
        Case "keySmallIconView"
            frmServices.lstvXplorer.Arrange = lvwAutoTop
            frmServices.lstvXplorer.View = lvwSmallIcon
        Case "keyQuit"
            End
    End Select
End Sub

Public Sub processAction(key As String)
    
    DoEvents
    
    Select Case (key)
        Case "keyKill"
            Call frmServices.oProcesses.Kill
        Case "keyProperties"
            Call frmServices.oProcesses.showProperties
        Case "keyAssociators"
            Call frmServices.oProcesses.showAssociators
    End Select

End Sub

Private Sub showNotification(pClient As clsClient)
    Load frmNotify
    
'    If (Not pClient Is Nothing) Then _
'        frmNotify.cmbClients.ListIndex = pClient.ID - 1
        
    Call frmNotify.Show(vbModal)
End Sub

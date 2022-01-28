Attribute VB_Name = "mdlExplorerServer"
Option Explicit
Option Private Module

Public Enum enmConversionUnit
    
    enmByte = 1
    enmKB = 2
    enmMB = 3
    enmGB = 4
    enmBest = 5

End Enum

Public colFolders As New Collection
Public oActiveFolder As clsFolder

Private Const cDriveDelimiter As String = ">"
Private Const cNextLevelDelimiter As String = "\"
Private Const cPrevLevelDelimiter As String = "/"
Private Const cSectionDelimiter As String = "|"
Private Const cFolderDelimiter As String = "^"
Private Const cFileDelimiter As String = "*"
Private Const cAttributeDelimiter As String = ":"

Private Const cIsDrive As Byte = 0
Private Const cIsFolder As Byte = 1
Private Const cIsFile As Byte = 2

Private colMoves As New Collection

Public Function buildExplorerTree(ByRef fsString As String, Optional ByVal parentNode As MSComctlLib.Node = Nothing) As MSComctlLib.Node
    
    On Error GoTo Err_Handle
    Err.Clear
    
    Dim vInfo As Variant
    Dim oNode As MSComctlLib.Node
    
    If (parentNode Is Nothing) Then
        
        Dim vDriveLevel As Variant
        Dim vNextLevel As Variant
            
        Dim vNextLevelCount As Integer
        Dim vDriveCount As Integer
        
        Dim oRootNode As MSComctlLib.Node
        Dim oParentFolderNode As MSComctlLib.Node
        
        Dim oDrive As clsDrive
        
        With frmServices.trvDirectories
            .Nodes.Clear
            .ImageList = frmMDI.imgList16.Object
            Set oRootNode = .Nodes.Add(, , "keyRoot", "Virtual Client (" & activeClient.Alias & ")", frmMDI.imgList16.ListImages("imgMyComputer").Index)
        End With
        
        vDriveLevel = Split(fsString, cDriveDelimiter)
        
        vDriveCount = UBound(vDriveLevel)
        
        Do While (vDriveCount >= LBound(vDriveLevel))
            
            vNextLevel = Split(vDriveLevel(vDriveCount), cNextLevelDelimiter)
            If (UBound(vNextLevel) > 0) Then
            
                vInfo = Split(vNextLevel(0), cAttributeDelimiter)
                
                Set oDrive = New clsDrive
                
                With oDrive
                    .DriveLetter = vInfo(0)
                    .DriveType = vInfo(1)
                    .FileSystem = vInfo(2)
                    .FreeSpace = vInfo(3)
                    .TotalSize = vInfo(4)
                    .VolumeName = IIf(vInfo(5) = vbNullString, "Local Disk", vInfo(5))
                    
                    Set oNode = frmServices.trvDirectories.Nodes.Add(, , , .VolumeName & " (" & .DriveLetter & ":)")
                    With oNode
                        .Image = frmMDI.imgList16.ListImages("imgHardDrive").Index
                        Set .Parent = oRootNode
                        .key = oDrive.DriveLetter & ":"
                        .Tag = cIsDrive
                    End With
                    
                    '********** Eliminating first cell
                    vNextLevel(0) = vbNullString
                    
                    Call colFolders.Add(oDrive, oNode.key)
                    Call buildExplorerTree(Join(vNextLevel, "\"), oNode)
                    
                    oNode.Expanded = False
                
                End With
                
                Set oDrive = Nothing
                
            End If
            
            vDriveCount = vDriveCount - 1
        
        Loop
        
        frmServices.trvDirectories.Nodes("keyRoot").Selected = True
        Call buildExplorerListView(frmServices.trvDirectories.Nodes("keyRoot"))
        
    Else
        
        Dim i As Long
        
        Dim arySymbols() As Byte
        Dim ID As Variant
        Dim vObject As String
        
        Dim oFolder As clsFolder
        Dim oFile As clsFile
        
        arySymbols = StrConv(fsString, vbFromUnicode)
        
        vObject = vbNullString
        
        If (parentNode.Tag = cIsDrive) Then
            
            Set oFile = New clsFile
            
            Call setPrgBar("Clearing...", 0, colFolders(parentNode.key).Files.Count)
            DoEvents
            
            For Each oFile In colFolders(parentNode.key).Files
                prgBar.Value = prgBar.Value + 1
                Call colFolders(parentNode.key).Files.Remove(parentNode.key & "\" & oFile.Name)
            Next oFile
            
            Set oFile = Nothing
            
        End If
        
        Call setPrgBar("Building...", LBound(arySymbols), UBound(arySymbols) + 1)
        DoEvents
        
        For ID = LBound(arySymbols) To UBound(arySymbols)
            
            prgBar.Value = prgBar.Value + 1
            
            Select Case (Chr(arySymbols(ID)))
                Case cPrevLevelDelimiter
                    If (Not parentNode Is Nothing) Then
                        parentNode.Expanded = False
                        Set parentNode = parentNode.Parent
                    End If
                
                Case cNextLevelDelimiter
                
                    If (Not oNode Is Nothing) Then
                        parentNode.Expanded = False
                        Set parentNode = oNode
                    End If
                
                Case cFolderDelimiter
                    
                    vInfo = Split(vObject, cAttributeDelimiter)
                    
                    If (UBound(vInfo) > 0) Then
                        Set oFolder = New clsFolder
                        
                        '********************************
                        '   Info sequence of Folder
                        '       Name
                        '       Attributes
                        '       TotalSubFolders
                        '       TotalFiles
                        '       Size
                        '********************************
                        With oFolder
                            .Name = vInfo(0)
                            .Attributes = vInfo(1)
                            .TotalSubFolders = vInfo(2)
                            .TotalFiles = vInfo(3)
                            .Size = vInfo(4)
                            
                            Set oNode = frmServices.trvDirectories.Nodes.Add(, , , .Name)
                            
                            With oNode
                                .Image = frmMDI.imgList16.ListImages("imgFolderClose1").Index
                                Set .Parent = parentNode
                                .key = .Parent.key & "\" & .text
                                .Tag = cIsFolder
                            End With
                                                    
                            .ParentFolderPath = oNode.Parent.key
                                                                                                            
                            Call colFolders.Add(oFolder, oNode.key)
                        
                            Set oFile = New clsFile
                            For Each oFile In colFolders(oNode.key).Files
                                Call colFolders(oNode.key).Files.Remove(oNode.key & "\" & oFile.Name)
                            Next oFile
                        
                        End With
                        
                        Set oFolder = Nothing
                        
                    End If
                
                    vObject = vbNullString
                    
                Case cFileDelimiter
                    
                    vInfo = Split(vObject, cAttributeDelimiter)
                    
                    If (UBound(vInfo) > 0) Then
                        Set oFile = New clsFile
                    
                        '********************************
                        '   Info sequence of File
                        '       Name
                        '       Attributes
                        '       Size
                        '********************************
                        With oFile
                            
                            .Name = vInfo(0)
                            .Attributes = vInfo(1)
                            .Size = vInfo(2)
                                
                            Call colFolders.Item(parentNode.key).Files.Add(oFile, .Name)
                        
                        End With
                        
                    End If
                    
                    vObject = vbNullString
                    
                Case Else
                
                    vObject = vObject & Chr(arySymbols(ID))
                    
            End Select
        
        Next ID
        
        Call offPrgBar
        
    End If

Exit_Handle:

    frmServices.cbXplorer.Visible = True
    
    Exit Function

Err_Handle:
    
    Select Case (Err.Number)
        Case 457
                
            If (oNode Is Nothing) Then Exit Function
            
            colFolders.Remove (oNode.key)
            
            Resume
            
        Case Else
            'Resume Next
            MsgBox Err.Description
    End Select
    

End Function

Public Sub buildExplorerListView(ByVal treeNode As MSComctlLib.Node)
    On Error Resume Next
    
    Dim oChildNode As MSComctlLib.Node
    Dim colFiles As Collection
    
    Call hideParentInfo
    Call hideItemInfo

    frmServices.lstvXplorer.ListItems.Clear
    
    If (Not treeNode.Parent Is Nothing) Then _
        Call colMoves.Add(treeNode)
    
    If (colMoves.Count > 1) Then
        frmMDI.tlbNavigation.Buttons("keyMovePrev").Enabled = True
    End If

    If (Not treeNode.Parent Is Nothing) Then
        frmMDI.tlbNavigation.Buttons("keyMoveUp").Enabled = True
    End If
    
    treeNode.Selected = True
    
    If (Not treeNode.Child Is Nothing) Then
        
        treeNode.Expanded = True
        
        Call setPrgBar("Showing...", 0, treeNode.Children + 1)
        DoEvents
        
        Set oChildNode = treeNode.Child.FirstSibling
        
        Do While (Not oChildNode Is Nothing)
            
            prgBar.Value = prgBar.Value + 1
            'DoEvents
            
            With oChildNode
                With frmServices.lstvXplorer.ListItems.Add(, .key, .text)
                    .Icon = frmMDI.imgList32.ListImages(frmMDI.imgList16.ListImages(oChildNode.Image).key).Index
                    .SmallIcon = oChildNode.Image
                    .Tag = oChildNode.Tag
                End With
            End With
            
            Set oChildNode = oChildNode.Next
            
        Loop
        
    End If
        
    If (Not treeNode.Tag = vbNullString) Then
        Set colFiles = colFolders(treeNode.key).Files
        
        If (Not colFiles Is Nothing) Then
            
            Dim oFile As clsFile
            
            Set oFile = New clsFile
            
            Call setPrgBar("Showing...", 0, colFiles.Count)
            DoEvents
            
            For Each oFile In colFiles
                
                prgBar.Value = prgBar.Value + 1
                'DoEvents
                
                With oFile
                    With frmServices.lstvXplorer.ListItems.Add(, treeNode.key & "\" & .Name, .Name)
                        .Icon = frmMDI.imgList32.ListImages("imgFile1").Index
                        .SmallIcon = frmMDI.imgList16.ListImages("imgFile1").Index
                        .Tag = cIsFile
                    End With
                End With
            Next oFile
            
            Set oFile = Nothing
            Set colFiles = Nothing
            
        End If
    End If
    
    Call offPrgBar
    
    With frmServices.lstvXplorer
        If (.ListItems.Count > 0) Then
            If (Not .ListItems(1) Is Nothing) Then _
                Call showParentInfo(.ListItems(1))
        End If
    End With
End Sub

Public Sub showParentInfo(Item As MSComctlLib.ListItem)
    
    Dim oParentNode As MSComctlLib.Node
    
    If (Item Is Nothing) Then Exit Sub
    
    If (frmServices.lstvXplorer.ListItems(1).Tag = cIsFile) Then
        Dim vParentCode As String
        
        vParentCode = left(Item.key, InStrRev(Item.key, "\") - 1)
        Set oParentNode = frmServices.trvDirectories.Nodes(vParentCode)
    Else
        Set oParentNode = frmServices.trvDirectories.Nodes(Item.key).Parent
    End If

    
    Set frmServices.imgObjectIcon.Picture = frmMDI.imgList32.ListImages(frmMDI.imgList16.ListImages(oParentNode.Image).key).Picture
    frmServices.imgObjectIcon.Visible = True
    
    frmServices.lblParentName.tOp = frmServices.imgObjectIcon.tOp + frmServices.imgObjectIcon.Height + 100
    
    If (Right(Item.key, 1) = ":") Then
        
        With oParentNode
            frmServices.lblParentName.Caption = .text
            frmServices.lblParentName.Visible = True
        End With
    
    Else
    
        If (oParentNode.Tag = cIsDrive) Then
            With colFolders(oParentNode.key)
                frmServices.lblParentName.Caption = .VolumeName & " (" & .DriveLetter & ":)"
                frmServices.lblParentName.Visible = True
            End With
        Else
            With colFolders(oParentNode.key)
                frmServices.lblParentName.Caption = .Name
                frmServices.lblParentName.Visible = True
            End With
        End If
    
    End If

    Call hideItemInfo

End Sub

Public Sub showItemInfo(ByVal Item As MSComctlLib.ListItem)
    
    Const cItemInfoTop As Integer = 1000
    
    If (Item Is Nothing) Then Exit Sub
    
    If (Item.Tag = cIsDrive) Then
        
        With colFolders(Item.key)
        
            frmServices.lblObjectName.Caption = .VolumeName & " (" & .DriveLetter & ":)"
            frmServices.lblObjectName.tOp = frmServices.lblParentName.tOp + frmServices.lblParentName.Height + cItemInfoTop
            frmServices.lblObjectName.Visible = True
            
            frmServices.lblObjectType.Caption = "Local Disk"
            frmServices.lblObjectType.tOp = frmServices.lblObjectName.tOp + frmServices.lblObjectName.Height + 20
            frmServices.lblObjectType.Visible = True
            
            frmServices.lblDescriptionHead.Caption = "File System" & vbCrLf & vbCrLf & _
                                            "Capacity   " & vbCrLf & _
                                            "Used       " & vbCrLf & _
                                            "Free       "
            
            frmServices.lblDescriptionHead.tOp = frmServices.lblObjectType.tOp + frmServices.lblObjectType.Height + 200
            frmServices.lblDescriptionHead.Visible = True
            
            frmServices.lblDescription.Caption = ": " & .FileSystem & vbCrLf & vbCrLf & _
                                        ": " & ConvertUnit(.TotalSize, enmMB, enmBest, True) & vbCrLf & _
                                        ": " & ConvertUnit((((.TotalSize * 1024) * 1024) - .FreeSpace), enmByte, enmBest, True) & vbCrLf & _
                                        ": " & ConvertUnit(.FreeSpace, enmByte, enmBest, True)

            frmServices.lblDescription.tOp = frmServices.lblDescriptionHead.tOp
            frmServices.lblDescription.left = frmServices.lblDescriptionHead.left + frmServices.lblDescriptionHead.Width
            frmServices.lblDescription.Visible = True
            
            frmServices.chartHDD.Width = frmServices.picObjectInfo.ScaleWidth * 1.1
            frmServices.chartHDD.Height = frmServices.chartHDD.Width
            
            frmServices.chartHDD.Column = 1
            frmServices.chartHDD.Data = (.TotalSize - .FreeSpace)

            frmServices.chartHDD.Column = 2
            frmServices.chartHDD.Data = .FreeSpace
            
            frmServices.chartHDD.tOp = frmServices.lblDescriptionHead.tOp + frmServices.lblDescriptionHead.Height '+ 200
            frmServices.chartHDD.left = frmServices.lblDescriptionHead.left - 600
            frmServices.chartHDD.Visible = True
            
            frmServices.chartHDD.Legend.Location.LocationType = VtChLocationTypeBottomRight
            
        End With
        
    ElseIf (Item.Tag = cIsFolder) Then
        
        With colFolders(Item.key)
            frmServices.lblObjectName.Caption = .Name
            frmServices.lblObjectName.tOp = frmServices.lblParentName.tOp + frmServices.lblParentName.Height + cItemInfoTop
            frmServices.lblObjectName.Visible = True
        
            frmServices.lblObjectType.Caption = "File Folder"
            frmServices.lblObjectType.tOp = frmServices.lblObjectName.tOp + frmServices.lblObjectName.Height + 20
            frmServices.lblObjectType.Visible = True
        
            frmServices.lblDescriptionHead.Caption = "Subfolders " & vbCrLf & _
                                                "Files      " & vbCrLf & _
                                                "Size       " & vbCrLf & _
                                                "Attributes "
            
            frmServices.lblDescriptionHead.tOp = frmServices.lblObjectType.tOp + frmServices.lblObjectType.Height + 200
            frmServices.lblDescriptionHead.Visible = True
            
            frmServices.lblDescription.Caption = ": " & .TotalSubFolders & vbCrLf & _
                                            ": " & .TotalFiles & vbCrLf & _
                                            ": " & ConvertUnit(.Size, enmByte, enmBest, True) & vbCrLf & _
                                            ": " & defineAttributes(.Attributes)
            
            frmServices.lblDescription.tOp = frmServices.lblDescriptionHead.tOp
            frmServices.lblDescription.left = frmServices.lblDescriptionHead.left + frmServices.lblDescriptionHead.Width
            frmServices.lblDescription.Visible = True
        
        End With
    
    Else
        
        Dim vParentCode As String
        Dim vFileCode As String
        Dim vSubInfo As String
        Dim vExtPos As Byte
        
        vParentCode = left(Item.key, InStrRev(Item.key, "\") - 1)
        vFileCode = Right(Item.key, Len(Item.key) - InStrRev(Item.key, "\"))
        vExtPos = InStrRev(vFileCode, ".")
        
        If (vExtPos > 0) Then
            vSubInfo = UCase(Right(vFileCode, Len(vFileCode) - vExtPos)) & " File"
        Else
            vSubInfo = "Unknown Type"
        End If
        
        With colFolders(vParentCode).Files(vFileCode)
            frmServices.lblObjectName.Caption = .Name
            frmServices.lblObjectName.tOp = frmServices.lblParentName.tOp + frmServices.lblParentName.Height + cItemInfoTop
            frmServices.lblObjectName.Visible = True
        
            frmServices.lblObjectType.Caption = vSubInfo
            frmServices.lblObjectType.tOp = frmServices.lblObjectName.tOp + frmServices.lblObjectName.Height + 20
            frmServices.lblObjectType.Visible = True
        
            frmServices.lblDescriptionHead.Caption = "Size       " & vbCrLf & _
                                                "Attributes "
            
            frmServices.lblDescriptionHead.tOp = frmServices.lblObjectType.tOp + frmServices.lblObjectType.Height + 200
            frmServices.lblDescriptionHead.Visible = True

            frmServices.lblDescription.Caption = ": " & ConvertUnit(.Size, enmByte, enmBest, True) & vbCrLf & _
                                            ": " & defineAttributes(.Attributes)
            
            frmServices.lblDescription.tOp = frmServices.lblDescriptionHead.tOp
            frmServices.lblDescription.left = frmServices.lblDescriptionHead.left + frmServices.lblDescriptionHead.Width
            frmServices.lblDescription.Visible = True
        
        End With
    End If

End Sub

Public Sub hideItemInfo()
    
    frmServices.lblObjectName.Visible = False
    frmServices.lblObjectType.Visible = False
    frmServices.lblDescriptionHead.Visible = False
    frmServices.lblDescription.Visible = False
    frmServices.chartHDD.Visible = False
    
End Sub

Public Sub hideParentInfo()
    frmServices.imgObjectIcon.Visible = False
    frmServices.lblParentName.Visible = False
End Sub

Private Function ConvertUnit(Value As Variant, UnitOfValue As enmConversionUnit, Optional ConvertTo As enmConversionUnit = enmBest, Optional SuffixUnit As Boolean = True) As String
    
    Dim i As enmConversionUnit
    
    Select Case (UnitOfValue)
        Case enmBest
            ConvertUnit = ConvertUnit(Value, enmByte, ConvertTo, SuffixUnit)
    End Select
    
    For i = UnitOfValue To ConvertTo
        
        If (ConvertTo = enmBest And Value < 1024) Then Exit For
        
        Value = Value / 1024
        
        If (i = ConvertTo) Then Exit For
        
    Next i
    
    ConvertUnit = Value
    
    Dim vPos As Byte
    
    vPos = InStr(1, ConvertUnit, ".")
    If (vPos > 0) Then ConvertUnit = left(ConvertUnit, vPos + 2)
    
    If (i = enmBest) Then i = enmGB
    
    Select Case (i)
        Case enmByte
            ConvertUnit = ConvertUnit & " Bytes"
        Case enmKB
            ConvertUnit = ConvertUnit & " KBytes"
        Case enmMB
            ConvertUnit = ConvertUnit & " MBytes"
        Case enmGB
            ConvertUnit = ConvertUnit & " GBytes"
    End Select
        
End Function

Private Function defineAttributes(Value As Integer) As String

    If (Value = 0) Then
        defineAttributes = "(normal)"
    Else
        
        defineAttributes = vbNullString
        'Compressed File
        If (Value >= 128) Then
            Value = Value - 128
'            defineAttributes = defineAttributes  & "c"
'        Else
'            defineAttributes = defineAttributes  & "-"
        End If
        
        'Link or Shortcut
        If (Value >= 64) Then
            Value = Value - 64
'            defineAttributes = defineAttributes & "l"
'        Else
'            defineAttributes = defineAttributes & "-"
        End If
        
        'Archived
        If (Value >= 32) Then
            Value = Value - 32
            defineAttributes = defineAttributes & "a"
        Else
            defineAttributes = defineAttributes & "-"
        End If
        
        'Directory
        If (Value >= 16) Then
            Value = Value - 16
            defineAttributes = "d" & defineAttributes
        Else
            defineAttributes = "-" & defineAttributes
        End If
        
        'Volume of disk drive
        If (Value >= 8) Then
            Value = Value - 8
'            defineAttributes = defineAttributes & "v"
'        Else
'            defineAttributes = defineAttributes & "-"
        End If
        
        'System
        If (Value >= 4) Then
            Value = Value - 4
            defineAttributes = defineAttributes & "s"
        Else
            defineAttributes = defineAttributes & "-"
        End If
        
        'Hidden
        If (Value >= 2) Then
            Value = Value - 2
            defineAttributes = defineAttributes & "h"
        Else
            defineAttributes = defineAttributes & "-"
        End If
        
        'Read only
        If (Value >= 1) Then
            Value = Value - 1
            defineAttributes = defineAttributes & "r"
        Else
            defineAttributes = defineAttributes & "-"
        End If
        
    End If
    
End Function

Public Sub setToolBarExplorer()
    Set frmMDI.tlbNavigation.ImageList = frmMDI.imgList16.Object
    
    frmMDI.tlbNavigation.Buttons("keyIconView").Image = frmMDI.imgList16.ListImages("imgIconView").Index
    frmMDI.tlbNavigation.Buttons("keyListView").Image = frmMDI.imgList16.ListImages("imgListView").Index
    frmMDI.tlbNavigation.Buttons("keySmallIconView").Image = frmMDI.imgList16.ListImages("imgSmallIconView").Index
    frmMDI.tlbNavigation.Buttons("keyMovePrev").Image = frmMDI.imgList16.ListImages("imgMoveBack").Index
    frmMDI.tlbNavigation.Buttons("keyMoveNext").Image = frmMDI.imgList16.ListImages("imgMoveNext").Index
    frmMDI.tlbNavigation.Buttons("keyMoveUp").Image = frmMDI.imgList16.ListImages("imgMoveUpLevel").Index
    frmMDI.tlbNavigation.Buttons("keyCut").Image = frmMDI.imgList16.ListImages("imgCut").Index
    frmMDI.tlbNavigation.Buttons("keyCopy").Image = frmMDI.imgList16.ListImages("imgCopy").Index
    frmMDI.tlbNavigation.Buttons("keyPaste").Image = frmMDI.imgList16.ListImages("imgPaste").Index
    frmMDI.tlbNavigation.Buttons("keyDelete").Image = frmMDI.imgList16.ListImages("imgDelete").Index
    frmMDI.tlbNavigation.Buttons("keySearch").Image = frmMDI.imgList16.ListImages("imgSearch").Index
    frmMDI.tlbNavigation.Buttons("keyFolders").Image = frmMDI.imgList16.ListImages("imgFolders").Index
    
'    frmservices.tlbNavigation.Buttons("keyRefresh").Caption = vbNullString
'    frmservices.tlbNavigation.Buttons("keyIconView").Caption = vbNullString
'    frmservices.tlbNavigation.Buttons("keyListView").Caption = vbNullString
'    frmservices.tlbNavigation.Buttons("keySmallIconView").Caption = vbNullString
'    frmservices.tlbNavigation.Buttons("keyMoveUp").Caption = vbNullString
'    frmservices.tlbNavigation.Buttons("keyCut").Caption = vbNullString
'    frmservices.tlbNavigation.Buttons("keyCopy").Caption = vbNullString
'    frmservices.tlbNavigation.Buttons("keyPaste").Caption = vbNullString
'    frmservices.tlbNavigation.Buttons("keyDelete").Caption = vbNullString
    
'    frmServices.tlbNavigation.Buttons("keyMovePrev").Caption = vbNullString
'    frmServices.tlbNavigation.Buttons("keyMoveNext").Caption = vbNullString
'    frmServices.tlbNavigation.Buttons("keyFolders").Caption = vbNullString
'    frmServices.tlbNavigation.Buttons("keySearch").Caption = vbNullString
End Sub

Public Sub moveBack()
    Dim oTreeNode As MSComctlLib.Node
    
    If (colMoves.Count > 1) Then
        
        Call colMoves.Remove(colMoves.Count)
        
        Set oTreeNode = colMoves(colMoves.Count)
        
        Call colMoves.Remove(colMoves.Count)
        
        Call buildExplorerListView(oTreeNode)
        
        If (colMoves.Count <= 1) Then
            frmMDI.tlbNavigation.Buttons("keyMovePrev").Enabled = False
        End If

        If (oTreeNode.Parent Is Nothing) Then
            frmMDI.tlbNavigation.Buttons("keyMoveUp").Enabled = False
        End If

        Set oTreeNode = Nothing
            
    End If

End Sub

Public Sub moveUp()
    Dim oTreeNode As MSComctlLib.Node
    
    If (colMoves.Count >= 1) Then
        
        Set oTreeNode = colMoves(colMoves.Count)
        
        If (Not oTreeNode.Parent Is Nothing) Then
            Call buildExplorerListView(oTreeNode.Parent)
        End If
        
        If (colMoves.Count <= 1) Then
            frmMDI.tlbNavigation.Buttons("keyMovePrev").Enabled = False
            frmMDI.tlbNavigation.Buttons("keyMoveUp").Enabled = False
        End If

        If (oTreeNode.Parent Is Nothing) Then
            frmMDI.tlbNavigation.Buttons("keyMoveUp").Enabled = False
        End If

        Set oTreeNode = Nothing
    
    End If

End Sub


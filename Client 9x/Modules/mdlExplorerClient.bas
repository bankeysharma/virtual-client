Attribute VB_Name = "mdlExplorerClient"
Option Explicit
Option Private Module

Private oFSO As New Scripting.FileSystemObject

Private Const cDriveDelimiter As String = ">"
Private Const cNextLevelDelimiter As String = "\"
Private Const cPrevLevelDelimiter As String = "/"
Private Const cFolderDelimiter As String = "^"
Private Const cFileDelimiter As String = "*"
Private Const cSectionDelimiter As String = "|"
Private Const cAttributeDelimiter As String = ":"

Public Function ScanFS(Path As String, Optional DepthLevel As Integer = 0) As String
    On Error Resume Next
    
    If (Path = vbNullString) Then
        Dim oAllDrives As Scripting.Drives
        Dim oDrive As Scripting.Drive
        
        Set oAllDrives = oFSO.Drives
        
        For Each oDrive In oAllDrives
            
            '********************************
            '   Info sequence of Drive
            '       DriveLetter
            '       DriveType
            '       FileSystem
            '       FreeSpace
            '       TotalSize
            '       VolumeName
            '********************************
            With oDrive
                If (.DriveType = Fixed) Then _
                ScanFS = ScanFS & cDriveDelimiter & _
                            .DriveLetter & cAttributeDelimiter & _
                            .DriveType & cAttributeDelimiter & _
                            .FileSystem & cAttributeDelimiter & _
                            .FreeSpace & cAttributeDelimiter & _
                            ((.TotalSize / 1024) / 1024) & cAttributeDelimiter & _
                            .VolumeName & _
                            cNextLevelDelimiter & _
                            ScanFS(.RootFolder, DepthLevel)
            End With
        
        Next oDrive
    
    Else
        
        Dim oAllFolders As Scripting.Folders
        Dim oFolder As Scripting.Folder
        Dim oAllFiles As Scripting.Files
        Dim oFile As Scripting.File
        
        If (DepthLevel = -1) Then Exit Function
        
        Set oAllFolders = oFSO.GetFolder(Path).SubFolders
        
        For Each oFolder In oAllFolders
            '********************************
            '   Info sequence of Folder
            '       Name
            '       Attributes
            '       TotalSubFolders
            '       TotalFiles
            '       Size
            '********************************
            
            With oFolder
                DoEvents
                ScanFS = ScanFS & _
                            .Name & cAttributeDelimiter & _
                            .Attributes & cAttributeDelimiter & _
                            .SubFolders.Count & cAttributeDelimiter & _
                            .Files.Count & cAttributeDelimiter & _
                            .Size & _
                            cFolderDelimiter & _
                            IIf(DepthLevel - 1 >= 0, cNextLevelDelimiter, vbNullString) & _
                            ScanFS(oFolder.Path, DepthLevel - 1)
            
            End With
                        
        Next oFolder
        
        Set oAllFiles = oFSO.GetFolder(Path).Files
        
        '********************************
        '   Info sequence of File
        '       Name
        '       Attributes
        '       Size
        '********************************
        For Each oFile In oAllFiles
            DoEvents
            With oFile
                ScanFS = ScanFS & _
                            .Name & cAttributeDelimiter & _
                            .Attributes & cAttributeDelimiter & _
                            .Size & _
                            cFileDelimiter
            End With
            
        Next oFile
        
        ScanFS = ScanFS & cPrevLevelDelimiter

        
    End If

End Function


Attribute VB_Name = "ExtractFileData"
Sub getFiles()

    Excel.Application.ScreenUpdating = True
    
    Dim SourceFolderName As String
    Dim drivepath As String

    ' Clear the previous files
    Call clear
    
    If Sheets("FileInfo").Cells(2, 1) = "" Then
        MsgBox ("Enter a path in the yellow box first!")
        End
    End If

    ' Which folder to check
    drivepath = Sheets("FileInfo").Cells(2, 1)
    
    ' make sure it ends in a slash
    lastChar = Mid(drivepath, Len(drivepath), 1)
    If lastChar = "/" Or lastChar = "\" Then
        ' Already has a slash
    Else
        ' Need to add a slash
        drivepath = drivepath & "\"
    End If
    
    SourceFolderName = drivepath
    
    ' Make sure the folder exists
    If Not FolderExists(SourceFolderName) Then
        MsgBox ("Folder does not exist" & vbNewLine & vbNewLine & "Ensure the correct FOLDER path in the yellow box")
        End
    Else
        ' Add a hyperlink to the folder
        Sheets("FileInfo").Cells(3, 1).Formula = "=HYPERLINK(""" & SourceFolderName & """,""" & "Open Folder" & """)"
    End If

    ' read all files
    Excel.Application.ScreenUpdating = True
    Sheets("FileInfo").Cells(6, 1) = "Reading files, please wait..."
    Application.Wait (Now + TimeValue("00:00:01"))
    Excel.Application.ScreenUpdating = False

    Call readpath(SourceFolderName)

    Excel.Application.ScreenUpdating = True
    Sheets("FileInfo").Cells(6, 1) = "Finished"
    Application.Wait (Now + TimeValue("00:00:01"))
    Excel.Application.ScreenUpdating = False

End Sub

Public Function FolderExists(strFolderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function

Private Sub readpath(SourceFolderName As String)

    ' read all the files on the j drive for the current project
    On Error Resume Next
    
    Dim FSO As New FileSystemObject, SourceFolder As Folder, Subfolder As Folder, FileItem As File
    
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    
    For Each FileItem In SourceFolder.Files
            
        ' Cycle through each file
            
        ' separate file name & extension
        For g = Len(FileItem.Name) To 1 Step -1
        
            If Mid(FileItem.Name, g, 1) = "." Then
                trimmedName = Mid(FileItem.Name, 1, g - 1)
                fileExtension = Mid(FileItem.Name, g + 1, Len(FileItem.Name) - g)
                Exit For
            End If
            
        Next g
        
        ' Check wheter the file is to be included
        includeFile = False
        
        ' Exclude files with ~ in their name
        If InStr(1, LCase(trimmedName), LCase("~")) Then
        
            ' Don't include
            
        Else
        
            ' No ~ , so continue checking
            
'            ' Check if it's the correct type
'            For fileTypeRow = 2 To 7
'
'                includedFileType = Sheets("Stages").Cells(fileTypeRow, 3)
'
'                If LCase(fileExtension) = LCase(includedFileType) Then

                    includeFile = True
'                    Exit For
'
'                End If
'
'            Next fileTypeRow
        
        End If
        
        ' Only record the file if it is to be included
        If includeFile = True Then
        
            ' Find a blank row
            For i = 9 To 10000
                If Sheets("FileInfo").Cells(i, 1) = "" Then
                    Exit For
                End If
            Next i
                    
            ' Output the file info
            Sheets("FileInfo").Cells(i, 2) = fileExtension
            Sheets("FileInfo").Cells(i, 3) = FileItem.Path
            Sheets("FileInfo").Cells(i, 4) = FileItem.DateCreated
            Sheets("FileInfo").Cells(i, 5) = FileItem.DateLastModified
            Sheets("FileInfo").Cells(i, 1) = trimmedName
            Sheets("FileInfo").Cells(i, 6) = FileItem.Size / 1000000    ' convert to MB from bytes
            
            ' trim the filepath down to just the path
            ' Sheets("FileInfo").Cells(i, 3) = Mid(Sheets("FileInfo").Cells(i, 3), 1, Len(Sheets("FileInfo").Cells(i, 3)) - Len(FileItem.Name))
            filePath = Sheets("FileInfo").Cells(i, 3)
            Sheets("FileInfo").Cells(i, 3).Formula = "=HYPERLINK(""" & filePath & """,""" & "Open" & """)"
        
        End If
    
        
    Next FileItem
    
    '  IncludeSubfolders Then
    For Each Subfolder In SourceFolder.SubFolders
        readpath Subfolder.Path
    Next Subfolder
    
    
    Set FileItem = Nothing
    Set SourceFolder = Nothing

End Sub

Sub clear()
    '
    ' clear Macro
    '
    Sheets("FileInfo").Cells(3, 1) = "" ' Clear the open folder link
    Sheets("FileInfo").Cells(6, 1) = "Enter a folder path above and press the button..." ' Reset the status
    Sheets("FileInfo").Select
    
      Range("A9:G10000").Select
       Selection.ClearContents
            Sheets("FileInfo").Range("A11").Select
    
End Sub

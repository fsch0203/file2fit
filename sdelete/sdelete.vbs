Dim Arg, sfScript, spfeSources, cmd, oShell, _
	spfeScript, spfeSDelete, aFiles, aFolders
Set Arg = WScript.Arguments
If Arg.Count>0 Then
	spfeSources = Arg(0)
	Set oShell = CreateObject ("WScript.Shell")
	spfeScript = Wscript.ScriptFullName 'path + filename of this vbs-script
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set oFile = fso.GetFile(spfeScript) 'get filename of this vbs
	sfScript = fso.GetParentFolderName(oFile) 
	' spfeSDelete = sfScript & "\sdelete64.exe"
	spfeSDelete = sfScript & "\sdelete\sdelete.exe"
    Call addSelectedFiles()

    Dim warning, item, okay
    warning = "The following " & UBound(aFolders) + 1 + UBound(aFiles) + 1 & " folder(s) and/or file(s) will be permanently deleted: " & vbnewline & vbnewline
    for each item in aFolders
        warning = warning & item & vbnewline
    next
    for each item in aFiles
        warning = warning & item & vbnewline
    next
    resp = msgbox (warning, 49, "Warning")

    If resp = 1 Then
        for each file in aFiles
            cmd = spfeSDelete & " -p 1 " & CHR(34) & file & CHR(34)
            oShell.Run cmd, 0, True
        next
        for each folder in aFolders
            cmd = spfeSDelete & " -p 1 " & CHR(34) & folder & CHR(34)
            oShell.Run cmd, 0, True
        next
    End If
    
    Set fso = Nothing
    Set oShell = Nothing
	set Arg = Nothing
Else
    msgbox "Nothing selected"
End If

Sub addSelectedFiles()
    aFiles = Array() 'array for all the files to be converted
    aFolders = Array() 'array for all folders to be made on target
    Dim spfeSource, spfeTarget, i
    If fso.FileExists(spfeSources) Then
        Set f = fso.OpenTextFile(spfeSources) 'tmp file with filepaths of files to be encrypted 
        Do Until f.AtEndOfStream
            spfeSource = LCase(f.ReadLine)
            If Right(spfeSource, 1) = "\" then 'folder
                ' if recursive.checked = true Then 
                if true Then 
                    Call getFiles(spfeSource) 
                end if
            Else 'file
                spfeTarget = Replace(spfeSource, spSource, spTarget)
                ' Call pushItem(aFiles, spfeSource & "|" & spfeTarget)
                Call pushItem(aFiles, spfeSource)
            End If
        Loop
        f.Close
        ' progress.innerhtml = "To do: " & UBound(aFiles) + 1 & " files"
        for i = 0 to UBound(aFolders)
            aFolders(i) = Replace(aFolders(i), spSource, spTarget)
        next
        ' showArray(aFiles)
        ' showArray(aFolders)
    Else
        ' progress.innerhtml = "No files or folders selected"
    End If
End Sub

Sub pushItem(arr, val) 'push item to array
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
End Sub

Sub showArray(array)
    Dim list, item
    list = "Total # items in array: " & UBound(array) + 1 & vbnewline
    for each item in array
        list = list & item & vbnewline
    next
    msgbox list
End Sub

Function getExtension(sfeFile)
    Dim ext
    ext = Right(sfeFile, 5)
    If Left(ext, 1) = "." Then 'e.g. flac
        ext = Right(ext, 4)
    ElseIf Mid(ext, 2, 1) = "." Then 'e.g. mp3
        ext = Right(ext, 3)
    ElseIf Mid(ext, 3, 1) = "." Then 'e.g. md
        ext = Right(ext, 2)
    Else
        ext = ""
    End If
    getExtension = ext
End Function

Sub getFiles(sFolder) 'put all files from folder and subfolders in aFiles
    Dim spfeSource, spfeTarget,oFile
    Set oFolder = fso.GetFolder(sFolder)
    Set colFiles = oFolder.Files
    For Each oFile in colFiles
        spfeSource = sFolder & oFile.Name
        Call pushItem(aFiles, spfeSource)
    Next
    getSubfolders fso.GetFolder(sFolder)
End Sub

Sub getSubFolders(Folder) 'part of recursive procedure
    Call pushItem(aFolders, LCase(Folder))
    For Each Subfolder in Folder.SubFolders
        Set oFolder = fso.GetFolder(Subfolder.Path)
        Set colFiles = oFolder.Files
        For Each oFile in colFiles
            if checkExt(getExtension(oFile.Name)) then
                spfeSource = LCase(Subfolder.Path) & "\" & oFile.Name
                Call pushItem(aFiles, spfeSource)
            end if
        Next
        getSubFolders Subfolder
    Next
End Sub


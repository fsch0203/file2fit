Dim spfePdftk, spfeSources, spTarget
Dim aFiles, fso
' WScript.Echo "init"
Set fso = CreateObject("Scripting.FileSystemObject")

call Init()
If addSelectedFiles() > 0 Then
    Call handlePdf()
Else
    msgbox "Select one or more pdf-files", vbCritical, "Error"
End If

Sub Init()
    Dim objFile, strPathFileScript
    strPathFileScript = Wscript.ScriptFullName 'path + filename of this vbs-script
    ' WScript.Echo strPathFileScript
    Set objFile = fso.GetFile(strPathFileScript) 'get filename of this vbs
    ' WScript.Echo objFile.Name
    strFolder = fso.GetParentFolderName(objFile) 
    spfePdftk = strFolder & "\pdftk\pdftk.exe"
    Set Arg = WScript.Arguments
    ' WScript.Echo Arg.Count
    If Arg.Count>0 Then
        spfeSources = Arg(0)
        spTarget = Arg(1)
    End If
End Sub

Function addSelectedFiles()
    aFiles = Array() 'array for all the files to be converted
    Dim spfeSource, n
    n = 0
    If fso.FileExists(spfeSources) Then
        Set f = fso.OpenTextFile(spfeSources) 'tmp file with filepaths of files to be encrypted 
        Do Until f.AtEndOfStream
            spfeSource = LCase(f.ReadLine)
            if Right(spfeSource, 3) = "pdf" then
                Call pushItem(aFiles, spfeSource)
                n = n + 1
            end if
        Loop
        f.Close
    End If
    addSelectedFiles = n
End Function

Sub handlePdf()
    Dim file, cmd, strFileSource
    If Ubound(aFiles) > 0 Then 'join
        cmd = spfePdftk
        for each file in aFiles
            cmd = cmd & " " & CHR(34) & file & CHR(34)
        next
        cmd  = cmd & " cat output " & CHR(34) & spTarget & "combined.pdf" & CHR(34)
        ' WScript.Echo cmd
    Else 'split
        Set objFile = fso.GetFile(aFiles(0)) 'get filename
        strFileSource = Left(objFile.Name, Len(objFile.Name)-4)
        cmd = spfePdftk & " " & CHR(34) & aFiles(0) & CHR(34) & " burst output " & _
            CHR(34) & spTarget & strFileSource & "_%04d.pdf" & CHR(34)
        ' WScript.Echo cmd
    End If
    Dim objShell
    Set objShell = CreateObject ("WScript.Shell")
    objShell.Run cmd, 0, True

End Sub

Sub pushItem(arr, val) 'push item to array
    ReDim Preserve arr(UBound(arr) + 1)
    arr(UBound(arr)) = val
End Sub



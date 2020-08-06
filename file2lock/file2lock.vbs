Dim Arg, spfeSource, spTarget, spfeSources, cmd, password, oShell, _
	spfeScript, spfe7zip, oTarget, sfeTarget, fso, oFile, sfScript, inputFiles, spSource
Set Arg = WScript.Arguments
If Arg.Count > 2 Then
    spSource = Arg(0)
	spTarget = Arg(1)
	spfeSources = Arg(2)
	password=InputBox("Give password if you want encryption","AES encryption")
	If Len(password) > 0 then
        password = " -p" & password
    Else
        password = ""
    End If
    Set oShell = CreateObject ("WScript.Shell")
    spfeScript = Wscript.ScriptFullName 'path + filename of this vbs-script
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.GetFile(spfeScript) 'get filename of this vbs
    sfScript = fso.GetParentFolderName(oFile) 
    spfe7zip = sfScript & "\7zip\7za.exe"
    Set f = fso.OpenTextFile(spfeSources) 'tmp file with filepaths of files to be encrypted 
    spfeSource = f.ReadLine
    inputFiles = CHR(34) & spfeSource & CHR(34)
    sfeTarget = Mid(spfeSource, Len(spSource)+1)
    If Right(sfeTarget, 1) = "\" then 'folder
        sfeTarget = Left(sfeTarget, Len(sfeTarget) - 1)
    End If
    ' msgbox sfeTarget
    Dim n : n = 1
    Do Until f.AtEndOfStream
        inputFiles = inputFiles & " " & CHR(34) & f.ReadLine & CHR(34)
        n = n + 1
    Loop
    If n > 1 Then
        sfeTarget = InputBox("Give name of 7z-file", "Input", sfeTarget)
    End If
    f.Close
    cmd = spfe7zip & " a " & CHR(34) & spTarget & sfeTarget & ".7z" & CHR(34) & password & " -mhe " & inputFiles
    oShell.Run cmd, 0, True
    Set oShell = Nothing
	set Arg = Nothing
End If
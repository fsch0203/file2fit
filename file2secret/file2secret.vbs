Dim Arg, spfeSource, spTarget, spfeSources, cmd, password, oShell, _
	spfeScript, spfeMySecret
Set Arg = WScript.Arguments
If Arg.Count>0 Then
	spfeSources = Arg(0)
	spTarget = Arg(1)
	password=InputBox("Give password:","MySecret")
	if Len(password) > 0 then
		Set oShell = CreateObject ("WScript.Shell")
		spfeScript = Wscript.ScriptFullName 'path + filename of this vbs-script
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.GetFile(spfeScript) 'get filename of this vbs
		strFolder = objFSO.GetParentFolderName(objFile) 
		spfeMySecret = strFolder & "\mysecret\MySecret.exe"
		' WScript.Echo spfeMySecret
		Set f = objFSO.OpenTextFile(spfeSources) 'tmp file with filepaths of files to be encrypted 
		Do Until f.AtEndOfStream
			spfeSource = f.ReadLine
            If Right(spfeSource, 1) = "\" then 'folder
                msgbox spfeSource & " is not a file", 16, "Error"
            Else
                Set objTarget = objFSO.GetFile(spfeSource)
                strFileTarget = objFSO.GetFileName(objTarget) 'isolate filename
                If Right(strFileTarget,3) = "mys" Then
                    strFileTarget = Left(strFileTarget, Len(strFileTarget)-4)
                    cmd = spfeMySecret & " -d -p " & password & " " & CHR(34) & spfeSource & CHR(34) & _
                        " " & CHR(34) & spTarget & strFileTarget & CHR(34)
                    oShell.Run cmd, 0, True
                Else
                    cmd = spfeMySecret & " -e -p " & password & " " & CHR(34) & spfeSource & CHR(34) & _
                        " " & CHR(34) & spTarget & strFileTarget & ".mys" & CHR(34)  
                    ' WScript.Echo cmd
                    oShell.Run cmd, 0, True
                End If
            End If
		Loop
		f.Close
	end if
	set Arg = Nothing
End If
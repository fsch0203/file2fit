Dim spScript, spfeSources, spSource, spTarget
Dim spfeIrfan, spfeCWebp, spfeDWebp, spfePandoc, spfeLame, spfeFlac, spfeOggenc, spfeOggdec
Dim aFiles, aFolders
Dim sCategory
Dim sIrfanSource, sIrfanTarget
sImageFiles = "bmp,gif,ico,jpg,jpeg,png,psd,psp,tga,tif,tiff,wmf,webp"
sMusicFiles = "wav,mp3,flac,ape,ogg"
sDocFiles = "md,html,epub,txt,tex,xml"
sIrfanSource = "bmp,gif,ico,jpg,jpeg,png,psd,psp,tga,tif,tiff,wmf"
sIrfanTarget = "bmp,gif,ico,jpg,png,tif"
sCWebpSource = "jpg,jpeg,png,tif,tiff"
sCWebpTarget = "webp"
sDWebpSource = "webp"
sDWebpTarget = "png"
sPandocSource = "md,html,epub,txt,tex,xml"
sPandocTarget = "md,html,epub,pdf,docx,odt,xml,wiki"
sLameSource = "wav,mp3"
sLameTarget = "wav,mp3"
sFlacSource = "wav,flac"
sFlacTarget = "wav,flac"
sOggencSource = "wav,flac"
sOggencTarget = "ogg"
sOggdecSource = "ogg"
sOggdecTarget = "wav"
sApeSource = "ape,wav"
sApeTarget = "ape,wav"
Set fso = CreateObject("Scripting.FileSystemObject")

Sub window_onload 'will be called when the application loads
    call Init()
    sCategory = getCategory()
    If sCategory <> "" Then
        Call setSelectBox()
        Call addSelectedFiles()
    Else
        selecttitle.innerhtml = "Category of files to convert is unknown. Please select category."
        selectbox.innerhtml = "<select class='selectbox' id=fileformat title='Select category of files to convert' name='sb-tooltip'>" &_ 
            "<option value='music'>music</option>" & _
            "<option value='image'>images</option>" &_
            "<option value='doc'>documents</option>" &_
            "</select>"
        fileformat.focus
    End If
End Sub

Sub window_onunload
    Set fso = Nothing
    Set oShell = Nothing
    Set oTarget = Nothing
    oFolder.Close
    colFiles.Close
    oFolder.Close
  'This method will be called
  'when the application exits
End Sub

Sub keyPress()
    If CStr(window.event.keyCode) = 13 Then 'Enter
        Call Send2()
    ElseIf CStr(window.event.keyCode) = 27 Then 'Escape
        self.Close()
    End If
End Sub

Sub Send2()
    If fileformat.value = "music" OR fileformat.value = "image" OR fileformat.value = "doc" Then 'category is now selected
        sCategory = fileformat.value
        Call setSelectBox()
        recursive.checked = true 'category wasn't kwown, so probably a folder was selected
        Call addSelectedFiles()
    Else
        call makeTargetFolders()
        call convertFiles(fileformat.value)
        self.Close()
    End If
End Sub

Sub Init()
    Set oFile = fso.GetFile(Self.location.pathname) 'get filename of this hta
    spScript = fso.GetParentFolderName(oFile) 'get hta-folder (is main folder)

    Window.ResizeTo 600, 470
    posX = CInt((window.screen.width - document.body.offsetWidth) / 2)
    posY = CInt((window.screen.height - document.body.offsetHeight) / 2)
    If posX < 0 Then posX = 0
    If posY < 0 Then posY = 0
    window.moveTo posX, posY

    if fso.FileExists(spScript & "\arguments.txt") then
        Set a = fso.OpenTextFile(spScript & "\arguments.txt")
        spSource = LCase(a.ReadLine) 'target folder from TC
        spTarget = LCase(a.ReadLine) 'target folder from TC
        spfeSources = LCase(a.ReadLine) 'tmp-file from TC with all source files
        ' msgbox spSource & "  " & spTarget & " " & spfeSources
        a.Close
    else
        msgbox "Nothing selected"
    end if
    if fso.FileExists(spScript & "\arguments.txt") then
        fso.DeleteFile(spScript & "\arguments.txt")
    end if

    spfeIrfan = spScript & "\irfan\i_view32.exe"
    spfeCWebp = spScript & "\webp\cwebp.exe"
    spfeDWebp = spScript & "\webp\dwebp.exe"
    spfePandoc = "pandoc.exe"
    spfeLame = spScript & "\lame\lame.exe"
    spfeFlac = spScript & "\flac\flac.exe"
    spfeOggenc = spScript & "\ogg\oggenc.exe"
    spfeOggdec = spScript & "\ogg\oggdec.exe"
    spfeApe = spScript & "\ape\mac.exe"
End Sub

Function getCategory()
    Dim spfeSource, ext, sCategory
    sCategory = ""
    If fso.FileExists(spfeSources) Then
        Set f = fso.OpenTextFile(spfeSources) 'tmp file with filepaths of files to be encrypted 
        Do Until f.AtEndOfStream
            spfeSource = LCase(f.ReadLine)
            If Right(spfeSource, 1) <> "\" then 'file
                ext = getExtension(spfeSource)
                If (Instr(sMusicFiles, ext)) Then
                    sCategory = "music"
                    Exit Do
                Elseif (Instr(sImageFiles, ext)) Then
                    sCategory = "image"
                    Exit Do
                ElseIf (Instr(sDocFiles, ext)) Then
                    sCategory = "doc"
                    Exit Do
                End If
            End If
        Loop
        f.Close
    End If
    getCategory = sCategory
End Function

Function setSelectBox()
    If sCategory = "music" Then 
        selecttitle.innerhtml = "Convert music file from wav, flac, mp3, ogg or to new format:"
        selectbox.innerhtml = "<select class='selectbox' id=fileformat title='Select new music format' name='sb-tooltip'>" &_ 
            "<option value='mp3'>mp3</option>" & _
            "<option value='wav'>wav</option>" &_
            "<option value='flac'>flac</option>" &_
            "<option value='ogg'>ogg</option>" &_
            "<option value='ape'>ape</option>" &_
            "</select>"
    Elseif sCategory = "image" Then
        selecttitle.innerhtml = "Convert image file from bmp, gif, ico, jpg, jpeg, png, psd, psp, tga, tif, tiff, wmf or webp to new format:"
        selectbox.innerhtml = "<select class='selectbox' id=fileformat title='Select new image format' name='sb-tooltip'>" &_ 
            "<option value='jpg'>jpg</option>" & _
            "<option value='png'>png</option>" &_
            "<option value='webp'>webp</option>" &_
            "<option value='bmp'>bmp</option>" &_
            "<option value='tif'>tif</option>" &_
            "<option value='ico'>ico</option>" &_
            "</select>"
    ElseIf sCategory = "doc" Then
        selecttitle.innerhtml = "Convert document file from md, html, epub, txt, tex or xml to new format:"
        selectbox.innerhtml = "<select class='selectbox' id=fileformat title='Select new doc-format' name='sb-tooltip'>" &_ 
            "<option value='md'>md</option>" & _
            "<option value='html'>html</option>" &_
            "<option value='epub'>epub</option>" &_
            "<option value='pdf'>pdf</option>" &_
            "<option value='docx'>docx</option>" &_
            "<option value='odt'>odt</option>" &_
            "<option value='xml'>xml</option>" &_
            "<option value='wiki'>wiki</option>" &_
            "</select>"
    End If
    fileformat.focus
End Function

Sub addSelectedFiles()
    aFiles = Array() 'array for all the files to be converted
    aFolders = Array() 'array for all folders to be made on target
    Dim spfeSource, spfeTarget, i
    If fso.FileExists(spfeSources) Then
        Set f = fso.OpenTextFile(spfeSources) 'tmp file with filepaths of files to be encrypted 
        Do Until f.AtEndOfStream
            spfeSource = LCase(f.ReadLine)
            If Right(spfeSource, 1) = "\" then 'folder
                if recursive.checked = true Then 
                    Call getFiles(spfeSource) 
                end if
            Else 'file
                if checkExt(getExtension(spfeSource)) then
                    spfeTarget = Replace(spfeSource, spSource, spTarget)
                    Call pushItem(aFiles, spfeSource & "|" & spfeTarget)
                end if
            End If
        Loop
        f.Close
        progress.innerhtml = "To do: " & UBound(aFiles) + 1 & " files"
        for i = 0 to UBound(aFolders)
            aFolders(i) = Replace(aFolders(i), spSource, spTarget)
        next
        ' showArray(aFiles)
        ' showArray(aFolders)
    Else
        progress.innerhtml = "No files or folders selected"
    End If
End Sub

Sub getFiles(sFolder) 'put all files from folder and subfolders in aFiles
    Dim spfeSource, spfeTarget,oFile
    Set oFolder = fso.GetFolder(sFolder)
    Set colFiles = oFolder.Files
    For Each oFile in colFiles
        if checkExt(getExtension(oFile.Name)) then
            spfeSource = sFolder & oFile.Name
            spfeTarget = Replace(spfeSource, spSource, spTarget)
            Call pushItem(aFiles, spfeSource & "|" & spfeTarget)
        end if
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
                spfeTarget = Replace(spfeSource, spSource, spTarget)
                Call pushItem(aFiles, spfeSource & "|" & spfeTarget)
            end if
        Next
        getSubFolders Subfolder
    Next
End Sub

'check if file is part of category
function checkExt(ext) 
    Dim check
    check = false
    If sCategory = "music" Then
        If InStr(sMusicFiles, ext) > 0 Then
            check = true
        End If
    ElseIf sCategory = "image" Then
        If InStr(sImageFiles, ext) > 0 Then
            check = true
        End If
    ElseIf sCategory = "doc" Then
        If InStr(sDocFiles, ext) > 0 Then
            check = true
        End If
    End If
    checkExt = check
end function

Sub makeTargetFolders()
    dim folder
    for each folder in aFolders
        If NOT fso.FolderExists(folder) Then
            Set oFolder = fso.CreateFolder(folder)
        End If
    next
End Sub

Sub convertFiles(seTarget)
    Dim item, n, split
    n = 0
    progress.innerhtml = "Progress: " & n & " of " & UBound(aFiles) + 1 & " files"    
    for each item in aFiles
        split = InStr(item, "|")
        src = Left(item, split - 1)
        trg = Right(item, Len(item) - split) 'still got the extension of the source-file
        x = convertFile (src, trg, seTarget) 'seTarget is new extension
        n = n + x
        progress.innerhtml = "Progress: " & n & " of " & UBound(aFiles) + 1 & " files"    
    next
End Sub


Function convertFile(spfeSource, spfeTarget, seTarget)
    Dim success, sfeSource, seSource
    Set oShell = CreateObject ("WScript.Shell")
    Set oTarget = fso.GetFile(spfeSource)
    sfeSource = LCase(fso.GetFileName(oTarget)) 'isolate filename
    seSource = getExtension(sfeSource)
    currentfile.innerhtml = "Current file: " & sfeSource
    succes = 1
    If sCategory = "image" Then
        If InStr(sIrfanSource, seSource) > 0 AND InStr(sIrfanTarget, seTarget) > 0 Then
            cmd = spfeIrfan & " " & CHR(34) & spfeSource & CHR(34) & " /convert=" & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & seTarget & CHR(34)
            oShell.Run cmd, 0, True
        ElseIf (seSource = "jpg" Or seSource = "png" Or seSource = "tif") AND seTarget = "webp" Then
            cmd = spfeCwebp & " -q 80 " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & "webp" & CHR(34)
            ' msgbox cmd
            oShell.Run cmd, 0, True
        ElseIf getExtension(sfeSource) = "webp" AND seTarget = "png"  Then
            cmd = spfeDwebp & " " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & "png" & CHR(34)
            ' msgbox cmd
            oShell.Run cmd, 0, True
        ElseIf getExtension(sfeSource) = "webp" AND seTarget <> "png" AND InStr(sIrfanTarget, seTarget) > 0 Then
            cmd = spfeDwebp & " " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & "png" & CHR(34)
            ' msgbox cmd
            oShell.Run cmd, 0, True
            cmd = spfeIrfan & " " & CHR(34) & getFileNoExt(spfeTarget) & "." & "png" & CHR(34) & " /convert=" & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & seTarget & CHR(34)
            oShell.Run cmd, 0, True
            cmd = getFileNoExt(spfeTarget) & "." & "png"
            if fso.FileExists(cmd) then
                fso.DeleteFile(cmd)
            end if
        ElseIf InStr(sIrfanSource, seSource) > 0  AND seSource <> "jpg" AND seSource <> "png" AND seSource <> "tif" AND seTarget = "webp" Then
            cmd = spfeIrfan & " " & CHR(34) & spfeSource & CHR(34) & " /convert=" & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & "png" & CHR(34)
            oShell.Run cmd, 0, True
            cmd = spfeCwebp & " -q 80 " & CHR(34) & getFileNoExt(spfeTarget) & "." & "png" & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & "webp" & CHR(34)
            ' msgbox cmd
            oShell.Run cmd, 0, True
            cmd = getFileNoExt(spfeTarget) & "." & "png"
            if fso.FileExists(cmd) then
                fso.DeleteFile(cmd)
            end if
        Else
            succes = 0
        End If
    ElseIf sCategory = "doc" Then
        If InStr(sPandocSource, seSource) > 0 AND InStr(sPandocTarget, seTarget) > 0 Then
            cmd = spfePandoc & " -s " & CHR(34) &  spfeSource & CHR(34) & " -o " & CHR(34) & _
            getFileNoExt(spfeTarget) & "." & seTarget & CHR(34)
            ' msgbox cmd
            oShell.Run cmd, 0, True
        End If
    ElseIf sCategory = "music" Then
        If getExtension(sfeSource) = "wav" AND seTarget = "flac" Then
            cmd = spfeFlac & " -6 -s " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            succes = 1
        ElseIf getExtension(sfeSource) = "flac" AND seTarget = "wav" Then
            cmd = spfeFlac & " -s -d " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            succes = 1
        ElseIf getExtension(sfeSource) = "wav" AND seTarget = "mp3" Then
            cmd = spfeLame & " -V2 " & CHR(34) & spfeSource & CHR(34) & " " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            succes = 1
        ElseIf getExtension(sfeSource) = "mp3" AND seTarget = "wav" Then
            cmd = spfeLame & " --decode " & CHR(34) & spfeSource & CHR(34) & " " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            ' msgbox cmd
            oShell.Run cmd, 0, True
            succes = 1
        ElseIf getExtension(sfeSource) = "ogg" AND seTarget = "wav" Then
            cmd = spfeOggdec & " " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            succes = 1
        ElseIf getExtension(sfeSource) = "wav" AND seTarget = "ogg" Then
            cmd = spfeOggenc & " -q6 " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            succes = 1
        ElseIf getExtension(sfeSource) = "flac" AND seTarget = "ogg" Then
            cmd = spfeOggenc & " -q6 " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            succes = 1
        ElseIf getExtension(sfeSource) = "wav" AND seTarget = "ape" Then
            cmd = spfeApe & " " & CHR(34) & spfeSource & CHR(34) & " " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34) & " -c2000"
            oShell.Run cmd, 0, True
            succes = 1
        ElseIf getExtension(sfeSource) = "ape" AND seTarget = "wav" Then
            cmd = spfeApe & " " & CHR(34) & spfeSource & CHR(34) & " " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34) & " -d"
            oShell.Run cmd, 0, True
            succes = 1
        ElseIf getExtension(sfeSource) = "flac" AND seTarget = "mp3" Then
            cmd = spfeFlac & " -s -d " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                "wav" & CHR(34)
            oShell.Run cmd, 0, True
            cmd = spfeLame & " -V2 " & CHR(34) & getFileNoExt(spfeTarget) & "." & _
                "wav" & CHR(34) & " " & CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            cmd = getFileNoExt(spfeTarget) & "." & "wav"
            if fso.FileExists(cmd) then
                fso.DeleteFile(cmd)
            end if
            succes = 1
        ElseIf getExtension(sfeSource) = "mp3" AND seTarget = "flac" Then
            cmd = spfeLame & " --decode " & CHR(34) & spfeSource & CHR(34) & " " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                "wav" & CHR(34)
            oShell.Run cmd, 0, True
            cmd = spfeFlac & " -6 -s " & CHR(34) & getFileNoExt(spfeTarget) & "." & _
                "wav" & CHR(34) & " -o " & CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            cmd = getFileNoExt(spfeTarget) & "." & "wav"
            if fso.FileExists(cmd) then
                fso.DeleteFile(cmd)
            end if
            succes = 1
        ElseIf getExtension(sfeSource) = "mp3" AND seTarget = "ogg" Then
            cmd = spfeLame & " --decode " & CHR(34) & spfeSource & CHR(34) & " " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                "wav" & CHR(34)
            oShell.Run cmd, 0, True
            cmd = spfeOggenc & " -q6 " & CHR(34) & getFileNoExt(spfeTarget) & "." & _
                "wav" & CHR(34) & " -o " & CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            cmd = getFileNoExt(spfeTarget) & "." & "wav"
            if fso.FileExists(cmd) then
                fso.DeleteFile(cmd)
            end if
            succes = 1
        ElseIf getExtension(sfeSource) = "ogg" AND seTarget = "mp3" Then
            cmd = spfeOggdec & " " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                "wav" & CHR(34)
            oShell.Run cmd, 0, True
            cmd = spfeLame & " -V2 " & CHR(34) & getFileNoExt(spfeTarget) & "." & _
                "wav" & CHR(34) & " " & CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            cmd = getFileNoExt(spfeTarget) & "." & "wav"
            if fso.FileExists(cmd) then
                fso.DeleteFile(cmd)
            end if
            succes = 1
        ElseIf getExtension(sfeSource) = "ogg" AND seTarget = "flac" Then
            cmd = spfeOggdec & " " & CHR(34) & spfeSource & CHR(34) & " -o " & _
                CHR(34) & getFileNoExt(spfeTarget) & "." & _
                "wav" & CHR(34)
            oShell.Run cmd, 0, True
            cmd = spfeFlac & " -6 -s " & CHR(34) & getFileNoExt(spfeTarget) & "." & _
                "wav" & CHR(34) & " -o " & CHR(34) & getFileNoExt(spfeTarget) & "." & _
                seTarget & CHR(34)
            oShell.Run cmd, 0, True
            cmd = getFileNoExt(spfeTarget) & "." & "wav"
            if fso.FileExists(cmd) then
                fso.DeleteFile(cmd)
            end if
            succes = 1
        Else
            succes = 0
        End If
    Else
        succes = 0
    End If
    
    convertFile = succes
End Function

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

Function getFileNoExt(sfeFile)
    Dim ext, file
    ext = Right(sfeFile, 5)
    If Left(ext, 1) = "." Then
        file = Left(sfeFile, Len(sfeFile) - 5)
    ElseIf Mid(ext, 2, 1) = "." Then
        file = Left(sfeFile, Len(sfeFile) - 4)
    ElseIf Mid(ext, 3, 1) = "." Then
        file = Left(sfeFile, Len(sfeFile) - 3)
    Else
        file = ""
    End If
    getFileNoExt = file
End Function

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



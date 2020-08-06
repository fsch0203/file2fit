Dim Arg
Set fso = CreateObject("Scripting.FileSystemObject") 
Set oFile = fso.GetFile(Wscript.ScriptFullName) 
sFolder = oFile.ParentFolder 

Set Arg = WScript.Arguments
If Arg.Count>1 Then
    Set a = fso.CreateTextFile(sFolder & "\arguments.txt", True)
    ' WScript.echo Arg(2)
    a.WriteLine(Arg(0)) 'source path
    a.WriteLine(Arg(1)) 'target path
    a.WriteLine(Arg(2)) 'source files
    a.Close
End If

sHtaFilePath = sFolder & "\dialog.hta"
Dim oShell: Set oShell = CreateObject("WScript.Shell")
oShell.Run sHtaFilePath, 1, True

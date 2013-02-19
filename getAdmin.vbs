
Set oShell = CreateObject("Shell.Application")
Set FSO = CreateObject("Scripting.FileSystemObject")
strPath = FSO.getParentFolderName(WScript.ScriptFullName)
If FSO.FileExists(strPath & "\run_spybot.vbs") Then
     oShell.ShellExecute "wscript.exe", _ 
        Chr(34) & strPath & "\run_spybot.vbs" & Chr(34), "", "runas", 1
Else
	MsgBox "Script file run_spybot.vbs not found"
End if

Set objFSO=CreateObject("Scripting.FileSystemObject")
nl = vbNewline

If Not objFSO.FolderExists("C:\spybot_script") Then 
	objFSO.CreateFolder "C:\spybot_script"
End If

outFile="C:\spybot_script\output.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write "test" & nl
objFile.Write "Running Spybot" & nl

Set objShell = CreateObject("WScript.Shell")
q = """"
strCmd = "cmd.exe /K  C:\" & q & "Program Files (x86)" & q & "\Spybot\SpybotSD.exe /autocheck"
objShell.Run(strCmd)
objFile.Write "Spybot Running..." & nl
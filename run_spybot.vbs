
'Setting output file location
outFile="spybot_output.txt"
nl = vbNewline
Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(outFile,True)

objFile.Write "Running Spybot at" &  Date & " " & Time & nl

Set objShell = CreateObject("WScript.Shell")
q = """" 'Insert quotes

'Runs Spybot via Command Line
strCmd = "cmd.exe /C  C:\" & q & "Program Files (x86)" & q & "\Spybot\SpybotSD.exe /autocheck"
objShell.Run strCmd, 0, true
objFile.Write "Spybot Done at " & Date & " " & Time & nl

'Get the spybot log file and find the newest file by comparison. Also saving the second newest file.
Dim fNewest, fSecond
set oFolder=createobject("scripting.filesystemobject").getfolder("C:\ProgramData\Spybot - Search & Destroy\Logs")
For Each aFile In oFolder.Files
    If fNewest = "" Then
        Set fNewest = aFile
    Else
        If fNewest.DateCreated < aFile.DateCreated Then
			Set fSecond = fNewest
            Set fNewest = aFile
        End If
    End If
Next

objFile.Write(fNewest.Name) & nl 'Debug output
objFile.Write(fSecond.Name) & nl 'Debug output

'Reading the log file and finding threats found
Const forReading = 1

Set objRegEx = CreateObject("VBScript.RegExp")
objRegEx.Pattern = "found:" 'Search string

'CStr(fSecond.Name) takes the name of the second newest log file, (.log instead of .txt) and searches for search string within.
objFile.Write("C:\ProgramData\Spybot - Search & Destroy\Logs\" & CStr(fSecond.Name) & nl)

Set WshShell = CreateObject ("Wscript.Shell")
Set logFSO = CreateObject("Scripting.FileSystemObject")
Set logFile = logFSO.OpenTextFile("C:\ProgramData\Spybot - Search & Destroy\Logs\" & CStr(fSecond.Name), forReading)
'Loop that reads each line of the log file to determine if search string is found.
Dim checkFound
checkFound = 0
Do Until logFile.AtEndOfStream
	strSearchString = logFile.ReadLine
	Set colMatches = objRegEx.Execute(strSearchString)
	If colMatches.Count > 0 Then
		For Each strMatch in colMatches
			objFile.Write(strSearchString & nl)
			checkFound = 1
		Next
	End If
Loop
'If search string is not found
If checkFound = 0 Then
	objFile.Write("Nothing Found!" & nl)
End If

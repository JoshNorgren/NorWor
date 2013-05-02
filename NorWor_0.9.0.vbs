'Initial variable declarations'
sHostName = "."
strComputer = "."
Set objShell = WScript.CreateObject("WScript.Shell")
strDesktopFolder = objShell.SpecialFolders("Desktop")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(strDesktopFolder + "\NorWor.html", True)

'Top Table Miscelanny'
CompName = objShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
StartTime =  Date & " " & Time
CompStatus = "OK!"

'HTML formatting nonsense' See comments at very bottom for table template if you want to add new sections to the log file'

TableTR1 = "<td width=120 valign=top style='width:90pt;border-top:solid #666666 1.0pt; border-left:none;border-bottom:solid #666666 1.0pt;border-right:none;  padding:0in 0in 0in 0in'>"
TableTR2 = "<td width=240 valign=top style='width:180pt;border-top:solid #666666 1.0pt;  border-left:none;border-bottom:solid #666666 1.0pt;border-right:none;  padding:0in 0in 0in 0in'>"
TableFormat1 = "<td width=120 valign=top style='width:90pt;border:none;border-bottom:solid #666666 1.0pt;  padding:0in 0in 0in 0in'>"
TableFormat2 = "<td width=240 valign=top style='width:180pt;border:none;border-bottom:solid #666666 1.0pt;  padding:0in 0in 0in 0in'>"
TableStart = "<table class=ListTable2 border=1 cellspacing=0 cellpadding=0 width='100%' style='width:100.0%;border-collapse: collapse;border:none'>"


'END HTML formatting nonsense'

msgbox "Scanning Computer..." & vbNewline & "This may take a few minutes."


'GET HARD DRIVE STATS'
GB = 1024 *1024 * 1024 'Number of bytes in a gigabyte'
kbGB = 1024 * 1024 ' Number of kilobytes in a GB'
HDLOGTEXT = ""
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")


For Each objItem in colItems
If (objItem.VolumeDirty = TRUE) then
	HDDirty = " Warning! This disk needs to be cleaned. You should run disk check."
	CompStatus = "Problems Found!"
	DiskPClass = "<p class=Warning>"
	CompStatusDesc = CompStatusDesc  & "<a href='#Disks'>One of the disks attached to this computer is dirty.</a>"	
Else
	HDDirty = "Disk is clean."
	DiskPClass = "<p class=MsoNormal>"
end if
if (IsNull(objItem.size)) then
	
Else
	if ( (objitem.FreeSpace / objitem.size) < 0.25) then
		CompStatus = "Problems Found!"
		DiskPClass2 = "<p class=Warning>"
		CompStatusDesc = CompStatusDesc  & "<a href='#Disks'>One of the disks attached to this computer is nearing full capacity.</a>"
	else
		DiskPClass2 = "<p class=MsoNormal>"
	end if
	HDLOGTEXT = HDLOGTEXT + _
	"<h2>" & objitem.Description & " - " & objItem.DeviceID & "</h2>" & _
	tablestart   & _
	" <tr> " & TableTR1 & " <h3>File System:</h3></td>" & TableTR2 & "  <p class=MsoNormal>" & objItem.Filesystem & "</p></td> "   & _
	"  " & TableTR1 & "  <h3>Total Hard Drive Size:</h3>  </td>  " & TableTR2  & "  <p class=MsoNormal>" & Int(objItem.size / GB) & " GB</p>  </td> </tr>"   & _
	" <tr>  " & TableFormat1 & "  <h3>Free space:</h3>  </td>  " & TableFormat2  & DiskPClass2 & Int(objItem.FreeSpace / GB) & " GB</p>  </td> "   & _
	" " & TableFormat1 & "  <h3>Caption:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & objItem.Caption & "</p>  </td> </tr>"   & _
	" <tr>  " & TableFormat1 & "  <h3>SerialNumber:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & objItem.VolumeSerialNumber & " </p>  </td> "   & _
	"  " & TableFormat1 & "  <h3>VolumeName:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & objItem.VolumeName & " </p>  </td> </tr>"   & _
	" <tr>  " & TableFormat1 & "  <h3>Disk Clean?</h3>  </td>  " & TableFormat2  & DiskPClass & HDDirty & " </p>  </td> "   & _
	"  " & TableFormat1 & "  <h3>&nbsp;</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>&nbsp;</p>  </td> </tr>"   & _
	"</table>"  
end if

Next


'GET CPU USAGE' 
Samples = 5
CPUtotal = 0
Set RefresherObject = CreateObject("WbemScripting.SWbemRefresher")
Set objWMIService = GetObject("winmgmts:\\" & strCOmputer & "\root\cimv2")

Set ProcessorObjects = _
	RefresherObject.Addenum(objWMIService, "Win32_processor").ObjectSet

RefresherObject.Refresh



For i = 1 to Samples
	RefresherObject.Refresh
	For Each Sampling in ProcessorObjects
		If (IsNull(Sampling.LoadPercentage)) then
			CPUtotal = CPUtotal
		else
			CPUtotal = CPUtotal + Sampling.LoadPercentage
		end if
	Next
Next
CPUavg= CPUtotal / Samples

Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objItem in colItems
     CpuName= objItem.Name
     CpuAddrWidth= objItem.AddressWidth
     CPUDescription= objItem.Description
     CPUManufacturer= objItem.Manufacturer
     CPUProcessorId= objItem.ProcessorId
     CPURevision= objItem.Revision
     CPUSocket= objItem.SocketDesignation
	 CPUL2Size= objItem.L2CacheSize
	 CPUL2Speed= objItem.L2CacheSpeed
	 CPUCurrentClock= objItem.CurrentClockSpeed
	 CPUMaxClock= objItem.MaxClockSpeed
Next





'GET MEM USAGE' 'This part DEF needs to go before we start running any programs'
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")



Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")


For Each objOperatingSystem in colSettings 
    freeMem = objOperatingSystem.FreePhysicalMemory
Next

Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
	
For Each objOperatingSystem in colSettings 
	totalMem = objOperatingSystem.TotalPhysicalMemory / 1024
Next



'get other mem Stats'
Set Memory = GetObject("winmgmts:{impersonationLevel=impersonate}!//" _
 & sHostName & "/root/cimv2:Win32_PerfFormattedData_PerfOS_Memory=@")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory",,48)

For Each objItem in colItems
    AvailableGB = Round(objItem.AvailableBytes / GB,3)
    CommitLimit = Round(objItem.CommitLimit / GB,3)
    CommittedGB = Round(objItem.CommittedBytes / GB,3)

Next

Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_PhysicalMemory",,48)
For Each objItem in colItems
	MEMLOGTEXT = MEMLOGTEXT + _
	"<h2>" & objItem.Description  & " - " & objItem.BankLabel & "</h2>" & _
	tablestart   & _
	" <tr> " & TableTR1 & " <h3>Capacity: </h3></td>" & TableTR2 & "  <p class=MsoNormal>" & Round(objItem.Capacity/ GB,1) & "GB" & "</p></td> "   & _
	"  " & TableTR1 & "  <h3>Speed: </h3>  </td>  " & TableTR2  & "  <p class=MsoNormal>" & objItem.Speed & "</p>  </td> </tr>"   & _
	" <tr>  " & TableFormat1 & "  <h3>Form Factor: </h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & objItem.FormFactor & " </p>  </td> "   & _
	"  " & TableFormat1 & "  <h3>Memory Type: </h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & objItem.MemoryType & " </p>  </td> </tr>"   & _
	" <tr>  " & TableFormat1 & "  <h3>Device Locator:  </h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & objItem.DeviceLocator & " </p>  </td> "   & _
	"  " & TableFormat1 & "  <h3>&nbsp;</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>&nbsp;</p>  </td> </tr>"   & _
	"</table>"  
Next

'Get OS info'
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")


For Each os in oss 											'Most of these are unnecessary but I'm leaving them in for now'
    BootDevice = os.BootDevice
    BuildNumber = os.BuildNumber
    BuildType = os.BuildType
    Caption = os.Caption
    CodeSet = os.CodeSet
    CountryCode = os.CountryCode
   
    EncryptionLevel = os.EncryptionLevel
    dtmConvertedDate.Value = os.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate
   
    OSProductSuite = os.OSProductSuite
    OSType = os.OSType
    Primary = os.Primary
    RegisteredUser = os.RegisteredUser
    SerialNumber = os.SerialNumber
    Version = os.Version
	SPmajor = os.ServicePackMajorVersion
	SPminor = os.ServicePackMinorVersion
	InstallDate = os.InstallDate
Next

'Operating system test' 
' Right now this just prints the OS in the log but this test could be placed earlier and used to determine script behavior based on OS'
strOS = " "
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystems
    strOS= objOperatingSystem.Caption
Next
if instr(strOS,"7") then strOS = "Windows 7"
if instr(strOS,"Vista") then strOS = "Windows Vista"
if instr(strOS,"XP") then strOS = "Windows XP"

IF GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth  = 64 THEN
strOS= strOS & " 64 bit"  
else

strOS= strOS & " 32 bit"
END IF





'Firewall test'


Dim CurrentProfiles
Dim LowerBound
Dim UpperBound
Dim iterate
Dim excludedinterfacesarray

' Profile Type
Const NET_FW_PROFILE2_DOMAIN = 1
Const NET_FW_PROFILE2_PRIVATE = 2
Const NET_FW_PROFILE2_PUBLIC = 4

' Action
Const NET_FW_ACTION_BLOCK = 0
Const NET_FW_ACTION_ALLOW = 1


' Create the FwPolicy2 object.
Dim fwPolicy2
Set fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")

CurrentProfiles = fwPolicy2.CurrentProfileTypes

'// The returned 'CurrentProfiles' bitmask can
'// have more than 1 bit set if multiple profiles 
'// are active or current at the same time

if ( CurrentProfiles AND NET_FW_PROFILE2_DOMAIN ) then
   if fwPolicy2.FirewallEnabled(NET_FW_PROFILE2_DOMAIN) = TRUE then
      DomainProfStatus = "Firewall is ON on domain profile."
   else
      DomainProfStatus = "Firewall is OFF on domain profile."
   end if
else
    DomainProfStatus = "0"
end if

if ( CurrentProfiles AND NET_FW_PROFILE2_PRIVATE ) then
   if fwPolicy2.FirewallEnabled(NET_FW_PROFILE2_PRIVATE) = TRUE then
      PrivateProfStatus = "Firewall is ON on private profile."
   else
      PrivateProfStatus = "Firewall is OFF on private profile."
   end if
else
    PrivateProfStatus = "0"
end if

if ( CurrentProfiles AND NET_FW_PROFILE2_PUBLIC ) then
   if fwPolicy2.FirewallEnabled(NET_FW_PROFILE2_PUBLIC) = TRUE then
      PublicProfStatus = "Firewall is ON on public profile."
   else
      PublicProfStatus = "Firewall is OFF on public profile."
   end if
else
    PublicProfStatus = "0"
end if



'Checks Event log for Unexpected Shutdowns' 'This part is REALLY innefficient and slow if you have a sufficiently large event log'


	'Checks for ALL unexpected shutdowns (AKA blue screens)'
Set colLoggedEvents = objWMIService.ExecQuery _ 
("Select * from Win32_NTLogEvent Where Logfile = 'System' and " _ 
& "EventCode = '6008'")
TotalBlueScreens = colLoggedEvents.Count

Const CONVERT_TO_LOCAL_TIME = True
Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
dtmStartDate.SetVarDate dateadd("m", -1, now)' CONVERT_TO_LOCAL_TIME

	' Checks for unexpected shutdowns in PAST MONTH'
Set colLoggedEvents = objWMIService.ExecQuery _	
("Select * from Win32_NTLogEvent Where Logfile = 'System' and " _
& "EventCode = '6008' and TimeWritten > '" & dtmStartDate & "'")
RecentBlueScreens = colLoggedEvents.Count

'Spybot
sbResult = MsgBox ("Do you wish to run an anti-virus check? This may take several minutes to a couple hours depending on the size of your hard drives.", _
    vbYesNo, "Anti-Virus")

Select Case sbResult
Case vbYes
    
	nl = vbNewline
	q = """" 'Insert quotes

	'Runs Spybot via Command Line

	strCmd = "cmd.exe /C  \Scripts\Spybot\SpybotSD.exe /autocheck /autoclose"
	objShell.Run strCmd, 0, true


	'Get the spybot log file and find the newest file by comparison. Also saving the second newest file.
	Dim fNewest, fSecond
	set oFolder=createobject("scripting.filesystemobject").getfolder("C:\ProgramData\Spybot - Search & Destroy\Logs")
	For Each aFile In oFolder.Files
		If fNewest = "" Then
			Set fNewest = aFile
			Set fSecond = aFile
		Else
			If fNewest.DateCreated < aFile.DateCreated Then
				Set fSecond = fNewest
				Set fNewest = aFile
			End If
		End If
	Next

	'Reading the log file and finding threats found
	Const forReading = 1

	Set objRegEx = CreateObject("VBScript.RegExp")
	objRegEx.Pattern = "found:" 'Search string

	Set WshShell = CreateObject ("Wscript.Shell")
	Set logFSO = CreateObject("Scripting.FileSystemObject")
	Set logFile = logFSO.OpenTextFile("C:\ProgramData\Spybot - Search & Destroy\Logs\" & CStr(fSecond.Name), forReading)
	'Loop that reads each line of the log file to determine if search string is found.
	Dim checkFound, sbFound
	checkFound = 0
	sbFound = ""
	Do Until logFile.AtEndOfStream
		strSearchString = logFile.ReadLine
		Set colMatches = objRegEx.Execute(strSearchString)
		If colMatches.Count > 0 Then
			For Each strMatch in colMatches
				sbFound = sbFound + strSearchString + nl
				checkFound = checkFound + 1
			Next
		End If
	Loop
	'If search string is not found
	If checkFound = 0 Then
		sbFound = "Nothing Found!" & nl
	End If
	ranSpybot = true
Case vbNo
    checkFound = 0
	sbFound = "You did not run an anti-virus check."
	ranSpybot = false
End Select



'General status checks'
if (CheckFound >= 1) then
	CompStatus = "Problems found!"
	CompStatusDesc = CompStatusDesc & "<p class=MsoNormal><a href='#VirusInfo'>Virus Alerts detected!</a></p> " 
	VirusPClass = "<p class=Warning>"
Else
	VirusPClass = "<p class=MsoNormal>"
end if

if ((Round((totalMem - freeMem) / totalMem * 100)) >= 80 ) then
	CompStatus = "Problems Found!"
	CompStatusDesc = CompStatusDesc & "<p class=MsoNormal><a href='#MemInfo'>High Memory use detected!</a></p>" 
	MemPClass = "<p class=Warning>"
Else
	MemPClass = "<p class=MsoNormal>"
end if

if (CPUAVG >= 75) then
	CompStatus = "Problems Found!"
	CompStatusDesc = CompStatusDesc  & "<p class=MsoNormal><a href='#CPUInfo'>High CPU usage detected!</a></p>" 
	CPUPClass = "<p class=Warning>"
Else
	CPUPClass = "<p class=MsoNormal>"
end if

if (DomainProfStatus <> "0" AND PrivateProfStatus <> "0" AND PublicProfStatus <> "0") then
	CompStatus = "Problems Found!"
	CompStatusDesc = CompStatusDesc  & "<p class=MsoNormal><a href='#FWInfo'>No Firewall detected!</a></p>"
end if

if (RecentBlueScreens >= 1) then
	CompStatus = "Problems Found!"
	CompStatusDesc = CompStatusDesc  & "<p class=MsoNormal><a href='#ShutdownInfo'>This computer has suffered from unexpected shutdowns in the past.</a></p>" 
	BluePClass = "<p class=Warning>"
Else
	BluePClass = "<p class=MsoNormal>"
end if




'Writes log file'

ReportTime =  Date & " " & Time

objFile.WriteLine"<html><head><meta http-equiv=Content-Type content='text/html; charset=windows-1252'><meta name=Generator content='Microsoft Word 14 (filtered)'><style><!-- /* Font Definitions */@font-face	{font-family:'Century Gothic';	panose-1:2 11 5 2 2 2 2 2 2 4;} /* Style Definitions */ p.MsoNormal, li.MsoNormal, div.MsoNormal	{margin-top:5.0pt;	margin-right:0in;	margin-bottom:5.0pt;	margin-left:0in;	font-size:9.0pt;	font-family:'Century Gothic','sans-serif';	color:black;} p.Warning, li.Warning, div.Warning	{margin-top:5.0pt;	margin-right:0in;	margin-bottom:5.0pt;	margin-left:0in;	font-size:9.0pt;	font-family:'Century Gothic','sans-serif';	color:red; font-weight:bold;} h1	{margin-top:12.0pt;	margin-right:0in;	margin-bottom:12.0pt;	margin-left:0in;	font-size:12.0pt;	font-family:'Century Gothic','sans-serif';	color:#E48312;	text-transform:uppercase;}h2	{margin-top:12.0pt;	margin-right:0in;	margin-bottom:5.0pt;	margin-left:0in;	background:#EADBD4;	font-size:11.0pt;	font-family:'Century Gothic','sans-serif';	color:#865640;	font-weight:normal;}h3	{margin-top:5.0pt;	margin-right:0in;	margin-bottom:5.0pt;	margin-left:0in;	font-size:9.0pt;	font-family:'Century Gothic','sans-serif';	color:#BD582C;	font-weight:normal;}p.Companyname, li.Companyname, div.Companyname	{mso-style-name:'Company name';	margin:0in;	margin-bottom:.0001pt;	text-align:center;	font-size:14.0pt;	font-family:'Century Gothic','sans-serif';	color:#49533D;	font-weight:bold;}--></style></head>"
objfile.Writeline"<body lang=EN-US><div class=WordSection1>"

objfile.Writeline"<p class=Companyname>NorWor Computer Diagnostics Tool</p><h1>Computer Status Report</h1><a name='CompInfo'><h2>Computer information</h2></a>"
objfile.Writeline"<table class=ListTable2 border=1 cellspacing=0 cellpadding=0 summary='Computer information' width='100%' style='width:100.0%;border-collapse: collapse;border:none'>"
objfile.Writeline"<tr>" & TableTR1 & "<h3>Computer Name:</h3> </td>"
objfile.Writeline TableTR2 & "<p class=MsoNormal>" & CompName & "</p> </td>"
objfile.Writeline "<td width=120 valign=top style='width:90pt;border-top:solid #666666 1.0pt;  border-left:none;border-bottom:solid #666666 1.0pt;border-right:none;  padding:0in 0in 0in 0in'><h3>Operating System:</h3> </td>"
objfile.Writeline " <td width=240 valign=top style='width:180pt;border-top:solid #666666 1.0pt;  border-left:none;border-bottom:solid #666666 1.0pt;border-right:none;  padding:0in 0in 0in 0in'>  <p class=MsoNormal>" & StrOS & "</p>  </td> </tr>"
objfile.Writeline" <tr>" & TableFormat1 &"  <h3>Scan Begin:</h3>  </td>" &  TableFormat2 &"  <p class=MsoNormal>" & StartTime & "</p>  </td>"
objfile.Writeline  TableFormat1 &"  <h3>Report end date</h3>  </td>" & TableFormat2 &"  <p class=MsoNormal>" & ReportTime & "</p>  </td> </tr>"
objfile.Writeline "<tr>" & TableFormat1 &"  <h3>General Status</h3> </td>" & TableFormat2 &"  <p class=MsoNormal>" & CompStatus & "</p>  </td>"
objfile.Writeline TableFormat1 &"  <h3>Antivirus Alerts</h3>  </td>" & TableFormat2 & VirusPClass & checkFound & "</p>  </td> </tr></table>"

if (CompStatus = "Problems Found!") then
	objfile.Writeline "<h2>Problems Found:</h2>" & CompStatusDesc 
end if
objfile.Writeline "<a name='Disks'></a>"
objFile.WriteLine HDLOGTEXT

objfile.Writeline "<a name='CPUInfo'><h2>CPU Usage Stats:</h2></a>"
objfile.Writeline tablestart
Objfile.Writeline " <tr> " & TableTR1 & " <h3>Processor Name</h3></td>" & TableTR2 & "  <p class=MsoNormal>" & CpuName & "</p></td> "
objfile.Writeline " " & TableTR1 & "  <h3>Average CPU usage:</h3>  </td>  " & TableTR2  & CPUPClass & CPUAVG & "%</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Address width</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CpuAddrWidth & "-Bit</p>  </td> "
objfile.Writeline " " & TableFormat1 & "  <h3>Description</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CPUDescription & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Manufacturer</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CPUManufacturer & "</p>  </td> "
objfile.Writeline "  " & TableFormat1 & "  <h3>Processor ID</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CPUProcessorId & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Revision</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CPURevision & "</p>  </td> "
objfile.Writeline "  " & TableFormat1 & "  <h3>Socket Type</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CPUSocket & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Current Clock Speed</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CPUCurrentClock & "</p>  </td> "
objfile.Writeline "  " & TableFormat1 & "  <h3>Maximum Clock SPeed</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CPUCurrentClock & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>L2 Cache Size</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CPUL2Size & "</p>  </td> "
objfile.Writeline "  " & TableFormat1 & "  <h3>&nbsp;</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>&nbsp;</p>  </td> </tr>"
objfile.writeline "</table>"

objfile.Writeline "<a name='MemInfo'><h2>Memory Statistics</h2></a>"
objfile.Writeline tablestart
Objfile.Writeline " <tr> " & TableTR1 & " <h3>Total Physical Memory:</h3></td>" & TableTR2 & "  <p class=MsoNormal>" & Round(totalMem / kbGB,1) &  "GB</p></td> "
objfile.Writeline " " & TableTR1 & "  <h3>Free Physical Memory:</h3>  </td>  " & TableTR2  & "  <p class=MsoNormal>" & Round(freeMem / kbGB,1) &  "GB</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Memory Usage: </h3>  </td>  " & TableFormat2  & MemPClass & Round((totalMem - freeMem) / totalMem * 100) & "%</p>  </td> "
objfile.Writeline "   " & TableFormat1 & "  <h3>Availible memory:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & AvailableGB & " GB</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Commit Limit:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CommitLimit & " GB</p>  </td> "
objfile.Writeline "  " & TableFormat1 & "  <h3>Committed memory:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CommittedGB & " GB</p>  </td> </tr>"
objfile.writeline "</table>"

objfile.writeline MEMLOGTEXT

objfile.Writeline "<a name='OSInfo'><h2>Operating System Info</h2></a>"
objfile.Writeline tablestart
Objfile.Writeline " <tr> " & TableTR1 & " <h3>Operating System:</h3></td>" & TableTR2 & "  <p class=MsoNormal>" & strOS & "</p></td> "
objfile.Writeline "  " & TableTR1 & "  <h3>Boot Device:</h3>  </td>  " & TableTR2  & "  <p class=MsoNormal>" & BootDevice & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Build Number:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & BuildNumber & "</p>  </td> "
objfile.Writeline " " & TableFormat1 & "  <h3>Build Type:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & BuildType & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Caption:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & Caption & "</p>  </td> "
objfile.Writeline "   " & TableFormat1 & "  <h3>Code Set:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CodeSet & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Country Code:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & CountryCode & "</p>  </td> "
objfile.Writeline "  " & TableFormat1 & "  <h3>Encryption Level:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & EncryptionLevel & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Install Date:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & dtmInstallDate & "</p>  </td> "
objfile.Writeline "  " & TableFormat1 & "  <h3>&nbsp;</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>&nbsp;</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>OS Product Suite:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & OSProductSuite & "</p>  </td> "
objfile.Writeline "   " & TableFormat1 & "  <h3>OS Type:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & OSType & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Primary:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & Primary & "</p>  </td> "
objfile.Writeline "   " & TableFormat1 & "  <h3>Registered User:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & RegisteredUser & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Serial Number:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & SerialNumber & "</p>  </td> "
objfile.Writeline "   " & TableFormat1 & "  <h3>Version:</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & Version & "</p>  </td> </tr>"
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Service Pack version (major):</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & SPmajor & "</p>  </td> "
objfile.Writeline "   " & TableFormat1 & "  <h3>Service Pack version (minor):</h3>  </td>  " & TableFormat2  & "  <p class=MsoNormal>" & SPminor & "</p>  </td> </tr>"
objfile.writeline "</table>"

objfile.Writeline "<a name='VirusInfo'><h2>Virus Scan Results</h2></a><p class=MsoNormal>" & sbFound &"</p>"

objfile.WriteLine "<a name='FWInfo'><h2>Firewall Status</h2></a>"
if (DomainProfStatus <> "0") then
   objfile.WriteLine "<p class=MsoNormal>" & DomainProfStatus & "</p>"
end if   
if (PrivateProfStatus <> "0") then
   objfile.WriteLine "<p class=MsoNormal>" & PrivateProfStatus & "</p>"
end if   
if (PublicProfStatus <> "0") then
   objfile.WriteLine "<p class=MsoNormal>" & PublicProfStatus & "</p>"
end if 
if (PublicProfStatus = "0" AND DomainProfStatus= "0" AND PrivateProfStatus = "0") then
	objfile.Writeline "<p class=Warning>WARNING! No firewall has been detected!</p>"
end if

objfile.Writeline "<a name='ShutdownInfo'><h2>Unexpected Shutdowns</h2></a>"
objfile.Writeline tablestart
objfile.Writeline " <tr>  " & TableFormat1 & "  <h3>Total Unexpected Shutdowns (all recorded errors):</h3>  </td>  " & TableFormat2  & "   <p class=MsoNormal>"  & TotalBlueScreens &"</p>  </td> "
objfile.Writeline "   " & TableFormat1 & "  <h3>Recent Unexpected Shutdowns (only within last month):</h3>  </td>  " & TableFormat2  & BluePClass & RecentBlueScreens & "</p>  </td> </tr>"
objfile.writeline "</table>"

objfile.Writeline"</div></body></html>"


msgbox "Scan Complete! Results saved to your desktop as 'NorWor.html'" 
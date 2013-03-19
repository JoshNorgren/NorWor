'Initial variable declarations'
sHostName = "."
strComputer = "."
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("results.log", True)

msgbox "Scanning Computer..."

msgbox "Obtaining Harddrive statistics..."
'GET HARD DRIVE SPACE'
GB = 1024 *1024 * 1024 'Number of bytes in a gigabyte'
kbGB = 1024 * 1024 ' Number of kilobytes in a GB'
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='c:'")
FreeMegaBytes = objLogicalDisk.FreeSpace / GB
TotalSpace = objLogicalDisk.Size / GB
FSType = objLogicalDisk.Filesystem
HDCaption = objLogicalDisk.Caption
HDDescription = objLogicalDisk.Description
HDDeviceID = objLogicalDisk.DeviceID
HDLastError = objLogicalDisk.LastErrorCode
HDStatus = objLogicalDisk.Status
HDVolumeDirty = objLogicalDisk.VolumeDirty
HDVolumeName = objLogicalDisk.VolumeName
HDSN = objLogicalDisk.VolumeSerialNumber



msgbox "Obtaining CPU statistics..."
'GET CPU USAGE' 
Samples = 5
CPUtotal = 0
Set RefresherObject = CreateObject("WbemScripting.SWbemRefresher")
Set objWMIService = GetObject("winmgmts:\\" & strCOmputer & "\root\cimv2")

Set ProcessorObjects = _
	RefresherObject.Addenum(objWMIService, "Win32_processor").ObjectSet

RefresherObject.Refresh


objfile.Writeline "DEBUG INFO (this will go away in final versions)"
For i = 1 to Samples
	RefresherObject.Refresh
	For Each Sampling in ProcessorObjects
		ObjFile.WriteLine "     Sample " & i & " " & Sampling.DeviceID & " usage: " & Sampling.LoadPercentage & "%"
		If (IsNull(Sampling.LoadPercentage)) then
			CPUtotal = CPUtotal
		else
			CPUtotal = CPUtotal + Sampling.LoadPercentage
		end if
	Next
Next
CPUavg= CPUtotal / Samples
objfile.Writeline""
msgbox "Obtaining Memory statistics..."



'GET MEM USAGE' 'This part DEF needs to go before we start running any programs'
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" _
    & strComputer & "\root\cimv2")

'	'CALCULATES AVG MEM USE'

Set MemRefresherObject = CreateObject("WbemScripting.SWbemRefresher")
Set MemoryObjects = MemRefresherObject.Addenum(objWMIService, "Win32_OperatingSystem").ObjectSet
MemRefresherObject.Refresh

ObjFile.Writeline "Mem AVG"

For i = 1 to Samples
	MemRefresherObject.Refresh
	Set colSettings = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	For Each objOperatingSystem in colSettings 
		freeMem = objOperatingSystem.FreePhysicalMemory
	Next

	Set colSettings = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
	
	For Each objOperatingSystem in colSettings 
		totalMem = objOperatingSystem.TotalPhysicalMemory / 1024
	Next
	memSample = Round((totalMem - freeMem) / totalMem * 100)
	ObjFile.WriteLine "     Sample " & i & " " & "Mem usage: " & memSample & "%"
	If (IsNull(memSample)) then
		MemUseTotal = MemUsetotal
	else
		MemUsetotal = MemUsetotal + memSample
	end if	
Next
MemAvg = memUseTotal / Samples
objfile.Writeline ""
'	'END MEM AVG'

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

msgbox "Obtaining Operating System statistics..."


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


msgbox "Testing for firewall..."


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
      DomainProfStatus = "     Firewall is ON on domain profile."
   else
      DomainProfStatus = "     Firewall is OFF on domain profile."
   end if
else
    DomainProfStatus = "0"
end if

if ( CurrentProfiles AND NET_FW_PROFILE2_PRIVATE ) then
   if fwPolicy2.FirewallEnabled(NET_FW_PROFILE2_PRIVATE) = TRUE then
      PrivateProfStatus = "     Firewall is ON on private profile."
   else
      PrivateProfStatus = "     Firewall is OFF on private profile."
   end if
else
    PrivateProfStatus = "0"
end if

if ( CurrentProfiles AND NET_FW_PROFILE2_PUBLIC ) then
   if fwPolicy2.FirewallEnabled(NET_FW_PROFILE2_PUBLIC) = TRUE then
      PublicProfStatus = "     Firewall is ON on public profile."
   else
      PublicProfStatus = "     Firewall is OFF on public profile."
   end if
else
    PublicProfStatus = "0"
end if



'Checks Event log for Unexpected Shutdowns' 'This part is REALLY innefficient and slow if you have a sufficiently large event log'
msgbox "Checking computer for errors." &VBnewline & "This could take a moment, please be patient."
Set colLoggedEvents = objWMIService.ExecQuery _
("Select * from Win32_NTLogEvent Where Logfile = 'System' and " _
& "EventCode = '6008'")


'Writes log file'

Objfile.Writeline "          SYSTEM STATISTICS"
Objfile.Writeline ""
objFile.WriteLine "Hard drive info "
objFile.WriteLine "     DeviceID: " & HDDeviceID
objFile.WriteLine "     File System: " & FSType
objFile.WriteLine "     Total Hard Drive Size: " & Int(TotalSpace) & " GB"
objFile.WriteLine "     Free space: " & Int(FreeMegabytes) & " GB"
objFile.WriteLine "     Caption: " & HDCaption
objFile.WriteLine "     Description: " & HDDescription
objFile.WriteLine "     Last Error Code: " & HDLastError
objFile.WriteLine "     Current Status: " & HSStatus
If (HDVolumeDirty = TRUE) then
	objFile.WriteLine "     Warning! Your Harddrive needs to be cleaned. You should run disk check."
end if
objFile.WriteLine "     VolumeName: " & HDVolumeName
objFile.WriteLine "     SerialNumber: " & HDSN
Objfile.Writeline ""
ObjFile.Writeline "CPU Usage Stats:"
ObjFile.WriteLine "     CPU TOTAL: " & CPUtotal 
ObjFile.WriteLine "     Average CPU usage: " & CPUAVG & "%"
Objfile.Writeline ""
objFile.WriteLine "Memory Statistics"
objFile.Writeline "     Free Physical memory: " & Round(freeMem / kbGB,1) & " GB"
objFile.Writeline "     Total Physical Memory: " & Round(totalMem / kbGB,1) & " GB"
objFile.Writeline "     Memory Usage: " & Round((totalMem - freeMem) / totalMem * 100) & "%"
objFile.Writeline "     Average Memory Usage: " & memAvg
objFile.WriteLine "     Availible memory: " & AvailableGB & " GB"
objFile.WriteLine "     Commit Limit: " & CommitLimit & " GB"
objFile.WriteLine "     Committed memory: " & CommittedGB & " GB"
Objfile.Writeline ""
objFile.WriteLine "Operating System Info"
objFile.WriteLine "     Operating System: " & strOS
objFile.WriteLine "     Boot Device: " & BootDevice
objFile.WriteLine "     Build Number: " & BuildNumber
objFile.WriteLine "     Build Type: " & BuildType
objFile.WriteLine "     Caption: " & Caption
objFile.WriteLine "     Code Set: " & CodeSet
objFile.WriteLine "     Country Code: " & CountryCode
objFile.WriteLine "     Encryption Level: " & EncryptionLevel
objFile.WriteLine "     Install Date: " & dtmInstallDate 
objFile.WriteLine "     Licensed Users: " & NumberOfLicensedUsers
objFile.WriteLine "     OS Product Suite: " & OSProductSuite
objFile.WriteLine "     OS Type: " & OSType
objFile.WriteLine "     Primary: " & Primary
objFile.WriteLine "     Registered User: " & RegisteredUser
objFile.WriteLine "     Serial Number: " & SerialNumber
objFile.WriteLine "     Version: " & Version
objFile.WriteLine "     Service Pack version (major): " & SPmajor 
objFile.WriteLine "     Service Pack version (minor): " & SPminor 
objFile.WriteLine "     Install Date: " & InstallDate 
Objfile.Writeline ""
objfile.WriteLine "Firewall Status"
if (DomainProfStatus <> "0") then
   objfile.WriteLine DomainProfStatus
end if   
if (PrivateProfStatus <> "0") then
   objfile.WriteLine PrivateProfStatus
end if   
if (PublicProfStatus <> "0") then
   objfile.WriteLine PublicProfStatus
end if   
Objfile.Writeline ""
objFile.Writeline "Unexpected shutdowns: " & colLoggedEvents.Count

msgbox "Scan Complete! Results saved to results.log" 'Leave this line at the very end of the script for debugging purposes'
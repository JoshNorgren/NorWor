'Initial variable declarations'
sHostName = "."
strComputer = "."
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile("results.log", True)

Wscript.Echo "Scanning Computer..."

Objfile.Writeline "          SYSTEM STATISTICS"

Objfile.Writeline ""
Wscript.Echo "Obtaining Harddrive statistics..."
'GET HARD DRIVE SPACE'
GB = 1024 *1024 * 1024 'Number of bytes in a gigabyte'
kbGB = 1024 * 1024 ' Number of kilobytes in a GB'
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set objLogicalDisk = objWMIService.Get("Win32_LogicalDisk.DeviceID='c:'")
FreeMegaBytes = objLogicalDisk.FreeSpace / GB
TotalSpace = objLogicalDisk.Size / GB
FSType = objLogicalDisk.Filesystem

objFile.WriteLine "Hard drive info "
objFile.WriteLine "     File System: " & FSType
objFile.WriteLine "     Total Hard Drive Size: " & Int(TotalSpace) & " GB"
objFile.WriteLine "     Free space: " & Int(FreeMegabytes) & " GB"


Objfile.Writeline ""
Wscript.Echo "Obtaining CPU statistics..."
'GET CPU USAGE' 'I'll need to rework this part to calculate the samples it gets into an average. Shouldn't be too hard but it's not there yet so... :V'
Samples = 5
Set RefresherObject = CreateObject("WbemScripting.SWbemRefresher")
Set objWMIService = GetObject("winmgmts:\\" & strCOmputer & "\root\cimv2")

Set ProcessorObjects = _
	RefresherObject.Addenum(objWMIService, "Win32_processor").ObjectSet

RefresherObject.Refresh
ObjFile.Writeline "CPU Usage Stats:"

For i = 1 to Samples
	RefresherObject.Refresh
	For Each Sampling in ProcessorObjects
		ObjFile.WriteLine "     Sample " & i & " " & Sampling.DeviceID & " usage: " & Sampling.LoadPercentage & "%"
	Next
Next


Objfile.Writeline ""
Wscript.Echo "Obtaining Memory statistics..."



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
objFile.WriteLine "Memory Statistics"
objFile.Writeline "     Free Physical memory: " & Round(freeMem / kbGB,1) & " GB"
objFile.Writeline "     Total Physical Memory: " & Round(totalMem / kbGB,1) & " GB"
objFile.Writeline "     Memory Usage: " & Round((totalMem - freeMem) / totalMem * 100) & "%"


'get other mem Stats'
Set Memory = GetObject("winmgmts:{impersonationLevel=impersonate}!//" _
 & sHostName & "/root/cimv2:Win32_PerfFormattedData_PerfOS_Memory=@")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory",,48)

For Each objItem in colItems
    AvailableGB = Round(objItem.AvailableBytes / GB,3)
    CommitLimit = Round(objItem.CommitLimit / GB,3)
    CommittedGB = Round(objItem.CommittedBytes / GB,3)

Next

objFile.WriteLine "     Availible memory: " & AvailableGB & " GB"
objFile.WriteLine "     Commit Limit: " & CommitLimit & " GB"
objFile.WriteLine "     Committed memory: " & CommittedGB & " GB"



Objfile.Writeline ""
Wscript.Echo "Obtaining Operating System statistics..."


'Get OS info'
Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
Set oss = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

objFile.WriteLine "Operating System Info"
For Each os in oss 											'Most of these are unnecessary but I'm leaving them in for now'
    objFile.WriteLine "     Boot Device: " & os.BootDevice
    objFile.WriteLine "     Build Number: " & os.BuildNumber
    objFile.WriteLine "     Build Type: " & os.BuildType
    objFile.WriteLine "     Caption: " & os.Caption
    objFile.WriteLine "     Code Set: " & os.CodeSet
    objFile.WriteLine "     Country Code: " & os.CountryCode
    objFile.WriteLine "     Debug: " & os.Debug
    objFile.WriteLine "     Encryption Level: " & os.EncryptionLevel
    dtmConvertedDate.Value = os.InstallDate
    dtmInstallDate = dtmConvertedDate.GetVarDate
    objFile.WriteLine "     Install Date: " & dtmInstallDate 
    objFile.WriteLine "     Licensed Users: " & os.NumberOfLicensedUsers
    objFile.WriteLine "     Organization: " & os.Organization
    objFile.WriteLine "     OS Language: " & os.OSLanguage
    objFile.WriteLine "     OS Product Suite: " & os.OSProductSuite
    objFile.WriteLine "     OS Type: " & os.OSType
    objFile.WriteLine "     Primary: " & os.Primary
    objFile.WriteLine "     Registered User: " & os.RegisteredUser
    objFile.WriteLine "     Serial Number: " & os.SerialNumber
    objFile.WriteLine "     Version: " & os.Version
Next



Objfile.Writeline ""
Wscript.Echo "Testing for firewall..."


'Firewall test'
'(I straight up borrowed this from somewhere, I'll need to see if I can peel the unnecessary bits off of it and make it less of a huge thing)'
objfile.WriteLine "Firewall Status"


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
      objFile.Writeline("     Firewall is ON on domain profile.")
   else
      objFile.Writeline("     Firewall is OFF on domain profile.")
   end if
end if

if ( CurrentProfiles AND NET_FW_PROFILE2_PRIVATE ) then
   if fwPolicy2.FirewallEnabled(NET_FW_PROFILE2_PRIVATE) = TRUE then
      objFile.Writeline("     Firewall is ON on private profile.")
   else
      objFile.Writeline("     Firewall is OFF on private profile.")
   end if
end if

if ( CurrentProfiles AND NET_FW_PROFILE2_PUBLIC ) then
   if fwPolicy2.FirewallEnabled(NET_FW_PROFILE2_PUBLIC) = TRUE then
      objFile.Writeline("     Firewall is ON on public profile.")
   else
      objFile.Writeline("     Firewall is OFF on public profile.")
   end if
end if



'Checks Event log for Unexpected Shutdowns' 'This part is REALLY innefficient and slow if you have a sufficiently large event log'
WScript.Echo "Checking computer for errors." &VBnewline & "This could take a moment, please be patient."
Set colLoggedEvents = objWMIService.ExecQuery _
("Select * from Win32_NTLogEvent Where Logfile = 'System' and " _
& "EventCode = '6008'")
Objfile.Writeline ""
objFile.Writeline "Unexpected shutdowns: " & colLoggedEvents.Count





Wscript.Echo "Scan Complete! Results saved to results.log" 'Leave this line at the very end of the script for debugging purposes'
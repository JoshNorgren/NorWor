strComputer = "."


' Enable error handling
On Error Resume Next

' Connect to specified computer
Set objWMIService = GetObject( "winmgmts://" & strComputer & "/root/cimv2" )
' Display error number and description if applicable
If Err Then ShowError

' Query video adapter properties
Set colItems = objWMIService.ExecQuery( "Select * from Win32_VideoController", , 48 )
' Display error number and description if applicable
If Err Then ShowError

' Prepare display of results
For Each objItem in colItems
   strmsg = objItem.Name &  " "  & Int( objItem.AdapterRAM / 1024000 )  & "MB"
Next

PWD = Left(WScript.ScriptFullName,(Len(WScript.ScriptFullName) - (Len(WScript.ScriptName) + 1)))
' Display results
Dim objResult
strmsg = strmsg & " - " & GetOSINFO
Set objShell = WScript.CreateObject("WScript.Shell")    
msgbox "Starting DXDIAG"
objResult = objShell.Run("dxdiag /t " & pwd & "\" & strMsg)
if err <>  ""then msgbox "DXDIAG Running"

'Done

WScript.Quit(0)

'/////////////////////////////////////////////////////////////////
'                       END MAIN
'/////////////////////////////////////////////////////////////////

Sub ShowError()
        strMsg = vbCrLf & "Error # " & Err.Number & vbCrLf & _
                 Err.Description & vbCrLf & vbCrLf
        Syntax
End Sub


function GetOSINFO()
strINFO = " "
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystems
    strINFO= objOperatingSystem.Caption
Next
if instr(strInfo,"7") then strINFO = "7"
if instr(strInfo,"Vista") then strINFO = "Vista"
if instr(strInfo,"XP") then strINFO = "XP"

IF GetObject("winmgmts:root\cimv2:Win32_Processor='cpu0'").AddressWidth  = 64 THEN
strINFO= strINFO & " 64 bit"  
else

strINFO= strINFO & " 32 bit"
END IF
GetOSINFO= strINFO
end function
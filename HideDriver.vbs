'
'
'	Hiding the driver in Windows
'
'   Set the HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDrives Data
'	Need Administrator authorization, so use "Shell.application.shellexecute" to RunAs
'
On Error Resume Next

Dim strDriver
Dim cShell, cExeRs, sData, nDrivers, sDriver, nPos, bRevert

Function ReadDrivers ()
	ReadDrivers = 0
	ReadDrivers = cShell.RegRead ("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDrives")
End Function


Set cShell = Wscript.CreateObject("WScript.Shell")

strDriver = UCase(InputBox ("Hidden the Driver:" & vbcrlf & _
														"[~][Driver Letter][...]" & vbcrlf & _
														"Use ~ to show the driver." & vbcrlf & _
 														"Don't input anything to show the current hidden drivers.", "Which driver you like to hide or show?"))
nDrivers = ReadDrivers ()

If strDriver <> "" Then
	bRevert = False
	For nPos = 1 to Len (strDriver)
		sDriver = Mid (strDriver, nPos, 1)
		If sDriver >= "A" And sDriver <= "Z" Then
			If bRevert Then	'Show the drivers
				nDrivers = (nDrivers AND NOT(2 ^ (ASC(sDriver)-ASC("A"))))
			Else
				nDrivers = (nDrivers OR (2 ^ (ASC(sDriver)-ASC("A"))))
			End If
			bRevert = False
		ElseIf sDriver = "~" Then
			bRevert = True
		Else
			bRevert = False
		End If
	Next
	WScript.CreateObject("Shell.application").shellexecute "REG", "ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer /v NoDrives /t REG_DWORD /d " & nDrivers & " /f","","RunAs",0
ElseIf nDrivers <> 0 Then
	strDriver = ""
	For nPos = 1 to 26
		If (nDrivers AND (2 ^ (nPos - 1))) <> 0 Then
			strDriver = strDriver & Chr(Asc("A") + nPos - 1)
		End If
	Next
	MsgBox "Drivers [" & strDriver & "] is hidden."
Else
	MsgBox "No any driver is hidden."
End If

Set cShell = Nothing
WScript.Quit

'
' End of file
'

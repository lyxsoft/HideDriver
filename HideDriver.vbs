On Error Resume Next

Dim strDriver
Dim cShell, cExeRs, sData, nDrivers

Function ReadDrivers ()
	ReadDrivers = 0
	ReadDrivers = cShell.RegRead ("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoDrives")
End Function


Set cShell = Wscript.CreateObject("WScript.Shell")


strDriver = UCase(InputBox ("Hidden the Driver:" & vbcrlf & "Don't input to check the current hidden drivers.", "Which driver you like to hide?"))
nDrivers = ReadDrivers ()

If strDriver <> "" Then
	WScript.CreateObject("Shell.application").shellexecute "REG", "ADD HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer /v NoDrives /t REG_DWORD /d " & (nDrivers OR (2 ^ (ASC(strDriver)-ASC("A")))),"","runas",0
ElseIf nDrivers <> 0 Then
	Dim nPos, sDrivers

	For nPos = 1 to 26
		If (nDrivers AND (2 ^ (nPos - 1))) <> 0 Then
			sDrivers = sDrivers & Chr(Asc("A") + nPos - 1)
		End If
	Next
	MsgBox "Drivers [" & sDrivers & "] is hidden."
Else
	MsgBox "No any driver is hidden."
End If

Set cShell = Nothing


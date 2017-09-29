' Check wheather it's passed 20:00 every 5 minutes
oneMinInMillis = 60000
Do While (Hour(Now) < 8)
	WScript.sleep oneMinInMillis*5
Loop


Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

Do
	checkTimer = 1
	secondsBetweenPops = 5
	For i = 0 To checkTimer
		shutDownAnswer = objShell.Popup("Doriti sa inchideti calculatorul?" & vbCrLf & "Se inchide in " & (checkTimer-i)*secondsBetweenPops & " secunde!", secondsBetweenPops, "Inchidere", vbYesNo)
		If (shutDownAnswer <> -1) Then
			Exit For
		End If
	Next
	
	
	Select Case shutDownAnswer

	'vbYes
	Case 6
		' Delay for half a second because it's a different prompt
		WScript.sleep 500

		For i = 0 To checkTimer
			shutDownAnswer = objShell.Popup("Sunteti sigur?" & vbCrLf & "Se inchide in " & (checkTimer-i)*secondsBetweenPops & " secunde!", secondsBetweenPops, "Inchidere", vbYesNo)
			If (shutDownAnswer <> -1) Then
				Exit For
			End If
		Next
		

	'vbNo
	Case 7
		' Moved vbNo outside of the Select Case for program flow reasons
		' If the statements are left here:
		' 	Selecting to shutdown brings up the "Are you sure" dialogue
		' 	Selecting to delay the shutdown here, won't process "case 7"


	'nothing pressed
	Case -1
		' Delay for half a second because it's a different prompt
		WScript.sleep 500
		
		For i = 0 To checkTimer
			shutDownAnswer = objShell.Popup("Nu ati selectat nici o optiune!" & vbCrLf & "Doriti sa inchideti calculatorul?" & vbCrLf & "Se inchide in " & (checkTimer-i)*secondsBetweenPops & " secunde!", secondsBetweenPops, "Inchidere", vbYesNo)
			If (shutDownAnswer <> -1) Then
				Exit For
			End If
		Next

	End Select
	


	If (shutDownAnswer = 6 Or shutDownAnswer = -1) Then
		'objShell.Run "C:\WINDOWS\system32\shutdown.exe -s -t 0"
		MsgBox "Selected NOTHING and shuting down..."
	End If

	If (shutDownAnswer = 7) Then
		MsgBox "Inchidere amanata " & (oneMinInMillis*30)/60000 & " minute."
		WScript.sleep 1500'oneMinInMillis*30
	End If
	
Loop While (shutDownAnswer = 7)
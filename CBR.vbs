oneMinInMilli = 60000
Do While (Hour(Now) < 8)
	WScript.sleep oneMinInMilli*5
Loop



oneMinInSec = 60
Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

Do
	checkTimer = 1
	For i = 0 To checkTimer
		shutDownAnswer = objShell.Popup("Doriti sa inchideti calculatorul?" & vbCrLf & "Se inchide in " & (checkTimer-i)*5 & " secunde!", 5, "Inchidere", vbYesNo)
		If (shutDownAnswer <> -1) Then
			Exit For
		End If
	Next
	
	
	Select Case shutDownAnswer
	Case 6	'vbYes
		WScript.sleep 500
		For i = 0 To checkTimer
			shutDownAnswer = objShell.Popup("Sunteti sigur?" & vbCrLf & "Se inchide in " & (checkTimer-i)*5 & " secunde!", 5, "Inchidere", vbYesNo)
			If (shutDownAnswer <> -1) Then
				Exit For
			End If
		Next
		
		If (shutDownAnswer = 6) Then
			'objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 0"
			MsgBox "shuting down........."
		End If
		
	Case 7	'vbNo
		' Moved vbNo outside of the Select Case for program flow reasons
		' If the statements are left here:
		' 	Selecting to shutdown brings up the "Are you sure" dialogue
		' 	Selecting to delay the shutdown here, won't process "case 7"
	Case -1	'nothing pressed
		For i = 0 To checkTimer
			shutDownAnswer = objShell.Popup("Nu ati selectat nici o optiune!" & vbCrLf & "Doriti sa inchideti calculatorul?" & vbCrLf & "Se inchide in " & (checkTimer-i)*5 & " secunde!", 5, "Inchidere", vbYesNo)
			If (shutDownAnswer <> -1) Then
				Exit For
			End If
		Next
		If (shutDownAnswer = 6) Then
			'objShell.Run "C:\WINDOWS\system32\shutdown.exe -r -t 0"
			MsgBox "shuting down........."
		End If
	End Select
	
	If (shutDownAnswer = 7) Then
		MsgBox "Inchidere amanata " & oneMinInSec & " minute."
		WScript.sleep 2000
	End If
	
Loop While (shutDownAnswer = 7)

MsgBox "computer is shut down at this point"
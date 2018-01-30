Option Explicit

Const LogFile = "LocalAdmins.log"
Const resultFile = "LocalAdministratorsMembership.csv"
Const inputFile = "c:\scan\targets.txt"


Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

Dim shl
Set shl = WScript.CreateObject("WScript.Shell")

Dim fil
Set fil = fso.OpenTextFile(inputFile)

Dim results
Set results = fso.CreateTextFile(resultFile, True) 

WriteToLog "Beginning Pass of " & inputFile & " at " & Now() 
WScript.Echo "Beginning Pass of " & inputFile & " at " & Now() 
'On Error Resume Next

Dim grp
Dim line
Dim exec
Dim pingResults
Dim member

While Not fil.AtEndOfStream
	line = fil.ReadLine 
	
	Set exec = shl.Exec("ping -n 2 -w 1000 " & line)
  	pingResults = LCase(exec.StdOut.ReadAll)
 	
 	If InStr(pingResults, "reply from") Then
 		WriteToLog line & " responded to ping"
 		 		
 		'On Error Resume Next 

		Set grp = GetObject("WinNT://" & line & "/Administrators")
		
		results.WriteLine line & ",Administrators,"
		
		For Each member In grp.Members
			
			WriteToLog line & ": Administrators - " & member.Name
			results.WriteLine ",," & member.Name
		Next
	Else
		WriteToLog line & " did not respond to ping"
		
	End If 
Wend

results.Close 

Sub WriteToLog(LogData)
	On Error Resume Next

	Dim fil	
	'8 = ForAppending
	Set fil = fso.OpenTextFile(LogFile, 8, True)
		
	fil.WriteLine(LogData)	

	fil.Close
	Set fil = Nothing
End Sub

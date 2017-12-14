

on error resume next
Err.clear
' Define Variables
Dim strDomainDN, strRet, strAD1, OUFilter
Dim strEmailTo, strEmailFrom, strEmailServer

SendEmail = 0 ' Set to 1 to send the log as an e-mail
strEmailFrom = "<SMTP Address>" ' log e-mail FROM address
strEmailTo = "" ' log e-mail TO address
strEmailServer = "<SMTP SERVER>" ' your e-mail server
OUFilter = "" ' eg: OU=Employees, end it with a comma if you have anything here.

' Domain List
Dim rootDSE
Set rootDSE = GetObject("LDAP://RootDSE")
strConfig = rootDSE.Get("configurationNamingContext")
DomainContainer =  rootDSE.Get("defaultNamingContext")
Set conn = CreateObject("ADODB.Connection")
conn.Provider = "ADSDSOObject"
conn.Open "ADs Provider"

' Use ADO to identify all domain controllers. Need to query them all since LastLogon isn't replicated.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

strBase = "<LDAP://" & strConfig & ">"
strFilter = "(objectClass=nTDSDSA)"
strAttributes = "AdsPath"
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 120
adoCommand.Properties("Cache Results") = False

Set adoRecordset = adoCommand.Execute

' Generate a list of all AD controllers
d = 0
Do Until adoRecordset.EOF
    Set objDC = _
        GetObject(GetObject(adoRecordset.Fields("AdsPath").Value).Parent)
    ReDim Preserve arrstrDCs(d)
    arrstrDCs(d) = objDC.DNSHostName
    d = d + 1
    adoRecordset.MoveNext
Loop
adoRecordset.Close

' Going to need a place to hold LastLogon for comparison
Set objList = CreateObject("Scripting.Dictionary")
objList.CompareMode = vbTextCompare

' Need to know if we want Users, Computers or Both
Dim ldapStr

If WScript.Arguments.Count <> 1 Then ' Nothing passed on launch
	strRet = TRIM(InputBox ("Enter the objects you want to report on." & CHR(10) & "Include the optional letter D to include disabled objects." & CHR(10) & "Include the optional letter X for only disabled objects." & CHR(13) & CHR(13) _
		& "1" & CHR(91) & "D" & CHR(93) & CHR(91) & "X" & CHR(93) & " - Users" & CHR(10) & "2" & CHR(91) & "D" & CHR(93) & CHR(91) & "X" & CHR(93) & " - Computers" & CHR(10) & "3" & CHR(91) & "D" & CHR(93) & CHR(91) & "X" & CHR(93) & " - Users and Computers"))
Else
	strRet = wscript.Arguments(0)
End If

Select Case LEFT(strRet,1)
	' Users
	Case "1"
		ldapFltr = "(&(objectCategory=person)(objectClass=user)"

	' Computers
	Case "2"
		ldapFltr = "(&(objectCategory=computer)"

	' Users and Computers
	Case "3"
		ldapFltr = "(&(|(&(objectCategory=person)(objectClass=user))(objectCategory=computer))"

	Case else
		If WScript.Arguments.Count <> 1 Then
			wscript.echo "You didn't specify an option."  & CHR(10) & "Ending script."
		End If
		wscript.quit
End Select

If UCASE(RIGHT(strRet,1)) = "X" Then
	ldapFltr = ldapFltr & "(UserAccountControl:1.2.840.113556.1.4.803:=2))"
End If
IF UCASE(RIGHT(strRet,1)) = "D" Then
	ldapFltr = ldapFltr & ")"
End If
If (UCASE(RIGHT(strRet,1)) <> "D" and UCase(RIGHT(strRet,1))<>"X") Then
	ldapFltr = ldapFltr & "!(UserAccountControl:1.2.840.113556.1.4.803:=2))"
End if

' Make a log file
Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptPath = objfso.GetParentFolderName(WScript.ScriptFullName) & "\"
strLogName = "LastLogon_" & TwoDigits(Year(now)) & TwoDigits(Month(now)) & TwoDigits(Day(now)) & TwoDigits(Hour(now)) & TwoDigits(Minute(now)) & TwoDigits(Second(now)) & ".csv"
strLogFile = strScriptPath & StrLogName
Set objLogFile = objFSO.CreateTextFile(strLogFile,1)

' Obtain local Time Zone bias from machine registry.
' This bias changes with Daylight Savings Time.
Set objShell = CreateObject("Wscript.Shell")
lngBiasKey = objShell.RegRead("HKLM\System\CurrentControlSet\Control\" _
    & "TimeZoneInformation\ActiveTimeBias")
If (UCase(TypeName(lngBiasKey)) = "LONG") Then
    lngBias = lngBiasKey
ElseIf (UCase(TypeName(lngBiasKey)) = "VARIANT()") Then
    lngBias = 0
    For k = 0 To UBound(lngBiasKey)
        lngBias = lngBias + (lngBiasKey(k) * 256^k)
    Next
End If


For d = 0 To Ubound(arrstrDCs)
	nErrNo = 0

	'Create the LDAP query and execute
	strBase = "<LDAP://" & arrstrDCs(d) & "/" & OUFilter & DomainContainer & ">"
	strAttributes = "sAMAccountName,lastLogon"
    strQuery = strBase & ";" & ldapFltr & ";" & strAttributes & ";subtree"

	adoCommand.CommandText = strQuery
	Set adoRecordset = adoCommand.Execute
	nErrNo = Err.Number
	If nErrNo <> 0 Then 
		objLogFile.WriteLine "Domain Controller not available: " & arrstrDCs(d) & " " & nErrNo
	Else
	
	'Hold one responding controller name, we're going to need it later.
	If IsEmpty(strAD1) Then
		strAD1 = arrstrDCs(d)
	End If
	
	'Process it
		While NOT adoRecordset.EOF
			Set objLastLogon = adoRecordset.Fields("lastLogon").Value
			Set strcn = adoRecordset.Fields(1).Value

			
			IF IsEmpty(adoRecordset.Fields("lastLogon").Value) Then
				lngHigh = 0
				lngLow = 0
			ElseIF IsNull(adoRecordset.Fields("lastLogon").Value) Then
				lngHigh = 0
				lngLow = 0
			Else			
				lngHigh = objLastLogon.HighPart
				lngLow = objLastLogon.LowPart
			End If

			
			If (lngLow < 0) Then
				lngHigh = lngHigh + 1
			End If
			   
			If (lngHigh = 0) And (lngLow = 0) Then
				strLastLogon = CDATE(#1/1/1601#) 'This should be never
			Else
				strLastLogon = CDate(#1/1/1601# + (((lngHigh * (2 ^ 32)) + lngLow)/600000000 - lngBias)/1440)
			End If

			If (objList.Exists(adoRecordset.Fields("sAMAccountName").Value) = True) Then
				If (strLastLogon > objList(adoRecordset.Fields("sAMAccountName").Value)) Then
					objList.Item(adoRecordset.Fields("sAMAccountName").Value) = strLastLogon
				End If
			Else
				objList.Add adoRecordset.Fields("sAMAccountName").Value, strLastLogon
			End If
			
			adoRecordset.MoveNext
		Wend
	End If
	adoRecordset.Close
	Err.clear
Next

'Query again so we can report other fields
strBase = "<LDAP://" & strAD1 & "/" & OUFilter & DomainContainer & ">"
strAttributes = "sAMAccountName,cn,givenName,sn,description,distinguishedName,objectCategory,mail"
strQuery = strBase & ";" & ldapFltr & ";" & strAttributes & ";subtree"

adoCommand.CommandText = strQuery
Set adoRecordset = adoCommand.Execute

' Write compiled data to the log
objLogFile.WriteLine "Display Name, Logon Name, First Name, Last Name, E-Mail, Last Logon, Days Since Last Logon, Category, description, distinguishedName"

While NOT adoRecordset.EOF
	strObjectCategory = Mid(adoRecordset.Fields("objectCategory").value, 4, InStr(adoRecordset.Fields("objectCategory").value,",")-4)
	period=Now()-objList.Item(adoRecordset.Fields("sAMAccountName").Value)
	If period > 0 then
		desc_str=" "
		For Each item In adoRecordset.Fields("Description").Value
			desc_str = desc_str & item
		Next
		
		objLogFile.WriteLine CHR(34) & adoRecordset.Fields("cn").Value &  CHR(34) & "," &  CHR(34) & adoRecordset.Fields("sAMAccountName").Value & CHR(34) & "," _
			& CHR(34) & adoRecordset.Fields("givenName").Value & CHR(34) & "," & CHR(34) & adoRecordset.Fields("sn").Value & CHR(34) & "," & CHR(34) & adoRecordset.Fields("mail").Value & CHR(34) & "," _
			& CHR(34) & objList.Item(adoRecordset.Fields("sAMAccountName").Value) & CHR(34) & "," & CHR(34) & Int(period-0.5) & CHR(34) & "," & CHR(34) & strObjectCategory & CHR(34) _
			& "," & CHR(34) & desc_str & CHR(34)& "," & CHR(34) & adoRecordset.Fields("distinguishedName").Value & CHR(34)
	End if
	adoRecordset.MoveNext
Wend
adoRecordset.Close
objLogFile.close

' Email It

IF SendEmail = 1 Then
	Set objLogEmail = CreateObject("CDO.Message")
	objLogEmail.From = strEmailFrom
	objLogEmail.To = strEmailTo
	objLogEmail.Fields.Item("urn:schemas:mailheader:X-Priority") = 1
	objLogEmail.Fields.Item("urn:schemas:httpmail:importance") = 1
	objLogEmail.Fields.Update
	objLogEmail.Subject = "LastLogon Log"
		objLogEmail.HTMLBody = "<HTML><HEAD>" & vbcrlf & "</HEAD><BODY>"
		objLogEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		objLogEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strEmailServer
		objLogEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objLogEmail.Configuration.Fields.Update
		objLogEmail.AddAttachment strLogFile
		objLogEmail.Send 
End If

adoConnection.Close
Set RootDSE = Nothing
Set adoConnection = Nothing
Set adoCommand = Nothing
Set adoRecordset = Nothing
Set objDC = Nothing
Set objList = Nothing
Set objShell = Nothing
Set strAD1 = Nothing

If WScript.Arguments.Count <> 1 Then
	wscript.echo "Done!"
End If

Function TwoDigits(t)
   TwoDigits = Right("00" & t,2)
End Function

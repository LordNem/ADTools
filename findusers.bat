@echo off
@echo Enter Date format in this type YYYYMMDD Below
SET /P NEWDATE=

dsquery * domainroot -filter "(&(objectcategory=person)(objectclass=user)(whencreated>=%NEWDATE%000000.0Z))" -attr sAMAccountName mail description Department PhysicalDeliveryOfficeName WhenCreated -limit 0 >> C:\OUTPUT.txt
c:\OUTPUT.txt
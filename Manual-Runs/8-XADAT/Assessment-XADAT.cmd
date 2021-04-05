@echo off
cls
color 1a
:: Version 1.0 -- 8/12/2013 -- Initial script creation and development
:: Version 1.1 -- 9/01/2013 -- Added positive feedback for all of the script results
:: Version 1.2 -- 9/11/2013 -- corrected powershell execution policy issue and dsquery/dsget versioning issues
:: Version 1.2.1 -- 9/14/2013 -- corrected formatting of data output for users with non-expiring passwords, also added a check for PS version 2.0
::                            -- also corrected issue with UAC lookup
::
:: Date: 8/12/2013
:: Author: Bruce Portz, Kelly Bush - Directories Architect, Xerox / ACS

::for /f "tokens=1-4 delims=/ " %%s in ('date /t') do set dt=%%t-%%u-%%v

set log=%USERPROFILE%\Desktop\ADData\assessment-output.log
set notools=0

:BEGIN

echo.
echo.
echo.

echo **************************************************************************
echo *                                                                        *
echo *                                                                        *
echo *                 Active Directory Assessment Tool                       *
echo *                 Version 1.2.1                                          *
echo *                 AD RAP Lite                                            *
echo *                                                                        *
echo **************************************************************************
echo,
echo.
echo.


:Startup

echo  Setting the default scripting engine...
%systemroot%\system32\cscript //h:cscript //s>NUL


:: Check OS version info
echo  Getting the OS version information...
for /f "delims=" %%v in ('cscript //nologo 3264.vbs') do (Set OSVer=%%v)


:: Checking the Powershell version info
powershell $psversiontable.psversion.major >temp.zzz
for /f %%i in (temp.zzz) do (set temp123=%%i)

del temp.zzz /q
if not %temp123% GEQ 2 GOTO PSPROBLEM


echo Checking for tools....
if not exist %systemroot%\system32\dsquery.exe GOTO SETUP






:READY

date/t >%log%
time/t >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%

echo OS Version is %OSVer% bit.....
echo OS Version is %OSVer% bit..... >>%log%

echo. >>%log%
echo Now gathering Forest information.....
echo Now gathering Forest information..... >>%log%

powershell.exe -executionpolicy unrestricted -command .\forest-info.ps1 >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now gathering Trust information.....
echo Now gathering Trust information..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo check-trusts.vbs %logonserver% >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now checking Site Links.....
echo Now checking Site Links..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo check-sitelinks.vbs >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now checking for Sites without Subnets.....
echo Now checking for Sites without Subnets..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo check-sites.vbs >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now comparing GPOs between SYSVOL and AD.....
echo Now comparing GPOs between SYSVOL and AD..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo GPO-GUIDs-AD.vbs >GPOs-in-AD-but-not-in-sysvol.txt
dir /b \\%userdnsdomain%\sysvol\%userdnsdomain%\Policies | find /i "{" >GPOs-in-SYSVOL-but-not-in-AD.txt
%systemroot%\system32\cscript //nologo sort-data.vbs GPOs-in-AD-but-not-in-sysvol.txt
%systemroot%\system32\cscript //nologo sort-data.vbs GPOs-in-SYSVOL-but-not-in-AD.txt
%systemroot%\system32\cscript //nologo check-consistency.vbs >>%log%
del GPOs-in*.txt /q
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now gathering Orphaned GPOS.....
echo Now gathering Orphaned GPOS..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo orphaned-GPOs.vbs >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now gathering Domain Account Policies.....
echo Now gathering Domain Account Policies..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo account-policies.vbs >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now gathering Domain Controller Security info.....
echo Now gathering Domain Controller Security info..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo check-uac-settings.vbs >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now gathering List of Empty Groups.....
echo Now gathering List of Empty Groups..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo check-empty-groups.vbs >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now gathering List of Users with non-expiring passwords.....
echo Now gathering List of Users with non-expiring passwords..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo users-non-expiring-pwd.vbs >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now checking domain time sync status.....
echo Now checking domain time sync status..... >>%log%
echo. >>%log%

w32tm /monitor >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%


echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now gathering detailed DC operating system data.....
echo Now gathering detailed DC operating system data..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo DC-info-1.vbs >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now gathering detailed DC advertising data.....
echo Now gathering detailed DC advertising data..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo DC-info-2.vbs >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

echo *********************************************************************************** >>%log%
echo. >>%log%
echo Now gathering detailed DC time configuration.....
echo Now gathering detailed DC time configuration..... >>%log%
echo. >>%log%

%systemroot%\system32\cscript //nologo DC-info-3.vbs >>%log%
echo. >>%log%
echo. >>%log%
w32tm /monitor >>%log%
echo. >>%log%
echo. >>%log%
echo. >>%log%

GOTO END

:HELP

cls

echo.
echo.
echo                       T O O L   O V E R V I E W
echo *****************************************************************************
echo.
echo   This assessment tool is designed for performing an "ADRAP-Lite" assessment
echo on the Active Directory forest and domains in the forest.  The tool will be
echo able to collect some of the information from the forest root domain as well as
echo other domains, however without running this in the forest root domain with EA
echo credentials, some of the information will be missed.  It is recommended that
echo the assessment tool be run in each domain in the forest.
echo.
echo.
echo   The assessment tool does require that the person executing the assessment be
echo logged in with Domain Admin credentials.  Most of the information can be 
echo collected without DA credentials but the secured information will be skipped.
echo When an assessment is executed, the results will be saved in a log file in the
echo same directory as the assessment tool with the filename: 
echo "assessment-output-<CURRENT DATE>.log".  All of the information is categorized
echo in the log file by the script that gathered the information so it is more or
echo compartmentalized.
echo.
echo.
pause
cls

echo.
echo.
echo.

echo   The script pulls the following information and collects it in the log file:
echo.
echo     Forest information
echo     Trust Health information
echo     Conflicting (overlapping) subnets
echo     AD Sites without Subnets bound to it
echo     List of OUs in the domain and the GPOs bound to each
echo     List of orphaned GPOs
echo     GPO consistency (compare AD against SYSVOL)
echo     Domain Account (password) Policy settings
echo     Domain controller security (UAC) setting
echo     List of empty groups in the domain
echo     List of user accounts with "Password never expires"
echo     List of user accounts that are disabled
echo     List of Members in "sensitive"groups, meaning
echo        -- Enterprise Admins
echo        -- Schema Admins
echo        -- Domain Admins
echo.
pause
cls
echo.
echo.
echo.

echo   The script also has the ability to pull detailed information from the domain
echo controllers themselves - however - this can be very slow to execute, since it
echo depends on the connectivity between the script source and domain controllers as
echo well as the network speed, latency, bandwidth, etc.
echo.
echo.
echo Press ENTER to go back to the main menu...
pause>nul
GOTO BEGIN

:PSPROBLEM
color 4f
cls
echo.
echo. >>%log%
echo There is a problem -- we need at least Powershell version 2.0 installed.....
echo There is a problem -- we need at least Powershell version 2.0 installed..... >>%log%
echo.
echo. >>%log%
echo.
echo. >>%log%
echo Please install Powershell version 2.0 and re-run....
echo Please install Powershell version 2.0 and re-run.... >>%log%

pause

GOTO END
:SETUP

echo. >>%log%
echo Adding setup files to server....
echo Adding setup files to server.... >>%log%
echo. >>%log%

set notools=1

if %OSVer% ==64 copy .\64\ds*.* %systemroot%\system32\ds*.* /Y
if %OSVer% ==32 copy .\32\ds*.* %systemroot%\system32\ds*.* /Y

GOTO READY


:CLEANUP

del %systemroot%\system32\dsquery.* /q
del %systemroot%\system32\dsget.exe /q

set notools=0

:END

if %notools% == 1 GOTO CLEANUP
start notepad %log%
cls
color 1a

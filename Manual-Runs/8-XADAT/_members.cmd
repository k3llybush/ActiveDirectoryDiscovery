@echo off

set cnt=%1
set dom=%~2
set log=%3


if %cnt% == 1 GOTO ROOT

GOTO CHILD

:ROOT
echo Checking Root Domain %dom% >>%log%
echo Enterprise Admins >>%log%
echo ------------------------------------------------------- >>%log%

dsquery group forestroot -name "Enterprise Admins" | dsget group -members | dsget user -display -samid >>%log%
echo. >>%log%
echo. >>%log%

echo Schema Admins >>%log%
echo ------------------------------------------------------- >>%log%

dsquery group forestroot -name "Schema Admins" | dsget group -members | dsget user -display -samid >>%log%
echo. >>%log%
echo. >>%log%

echo Domain Admins >>%log%
echo ------------------------------------------------------- >>%log%

dsquery group domainroot -name "Domain Admins" | dsget group -members | dsget user -display -samid >>%log%
echo. >>%log%
echo. >>%log%

GOTO END

:CHILD
echo Checking Child Domain %dom% >>%log%
echo Domain Admins >>%log%
echo ------------------------------------------------------- >>%log%

dsquery group domainroot -name "Domain Admins" | dsget group -members | dsget user -display -samid >>%log%
echo. >>%log%
echo. >>%log%

GOTO END





:END
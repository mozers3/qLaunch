@echo off

rem Install-Module ps2exe before
rem or download from here: https://github.com/MScholtes/PS2EXE

chcp 65001>nul
CD /D "%~dp0"
set ps1=qLaunch.ps1
set exe=qLaunch.exe
set ico=qLaunch.ico
set title="ps Quick Launch"
set product="ps Quick Launch"
set ver="1.0.0"
set c=mozers™
set tm=mozers™
set comp=https://github.com/mozers3/qLaunch

powershell -Command Invoke-ps2exe -inputFile '%ps1%' -outputFile '%exe%' -x64 -noConsole -verbose -iconFile '%ico%' -title '%title%' -product '%product%' -copyright '%c%' -trademark '%tm%' -company '%comp%' -version '%ver%'

pause
